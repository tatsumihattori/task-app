// ============================================================
// TaskSync.gs — Sheets ↔ Calendar 双方向タスク同期
// ============================================================
// 【セットアップ手順】
//   1. Google Sheetsを開き、拡張機能 > Apps Script を選択
//   2. このスクリプトを貼り付けて保存
//   3. setup() を一度だけ実行してトリガーを登録
//   4. 初回実行時に「権限を許可」のダイアログが表示されるので許可する
// ============================================================

// ==================== 設定 ====================
const CONFIG = {
  SHEET_NAME: "タスク",          // シート名（必要に応じて変更）
  CALENDAR_ID: "c_60e6f80451cdf1045be018eff7cfa2fbe19fac4bc8dec036d29c8701eb9df4fe@group.calendar.google.com",
  ERROR_NOTIFY_EMAIL: "tatsumi.hattori@laboratory.work",  // エラー通知先（空文字で無効化）

  // ステータスごとのCalendarイベント色（CalendarApp.EventColor の数値文字列）
  STATUS_COLORS: {
    "未着手":    "2",  // Pale Green
    "進行中":    "9",  // Blue
    "完了":      "8",  // Gray
    "キャンセル": "7",  // Cyan
  },

  SYNC_DAYS_AHEAD: 90,           // Calendarからの取得範囲（今日から何日先まで）
  SYNC_DAYS_BACK: 7,             // Calendarからの取得範囲（今日から何日前まで）

  CHAT_NOTIFY_DAYS_BEFORE: 3,    // 期日の何日前から通知を開始するか（当日を含めて毎日通知）

  // 列番号（1始まり）
  COL: {
    TASK_NAME:    1,  // A: タスク名
    START:        2,  // B: 開始日時
    END:          3,  // C: 終了日時
    STATUS:       4,  // D: ステータス
    ASSIGNEE:     5,  // E: 担当者
    MEMO:         6,  // F: メモ/説明
    EVENT_ID:     7,  // G: Calendar Event ID（自動管理・編集不要）
    LAST_SYNC:    8,  // H: 最終同期日時（自動管理・編集不要）
    CREATED_BY:   9,  // I: 作成者（自動管理・編集不要）
    UPDATED_BY:  10,  // J: 最終更新者（自動管理・編集不要）
    PROJECT:     11,  // K: プロジェクト名（表示のみ・GAS未使用）
    DEADLINE:    12,  // L: 期日（Google Chat期日通知に使用）
  },

  HEADER_ROW: 1,
  DATA_START_ROW: 2,
};

// ==================== メインエントリ ====================

/**
 * 初回セットアップ：ヘッダー作成 & トリガー登録
 * ※ 一度だけ手動で実行してください
 */
function setup() {
  // _createHeader();
  _registerTriggers();
  SpreadsheetApp.getUi().alert(
    "✅ セットアップ完了！\n\n" +
    "・Sheetsを編集すると自動でCalendarに反映されます\n" +
    "・CalendarからSheetsへの同期は10分ごとに実行されます\n" +
    "・期日が近いタスクのGoogle Chat通知は毎日9時に実行されます\n" +
    "　（スクリプトプロパティ CHAT_WEBHOOK_URL が未設定の場合は通知されません）"
  );
}

/**
 * カレンダーをリセットしてシートの内容で全件再作成する（手動実行専用）
 * データ不整合やカレンダー移行後のリカバリに使用。
 */
function resetAndResyncToCalendar() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.alert(
    "カレンダーリセット確認",
    "シートに紐付いた全カレンダーイベントを削除し、シートの内容で再作成します。\nよろしいですか？",
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    ui.alert("別の同期処理が実行中です。しばらく待ってから再試行してください。");
    return;
  }

  try {
    const sheet       = _getOrCreateSheet();
    const calendarMap = _getAssigneeCalendarMap();
    const allCalIds   = [...calendarMap.values()];
    const col         = CONFIG.COL;
    const lastRow     = sheet.getLastRow();

    if (lastRow < CONFIG.DATA_START_ROW) {
      ui.alert("同期対象のデータがありません。");
      return;
    }

    const NUM_COLS = Object.keys(col).length;
    const rows = sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.HEADER_ROW, NUM_COLS).getValues();

    // Step 1: 既存イベントを削除して EVENT_ID をクリア
    let deletedCount = 0;
    rows.forEach(function(rowData, i) {
      const eventId  = rowData[col.EVENT_ID - 1];
      const assignee = String(rowData[col.ASSIGNEE - 1]).trim();
      if (!eventId) return;

      try {
        let deleted = false;
        if (assignee && calendarMap.has(assignee)) {
          const cal = CalendarApp.getCalendarById(calendarMap.get(assignee));
          const ev  = cal ? cal.getEventById(eventId) : null;
          if (ev) { ev.deleteEvent(); deleted = true; }
        }
        if (!deleted) {
          const found = _findEventInCalendars(eventId, allCalIds);
          if (found) found.event.deleteEvent();
        }
      } catch(e) {
        Logger.log("削除エラー: " + eventId + " " + e.message);
      }

      sheet.getRange(CONFIG.DATA_START_ROW + i, col.EVENT_ID).clearContent();
      deletedCount++;
    });

    // Step 2: シートの全行を再同期
    let createdCount = 0;
    for (let r = CONFIG.DATA_START_ROW; r <= lastRow; r++) {
      _syncRowToCalendar(sheet, r);
      createdCount++;
    }

    ui.alert("完了！\n" + deletedCount + "件削除 → " + createdCount + "行を再同期しました。");

  } catch(e) {
    _notifyError("resetAndResyncToCalendar", e);
    ui.alert("エラーが発生しました: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Sheets編集時に呼ばれるトリガー関数（installable onEdit）
 */
function onEditTrigger(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  const startRow = e.range.getRow();
  const numRows  = e.range.getNumRows();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log("onEditTrigger: ロック取得失敗（別の同期処理が実行中）");
    return;
  }
  try {
    for (let i = 0; i < numRows; i++) {
      const row = startRow + i;
      if (row <= CONFIG.HEADER_ROW) continue;
      _syncRowToCalendar(sheet, row);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * 定期実行：CalendarからSheetsへ同期
 * - 新規イベント    → 行を追加
 * - 既存イベントの変更 → 該当行を更新（日時・タイトル・担当者・メモ）
 */
function syncCalendarToSheets() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log("syncCalendarToSheets: ロック取得失敗（別の同期処理が実行中）スキップ");
    return;
  }
  try {
    _syncCalendarToSheetsBody();
  } catch(e) {
    _notifyError("syncCalendarToSheets", e);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 定期実行：期日が近いタスクをGoogle Chatに通知（担当者をメンション）
 * Webhook URLはスクリプトプロパティ "CHAT_WEBHOOK_URL" から読み込む（コードには書かない）
 */
function notifyUpcomingDeadlines() {
  try {
    _notifyUpcomingDeadlinesBody();
  } catch(e) {
    _notifyError("notifyUpcomingDeadlines", e);
  }
}

function _syncCalendarToSheetsBody() {
  const sheet = _getOrCreateSheet();

  const now   = new Date();
  const start = new Date(now.getTime() - CONFIG.SYNC_DAYS_BACK  * 24 * 60 * 60 * 1000);
  const end   = new Date(now.getTime() + CONFIG.SYNC_DAYS_AHEAD * 24 * 60 * 60 * 1000);

  // 担当者マスタの全カレンダーからイベントを収集
  const calendarMap = _getAssigneeCalendarMap();
  if (calendarMap.size === 0) {
    throw new Error("担当者マスタが空、またはシートが見つかりません。Calendar→Sheets同期を中断しました。");
  }
  const uniqueCalIds = [...new Set(calendarMap.values())];

  // calendarId → 担当者名 の逆引きマップ（タイトルに @名前 がない場合のフォールバック）
  const calIdToAssignee = new Map();
  calendarMap.forEach(function(calId, name) {
    if (!calIdToAssignee.has(calId)) calIdToAssignee.set(calId, name);
  });

  // イベントを収集しつつ取得元カレンダーIDを保持
  const events = []; // { event, sourceCalId }
  uniqueCalIds.forEach(function(calId) {
    const cal = CalendarApp.getCalendarById(calId);
    if (cal) cal.getEvents(start, end).forEach(function(ev) {
      events.push({ event: ev, sourceCalId: calId });
    });
  });

  const col      = CONFIG.COL;
  const NUM_COLS = Object.keys(CONFIG.COL).length;
  const lastRow  = sheet.getLastRow();
  const syncTime = new Date();

  // シート全データを一括読み込み（ヘッダー除く）
  const sheetRows = lastRow >= CONFIG.DATA_START_ROW
    ? sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.HEADER_ROW, NUM_COLS).getValues()
    : [];

  // EventID → { rowNum, data } のマップ構築（_removeDeletedEvents 用に eventIdToRow も同時に作る）
  const eventIdToEntry = new Map();
  const eventIdToRow   = new Map();
  sheetRows.forEach(function(rowData, i) {
    const id = rowData[col.EVENT_ID - 1];
    if (id) {
      const rowNum = CONFIG.DATA_START_ROW + i;
      eventIdToEntry.set(id, { rowNum: rowNum, data: rowData });
      eventIdToRow.set(id, rowNum);
    }
  });

  // カレンダーイベント色 → ステータス の逆引きマップ
  const colorToStatus = new Map();
  Object.keys(CONFIG.STATUS_COLORS).forEach(function(s) {
    colorToStatus.set(CONFIG.STATUS_COLORS[s], s);
  });

  let addedCount   = 0;
  let updatedCount = 0;
  const pendingUpdates = []; // { rowNum, values }[]
  const pendingNewRows = []; // values[][]

  events.forEach(function(entry) {
    const event       = entry.event;
    const sourceCalId = entry.sourceCalId;
    const eventId     = event.getId();
    const newTaskName = _sanitizeForSheet(event.getTitle());
    // 終日イベントは開始日時を空欄のまま保つ（Sheets→Calendar側の「開始日時が空＝終日タスク」という表現を維持する）
    const newStart    = event.isAllDayEvent() ? "" : _formatDateTime(event.getStartTime());
    const newEnd      = event.isAllDayEvent() ? "" : _formatDateTime(event.getEndTime());
    const assignee    = calIdToAssignee.get(sourceCalId) || "";
    const memo        = _sanitizeForSheet((event.getDescription() || "").trim());
    const creator     = event.getCreators()[0] || "Calendar";
    // イベント色がSTATUS_COLORSに対応している場合のみステータスを更新
    const eventColor = String(event.getColor());
    Logger.log("色確認: [" + event.getTitle() + "] color=" + eventColor + " → " + (colorToStatus.get(eventColor) || "(未対応)"));
    const newStatus   = colorToStatus.get(eventColor) || null;

    if (eventIdToEntry.has(eventId)) {
      // ---- 既存行：差分チェック（読み取り済みデータを使用） ----
      const entry = eventIdToEntry.get(eventId);
      const d     = entry.data;

      const statusChanged = newStatus !== null && String(d[col.STATUS - 1]) !== newStatus;
      const changed =
        statusChanged                                          ||
        String(d[col.TASK_NAME - 1])          !== newTaskName ||
        _formatDateTime(d[col.START - 1])      !== newStart    ||
        _formatDateTime(d[col.END - 1])        !== newEnd      ||
        String(d[col.ASSIGNEE - 1])            !== assignee    ||
        String(d[col.MEMO - 1])                !== memo;

      if (changed) {
        const updated = d.slice(); // 既存値をコピーし変更列だけ上書き
        updated[col.TASK_NAME - 1]  = newTaskName;
        if (newStatus !== null) updated[col.STATUS - 1] = newStatus;
        updated[col.START - 1]      = newStart;
        updated[col.END - 1]        = newEnd;
        updated[col.ASSIGNEE - 1]   = assignee;
        updated[col.MEMO - 1]       = memo;
        updated[col.LAST_SYNC - 1]  = syncTime;
        updated[col.UPDATED_BY - 1] = creator;
        pendingUpdates.push({ rowNum: entry.rowNum, values: updated });
        Logger.log("更新: row=" + entry.rowNum + " [" + newTaskName + "]" + (statusChanged ? " ステータス→" + newStatus : ""));
        updatedCount++;
      }
    } else {
      // ---- 新規行 ----
      const newRow = new Array(NUM_COLS).fill("");
      newRow[col.TASK_NAME - 1]  = newTaskName;
      newRow[col.START - 1]      = newStart;
      newRow[col.END - 1]        = newEnd;
      newRow[col.STATUS - 1]     = newStatus || "未着手";
      newRow[col.ASSIGNEE - 1]   = assignee;
      newRow[col.MEMO - 1]       = memo;
      newRow[col.EVENT_ID - 1]   = eventId;
      newRow[col.LAST_SYNC - 1]  = syncTime;
      newRow[col.CREATED_BY - 1] = creator;
      newRow[col.UPDATED_BY - 1] = creator;
      pendingNewRows.push(newRow);
      addedCount++;
    }
  });

  // バッチ書き込み：更新行（行ごとに1回 setValues）
  pendingUpdates.forEach(function(u) {
    sheet.getRange(u.rowNum, 1, 1, NUM_COLS).setValues([u.values]);
  });

  // バッチ書き込み：新規行（連続範囲に一括 setValues）
  if (pendingNewRows.length > 0) {
    const appendStart = sheet.getLastRow() + 1;
    sheet.getRange(appendStart, 1, pendingNewRows.length, NUM_COLS).setValues(pendingNewRows);
  }

  Logger.log("Calendar -> Sheets 同期完了: " + addedCount + "件追加 / " + updatedCount + "件更新");

  // カレンダーから削除されたイベントをSheetsからも削除
  _removeDeletedEvents(sheet, events, eventIdToRow);
}

/**
 * カレンダー上に存在しなくなったイベントのSheets行を削除する
 * ※ 同期取得範囲（SYNC_DAYS_BACK〜SYNC_DAYS_AHEAD）外の行は判定対象外
 *   （範囲外はそもそもカレンダーから取得していないため、削除判断ができない）
 */
function _removeDeletedEvents(sheet, calendarEvents, eventIdToRow) {
  if (eventIdToRow.size === 0) return;

  const now       = new Date();
  const syncStart = new Date(now.getTime() - CONFIG.SYNC_DAYS_BACK  * 24 * 60 * 60 * 1000);
  const syncEnd   = new Date(now.getTime() + CONFIG.SYNC_DAYS_AHEAD * 24 * 60 * 60 * 1000);

  const activeIds = new Set(calendarEvents.map(function(e) { return e.event.getId(); }));

  const rowsToDelete = [];
  eventIdToRow.forEach(function(row, eventId) {
    if (activeIds.has(eventId)) return;

    // 同期範囲外の行は「カレンダーから消えた」か判断できないのでスキップ
    const startVal  = sheet.getRange(row, CONFIG.COL.START).getDisplayValue();
    const startDate = _sheetDateToCalendarDate(startVal);
    if (!startDate || startDate < syncStart || startDate > syncEnd) return;

    rowsToDelete.push(row);
  });

  // 下の行から削除（行番号のずれを防ぐため降順）
  rowsToDelete.sort(function(a, b) { return b - a; });
  rowsToDelete.forEach(function(row) {
    sheet.deleteRow(row);
    Logger.log("削除: row=" + row);
  });

  if (rowsToDelete.length > 0) {
    Logger.log("Calendar -> Sheets 削除: " + rowsToDelete.length + "件");
  }
}

// ==================== 内部関数 ====================

/**
 * 指定行のタスクをCalendarに反映（作成 or 更新 or 削除）
 * 担当者マスタに登録されていない担当者の場合はスキップ。
 */
function _syncRowToCalendar(sheet, row) {
  const col      = CONFIG.COL;
  const taskName = sheet.getRange(row, col.TASK_NAME).getValue();
  const startVal = sheet.getRange(row, col.START).getDisplayValue();
  const endVal   = sheet.getRange(row, col.END).getDisplayValue();
  let   status   = sheet.getRange(row, col.STATUS).getValue();
  const assignee = String(sheet.getRange(row, col.ASSIGNEE).getValue()).trim();
  const memo     = sheet.getRange(row, col.MEMO).getValue();
  const eventId  = sheet.getRange(row, col.EVENT_ID).getValue();

  const calendarMap = _getAssigneeCalendarMap();

  // タスク名が空なら削除扱い（担当者の状態に関わらず、全カレンダーから検索して削除）
  if (!taskName) {
    if (eventId) {
      try {
        const found = _findEventInCalendars(eventId, [...calendarMap.values()]);
        if (found) found.event.deleteEvent();
      } catch(e) {}
      sheet.getRange(row, col.EVENT_ID).clearContent();
    }
    return;
  }

  // 担当者マスタから対象カレンダーを決定（未登録はスキップ。削除はしない）
  if (!assignee || !calendarMap.has(assignee)) {
    Logger.log("スキップ: row=" + row + " 担当者「" + assignee + "」はマスタに未登録");
    return;
  }

  const targetCal = CalendarApp.getCalendarById(calendarMap.get(assignee));
  if (!targetCal) {
    Logger.log("カレンダーが見つかりません (担当者: " + assignee + ")");
    return;
  }

  if (!status) {
    status = "未着手";
    sheet.getRange(row, col.STATUS).setValue(status);
  }

  // 開始日時が空なら終日タスクとして登録する（日付はトリガー発火日）
  const isAllDay = !startVal;
  let startDate;
  let endDate = null;

  if (isAllDay) {
    startDate = new Date();
  } else {
    // getDisplayValue() の文字列をそのままパースするため TZ 変換は発生しない
    startDate = _sheetDateToCalendarDate(startVal);
    if (!startDate) {
      sheet.getRange(row, col.LAST_SYNC).setValue("⚠ 日時フォーマット不正");
      Logger.log("バリデーションエラー: row=" + row + " 開始日時のフォーマットが不正");
      return;
    }

    endDate = endVal
      ? _sheetDateToCalendarDate(endVal)
      : new Date(startDate.getTime() + 60 * 60 * 1000);

    if (endVal && !endDate) {
      sheet.getRange(row, col.LAST_SYNC).setValue("⚠ 日時フォーマット不正");
      Logger.log("バリデーションエラー: row=" + row + " 終了日時のフォーマットが不正");
      return;
    }

    if (endDate && endDate <= startDate) {
      sheet.getRange(row, col.LAST_SYNC).setValue("⚠ 終了 < 開始");
      Logger.log("バリデーションエラー: row=" + row + " 終了日時が開始日時より前");
      return;
    }
  }

  const title       = taskName;
  const description = memo || "";

  try {
    const operator  = Session.getActiveUser().getEmail();
    const allCalIds = [...calendarMap.values()];

    if (eventId) {
      // まず target calendar で検索
      let ev = null;
      try { ev = targetCal.getEventById(eventId); } catch(ignore) {}

      if (ev) {
        // 同じカレンダー → 更新
        ev.setTitle(title);
        if (isAllDay) {
          ev.setAllDayDate(startDate);
        } else {
          ev.setTime(startDate, endDate);
        }
        ev.setDescription(description);
        const color = CONFIG.STATUS_COLORS[status];
        if (color) ev.setColor(color);
        sheet.getRange(row, col.LAST_SYNC).setValue(new Date());
        sheet.getRange(row, col.UPDATED_BY).setValue(operator);
        Logger.log("更新: " + title + " by " + operator);
        return;
      }

      // 別カレンダーに存在する（担当者変更）→ 旧イベントを削除して再作成
      const found = _findEventInCalendars(eventId, allCalIds);
      if (found) {
        try { found.event.deleteEvent(); } catch(ignore) {}
        Logger.log("担当者変更により旧カレンダーからイベント削除");
      }
    }

    // 新規作成
    const newEvent = isAllDay
      ? targetCal.createAllDayEvent(title, startDate, { description: description })
      : targetCal.createEvent(title, startDate, endDate, { description: description });
    const color = CONFIG.STATUS_COLORS[status];
    if (color) newEvent.setColor(color);
    sheet.getRange(row, col.EVENT_ID).setValue(newEvent.getId());
    sheet.getRange(row, col.LAST_SYNC).setValue(new Date());
    sheet.getRange(row, col.CREATED_BY).setValue(operator);
    sheet.getRange(row, col.UPDATED_BY).setValue(operator);
    Logger.log("作成: " + title + " by " + operator);

  } catch(e) {
    _notifyError("_syncRowToCalendar row=" + row, e);
  }
}

/**
 * エラーをログ出力し、設定があればメール通知する
 */
function _notifyError(context, err) {
  const msg = context + ": " + err.message;
  Logger.log(msg);
  if (!CONFIG.ERROR_NOTIFY_EMAIL) return;
  try {
    MailApp.sendEmail(CONFIG.ERROR_NOTIFY_EMAIL, "[TaskSync エラー] " + context, msg);
  } catch(mailErr) {
    Logger.log("メール送信失敗: " + mailErr.message);
  }
}

/**
 * 期日が近いタスクをGoogle Chatに通知する（担当者をメンション）
 * Webhook URLはスクリプトプロパティ "CHAT_WEBHOOK_URL" から取得する（ソースコードには書かない）
 */
function _notifyUpcomingDeadlinesBody() {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty("CHAT_WEBHOOK_URL");
  if (!webhookUrl) {
    Logger.log("CHAT_WEBHOOK_URL未設定のため期日通知をスキップ");
    return;
  }

  const sheet   = _getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return;

  const col      = CONFIG.COL;
  const NUM_COLS = Object.keys(col).length;
  const rows     = sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.HEADER_ROW, NUM_COLS).getValues();

  const chatUserMap = _getAssigneeChatUserMap();

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const limit = new Date(today.getTime() + CONFIG.CHAT_NOTIFY_DAYS_BEFORE * 24 * 60 * 60 * 1000);

  rows.forEach(function(rowData) {
    const status = String(rowData[col.STATUS - 1]);
    if (status === "完了" || status === "キャンセル") return;

    const deadlineVal = rowData[col.DEADLINE - 1];
    if (!deadlineVal) return;
    const deadline = (deadlineVal instanceof Date) ? new Date(deadlineVal) : new Date(String(deadlineVal));
    if (isNaN(deadline.getTime())) return;
    deadline.setHours(0, 0, 0, 0);

    if (deadline < today || deadline > limit) return;

    const taskName    = rowData[col.TASK_NAME - 1];
    const assignee    = String(rowData[col.ASSIGNEE - 1]).trim();
    const chatUserId  = chatUserMap.get(assignee);
    const mention     = chatUserId ? "<" + chatUserId + ">" : (assignee || "(担当者未設定)");
    const deadlineStr = _formatDateTime(deadline).split(" ")[0];

    const text = "⏰ " + mention + " タスク「" + taskName + "」の期日が近づいています（期日: " + deadlineStr + "）";

    try {
      UrlFetchApp.fetch(webhookUrl, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({ text: text }),
      });
    } catch(e) {
      Logger.log("Chat通知送信失敗: " + taskName + " " + e.message);
    }
  });
}

/**
 * ヘッダー行を作成
 */
function _createHeader() {
  const sheet = _getOrCreateSheet();
  const headers = ["タスク名", "開始日時", "終了日時", "ステータス", "担当者", "メモ/説明", "Event ID (自動)", "最終同期 (自動)", "作成者 (自動)", "最終更新者 (自動)"];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);

  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4A90D9");
  headerRange.setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 200); // タスク名
  sheet.setColumnWidth(2, 160); // 開始日時
  sheet.setColumnWidth(3, 160); // 終了日時
  sheet.setColumnWidth(4, 100); // ステータス
  sheet.setColumnWidth(5, 120); // 担当者
  sheet.setColumnWidth(6, 220); // メモ
  sheet.setColumnWidth(7, 220); // Event ID
  sheet.setColumnWidth(8, 160); // 最終同期
  sheet.setColumnWidth(9, 180); // 作成者
  sheet.setColumnWidth(10, 180); // 最終更新者

  // ステータスのドロップダウン
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["未着手", "進行中", "完了", "キャンセル"], true)
    .build();
  sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.COL.STATUS, 1000, 1).setDataValidation(statusRule);

  SpreadsheetApp.flush();
}

/**
 * トリガーを登録（重複防止あり）
 */
function _registerTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (["onEditTrigger", "syncCalendarToSheets", "notifyUpcomingDeadlines"].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("onEditTrigger")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  ScriptApp.newTrigger("syncCalendarToSheets")
    .timeBased()
    .everyMinutes(10)
    .create();

  ScriptApp.newTrigger("notifyUpcomingDeadlines")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
}

/**
 * シートを取得（なければ作成）
 */
function _getOrCreateSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }
  return sheet;
}

/**
 * TZ変換を一切挟まないのでズレが起きない。
 */
function _sheetDateToCalendarDate(val) {
  if (!val) return null;
  const s = String(val);
  const parts = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})[\s](\d{1,2}):(\d{1,2})/);
  if (!parts) return null;
  return new Date(
    parseInt(parts[1]),
    parseInt(parts[2]) - 1,
    parseInt(parts[3]),
    parseInt(parts[4]),
    parseInt(parts[5])
  );
}

/**
 * 担当者マスタシート（GID: 881583119）を読み込み、担当者名 → カレンダーID の Map を返す
 */
function _getAssigneeCalendarMap() {
  const sheet = SpreadsheetApp.getActive().getSheets()
    .find(function(s) { return s.getSheetId() === 881583119; });
  if (!sheet || sheet.getLastRow() < 2) return new Map();

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const map  = new Map();
  data.forEach(function(row) {
    const name  = String(row[0]).trim();
    const calId = String(row[1]).trim();
    if (name && calId) map.set(name, calId);
  });
  return map;
}

/**
 * 担当者マスタシート（GID: 881583119）のC列を読み込み、担当者名 → ChatユーザーID（"users/xxxx"形式）の Map を返す
 * C列が未入力の担当者はMapに含まれない（呼び出し側で担当者名へのフォールバックが必要）
 */
function _getAssigneeChatUserMap() {
  const sheet = SpreadsheetApp.getActive().getSheets()
    .find(function(s) { return s.getSheetId() === 881583119; });
  if (!sheet || sheet.getLastRow() < 2) return new Map();

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const map  = new Map();
  data.forEach(function(row) {
    const name       = String(row[0]).trim();
    const chatUserId = String(row[2] || "").trim();
    if (name && chatUserId) map.set(name, chatUserId);
  });
  return map;
}

/**
 * 複数カレンダーIDからイベントIDを検索し、最初に見つかった { calendar, event } を返す
 * 見つからなければ null
 */
function _findEventInCalendars(eventId, calendarIds) {
  for (var i = 0; i < calendarIds.length; i++) {
    var cal = CalendarApp.getCalendarById(calendarIds[i]);
    if (!cal) continue;
    try {
      var ev = cal.getEventById(eventId);
      if (ev) return { calendar: cal, event: ev };
    } catch(ignore) {}
  }
  return null;
}

/**
 * Sheetsのセル値（Dateオブジェクト）と文字列の両方を受け付ける
 */
function _formatDateTime(date) {
  if (!date) return "";
  // 文字列で渡された場合はDateに変換（Sheets内部値の再変換は行わない）
  const d = (date instanceof Date) ? date : new Date(date);
  if (isNaN(d.getTime())) return "";
  const y  = d.getFullYear();
  const mo = String(d.getMonth() + 1).padStart(2, "0");
  const dy = String(d.getDate()).padStart(2, "0");
  const h  = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  return y + "/" + mo + "/" + dy + " " + h + ":" + mi;
}

/**
 * Calendar由来の文字列がSheets上で数式として評価されないよう、
 * 先頭が =+-@ の場合は ' を付与してテキスト扱いにする（数式インジェクション対策）
 */
function _sanitizeForSheet(value) {
  if (typeof value !== "string") return value;
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}