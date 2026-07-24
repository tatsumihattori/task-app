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
    PROJECT:     11,  // K: プロジェクト名（Calendarイベント説明欄に表示）
    DEADLINE:    12,  // L: 期日（Google Chat期日通知・Calendarイベント説明欄に使用）
    CLIENT:      13,  // M: 取引先（Calendarイベント説明欄に表示）
  },

  HEADER_ROW: 1,
  DATA_START_ROW: 2,
};

// ==================== メインエントリ ====================

/**
 * スプレッドシートを開いたときに独自メニューを追加する（simple trigger）
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("TaskSync")
    .addItem("Calendar → Sheets を今すぐ同期", "manualSyncCalendarToSheets")
    .addToUi();
}

/**
 * 初回セットアップ：ヘッダー作成 & トリガー登録
 * ※ 一度だけ手動で実行してください
 */
function setup() {
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

      const preferredCalId = (assignee && calendarMap.has(assignee)) ? calendarMap.get(assignee) : null;
      _deleteEventAcrossCalendars(eventId, allCalIds, preferredCalId);

      sheet.getRange(CONFIG.DATA_START_ROW + i, col.EVENT_ID).clearContent();
      deletedCount++;
    });

    // Step 2: シートの全行を再同期
    let createdCount = 0;
    const labelCache = new Map();
    for (let r = CONFIG.DATA_START_ROW; r <= lastRow; r++) {
      _syncRowToCalendar(sheet, r, labelCache, calendarMap);
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
    const labelCache  = new Map();
    const calendarMap = _getAssigneeCalendarMap();
    for (let i = 0; i < numRows; i++) {
      const row = startRow + i;
      if (row <= CONFIG.HEADER_ROW) continue;
      _syncRowToCalendar(sheet, row, labelCache, calendarMap);
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
 * メニューからの手動実行：CalendarからSheetsへの同期を即時実行する
 * （10分ごとの定期実行を待たずに反映を確認したい場合用）
 */
function manualSyncCalendarToSheets() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    ui.alert("別の同期処理が実行中です。しばらく待ってから再試行してください。");
    return;
  }
  try {
    _syncCalendarToSheetsBody();
    ui.alert("✅ Calendar → Sheets の同期が完了しました。");
  } catch (e) {
    _notifyError("manualSyncCalendarToSheets", e);
    ui.alert("エラーが発生しました: " + e.message);
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

  // ラベル名の妥当性チェック用（ラベル名はSTATUS_COLORSのキーと完全一致する運用が前提。CLAUDE.md参照）
  const validStatuses  = new Set(Object.keys(CONFIG.STATUS_COLORS));
  // calendarId → (labelId → labelName) キャッシュ（同一実行内で同じカレンダーのラベル情報を再取得しない）
  const labelNameCache = new Map();

  let addedCount   = 0;
  let updatedCount = 0;
  const pendingUpdates = []; // { rowNum, values }[]
  const pendingNewRows = []; // values[][]
  const eventErrors    = []; // { eventId, message }[]（イベント単位の処理失敗を集約し、ループ後にまとめて1回だけ通知する）

  events.forEach(function(entry) {
    const event       = entry.event;
    const sourceCalId = entry.sourceCalId;

    let eventId;
    try {
      eventId = event.getId();
    } catch (idErr) {
      Logger.log("イベントID取得失敗のためスキップ: " + idErr.message);
      eventErrors.push({ eventId: "(不明)", message: idErr.message });
      return;
    }

    // 対象イベントが同一実行中にCalendar側で削除・変更されると、CalendarAppのイベントオブジェクトへの
    // アクセス（getTitle/getStartTime/getColor等）やAdvanced Service呼び出しが例外を投げることがある。
    // 1件の失敗でforEach全体・ひいては後続の _removeDeletedEvents 呼び出しまで巻き込んで中断させないよう、
    // イベント単位でtry/catchし、失敗したイベントだけスキップして次へ進む。
    // 既存行は前回値を維持、新規イベントは次回同期で再試行される（いずれも安全側）。
    try {
      const rawTitle    = event.getTitle();
      const newTaskName = _sanitizeForSheet(rawTitle);
      // 終日イベントは開始日時を空欄のまま保つ（Sheets→Calendar側の「開始日時が空＝終日タスク」という表現を維持する）
      const newStart    = event.isAllDayEvent() ? "" : _formatDateTime(event.getStartTime());
      const newEnd      = event.isAllDayEvent() ? "" : _formatDateTime(event.getEndTime());
      const assignee    = calIdToAssignee.get(sourceCalId) || "";
      const description = event.getDescription() || "";
      const memo        = _sanitizeForSheet(_extractMemoFromDescription(description).trim());
      const creator     = event.getCreators()[0] || "Calendar";
      // プロジェクト/取引先/期日: 区切り文字が見つかった場合のみヘッダーブロックとして解析する。
      // 見つからない場合（区切り文字追加前の既存イベント、または手動でメモのみに書き換えられた場合）は
      // null を返し、K/L/M列への反映をスキップする（誤って空欄化しないための後方互換フォールバック）。
      const headerFields   = _extractProjectClientDeadlineFromDescription(description);
      const newProject     = headerFields ? _sanitizeForSheet(headerFields.project) : null;
      const newClient      = headerFields ? _sanitizeForSheet(headerFields.client)  : null;
      const newDeadlineStr = headerFields ? headerFields.deadline : null; // "" または "YYYY/MM/DD"
      // ステータス判定: まずCalendarApp（追加API呼び出し不要）の色で判定を試みる。
      // 判定できない場合のみ Advanced Service で colorId と eventLabelId をまとめて取得する。
      // eventLabelId はCalendarの「ラベル」機能に対応し、UIでの手動変更を確実に反映するため colorId より優先する
      // （実機検証で colorId は手動変更後に欠落・不一致することがあったが eventLabelId は常に反映されていた）。
      let newStatus  = colorToStatus.get(String(event.getColor())) || null;
      let resolvedBy = newStatus ? "CalendarApp色" : null;

      if (newStatus === null) {
        const adv = _getEventColorAndLabelViaAdvancedService(sourceCalId, eventId);
        if (adv.eventLabelId) {
          const labelName = _getCalendarLabelMaps(sourceCalId, labelNameCache).idToName.get(adv.eventLabelId);
          if (labelName && validStatuses.has(labelName)) {
            newStatus  = labelName;
            resolvedBy = "ラベル(" + labelName + ")";
          }
        }
        if (newStatus === null && adv.colorId && colorToStatus.has(adv.colorId)) {
          newStatus  = colorToStatus.get(adv.colorId);
          resolvedBy = "colorId(" + adv.colorId + ")";
        }
      }

      Logger.log("色確認: [" + rawTitle + "] " + (resolvedBy ? resolvedBy + " → " + newStatus : "(未対応)"));

      if (eventIdToEntry.has(eventId)) {
        // ---- 既存行：差分チェック（読み取り済みデータを使用） ----
        const entry = eventIdToEntry.get(eventId);
        const d     = entry.data;

        const statusChanged   = newStatus !== null && String(d[col.STATUS - 1]) !== newStatus;
        const projectChanged  = newProject     !== null && String(d[col.PROJECT - 1])     !== newProject;
        const clientChanged   = newClient      !== null && String(d[col.CLIENT - 1])      !== newClient;
        const deadlineChanged = newDeadlineStr !== null && _formatDateOnly(d[col.DEADLINE - 1]) !== newDeadlineStr;
        const changed =
          statusChanged                                          ||
          projectChanged                                         ||
          clientChanged                                          ||
          deadlineChanged                                        ||
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
          if (newProject !== null) updated[col.PROJECT - 1] = newProject;
          if (newClient  !== null) updated[col.CLIENT  - 1] = newClient;
          if (newDeadlineStr !== null) updated[col.DEADLINE - 1] = newDeadlineStr ? _parseDateOnly(newDeadlineStr) : "";
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
        if (newProject) newRow[col.PROJECT - 1]  = newProject;
        if (newClient)  newRow[col.CLIENT - 1]   = newClient;
        if (newDeadlineStr) newRow[col.DEADLINE - 1] = _parseDateOnly(newDeadlineStr);
        pendingNewRows.push(newRow);
        addedCount++;
      }
    } catch (e) {
      Logger.log("イベント処理失敗のためスキップ: eventId=" + eventId + " " + e.message);
      eventErrors.push({ eventId: eventId, message: e.message });
    }
  });

  if (eventErrors.length > 0) {
    const summary = eventErrors.length + "件のイベント処理でエラー: " +
      eventErrors.slice(0, 5).map(function(er) { return er.eventId + "(" + er.message + ")"; }).join(" / ") +
      (eventErrors.length > 5 ? " ...他" + (eventErrors.length - 5) + "件" : "");
    Logger.log(summary);
    _notifyError("_syncCalendarToSheetsBody（イベント単位エラー）", new Error(summary));
  }

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
  // start/end をそのまま渡すことで、上のイベント取得範囲と削除判定範囲を同一のものにする
  // （forEachループの実行時間が伸びても、範囲がズレないようにするため）
  _removeDeletedEvents(sheet, events, eventIdToRow, uniqueCalIds, start, end);
}

/**
 * カレンダー上に存在しなくなったイベントのSheets行を削除する
 * ※ 同期取得範囲（SYNC_DAYS_BACK〜SYNC_DAYS_AHEAD）外の行は、日付だけでは「カレンダーから消えた」のか
 *   「範囲外なだけ」なのか判断できない。開始日時（B列）が読める行はそのレンジ判定でスキップする。
 *   終日タスク（B列が仕様上つねに空欄）やフォーマット不正の行はレンジ判定ができないため、
 *   代わりに calendarIds 全件を対象に getEventById による実在確認（日付に縛られない）にフォールバックする。
 * syncStart/syncEnd は呼び出し元（_syncCalendarToSheetsBody）のイベント取得範囲と同一のものを渡すこと。
 */
function _removeDeletedEvents(sheet, calendarEvents, eventIdToRow, calendarIds, syncStart, syncEnd) {
  if (eventIdToRow.size === 0) return;

  const activeIds = new Set(calendarEvents.map(function(e) { return e.event.getId(); }));

  const rowsToDelete = [];
  eventIdToRow.forEach(function(row, eventId) {
    if (activeIds.has(eventId)) return;

    const startVal  = sheet.getRange(row, CONFIG.COL.START).getDisplayValue();
    const startDate = _sheetDateToCalendarDate(startVal);

    if (!startDate) {
      // 終日タスク／フォーマット不正行：日付で範囲内外を判断できないので直接検索で実在確認する。
      // CalendarApp.getEventById() は削除済み（status:"cancelled"のトゥームストーン）イベントでも
      // 非nullを返すことがあるため、見つかった場合でも Advanced Service の status で真偽を確認する。
      const found = _findEventInCalendars(eventId, calendarIds);
      if (found && _isEventActiveViaAdvancedService(found.calendar.getId(), eventId)) return; // 有効なイベントとして存在する → 削除しない
      rowsToDelete.push(row);
      return;
    }

    // 同期範囲外の行は「カレンダーから消えた」か判断できないのでスキップ
    if (startDate < syncStart || startDate > syncEnd) return;

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
 * calendarMap（担当者名→カレンダーID）は呼び出し元（onEditTrigger/resetAndResyncToCalendar）が
 * ループの外で1回だけ取得したものを受け取る（行ごとにマスタシートを再読み込みしないため）。
 */
function _syncRowToCalendar(sheet, row, labelCache, calendarMap) {
  const col      = CONFIG.COL;
  const taskName = sheet.getRange(row, col.TASK_NAME).getValue();
  const startVal = sheet.getRange(row, col.START).getDisplayValue();
  const endVal   = sheet.getRange(row, col.END).getDisplayValue();
  let   status   = String(sheet.getRange(row, col.STATUS).getValue()).trim();
  const assignee = String(sheet.getRange(row, col.ASSIGNEE).getValue()).trim();
  const memo     = sheet.getRange(row, col.MEMO).getValue();
  const project  = sheet.getRange(row, col.PROJECT).getValue();
  const deadline = sheet.getRange(row, col.DEADLINE).getValue();
  const client   = sheet.getRange(row, col.CLIENT).getValue();
  const eventId  = sheet.getRange(row, col.EVENT_ID).getValue();

  // タスク名が空なら削除扱い（担当者の状態に関わらず、全カレンダーから検索して削除）
  if (!taskName) {
    if (eventId) {
      const preferredCalId = (assignee && calendarMap.has(assignee)) ? calendarMap.get(assignee) : null;
      _deleteEventAcrossCalendars(eventId, [...calendarMap.values()], preferredCalId);
      sheet.getRange(row, col.EVENT_ID).clearContent();
    }
    return;
  }

  // 担当者マスタから対象カレンダーを決定（未登録はスキップ。削除はしない）
  if (!assignee || !calendarMap.has(assignee)) {
    Logger.log("スキップ: row=" + row + " 担当者「" + assignee + "」はマスタに未登録");
    return;
  }

  const targetCalId = calendarMap.get(assignee);
  const targetCal   = CalendarApp.getCalendarById(targetCalId);
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
  const description = _buildEventDescription(project, client, deadline, memo);

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
        _applyStatusColorAndLabel(ev, eventId, targetCalId, status, row, labelCache);
        sheet.getRange(row, col.LAST_SYNC).setValue(new Date());
        sheet.getRange(row, col.UPDATED_BY).setValue(operator);
        Logger.log("更新: " + title + " by " + operator);
        return;
      }

      // 別カレンダーに存在する（担当者変更）→ 旧イベントを削除して再作成
      if (_deleteEventAcrossCalendars(eventId, allCalIds, null)) {
        Logger.log("担当者変更により旧カレンダーからイベント削除");
      }
    }

    // 新規作成
    const newEvent = isAllDay
      ? targetCal.createAllDayEvent(title, startDate, { description: description })
      : targetCal.createEvent(title, startDate, endDate, { description: description });
    const newEventId = newEvent.getId();
    _applyStatusColorAndLabel(newEvent, newEventId, targetCalId, status, row, labelCache);
    sheet.getRange(row, col.EVENT_ID).setValue(newEventId);
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

  // 担当者ごとにタスクをまとめてから1人1メッセージで送る
  const tasksByAssignee = new Map(); // assignee(表示名) -> [{ taskName, deadline, deadlineStr }]

  rows.forEach(function(rowData) {
    const status = String(rowData[col.STATUS - 1]);
    if (status === "完了" || status === "キャンセル") return;

    const deadlineVal = rowData[col.DEADLINE - 1];
    if (!deadlineVal) return;
    const deadline = (deadlineVal instanceof Date) ? new Date(deadlineVal) : new Date(String(deadlineVal));
    if (isNaN(deadline.getTime())) return;
    deadline.setHours(0, 0, 0, 0);

    if (deadline < today || deadline > limit) return;

    const taskName = rowData[col.TASK_NAME - 1];
    const assignee = String(rowData[col.ASSIGNEE - 1]).trim() || "(担当者未設定)";

    if (!tasksByAssignee.has(assignee)) tasksByAssignee.set(assignee, []);
    tasksByAssignee.get(assignee).push({
      taskName: taskName,
      deadline: deadline,
      deadlineStr: _formatDateOnly(deadline),
    });
  });

  tasksByAssignee.forEach(function(tasks, assignee) {
    tasks.sort(function(a, b) { return a.deadline - b.deadline; });

    const chatUserId = chatUserMap.get(assignee);
    const mention     = chatUserId ? "<" + chatUserId + ">" : assignee;
    const taskLines   = tasks.map(function(t) {
      return "・" + t.taskName + "（期日: " + t.deadlineStr + "）";
    }).join("\n");

    const text = "⏰ " + mention + " 期日が近づいているタスクがあります\n" + taskLines;

    try {
      UrlFetchApp.fetch(webhookUrl, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({ text: text }),
      });
    } catch(e) {
      Logger.log("Chat通知送信失敗: " + assignee + " " + e.message);
    }
  });
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
 * 担当者マスタシート（GID: 881583119）のA〜C列（担当者名・カレンダーID・ChatユーザーID）を読み込む。
 * シートが無い/データ行が無い場合は空配列を返す。
 */
function _readAssigneeMasterRows() {
  const sheet = SpreadsheetApp.getActive().getSheets()
    .find(function(s) { return s.getSheetId() === 881583119; });
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
}

/**
 * 担当者マスタシートを読み込み、担当者名 → カレンダーID の Map を返す
 */
function _getAssigneeCalendarMap() {
  const map = new Map();
  _readAssigneeMasterRows().forEach(function(row) {
    const name  = String(row[0]).trim();
    const calId = String(row[1]).trim();
    if (name && calId) map.set(name, calId);
  });
  return map;
}

/**
 * 担当者マスタシートのC列を読み込み、担当者名 → ChatユーザーID（"users/xxxx"形式）の Map を返す
 * C列が未入力の担当者はMapに含まれない（呼び出し側で担当者名へのフォールバックが必要）
 */
function _getAssigneeChatUserMap() {
  const map = new Map();
  _readAssigneeMasterRows().forEach(function(row) {
    const name       = String(row[0]).trim();
    const chatUserId = String(row[2] || "").trim();
    if (name && chatUserId) map.set(name, chatUserId);
  });
  return map;
}

/**
 * CalendarApp由来のEvent ID（"xxxx@google.com"形式）を、Advanced Service (REST API) が
 * 要求する素のID（"@"以降を除いたもの）に変換する。
 */
function _toRawEventId(eventId) {
  return eventId.split("@")[0];
}

/**
 * Calendar Advanced Service (REST API) でイベントの colorId と eventLabelId をまとめて取得する。
 * CalendarApp.getColor() はUIで手動変更した色/ラベルを返さないための回避策。
 * 取得できなかった項目は null を返す（呼び出し側でフォールバックする）。
 */
function _getEventColorAndLabelViaAdvancedService(calendarId, eventId) {
  try {
    const rawEventId = _toRawEventId(eventId);
    const advEvent = Calendar.Events.get(calendarId, rawEventId, { fields: "colorId,eventLabelId" });
    return { colorId: advEvent.colorId || null, eventLabelId: advEvent.eventLabelId || null };
  } catch(e) {
    Logger.log("Advanced Serviceでの色/ラベル取得失敗: " + eventId + " " + e.message);
    return { colorId: null, eventLabelId: null };
  }
}

/**
 * Advanced Service (REST API) でイベントの status を確認し、真に有効なイベントかを判定する。
 * CalendarApp.getEventById() は削除（キャンセル）済みイベントでも一定期間は非nullのオブジェクトを
 * 返すことがあり（Calendar側は削除を status:"cancelled" のトゥームストーンとして保持するため）、
 * 実在確認としては不十分。_removeDeletedEvents の削除判定でのみ、この関数で status を併せて確認する。
 * 判定できない場合（API通信エラー等、cancelledと確定できない場合）は誤削除を避けるため true（有効）を返す。
 */
function _isEventActiveViaAdvancedService(calendarId, eventId) {
  try {
    const rawEventId = _toRawEventId(eventId);
    const advEvent = Calendar.Events.get(calendarId, rawEventId, { fields: "status" });
    return advEvent.status !== "cancelled";
  } catch(e) {
    Logger.log("Advanced Serviceでのstatus確認失敗（安全側に倒し有効扱いとする）: " + eventId + " " + e.message);
    return true;
  }
}

/**
 * Calendar Advanced Service (REST API) でイベントに eventLabelId を直接書き込む。
 * ラベル機能が有効なカレンダーでは CalendarApp.setColor() が色を反映しないための回避策（Issue #35）。
 * 失敗してもログのみ（色付けは見た目の補助のため、タスク作成/更新自体は止めない）。
 */
function _setEventLabel(calendarId, eventId, labelId) {
  try {
    const rawEventId = _toRawEventId(eventId);
    Calendar.Events.patch({ eventLabelId: labelId }, calendarId, rawEventId, { eventLabelVersion: 1 });
  } catch(e) {
    Logger.log("ラベル書き込み失敗: " + eventId + " " + e.message);
  }
}

/**
 * 指定カレンダーの labelProperties.eventLabels を取得し、
 * { idToName: Map(labelId->labelName), nameToId: Map(labelName->labelId) } を返す。
 * cache（calendarId → 結果）で同一実行内の再取得を避ける。取得失敗時は両方空のMapを返す。
 */
function _getCalendarLabelMaps(calendarId, cache) {
  if (cache.has(calendarId)) return cache.get(calendarId);
  const idToName = new Map();
  const nameToId = new Map();
  try {
    const cal    = Calendar.Calendars.get(calendarId, { fields: "labelProperties" });
    const labels = (cal.labelProperties && cal.labelProperties.eventLabels) || [];
    labels.forEach(function(label) {
      idToName.set(label.id, label.name);
      nameToId.set(label.name, label.id);
    });
  } catch(e) {
    Logger.log("ラベル情報取得失敗: " + calendarId + " " + e.message);
  }
  const result = { idToName: idToName, nameToId: nameToId };
  cache.set(calendarId, result);
  return result;
}

/**
 * eventIdのイベントを削除する。preferredCalId（担当者の現在のカレンダー）が分かっていれば
 * まずそこだけを直接検索して削除を試み、見つからなければ calendarIds 全体を
 * _findEventInCalendars でフォールバック検索する（Issue #15：3箇所に散在していた削除ロジックの共通化）。
 * 削除できた場合は true、見つからない/削除失敗の場合は false を返す（例外は投げない）。
 */
function _deleteEventAcrossCalendars(eventId, calendarIds, preferredCalId) {
  try {
    if (preferredCalId) {
      const cal = CalendarApp.getCalendarById(preferredCalId);
      const ev  = cal ? cal.getEventById(eventId) : null;
      if (ev) { ev.deleteEvent(); return true; }
    }
    const found = _findEventInCalendars(eventId, calendarIds);
    if (found) { found.event.deleteEvent(); return true; }
  } catch(e) {
    Logger.log("削除エラー: " + eventId + " " + e.message);
  }
  return false;
}

/**
 * イベントにステータス色（setColor）を設定し、ラベル機能が有効なカレンダーであれば
 * eventLabelId も直接書き込む（Issue #35。setColor()だけではラベル有効カレンダーに反映されないため）。
 * 未対応のステータス文字列の場合は警告ログのみを出し、色設定はスキップする。
 */
function _applyStatusColorAndLabel(event, eventId, calendarId, status, row, labelCache) {
  const color = CONFIG.STATUS_COLORS[status];
  if (color) {
    event.setColor(color);
  } else {
    Logger.log("警告: row=" + row + " 未対応のステータス「" + status + "」のためCalendar色を設定しませんでした");
  }
  const labelId = _getCalendarLabelMaps(calendarId, labelCache).nameToId.get(status);
  if (labelId) _setEventLabel(calendarId, eventId, labelId);
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

// Calendarイベント説明欄で「プロジェクト/取引先/期日」ブロックとメモ本文を区切る文字列
const _DESCRIPTION_MEMO_DELIMITER = "\n---\n";

/**
 * Sheetsのセル値（Dateオブジェクト or 文字列）を "YYYY/MM/DD" 形式の日付のみの文字列に変換する。
 * パースできない場合は元の値をそのまま文字列化して返す（情報を消さない）。
 */
function _formatDateOnly(val) {
  if (!val) return "";
  const d = (val instanceof Date) ? val : new Date(String(val));
  if (isNaN(d.getTime())) return String(val);
  return _formatDateTime(d).split(" ")[0];
}

/**
 * プロジェクト名・取引先・期日・メモからCalendarイベントの説明欄テキストを組み立てる。
 * プロジェクト/取引先/期日がすべて空なら、メモのみ（従来と同じ形式）を返す。
 */
function _buildEventDescription(project, client, deadline, memo) {
  const lines = [];
  if (project) lines.push("プロジェクト: " + project);
  if (client)  lines.push("取引先: " + client);
  const deadlineStr = _formatDateOnly(deadline);
  if (deadlineStr) lines.push("期日: " + deadlineStr);

  if (lines.length === 0) return memo || "";
  return lines.join("\n") + _DESCRIPTION_MEMO_DELIMITER + (memo || "");
}

/**
 * Calendarイベントの説明欄からメモ本文だけを取り出す（`_buildEventDescription` の逆変換）。
 * 区切り文字が見つからない場合は説明欄全体をメモとして扱う（従来の挙動と同じフォールバック）。
 */
function _extractMemoFromDescription(description) {
  const idx = description.indexOf(_DESCRIPTION_MEMO_DELIMITER);
  if (idx === -1) return description;
  return description.slice(idx + _DESCRIPTION_MEMO_DELIMITER.length);
}

/**
 * Calendarイベントの説明欄からプロジェクト名・取引先・期日（文字列）を取り出す
 * （`_buildEventDescription` の逆変換。Issue #46）。
 * 区切り文字が見つからない場合（区切り文字追加前の既存イベント、または手動でメモのみに
 * 書き換えられた場合）は null を返す。呼び出し側はこの場合K/L/M列への反映をスキップし、
 * 誤って値を空欄化しないようにする。
 * 区切り文字が見つかった場合、各行（プロジェクト/取引先/期日）は存在すればその値、
 * 存在しなければ空文字（＝Calendar側でその項目が削除された）を返す。
 */
function _extractProjectClientDeadlineFromDescription(description) {
  const idx = description.indexOf(_DESCRIPTION_MEMO_DELIMITER);
  if (idx === -1) return null;
  const header       = description.slice(0, idx);
  const projectMatch  = header.match(/^プロジェクト:\s?(.*)$/m);
  const clientMatch   = header.match(/^取引先:\s?(.*)$/m);
  const deadlineMatch = header.match(/^期日:\s?(.*)$/m);
  return {
    project:  projectMatch  ? projectMatch[1].trim()  : "",
    client:   clientMatch   ? clientMatch[1].trim()   : "",
    deadline: deadlineMatch ? deadlineMatch[1].trim() : "",
  };
}

/**
 * "YYYY/MM/DD"（または "YYYY-MM-DD"）形式の日付のみの文字列を Date に変換する。
 * `_sheetDateToCalendarDate` の日付のみ版（時刻を含まない）。パースできない場合は null。
 */
function _parseDateOnly(str) {
  const m = String(str).match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (!m) return null;
  return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
}