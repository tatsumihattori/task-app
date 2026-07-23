# TaskSync — Sheets ↔ Calendar 双方向同期 引き継ぎメモ

このファイルは開発の引き継ぎ・コンテキスト復元用メモです。
`TaskSync.js` と一緒に Apps Script プロジェクトのディレクトリに置いてください。

## 概要

Google Sheets と Google Calendar の双方向タスク同期システム。Google Apps Script (GAS) で実装。

- **Sheets → Calendar**: セル編集時に即座反映（`onEdit` インストーラブルトリガー）
- **Calendar → Sheets**: 10分ごとのポーリングで反映（時間ベーストリガー）。スプレッドシートのカスタムメニュー「TaskSync」からも手動即時実行可能（`onOpen()` の simple trigger でメニュー登録、実体は `manualSyncCalendarToSheets()` → `_syncCalendarToSheetsBody()`。Issue #40）
- **期日通知**: 期日が近いタスクをGoogle Chatに毎日通知（時間ベーストリガー、`notifyUpcomingDeadlines`）

## シート構成

| 列 | 内容 | 備考 |
|---|---|---|
| A | タスク名 | 空にするとカレンダー側のイベントを削除する（唯一の削除トリガー） |
| B | 開始日時 | `YYYY/MM/DD HH:mm` 形式（1桁の月・日・時・分も許容）。**空欄の場合は終日タスクとして登録される**（日付はトリガー実行日） |
| C | 終了日時 | 同上。空欄時は開始+1時間（終日タスクの場合は無視される） |
| D | ステータス | 未着手 / 進行中 / 完了 / キャンセル。未入力の場合はSheets→Calendar同期時に「未着手」が自動補完される |
| E | 担当者 | 担当者マスタに登録された名前のみ有効 |
| F | メモ/説明 | カレンダーイベントの説明欄に同期 |
| G | Event ID | 自動管理・編集不可 |
| H | 最終同期日時 | 自動管理。バリデーションエラー時は警告文字列を書き込む |
| I | 作成者 | 自動管理 |
| J | 最終更新者 | 自動管理 |
| K | プロジェクト名 | Calendarイベント説明欄に表示 |
| L | 期日 | Google Chat期日通知に使用（`CHAT_NOTIFY_DAYS_BEFORE`日前から当日まで毎日通知。完了/キャンセルは通知対象外）。Calendarイベント説明欄にも表示 |
| M | 取引先 | Calendarイベント説明欄に表示 |

カレンダー側のイベントタイトル形式: **タスク名のみ**（ステータス・担当者はタイトルに含めない）
説明欄は「プロジェクト/取引先/期日」ブロック（いずれか入力されている場合のみ）＋メモ。詳細は「Calendarイベント説明欄の構成」参照。

## CONFIG（主要設定）

```js
STATUS_COLORS: {
  "未着手":    "2",  // Sage (Pale Green)
  "進行中":    "9",  // Blueberry (Blue)
  "完了":      "8",  // Graphite (Gray)
  "キャンセル": "7",  // Peacock (Cyan)
}
SYNC_DAYS_AHEAD: 90  // Calendar取得範囲：未来
SYNC_DAYS_BACK: 7    // Calendar取得範囲：過去
ERROR_NOTIFY_EMAIL: "tatsumi.hattori@laboratory.work"
CHAT_NOTIFY_DAYS_BEFORE: 3  // 期日の何日前から通知するか（当日まで毎日）
```

COL は A=1 〜 M=13 の13列構成。`Object.keys(CONFIG.COL).length` で `NUM_COLS` を動的計算しているため、列追加時は COL にキーを追加するだけでよい。
※ ヘッダー行の自動生成（旧 `_createHeader()`）は、呼び出しがコメントアウトされたまま10列(A〜J)分の定義しか持たず13列構成と乖離していたため削除した（Issue #11対応）。ヘッダーは手動で管理する。

## 担当者マスタ（GID: 881583119）

A列=担当者名、B列=カレンダーID、C列=ChatユーザーID（`users/xxxxxxxxxxxxxxxxxxx`形式。任意項目）。
`_getAssigneeCalendarMap()` でA・B列をMapに読み込む。マスタ未登録の担当者はSheets→Calendar同期をスキップ（削除は担当者の状態に関わらず行われる。詳細は「運用上の合意事項」参照）。
カレンダーID→担当者名の逆引きマップも同関数内で生成（`calIdToAssignee`）し、Calendar→Sheets同期での担当者列埋めに使用。
C列は `_getAssigneeChatUserMap()` で読み込み、Google Chat期日通知でのメンションに使用（未入力の担当者は名前をテキストで含めるだけでメンションは飛ばない）。
新規担当者のカレンダーには、"未着手"/"進行中"/"完了"/"キャンセル" の4つのラベルをCalendar UIで作成しておくこと。Calendar側での手動ステータス変更の検知（詳細は「解決済みの重要な技術的ハマりどころ」項目6参照）だけでなく、Sheets→Calendar同期時の色表示（項目7参照）にも必要。未作成でも同期自体は壊れず、Calendar→Sheetsは`colorId`ベースの判定に、Sheets→Calendarは`setColor()`のみにフォールバックする。

## Google Chat連携（期日通知）

- `notifyUpcomingDeadlines()`（毎日9時、時間ベーストリガー）→ `_notifyUpcomingDeadlinesBody()` が本体
- Webhook URLは**ソースコードには書かない**。Apps Scriptエディタの「プロジェクトの設定 > スクリプト プロパティ」で `CHAT_WEBHOOK_URL` を設定する（`key`/`token` を含む実質的な認証情報のため）。未設定の場合は通知をスキップするだけでエラーにはならない
- 担当者マスタC列（ChatユーザーID）が設定されていれば `<users/xxxx>` 形式で本文に埋め込み、Chatが実際のメンション（通知）として扱う。未設定なら担当者名をテキストとして含めるだけ
- 通知は**タスク単位ではなく担当者単位**で1メッセージにまとめて送信する（同じ担当者の対象タスクが複数あっても1通）。タスクは期日の近い順に並べる
- 重複通知の抑制は行っていない。期日が通知範囲（`CHAT_NOTIFY_DAYS_BEFORE`日前〜当日）に入っている間は毎日通知される（意図的な仕様。小規模運用のため簡略化）
- ステータスが「完了」「キャンセル」のタスクは通知対象外

## Calendarイベント説明欄の構成

`_buildEventDescription()` (Sheets→Calendar) が説明欄のテキストを組み立てる。プロジェクト名(K)・取引先(M)・期日(L)のうち入力されている項目だけを先頭にまとめ、区切り文字 `_DESCRIPTION_MEMO_DELIMITER`（`"\n---\n"`）を挟んでメモ(F)を続ける。K/L/Mがすべて空の場合は区切り文字を付けず、メモのみ（従来と同じ形式）を出力する。

```
プロジェクト: {K}
取引先: {M}
期日: {L}
---
{メモ本文}
```

Calendar→Sheets同期（`_syncCalendarToSheetsBody`）は `_extractMemoFromDescription()` で区切り文字より後ろだけを抜き出してF列（メモ）に書き戻す。**プロジェクト/取引先/期日の3項目はSheets→Calendarの一方向のみ**（Calendar側で説明欄の先頭ブロックを直接編集しても、K/L/M列には反映されない）。

区切り文字が見つからない場合（区切り文字を追加する前に作られた既存イベント、またはCalendar側で区切り文字自体を消してしまった場合）は、説明欄全体をメモとして扱う後方互換フォールバックになっている。

## ステータスとカレンダーイベント色の対応

**Sheets → Calendar**（`_syncRowToCalendar`）: `CONFIG.STATUS_COLORS[status]` で色を決定し `event.setColor()` で設定する。ただしラベル機能（`labelProperties`）が設定されているカレンダーでは `setColor()` だけでは色が反映されない（`colorId`・`eventLabelId` とも設定されず、Calendar UI上も無色のまま。Issue #35）。そのため `setColor()` に加えて、`_getCalendarLabelMaps()` でステータス名（`status`）と文字列完全一致するラベルIDを引き、見つかれば `_setEventLabel()` で `Calendar.Events.patch()` により `eventLabelId` を直接書き込む（Advanced Service）。ラベルが無い/名前が一致しないカレンダーは従来通り `setColor()` の結果のみに依存する（フォールバック、退行なし）。ラベル書き込み失敗はログに残すのみでタスク作成/更新自体は継続する（詳細は「解決済みの重要な技術的ハマりどころ」項目7参照）。

**Calendar → Sheets**（`_syncCalendarToSheetsBody`）: 3段階でステータスを判定する。

1. `colorToStatus` 逆引きマップを使い、まず `event.getColor()`（CalendarApp）で色を読む。追加API呼び出し不要な高速パス。
2. 判定できない場合のみ、`_getEventColorAndLabelViaAdvancedService()` で Calendar Advanced Service (REST API, `Calendar.Events.get()`) を呼び、`colorId` と `eventLabelId`（Calendarの「ラベル」機能。Issue #32参照）をまとめて取得する。`eventLabelId` が取れた場合、`_getLabelNameMap()` でそのカレンダーの `labelProperties.eventLabels` からラベル名を解決し、ラベル名が `STATUS_COLORS` のキー（"未着手"/"進行中"/"完了"/"キャンセル"）と完全一致すればそれをステータスとして採用する（**colorIdより優先**。理由は「解決済みの重要な技術的ハマりどころ」項目6を参照）。
3. ラベルで判定できない場合のみ、`colorId` と `colorToStatus` で判定する（`CONFIG.STATUS_COLORS` の値 "2"/"9"/"8"/"7" と同じ数値文字列表現のため変換不要）。

いずれでも判定できない場合はステータスを更新しない（既存値を維持）。Advanced Serviceの呼び出しはイベント単位のAPIコールになるため、CalendarApp側で色が判定できる場合（プログラムでsetColorした直後など）はスキップして呼ばないようにしている。ラベル名解決にはカレンダー単位の `labelProperties` 取得が必要だが、同一実行内は `labelNameCache` でカレンダーごとに1回だけ取得する。

**この判定はラベルが存在するカレンダーが前提**：各担当者カレンダーに、"未着手"/"進行中"/"完了"/"キャンセル" という名前のラベルが事前に作成されている必要がある（現在は全担当者カレンダーに設定済み）。ラベルが無いカレンダー・イベントは自動的にcolorId判定にフォールバックするため、退行はしない。

## 解決済みの重要な技術的ハマりどころ

### 1. 日時のタイムゾーンズレ（最重要）

**症状**: Sheetsで10:00と入力したのに、Calendarには19:00で登録される（9時間ズレ＝JSTのオフセット分）。

**原因**: `Range.getValue()` で取得したDateオブジェクトはすでに内部的にUTCシフトしており、そこにさらにタイムゾーン補正を加えると二重にズレる。

**解決策**: `getDisplayValue()` でセルの表示文字列を取得し、正規表現でパースして `new Date(年, 月, 日, 時, 分)` で直接組み立てる（TZ変換を一切挟まない）。実装は `_sheetDateToCalendarDate()` を参照。

→ **今後似た日時バグが出たら、まずこの関数とその呼び出し元を疑うこと。**

正規表現は `\d{1,2}` で1・2桁両方に対応済み（Sheetsのセル書式次第で時刻が `9:30` のように1桁になるため）。

### 2. `CalendarEvent.getColor()` の制約（Advanced Serviceで対応済み）

GASの CalendarApp 経由では、UIで手動変更したイベント色を `getColor()` で取得できない（空文字を返す）。GASがプログラムで `setColor()` した色のみ取得可能。
→ `_getEventColorAndLabelViaAdvancedService()`（旧 `_getEventColorIdViaAdvancedService()`。Issue #32対応でラベル取得も統合）で Calendar Advanced Service の `Calendar.Events.get()` を呼び、実際の `colorId` を取得することで解決済み。`appsscript.json` の `dependencies.enabledAdvancedServices` にCalendar API (v3) を追加してある。詳細は「ステータスとカレンダーイベント色の対応」セクション参照。ただし項目6の通り、`colorId` 自体もCalendarの新UI（ラベル機能）下では不完全になるケースがあり、現在は `eventLabelId` を優先している。

### 3. 担当者変更時のカレンダーイベント移動

担当者が変わるとターゲットカレンダーが変わる。`_syncRowToCalendar` では：

1. 新担当者のカレンダーでEventIDを検索
2. 見つからなければ `_findEventInCalendars` で全カレンダーを検索
3. 旧カレンダーのイベントを削除 → 新カレンダーに再作成

### 4. CalendarApp由来のEvent IDをAdvanced Serviceにそのまま渡すと失敗する

**症状**: Advanced Serviceでの色取得関数（現 `_getEventColorAndLabelViaAdvancedService()`）を実装したのに、Calendar側で手動変更した色が一向にSheetsへ反映されない。

**原因**: `CalendarApp`（`event.getId()`）が返すEvent IDは `"xxxxx@google.com"` という末尾付きの形式。一方、Calendar Advanced Service（REST API）の `Calendar.Events.get()` はこの末尾を含まない素のIDを期待するため、そのまま渡すと404エラーになる。`catch` で例外を握りつぶしてログに出すだけの実装だったため、常に失敗していることに気づきにくかった。

**解決策**: `eventId.split("@")[0]` で末尾を取り除いてからAdvanced Serviceに渡す。

→ **CalendarAppのEvent IDをAdvanced Service（またはCalendar API）に渡す処理を新しく書くときは、必ずこの変換を思い出すこと。**

### 5. 終日タスクの行がCalendar側でイベント削除してもSheetsから消えない（Issue #31で対応済み）

**症状**: 終日タスク（B列＝開始日時が空欄）のCalendarイベントを削除しても、対応するSheets行が`syncCalendarToSheets`のポーリングで永久に削除されない。

**原因**: `_removeDeletedEvents()` は、行が同期取得範囲（`SYNC_DAYS_BACK`〜`SYNC_DAYS_AHEAD`）内かどうかをB列の値から`_sheetDateToCalendarDate()`で判定していたが、終日タスクはB列が仕様上つねに空欄のため、この関数は常に`null`を返す。`if (!startDate || ...) return;` の分岐によって、日付が読めない行は「範囲外かどうか判断できないので保守的にスキップ」という扱いになり、結果として終日タスクの行は削除判定の対象に一切ならなかった。

単純に「終日タスクは常に削除対象にする」と直すのは危険：終日タスクのCalendar上の日付は`_syncRowToCalendar`の`isAllDay`分岐で編集の都度`new Date()`（今日）に上書きされるため、長期間編集されていない終日タスクほど古い日付のまま止まる。その状態で`SYNC_DAYS_BACK`（7日）を過ぎると、Calendar上にまだ存在していてもポーリングの取得範囲から外れて「消えた」と誤判定し、Sheets行を誤って削除してしまう（実データ消失の回帰）。

**解決策**: 日付で範囲内外を判断できない行（終日タスク、およびフォーマット不正行）に限り、既存の`_findEventInCalendars()`（`getEventById`ベースで日付に縛られない直接検索。担当者変更時の旧イベント検索や`resetAndResyncToCalendar`で既に使われているのと同じ仕組み）で実在確認するようフォールバックさせた。見つかれば「範囲外なだけでまだ存在する」として削除しない、見つからなければ本当に削除されたと判断して行を削除する。

→ **同期範囲の内外判定にB列の日付だけを当てにできないケース（終日タスク・フォーマット不正）が出てきたら、まずこのフォールバックとその周辺を疑うこと。**

### 6. Calendar新UIの「ラベル」機能により`colorId`だけでの手動変更検知が不完全になっていた（Issue #32で対応済み）

**症状**: Calendar UIで手動色変更したイベントの一部が、Advanced Service経由でも`colorId`が正しく取得できず、Sheets側にステータスが反映されないことがあった。

**原因**: Googleカレンダーには`colorId`（固定11色の列挙値）を supersede する「ラベル」機能があり、`eventLabelId`（カレンダー単位で`labelProperties.eventLabels`に定義するカスタムラベルのUUID）で色を管理するようになっている。実機検証では、Calendar UIで4回手動色変更したところ`eventLabelId`は毎回確実に新しい値が書き込まれた一方、`colorId`は4回中2回しか返らず、値も`CONFIG.STATUS_COLORS`と一致しなかった。

**解決策**: このカレンダーの`labelProperties.eventLabels`には、`CONFIG.STATUS_COLORS`のキーと文字列完全一致する名前（"未着手"/"進行中"/"完了"/"キャンセル"）のラベルが担当者マスタの全カレンダーに既に用意されていたため、`eventLabelId` → ラベル名 → ステータス、という解決をラベル名の完全一致でそのまま行える（hexマッチング等は不要）。`_getEventColorAndLabelViaAdvancedService()`で`colorId`と`eventLabelId`をまとめて取得し、`eventLabelId`が解決できればそちらを優先、できなければ`colorId`にフォールバックする2段構えにした。書き込み側（`_syncRowToCalendar`の`setColor()`）は変更していない。

→ **新しく担当者を追加する場合、そのカレンダーにも同じ4つのラベル（未着手/進行中/完了/キャンセル）をCalendar UIで作成しておくこと。ラベルが無いカレンダーは自動的に`colorId`判定にフォールバックする（動作はするがCalendar手動変更の検知精度が落ちる）。**

### 7. Calendar新UIの「ラベル」機能により、書き込み側の`CalendarApp.setColor()`が無効化されていた（Issue #35で対応済み）

**症状**: Sheetsからタスクを新規作成/更新しても、Calendar側でイベントに色が一切つかない（デフォルト色のまま）。`_syncRowToCalendar`の`setColor()`自体は例外を投げず、Event IDも正常に払い出される。

**原因**: 項目6（読み取り側）の裏返し。`labelProperties`（ラベル機能）が設定されているカレンダーでは、`CalendarApp.setColor()`で書き込んだ色がCalendar側に一切反映されない。実機検証では、`setColor()`実行直後にAdvanced Service経由でイベントを取得しても`colorId`・`eventLabelId`とも`null`で、Calendar UI上も無色のままだった。同じイベントに対してAdvanced Serviceで`Calendar.Events.patch({eventLabelId: <ラベルのUUID>}, ..., {eventLabelVersion: 1})`を直接呼ぶと、API応答に`eventLabelId`が反映され、Calendar UI上でも実際に色が表示されることを目視確認済み。

**解決策**: `_syncRowToCalendar`で`setColor()`を呼んだ後、対象カレンダーの`labelProperties.eventLabels`（`_getCalendarLabelMaps()`で取得、`name→id`方向）からステータス名と完全一致するラベルを探し、見つかれば`_setEventLabel()`で`eventLabelId`を直接書き込む。ラベルが無い/名前が一致しないカレンダーは`setColor()`の結果のみに依存する（従来通り、退行なし）。`setColor()`自体は削除していない（ラベル無しカレンダーでは引き続き唯一の色付け手段のため）。

→ **新しく担当者を追加する場合、そのカレンダーに4つのラベル（未着手/進行中/完了/キャンセル）を作成しておくことが、読み取り側（項目6）だけでなく書き込み側の色表示にも必要。**

### 8. `events.forEach` 内の未捕捉例外、および `CalendarApp.getEventById()` の削除済みイベント誤検知により、Calendar削除がSheetsに反映されないことがあった（Issue #31再発）

**症状**: Issue #31（終日タスク限定の別原因、155166aで対応済み）の修正後も、Calendar側でイベントを削除してもSheets側の行が消えない事象が再発した。原因は2つあり、実機検証で切り分けた。

**原因1（副次的に発見・修正）**: `_syncCalendarToSheetsBody()`内の`events.forEach(...)`（イベントごとの差分検出・更新/新規行の組み立て）にはtry/catchが無く、ループ内で1件でも例外が発生すると forEach 全体、ひいては `_syncCalendarToSheetsBody` 全体がそこで中断される構造になっていた。削除判定を行う `_removeDeletedEvents()` はこのforEachの**後**で呼ばれる設計のため、ループが中断すると一度も実行されない。ステータス判定に使う `_getEventColorAndLabelViaAdvancedService()` / `_getCalendarLabelMaps()` 自体は内部でtry/catch済みで例外を握りつぶす設計だったため直接の原因ではなかったが、その前後で `CalendarApp` 経由でイベントオブジェクトのプロパティに直接アクセスする呼び出し（`event.getTitle()` / `isAllDayEvent()` / `getStartTime()` / `getEndTime()` / `getDescription()` / `getCreators()` / `getColor()` / `getId()`）は無防備で、対象イベントが同一実行中にCalendar側で削除・変更されると失敗しうる。これを個別のtry/catchで保護し、失敗したイベントだけスキップするよう修正（エラーはループ後にまとめて1回だけ通知）。

**原因2（実際の再発原因）**: 原因1の修正をpushしてもなお、終日タスクを削除したケースで再現した。`_removeDeletedEvents`は日付で範囲内外を判断できない行（終日タスク／フォーマット不正行）を`_findEventInCalendars()`で直接実在確認するが、この関数は`cal.getEventById(eventId)`の**戻り値が非nullかどうかだけ**を実在の判定基準にしていた。Google Calendarは削除されたイベントを即座に完全削除するのではなく、`status: "cancelled"`のトゥームストーンとしてしばらく保持する仕様のため、`CalendarApp.getEventById()`は**削除済みのイベントでも非nullのオブジェクトを返すことがある**。結果、終日タスクを削除しても`_findEventInCalendars`が誤って「まだ存在する」と判定し、`_removeDeletedEvents`が該当行を削除対象に含めていなかった。

**解決策**: 新設した`_isEventActiveViaAdvancedService(calendarId, eventId)`で、Calendar Advanced Serviceの`Calendar.Events.get(..., {fields: "status"})`を呼び、`status !== "cancelled"`かどうかで真の有効性を確認する。`_removeDeletedEvents`の終日タスク用フォールバック分岐で、`_findEventInCalendars`が見つけた場合でもこのstatus確認を併せて行い、`status === "cancelled"`なら削除対象とみなすようにした。API通信エラー等でstatusを確定できない場合は、誤削除（データ消失）を避けるため「有効（削除しない）」を返す安全側設計にしている。`_findEventInCalendars`自体（他の4箇所の呼び出し元）は変更していない。

→ **`CalendarApp.getEventById()`の戻り値の非null判定だけを「イベントが存在する」根拠にしないこと。削除済みイベントの検知（=本当に消えたかどうかの確認）が必要な場面では、Advanced Serviceで`status`フィールドを確認すること。CalendarAppの型はこのstatus情報を公開していない。**
→ **`events.forEach`のようなイベント単位ループを新しく書く/触るときは、CalendarAppのイベントオブジェクトへの直接アクセスも「同一実行中に対象がCalendar側で削除/変更されて失敗しうる」ことを前提に、必ずイベント単位のtry/catchで保護すること。**

## 既知の未対応課題

1. **Sheets側で行ごと削除した場合、Calendar側に反映されない**（重要度：中）— タスク名セルを空にする「中身だけ消す削除」は対応済み（担当者の状態に関わらず動作する）。行削除自体は `onChange` トリガーが必要だが実装コストが高いため**運用ルールで対応**（行削除禁止・中身消しで代替）

過去にあった以下の課題は対応済み: Calendar起点のステータス変更（色変更）のSheets反映（Calendar Advanced Serviceで対応）、期日の通知機能（Google Chat連携で対応）、新規行の初期値補完（ステータス未入力→未着手、開始日時未入力→終日タスク）。

## 運用上の合意事項

- カレンダーイベントのタイトルは**タスク名のみ**（`[ステータス] @担当者` は含めない）
- 担当者はカレンダー（どのカレンダーに入っているか）で判別
- ステータスはイベント色で視覚表示。**変更はSheets側から行うことを推奨**（Calendar UIでの手動色/ラベル変更もAdvanced Service経由でSheetsに反映されるが、Sheets側が正のデータソースという運用方針は変えない）
- タスクを削除したい場合は**行削除ではなく、タスク名を空にする**（開始日時を空にするのは「終日タスクにする」という意味になり、削除トリガーではなくなった）
- 色付け（ヘッダー背景色、ゼブラストライプ）は不要（削除済み）

## GASプロジェクトの種別（clasp連携の前提）

- **Bound script**（スプレッドシート紐付け型）。`SpreadsheetApp.getActive()` がそのまま対象シートを指す。
- clasp で管理。`clasp push` でGASにデプロイ。認証切れ時は `clasp login` で再認証。
- 対象スプレッドシートのID: `1IJm9RBFAJ_bxpLsxiWEL1yiOHMPH20dwDLdrp_Kz0aPUhnNHMzObMon9`
- `.clasp.json` の `scriptId` はスプレッドシートの「拡張機能 > Apps Script」から開いたエディタのURLから取得。
- `appsscript.json` で Calendar Advanced Service (v3) を有効化済み。`clasp push` 後の初回実行時に、通常より広い権限（Calendar API）の再承認ダイアログが出ることがある。エラーが出る場合はAppsScriptエディタの「サービス」一覧にCalendar APIが表示されているか確認すること。

## Session.getActiveUser() の既知の制約

`Session.getActiveUser().getEmail()` は、Google Workspaceの組織設定によっては空文字を返すことがある。その場合、作成者・最終更新者列（I・J）が空欄になる。コードのバグではない。

## setup() の再実行が必要なタイミング

- トリガーが消えた/動かなくなったとき（`_registerTriggers()` が重複削除→再登録するため安全）
- 初回セットアップ時

列構成変更時のヘッダー行は自動生成の仕組みがないため、手動で更新すること。

## テスト環境について

未整備。本番のスプレッドシート・カレンダーに対して直接動作確認しながら開発している。
