# TaskSync — Sheets ↔ Calendar 双方向同期 引き継ぎメモ

このファイルは開発の引き継ぎ・コンテキスト復元用メモです。
`TaskSync.js` と一緒に Apps Script プロジェクトのディレクトリに置いてください。

## 概要

Google Sheets と Google Calendar の双方向タスク同期システム。Google Apps Script (GAS) で実装。

- **Sheets → Calendar**: セル編集時に即座反映（`onEdit` インストーラブルトリガー）
- **Calendar → Sheets**: 10分ごとのポーリングで反映（時間ベーストリガー）
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
※ `_createHeader()` のヘッダー配列はA〜J列（10列）までしか定義されておらず、既にこの13列構成と一致していない（Issue #11、未対応）。`setup()`内でコメントアウトされているため実害はないが、将来コメントアウトを解除する際は先に更新すること。

## 担当者マスタ（GID: 881583119）

A列=担当者名、B列=カレンダーID、C列=ChatユーザーID（`users/xxxxxxxxxxxxxxxxxxx`形式。任意項目）。
`_getAssigneeCalendarMap()` でA・B列をMapに読み込む。マスタ未登録の担当者はSheets→Calendar同期をスキップ（削除は担当者の状態に関わらず行われる。詳細は「運用上の合意事項」参照）。
カレンダーID→担当者名の逆引きマップも同関数内で生成（`calIdToAssignee`）し、Calendar→Sheets同期での担当者列埋めに使用。
C列は `_getAssigneeChatUserMap()` で読み込み、Google Chat期日通知でのメンションに使用（未入力の担当者は名前をテキストで含めるだけでメンションは飛ばない）。

## Google Chat連携（期日通知）

- `notifyUpcomingDeadlines()`（毎日9時、時間ベーストリガー）→ `_notifyUpcomingDeadlinesBody()` が本体
- Webhook URLは**ソースコードには書かない**。Apps Scriptエディタの「プロジェクトの設定 > スクリプト プロパティ」で `CHAT_WEBHOOK_URL` を設定する（`key`/`token` を含む実質的な認証情報のため）。未設定の場合は通知をスキップするだけでエラーにはならない
- 担当者マスタC列（ChatユーザーID）が設定されていれば `<users/xxxx>` 形式で本文に埋め込み、Chatが実際のメンション（通知）として扱う。未設定なら担当者名をテキストとして含めるだけ
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

**Sheets → Calendar**（`_syncRowToCalendar`）: `CONFIG.STATUS_COLORS[status]` で色を決定し `event.setColor()` で設定。正常動作。

**Calendar → Sheets**（`_syncCalendarToSheetsBody`）: `colorToStatus` 逆引きマップを使い、まず `event.getColor()`（CalendarApp）で色を読む。**GASの `CalendarApp.getColor()` はUIで手動変更した色を返さない**（空文字を返す）ため、その値が `STATUS_COLORS` のいずれにも一致しない場合のみ、`_getEventColorIdViaAdvancedService()` で Calendar Advanced Service (REST API, `Calendar.Events.get()` の `colorId`) を呼び、実際の色を取得する。取得できた `colorId` は `CONFIG.STATUS_COLORS` の値（"2"/"9"/"8"/"7"）と同じ数値文字列表現のため変換不要。
Advanced Serviceの呼び出しはイベント単位のAPIコールになるため、CalendarApp側で色が判定できる場合（プログラムでsetColorした直後など）はスキップして呼ばないようにしている。

## 解決済みの重要な技術的ハマりどころ

### 1. 日時のタイムゾーンズレ（最重要）

**症状**: Sheetsで10:00と入力したのに、Calendarには19:00で登録される（9時間ズレ＝JSTのオフセット分）。

**原因**: `Range.getValue()` で取得したDateオブジェクトはすでに内部的にUTCシフトしており、そこにさらにタイムゾーン補正を加えると二重にズレる。

**解決策**: `getDisplayValue()` でセルの表示文字列を取得し、正規表現でパースして `new Date(年, 月, 日, 時, 分)` で直接組み立てる（TZ変換を一切挟まない）。実装は `_sheetDateToCalendarDate()` を参照。

→ **今後似た日時バグが出たら、まずこの関数とその呼び出し元を疑うこと。**

正規表現は `\d{1,2}` で1・2桁両方に対応済み（Sheetsのセル書式次第で時刻が `9:30` のように1桁になるため）。

### 2. `CalendarEvent.getColor()` の制約（Advanced Serviceで対応済み）

GASの CalendarApp 経由では、UIで手動変更したイベント色を `getColor()` で取得できない（空文字を返す）。GASがプログラムで `setColor()` した色のみ取得可能。
→ `_getEventColorIdViaAdvancedService()` で Calendar Advanced Service の `Calendar.Events.get()` を呼び、実際の `colorId` を取得することで解決済み。`appsscript.json` の `dependencies.enabledAdvancedServices` にCalendar API (v3) を追加してある。詳細は「ステータスとカレンダーイベント色の対応」セクション参照。

### 3. 担当者変更時のカレンダーイベント移動

担当者が変わるとターゲットカレンダーが変わる。`_syncRowToCalendar` では：

1. 新担当者のカレンダーでEventIDを検索
2. 見つからなければ `_findEventInCalendars` で全カレンダーを検索
3. 旧カレンダーのイベントを削除 → 新カレンダーに再作成

## 既知の未対応課題

1. **Sheets側で行ごと削除した場合、Calendar側に反映されない**（重要度：中）— タスク名セルを空にする「中身だけ消す削除」は対応済み（担当者の状態に関わらず動作する）。行削除自体は `onChange` トリガーが必要だが実装コストが高いため**運用ルールで対応**（行削除禁止・中身消しで代替）

過去にあった以下の課題は対応済み: Calendar起点のステータス変更（色変更）のSheets反映（Calendar Advanced Serviceで対応）、期日の通知機能（Google Chat連携で対応）、新規行の初期値補完（ステータス未入力→未着手、開始日時未入力→終日タスク）。

## 運用上の合意事項

- カレンダーイベントのタイトルは**タスク名のみ**（`[ステータス] @担当者` は含めない）
- 担当者はカレンダー（どのカレンダーに入っているか）で判別
- ステータスはイベント色で視覚表示。**変更はSheets側から行う**（Calendat UIでの色変更はSheetsに反映されない）
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

列構成変更時は `_createHeader()` が `setup()` 内でコメントアウトされているため、ヘッダーは手動で更新すること。

## テスト環境について

未整備。本番のスプレッドシート・カレンダーに対して直接動作確認しながら開発している。
