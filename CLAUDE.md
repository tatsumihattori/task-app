# TaskSync — Sheets ↔ Calendar 双方向同期 引き継ぎメモ

このファイルは開発の引き継ぎ・コンテキスト復元用メモです。
`TaskSync.js` と一緒に Apps Script プロジェクトのディレクトリに置いてください。

## 概要

Google Sheets と Google Calendar の双方向タスク同期システム。Google Apps Script (GAS) で実装。

- **Sheets → Calendar**: セル編集時に即座反映（`onEdit` インストーラブルトリガー）
- **Calendar → Sheets**: 10分ごとのポーリングで反映（時間ベーストリガー）

## シート構成

| 列 | 内容 | 備考 |
|---|---|---|
| A | タスク名 | |
| B | 開始日時 | `YYYY/MM/DD HH:mm` 形式（1桁の月・日・時・分も許容） |
| C | 終了日時 | 同上。空欄時は開始+1時間 |
| D | ステータス | 未着手 / 進行中 / 完了 / キャンセル |
| E | 担当者 | 担当者マスタに登録された名前のみ有効 |
| F | メモ/説明 | カレンダーイベントの説明欄に同期 |
| G | Event ID | 自動管理・編集不可 |
| H | 最終同期日時 | 自動管理。バリデーションエラー時は警告文字列を書き込む |
| I | 作成者 | 自動管理 |
| J | 最終更新者 | 自動管理 |
| K | プロジェクト名 | 表示のみ・GAS未使用 |
| L | 期日 | 表示のみ・GAS未使用（将来の通知機能用に予約） |

カレンダー側のイベントタイトル形式: **タスク名のみ**（ステータス・担当者はタイトルに含めない）
説明欄はメモのみ。

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
```

COL は A=1 〜 L=12 の12列構成。`Object.keys(CONFIG.COL).length` で `NUM_COLS` を動的計算しているため、列追加時は COL にキーを追加するだけでよい。

## 担当者マスタ（GID: 881583119）

A列=担当者名、B列=カレンダーID。
`_getAssigneeCalendarMap()` でMapに読み込む。マスタ未登録の担当者はSheets→Calendar同期をスキップ。
カレンダーID→担当者名の逆引きマップも同関数内で生成（`calIdToAssignee`）し、Calendar→Sheets同期での担当者列埋めに使用。

## ステータスとカレンダーイベント色の対応

**Sheets → Calendar**（`_syncRowToCalendar`）: `CONFIG.STATUS_COLORS[status]` で色を決定し `event.setColor()` で設定。正常動作。

**Calendar → Sheets**（`_syncCalendarToSheetsBody`）: `colorToStatus` 逆引きマップを使い `event.getColor()` で色を読む。ただし **GASの `CalendarApp.getColor()` はUIで手動変更した色を返さない**（空文字を返す）ため、カレンダー側での手動色変更はステータスに反映されない。
→ 回避策として Calendar Advanced Service (REST API) の `Calendar.Events.list()` を使う方法が検討中（未実装）。

## 解決済みの重要な技術的ハマりどころ

### 1. 日時のタイムゾーンズレ（最重要）

**症状**: Sheetsで10:00と入力したのに、Calendarには19:00で登録される（9時間ズレ＝JSTのオフセット分）。

**原因**: `Range.getValue()` で取得したDateオブジェクトはすでに内部的にUTCシフトしており、そこにさらにタイムゾーン補正を加えると二重にズレる。

**解決策**: `getDisplayValue()` でセルの表示文字列を取得し、正規表現でパースして `new Date(年, 月, 日, 時, 分)` で直接組み立てる（TZ変換を一切挟まない）。実装は `_sheetDateToCalendarDate()` を参照。

→ **今後似た日時バグが出たら、まずこの関数とその呼び出し元を疑うこと。**

正規表現は `\d{1,2}` で1・2桁両方に対応済み（Sheetsのセル書式次第で時刻が `9:30` のように1桁になるため）。

### 2. `CalendarEvent.getColor()` の制約

GASの CalendarApp 経由では、UIで手動変更したイベント色を `getColor()` で取得できない（空文字を返す）。GASがプログラムで `setColor()` した色のみ取得可能。詳細は「ステータスとカレンダーイベント色の対応」セクション参照。

### 3. 担当者変更時のカレンダーイベント移動

担当者が変わるとターゲットカレンダーが変わる。`_syncRowToCalendar` では：

1. 新担当者のカレンダーでEventIDを検索
2. 見つからなければ `_findEventInCalendars` で全カレンダーを検索
3. 旧カレンダーのイベントを削除 → 新カレンダーに再作成

## 既知の未対応課題

1. **Sheets側で行ごと削除した場合、Calendar側に反映されない**（重要度：中）— タスク名・日時セルを空にする「中身だけ消す削除」は対応済み。行削除自体は `onChange` トリガーが必要だが実装コストが高いため**運用ルールで対応**（行削除禁止・中身消しで代替）
2. **Calendar起点のステータス変更（色変更）がSheetsに反映されない**（重要度：中）— Calendar Advanced Service を使えば解決可能、未実装
3. **期日の通知機能がない**（重要度：中）— L列の期日をもとにチャット/メール通知する機能を将来実装予定（現在はGAS未使用）
4. **新規行の初期値補完がない**（重要度：低）— ステータス等のデフォルト値自動入力

## 運用上の合意事項

- カレンダーイベントのタイトルは**タスク名のみ**（`[ステータス] @担当者` は含めない）
- 担当者はカレンダー（どのカレンダーに入っているか）で判別
- ステータスはイベント色で視覚表示。**変更はSheets側から行う**（Calendat UIでの色変更はSheetsに反映されない）
- タスクを削除したい場合は**行削除ではなく、タスク名または開始日時を空にする**
- 色付け（ヘッダー背景色、ゼブラストライプ）は不要（削除済み）

## GASプロジェクトの種別（clasp連携の前提）

- **Bound script**（スプレッドシート紐付け型）。`SpreadsheetApp.getActive()` がそのまま対象シートを指す。
- clasp で管理。`clasp push` でGASにデプロイ。認証切れ時は `clasp login` で再認証。
- 対象スプレッドシートのID: `1IJm9RBFAJ_bxpLsxiWEL1yiOHMPH20dwDLdrp_Kz0aPUhnNHMzObMon9`
- `.clasp.json` の `scriptId` はスプレッドシートの「拡張機能 > Apps Script」から開いたエディタのURLから取得。

## Session.getActiveUser() の既知の制約

`Session.getActiveUser().getEmail()` は、Google Workspaceの組織設定によっては空文字を返すことがある。その場合、作成者・最終更新者列（I・J）が空欄になる。コードのバグではない。

## setup() の再実行が必要なタイミング

- トリガーが消えた/動かなくなったとき（`_registerTriggers()` が重複削除→再登録するため安全）
- 初回セットアップ時

列構成変更時は `_createHeader()` が `setup()` 内でコメントアウトされているため、ヘッダーは手動で更新すること。

## テスト環境について

未整備。本番のスプレッドシート・カレンダーに対して直接動作確認しながら開発している。
