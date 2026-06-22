# GMS Calendar Sync 仕様書

> 最終更新: 2026-06-22（Microsoft認証エンドポイントをテナント固定URLに変更）

---

## 1. プロジェクト概要

| 項目 | 内容 |
|------|------|
| 名称 | GMS Calendar Sync |
| バージョン | 1.0.0 |
| 種別 | Chrome拡張機能（Manifest V3） |
| 拡張機能ID | `kfeidfgokolbgomcmokifecaajkcobhl` |
| 目的 | GoogleカレンダーとMicrosoft 365（Outlook）カレンダーの双方向同期 |

---

## 2. ファイル構成

```
chrome-extention-calendersync-gms/
├── manifest.json              # 拡張機能設定（権限・OAuth2・エントリポイント）
├── popup.html                 # メインポップアップUI
├── popup.css                  # ポップアップスタイル
├── popup.js                   # ポップアップロジック
├── background.js              # バックグラウンドサービスワーカー（同期・認証処理）
├── test-no-duplicates.mjs     # Node.js製E2Eテストハーネス（重複防止・ログ記録検証）
└── spec.md                    # 本仕様書
```

---

## 3. 認証

| サービス | 方式 | Client ID の管理 |
|---------|------|----------------|
| Google | Chrome Identity API（`getAuthToken`） | `manifest.json` の `oauth2.client_id` に直接記載 |
| Microsoft | `chrome.identity.launchWebAuthFlow`（OAuth2 暗黙的フロー） | `background.js` の `MICROSOFT_CLIENT_ID` 定数にハードコード |

### Microsoft OAuth 設定

| 項目 | 値 |
|------|---|
| テナント | `0e75c023-b71b-4f6a-8550-a6970bee3358`（固定テナント） |
| 認証エンドポイント | `https://login.microsoftonline.com/0e75c023-b71b-4f6a-8550-a6970bee3358/oauth2/v2.0/authorize`（テナント固定） |
| スコープ | `Calendars.ReadWrite offline_access` |
| リダイレクトURI | `https://kfeidfgokolbgomcmokifecaajkcobhl.chromiumapp.org/` |
| フロー | 暗黙的フロー（Azure Portal の「認証」→「アクセストークン」を有効化済みが必要） |
| Azure アプリ設定 | サポートされているアカウントの種類：「この組織のディレクトリ内のアカウントのみ（単一テナント）」 |

### MSトークン期限切れ検知

- `fetchMicrosoftEvents` が 401 を受信すると、`msTokenExpired: true` をストレージに保存
- 拡張機能アイコンに赤バッジ（`!`）を表示
- ポップアップに「⚠️ トークン期限切れ — 再認証してください」と表示し、「今すぐ同期」を無効化
- 再認証成功または同期成功時に `msTokenExpired: false` へリセット

---

## 4. カレンダー同期

### 4.1 同期仕様

| 項目 | 内容 |
|------|------|
| 同期方向 | 双方向（Google ↔ Microsoft） |
| 同期対象期間 | 過去7日〜将来180日（約6ヶ月） |
| 自動同期間隔 | 15分ごと（`chrome.alarms`） |
| 手動同期 | ポップアップの「今すぐ同期」ボタン |
| Google 取得上限 | 2500件/ページ（`nextPageToken` でページネーション対応、全件取得） |
| Microsoft 取得上限 | 500件/ページ（`@odata.nextLink` でページネーション対応、全件取得） |
| 同時実行防止 | `isSyncing` フラグによるガード（重複実行をスキップ） |

### 4.2 同期フロー

```
1. トークン取得（非インタラクティブ）
2. Google・Microsoft 両カレンダーからイベント全件取得
   （重複チェックは全件参照が必要なため updatedMin/startDateTime フィルタなし）
3. API拡張プロパティから同期済みキャッシュを再構築
   （ストレージがクリアされた場合でも判定1が確実に機能するよう補完）
4. lastSyncTime 以降に更新されたイベントのみを処理対象としてフィルタ
5. Google → Microsoft 同期（3段階の重複チェックを通過したイベントのみ作成）
   - 作成成功時のみ syncedToMs に記録し googleToMs をカウント
   - 作成失敗時は googleToMsFailed をカウント
6. Microsoft → Google 同期（3段階の重複チェックを通過したイベントのみ作成）
   - 作成成功時のみ syncedToGoogle に記録し msToGoogle をカウント
   - 作成失敗時は msToGoogleFailed をカウント
7. ローカルキャッシュ・lastSyncTime を保存、MSトークン期限切れフラグをクリア
8. 同期ログを保存（成功/失敗カウント含む）
```

### 4.3 重複防止の仕組み（三段階チェック）

重複チェックは **全取得イベント**（処理対象フィルタ前の全件）を参照する。
これにより「前回同期済みだが lastSyncTime 以前に作成されたイベント」の重複を確実に防ぐ。

#### Google → Microsoft

| 判定 | 手段 | 説明 |
|------|------|------|
| 判定1 | `syncedToMs`（ローカルキャッシュ） | 同期済み Google イベント ID を記録。最速・最優先 |
| 判定2 | MS 拡張プロパティ `original_google_id` | `$expand` で取得。一致すればキャッシュも補完してスキップ |
| 判定3 | タイトル＋日時フォールバック | `$expand` 不安定時の最終手段。終日は日付一致、時刻ありは `msDatetimeToUTC` でUTC正規化後1分以内を同一とみなす |

#### Microsoft → Google

| 判定 | 手段 | 説明 |
|------|------|------|
| 判定1 | `syncedToGoogle`（ローカルキャッシュ） | 同期済み MS イベント ID を記録。最速・最優先 |
| 判定2 | Google 拡張プロパティ `original_ms_id` | Google API は常に返却するため信頼性が高い。全件参照 |
| 判定3 | タイトル＋日時フォールバック | キャッシュなし・`$expand` 失敗時の補完。終日は日付一致、時刻ありは先頭19文字（`YYYY-MM-DDTHH:mm:ss`）で1分以内を確認 |

> **注意：** MS Graph API の `calendarView` における `$expand`（`singleValueExtendedProperties`）は不安定なことがあり、判定2が機能しないケースがある。その場合、判定3が最終的な重複防止として機能する。

#### キャッシュ再構築（同期開始時に毎回実行）

```js
// MS拡張プロパティから syncedToMs を補完
for (const mEvent of allMsEvents) {
  const prop = (mEvent.singleValueExtendedProperties || []).find(
    p => p.id.includes("Name original_google_id")
  );
  if (prop?.value) syncedToMs[prop.value] = true;
}

// Google拡張プロパティから syncedToGoogle を補完
for (const gEvent of allGoogleEvents) {
  const msId = gEvent.extendedProperties?.private?.original_ms_id;
  if (msId) syncedToGoogle[msId] = true;
}
```

`chrome.storage.local` がクリアされても、APIが拡張プロパティを返す限り判定1が有効になる。

### 4.4 API 拡張プロパティ

| プロパティ名 | 設定先 | 値 |
|------------|-------|---|
| `original_google_id` | Microsoft イベント (`singleValueExtendedProperties`) | 同期元の Google イベント ID |
| `synced_from` | Microsoft イベント (`singleValueExtendedProperties`) | `"google"` 固定 |
| `original_ms_id` | Google イベント (`extendedProperties.private`) | 同期元の Microsoft イベント ID |
| `synced_from` | Google イベント (`extendedProperties.private`) | `"microsoft"` 固定 |

MS Graph の GUID: `{66f5a359-4659-4830-9070-00040ec6ac6e}`

### 4.5 タイムゾーン正規化ヘルパー

```js
function msDatetimeToUTC(datetimeStr, timezoneName) → number (UTC epoch ms)
```

MS Graph は `dateTime` をタイムゾーンなしのローカル時刻文字列で返す（例: `"2026-06-22T10:00:00"`）。
この関数は `Intl.DateTimeFormat.formatToParts` を使い、
`timezoneName`（例: `"Asia/Tokyo"`）のオフセットを計算してUTCエポックmsに変換する。
タイムゾーン名が不明な場合はシステムローカル時刻としてフォールバック解析する。

### 4.6 イベント作成処理

#### `createEventInMicrosoft(token, gEvent)` → `boolean`

- `gEvent.start.date` がある場合は終日イベント（`isAllDay: true`）として作成
- 日時が取得できない予定は `return false` でスキップ
- API エラー時は `console.error` を出力して `return false`（呼び出し元が失敗カウンタを増加）
- 成功時 `return true`

#### `createEventInGoogle(token, mEvent)` → `boolean`

- `mEvent.isAllDay` が不安定なため補完ロジックあり:
  開始・終了がともに `T00:00:00` の場合も終日とみなす
- MSのローカル時刻文字列から先頭19文字のみ使用（`Z` を付けない）
- API エラー時は `console.error` を出力して `return false`（呼び出し元が失敗カウンタを増加）
- 成功時 `return true`

---

## 5. 同期ログ

### 5.1 仕様

| 項目 | 内容 |
|------|------|
| 保存先 | `chrome.storage.local`（キー：`syncLogs`） |
| 保持件数 | 最新10件 |
| 表示場所 | ポップアップ下部 |
| エラー耐性 | `saveSyncLog` はストレージ障害時に例外を投げない（`try/catch` で保護し `console.error` のみ出力） |

### 5.2 ログエントリ形式

成功時：
```json
{
  "timestamp": "2026-06-22T10:30:00.000Z",
  "success": true,
  "timeMin": "2026-06-15T00:00:00.000Z",
  "timeMax": "2026-12-19T00:00:00.000Z",
  "googleToMs": 3,
  "msToGoogle": 1,
  "googleToMsFailed": 0,
  "msToGoogleFailed": 0
}
```

エラー時：
```json
{
  "timestamp": "2026-06-22T10:30:00.000Z",
  "success": false,
  "error": "認証が必要です: ..."
}
```

### 5.3 ポップアップ表示

| 状態 | 表示 |
|------|------|
| 成功（失敗なし） | `✅ 6/15〜12/19  G→M: 3件 / M→G: 1件` |
| 成功（一部失敗） | `✅ 6/15〜12/19  G→M: 2件 / M→G: 1件  ⚠️ 失敗 G→M:1 M→G:0`（オレンジ色） |
| エラー | `❌ 認証が必要です: ...` |
| ログなし | 「ログなし」と表示 |

> **注意:** ポップアップは開いた時点のストレージ内容を表示する。バックグラウンド自動同期のログは次回ポップアップを開いた時に反映される。

---

## 6. ストレージキー一覧

| キー | 型 | 内容 | 更新タイミング |
|------|----|------|----------------|
| `googleToken` | string | Google アクセストークン | 認証成功時 |
| `msToken` | string | Microsoft アクセストークン | 認証成功時 |
| `msTokenExpired` | boolean | MSトークン期限切れフラグ | 401受信時に `true`、同期成功・再認証時に `false` |
| `syncedToMs` | object | `{ [googleEventId]: true }` — Google→MS 同期済み ID マップ | 同期成功時 |
| `syncedToGoogle` | object | `{ [msEventId]: true }` — MS→Google 同期済み ID マップ | 同期成功時 |
| `lastSyncTime` | string | 前回同期完了時刻（ISO8601） | 同期成功時 |
| `syncLogs` | array | 同期ログ（最新10件） | 同期完了・エラー時 |

---

## 7. 権限

| 権限 | 用途 |
|------|------|
| `identity` | OAuth2 認証フロー |
| `storage` | トークン・キャッシュ・ログの保存 |
| `alarms` | 15分ごとの定期同期 |

---

## 8. ポップアップ UI

```
┌─────────────────────────────────┐
│ カレンダー同期                   │
│ GoogleとMicrosoft 365のカレン…  │
├─────────────────────────────────┤
│ [G Googleでサインイン          ] │
│                         未ログイン│
│ [M Microsoftでサインイン       ] │
│                         未ログイン│
├─────────────────────────────────┤
│ [        今すぐ同期            ] │
│                          準備完了│
├─────────────────────────────────┤
│ 同期ログ                   クリア│
│ 6/22 10:30 ✅ 6/15〜12/19      │
│ G→M: 3件 / M→G: 1件           │
│ 6/22 09:15 ✅ 6/15〜12/19      │
│ G→M: 0件 / M→G: 0件           │
└─────────────────────────────────┘
```

| 要素 | 動作 |
|------|------|
| 「今すぐ同期」ボタン | 両サービスにログイン済みかつトークン未期限切れの場合のみ有効 |
| 「クリア」リンク | `syncLogs` を削除してリストをリセット |
| Google / MS ステータス | ストレージのトークン有無に応じてテキスト更新 |
| 同期ステータス | 「同期中...」→「同期完了 (HH:MM:SS)」または「同期失敗: ...」 |

---

## 9. テスト

`test-no-duplicates.mjs`（Node.js VM モジュールで `background.js` 実コードを実行）

| # | テスト内容 |
|---|-----------|
| 1 | Google→MS 初回同期: カウントが正しく記録される |
| 2 | Google→MS 2回目同期: 重複ゼロ（キャッシュ判定1） |
| 3 | MS→Google 初回同期: カウントが正しく記録される |
| 4 | MS→Google 2回目同期: 重複ゼロ（キャッシュ判定1） |
| 5 | ストレージクリア後: 拡張プロパティからキャッシュ再構築し重複ゼロ |
| 6 | `createEventInMicrosoft` 失敗時: `googleToMsFailed` がカウントされる |
| 7 | `createEventInGoogle` 失敗時: `msToGoogleFailed` がカウントされる |
| 8 | ストレージエラー時: `saveSyncLog` は例外を投げず同期フローを継続する |

実行方法:
```sh
node test-no-duplicates.mjs
```

---

## 10. 既知の注意事項

| 項目 | 内容 |
|------|------|
| MS `$expand` の不安定性 | MS Graph `calendarView` の `singleValueExtendedProperties` `$expand` は返却が不安定。その場合は判定3（タイトル＋日時）が最終防衛線として機能する |
| MS `isAllDay` の不安定性 | `isAllDay: false` でも開始・終了が `T00:00:00` の場合があるため `createEventInGoogle` 内で補完判定を行っている |
| MSトークンの有効期限 | 暗黙的フローで取得したアクセストークンは通常1時間で失効。ポップアップからの再認証が必要 |
| ログの初期状態 | インストール直後または再読み込み後、自動同期（15分）が走るまで「ログなし」と表示される。「今すぐ同期」ボタンで即時確認可能 |
