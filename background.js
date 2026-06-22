const SYNC_INTERVAL_MINUTES = 15;
const MICROSOFT_CLIENT_ID = "33bb455e-161c-42be-a4fa-c059fdc2a0bc";

// 同時実行防止フラグ
let isSyncing = false;

// ==========================================
// 1. メッセージリスナー (Popupからの要求を処理)
// ==========================================
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'loginGoogle') {
    getGoogleToken(true)
      .then(token => sendResponse({ success: true, token }))
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true; // 非同期応答を示す
  }
  
  if (request.action === 'loginMicrosoft') {
    getMicrosoftToken(true)
      .then(token => sendResponse({ success: true, token }))
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true;
  }
  
  if (request.action === 'syncNow') {
    executeSync()
      .then(() => sendResponse({ success: true }))
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true;
  }
});

// ==========================================
// 2. 認証関連処理
// ==========================================

async function getGoogleToken(interactive = false) {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive }, async function(token) {
      if (chrome.runtime.lastError) {
        return reject(chrome.runtime.lastError);
      }
      await chrome.storage.local.set({ googleToken: token });
      resolve(token);
    });
  });
}

async function getMicrosoftToken(interactive = false) {
  const clientId = MICROSOFT_CLIENT_ID;
  const redirectUri = chrome.identity.getRedirectURL();
  const scope = "Calendars.ReadWrite offline_access";
  const authUrl = `https://login.microsoftonline.com/0e75c023-b71b-4f6a-8550-a6970bee3358/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${encodeURIComponent(redirectUri)}&scope=${encodeURIComponent(scope)}`;

  return new Promise((resolve, reject) => {
    chrome.identity.launchWebAuthFlow({ url: authUrl, interactive }, async function(redirectUrl) {
      if (chrome.runtime.lastError) {
        return reject(chrome.runtime.lastError);
      }
      if (!redirectUrl) {
        return reject(new Error("認証フローが中断されました"));
      }

      const hash = new URL(redirectUrl).hash.substring(1);
      const urlParams = new URLSearchParams(hash);
      const token = urlParams.get('access_token');
      
      if (token) {
        await chrome.storage.local.set({ msToken: token, msTokenExpired: false });
        await chrome.action.setBadgeText({ text: '' });
        resolve(token);
      } else {
        reject(new Error("アクセストークンを取得できませんでした"));
      }
    });
  });
}

// ==========================================
// 3. アラーム（定期同期）処理
// ==========================================

// インストール/起動時にアラームを設定
chrome.runtime.onInstalled.addListener(() => {
  chrome.alarms.create("sync-calendar", { periodInMinutes: SYNC_INTERVAL_MINUTES });
});

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === "sync-calendar") {
    // バックグラウンド同期（非インタラクティブ）
    executeSync().catch(err => console.error("定期同期エラー:", err));
  }
});

// ==========================================
// 4. 同期ロジック
// ==========================================

async function executeSync() {
  // ② 同時実行ガード
  if (isSyncing) {
    console.log("同期処理がすでに実行中のためスキップします");
    return;
  }
  isSyncing = true;
  console.log("同期処理を開始します...");

  try {
    await _doSync();
  } catch (err) {
    console.error("同期エラー:", err.message);
    await saveSyncLog({ success: false, error: err.message });
    throw err;
  } finally {
    isSyncing = false;
  }
}

async function _doSync() {
  let googleToken, msToken;
  try {
    googleToken = await getGoogleToken(false);
    const storageData = await chrome.storage.local.get(['msToken']);
    msToken = storageData.msToken;
    if (!msToken) throw new Error("Microsoftトークンがありません");
  } catch (err) {
    throw new Error("認証が必要です: " + err.message);
  }

  const now = new Date();
  const timeMin = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000).toISOString();
  const timeMax = new Date(now.getTime() + 180 * 24 * 60 * 60 * 1000).toISOString();

  const { lastSyncTime, syncedToMs = {}, syncedToGoogle = {} } =
    await chrome.storage.local.get(['lastSyncTime', 'syncedToMs', 'syncedToGoogle']);

  // 重複チェックは全件が必要なため updatedMin なしで全取得
  // 処理対象のみ lastModifiedDateTime / updated でクライアント側フィルタ
  const allGoogleEvents = await fetchGoogleEvents(googleToken, timeMin, timeMax);
  const allMsEvents    = await fetchMicrosoftEvents(msToken, timeMin, timeMax);

  // API拡張プロパティからキャッシュを再構築
  // ストレージがクリアされた後でも $expand が返す限り判定1で確実にスキップできる
  for (const mEvent of allMsEvents) {
    const prop = (mEvent.singleValueExtendedProperties || []).find(
      p => p.id.includes("Name original_google_id")
    );
    if (prop?.value) syncedToMs[prop.value] = true;
  }
  for (const gEvent of allGoogleEvents) {
    const msId = gEvent.extendedProperties?.private?.original_ms_id;
    if (msId) syncedToGoogle[msId] = true;
  }

  const googleEventsToProcess = lastSyncTime
    ? allGoogleEvents.filter(e => new Date(e.updated) > new Date(lastSyncTime))
    : allGoogleEvents;
  const msEventsToProcess = lastSyncTime
    ? allMsEvents.filter(e => new Date(e.lastModifiedDateTime) > new Date(lastSyncTime))
    : allMsEvents;

  let googleToMs = 0;
  let msToGoogle = 0;
  let googleToMsFailed = 0;
  let msToGoogleFailed = 0;

  // ── Google -> MS ──────────────────────────────────────────
  for (const gEvent of googleEventsToProcess) {
    // MS由来イベントはスキップ（無限ループ防止）
    if (gEvent.extendedProperties?.private?.synced_from === 'microsoft') continue;

    // 判定1: ローカルキャッシュ（最速・最優先）
    if (syncedToMs[gEvent.id]) continue;

    // 判定2: MS拡張プロパティ（$expand が返った場合のみ有効）
    const existsInMsByProp = allMsEvents.some(mEvent =>
      (mEvent.singleValueExtendedProperties || []).some(p =>
        p.id.includes("Name original_google_id") && p.value === gEvent.id
      )
    );
    if (existsInMsByProp) { syncedToMs[gEvent.id] = true; continue; }

    // 判定3: 終日イベント — 日付＋件名で照合（$expand 不安定時の補完）
    // isAllDay チェックを外す: タイムゾーン処理によって false で返るケースがあるため
    if (gEvent.start.date) {
      const existsByTitleDate = allMsEvents.some(m =>
        m.subject === (gEvent.summary || "(タイトルなし)") &&
        m.start?.dateTime?.substring(0, 10) === gEvent.start.date
      );
      if (existsByTitleDate) { syncedToMs[gEvent.id] = true; continue; }
    } else {
      // 判定3': 時刻ありイベント — 開始日時（1分以内）＋件名で照合
      // MSはタイムゾーンなしのローカル時刻を返すため msDatetimeToUTC で正規化して比較
      const gStartMs = gEvent.start.dateTime ? new Date(gEvent.start.dateTime).getTime() : null;
      if (gStartMs) {
        const existsByTitleTime = allMsEvents.some(m =>
          !m.isAllDay &&
          m.subject === (gEvent.summary || "(タイトルなし)") &&
          m.start?.dateTime &&
          Math.abs(msDatetimeToUTC(m.start.dateTime, m.start.timeZone) - gStartMs) < 60000
        );
        if (existsByTitleTime) { syncedToMs[gEvent.id] = true; continue; }
      }
    }

    console.log(`[Google -> MS] 追加: ${gEvent.summary}`);
    const okToMs = await createEventInMicrosoft(msToken, gEvent);
    if (okToMs) {
      syncedToMs[gEvent.id] = true;
      googleToMs++;
    } else {
      googleToMsFailed++;
    }
  }

  // ── MS -> Google ──────────────────────────────────────────
  for (const mEvent of msEventsToProcess) {
    // Google由来イベントはスキップ（無限ループ防止）
    const extProps = mEvent.singleValueExtendedProperties || [];
    const syncProp = extProps.find(p => p.id.includes("Name synced_from"));
    if (syncProp?.value === 'google') continue;

    // 判定1: ローカルキャッシュ（最速・最優先）
    if (syncedToGoogle[mEvent.id]) continue;

    // 判定2: 全Googleイベントで original_ms_id を照合（穴①の修正: フィルタなし全件参照）
    const existsInGoogle = allGoogleEvents.some(g =>
      g.extendedProperties?.private?.original_ms_id === mEvent.id
    );
    if (existsInGoogle) { syncedToGoogle[mEvent.id] = true; continue; }

    // 判定3: タイトル＋日時で照合（$expand 不安定・キャッシュなし時の補完）
    // MSはタイムゾーンなしのローカル時刻を返す。GoogleはZ付きで保存されるため、
    // 先頭19文字（YYYY-MM-DDTHH:mm:ss）を統一フォーマットとして比較する
    if (mEvent.start?.dateTime) {
      const mDate = mEvent.start.dateTime.substring(0, 10);
      const mNorm = mEvent.start.dateTime.substring(0, 19); // "YYYY-MM-DDTHH:mm:ss"
      const existsByTitleMatch = allGoogleEvents.some(g => {
        if ((g.summary || "(タイトルなし)") !== (mEvent.subject || "(タイトルなし)")) return false;
        // 終日Google予定: 日付で比較
        if (g.start?.date) return g.start.date === mDate;
        // 時刻ありGoogle予定: タイムゾーン記号を除いた先頭19文字で1分以内の一致を確認
        if (g.start?.dateTime) {
          const gNorm = g.start.dateTime.substring(0, 19);
          return Math.abs(new Date(gNorm).getTime() - new Date(mNorm).getTime()) < 60000;
        }
        return false;
      });
      if (existsByTitleMatch) { syncedToGoogle[mEvent.id] = true; continue; }
    }

    console.log(`[MS -> Google] 追加: ${mEvent.subject}`);
    const okToGoogle = await createEventInGoogle(googleToken, mEvent);
    if (okToGoogle) {
      syncedToGoogle[mEvent.id] = true;
      msToGoogle++;
    } else {
      msToGoogleFailed++;
    }
  }

  // キャッシュと lastSyncTime を保存・MSトークン期限切れフラグをクリア
  await chrome.storage.local.set({ syncedToMs, syncedToGoogle, lastSyncTime: now.toISOString(), msTokenExpired: false });
  await chrome.action.setBadgeText({ text: '' });

  await saveSyncLog({ success: true, timeMin, timeMax, googleToMs, msToGoogle,
    googleToMsFailed, msToGoogleFailed });
  console.log(`同期処理が完了しました。(G→M: ${googleToMs}件成功/${googleToMsFailed}件失敗, M→G: ${msToGoogle}件成功/${msToGoogleFailed}件失敗)`);
}

async function saveSyncLog(entry) {
  try {
    const { syncLogs = [] } = await chrome.storage.local.get('syncLogs');
    syncLogs.unshift({ ...entry, timestamp: new Date().toISOString() });
    if (syncLogs.length > 10) syncLogs.length = 10; // 最新10件のみ保持
    await chrome.storage.local.set({ syncLogs });
  } catch (err) {
    // ストレージ書き込み失敗でもログ保存の失敗が同期処理全体を壊さないようにする
    console.error("ログの保存に失敗:", err.message);
  }
}

// ==========================================
// 5. API通信処理
// ==========================================

// MSが返すnaive日時文字列（タイムゾーンなし）をUTCエポックmsに変換する
// Intl.DateTimeFormat でオフセットを計算し、システムTZに依存しない比較を実現する
function msDatetimeToUTC(datetimeStr, timezoneName) {
  const naive = datetimeStr.substring(0, 19); // "YYYY-MM-DDTHH:mm:ss"
  if (!timezoneName || timezoneName.toUpperCase() === 'UTC') {
    return new Date(naive + 'Z').getTime();
  }
  try {
    const asUTCEpoch = new Date(naive + 'Z').getTime();
    const parts = new Intl.DateTimeFormat('en-US', {
      timeZone: timezoneName,
      year: 'numeric', month: '2-digit', day: '2-digit',
      hour: '2-digit', minute: '2-digit', second: '2-digit',
      hour12: false
    }).formatToParts(new Date(asUTCEpoch));
    const get = t => parts.find(p => p.type === t).value;
    const localStr = `${get('year')}-${get('month')}-${get('day')}T${get('hour')}:${get('minute')}:${get('second')}`;
    const offset = new Date(localStr + 'Z').getTime() - asUTCEpoch;
    return asUTCEpoch - offset;
  } catch {
    // タイムゾーン名が不明な場合はシステムローカル時刻として解析（フォールバック）
    return new Date(naive).getTime();
  }
}

async function fetchGoogleEvents(token, timeMin, timeMax) {
  let allItems = [];
  let pageToken = null;

  do {
    const url = new URL('https://www.googleapis.com/calendar/v3/calendars/primary/events');
    url.searchParams.append('timeMin', timeMin);
    url.searchParams.append('timeMax', timeMax);
    url.searchParams.append('singleEvents', 'true');
    url.searchParams.append('maxResults', '2500');
    if (pageToken) url.searchParams.append('pageToken', pageToken);

    const response = await fetch(url.toString(), {
      headers: { 'Authorization': `Bearer ${token}` }
    });
    if (!response.ok) throw new Error("Google Calendar API Error: " + response.statusText);
    const data = await response.json();
    allItems = allItems.concat(data.items || []);
    pageToken = data.nextPageToken || null;
  } while (pageToken);

  return allItems;
}

async function fetchMicrosoftEvents(token, timeMin, timeMax) {
  let allEvents = [];

  const initialUrl = new URL('https://graph.microsoft.com/v1.0/me/calendarView');
  initialUrl.searchParams.append('startDateTime', timeMin);
  initialUrl.searchParams.append('endDateTime', timeMax);
  initialUrl.searchParams.append('$top', '500');
  // カスタム拡張プロパティも一緒に取得するための$expand
  initialUrl.searchParams.append('$expand', "singleValueExtendedProperties($filter=id eq 'String {66f5a359-4659-4830-9070-00040ec6ac6e} Name original_google_id' or id eq 'String {66f5a359-4659-4830-9070-00040ec6ac6e} Name synced_from')");

  let nextUrl = initialUrl.toString();

  while (nextUrl) {
    const response = await fetch(nextUrl, {
      headers: { 'Authorization': `Bearer ${token}` }
    });

    // ① 401 = トークン期限切れ → フラグ＆バッジをセット
    if (response.status === 401) {
      await chrome.storage.local.set({ msTokenExpired: true });
      await chrome.action.setBadgeText({ text: '!' });
      await chrome.action.setBadgeBackgroundColor({ color: '#e53935' });
      isSyncing = false;
      throw new Error('MS_TOKEN_EXPIRED');
    }

    if (!response.ok) throw new Error("Microsoft Graph API Error: " + response.statusText);
    const data = await response.json();
    allEvents = allEvents.concat(data.value || []);
    nextUrl = data['@odata.nextLink'] || null;
  }

  return allEvents;
}

async function createEventInMicrosoft(token, gEvent) {
  // ISO8601フォーマットだが、Graph APIのDateTimeTimeZone要件に合わせて整形
  const formatDateTime = (dt) => {
    if (dt.dateTime) {
      return { dateTime: dt.dateTime, timeZone: dt.timeZone || 'UTC' };
    } else if (dt.date) {
      return { dateTime: dt.date + 'T00:00:00', timeZone: 'UTC' };
    }
    return null;
  };

  const start = formatDateTime(gEvent.start);
  const end = formatDateTime(gEvent.end);
  if (!start || !end) return false; // 日時がない予定は無視

  const payload = {
    subject: gEvent.summary || "(タイトルなし)",
    body: {
      contentType: "HTML",
      content: gEvent.description || ""
    },
    start: start,
    end: end,
    isAllDay: !!gEvent.start.date,
    singleValueExtendedProperties: [
      {
        id: "String {66f5a359-4659-4830-9070-00040ec6ac6e} Name original_google_id",
        value: gEvent.id
      },
      {
        id: "String {66f5a359-4659-4830-9070-00040ec6ac6e} Name synced_from",
        value: "google"
      }
    ]
  };

  const response = await fetch('https://graph.microsoft.com/v1.0/me/events', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(payload)
  });
  if (!response.ok) {
    const errData = await response.text();
    console.error("MSへの追加に失敗:", errData);
    return false;
  }
  return true;
}

async function createEventInGoogle(token, mEvent) {
  // isAllDay が不安定な場合の補完: 開始・終了がともに T00:00:00（深夜0時）なら終日とみなす
  const isEffectivelyAllDay = mEvent.isAllDay ||
    (/T00:00:00/.test(mEvent.start?.dateTime || '') &&
     /T00:00:00/.test(mEvent.end?.dateTime || ''));

  const formatDateTime = (dt, isAllDay) => {
    if (isAllDay) {
      return { date: dt.dateTime.split('T')[0] };
    } else {
      // MSはローカル時刻（タイムゾーンなし）を返すため、先頭19文字のみ使用し'Z'は付けない
      return { dateTime: dt.dateTime.substring(0, 19) };
    }
  };

  const payload = {
    summary: mEvent.subject || "(タイトルなし)",
    description: mEvent.body?.content || "",
    start: formatDateTime(mEvent.start, isEffectivelyAllDay),
    end: formatDateTime(mEvent.end, isEffectivelyAllDay),
    extendedProperties: {
      private: {
        original_ms_id: mEvent.id,
        synced_from: "microsoft"
      }
    }
  };

  const response = await fetch('https://www.googleapis.com/calendar/v3/calendars/primary/events', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(payload)
  });
  
  if (!response.ok) {
    const errData = await response.text();
    console.error("Googleへの追加に失敗:", errData);
    return false;
  }
  return true;
}
