const MICROSOFT_CLIENT_ID = "YOUR_MICROSOFT_CLIENT_ID"; // ユーザーが後で設定する
const SYNC_INTERVAL_MINUTES = 15;

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
  const redirectUri = chrome.identity.getRedirectURL();
  const scope = "Calendars.ReadWrite offline_access";
  const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${MICROSOFT_CLIENT_ID}&response_type=token&redirect_uri=${encodeURIComponent(redirectUri)}&scope=${encodeURIComponent(scope)}`;

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
        await chrome.storage.local.set({ msToken: token });
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
  console.log("同期処理を開始します...");
  
  // 非インタラクティブでトークン取得を試みる
  let googleToken, msToken;
  try {
    googleToken = await getGoogleToken(false);
    // MSはトークンの有効期限管理やリフレッシュフローが必要ですが、
    // 簡略化のため一旦ストレージのキャッシュを利用または再度フローを走らせます。
    // ※実際にはMSトークンのリフレッシュロジックが必要です
    const storageData = await chrome.storage.local.get(['msToken']);
    msToken = storageData.msToken;
    if (!msToken) throw new Error("Microsoftトークンがありません");
  } catch (err) {
    throw new Error("認証が必要です: " + err.message);
  }

  // 同期対象期間（例: 過去7日から向こう30日）
  const now = new Date();
  const timeMin = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000).toISOString();
  const timeMax = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000).toISOString();

  // 1. Googleカレンダーから予定を取得
  const googleEvents = await fetchGoogleEvents(googleToken, timeMin, timeMax);
  
  // 2. Microsoftカレンダーから予定を取得
  const msEvents = await fetchMicrosoftEvents(msToken, timeMin, timeMax);

  // 3. Google -> MS への同期 (Google由来の予定で、まだMSに無いものをコピー)
  for (const gEvent of googleEvents) {
    // 既にMSから同期された予定はスキップ（無限ループ防止）
    if (gEvent.extendedProperties?.private?.synced_from === 'microsoft') {
      continue;
    }
    
    // MS側に既に同等の予定があるかチェック（IDベース）
    const existsInMs = msEvents.some(mEvent => {
      // MSのカスタムプロパティを探す
      const extProps = mEvent.singleValueExtendedProperties || [];
      const prop = extProps.find(p => p.id.includes("String {66f5a359-4659-4830-9070-00040ec6ac6e} Name original_google_id"));
      return prop && prop.value === gEvent.id;
    });

    if (!existsInMs) {
      console.log(`[Google -> MS] 追加: ${gEvent.summary}`);
      await createEventInMicrosoft(msToken, gEvent);
    }
  }

  // 4. MS -> Google への同期 (MS由来の予定で、まだGoogleに無いものをコピー)
  for (const mEvent of msEvents) {
    // 既にGoogleから同期された予定はスキップ
    const extProps = mEvent.singleValueExtendedProperties || [];
    const syncProp = extProps.find(p => p.id.includes("String {66f5a359-4659-4830-9070-00040ec6ac6e} Name synced_from"));
    if (syncProp && syncProp.value === 'google') {
      continue;
    }

    // Google側に既に同等の予定があるかチェック
    const existsInGoogle = googleEvents.some(gEvent => {
      return gEvent.extendedProperties?.private?.original_ms_id === mEvent.id;
    });

    if (!existsInGoogle) {
      console.log(`[MS -> Google] 追加: ${mEvent.subject}`);
      await createEventInGoogle(googleToken, mEvent);
    }
  }

  console.log("同期処理が完了しました。");
}

// ==========================================
// 5. API通信処理
// ==========================================

async function fetchGoogleEvents(token, timeMin, timeMax) {
  const url = new URL('https://www.googleapis.com/calendar/v3/calendars/primary/events');
  url.searchParams.append('timeMin', timeMin);
  url.searchParams.append('timeMax', timeMax);
  url.searchParams.append('singleEvents', 'true'); // 繰り返し予定を展開
  
  const response = await fetch(url.toString(), {
    headers: { 'Authorization': `Bearer ${token}` }
  });
  if (!response.ok) throw new Error("Google Calendar API Error: " + response.statusText);
  const data = await response.json();
  return data.items || [];
}

async function fetchMicrosoftEvents(token, timeMin, timeMax) {
  const url = new URL('https://graph.microsoft.com/v1.0/me/calendarView');
  url.searchParams.append('startDateTime', timeMin);
  url.searchParams.append('endDateTime', timeMax);
  // カスタム拡張プロパティも一緒に取得するための$expand
  url.searchParams.append('$expand', "singleValueExtendedProperties($filter=id eq 'String {66f5a359-4659-4830-9070-00040ec6ac6e} Name original_google_id' or id eq 'String {66f5a359-4659-4830-9070-00040ec6ac6e} Name synced_from')");

  const response = await fetch(url.toString(), {
    headers: { 'Authorization': `Bearer ${token}` }
  });
  if (!response.ok) throw new Error("Microsoft Graph API Error: " + response.statusText);
  const data = await response.json();
  return data.value || [];
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
  if (!start || !end) return; // 日時がない予定は無視

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
  }
}

async function createEventInGoogle(token, mEvent) {
  const formatDateTime = (dt, isAllDay) => {
    if (isAllDay) {
      return { date: dt.dateTime.split('T')[0] };
    } else {
      return { dateTime: dt.dateTime + 'Z' }; // 簡略化のためUTCとして扱うなど調整が必要な場合あり。Graph APIのレスポンスに基づく
    }
  };

  const payload = {
    summary: mEvent.subject || "(タイトルなし)",
    description: mEvent.body?.content || "",
    start: formatDateTime(mEvent.start, mEvent.isAllDay),
    end: formatDateTime(mEvent.end, mEvent.isAllDay),
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
  }
}
