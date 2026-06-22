document.addEventListener('DOMContentLoaded', async () => {
  const btnGoogleLogin = document.getElementById('btn-google-login');
  const btnMsLogin = document.getElementById('btn-ms-login');
  const btnSyncNow = document.getElementById('btn-sync-now');
  
  const googleStatus = document.getElementById('google-status');
  const msStatus = document.getElementById('ms-status');
  const syncStatus = document.getElementById('sync-status');

  // 初期状態のチェック
  await updateStatus();

  btnGoogleLogin.addEventListener('click', async () => {
    googleStatus.textContent = '認証中...';
    try {
      const response = await chrome.runtime.sendMessage({ action: 'loginGoogle' });
      if (response && response.success) {
        googleStatus.textContent = 'ログイン済み';
        btnGoogleLogin.textContent = 'Googleで再認証';
      } else {
        googleStatus.textContent = 'エラー: ' + (response?.error || '不明なエラー');
      }
    } catch (err) {
      googleStatus.textContent = 'エラー: ' + err.message;
    }
    await updateStatus(true);
  });

  btnMsLogin.addEventListener('click', async () => {
    msStatus.textContent = '認証中...';
    try {
      const response = await chrome.runtime.sendMessage({ action: 'loginMicrosoft' });
      if (response && response.success) {
        msStatus.textContent = 'ログイン済み';
        btnMsLogin.textContent = 'Microsoftで再認証';
      } else {
        msStatus.textContent = 'エラー: ' + (response?.error || '不明なエラー');
      }
    } catch (err) {
      msStatus.textContent = 'エラー: ' + err.message;
    }
    await updateStatus(true);
  });

  btnSyncNow.addEventListener('click', async () => {
    btnSyncNow.disabled = true;
    syncStatus.textContent = '同期中...';
    try {
      const response = await chrome.runtime.sendMessage({ action: 'syncNow' });
      if (response && response.success) {
        syncStatus.textContent = '同期完了 (' + new Date().toLocaleTimeString() + ')';
      } else {
        syncStatus.textContent = '同期失敗: ' + (response?.error || '不明なエラー');
      }
    } catch (err) {
      syncStatus.textContent = 'エラー: ' + err.message;
    } finally {
      await updateStatus();
      await renderLogs();
    }
  });

  await renderLogs();

  document.getElementById('log-clear').addEventListener('click', async () => {
    await chrome.storage.local.remove('syncLogs');
    await renderLogs();
  });

  async function renderLogs() {
    const { syncLogs = [] } = await chrome.storage.local.get('syncLogs');
    const list = document.getElementById('sync-log-list');
    if (syncLogs.length === 0) {
      list.innerHTML = '<div class="log-empty">ログなし</div>';
      return;
    }
    list.innerHTML = syncLogs.map(log => {
      const time = new Date(log.timestamp).toLocaleString('ja-JP', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' });
      if (!log.success) {
        return `<div class="log-entry log-error"><span class="log-time">${time}</span> ❌ ${log.error}</div>`;
      }
      const from = new Date(log.timeMin).toLocaleDateString('ja-JP', { month: 'numeric', day: 'numeric' });
      const to   = new Date(log.timeMax).toLocaleDateString('ja-JP', { month: 'numeric', day: 'numeric' });
      const failedG = log.googleToMsFailed || 0;
      const failedM = log.msToGoogleFailed || 0;
      const failedPart = (failedG + failedM > 0)
        ? ` <span class="log-failed">⚠️ 失敗 G→M:${failedG} M→G:${failedM}</span>` : '';
      return `<div class="log-entry log-success">
        <span class="log-time">${time}</span> ✅ ${from}〜${to}
        <span class="log-counts">G→M: ${log.googleToMs}件 / M→G: ${log.msToGoogle}件${failedPart}</span>
      </div>`;
    }).join('');
  }

  // preserveMessages=true のとき、エラー中のステータス表示を上書きしない
  async function updateStatus(preserveMessages = false) {
    const { googleToken, msToken, msTokenExpired } = await chrome.storage.local.get(['googleToken', 'msToken', 'msTokenExpired']);

    if (!preserveMessages) {
      if (googleToken) {
        googleStatus.textContent = 'ログイン済み';
        btnGoogleLogin.textContent = 'Googleで再認証';
      } else {
        googleStatus.textContent = '未ログイン';
      }

      if (msTokenExpired) {
        // ① トークン期限切れ警告
        msStatus.textContent = '⚠️ トークン期限切れ — 再認証してください';
        msStatus.style.color = '#c5221f';
        btnMsLogin.textContent = 'Microsoftで再認証';
      } else if (msToken) {
        msStatus.textContent = 'ログイン済み';
        msStatus.style.color = '';
        btnMsLogin.textContent = 'Microsoftで再認証';
      } else {
        msStatus.textContent = '未ログイン';
        msStatus.style.color = '';
      }
    }

    btnSyncNow.disabled = !(googleToken && msToken && !msTokenExpired);
  }
});
