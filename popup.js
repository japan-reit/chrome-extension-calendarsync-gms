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
      } else {
        googleStatus.textContent = 'ログイン失敗: ' + (response?.error || '不明なエラー');
      }
    } catch (err) {
      googleStatus.textContent = 'エラー: ' + err.message;
    }
    await updateStatus();
  });

  btnMsLogin.addEventListener('click', async () => {
    msStatus.textContent = '認証中...';
    try {
      const response = await chrome.runtime.sendMessage({ action: 'loginMicrosoft' });
      if (response && response.success) {
        msStatus.textContent = 'ログイン済み';
      } else {
        msStatus.textContent = 'ログイン失敗: ' + (response?.error || '不明なエラー');
      }
    } catch (err) {
      msStatus.textContent = 'エラー: ' + err.message;
    }
    await updateStatus();
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
      await updateStatus(); // Update disabled state of sync button
    }
  });

  async function updateStatus() {
    const { googleToken, msToken } = await chrome.storage.local.get(['googleToken', 'msToken']);
    
    if (googleToken) {
      googleStatus.textContent = 'ログイン済み';
      btnGoogleLogin.textContent = 'Googleで再認証';
    } else {
      googleStatus.textContent = '未ログイン';
    }

    if (msToken) {
      msStatus.textContent = 'ログイン済み';
      btnMsLogin.textContent = 'Microsoftで再認証';
    } else {
      msStatus.textContent = '未ログイン';
    }

    // 両方ログインしていれば同期ボタンを有効化
    if (googleToken && msToken) {
      btnSyncNow.disabled = false;
    } else {
      btnSyncNow.disabled = true;
    }
  }
});
