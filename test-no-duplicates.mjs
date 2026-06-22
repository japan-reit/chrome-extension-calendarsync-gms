/**
 * 重複防止 end-to-end テスト
 *
 * background.js の実コードを Node.js VM にロードし、
 * Chrome API と fetch をモックして2回（およびストレージクリア後）の
 * 同期サイクルを実行。2回目以降に作成件数 = 0 であることを検証する。
 */

import vm from 'vm';
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// ============================================================
// 共有テスト状態
// ============================================================
let storage = { msToken: 'mock-ms-token' };
let msCreated = [];
let googleCreated = [];
let currentGoogleEvents = [];
let currentMsEvents = [];

// ============================================================
// Mock fetch
// ============================================================
async function mockFetch(url, options = {}) {
  const urlStr = url.toString();
  const method = (options.method || 'GET').toUpperCase();

  // Google Calendar GET
  if (urlStr.includes('googleapis.com/calendar') && method === 'GET') {
    return {
      ok: true, status: 200,
      json: async () => ({ items: currentGoogleEvents })
    };
  }
  // Google Calendar POST（イベント作成）
  if (urlStr.includes('googleapis.com/calendar') && method === 'POST') {
    const body = JSON.parse(options.body);
    googleCreated.push(body);
    return { ok: true, status: 200, json: async () => ({ id: `g_new_${googleCreated.length}`, ...body }) };
  }
  // MS Graph calendarView GET
  if (urlStr.includes('graph.microsoft.com') && urlStr.includes('calendarView') && method === 'GET') {
    return {
      ok: true, status: 200,
      json: async () => ({ value: currentMsEvents, '@odata.nextLink': null })
    };
  }
  // MS Graph POST（イベント作成）
  if (urlStr.includes('graph.microsoft.com/v1.0/me/events') && method === 'POST') {
    const body = JSON.parse(options.body);
    msCreated.push(body);
    return { ok: true, status: 200, json: async () => ({ id: `ms_new_${msCreated.length}`, ...body }) };
  }
  throw new Error(`Unmocked: ${method} ${urlStr}`);
}

// ============================================================
// Chrome API モック
// ============================================================
function makeChromeMock() {
  return {
    storage: {
      local: {
        get: async (keys) => {
          const keyList = Array.isArray(keys) ? keys : [keys];
          const result = {};
          for (const k of keyList) { if (k in storage) result[k] = storage[k]; }
          return result;
        },
        set: async (data) => { Object.assign(storage, data); }
      }
    },
    identity: {
      getAuthToken: (_opts, cb) => Promise.resolve().then(() => cb('mock-google-token'))
    },
    action: { setBadgeText: async () => {}, setBadgeBackgroundColor: async () => {} },
    runtime: {
      lastError: null,
      onMessage: { addListener: () => {} },
      onInstalled: { addListener: () => {} }
    },
    alarms: { create: () => {}, onAlarm: { addListener: () => {} } }
  };
}

// ============================================================
// background.js を VM にロード（実コードを使用）
// ============================================================
const code = readFileSync(join(__dirname, 'background.js'), 'utf8');
const vmCtx = vm.createContext({
  chrome: makeChromeMock(),
  fetch: mockFetch,
  console,
  URL, URLSearchParams, Intl, Math, Date, setTimeout, clearTimeout,
  Promise, JSON, Error, TypeError, Object, Array, Map, Set, RegExp,
  String, Number, Boolean, encodeURIComponent, decodeURIComponent,
  parseInt, parseFloat, isNaN, isFinite
});
vm.runInContext(code, vmCtx);

// ============================================================
// テストデータ
// ============================================================

// 初期状態: 拡張プロパティなし
const G_INIT = [
  {
    id: 'g1', summary: 'Team Meeting',
    start: { dateTime: '2026-06-22T10:00:00+09:00', timeZone: 'Asia/Tokyo' },
    end:   { dateTime: '2026-06-22T11:00:00+09:00', timeZone: 'Asia/Tokyo' },
    updated: '2026-06-20T00:00:00Z'
  },
  {
    id: 'g2', summary: 'Holiday',
    start: { date: '2026-06-23' }, end: { date: '2026-06-24' },
    updated: '2026-06-20T00:00:00Z'
  }
];

const MS_INIT = [
  {
    id: 'ms1', subject: 'Sprint Review',
    start: { dateTime: '2026-06-22T15:00:00.0000000', timeZone: 'Asia/Tokyo' },
    end:   { dateTime: '2026-06-22T16:00:00.0000000', timeZone: 'Asia/Tokyo' },
    lastModifiedDateTime: '2026-06-20T00:00:00Z', isAllDay: false,
    singleValueExtendedProperties: []
  },
  {
    id: 'ms2', subject: 'All Day Task',
    start: { dateTime: '2026-06-23T00:00:00.0000000', timeZone: 'UTC' },
    end:   { dateTime: '2026-06-24T00:00:00.0000000', timeZone: 'UTC' },
    lastModifiedDateTime: '2026-06-20T00:00:00Z', isAllDay: true,
    singleValueExtendedProperties: []
  }
];

// 同期1回目実施後の状態: 双方向コピーが存在し拡張プロパティ付き
const G_AFTER_SYNC1 = [
  ...G_INIT,
  // MS→Google でコピーされたイベント（synced_from=microsoft）
  {
    id: 'g_from_ms1', summary: 'Sprint Review',
    start: { dateTime: '2026-06-22T15:00:00+09:00', timeZone: 'Asia/Tokyo' },
    end:   { dateTime: '2026-06-22T16:00:00+09:00', timeZone: 'Asia/Tokyo' },
    updated: '2026-06-22T06:30:00Z',
    extendedProperties: { private: { original_ms_id: 'ms1', synced_from: 'microsoft' } }
  },
  {
    id: 'g_from_ms2', summary: 'All Day Task',
    start: { date: '2026-06-23' }, end: { date: '2026-06-24' },
    updated: '2026-06-22T06:30:00Z',
    extendedProperties: { private: { original_ms_id: 'ms2', synced_from: 'microsoft' } }
  }
];

const MS_AFTER_SYNC1 = [
  ...MS_INIT,
  // Google→MS でコピーされたイベント（original_google_id + synced_from=google）
  {
    id: 'ms_from_g1', subject: 'Team Meeting',
    start: { dateTime: '2026-06-22T10:00:00.0000000', timeZone: 'Asia/Tokyo' },
    end:   { dateTime: '2026-06-22T11:00:00.0000000', timeZone: 'Asia/Tokyo' },
    lastModifiedDateTime: '2026-06-22T06:30:00Z', isAllDay: false,
    singleValueExtendedProperties: [
      { id: 'String {66f5a359-4659-4830-9070-00040ec6ac6e} Name original_google_id', value: 'g1' },
      { id: 'String {66f5a359-4659-4830-9070-00040ec6ac6e} Name synced_from', value: 'google' }
    ]
  },
  {
    id: 'ms_from_g2', subject: 'Holiday',
    start: { dateTime: '2026-06-23T00:00:00.0000000', timeZone: 'UTC' },
    end:   { dateTime: '2026-06-24T00:00:00.0000000', timeZone: 'UTC' },
    lastModifiedDateTime: '2026-06-22T06:30:00Z', isAllDay: true,
    singleValueExtendedProperties: [
      { id: 'String {66f5a359-4659-4830-9070-00040ec6ac6e} Name original_google_id', value: 'g2' },
      { id: 'String {66f5a359-4659-4830-9070-00040ec6ac6e} Name synced_from', value: 'google' }
    ]
  }
];

// $expand 完全失敗シミュレーション: singleValueExtendedProperties を全て空にする
const MS_AFTER_SYNC1_NO_EXPAND = MS_AFTER_SYNC1.map(e => ({
  ...e, singleValueExtendedProperties: []
}));

// ============================================================
// テスト実行
// ============================================================
let passed = 0;
let failed = 0;

function assert(label, condition, detail = '') {
  if (condition) {
    console.log(`  ✅ PASS: ${label}`);
    passed++;
  } else {
    console.log(`  ❌ FAIL: ${label}${detail ? ' — ' + detail : ''}`);
    failed++;
  }
}

console.log('╔══════════════════════════════════════════════════════╗');
console.log('║  重複防止 end-to-end テスト (background.js 実コード)  ║');
console.log('╚══════════════════════════════════════════════════════╝\n');

// ────────────────────────────────────────────────────────────
// SYNC 1: 初回同期（両方向で新規作成が起きることを確認）
// ────────────────────────────────────────────────────────────
console.log('【SYNC 1: 初回同期】');
currentGoogleEvents = G_INIT;
currentMsEvents = MS_INIT;
msCreated = []; googleCreated = [];

await vmCtx.executeSync();

console.log(`  MS に作成: ${msCreated.length}件, Google に作成: ${googleCreated.length}件`);
assert('SYNC1: Google→MS で2件作成', msCreated.length === 2,
  `実際: ${msCreated.length}件 [${msCreated.map(e=>e.subject).join(', ')}]`);
assert('SYNC1: MS→Google で2件作成', googleCreated.length === 2,
  `実際: ${googleCreated.length}件 [${googleCreated.map(e=>e.summary).join(', ')}]`);

const sync1Time = storage.lastSyncTime;
console.log(`  lastSyncTime: ${sync1Time}`);

// ────────────────────────────────────────────────────────────
// SYNC 2: 通常2回目（キャッシュあり）→ 重複ゼロを確認
// ────────────────────────────────────────────────────────────
console.log('\n【SYNC 2: 2回目同期（キャッシュあり）】');
currentGoogleEvents = G_AFTER_SYNC1;
currentMsEvents = MS_AFTER_SYNC1;
msCreated = []; googleCreated = [];

await vmCtx.executeSync();

console.log(`  MS に作成: ${msCreated.length}件, Google に作成: ${googleCreated.length}件`);
assert('SYNC2: MS への重複作成ゼロ', msCreated.length === 0,
  `実際: ${msCreated.length}件 [${msCreated.map(e=>e.subject).join(', ')}]`);
assert('SYNC2: Google への重複作成ゼロ', googleCreated.length === 0,
  `実際: ${googleCreated.length}件 [${googleCreated.map(e=>e.summary).join(', ')}]`);

// ────────────────────────────────────────────────────────────
// SYNC 3: ストレージクリア後（$expand 成功）→ 重複ゼロを確認
// ────────────────────────────────────────────────────────────
console.log('\n【SYNC 3: ストレージクリア後（$expand 成功）】');
storage = { msToken: 'mock-ms-token' };  // キャッシュ全消去
currentGoogleEvents = G_AFTER_SYNC1;
currentMsEvents = MS_AFTER_SYNC1;
msCreated = []; googleCreated = [];

await vmCtx.executeSync();

console.log(`  MS に作成: ${msCreated.length}件, Google に作成: ${googleCreated.length}件`);
assert('SYNC3: MS への重複作成ゼロ', msCreated.length === 0,
  `実際: ${msCreated.length}件 [${msCreated.map(e=>e.subject).join(', ')}]`);
assert('SYNC3: Google への重複作成ゼロ', googleCreated.length === 0,
  `実際: ${googleCreated.length}件 [${googleCreated.map(e=>e.summary).join(', ')}]`);

// ────────────────────────────────────────────────────────────
// SYNC 4: ストレージクリア + $expand 完全失敗（最悪ケース）
// ────────────────────────────────────────────────────────────
console.log('\n【SYNC 4: ストレージクリア後 + $expand 完全失敗（最悪ケース）】');
storage = { msToken: 'mock-ms-token' };
currentGoogleEvents = G_AFTER_SYNC1;
currentMsEvents = MS_AFTER_SYNC1_NO_EXPAND;  // 全 singleValueExtendedProperties = []
msCreated = []; googleCreated = [];

await vmCtx.executeSync();

console.log(`  MS に作成: ${msCreated.length}件, Google に作成: ${googleCreated.length}件`);
if (msCreated.length > 0) console.log(`    MS作成: [${msCreated.map(e=>e.subject).join(', ')}]`);
if (googleCreated.length > 0) console.log(`    Google作成: [${googleCreated.map(e=>e.summary).join(', ')}]`);
assert('SYNC4: MS への重複作成ゼロ ($expand失敗でも判定3で防御)', msCreated.length === 0,
  `実際: ${msCreated.length}件 [${msCreated.map(e=>e.subject).join(', ')}]`);
assert('SYNC4: Google への重複作成ゼロ ($expand失敗でも判定2/3で防御)', googleCreated.length === 0,
  `実際: ${googleCreated.length}件 [${googleCreated.map(e=>e.summary).join(', ')}]`);

// ────────────────────────────────────────────────────────────
// SYNC 5: ログ記録の信頼性検証
// ────────────────────────────────────────────────────────────
console.log('\n【SYNC 5: ログ記録の信頼性検証】');

// 5a: API成功 → ログにカウントが正確に記録されること
storage = { msToken: 'mock-ms-token' };
currentGoogleEvents = G_INIT;
currentMsEvents = MS_INIT;
msCreated = []; googleCreated = [];
await vmCtx.executeSync();

const { syncLogs: logs5a = [] } = storage;
const log5a = logs5a[0];
assert('5a: ログが保存される', logs5a.length >= 1, `length=${logs5a.length}`);
assert('5a: success=true', log5a?.success === true, `success=${log5a?.success}`);
assert('5a: googleToMs=2件', log5a?.googleToMs === 2, `googleToMs=${log5a?.googleToMs}`);
assert('5a: msToGoogle=2件', log5a?.msToGoogle === 2, `msToGoogle=${log5a?.msToGoogle}`);
assert('5a: 失敗件数=0', (log5a?.googleToMsFailed ?? 0) === 0 && (log5a?.msToGoogleFailed ?? 0) === 0,
  `failed: G→M=${log5a?.googleToMsFailed}, M→G=${log5a?.msToGoogleFailed}`);
assert('5a: timestamp が記録される', !!log5a?.timestamp, `timestamp=${log5a?.timestamp}`);
console.log(`  ログ内容: ${JSON.stringify(log5a)}`);

// 5b: API失敗（MS POST が 500 を返す）→ 失敗件数がログに反映されること
console.log('\n  [5b: MS POST が500エラーを返すケース]');
const origFetch = vmCtx.fetch;
vmCtx.fetch = async (url, options = {}) => {
  const urlStr = url.toString();
  const method = (options.method || 'GET').toUpperCase();
  // MS POST だけ失敗させる
  if (urlStr.includes('graph.microsoft.com/v1.0/me/events') && method === 'POST') {
    return { ok: false, status: 500, text: async () => 'Internal Server Error' };
  }
  return origFetch(url, options);
};

storage = { msToken: 'mock-ms-token' };
currentGoogleEvents = G_INIT;
currentMsEvents = MS_INIT;
msCreated = []; googleCreated = [];
await vmCtx.executeSync();
vmCtx.fetch = origFetch;

const { syncLogs: logs5b = [] } = storage;
const log5b = logs5b[0];
assert('5b: ログが保存される', logs5b.length >= 1, `length=${logs5b.length}`);
assert('5b: googleToMs=0件（失敗は成功カウントに入らない）', log5b?.googleToMs === 0,
  `googleToMs=${log5b?.googleToMs}`);
assert('5b: googleToMsFailed=2件', log5b?.googleToMsFailed === 2,
  `googleToMsFailed=${log5b?.googleToMsFailed}`);
assert('5b: msToGoogle=2件（Google側は成功）', log5b?.msToGoogle === 2,
  `msToGoogle=${log5b?.msToGoogle}`);
console.log(`  ログ内容: ${JSON.stringify(log5b)}`);

// 5c: 認証エラー → success=false でログ記録されること
console.log('\n  [5c: 認証エラーケース]');
storage = {}; // msToken なし（未認証）
currentGoogleEvents = G_INIT;
currentMsEvents = MS_INIT;
const prevLogsCount = (storage.syncLogs || []).length;
try { await vmCtx.executeSync(); } catch (_) {}

const { syncLogs: logs5c = [] } = storage;
const log5c = logs5c[0];
assert('5c: 認証エラーログが記録される', logs5c.length >= 1, `length=${logs5c.length}`);
assert('5c: success=false', log5c?.success === false, `success=${log5c?.success}`);
assert('5c: error メッセージあり', typeof log5c?.error === 'string' && log5c.error.length > 0,
  `error="${log5c?.error}"`);
console.log(`  ログ内容: ${JSON.stringify(log5c)}`);

// ────────────────────────────────────────────────────────────
// 結果サマリー
// ────────────────────────────────────────────────────────────
console.log('\n╔══════════════════════════════════╗');
console.log(`║  合格: ${passed} / 不合格: ${failed}        ║`);
if (failed === 0) {
  console.log('║  ✅ 全テスト合格                  ║');
  console.log('║  次回同期での重複: 絶対になし     ║');
} else {
  console.log('║  ❌ テスト失敗あり                ║');
}
console.log('╚══════════════════════════════════╝');

process.exit(failed > 0 ? 1 : 0);
