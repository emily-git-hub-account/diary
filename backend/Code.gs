/**
 * 日記 App 後端
 * - 資料存在這份試算表
 * - 網頁版（GitHub Pages）透過 doPost 呼叫這裡
 * - Claude 金鑰、App 密碼存在「指令碼屬性」，不會出現在網頁上
 */

const MODEL = 'claude-opus-5-5';

const SHEETS = {
  ENTRIES: '日記',
  CATEGORIES: '分類',
  LISTS: '清單',
  CHECKS: '打卡',
};

const HEADERS = {
  [SHEETS.ENTRIES]: ['id', '日期', '分類ID', '分類名稱', '整理後內容', '原始文字', '建立時間', '更新時間'],
  [SHEETS.CATEGORIES]: ['id', '名稱', '顏色', '排序', '啟用'],
  [SHEETS.LISTS]: ['id', '名稱', '開始日', '結束日', '項目'],
  [SHEETS.CHECKS]: ['日期', '清單ID', '已完成項目', '更新時間'],
};

const DEFAULT_CATEGORIES = [
  ['感恩', '#C8794F'],
  ['覺察', '#6F8FAF'],
  ['發現幸福', '#D4A23C'],
  ['習慣', '#6E9A78'],
];

// ───────────────────────── 試算表選單 ─────────────────────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📔 日記設定')
    .addItem('① 初始化（第一次用）', 'setup')
    .addItem('② 設定 App 密碼', 'setPassword')
    .addItem('③ 設定 Claude 金鑰', 'setApiKey')
    .addSeparator()
    .addItem('測試 Claude 連線', 'testClaude')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS[name]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold').setBackground('#F3EDE4');
    }
    // 全部存成純文字，避免日期被試算表自動轉格式
    sh.getRange(1, 1, sh.getMaxRows(), HEADERS[name].length).setNumberFormat('@');
  });
  const cat = ss.getSheetByName(SHEETS.CATEGORIES);
  if (cat.getLastRow() === 1) {
    DEFAULT_CATEGORIES.forEach((c, i) => cat.appendRow([newId_(), c[0], c[1], String(i + 1), 'TRUE']));
  }
  const first = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  SpreadsheetApp.getUi().alert('初始化完成 ✅');
}

function setPassword() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('設定 App 密碼', '打開日記 App 時要輸入的密碼（至少 6 個字）：', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const pw = res.getResponseText().trim();
  if (pw.length < 6) return ui.alert('密碼至少要 6 個字，請再設定一次。');
  PropertiesService.getScriptProperties().setProperty('APP_PASSWORD', pw);
  ui.alert('密碼設定好了 ✅');
}

function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('設定 Claude 金鑰', '貼上從 console.anthropic.com 建立的 API Key（sk-ant- 開頭）：', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const key = res.getResponseText().trim();
  if (!/^sk-ant-/.test(key)) return ui.alert('看起來不像 Claude 金鑰（應該是 sk-ant- 開頭），請再確認一次。');
  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_API_KEY', key);
  ui.alert('金鑰設定好了 ✅ 可以按「測試 Claude 連線」確認。');
}

function testClaude() {
  const ui = SpreadsheetApp.getUi();
  try {
    const out = polish_('嗯今天早上那個喝咖啡的時候突然覺得陽光很好然後就覺得很感謝', '感恩');
    ui.alert('連線成功 ✅\n\n整理結果：\n' + out);
  } catch (err) {
    ui.alert('連線失敗 ❌\n\n' + err.message);
  }
}

// ───────────────────────── Web API ─────────────────────────

function doGet() {
  return ContentService.createTextOutput('日記 App 後端運作中');
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: '格式錯誤' });
  }
  try {
    checkAuth_(body.password);
    const handler = ACTIONS[body.action];
    if (!handler) throw new Error('未知的動作：' + body.action);
    return json_({ ok: true, data: handler(body) });
  } catch (err) {
    return json_({ ok: false, error: err.message, auth: err.auth || false });
  }
}

const ACTIONS = {
  login: () => true,
  init: () => ({ categories: readCategories_(), lists: readLists_(), prompt: promptInfo_() }),
  savePrompt: b => savePrompt_(b.prompt),
  getEntries: b => readEntries_().filter(r => r.date >= b.from && r.date <= b.to),
  getEntriesByCategory: b => readEntries_().filter(r => r.categoryId === b.categoryId),
  saveEntry: b => withLock_(() => saveEntry_(b.entry)),
  deleteEntry: b => withLock_(() => deleteRowById_(SHEETS.ENTRIES, b.id)),
  polish: b => polish_(b.text, b.categoryName),
  saveCategories: b => withLock_(() => saveCategories_(b.categories)),
  saveList: b => withLock_(() => saveList_(b.list)),
  deleteList: b => withLock_(() => deleteRowById_(SHEETS.LISTS, b.id)),
  getChecks: b => readChecks_(b.listId),
  setCheck: b => withLock_(() => setCheck_(b.date, b.listId, b.done)),
};

function checkAuth_(password) {
  const props = PropertiesService.getScriptProperties();
  const real = props.getProperty('APP_PASSWORD');
  if (!real) throw authError_('還沒設定密碼，請到試算表選單「📔 日記設定 → ② 設定 App 密碼」。');
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= 10) throw authError_('密碼錯太多次，請 15 分鐘後再試。');
  if (password !== real) {
    cache.put('fails', String(fails + 1), 900);
    throw authError_('密碼不對');
  }
}

function authError_(msg) {
  const err = new Error(msg);
  err.auth = true;
  return err;
}

// ───────────────────────── 日記 ─────────────────────────

function readEntries_() {
  return readRows_(SHEETS.ENTRIES).map(r => ({
    id: r[0], date: r[1], categoryId: r[2], categoryName: r[3],
    content: r[4], raw: r[5], createdAt: r[6], updatedAt: r[7],
  }));
}

function saveEntry_(entry) {
  if (!entry.date || !entry.categoryId || !String(entry.content || '').trim()) {
    throw new Error('日期、分類、內容都要有');
  }
  const sh = sheet_(SHEETS.ENTRIES);
  const now = nowStr_();
  if (entry.id) {
    const row = findRow_(sh, entry.id);
    if (row) {
      const createdAt = sh.getRange(row, 7).getDisplayValue();
      sh.getRange(row, 1, 1, 8).setValues([[entry.id, entry.date, entry.categoryId, entry.categoryName,
        entry.content, entry.raw || '', createdAt, now]]);
      return entry.id;
    }
  }
  const id = newId_();
  sh.appendRow([id, entry.date, entry.categoryId, entry.categoryName, entry.content, entry.raw || '', now, now]);
  return id;
}

// ───────────────────────── 分類 ─────────────────────────

function readCategories_() {
  return readRows_(SHEETS.CATEGORIES)
    .map(r => ({ id: r[0], name: r[1], color: r[2], order: Number(r[3]) || 0, active: r[4] !== 'FALSE' }))
    .sort((a, b) => a.order - b.order);
}

function saveCategories_(categories) {
  const sh = sheet_(SHEETS.CATEGORIES);
  const rows = categories.map((c, i) => [c.id || newId_(), c.name, c.color, String(i + 1), c.active === false ? 'FALSE' : 'TRUE']);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 5).setValues(rows);
  return readCategories_();
}

// ───────────────────────── 清單與打卡 ─────────────────────────

function readLists_() {
  return readRows_(SHEETS.LISTS).map(r => ({
    id: r[0], name: r[1], start: r[2], end: r[3], items: safeJson_(r[4], []),
  }));
}

function saveList_(list) {
  if (!list.name || !list.start || !list.end || !list.items || !list.items.length) {
    throw new Error('名稱、期間、至少一個項目都要有');
  }
  const sh = sheet_(SHEETS.LISTS);
  const id = list.id || newId_();
  const values = [[id, list.name, list.start, list.end, JSON.stringify(list.items)]];
  const row = list.id && findRow_(sh, list.id);
  if (row) sh.getRange(row, 1, 1, 5).setValues(values);
  else sh.appendRow(values[0]);
  return readLists_();
}

function readChecks_(listId) {
  const out = {};
  readRows_(SHEETS.CHECKS).forEach(r => {
    if (r[1] === listId) out[r[0]] = safeJson_(r[2], []);
  });
  return out;
}

function setCheck_(date, listId, done) {
  const sh = sheet_(SHEETS.CHECKS);
  const data = sh.getDataRange().getDisplayValues();
  const values = [[date, listId, JSON.stringify(done || []), nowStr_()]];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === date && data[i][1] === listId) {
      sh.getRange(i + 1, 1, 1, 4).setValues(values);
      return true;
    }
  }
  sh.appendRow(values[0]);
  return true;
}

// ───────────────────────── Claude 整理 ─────────────────────────

const DEFAULT_PROMPT = [
  '你是日記整理助手。使用者用語音輸入寫日記，你把轉出來的口語文字整理成通順、好讀的繁體中文日記。',
  '規則：',
  '1. 保留原意、第一人稱、原本的語氣和具體細節，不要加入使用者沒說過的事。',
  '2. 修正聽寫錯字，刪掉「嗯、那個、然後然後」這類贅字和重複。',
  '3. 適當分段；如果講了好幾件不同的事，可以用條列。',
  '4. 不要說教、不要給建議、不要加評論或結語。',
  '5. 只輸出整理後的內文，不要標題、不要前言。',
].join('\n');

// 使用者在 App 設定頁改過就用她的版本，沒改過用預設
function getPrompt_() {
  return PropertiesService.getScriptProperties().getProperty('CUSTOM_PROMPT') || DEFAULT_PROMPT;
}

function savePrompt_(prompt) {
  const props = PropertiesService.getScriptProperties();
  prompt = String(prompt || '').trim();
  if (!prompt || prompt === DEFAULT_PROMPT) props.deleteProperty('CUSTOM_PROMPT');
  else props.setProperty('CUSTOM_PROMPT', prompt);
  return promptInfo_();
}

function promptInfo_() {
  return { prompt: getPrompt_(), defaultPrompt: DEFAULT_PROMPT };
}

function polish_(text, categoryName) {
  text = String(text || '').trim();
  if (!text) throw new Error('沒有內容可以整理');
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('還沒設定 Claude 金鑰，請到試算表選單「📔 日記設定 → ③ 設定 Claude 金鑰」。');

  const system = getPrompt_() + '\n\n這篇日記的分類是「' + (categoryName || '日記') + '」。';

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      'x-api-key': key,
      'anthropic-version': '2023-06-01',
      'anthropic-beta': 'server-side-fallback-2026-07-01',
    },
    payload: JSON.stringify({
      model: MODEL,
      max_tokens: 16000,
      output_config: { effort: 'low' },
      fallbacks: 'default',
      system: system,
      messages: [{ role: 'user', content: text }],
    }),
  });

  const code = res.getResponseCode();
  const body = safeJson_(res.getContentText(), {});
  if (code !== 200) {
    const msg = (body.error && body.error.message) || res.getContentText().slice(0, 200);
    if (code === 401) throw new Error('Claude 金鑰不正確，請重新設定。');
    if (code === 400 && /credit/i.test(msg)) throw new Error('Claude 帳戶餘額不足，請到 console.anthropic.com 儲值。');
    throw new Error('Claude 整理失敗（' + code + '）：' + msg);
  }
  if (body.stop_reason === 'refusal') throw new Error('Claude 這次沒辦法整理這段內容，請改用原文儲存。');
  const out = (body.content || []).filter(b => b.type === 'text').map(b => b.text).join('').trim();
  if (!out) throw new Error('Claude 沒有回傳內容，請再試一次。');
  return out;
}

// ───────────────────────── 共用工具 ─────────────────────────

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('找不到「' + name + '」分頁，請先按選單「① 初始化」。');
  return sh;
}

function readRows_(name) {
  const sh = sheet_(name);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS[name].length).getDisplayValues().filter(r => r[0] !== '');
}

function findRow_(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return 0;
}

function deleteRowById_(name, id) {
  const sh = sheet_(name);
  const row = findRow_(sh, id);
  if (row) sh.deleteRow(row);
  return true;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function safeJson_(s, fallback) {
  try { return JSON.parse(s); } catch (e) { return fallback; }
}

function newId_() {
  return Utilities.getUuid().slice(0, 8);
}

function nowStr_() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
}
