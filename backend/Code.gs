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
  SUMMARIES: '每日總結',
  PERIODS: '期間回顧',
  HABITS: '習慣',
  HABIT_CHECKS: '習慣打卡',
};

const HEADERS = {
  [SHEETS.ENTRIES]: ['id', '日期', '分類ID', '分類名稱', '整理後內容', '原始文字', '建立時間', '更新時間'],
  [SHEETS.CATEGORIES]: ['id', '名稱', '顏色', '排序', '啟用', '描述'],
  [SHEETS.LISTS]: ['id', '名稱', '開始日', '結束日', '項目'],
  [SHEETS.CHECKS]: ['日期', '清單ID', '已完成項目', '更新時間'],
  [SHEETS.SUMMARIES]: ['日期', '內容', '建立時間'],
  [SHEETS.PERIODS]: ['起日', '迄日', '內容', '統計', '建立時間'],
  [SHEETS.HABITS]: ['id', '名稱', '頻率', '設定', '開始日', '狀態', '排序'],
  [SHEETS.HABIT_CHECKS]: ['日期', '習慣ID'],
};

const DEFAULT_CATEGORIES = [
  ['感恩', '#C8794F', '今天讓你覺得開心、想說謝謝的人事物有哪些？'],
  ['覺察', '#6F8FAF', '今天注意到自己的哪些情緒、想法或反應？背後可能是什麼？'],
  ['發現幸福', '#D4A23C', '今天有哪些讓你覺得「好幸福」的小時刻？'],
  ['習慣', '#6E9A78', '今天在想培養的習慣上做得怎麼樣？有什麼發現？'],
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
    DEFAULT_CATEGORIES.forEach((c, i) => cat.appendRow([newId_(), c[0], c[1], String(i + 1), 'TRUE', c[2]]));
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
  init: () => {
    upgrade_();
    return {
      categories: readCategories_(), prompt: promptInfo_(),
      habits: readHabits_(), habitChecks: readHabitChecks_(),
      entryDates: [...new Set(readEntries_().map(e => e.date))].sort(),
    };
  },
  savePrompt: b => savePrompt_(b.kind, b.prompt),
  summarizeDay: b => summarizeDay_(b.date),
  getSummaries: b => readSummaries_().filter(r => r.date >= b.from && r.date <= b.to),
  summarizePeriod: b => summarizePeriod_(b.from, b.to),
  getPeriodSummaries: () => readPeriodSummaries_(),
  getEntries: b => readEntries_().filter(r => r.date >= b.from && r.date <= b.to),
  getEntriesByCategory: b => readEntries_().filter(r => r.categoryId === b.categoryId),
  saveEntry: b => withLock_(() => saveEntry_(b.entry)),
  deleteEntry: b => withLock_(() => deleteRowById_(SHEETS.ENTRIES, b.id)),
  polish: b => polish_(b.text, b.categoryName, b.categoryDesc),
  saveCategories: b => withLock_(() => saveCategories_(b.categories)),
  saveHabit: b => withLock_(() => saveHabit_(b.habit)),
  deleteHabit: b => withLock_(() => deleteHabit_(b.id)),
  setHabitCheck: b => withLock_(() => setHabitCheck_(b.date, b.habitId, b.done)),
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
    .map(r => ({ id: r[0], name: r[1], color: r[2], order: Number(r[3]) || 0, active: r[4] !== 'FALSE', desc: r[5] }))
    .sort((a, b) => a.order - b.order);
}

function saveCategories_(categories) {
  const sh = sheet_(SHEETS.CATEGORIES);
  const rows = categories.map((c, i) => [c.id || newId_(), c.name, c.color, String(i + 1),
    c.active === false ? 'FALSE' : 'TRUE', String(c.desc || '').trim()]);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 6).setNumberFormat('@').setValues(rows);
  return readCategories_();
}

// ───────────────────────── 習慣（每一項有自己的頻率）─────────────────────────
// 頻率：daily 每天／weekdays 每週固定星期（設定 = "1,3"，0 是星期日）／weekly 每週幾次（設定 = 次數）

const WEEKDAY_NAMES = ['日', '一', '二', '三', '四', '五', '六'];

function readHabits_() {
  return readRows_(SHEETS.HABITS).map(r => ({
    id: r[0], name: r[1], freq: r[2] || 'daily',
    days: r[2] === 'weekdays' ? String(r[3]).split(',').filter(x => x !== '').map(Number) : [],
    times: r[2] === 'weekly' ? Number(r[3]) || 1 : 0,
    start: r[4], active: r[5] !== 'archived', order: Number(r[6]) || 0,
  })).sort((a, b) => a.order - b.order);
}

function saveHabit_(h) {
  const name = String(h.name || '').trim();
  if (!name) throw new Error('請輸入名稱');
  if (['daily', 'weekdays', 'weekly'].indexOf(h.freq) < 0) throw new Error('請選擇頻率');
  if (h.freq === 'weekdays' && !(h.days || []).length) throw new Error('請至少選一個星期幾');
  if (!h.start) throw new Error('請選擇開始日');
  const setting = h.freq === 'weekdays' ? h.days.join(',') : h.freq === 'weekly' ? String(Math.min(7, Math.max(1, Number(h.times) || 1))) : '';
  const sh = sheet_(SHEETS.HABITS);
  const row = h.id && findRow_(sh, h.id);
  const order = row ? sh.getRange(row, 7).getDisplayValue() : String(sh.getLastRow());
  const values = [h.id || newId_(), name, h.freq, setting, h.start, h.active === false ? 'archived' : 'active', order];
  if (row) sh.getRange(row, 1, 1, 7).setValues([values]);
  else sh.appendRow(values);
  return readHabits_();
}

function deleteHabit_(id) {
  deleteRowById_(SHEETS.HABITS, id);
  const sh = sheet_(SHEETS.HABIT_CHECKS);
  const data = sh.getDataRange().getDisplayValues();
  for (let i = data.length - 1; i >= 1; i--) if (data[i][1] === id) sh.deleteRow(i + 1);
  return readHabits_();
}

function readHabitChecks_() {
  const out = {};
  readRows_(SHEETS.HABIT_CHECKS).forEach(r => (out[r[1]] = out[r[1]] || []).push(r[0]));
  return out;
}

function setHabitCheck_(date, habitId, done) {
  const sh = sheet_(SHEETS.HABIT_CHECKS);
  const data = sh.getDataRange().getDisplayValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === date && data[i][1] === habitId) {
      if (!done) sh.deleteRow(i + 1);
      return true;
    }
  }
  if (done) sh.appendRow([date, habitId]);
  return true;
}

function freqText_(h) {
  if (h.freq === 'weekdays') return '每週' + h.days.map(d => WEEKDAY_NAMES[d]).join('、');
  if (h.freq === 'weekly') return '每週 ' + h.times + ' 次';
  return '每天';
}

// 日期小工具（字串 yyyy-MM-dd，用 UTC 算避免時區跑掉）
function weekday_(d) { return new Date(d + 'T00:00:00Z').getUTCDay(); }
function addDays_(d, n) { return new Date(new Date(d + 'T00:00:00Z').getTime() + n * 86400000).toISOString().slice(0, 10); }
function weekStart_(d) { return addDays_(d, -((weekday_(d) + 6) % 7)); }
function todayStr_() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd'); }

// 給 AI 看的打卡文字
function habitLinesForDay_(date) {
  const checks = readHabitChecks_();
  const lines = [];
  let checked = 0;
  readHabits_().forEach(h => {
    const dates = checks[h.id] || [];
    const doneToday = dates.indexOf(date) >= 0;
    if (h.start > date || (!h.active && !doneToday)) return;
    if (doneToday) checked++;
    if (h.freq === 'weekly') {
      const ws = weekStart_(date);
      const n = dates.filter(d => d >= ws && d <= date).length;
      lines.push((doneToday ? '✓' : '・') + h.name + '（每週 ' + h.times + ' 次，本週到今天完成 ' + n + ' 次）');
    } else if (h.freq === 'daily' || h.days.indexOf(weekday_(date)) >= 0 || doneToday) {
      lines.push((doneToday ? '✓' : '✗') + h.name + '（' + freqText_(h) + '）');
    }
  });
  return { lines: lines, checked: checked };
}

function habitLinesForPeriod_(from, to) {
  const checks = readHabitChecks_();
  const end0 = [to, todayStr_()].sort()[0];
  const lines = [];
  readHabits_().forEach(h => {
    const s = h.start > from ? h.start : from;
    if (s > end0) return;
    const done = (checks[h.id] || []).filter(d => d >= s && d <= end0).length;
    if (!h.active && !done) return;
    const days = [];
    for (let d = s; d <= end0; d = addDays_(d, 1)) days.push(d);
    if (h.freq === 'daily') lines.push(h.name + '（每天）：完成 ' + done + ' / ' + days.length + ' 天');
    else if (h.freq === 'weekdays') {
      const due = days.filter(d => h.days.indexOf(weekday_(d)) >= 0).length;
      lines.push(h.name + '（' + freqText_(h) + '）：完成 ' + done + ' / ' + due + ' 次');
    } else {
      const weeks = new Set(days.map(weekStart_)).size;
      lines.push(h.name + '（每週 ' + h.times + ' 次）：共完成 ' + done + ' 次，期間共 ' + weeks + ' 週');
    }
  });
  return lines;
}

// ───────────────────────── 舊版資料升級 ─────────────────────────
// 1. 舊的「清單＋項目」搬成一項一項的習慣（每天），打卡紀錄一起搬
// 2. 分類補上「描述」欄，預設四個分類填入建議描述

function upgrade_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('UPGRADED_V2')) return;
  withLock_(() => {
    if (props.getProperty('UPGRADED_V2')) return;
    const ss = SpreadsheetApp.getActive();

    const cat = sheet_(SHEETS.CATEGORIES);
    cat.getRange(1, 6).setValue('描述').setFontWeight('bold').setBackground('#F3EDE4');
    cat.getRange(1, 6, cat.getMaxRows(), 1).setNumberFormat('@');
    if (cat.getLastRow() > 1) {
      const rows = cat.getRange(2, 1, cat.getLastRow() - 1, 6).getDisplayValues();
      rows.forEach((r, i) => {
        const def = DEFAULT_CATEGORIES.find(c => c[0] === r[1]);
        if (def && !r[5]) cat.getRange(i + 2, 6).setValue(def[2]);
      });
    }

    const habits = sheet_(SHEETS.HABITS);
    const habitChecks = sheet_(SHEETS.HABIT_CHECKS);
    const listSheet = ss.getSheetByName(SHEETS.LISTS);
    const checkSheet = ss.getSheetByName(SHEETS.CHECKS);
    if (listSheet && listSheet.getLastRow() > 1 && habits.getLastRow() < 2) {
      const today = todayStr_();
      const oldChecks = checkSheet && checkSheet.getLastRow() > 1
        ? checkSheet.getRange(2, 1, checkSheet.getLastRow() - 1, 3).getDisplayValues() : [];
      const habitRows = [];
      const checkRows = [];
      listSheet.getRange(2, 1, listSheet.getLastRow() - 1, 5).getDisplayValues().forEach(r => {
        if (!r[0]) return;
        const [listId, , start, end] = r;
        safeJson_(r[4], []).forEach(item => {
          const id = newId_();
          habitRows.push([id, item.name, 'daily', '', start, end < today ? 'archived' : 'active', String(habitRows.length + 1)]);
          oldChecks.forEach(c => {
            if (c[1] === listId && safeJson_(c[2], []).indexOf(item.id) >= 0) checkRows.push([c[0], id]);
          });
        });
      });
      if (habitRows.length) habits.getRange(2, 1, habitRows.length, 7).setValues(habitRows);
      if (checkRows.length) habitChecks.getRange(2, 1, checkRows.length, 2).setValues(checkRows);
    }
    props.setProperty('UPGRADED_V2', '1');
  });
}

// ───────────────────────── Claude 整理 ─────────────────────────

const DEFAULT_PROMPTS = {
  polish: [
    '你是日記整理助手。使用者用語音輸入寫日記，你把轉出來的口語文字整理成通順、好讀的繁體中文日記。',
    '規則：',
    '1. 保留原意、第一人稱、原本的語氣和具體細節，不要加入使用者沒說過的事。',
    '2. 修正聽寫錯字，刪掉「嗯、那個、然後然後」這類贅字和重複。',
    '3. 適當分段；如果講了好幾件不同的事，可以用條列。',
    '4. 不要說教、不要給建議、不要加評論或結語。',
    '5. 只輸出整理後的內文，不要標題、不要前言。',
  ].join('\n'),
  summary: [
    '你是一位睿智、溫和、話不多的人生導師，說話風格像納瓦爾（Naval Ravikant）：簡潔、直指本質，重視長期思維、內在平靜、自我覺察和複利。',
    '使用者會給你他今天的日記和習慣打卡紀錄。讀完之後，用繁體中文寫一段給他的話：',
    '3 到 5 句。先具體點出今天做到的一件事（只根據他寫的內容，不要編造），再給一個有智慧、可以帶著走的提醒。',
    '像信任的朋友一樣真誠，不說教、不灌雞湯、不堆驚嘆號。',
    '不要假裝是納瓦爾本人，也不要直接引用他的名言。只輸出這段話本身，不要標題、不要前言。',
  ].join('\n'),
  period: [
    '你是一位睿智、溫和、話不多的人生導師，說話風格像納瓦爾（Naval Ravikant）：簡潔、直指本質，重視長期思維、內在平靜、自我覺察和複利。',
    '使用者會給你他一段期間的日記、各分類篇數和習慣打卡統計。讀完之後，用繁體中文寫一段給他的話：',
    '4 到 6 句。點出這段時間反覆出現的主題或明顯的變化，具體肯定他做到的事（只根據他寫的內容，不要編造），再給一個可以帶進下一段時間的提醒。',
    '像信任的朋友一樣真誠，不說教、不灌雞湯、不堆驚嘆號。',
    '不要假裝是納瓦爾本人，也不要直接引用他的名言。只輸出這段話本身，不要標題、不要前言，也不用重複列出篇數。',
  ].join('\n'),
};
const PROMPT_KEYS = { polish: 'CUSTOM_PROMPT', summary: 'CUSTOM_SUMMARY_PROMPT', period: 'CUSTOM_PERIOD_PROMPT' };

// 使用者在 App 設定頁改過就用他的版本，沒改過用預設
function getPrompt_(kind) {
  return PropertiesService.getScriptProperties().getProperty(PROMPT_KEYS[kind]) || DEFAULT_PROMPTS[kind];
}

function savePrompt_(kind, prompt) {
  kind = kind || 'polish';
  if (!PROMPT_KEYS[kind]) throw new Error('未知的提示詞種類');
  const props = PropertiesService.getScriptProperties();
  prompt = String(prompt || '').trim();
  if (!prompt || prompt === DEFAULT_PROMPTS[kind]) props.deleteProperty(PROMPT_KEYS[kind]);
  else props.setProperty(PROMPT_KEYS[kind], prompt);
  return promptInfo_();
}

function promptInfo_() {
  return {
    prompt: getPrompt_('polish'), defaultPrompt: DEFAULT_PROMPTS.polish,
    summaryPrompt: getPrompt_('summary'), defaultSummaryPrompt: DEFAULT_PROMPTS.summary,
    periodPrompt: getPrompt_('period'), defaultPeriodPrompt: DEFAULT_PROMPTS.period,
  };
}

function polish_(text, categoryName, categoryDesc) {
  text = String(text || '').trim();
  if (!text) throw new Error('沒有內容可以整理');
  let system = getPrompt_('polish') + '\n\n這篇日記的分類是「' + (categoryName || '日記') + '」。';
  if (categoryDesc) system += '這個分類要寫的是：' + categoryDesc;
  return callClaude_(system, text);
}

// ───────────────────────── 每日總結 ─────────────────────────

function readSummaries_() {
  return readRows_(SHEETS.SUMMARIES).map(r => ({ date: r[0], content: r[1], createdAt: r[2] }));
}

function summarizeDay_(date) {
  if (!date) throw new Error('缺少日期');
  const entries = readEntries_().filter(e => e.date === date);
  const lines = ['日期：' + date, '', '【日記】'];
  if (entries.length) entries.forEach(e => lines.push('［' + e.categoryName + '］' + e.content, ''));
  else lines.push('（這天沒有寫日記）', '');

  const habit = habitLinesForDay_(date);
  if (habit.lines.length) lines.push('【習慣打卡】', ...habit.lines);
  if (!entries.length && !habit.checked) throw new Error('這天還沒有日記或打卡，先寫一點再來整理吧');

  const content = callClaude_(getPrompt_('summary'), lines.join('\n'));
  return withLock_(() => {
    const sh = sheet_(SHEETS.SUMMARIES);
    const now = nowStr_();
    const data = sh.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === date) {
        sh.getRange(i + 1, 1, 1, 3).setValues([[date, content, now]]);
        return { date: date, content: content, createdAt: now };
      }
    }
    sh.appendRow([date, content, now]);
    return { date: date, content: content, createdAt: now };
  });
}

// ───────────────────────── 期間回顧 ─────────────────────────

function readPeriodSummaries_() {
  return readRows_(SHEETS.PERIODS)
    .map(r => ({ from: r[0], to: r[1], content: r[2], stats: safeJson_(r[3], null), createdAt: r[4] }))
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

function summarizePeriod_(from, to) {
  if (!from || !to || to < from) throw new Error('請選擇正確的期間');
  if ((new Date(to) - new Date(from)) / 86400000 > 366) throw new Error('期間最長一年');
  const entries = readEntries_()
    .filter(e => e.date >= from && e.date <= to)
    .sort((a, b) => (a.date + a.createdAt).localeCompare(b.date + b.createdAt));
  if (!entries.length) throw new Error('這段期間沒有日記');

  const stats = periodStats_(entries);
  const lines = ['期間：' + from + ' 到 ' + to, '', '【各分類篇數】'];
  stats.categories.forEach(c => lines.push(c.name + '：' + c.count + ' 則'));
  lines.push('共 ' + stats.total + ' 則，有寫日記的天數：' + stats.days + ' 天', '');

  const habitLines = habitLinesForPeriod_(from, to);
  if (habitLines.length) lines.push('【習慣打卡】', ...habitLines, '');

  lines.push('【日記】');
  entries.forEach(e => lines.push(e.date + '［' + e.categoryName + '］' + e.content, ''));

  const content = callClaude_(getPrompt_('period'), lines.join('\n'));
  const now = nowStr_();
  withLock_(() => sheet_(SHEETS.PERIODS).appendRow([from, to, content, JSON.stringify(stats), now]));
  return { from: from, to: to, content: content, stats: stats, createdAt: now };
}

// 篇數由程式算，不交給 AI，數字才會準
function periodStats_(entries) {
  const cats = readCategories_();
  const counts = {};
  entries.forEach(e => counts[e.categoryId] = (counts[e.categoryId] || 0) + 1);
  const list = cats
    .filter(c => c.active || counts[c.id])
    .map(c => ({ name: c.name, color: c.color, count: counts[c.id] || 0 }));
  Object.keys(counts).forEach(id => {
    if (!cats.some(c => c.id === id)) {
      const e = entries.find(x => x.categoryId === id);
      list.push({ name: e.categoryName, color: '#999999', count: counts[id] });
    }
  });
  return { total: entries.length, days: new Set(entries.map(e => e.date)).size, categories: list };
}

function callClaude_(system, text) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('還沒設定 Claude 金鑰，請到試算表選單「📔 日記設定 → ③ 設定 Claude 金鑰」。');

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
  if (body.stop_reason === 'refusal') throw new Error('Claude 這次沒辦法處理這段內容，請稍後再試。');
  const out = (body.content || []).filter(b => b.type === 'text').map(b => b.text).join('').trim();
  if (!out) throw new Error('Claude 沒有回傳內容，請再試一次。');
  return out;
}

// ───────────────────────── 共用工具 ─────────────────────────

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  if ([SHEETS.SUMMARIES, SHEETS.PERIODS, SHEETS.HABITS, SHEETS.HABIT_CHECKS].indexOf(name) < 0) throw new Error('找不到「' + name + '」分頁，請先按選單「① 初始化」。');
  // 後來才加的分頁，第一次用到時自動建立
  sh = ss.insertSheet(name);
  sh.appendRow(HEADERS[name]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold').setBackground('#F3EDE4');
  sh.getRange(1, 1, sh.getMaxRows(), HEADERS[name].length).setNumberFormat('@');
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
