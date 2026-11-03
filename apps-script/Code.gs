/**
 * 每日工作追蹤系統｜Google 試算表 ＋ Apps Script
 *
 * 第一次使用：
 *   1. 試算表選單「代辦系統 → ① 初始化工作表」
 *   2. 在「人員」分頁填姓名、角色（主管／員工）、Email
 *   3. 部署 → 新增部署作業 → 網頁應用程式，把網址貼到「設定」分頁的「系統網址」
 *   4. 選單「代辦系統 → ② 安裝自動排程」
 *
 * 登入：
 *   帳號密碼寫在「人員」分頁的「帳號」「密碼」欄；預設管理者帳號 n821129x／n821129x
 *
 * 權限：
 *   主管：指派、轉派、確認完成、退回、刪除、看全部與統計圖表
 *   員工：自行新增工作、開始、回報、提報完成（送主管確認），只看得到自己的工作
 */

const SHEET = { PEOPLE: '人員', TASKS: '代辦事項', LOG: '回報歷程', CONFIG: '設定' };
const TZ = 'Asia/Taipei';
const STATUSES = ['未開始', '進行中', '等待外部', '暫停', '待確認', '完成'];
const STATUS_ALIAS = {
  '已完成': '完成', '完工': '完成', '結案': '完成', 'done': '完成',
  '處理中': '進行中', '執行中': '進行中', '開始': '進行中',
  '等待': '等待外部', '待回覆': '等待外部', '追蹤中': '等待外部',
  '暫緩': '暫停', '暫滯': '暫停', '擱置': '暫停',
  '提報完成': '待確認', '送審': '待確認', '待審核': '待確認'
};
const TASK_COLS = [
  ['id', 'ID'], ['type', '類型'], ['source', '來源'], ['project', '專案名稱'], ['content', '代辦內容'],
  ['owner', '負責人'], ['assignee', '承辦人'], ['due', '需完成日期'],
  ['assignedAt', '指派時間'], ['startedAt', '開始時間'], ['submittedAt', '提報時間'], ['doneAt', '完成時間'],
  ['status', '狀態'], ['progress', '進度'], ['latest', '最新回報'], ['waitingFor', '等待對象'],
  ['link', '連結'], ['updatedAt', '最後更新']
];
const PEOPLE_COLS = ['姓名', '角色', 'Email', 'LINE名稱', '帳號', '密碼'];
const ADMIN_ROW = ['管理者', '主管', '', '', 'n821129x', 'n821129x'];
const SESSION_DAYS = 14;
const LOG_COLS = ['時間', 'ID', '回報人', '狀態', '進度', '內容', '來源', '訊息ID'];
const CONFIG_DEFAULTS = [
  ['稱謂', '報告正翰', '每日回報第一句的稱謂'],
  ['系統名稱', '每日工作追蹤', '網頁標題與寄件者名稱'],
  ['系統網址', '', '部署網頁應用程式後貼上，信件中的「開啟系統」按鈕會用到'],
  ['每日摘要時間', 8, '每天幾點寄摘要（0–23），修改後要重新執行「安裝自動排程」'],
  ['週末寄送摘要', '否', '是／否'],
  ['未開始提醒天數', 2, '指派超過幾天仍未開始，列入主管摘要提醒'],
  ['等待外部提醒天數', 7, '等待外部超過幾天沒更新，列入主管摘要提醒']
];

/* ───────────── 選單與初始化 ───────────── */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('代辦系統')
    .addItem('① 初始化工作表', 'setup')
    .addItem('② 安裝自動排程（每日摘要＋讀取回信）', 'installTriggers')
    .addSeparator()
    .addItem('立即寄送每日摘要（測試）', 'dailyDigestNow')
    .addItem('立即讀取回信（測試）', 'scanInbox')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);
  ensureSheet_(ss, SHEET.PEOPLE, PEOPLE_COLS, [ADMIN_ROW, ['範例員工', '員工', '', '', 'staff01', '1234']]);
  ensurePeopleCols_();
  ensureSheet_(ss, SHEET.TASKS, TASK_COLS.map(c => c[1])).getRange('A:A').setNumberFormat('@');
  ensureTaskCols_();
  ensureSheet_(ss, SHEET.LOG, LOG_COLS);
  ensureSheet_(ss, SHEET.CONFIG, ['項目', '值', '說明'], CONFIG_DEFAULTS);
  ['工作表1', 'Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(tr => {
    if (['dailyDigest', 'scanInbox'].indexOf(tr.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(tr);
  });
  let h = parseInt(getConfig_()['每日摘要時間'], 10);
  if (isNaN(h) || h < 0 || h > 23) h = 8;
  ScriptApp.newTrigger('dailyDigest').timeBased().everyDays(1).atHour(h).inTimezone(TZ).create();
  ScriptApp.newTrigger('scanInbox').timeBased().everyMinutes(10).create();
}

function ensureSheet_(ss, name, headers, defaults) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#ECF1FE').setFontColor('#1C1B2E');
    sh.setFrozenRows(1);
    if (defaults && defaults.length) sh.getRange(2, 1, defaults.length, defaults[0].length).setValues(defaults);
  }
  return sh;
}

/** 舊版工作表缺少的欄位補在最後面 */
function ensureTaskCols_() {
  const sh = sheet_(SHEET.TASKS);
  const lastCol = Math.max(1, sh.getLastColumn());
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const missing = TASK_COLS.map(c => c[1]).filter(h => head.indexOf(h) < 0);
  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing])
      .setFontWeight('bold').setBackground('#ECF1FE').setFontColor('#1C1B2E');
  }
}

/* ───────────── 網頁 ───────────── */

function doGet() {
  const cfg = getConfig_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(String(cfg['系統名稱'] || '每日工作追蹤'))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ───────────── 登入 ───────────── */

function api_login(account, password) {
  ensurePeopleCols_();
  const acc = String(account || '').trim().toLowerCase();
  if (!acc || !password) throw new Error('請輸入帳號與密碼');
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + acc;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= 5) throw new Error('密碼錯誤太多次，請 10 分鐘後再試');
  const p = getPeople_().filter(x => x.account && x.account.toLowerCase() === acc)[0];
  if (!p || p.password !== String(password)) {
    cache.put(failKey, String(fails + 1), 600);
    throw new Error('帳號或密碼錯誤');
  }
  cache.remove(failKey);
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const props = PropertiesService.getScriptProperties();
  cleanSessions_(props);
  props.setProperty('s_' + token, JSON.stringify({ n: p.name, e: Date.now() + SESSION_DAYS * 864e5 }));
  return { token: token, data: bootstrap_(p.name) };
}

function api_logout(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty('s_' + token);
  return true;
}

/** 由登入憑證取得使用者姓名；失效時丟出「請重新登入」 */
function auth_(token) {
  const fail = () => { throw new Error('登入已過期，請重新登入'); };
  if (!token) fail();
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('s_' + token);
  if (!raw) fail();
  const s = JSON.parse(raw);
  if (s.e < Date.now()) { props.deleteProperty('s_' + token); fail(); }
  if (!findPerson_(getPeople_(), s.n)) { props.deleteProperty('s_' + token); fail(); }
  return s.n;
}

function cleanSessions_(props) {
  const all = props.getProperties(), now = Date.now();
  Object.keys(all).forEach(k => {
    if (k.indexOf('s_') !== 0) return;
    try { if (JSON.parse(all[k]).e < now) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
  });
}

/** 舊版人員表補上帳號密碼欄，並確保有管理者帳號 */
function ensurePeopleCols_() {
  const sh = sheet_(SHEET.PEOPLE);
  const lastCol = Math.max(1, sh.getLastColumn());
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const missing = PEOPLE_COLS.filter(h => head.indexOf(h) < 0);
  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing])
      .setFontWeight('bold').setBackground('#ECF1FE').setFontColor('#1C1B2E');
  }
  if (!getPeople_().some(p => p.account.toLowerCase() === ADMIN_ROW[4])) {
    const h2 = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    sh.appendRow(h2.map(h => { const i = PEOPLE_COLS.indexOf(h); return i >= 0 ? ADMIN_ROW[i] : ''; }));
  }
}

/* ───────────── 資料 API（第一個參數都是登入憑證） ───────────── */

function api_bootstrap(token) { return bootstrap_(auth_(token)); }

/** 主管拿到全部工作；員工只拿到自己的 */
function bootstrap_(actor) {
  const cfg = getConfig_();
  const all = getPeople_();
  const people = all.map(p => ({ name: p.name, role: p.role, hasEmail: !!p.email }));
  const me = findPerson_(all, actor);
  let tasks = [], projects = [];
  if (me) {
    tasks = readTasks_();
    projects = tasks.map(t => t.project).filter((v, i, a) => v && a.indexOf(v) === i);
    if (me.role !== '主管') tasks = tasks.filter(t => t.assignee === me.name);
  }
  return {
    me: me ? me.name : '',
    people: people,
    projects: projects,
    tasks: tasks.map(toClient_),
    config: { salutation: String(cfg['稱謂'] || '報告'), title: String(cfg['系統名稱'] || '每日工作追蹤') },
    today: fmt_(new Date(), 'yyyy-MM-dd')
  };
}

/** 主管：指派給任何人。員工：只能新增給自己（來源＝自行新增） */
function api_createTask(token, p) {
  const actor = auth_(token);
  const people = getPeople_();
  const me = mustPerson_(people, actor);
  const mgr = me.role === '主管';
  if (!p || !String(p.content || '').trim() || !p.due) throw new Error('請填寫工作內容與需完成日期');
  const assignee = mgr ? p.assignee : actor;
  if (!assignee) throw new Error('請選擇承辦人');
  if (!findPerson_(people, assignee)) throw new Error('人員表找不到承辦人：' + assignee);
  const selfAdd = assignee === actor;
  let owner = actor;
  if (!mgr) {
    const managers = people.filter(x => x.role === '主管');
    owner = (p.owner && isManager_(people, p.owner)) ? p.owner : (managers[0] ? managers[0].name : '');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let t;
  try {
    const now = new Date();
    t = {
      id: nextId_(readTasks_()),
      type: p.type === '臨時代辦' ? '臨時代辦' : '專案執行',
      source: (mgr && !selfAdd) ? '主管交辦' : '自行新增',
      project: String(p.project || '').trim(),
      content: String(p.content).trim(),
      owner: owner, assignee: assignee,
      due: parseDate_(p.due),
      assignedAt: now, startedAt: '', submittedAt: '', doneAt: '',
      status: '未開始', progress: '', latest: '', waitingFor: '',
      link: String(p.link || '').trim(), updatedAt: now
    };
    writeTask_(t);
    log_(t.id, actor, '未開始', '', t.source === '主管交辦' ? '指派給 ' + assignee : '自行新增', '網頁');
  } finally {
    lock.releaseLock();
  }
  if (t.source === '主管交辦' && p.notify !== false) sendAssignMail_(t, people, '');
  return bootstrap_(actor);
}

function api_start(token, id) {
  const actor = auth_(token);
  const people = getPeople_();
  const t = mustFindTask_(id);
  checkCanReport_(t, actor, people);
  const now = new Date();
  if (!t.startedAt) t.startedAt = now;
  if (t.status === '未開始') t.status = '進行中';
  t.updatedAt = now;
  writeTask_(t);
  log_(t.id, actor, t.status, t.progress, '開始執行', '網頁');
  return bootstrap_(actor);
}

function api_report(token, id, data) {
  const actor = auth_(token);
  const people = getPeople_();
  const t = mustFindTask_(id);
  checkCanReport_(t, actor, people);
  applyReport_(t, actor, data || {}, '網頁', '', people);
  return bootstrap_(actor);
}

function api_reassign(token, id, newAssignee) {
  const actor = auth_(token);
  const people = getPeople_();
  if (!isManager_(people, actor)) throw new Error('只有主管可以轉派');
  if (!findPerson_(people, newAssignee)) throw new Error('人員表找不到：' + newAssignee);
  const t = mustFindTask_(id);
  const old = t.assignee;
  if (old === newAssignee) return bootstrap_(actor);
  t.assignee = newAssignee;
  t.updatedAt = new Date();
  writeTask_(t);
  log_(t.id, actor, t.status, t.progress, '轉派：' + old + ' → ' + newAssignee, '網頁');
  sendAssignMail_(t, people, '此工作由 ' + old + ' 轉派給您。');
  return bootstrap_(actor);
}

function api_delete(token, id) {
  const actor = auth_(token);
  const people = getPeople_();
  if (!isManager_(people, actor)) throw new Error('只有主管可以刪除工作');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const t = mustFindTask_(id);
    sheet_(SHEET.TASKS).deleteRow(t._row);
    log_(t.id, actor, '刪除', '', '刪除：' + lineCore_(t) + '（承辦 ' + t.assignee + '）', '網頁');
  } finally {
    lock.releaseLock();
  }
  return bootstrap_(actor);
}

function api_history(token, id) {
  const actor = auth_(token);
  const people = getPeople_();
  const t = mustFindTask_(id);
  checkCanReport_(t, actor, people);
  const vals = sheet_(SHEET.LOG).getDataRange().getValues();
  const head = vals[0];
  const c = n => head.indexOf(n);
  return vals.slice(1)
    .filter(r => String(r[c('ID')]) === String(id))
    .map(r => ({
      time: fmt_(r[c('時間')], 'yyyy-MM-dd HH:mm'),
      actor: String(r[c('回報人')]), status: String(r[c('狀態')]),
      progress: r[c('進度')] === '' ? '' : Number(r[c('進度')]),
      content: String(r[c('內容')]), source: String(r[c('來源')])
    }))
    .reverse();
}

/* ───────────── 回報邏輯 ───────────── */

function applyReport_(t, actor, data, source, msgId, people) {
  const now = new Date();
  const mgr = isManager_(people, actor);
  const prev = t.status;
  let status = normStatus_(data.status) || prev || '未開始';
  if (status === '完成' && !mgr) status = '待確認';           // 員工按完成 → 送主管確認
  if (prev === '完成' && !mgr) throw new Error('此工作已由主管確認完成，如需修改請洽主管');
  if (status === '未開始' && (data.content || isNum_(data.progress))) status = '進行中';
  if (!t.startedAt && status !== '未開始') t.startedAt = now;
  if (isNum_(data.progress)) t.progress = Math.max(0, Math.min(100, Math.round(Number(data.progress))));

  if (status === '待確認') {
    if (prev !== '待確認') t.submittedAt = now;
    t.doneAt = '';
  } else if (status === '完成') {
    if (prev !== '完成') t.doneAt = now;
    if (!t.submittedAt) t.submittedAt = now;
    if (isNum_(t.progress)) t.progress = 100;
  } else {
    t.submittedAt = '';
    t.doneAt = '';
  }
  if (data.content) t.latest = String(data.content).trim();
  t.waitingFor = status === '等待外部' ? String(data.waitingFor || t.waitingFor || '').trim() : '';
  t.status = status;
  t.updatedAt = now;
  writeTask_(t);
  log_(t.id, actor, status, t.progress, data.content || '', source, msgId);

  if (status === '待確認' && prev !== '待確認') sendSubmitMail_(t, people);
  if (status === '完成' && prev !== '完成' && actor !== t.assignee) sendResultMail_(t, people, true, data.content);
  if (prev === '待確認' && status !== '待確認' && status !== '完成' && mgr) sendResultMail_(t, people, false, data.content);
}

function checkCanReport_(t, actor, people) {
  if (t.assignee !== actor && !isManager_(people, actor)) throw new Error('這件工作的承辦人是 ' + t.assignee);
}

function normStatus_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  if (STATUSES.indexOf(s) >= 0) return s;
  if (STATUS_ALIAS[s]) return STATUS_ALIAS[s];
  for (let i = 0; i < STATUSES.length; i++) if (s.indexOf(STATUSES[i]) >= 0) return STATUSES[i];
  return '';
}

/* ───────────── Email ───────────── */

function subjectFor_(t) {
  const title = (t.project ? t.project + '｜' : '') + t.content;
  return '【代辦 #' + t.id + '】' + (title.length > 40 ? title.slice(0, 40) + '…' : title);
}

function managersFor_(t, people) {
  const owner = findPerson_(people, t.owner);
  if (owner && owner.role === '主管' && owner.email) return [owner];
  return people.filter(p => p.role === '主管' && p.email);
}

function sendAssignMail_(t, people, note) {
  const p = findPerson_(people, t.assignee);
  if (!p || !p.email) return;
  const cfg = getConfig_();
  const rows = [['類型', t.type], ['專案', t.project || '—'], ['內容', t.content],
    ['需完成日期', fmt_(t.due, 'yyyy/MM/dd')], ['負責人', t.owner]];
  if (t.link) rows.push(['連結', t.link]);
  const guide = ['狀態：進行中', '進度：50', '內容：說明目前做了什麼'];
  const plain = [
    t.assignee + ' 您好，', '',
    note || (t.owner + ' 指派了一項工作給您：'), '',
    rows.map(r => r[0] + '：' + r[1]).join('\n'), '',
    cfg['系統網址'] ? '開啟系統：' + cfg['系統網址'] + '\n' : '',
    '──── 回報方式 ────',
    '直接回覆這封信，在最上面照以下格式填寫（進度可省略）：', '',
    guide.join('\n'), '',
    '狀態可填：進行中、等待外部、暫停、完成（完成會送主管確認）',
    '等待外部時可加一行「等待：叡揚」'
  ].join('\n');
  const html = mailShell_(cfg, '新的工作 #' + t.id,
    '<p style="margin:0 0 14px">' + esc_(t.assignee) + ' 您好，' + esc_(note || (t.owner + ' 指派了一項工作給您：')) + '</p>' +
    tableHtml_(rows) + guideHtml_(guide));
  GmailApp.sendEmail(p.email, subjectFor_(t), plain, { htmlBody: html, name: String(cfg['系統名稱'] || '每日工作追蹤') });
}

/** 員工提報完成 → 通知主管確認 */
function sendSubmitMail_(t, people) {
  const to = managersFor_(t, people);
  if (!to.length) return;
  const cfg = getConfig_();
  const rows = [['承辦人', t.assignee], ['來源', t.source || '主管交辦'], ['專案', t.project || '—'], ['內容', t.content],
    ['需完成日期', fmt_(t.due, 'yyyy/MM/dd')],
    ['反應時間', t.startedAt ? durText_(t.assignedAt, t.startedAt) + '（指派 → 開始）' : '—'],
    ['執行時間', t.startedAt ? durText_(t.startedAt, t.submittedAt) + '（開始 → 提報）' : '—']];
  if (t.latest) rows.push(['完成說明', t.latest]);
  const plain = rows.map(r => r[0] + '：' + r[1]).join('\n') + '\n\n請到系統按「確認完成」或「退回」。' +
    (cfg['系統網址'] ? '\n' + cfg['系統網址'] : '');
  const html = mailShell_(cfg, '待確認完成 #' + t.id, tableHtml_(rows) +
    '<p style="margin:14px 0 0;color:#6C6B80">請到系統按「確認完成」或「退回」。</p>');
  to.forEach(m => GmailApp.sendEmail(m.email, '【待確認 #' + t.id + '】' + t.assignee + '｜' + (t.project ? t.project + '｜' : '') + t.content,
    plain, { htmlBody: html, name: String(cfg['系統名稱'] || '每日工作追蹤') }));
}

/** 主管確認完成／退回 → 通知承辦人 */
function sendResultMail_(t, people, ok, note) {
  const p = findPerson_(people, t.assignee);
  if (!p || !p.email) return;
  const cfg = getConfig_();
  const title = ok ? '工作已確認完成 #' + t.id : '工作退回 #' + t.id;
  const rows = [['專案', t.project || '—'], ['內容', t.content], ['結果', ok ? '已確認完成' : '退回，請繼續處理']];
  if (note) rows.push(['主管說明', note]);
  const plain = rows.map(r => r[0] + '：' + r[1]).join('\n');
  GmailApp.sendEmail(p.email, (ok ? '【已完成 #' : '【退回 #') + t.id + '】' + (t.project ? t.project + '｜' : '') + t.content,
    plain, { htmlBody: mailShell_(cfg, title, tableHtml_(rows)), name: String(cfg['系統名稱'] || '每日工作追蹤') });
}

/* ───────────── 每日摘要 ───────────── */

function dailyDigestNow() { dailyDigest_(true); }
function dailyDigest() { dailyDigest_(false); }

function dailyDigest_(force) {
  const cfg = getConfig_();
  const now = new Date();
  const dow = Number(fmt_(now, 'u'));
  if (!force && dow >= 6 && String(cfg['週末寄送摘要']) !== '是') return;
  const people = getPeople_();
  const tasks = readTasks_();
  const today = fmt_(now, 'yyyy-MM-dd');
  const tomorrow = fmt_(new Date(now.getTime() + 864e5), 'yyyy-MM-dd');
  const sal = String(cfg['稱謂'] || '報告');
  const name = String(cfg['系統名稱'] || '每日工作追蹤');
  const md = fmt_(now, 'M/d');
  const recentDone = t => t.status === '完成' && t.doneAt instanceof Date && now - t.doneAt < 864e5;

  // 員工：自己的清單
  people.forEach(p => {
    if (!p.email) return;
    const mine = tasks.filter(t => t.assignee === p.name && t.status !== '完成');
    if (!mine.length) return;
    const text = buildReportText_(mine, sal, today, true);
    const alerts = [];
    mine.forEach(t => {
      const due = dateStr_(t.due);
      if (isOverdue_(t, today)) alerts.push('⚠ 逾期：' + lineCore_(t) + '（期限 ' + due + '）');
      else if (t.status !== '待確認' && (due === tomorrow || due === today)) alerts.push('⏰ ' + (due === today ? '今天' : '明天') + '到期：' + lineCore_(t));
    });
    const plain = (alerts.length ? alerts.join('\n') + '\n\n' : '') + text;
    const html = mailShell_(cfg, md + ' 今日預計作業事項',
      (alerts.length ? '<div style="background:#FDEDF4;border-radius:4px;padding:10px 14px;margin-bottom:14px;color:#A2275C">' +
        alerts.map(esc_).join('<br>') + '</div>' : '') + preHtml_(text));
    GmailApp.sendEmail(p.email, '【每日作業】' + md + ' ' + p.name + ' 今日預計作業事項', plain, { htmlBody: html, name: name });
  });

  // 主管：全員清單＋提醒
  const managers = people.filter(p => p.role === '主管' && p.email);
  if (!managers.length) return;
  const notStartedDays = Number(cfg['未開始提醒天數']) || 2;
  const waitDays = Number(cfg['等待外部提醒天數']) || 7;
  const alerts = [];
  tasks.forEach(t => {
    if (t.status === '完成') return;
    if (t.status === '待確認') alerts.push('✅ 待確認｜' + t.assignee + '｜' + lineCore_(t));
    if (isOverdue_(t, today)) alerts.push('⚠ 逾期｜' + t.assignee + '｜' + lineCore_(t) + '（期限 ' + dateStr_(t.due) + '）');
    if (t.status === '未開始' && t.assignedAt instanceof Date && (now - t.assignedAt) / 864e5 >= notStartedDays)
      alerts.push('⏳ 指派 ' + Math.floor((now - t.assignedAt) / 864e5) + ' 天未開始｜' + t.assignee + '｜' + lineCore_(t));
    if (t.status === '等待外部' && t.updatedAt instanceof Date && (now - t.updatedAt) / 864e5 >= waitDays)
      alerts.push('🕓 等待外部 ' + Math.floor((now - t.updatedAt) / 864e5) + ' 天未更新｜' + t.assignee + '｜' + lineCore_(t));
  });
  const sections = [];
  people.forEach(p => {
    const list = tasks.filter(t => t.assignee === p.name && (t.status !== '完成' || recentDone(t)));
    if (list.length) sections.push({ name: p.name, text: buildReportText_(list, sal, today, false) });
  });
  if (!sections.length && !alerts.length) return;
  const plain = (alerts.length ? '【需要注意】\n' + alerts.join('\n') + '\n\n' : '') +
    sections.map(s => '■ ' + s.name + '\n' + s.text).join('\n\n');
  const html = mailShell_(cfg, md + ' 團隊作業摘要',
    (alerts.length ? '<div style="background:#FDEDF4;border-radius:4px;padding:10px 14px;margin-bottom:16px;color:#A2275C"><b>需要注意</b><br>' +
      alerts.map(esc_).join('<br>') + '</div>' : '') +
    sections.map(s => '<div style="font-weight:700;color:#2D5BE3;margin:18px 0 6px">■ ' + esc_(s.name) + '</div>' + preHtml_(s.text)).join(''));
  managers.forEach(m => GmailApp.sendEmail(m.email, '【團隊摘要】' + md + ' 作業進度', plain, { htmlBody: html, name: name }));
}

/** 產生「報告正翰，關於今日預計作業事項如下」格式 */
function buildReportText_(tasks, salutation, today, withHeader) {
  const list = orderTasks_(tasks);
  const lines = withHeader ? [salutation + '，關於今日預計作業事項如下：', ''] : [];
  list.forEach((t, i) => {
    lines.push((i + 1) + '.' + reportLine_(t, today));
    if (t.link) lines.push(String(t.link));
  });
  return lines.join('\n');
}

function lineCore_(t) { return (t.project ? t.project + '，' : '') + t.content; }

function reportLine_(t, today) {
  let s = lineCore_(t);
  if (isNum_(t.progress)) s += '(' + t.progress + '%)';
  if (t.latest) s += '；' + String(t.latest).replace(/\s*\n+\s*/g, '，');
  s = s.replace(/[。．.]+$/, '');
  const tags = [];
  if (t.status === '完成') tags.push('已完成');
  else if (t.status === '待確認') tags.push('已完成，待主管確認');
  else if (t.status === '等待外部') tags.push('等待外部' + (t.waitingFor ? '：' + t.waitingFor : ''));
  else if (t.status === '暫停') tags.push('暫停');
  if (isOverdue_(t, today)) tags.push('逾期');
  if (tags.length) s += '（' + tags.join('，') + '）';
  return s + '。';
}

function orderTasks_(tasks) {
  const first = {};
  tasks.forEach(t => {
    const k = t.project ? t.project : '\uffff';
    const tm = toTime_(t.assignedAt);
    if (first[k] === undefined || tm < first[k]) first[k] = tm;
  });
  return tasks.slice().sort((a, b) => {
    const ka = a.project ? a.project : '\uffff', kb = b.project ? b.project : '\uffff';
    if (ka !== kb) {
      if (ka === '\uffff') return 1;
      if (kb === '\uffff') return -1;
      return first[ka] - first[kb] || (ka < kb ? -1 : 1);
    }
    return toTime_(a.assignedAt) - toTime_(b.assignedAt);
  });
}

function isOverdue_(t, today) {
  const d = dateStr_(t.due);
  return t.status !== '完成' && t.status !== '待確認' && !!d && d < today;
}

/* ───────────── 讀取回信 ───────────── */

function scanInbox() {
  const me = String(Session.getEffectiveUser().getEmail()).toLowerCase();
  const people = getPeople_();
  const processed = processedIds_();
  const okLabel = label_('代辦系統/已匯入');
  const failLabel = label_('代辦系統/需人工確認');
  const threads = GmailApp.search('subject:代辦 newer_than:7d', 0, 100);
  threads.forEach(th => {
    th.getMessages().forEach(msg => {
      const mid = msg.getId();
      if (processed[mid]) return;
      const m = String(msg.getSubject()).match(/代辦\s*#\s*(T\d{3,})/i);
      if (!m) return;
      const from = extractEmail_(msg.getFrom());
      if (from === me) return;
      let res;
      try {
        res = handleReply_(m[1].toUpperCase(), from, msg.getPlainBody(), people, mid);
      } catch (e) {
        res = { ok: false, error: e.message };
      }
      processed[mid] = true;
      if (res.ok) {
        th.addLabel(okLabel);
      } else {
        th.addLabel(failLabel);
        markFailed_(mid);
        msg.reply('這封回報沒有成功匯入系統：' + res.error + '\n\n' +
          '請直接回覆，在最上面照格式填寫：\n狀態：進行中\n進度：50\n內容：說明\n\n狀態可填：進行中、等待外部、暫停、完成');
      }
    });
  });
}

function handleReply_(id, from, body, people, mid) {
  const t = findTask_(id);
  if (!t) return { ok: false, error: '找不到工作 #' + id };
  const sender = people.filter(p => p.email && p.email === from)[0];
  if (!sender) return { ok: false, error: '您的 Email（' + from + '）不在人員名單中' };
  if (sender.name !== t.assignee && sender.role !== '主管') return { ok: false, error: '這件工作的承辦人是 ' + t.assignee };
  const data = parseReply_(stripQuoted_(body));
  if (!data) return { ok: false, error: '讀不到「狀態／進度／內容」任何一行' };
  if (data.statusRaw && !data.status) return { ok: false, error: '看不懂狀態「' + data.statusRaw + '」' };
  applyReport_(t, sender.name, data, 'Email', mid, people);
  return { ok: true };
}

/** 只保留回覆內容，去掉引用的原信 */
function stripQuoted_(body) {
  const lines = String(body || '').replace(/\r/g, '').split('\n');
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i].trim();
    if (/^>/.test(l)) break;
    if (/wrote:\s*$/i.test(l) || /寫道[:：]?\s*$/.test(l)) {
      if (out.length && /^(On |在|於)/.test(out[out.length - 1].trim())) out.pop();
      break;
    }
    if (/^-{2,}\s*(Original Message|原始郵件)/i.test(l) || /^(From|寄件者)\s*[:：]/i.test(l) || /^─{4,}/.test(l)) break;
    out.push(lines[i]);
  }
  return out.join('\n');
}

function parseReply_(text) {
  const out = { status: '', statusRaw: '', progress: '', content: '', waitingFor: '' };
  let cur = null, found = false, ended = false;
  String(text).split('\n').forEach(raw => {
    if (ended) return;
    const line = raw.trim();
    if (line === '--' || line === '-- ') { ended = true; return; }
    const m = line.match(/^(狀態|進度|內容|等待對象|等待)\s*[:：]\s*(.*)$/);
    if (m) {
      found = true;
      const k = m[1], v = m[2].trim();
      if (k === '狀態') { out.statusRaw = v; out.status = normStatus_(v); cur = null; }
      else if (k === '進度') {
        const n = parseInt(v.replace(/[^\d]/g, ''), 10);
        out.progress = isNaN(n) ? '' : Math.max(0, Math.min(100, n));
        cur = null;
      }
      else if (k === '內容') { out.content = v; cur = '內容'; }
      else { out.waitingFor = v; cur = null; }
      return;
    }
    if (cur === '內容' && line) out.content += (out.content ? '\n' : '') + line;
  });
  return found ? out : null;
}

function processedIds_() {
  const map = {};
  const vals = sheet_(SHEET.LOG).getDataRange().getValues();
  const col = vals[0].indexOf('訊息ID');
  if (col >= 0) vals.slice(1).forEach(r => { if (r[col]) map[String(r[col])] = true; });
  failedIds_().forEach(id => { map[id] = true; });
  return map;
}
function failedIds_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('failedMsgIds') || '[]'); }
  catch (e) { return []; }
}
function markFailed_(id) {
  const a = failedIds_();
  a.push(id);
  PropertiesService.getScriptProperties().setProperty('failedMsgIds', JSON.stringify(a.slice(-200)));
}
function label_(name) { return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name); }
function extractEmail_(s) {
  const m = String(s).match(/<([^>]+)>/);
  return (m ? m[1] : String(s)).trim().toLowerCase();
}

/* ───────────── 試算表存取 ───────────── */

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('找不到「' + name + '」分頁，請先執行「代辦系統 → ① 初始化工作表」');
  return sh;
}

function getConfig_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.CONFIG);
  const out = {};
  if (!sh) return out;
  sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) out[String(r[0]).trim()] = r[1]; });
  return out;
}

function getPeople_() {
  const vals = sheet_(SHEET.PEOPLE).getDataRange().getValues();
  const head = vals[0].map(h => String(h).trim());
  const col = n => head.indexOf(n);
  const get = (r, n) => col(n) >= 0 ? String(r[col(n)] == null ? '' : r[col(n)]).trim() : '';
  return vals.slice(1)
    .filter(r => get(r, '姓名'))
    .map(r => ({
      name: get(r, '姓名'),
      role: get(r, '角色') === '主管' ? '主管' : '員工',
      email: get(r, 'Email').toLowerCase(),
      line: get(r, 'LINE名稱'),
      account: get(r, '帳號'),
      password: get(r, '密碼')
    }));
}
function findPerson_(people, name) { return people.filter(p => p.name === name)[0]; }
function mustPerson_(people, name) {
  const p = findPerson_(people, name);
  if (!p) throw new Error('人員表找不到：' + name + '，請重新選擇身分');
  return p;
}
function isManager_(people, name) { const p = findPerson_(people, name); return !!p && p.role === '主管'; }

function readTasks_() {
  const vals = sheet_(SHEET.TASKS).getDataRange().getValues();
  if (vals.length < 2) return [];
  const head = vals[0];
  const idx = TASK_COLS.map(c => head.indexOf(c[1]));
  return vals.slice(1).map((r, i) => {
    const t = { _row: i + 2 };
    TASK_COLS.forEach((c, k) => { t[c[0]] = idx[k] >= 0 ? r[idx[k]] : ''; });
    t.id = String(t.id).trim();
    ['type', 'source', 'project', 'content', 'owner', 'assignee', 'status', 'latest', 'waitingFor', 'link']
      .forEach(k => { t[k] = String(t[k] == null ? '' : t[k]).trim(); });
    t.status = normStatus_(t.status) || '未開始';
    t.source = t.source || '主管交辦';
    t.progress = isNum_(t.progress) ? Number(t.progress) : '';
    return t;
  }).filter(t => t.id);
}

function writeTask_(t) {
  ensureTaskCols_();
  const sh = sheet_(SHEET.TASKS);
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const existing = t._row ? sh.getRange(t._row, 1, 1, head.length).getValues()[0] : head.map(() => '');
  const row = head.map((h, i) => {
    const c = TASK_COLS.filter(c => c[1] === h)[0];
    return c ? (t[c[0]] == null ? '' : t[c[0]]) : existing[i];
  });
  if (t._row) sh.getRange(t._row, 1, 1, row.length).setValues([row]);
  else { sh.appendRow(row); t._row = sh.getLastRow(); }
}

function findTask_(id) { return readTasks_().filter(t => t.id === String(id))[0]; }
function mustFindTask_(id) {
  const t = findTask_(id);
  if (!t) throw new Error('找不到工作 #' + id + '（可能已被刪除）');
  return t;
}

/** 編號也參考回報歷程，刪除後不會重複使用 */
function nextId_(tasks) {
  let max = 0;
  const bump = v => { const n = parseInt(String(v).replace(/\D/g, ''), 10); if (n > max) max = n; };
  tasks.forEach(t => bump(t.id));
  const log = sheet_(SHEET.LOG);
  if (log.getLastRow() > 1) log.getRange(2, 2, log.getLastRow() - 1, 1).getValues().forEach(r => bump(r[0]));
  return 'T' + ('000' + (max + 1)).slice(-4);
}

function log_(id, actor, status, progress, content, source, msgId) {
  sheet_(SHEET.LOG).appendRow([new Date(), id, actor, status, progress === '' ? '' : progress, content || '', source, msgId || '']);
}

function toClient_(t) {
  return {
    id: t.id, type: t.type || '專案執行', source: t.source || '主管交辦', project: t.project, content: t.content,
    owner: t.owner, assignee: t.assignee, due: dateStr_(t.due),
    assignedAt: fmt_(t.assignedAt, 'yyyy-MM-dd HH:mm'), startedAt: fmt_(t.startedAt, 'yyyy-MM-dd HH:mm'),
    submittedAt: fmt_(t.submittedAt, 'yyyy-MM-dd HH:mm'),
    doneAt: fmt_(t.doneAt, 'yyyy-MM-dd HH:mm'), updatedAt: fmt_(t.updatedAt, 'yyyy-MM-dd HH:mm'),
    status: t.status, progress: t.progress, latest: t.latest, waitingFor: t.waitingFor, link: t.link
  };
}

/* ───────────── 小工具 ───────────── */

function fmt_(v, pattern) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, TZ, pattern);
  return v == null ? '' : String(v).replace(/\//g, '-');
}
function dateStr_(v) { return fmt_(v, 'yyyy-MM-dd').slice(0, 10); }
function parseDate_(s) { return new Date(String(s).slice(0, 10) + 'T00:00:00+08:00'); }
function toTime_(v) {
  if (v instanceof Date) return v.getTime();
  const n = Date.parse(String(v || '').replace(' ', 'T'));
  return isNaN(n) ? 0 : n;
}
function isNum_(v) { return v !== '' && v !== null && v !== undefined && !isNaN(Number(v)); }
function durText_(a, b) {
  const ms = Math.max(0, toTime_(b) - toTime_(a));
  const m = Math.floor(ms / 6e4);
  if (m < 60) return m + ' 分鐘';
  const h = Math.floor(m / 60);
  if (h < 24) return h + ' 小時';
  return Math.floor(h / 24) + ' 天 ' + (h % 24) + ' 小時';
}
function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function mailShell_(cfg, title, inner) {
  const url = String(cfg['系統網址'] || '');
  return '<div style="font-family:\'Noto Sans TC\',\'Microsoft JhengHei\',Arial,sans-serif;background:#F1F1F3;padding:24px">' +
    '<div style="max-width:660px;margin:0 auto;background:#fff;border:1px solid #E1E1E8;border-radius:6px;overflow:hidden">' +
    '<div style="height:5px;background:linear-gradient(90deg,#2D5BE3,#8A55D6,#E0508A,#E0791F)"></div>' +
    '<div style="background:#2D5BE3;color:#fff;padding:16px 24px;font-size:17px;font-weight:700">' + esc_(title) + '</div>' +
    '<div style="padding:20px 24px;color:#1C1B2E;font-size:15px;line-height:1.75">' + inner + '</div>' +
    (url ? '<div style="padding:0 24px 22px"><a href="' + esc_(url) + '" style="display:inline-block;background:#2D5BE3;color:#fff;text-decoration:none;padding:9px 18px;border-radius:4px;font-weight:600">開啟系統</a></div>' : '') +
    '</div></div>';
}
function tableHtml_(rows) {
  return '<table style="border-collapse:collapse;width:100%;font-size:14.5px">' + rows.map(r =>
    '<tr><td style="padding:7px 10px;background:#ECF1FE;color:#2D5BE3;font-weight:600;width:110px;border-bottom:1px solid #fff;vertical-align:top">' + esc_(r[0]) +
    '</td><td style="padding:7px 12px;border-bottom:1px solid #EEE">' + esc_(r[1]) + '</td></tr>').join('') + '</table>';
}
function guideHtml_(lines) {
  return '<div style="margin-top:18px;background:#F6EFFD;border:1px solid #E6D2F8;border-radius:6px;padding:14px 16px">' +
    '<div style="font-weight:700;color:#3B1E63;margin-bottom:6px">直接回覆這封信即可回報</div>' +
    '<div style="font-size:13.5px;color:#6C6B80;margin-bottom:8px">在回信最上面照以下格式填寫，進度可省略：</div>' +
    '<pre style="margin:0;font-family:inherit;background:#fff;border-radius:4px;padding:10px 12px;font-size:14px">' + lines.map(esc_).join('\n') + '</pre>' +
    '<div style="font-size:13px;color:#6C6B80;margin-top:8px">狀態可填：進行中、等待外部、暫停、完成（完成會送主管確認）；等待外部時可加一行「等待：對象」</div></div>';
}
function preHtml_(text) {
  return '<pre style="white-space:pre-wrap;margin:0;font-family:inherit;font-size:14.5px;line-height:1.8">' + esc_(text) + '</pre>';
}
