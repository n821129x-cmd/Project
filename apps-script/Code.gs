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

const SHEET = { PEOPLE: '人員', TASKS: '代辦事項', LOG: '回報歷程', CONFIG: '設定', REPORTS: '日報' };
const REPORT_COLS = ['日期', '姓名', '寄出時間', 'LINE文字', 'AI摘要', '工作數'];
const DEFAULT_MODEL = 'gemini-3.5-flash-lite';
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
  ['等待外部提醒天數', 7, '等待外部超過幾天沒更新，列入主管摘要提醒'],
  ['Gemini模型', DEFAULT_MODEL, 'AI 彙整使用的模型；API 金鑰請放在 Apps Script「專案設定 → 指令碼屬性」的 GEMINI_API_KEY'],
  ['日報收件人', '', '留空＝寄給該員工的負責主管（或所有主管）；也可填 Email，多個用逗號分隔']
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

/** 取得資料試算表：綁定的試算表 → 已記錄的試算表 → 自動建立一份 */
function ss_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* 被刪除就重建 */ } }
  ss = SpreadsheetApp.create('每日工作追蹤資料');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function setup() {
  const ss = ss_();
  ss.setSpreadsheetTimeZone(TZ);
  ensureSheet_(ss, SHEET.PEOPLE, PEOPLE_COLS, [ADMIN_ROW, ['範例員工', '員工', '', '', 'staff01', '1234']]);
  ensurePeopleCols_();
  ensureSheet_(ss, SHEET.TASKS, TASK_COLS.map(c => c[1])).getRange('A:A').setNumberFormat('@');
  ensureTaskCols_();
  ensureSheet_(ss, SHEET.LOG, LOG_COLS);
  ensureSheet_(ss, SHEET.CONFIG, ['項目', '值', '說明'], CONFIG_DEFAULTS);
  ensureSheet_(ss, SHEET.REPORTS, REPORT_COLS);
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
    isSuper: !!(me && me.super),
    aiEnabled: !!geminiKey_(),
    reportedToday: me ? reportedOn_(me.name, fmt_(new Date(), 'yyyy-MM-dd')) : false,
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

/* ───────────── 日報（員工下班寄出） ───────────── */

/** AI 彙整今日進度（員工寄日報前預覽用） */
function api_aiDaily(token) {
  const actor = auth_(token);
  return aiDailyText_(actor);
}

/** 寄出日報：LINE 文字 ＋ AI 摘要 ＋ 工作表格 */
function api_sendDailyReport(token, lineText, aiText) {
  const actor = auth_(token);
  const people = getPeople_();
  const me = mustPerson_(people, actor);
  const cfg = getConfig_();
  const today = fmt_(new Date(), 'yyyy-MM-dd');
  const tasks = orderTasks_(readTasks_().filter(t => t.assignee === actor &&
    (t.status !== '完成' || (t.doneAt instanceof Date && fmt_(t.doneAt, 'yyyy-MM-dd') === today))));
  const text = String(lineText || '').trim() || buildReportText_(tasks, String(cfg['稱謂'] || '報告'), today, true);
  let ai = String(aiText || '').trim();
  if (!ai && geminiKey_()) { try { ai = aiDailyText_(actor); } catch (e) { ai = ''; } }

  // 收件人
  let to = String(cfg['日報收件人'] || '').split(/[,，;；\s]+/).filter(x => /@/.test(x));
  if (!to.length) {
    const owners = tasks.map(t => t.owner).filter((v, i, a) => v && a.indexOf(v) === i)
      .map(n => findPerson_(people, n)).filter(p => p && p.role === '主管' && p.email && p.name !== actor);
    to = (owners.length ? owners : people.filter(p => p.role === '主管' && p.email && p.name !== actor)).map(p => p.email);
  }
  if (!to.length) throw new Error('找不到收件主管的 Email，請主管到「人員管理」填寫 Email');
  to = to.filter((v, i, a) => a.indexOf(v) === i);

  const md = fmt_(new Date(), 'M/d');
  const html = mailShell_(cfg, md + ' ' + actor + ' 工作日報',
    sectionTitle_('1', 'LINE 回報文字') + preHtml_(text) +
    (ai ? sectionTitle_('2', 'AI 進度摘要') + '<div style="background:#F6EFFD;border:1px solid #E6D2F8;border-radius:6px;padding:12px 16px;font-size:14.5px;line-height:1.8">' + mdToHtml_(ai) + '</div>' : '') +
    sectionTitle_(ai ? '3' : '2', '工作明細（' + tasks.length + ' 件）') + taskTableHtml_(tasks, today));
  const plain = '【LINE 回報文字】\n' + text + (ai ? '\n\n【AI 進度摘要】\n' + ai : '') + '\n\n【工作明細】\n' +
    tasks.map(t => '#' + t.id + ' ' + lineCore_(t) + '｜' + t.status + (isNum_(t.progress) ? ' ' + t.progress + '%' : '') + '｜期限 ' + dateStr_(t.due)).join('\n');
  const opt = { htmlBody: html, name: actor + '（' + String(cfg['系統名稱'] || '每日工作追蹤') + '）' };
  if (me.email) { opt.replyTo = me.email; opt.cc = me.email; }
  GmailApp.sendEmail(to.join(','), '【日報】' + md + ' ' + actor + ' 工作回報', plain, opt);

  sheet_(SHEET.REPORTS).appendRow([today, actor, new Date(), text, ai, tasks.length]);
  return { sentTo: to, ai: ai, data: bootstrap_(actor) };
}

function reportedOn_(name, day) {
  const sh = ss_().getSheetByName(SHEET.REPORTS);
  if (!sh || sh.getLastRow() < 2) return false;
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().some(r => dateStr_(r[0]) === day && String(r[1]).trim() === name);
}

function prevWorkday_(d) {
  const x = new Date(d.getTime());
  do { x.setTime(x.getTime() - 864e5); } while ([6, 7].indexOf(Number(fmt_(x, 'u'))) >= 0);
  return fmt_(x, 'yyyy-MM-dd');
}

function sectionTitle_(n, t) {
  return '<div style="display:flex;align-items:center;gap:8px;margin:22px 0 10px;font-weight:700;font-size:15.5px;color:#1C1B2E">' +
    '<span style="display:inline-block;width:22px;height:22px;line-height:22px;text-align:center;border-radius:50%;background:#2D5BE3;color:#fff;font-size:12.5px">' + n + '</span>' + esc_(t) + '</div>';
}

function taskTableHtml_(tasks, today) {
  const color = { '未開始': ['#F2F2F5', '#6C6B80'], '進行中': ['#ECF1FE', '#2D5BE3'], '等待外部': ['#F6EFFD', '#8A55D6'],
    '暫停': ['#FDF1E5', '#B85E10'], '待確認': ['#FFF5D6', '#8A5E00'], '完成': ['#E3F5F1', '#0E7466'] };
  const th = 'padding:8px 10px;background:#ECF1FE;color:#2D5BE3;font-weight:600;text-align:left;font-size:13px;white-space:nowrap';
  const td = 'padding:8px 10px;border-bottom:1px solid #EEE;font-size:13.5px;vertical-align:top';
  if (!tasks.length) return '<p style="color:#6C6B80">目前沒有工作</p>';
  return '<table style="border-collapse:collapse;width:100%">' +
    '<tr><th style="' + th + '">編號</th><th style="' + th + '">專案</th><th style="' + th + '">工作內容</th><th style="' + th + '">狀態</th><th style="' + th + '">進度</th><th style="' + th + '">最新回報</th><th style="' + th + '">期限</th></tr>' +
    tasks.map(t => {
      const c = color[t.status] || color['未開始'];
      const od = isOverdue_(t, today);
      return '<tr><td style="' + td + ';color:#9A99AB">#' + t.id + '</td><td style="' + td + '">' + esc_(t.project || '臨時代辦') + '</td>' +
        '<td style="' + td + ';font-weight:600">' + esc_(t.content) + '</td>' +
        '<td style="' + td + '"><span style="background:' + c[0] + ';color:' + c[1] + ';border-radius:999px;padding:2px 9px;font-size:12.5px;font-weight:600;white-space:nowrap">' + t.status + '</span>' +
        (od ? ' <span style="background:#FDEDF4;color:#C23A72;border-radius:999px;padding:2px 9px;font-size:12.5px;font-weight:600">逾期</span>' : '') + '</td>' +
        '<td style="' + td + '">' + (isNum_(t.progress) ? t.progress + '%' : '—') + '</td>' +
        '<td style="' + td + ';color:#4A4960">' + esc_(t.latest || '') + (t.status === '等待外部' && t.waitingFor ? '（等待：' + esc_(t.waitingFor) + '）' : '') + '</td>' +
        '<td style="' + td + ';white-space:nowrap' + (od ? ';color:#C23A72;font-weight:600' : '') + '">' + dateStr_(t.due).slice(5).replace('-', '/') + '</td></tr>';
    }).join('') + '</table>';
}

/** 簡易 Markdown → HTML（標題、條列、粗體） */
function mdToHtml_(text) {
  const out = [];
  let inList = false;
  String(text).split('\n').forEach(raw => {
    const line = raw.trim();
    const li = line.match(/^[-*•]\s+(.*)$/) || line.match(/^\d+[.)、]\s+(.*)$/);
    if (li) { if (!inList) { out.push('<ul style="margin:4px 0 8px;padding-left:20px">'); inList = true; } out.push('<li>' + inline_(li[1]) + '</li>'); return; }
    if (inList) { out.push('</ul>'); inList = false; }
    if (!line) return;
    const h = line.match(/^#{1,4}\s+(.*)$/);
    if (h) out.push('<div style="font-weight:700;color:#3B1E63;margin:10px 0 4px">' + inline_(h[1]) + '</div>');
    else out.push('<div style="margin:3px 0">' + inline_(line) + '</div>');
  });
  if (inList) out.push('</ul>');
  return out.join('');
  function inline_(s) { return esc_(s).replace(/\*\*(.+?)\*\*/g, '<b>$1</b>'); }
}

/* ───────────── AI（Gemini） ───────────── */

function geminiKey_() {
  const k = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (k) return k.trim();
  const c = String(getConfig_()['Gemini API金鑰'] || '').trim();
  return c;
}

function gemini_(prompt) {
  const key = geminiKey_();
  if (!key) throw new Error('尚未設定 Gemini API 金鑰：請到 Apps Script「專案設定 → 指令碼屬性」新增 GEMINI_API_KEY');
  const model = String(getConfig_()['Gemini模型'] || DEFAULT_MODEL).trim();
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';
  const body = {
    systemInstruction: { parts: [{ text: '你是工程顧問公司的專案助理，使用繁體中文（台灣用語）撰寫，語氣精簡、具體、客觀，不要捏造資料中沒有的內容。輸出用簡單的 Markdown（## 小標、- 條列、**粗體**），不要使用表格。' }] },
    contents: [{ role: 'user', parts: [{ text: prompt }] }]
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body)
  });
  const code = res.getResponseCode();
  let json = {};
  try { json = JSON.parse(res.getContentText()); } catch (e) { /* 非 JSON */ }
  if (code !== 200) {
    const msg = (json.error && json.error.message) || res.getContentText().slice(0, 200);
    if (code === 429) throw new Error('AI 使用量已達上限，請稍後再試');
    if (code === 404) throw new Error('找不到 AI 模型「' + model + '」，請到「設定」分頁修改 Gemini模型');
    throw new Error('AI 回應錯誤（' + code + '）：' + msg);
  }
  const parts = (((json.candidates || [])[0] || {}).content || {}).parts || [];
  const text = parts.filter(p => !p.thought).map(p => p.text || '').join('').trim();
  if (!text) throw new Error('AI 沒有回傳內容，請再試一次');
  return text;
}

function taskBrief_(t, today) {
  return '#' + t.id + '｜' + (t.type === '臨時代辦' ? '臨時交辦' : '專案') + '｜' + lineCore_(t) + '｜狀態：' + t.status +
    (isNum_(t.progress) ? '｜進度：' + t.progress + '%' : '') + (t.latest ? '｜最新回報：' + t.latest : '') +
    (t.waitingFor ? '｜等待：' + t.waitingFor : '') + '｜期限：' + dateStr_(t.due) + (isOverdue_(t, today) ? '（已逾期）' : '');
}

function logsBetween_(from, to, name) {
  const vals = sheet_(SHEET.LOG).getDataRange().getValues().slice(1);
  return vals.filter(r => {
    const d = dateStr_(r[0]);
    return d >= from && d <= to && (!name || String(r[2]).trim() === name);
  }).map(r => ({ day: dateStr_(r[0]), time: fmt_(r[0], 'MM/dd HH:mm'), id: String(r[1]), actor: String(r[2]), status: String(r[3]), progress: r[4], content: String(r[5]), source: String(r[6]) }));
}

function aiDailyText_(name) {
  const today = fmt_(new Date(), 'yyyy-MM-dd');
  const tasks = readTasks_().filter(t => t.assignee === name &&
    (t.status !== '完成' || (t.doneAt instanceof Date && fmt_(t.doneAt, 'yyyy-MM-dd') === today)));
  const logs = logsBetween_(today, today, name).filter(l => l.status !== '刪除');
  const prompt = '請根據以下資料，替「' + name + '」寫今天的工作進度摘要，給主管快速閱讀。\n' +
    '要求：3～6 行條列；先寫今天完成或有進展的事，再寫卡關／等待外部的事與需要主管協助之處，最後一行寫明日重點。不要逐項照抄。\n\n' +
    '【今天的回報紀錄】\n' + (logs.length ? logs.map(l => l.time + ' #' + l.id + ' ' + l.status + (l.progress !== '' ? ' ' + l.progress + '%' : '') + ' ' + l.content).join('\n') : '（今天沒有在系統更新）') +
    '\n\n【手上的工作】\n' + (tasks.length ? tasks.map(t => taskBrief_(t, today)).join('\n') : '（無）');
  return gemini_(prompt);
}

/** 團隊總覽：單一員工的 AI 簡要說明 */
function api_aiPerson(token, name) {
  mustManager_(token);
  const today = fmt_(new Date(), 'yyyy-MM-dd');
  const from = fmt_(new Date(Date.now() - 6 * 864e5), 'yyyy-MM-dd');
  const tasks = readTasks_().filter(t => t.assignee === name && (t.status !== '完成' || (t.doneAt instanceof Date && dateStr_(t.doneAt) >= from)));
  const logs = logsBetween_(from, today, name).filter(l => l.status !== '刪除');
  const prompt = '請用 2～4 句話，說明「' + name + '」目前的工作狀況，給主管參考：手上件數與重點工作、近 7 天的進展、逾期或卡關、是否需要協助。不要用條列，不要寫標題。\n\n' +
    '【工作】\n' + (tasks.length ? tasks.map(t => taskBrief_(t, today)).join('\n') : '（無）') +
    '\n\n【近 7 天回報】\n' + (logs.length ? logs.map(l => l.time + ' #' + l.id + ' ' + l.status + ' ' + l.content).join('\n') : '（沒有回報紀錄）');
  return gemini_(prompt);
}

/** 最高管理者：指定區間的員工辦理狀況分析 */
function api_aiRange(token, from, to, who) {
  const actor = auth_(token);
  const me = findPerson_(getPeople_(), actor);
  if (!me || !me.super) throw new Error('只有最高管理者可以使用 AI 區間分析');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to) || from > to) throw new Error('請選擇正確的日期區間');
  const today = fmt_(new Date(), 'yyyy-MM-dd');
  const people = getPeople_().filter(p => !who || p.name === who);
  const all = readTasks_();
  const inRange = d => d && d >= from && d <= to;
  const blocks = people.map(p => {
    const mine = all.filter(t => t.assignee === p.name);
    const assigned = mine.filter(t => inRange(dateStr_(t.assignedAt)));
    const done = mine.filter(t => t.status === '完成' && inRange(dateStr_(t.doneAt)));
    const open = mine.filter(t => t.status !== '完成');
    const logs = logsBetween_(from, to, p.name).filter(l => l.status !== '刪除');
    if (!assigned.length && !done.length && !open.length && !logs.length) return '';
    return '### ' + p.name + '（' + p.role + '）\n' +
      '區間內新指派 ' + assigned.length + ' 件、完成 ' + done.length + ' 件、目前未完成 ' + open.length + ' 件（其中逾期 ' + open.filter(t => isOverdue_(t, today)).length + ' 件）、回報 ' + logs.length + ' 次\n' +
      '完成：\n' + (done.map(t => '- ' + lineCore_(t)).join('\n') || '- 無') + '\n' +
      '未完成：\n' + (open.map(t => '- ' + taskBrief_(t, today)).join('\n') || '- 無') + '\n' +
      '回報紀錄（最多 40 筆）：\n' + (logs.slice(-40).map(l => '- ' + l.time + ' #' + l.id + ' ' + l.status + ' ' + l.content).join('\n') || '- 無');
  }).filter(Boolean);
  if (!blocks.length) return '這段期間沒有任何工作或回報紀錄。';
  const prompt = '請分析 ' + from + ' 至 ' + to + ' 期間' + (who ? '「' + who + '」' : '團隊各員工') + '的工作辦理狀況，給公司最高管理者閱讀。\n' +
    '請依序輸出：\n## 整體概況（2～3 句，含完成件數與重點）\n## 各員工辦理狀況（每人 2～3 點：完成什麼、進行中重點、效率與回報情形）\n## 風險與卡關（逾期、等待外部過久、久未回報）\n## 建議（3 點以內，具體可執行）\n\n' +
    blocks.join('\n\n');
  return gemini_(prompt);
}

/* ───────────── 人員管理（主管） ───────────── */

function mustManager_(token) {
  const actor = auth_(token);
  if (!isManager_(getPeople_(), actor)) throw new Error('只有主管可以使用此功能');
  return actor;
}

function peopleForAdmin_() {
  return getPeople_().map(p => ({ name: p.name, role: p.role, email: p.email, line: p.line, account: p.account, hasPassword: !!p.password }));
}

function api_people(token) {
  mustManager_(token);
  return peopleForAdmin_();
}

/** 新增或修改一位人員；original 為修改前姓名（新增時留空）；密碼留空表示不變 */
function api_savePerson(token, p, original) {
  const actor = mustManager_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    upsertPerson_(p, original || '', actor);
  } finally {
    lock.releaseLock();
  }
  return { people: peopleForAdmin_(), data: bootstrap_(actor) };
}

function api_deletePerson(token, name) {
  const actor = mustManager_(token);
  if (name === actor) throw new Error('不能刪除自己的帳號');
  const people = getPeople_();
  const target = findPerson_(people, name);
  if (!target) throw new Error('找不到人員：' + name);
  if (target.role === '主管' && people.filter(p => p.role === '主管').length <= 1) throw new Error('至少要保留一位主管');
  const open = readTasks_().filter(t => t.assignee === name && t.status !== '完成').length;
  if (open) throw new Error(name + ' 還有 ' + open + ' 件未完成的工作，請先轉派');
  const sh = sheet_(SHEET.PEOPLE);
  const vals = sh.getDataRange().getValues();
  const col = vals[0].map(String).indexOf('姓名');
  for (let i = vals.length - 1; i >= 1; i--) if (String(vals[i][col]).trim() === name) sh.deleteRow(i + 1);
  return { people: peopleForAdmin_(), data: bootstrap_(actor) };
}

/** Excel 批次匯入人員：依姓名比對，已存在就更新，不存在就新增 */
function api_importPeople(token, rows) {
  const actor = mustManager_(token);
  const result = { added: 0, updated: 0, errors: [] };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    (rows || []).forEach((r, i) => {
      try {
        const exists = !!findPerson_(getPeople_(), String(r.name || '').trim());
        upsertPerson_(r, exists ? String(r.name).trim() : '', actor, true);
        if (exists) result.updated++; else result.added++;
      } catch (e) {
        result.errors.push('第 ' + (r._row || i + 2) + ' 列：' + e.message);
      }
    });
  } finally {
    lock.releaseLock();
  }
  result.people = peopleForAdmin_();
  result.data = bootstrap_(actor);
  return result;
}

function upsertPerson_(p, original, actor, fromImport) {
  const name = String(p.name || '').trim();
  const role = /主管|管理/.test(String(p.role || '')) ? '主管' : '員工';
  const email = String(p.email || '').trim().toLowerCase();
  const account = String(p.account || '').trim();
  const password = p.password == null ? '' : String(p.password);
  if (!name) throw new Error('姓名不可空白');
  if (email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Email 格式不正確：' + email);
  const people = getPeople_();
  const others = people.filter(x => x.name !== original);
  if (others.some(x => x.name === name)) throw new Error('姓名「' + name + '」已存在');
  if (account && others.some(x => x.account && x.account.toLowerCase() === account.toLowerCase())) throw new Error('帳號「' + account + '」已被使用');
  const old = original ? findPerson_(people, original) : null;
  if (original && !old) throw new Error('找不到人員：' + original);
  if (old && original === actor && role !== '主管') throw new Error('不能把自己改成員工');
  if (account && !password && !(old && old.password)) throw new Error('有帳號就要設定密碼');
  if (password && password.length < 4) throw new Error('密碼至少 4 個字');

  const sh = sheet_(SHEET.PEOPLE);
  ensurePeopleCols_();
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const rec = {
    '姓名': name, '角色': (old && old.roleRaw === '管理者' && role === '主管') ? '管理者' : role, 'Email': email,
    'LINE名稱': p.line == null && old ? old.line : String(p.line || '').trim(),
    '帳號': fromImport && !account && old ? old.account : account,
    '密碼': password || (old ? old.password : '')
  };
  if (old) {
    const vals = sh.getDataRange().getValues();
    const col = head.indexOf('姓名');
    let row = -1;
    for (let i = 1; i < vals.length; i++) if (String(vals[i][col]).trim() === original) { row = i + 1; break; }
    const cur = vals[row - 1];
    sh.getRange(row, 1, 1, head.length).setValues([head.map((h, i) => (h in rec) ? rec[h] : cur[i])]);
    if (name !== original) renameInTasks_(original, name);
  } else {
    sh.appendRow(head.map(h => (h in rec) ? rec[h] : ''));
  }
}

/** 改名時一併更新工作的承辦人／負責人 */
function renameInTasks_(from, to) {
  const sh = sheet_(SHEET.TASKS);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return;
  const head = vals[0].map(String);
  [head.indexOf('承辦人'), head.indexOf('負責人')].forEach(c => {
    if (c < 0) return;
    const colVals = vals.slice(1).map(r => [String(r[c]).trim() === from ? to : r[c]]);
    sh.getRange(2, c + 1, colVals.length, 1).setValues(colVals);
  });
}

/* ───────────── Excel 批次匯入工作（主管） ───────────── */

function api_importTasks(token, rows, notify) {
  const actor = mustManager_(token);
  const people = getPeople_();
  const result = { added: 0, errors: [] };
  const valid = [];
  (rows || []).forEach((r, i) => {
    const rowNo = r._row || i + 2;
    const content = String(r.content || '').trim();
    const assignee = String(r.assignee || '').trim();
    const due = String(r.due || '').trim();
    const errs = [];
    if (!content) errs.push('工作內容空白');
    if (!assignee) errs.push('承辦人空白');
    else if (!findPerson_(people, assignee)) errs.push('人員表沒有「' + assignee + '」');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(due)) errs.push('需完成日期格式不正確');
    const st = r.status ? normStatus_(r.status) : '未開始';
    if (r.status && !st) errs.push('看不懂狀態「' + r.status + '」');
    if (errs.length) { result.errors.push('第 ' + rowNo + ' 列：' + errs.join('、')); return; }
    valid.push({ r: r, content: content, assignee: assignee, due: due, status: st === '待確認' ? '進行中' : st });
  });
  if (!valid.length) { result.data = bootstrap_(actor); return result; }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const created = [];
  try {
    ensureTaskCols_();
    const sh = sheet_(SHEET.TASKS);
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    let next = parseInt(nextId_(readTasks_()).slice(1), 10);
    const now = new Date();
    const taskRows = [], logRows = [];
    valid.forEach(v => {
      const id = 'T' + ('000' + next++).slice(-4);
      const prog = isNum_(v.r.progress) ? Math.max(0, Math.min(100, Math.round(Number(v.r.progress)))) : '';
      const started = v.status !== '未開始';
      const t = {
        id: id, type: /臨時/.test(String(v.r.type || '')) ? '臨時代辦' : '專案執行', source: '主管交辦',
        project: String(v.r.project || '').trim(), content: v.content, owner: actor, assignee: v.assignee,
        due: parseDate_(v.due), assignedAt: now, startedAt: started ? now : '',
        submittedAt: v.status === '完成' ? now : '', doneAt: v.status === '完成' ? now : '',
        status: v.status, progress: v.status === '完成' && prog !== '' ? 100 : prog,
        latest: String(v.r.latest || '').trim(), waitingFor: v.status === '等待外部' ? String(v.r.waitingFor || '').trim() : '',
        link: String(v.r.link || '').trim(), updatedAt: now
      };
      taskRows.push(head.map(h => { const c = TASK_COLS.filter(c => c[1] === h)[0]; return c ? (t[c[0]] == null ? '' : t[c[0]]) : ''; }));
      logRows.push([now, id, actor, t.status, t.progress, 'Excel 匯入，指派給 ' + t.assignee, 'Excel', '']);
      created.push(t);
    });
    sh.getRange(sh.getLastRow() + 1, 1, taskRows.length, head.length).setValues(taskRows);
    const log = sheet_(SHEET.LOG);
    log.getRange(log.getLastRow() + 1, 1, logRows.length, logRows[0].length).setValues(logRows);
    result.added = created.length;
  } finally {
    lock.releaseLock();
  }
  if (notify) created.filter(t => t.status !== '完成').forEach(t => { try { sendAssignMail_(t, people, ''); } catch (e) { /* 額度用完就略過 */ } });
  result.data = bootstrap_(actor);
  return result;
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
  const lastDay = prevWorkday_(now);
  const missing = people.filter(p => p.role !== '主管' && tasks.some(t => t.assignee === p.name && t.status !== '完成') && !reportedOn_(p.name, lastDay));
  if (missing.length) alerts.unshift('📭 ' + lastDay.slice(5).replace('-', '/') + ' 未寄日報：' + missing.map(p => p.name).join('、'));
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
  let sh = ss_().getSheetByName(name);
  if (!sh) {                       // 第一次使用自動建立所有分頁
    setup();
    sh = ss_().getSheetByName(name);
  }
  if (!sh) throw new Error('找不到「' + name + '」分頁');
  return sh;
}

function getConfig_() {
  const sh = ss_().getSheetByName(SHEET.CONFIG);
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
      role: /^(主管|管理者)$/.test(get(r, '角色')) ? '主管' : '員工',
      roleRaw: get(r, '角色'),
      super: get(r, '角色') === '管理者' || get(r, '帳號').toLowerCase() === ADMIN_ROW[4],
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
