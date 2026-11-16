/**
 * 小朋友存錢筒 — Google Apps Script 後端
 * --------------------------------------------------------------
 * 資料庫：本試算表（只有擁有者能開啟，不需分享給任何人）
 * 前端：GitHub Pages 靜態網頁，透過 fetch 呼叫本程式的 /exec 網址
 * 身分驗證：前端用「Sign in with Google」取得 ID token，後端向 Google 驗證
 * 小孩端：用 Settings 裡的 kid_token（唯讀專屬連結），不需登入
 *
 * 部署方式：部署 → 新增部署作業 → 類型「網頁應用程式」
 *   執行身分：我（擁有者）
 *   誰可以存取：所有人
 */

// ===== 常數 =====
const SHEETS = {
  USERS:    { name: 'Users',    headers: ['email', 'name', 'role', 'active'] },
  LEDGER:   { name: 'Ledger',   headers: ['id', 'time', 'type', 'amount', 'source', 'actor_email', 'actor_name', 'note', 'ref_id', 'link_id'] },
  TASKS:    { name: 'Tasks',    headers: ['id', 'name', 'amount', 'stamp_target', 'stamp_bonus', 'active'] },
  REWARDS:  { name: 'Rewards',  headers: ['id', 'name', 'price', 'active'] },
  GOALS:    { name: 'Goals',    headers: ['id', 'name', 'target', 'status', 'created', 'done_time'] },
  SETTINGS: { name: 'Settings', headers: ['key', 'value'] },
};

// 流水帳類型
const TYPE_LABEL = {
  deposit:  '給錢',
  deduct:   '扣錢',
  task:     '任務',
  bonus:    '集章獎勵',
  spend:    '花費',
  redeem:   '兌換',
  interest: '利息',
  reversal: '撤銷',
};
const INCOME_TYPES = ['deposit', 'task', 'bonus', 'interest'];

// 權限：giver = 長輩（爺爺奶奶等），admin = 爸媽
const PERMS = {
  deposit: ['admin', 'giver'],
  deduct:  ['admin', 'giver'],
  task:    ['admin', 'giver'],
  spend:   ['admin'],
  redeem:  ['admin'],
  manage:  ['admin'],
};

const DEFAULT_SETTINGS = {
  kid_name: '寶貝',
  interest_step: 1000,      // 每滿多少元
  interest_per_step: 10,    // 每月多少利息
  last_interest_month: '',  // yyyy-MM，防止重複發息
};

// =====================================================================
// 一、初始化（第一次手動執行一次）
// =====================================================================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.values(SHEETS).forEach(def => {
    let sh = ss.getSheetByName(def.name);
    if (!sh) sh = ss.insertSheet(def.name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  // 刪掉預設的「工作表1 / Sheet1」
  ['工作表1', 'Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  // 預設設定
  const settings = readSettings_();
  Object.keys(DEFAULT_SETTINGS).forEach(k => {
    if (settings[k] === undefined) writeSetting_(k, DEFAULT_SETTINGS[k]);
  });

  // 小孩唯讀連結的 token
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('KID_TOKEN')) props.setProperty('KID_TOKEN', newToken_());
  if (!props.getProperty('CLIENT_ID')) props.setProperty('CLIENT_ID', 'PASTE_YOUR_CLIENT_ID.apps.googleusercontent.com');

  // 把執行者（你自己）設為管理員
  const me = Session.getEffectiveUser().getEmail().toLowerCase();
  const users = readRows_(SHEETS.USERS);
  if (me && !users.some(u => String(u.email).toLowerCase() === me)) {
    appendRow_(SHEETS.USERS, { email: me, name: '爸爸', role: 'admin', active: true });
  }

  // 範例任務 / 獎品（可在網頁或試算表修改）
  if (readRows_(SHEETS.TASKS).length === 0) {
    appendRow_(SHEETS.TASKS, { id: newId_('T'), name: '洗碗', amount: 10, stamp_target: 10, stamp_bonus: 50, active: true });
    appendRow_(SHEETS.TASKS, { id: newId_('T'), name: '整理書桌', amount: 20, stamp_target: 5, stamp_bonus: 30, active: true });
    appendRow_(SHEETS.TASKS, { id: newId_('T'), name: '準時睡覺', amount: 5, stamp_target: 20, stamp_bonus: 100, active: true });
  }

  installInterestTrigger_();
  Logger.log('✅ 初始化完成');
  Logger.log('小孩唯讀 token：' + props.getProperty('KID_TOKEN'));
  Logger.log('下一步：專案設定 → 指令碼屬性 → 把 CLIENT_ID 改成你的 OAuth 用戶端 ID');
}

/** 每月 1 日凌晨 1 點發利息 */
function installInterestTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'monthlyInterest')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('monthlyInterest').timeBased().onMonthDay(1).atHour(1).create();
}

/** 重新產生小孩連結（舊連結會失效） */
function resetKidToken() {
  const t = newToken_();
  PropertiesService.getScriptProperties().setProperty('KID_TOKEN', t);
  Logger.log('新的小孩 token：' + t);
  return t;
}

// =====================================================================
// 二、HTTP 入口
// =====================================================================
function doGet() {
  return json_({ ok: true, service: 'kids-savings', time: new Date().toISOString() });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    return json_(route_(body));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err), code: err && err.code || 'ERROR' });
  }
}

function route_(body) {
  const action = body.action;

  // --- 小孩端：只用 token，唯讀 ---
  if (action === 'kid') {
    const token = PropertiesService.getScriptProperties().getProperty('KID_TOKEN');
    if (!body.token || !token || body.token !== token) throw codeErr_('連結無效，請跟爸媽要新的連結', 'BAD_TOKEN');
    return { ok: true, data: kidView_() };
  }

  // --- 長輩 / 爸媽：需要 Google 登入 ---
  const user = authUser_(body.idToken);

  switch (action) {
    case 'state':
      return { ok: true, data: fullState_(user) };

    case 'deposit':
    case 'deduct':
    case 'spend':
      return withLock_(() => { addMoney_(user, action, body); return { ok: true, data: fullState_(user) }; });

    case 'task':
      return withLock_(() => { completeTask_(user, body); return { ok: true, data: fullState_(user) }; });

    case 'redeem':
      return withLock_(() => { redeem_(user, body); return { ok: true, data: fullState_(user) }; });

    case 'reverse':
      return withLock_(() => { reverse_(user, body.id); return { ok: true, data: fullState_(user) }; });

    case 'saveTask':
    case 'saveReward':
    case 'saveGoal':
    case 'saveUser':
    case 'saveSettings':
    case 'goalAchieve':
    case 'resetKidToken':
      requirePerm_(user, 'manage');
      return withLock_(() => { manage_(user, action, body); return { ok: true, data: fullState_(user) }; });

    default:
      throw codeErr_('未知的操作：' + action, 'BAD_ACTION');
  }
}

// =====================================================================
// 三、身分驗證
// =====================================================================
function authUser_(idToken) {
  if (!idToken) throw codeErr_('請先登入', 'AUTH');
  const info = verifyIdToken_(idToken);
  const user = readRows_(SHEETS.USERS).find(u =>
    String(u.email).trim().toLowerCase() === info.email && isTrue_(u.active));
  if (!user) throw codeErr_('此帳號（' + info.email + '）尚未被加入，請爸媽在「設定 → 成員」新增', 'NOT_MEMBER');
  return { email: info.email, name: String(user.name), role: String(user.role) };
}

/** 向 Google 驗證 ID token，結果快取到 token 到期前 */
function verifyIdToken_(idToken) {
  const cache = CacheService.getScriptCache();
  const key = 'tok_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken));
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const res = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw codeErr_('登入已過期，請重新登入', 'AUTH');

  const p = JSON.parse(res.getContentText());
  const clientId = PropertiesService.getScriptProperties().getProperty('CLIENT_ID');
  if (p.aud !== clientId) throw codeErr_('登入憑證不屬於本系統', 'AUTH');
  if (['accounts.google.com', 'https://accounts.google.com'].indexOf(p.iss) < 0) throw codeErr_('登入憑證來源錯誤', 'AUTH');
  if (String(p.email_verified) !== 'true') throw codeErr_('Google 帳號 email 未驗證', 'AUTH');
  const expMs = Number(p.exp) * 1000;
  if (!(expMs > Date.now())) throw codeErr_('登入已過期，請重新登入', 'AUTH');

  const info = { email: String(p.email).toLowerCase(), name: p.name || '' };
  const ttl = Math.min(21600, Math.floor((expMs - Date.now()) / 1000) - 30);
  if (ttl > 0) cache.put(key, JSON.stringify(info), ttl);
  return info;
}

function requirePerm_(user, perm) {
  if (PERMS[perm].indexOf(user.role) < 0) throw codeErr_('你沒有權限執行此操作', 'FORBIDDEN');
}

// =====================================================================
// 四、寫入操作
// =====================================================================
function addMoney_(user, type, body) {
  requirePerm_(user, type);
  const amt = parseAmount_(body.amount);
  const note = cleanText_(body.note);
  if (type === 'spend' && !note) throw codeErr_('請填寫花費項目', 'BAD_INPUT');
  const sign = type === 'deposit' ? 1 : -1;
  if (sign < 0) ensureEnough_(amt);
  appendLedger_(user, { type: type, amount: sign * amt, note: note });
}

function completeTask_(user, body) {
  requirePerm_(user, 'task');
  const task = readRows_(SHEETS.TASKS).find(t => t.id === body.taskId && isTrue_(t.active));
  if (!task) throw codeErr_('找不到這個任務', 'BAD_INPUT');
  const entry = appendLedger_(user, {
    type: 'task', amount: Number(task.amount) || 0, note: task.name + cleanSuffix_(body.note), ref_id: task.id,
  });

  // 集章：完成次數達到 stamp_target 的倍數時自動加碼
  const target = Number(task.stamp_target) || 0;
  const bonus = Number(task.stamp_bonus) || 0;
  if (target > 0 && bonus > 0) {
    const count = effectiveLedger_(readRows_(SHEETS.LEDGER))
      .filter(r => r.type === 'task' && r.ref_id === task.id).length;
    if (count % target === 0) {
      appendLedger_(user, {
        type: 'bonus', amount: bonus,
        note: task.name + ' 集滿 ' + target + ' 章！', ref_id: task.id, link_id: entry.id,
      });
    }
  }
}

function redeem_(user, body) {
  requirePerm_(user, 'redeem');
  const r = readRows_(SHEETS.REWARDS).find(x => x.id === body.rewardId && isTrue_(x.active));
  if (!r) throw codeErr_('找不到這個獎品', 'BAD_INPUT');
  const price = parseAmount_(r.price);
  ensureEnough_(price);
  appendLedger_(user, { type: 'redeem', amount: -price, note: '兌換：' + r.name, ref_id: r.id });
}

/** 撤銷：新增一筆反向紀錄，不刪除原紀錄 */
function reverse_(user, id) {
  const all = readRows_(SHEETS.LEDGER);
  const orig = all.find(r => r.id === id);
  if (!orig) throw codeErr_('找不到這筆紀錄', 'BAD_INPUT');
  if (orig.type === 'reversal') throw codeErr_('撤銷紀錄不能再撤銷', 'BAD_INPUT');
  if (orig.type === 'interest' || orig.type === 'bonus') throw codeErr_('利息與集章獎勵由系統產生，不能單獨撤銷', 'BAD_INPUT');
  if (all.some(r => r.type === 'reversal' && r.link_id === id)) throw codeErr_('這筆已經撤銷過了', 'BAD_INPUT');
  const own = String(orig.actor_email).toLowerCase() === user.email;
  if (!own && user.role !== 'admin') throw codeErr_('只能撤銷自己記的紀錄', 'FORBIDDEN');

  appendLedger_(user, {
    type: 'reversal', amount: -Number(orig.amount), source: orig.source,
    note: '撤銷：' + (TYPE_LABEL[orig.type] || orig.type) + ' ' + (orig.note || ''), link_id: orig.id,
  });

  // 若這筆任務觸發過集章獎勵，獎勵一併撤銷
  if (orig.type === 'task') {
    all.filter(r => r.type === 'bonus' && r.link_id === orig.id)
      .filter(b => !all.some(r => r.type === 'reversal' && r.link_id === b.id))
      .forEach(b => appendLedger_(user, {
        type: 'reversal', amount: -Number(b.amount), source: b.source,
        note: '撤銷：集章獎勵 ' + (b.note || ''), link_id: b.id,
      }));
  }
}

function manage_(user, action, b) {
  switch (action) {
    case 'saveTask': {
      const name = cleanText_(b.name);
      if (!name) throw codeErr_('請填任務名稱', 'BAD_INPUT');
      upsert_(SHEETS.TASKS, {
        id: b.id || newId_('T'), name: name,
        amount: parseAmount_(b.amount, true),
        stamp_target: Math.max(0, Math.floor(Number(b.stamp_target) || 0)),
        stamp_bonus: parseAmount_(b.stamp_bonus, true),
        active: b.active !== false,
      });
      break;
    }
    case 'saveReward': {
      const name = cleanText_(b.name);
      if (!name) throw codeErr_('請填獎品名稱', 'BAD_INPUT');
      upsert_(SHEETS.REWARDS, { id: b.id || newId_('R'), name: name, price: parseAmount_(b.price), active: b.active !== false });
      break;
    }
    case 'saveGoal': {
      const name = cleanText_(b.name);
      if (!name) throw codeErr_('請填目標名稱', 'BAD_INPUT');
      const existing = b.id ? readRows_(SHEETS.GOALS).find(g => g.id === b.id) : null;
      upsert_(SHEETS.GOALS, {
        id: b.id || newId_('G'), name: name, target: parseAmount_(b.target),
        status: b.status || (existing ? existing.status : 'active'),
        created: existing ? existing.created : new Date(), done_time: existing ? existing.done_time : '',
      });
      break;
    }
    case 'goalAchieve': {
      // 達成目標：扣掉目標金額（當作用存款買下），並標記完成
      const g = readRows_(SHEETS.GOALS).find(x => x.id === b.id && x.status === 'active');
      if (!g) throw codeErr_('找不到進行中的目標', 'BAD_INPUT');
      const price = parseAmount_(g.target);
      ensureEnough_(price);
      appendLedger_(user, { type: 'redeem', amount: -price, note: '達成目標：' + g.name, ref_id: g.id });
      g.status = 'done'; g.done_time = new Date();
      upsert_(SHEETS.GOALS, g);
      break;
    }
    case 'saveUser': {
      const email = String(b.email || '').trim().toLowerCase();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw codeErr_('email 格式不正確', 'BAD_INPUT');
      if (['admin', 'giver'].indexOf(b.role) < 0) throw codeErr_('角色只能是 admin 或 giver', 'BAD_INPUT');
      const name = cleanText_(b.name);
      if (!name) throw codeErr_('請填顯示名稱（例如：爺爺）', 'BAD_INPUT');
      if (email === user.email && (b.role !== 'admin' || b.active === false)) throw codeErr_('不能移除自己的管理員權限', 'BAD_INPUT');
      upsertBy_(SHEETS.USERS, 'email', { email: email, name: name, role: b.role, active: b.active !== false });
      break;
    }
    case 'saveSettings': {
      if (b.kid_name !== undefined) writeSetting_('kid_name', cleanText_(b.kid_name) || '寶貝');
      if (b.interest_step !== undefined) writeSetting_('interest_step', parseAmount_(b.interest_step));
      if (b.interest_per_step !== undefined) writeSetting_('interest_per_step', parseAmount_(b.interest_per_step, true));
      break;
    }
    case 'resetKidToken':
      resetKidToken();
      break;
  }
}

// =====================================================================
// 五、月息（時間觸發器呼叫）
// =====================================================================
function monthlyInterest() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    const month = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
    const s = readSettings_();
    if (String(s.last_interest_month) === month) return; // 本月已發過
    const amt = calcInterest_(balanceOf_(readRows_(SHEETS.LEDGER)), s);
    if (amt > 0) {
      appendLedger_({ email: 'system', name: '系統' }, {
        type: 'interest', amount: amt, source: '利息', note: month + ' 月息',
      });
    }
    writeSetting_('last_interest_month', month);
  } finally {
    lock.releaseLock();
  }
}

function calcInterest_(balance, s) {
  const step = Number(s.interest_step) || 0;
  const per = Number(s.interest_per_step) || 0;
  if (step <= 0 || per <= 0 || balance <= 0) return 0;
  return Math.floor(balance / step) * per;
}

// =====================================================================
// 六、讀取 / 彙整
// =====================================================================
function fullState_(user) {
  const ledger = readRows_(SHEETS.LEDGER);
  const eff = effectiveLedger_(ledger);
  const tasks = readRows_(SHEETS.TASKS);
  const settings = readSettings_();
  const reversedIds = {};
  ledger.forEach(r => { if (r.type === 'reversal') reversedIds[r.link_id] = true; });

  const isAdmin = user.role === 'admin';
  const recent = ledger.slice(-300).reverse().map(r => ({
    id: r.id, time: iso_(r.time), type: r.type, typeLabel: TYPE_LABEL[r.type] || r.type,
    amount: Number(r.amount), source: r.source, actor: r.actor_name, note: r.note,
    reversed: !!reversedIds[r.id],
    canReverse: r.type !== 'reversal' && r.type !== 'interest' && r.type !== 'bonus' && !reversedIds[r.id] &&
      (isAdmin || String(r.actor_email).toLowerCase() === user.email),
  }));

  const state = {
    me: user,
    kidName: settings.kid_name || '寶貝',
    balance: balanceOf_(ledger),
    bySource: bySource_(eff),
    stamps: stamps_(eff, tasks),
    tasks: tasks.filter(t => isTrue_(t.active)).map(cleanTask_),
    rewards: readRows_(SHEETS.REWARDS).filter(r => isTrue_(r.active)).map(r => ({ id: r.id, name: r.name, price: Number(r.price) })),
    goals: goals_(ledger),
    nextInterest: calcInterest_(balanceOf_(ledger), settings),
    history: monthlyHistory_(ledger),
    ledger: recent,
    perms: Object.keys(PERMS).filter(p => PERMS[p].indexOf(user.role) >= 0),
  };
  if (isAdmin) {
    state.admin = {
      users: readRows_(SHEETS.USERS).map(u => ({ email: u.email, name: u.name, role: u.role, active: isTrue_(u.active) })),
      allTasks: tasks.map(cleanTask_),
      allRewards: readRows_(SHEETS.REWARDS).map(r => ({ id: r.id, name: r.name, price: Number(r.price), active: isTrue_(r.active) })),
      settings: { kid_name: settings.kid_name, interest_step: Number(settings.interest_step), interest_per_step: Number(settings.interest_per_step) },
      kidToken: PropertiesService.getScriptProperties().getProperty('KID_TOKEN'),
    };
  }
  return state;
}

/** 小孩端只看到：總存款、目標進度、集章進度 */
function kidView_() {
  const ledger = readRows_(SHEETS.LEDGER);
  const eff = effectiveLedger_(ledger);
  const settings = readSettings_();
  return {
    kidName: settings.kid_name || '寶貝',
    balance: balanceOf_(ledger),
    goals: goals_(ledger).filter(g => g.status === 'active').map(g => ({ name: g.name, target: g.target, progress: g.progress })),
    doneGoals: goals_(ledger).filter(g => g.status === 'done').length,
    history: monthlyHistory_(ledger),
    stamps: stamps_(eff, readRows_(SHEETS.TASKS)).map(s => ({
      name: s.name, inCard: s.inCard, target: s.target, cards: s.cards, total: s.total,
    })),
    updated: new Date().toISOString(),
  };
}

/** 排除「已被撤銷」與「撤銷紀錄本身」後的有效流水帳 */
function effectiveLedger_(ledger) {
  const reversed = {};
  ledger.forEach(r => { if (r.type === 'reversal') reversed[r.link_id] = true; });
  return ledger.filter(r => r.type !== 'reversal' && !reversed[r.id]);
}

/** 每月月底累積存款（含撤銷，與總存款一致）：從第一筆紀錄的月份到本月 */
function monthlyHistory_(ledger) {
  const tz = 'Asia/Taipei';
  const byMonth = {};
  ledger.forEach(r => {
    const d = r.time instanceof Date ? r.time : new Date(r.time);
    if (isNaN(d)) return;
    const m = Utilities.formatDate(d, tz, 'yyyy-MM');
    byMonth[m] = round_((byMonth[m] || 0) + (Number(r.amount) || 0));
  });
  const months = Object.keys(byMonth).sort();
  if (!months.length) return [];
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  const last = months[months.length - 1] > now ? months[months.length - 1] : now;
  let y = Number(months[0].slice(0, 4)), mo = Number(months[0].slice(5, 7));
  const out = [];
  let cum = 0;
  while (out.length < 600) {
    const key = y + '-' + (mo < 10 ? '0' : '') + mo;
    const change = byMonth[key] || 0;
    cum = round_(cum + change);
    out.push({ month: key, balance: cum, change: change });
    if (key >= last) break;
    mo++; if (mo > 12) { mo = 1; y++; }
  }
  return out;
}

function balanceOf_(ledger) {
  return round_(ledger.reduce((s, r) => s + (Number(r.amount) || 0), 0));
}

/** 收入依來源分組（爺爺、奶奶、利息…），支出合計 */
function bySource_(eff) {
  const map = {};
  let out = 0;
  eff.forEach(r => {
    const a = Number(r.amount) || 0;
    if (INCOME_TYPES.indexOf(r.type) >= 0) {
      const k = r.source || '未知';
      map[k] = round_((map[k] || 0) + a);
    } else {
      out = round_(out + a);
    }
  });
  const income = Object.keys(map).map(k => ({ source: k, amount: map[k] })).sort((a, b) => b.amount - a.amount);
  return { income: income, outgoing: out };
}

function stamps_(eff, tasks) {
  return tasks.filter(t => isTrue_(t.active) && Number(t.stamp_target) > 0).map(t => {
    const total = eff.filter(r => r.type === 'task' && r.ref_id === t.id).length;
    const target = Number(t.stamp_target);
    return { id: t.id, name: t.name, total: total, target: target, cards: Math.floor(total / target), inCard: total % target, bonus: Number(t.stamp_bonus) || 0 };
  });
}

function goals_(ledger) {
  const bal = balanceOf_(ledger);
  return readRows_(SHEETS.GOALS).map(g => {
    const target = Number(g.target) || 0;
    return {
      id: g.id, name: g.name, target: target, status: g.status || 'active',
      progress: g.status === 'done' ? 1 : (target > 0 ? Math.max(0, Math.min(1, bal / target)) : 0),
      remain: Math.max(0, round_(target - bal)), done_time: iso_(g.done_time),
    };
  });
}

function cleanTask_(t) {
  return { id: t.id, name: t.name, amount: Number(t.amount) || 0, stamp_target: Number(t.stamp_target) || 0, stamp_bonus: Number(t.stamp_bonus) || 0, active: isTrue_(t.active) };
}

function ensureEnough_(amt) {
  const bal = balanceOf_(readRows_(SHEETS.LEDGER));
  if (bal - amt < 0) throw codeErr_('存款不足（目前 ' + bal + ' 元）', 'INSUFFICIENT');
}

// =====================================================================
// 七、試算表工具
// =====================================================================
function sheet_(def) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(def.name);
  if (!sh) throw codeErr_('找不到工作表 ' + def.name + '，請先執行 setup()', 'SETUP');
  return sh;
}

function readRows_(def) {
  const sh = sheet_(def);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, def.headers.length).getValues();
  return vals.filter(row => row.some(v => v !== '' && v !== null)).map(row => {
    const o = {};
    def.headers.forEach((h, i) => { o[h] = row[i]; });
    return o;
  });
}

function appendRow_(def, obj) {
  sheet_(def).appendRow(def.headers.map(h => (obj[h] === undefined ? '' : obj[h])));
}

function appendLedger_(user, e) {
  const entry = {
    id: newId_('L'), time: new Date(), type: e.type, amount: round_(e.amount),
    source: e.source || user.name, actor_email: user.email, actor_name: user.name,
    note: e.note || '', ref_id: e.ref_id || '', link_id: e.link_id || '',
  };
  appendRow_(SHEETS.LEDGER, entry);
  return entry;
}

function upsert_(def, obj) { upsertBy_(def, 'id', obj); }

function upsertBy_(def, keyField, obj) {
  const sh = sheet_(def);
  const last = sh.getLastRow();
  const col = def.headers.indexOf(keyField) + 1;
  if (last >= 2) {
    const keys = sh.getRange(2, col, last - 1, 1).getValues().map(r => String(r[0]).toLowerCase());
    const idx = keys.indexOf(String(obj[keyField]).toLowerCase());
    if (idx >= 0) {
      const row = def.headers.map(h => (obj[h] === undefined ? '' : obj[h]));
      sh.getRange(idx + 2, 1, 1, def.headers.length).setValues([row]);
      return;
    }
  }
  appendRow_(def, obj);
}

function readSettings_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS.name);
  const o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0] !== '') o[r[0]] = r[1]; });
  return o;
}

function writeSetting_(key, value) {
  upsertBy_(SHEETS.SETTINGS, 'key', { key: key, value: value });
}

// =====================================================================
// 八、小工具
// =====================================================================
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw codeErr_('系統忙碌中，請稍後再試', 'BUSY');
  try { return fn(); } finally { lock.releaseLock(); }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function codeErr_(msg, code) { const e = new Error(msg); e.code = code; return e; }

function parseAmount_(v, allowZero) {
  const n = Number(v);
  if (!isFinite(n) || n < 0 || (!allowZero && n === 0)) throw codeErr_('金額需為大於 0 的數字', 'BAD_INPUT');
  if (n > 1000000) throw codeErr_('金額過大', 'BAD_INPUT');
  return round_(n);
}

function cleanText_(s) { return String(s || '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, 100); }
function cleanSuffix_(s) { const t = cleanText_(s); return t ? '（' + t + '）' : ''; }
function round_(n) { return Math.round(Number(n) * 100) / 100; }
function isTrue_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function iso_(d) { return d instanceof Date ? d.toISOString() : (d ? String(d) : ''); }
function newId_(p) { return p + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyMMddHHmmss') + Math.random().toString(36).slice(2, 6); }
function newToken_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8); }
