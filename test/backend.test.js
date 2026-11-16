// Mock Apps Script services and exercise Code.gs end to end
// 執行：node test/backend.test.js apps-script/Code.gs
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

function makeSheet(name) {
  const rows = [];
  const sh = {
    name, rows,
    getLastRow: () => rows.length,
    getRange: (r, c, nr = 1, nc = 1) => ({
      getValues: () => { const out = []; for (let i = 0; i < nr; i++) { const row = rows[r - 1 + i] || []; out.push(Array.from({ length: nc }, (_, j) => row[c - 1 + j] ?? '')); } return out; },
      setValues: v => { v.forEach((row, i) => { rows[r - 1 + i] = rows[r - 1 + i] || []; row.forEach((x, j) => rows[r - 1 + i][c - 1 + j] = x); }); return { setFontWeight() {} }; },
    }),
    appendRow: v => rows.push(v.slice()),
    setFrozenRows() {},
  };
  return sh;
}
const sheets = { '工作表1': makeSheet('工作表1') };
const ss = {
  getSheetByName: n => sheets[n] || null,
  insertSheet: n => (sheets[n] = makeSheet(n)),
  getSheets: () => Object.values(sheets),
  deleteSheet: s => delete sheets[s.name],
};
const props = {};
const cache = {};
let tokenInfo = {};
let uid = 0;
const ctx = {
  console,
  SpreadsheetApp: { getActiveSpreadsheet: () => ss },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => (props[k] = v) }) },
  CacheService: { getScriptCache: () => ({ get: k => cache[k] ?? null, put: (k, v) => (cache[k] = v) }) },
  Session: { getEffectiveUser: () => ({ getEmail: () => 'Dad@gmail.com' }), getScriptTimeZone: () => 'Asia/Taipei' },
  LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock() {}, releaseLock() {} }) },
  ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => ({ timeBased: () => ({ onMonthDay: () => ({ atHour: () => ({ create() {} }) }) }) }) },
  Logger: { log() {} },
  Utilities: {
    computeDigest: (a, s) => Array.from(Buffer.from(require('crypto').createHash('sha256').update(s).digest())),
    DigestAlgorithm: { SHA_256: 'sha256' },
    base64EncodeWebSafe: b => Buffer.from(b).toString('base64url'),
    formatDate: (d, tz, f) => f === 'yyyy-MM' ? `2026-${String(d.getMonth() + 1).padStart(2, '0')}` : 'x' + (++uid),
    getUuid: () => require('crypto').randomUUID(),
  },
  UrlFetchApp: { fetch: url => { const t = decodeURIComponent(url.split('id_token=')[1]); const info = tokenInfo[t];
    return { getResponseCode: () => info ? 200 : 400, getContentText: () => JSON.stringify(info || {}) }; } },
  ContentService: { createTextOutput: s => ({ s, setMimeType() { return this; } }), MimeType: { JSON: 'json' } },
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(process.argv[2], 'utf8'), ctx);

const exp = Math.floor(Date.now() / 1000) + 3600;
function tok(name, email) { tokenInfo[name] = { aud: 'CID', iss: 'https://accounts.google.com', email, email_verified: 'true', exp: String(exp) }; return name; }
const call = body => { const r = ctx.doPost({ postData: { contents: JSON.stringify(body) } }); return JSON.parse(r.s); };

// setup
ctx.setup();
props.CLIENT_ID = 'CID';
assert(sheets.Ledger && !sheets['工作表1']);
const dad = tok('t_dad', 'dad@gmail.com');
const gpa = tok('t_gpa', 'grandpa@gmail.com');
const stranger = tok('t_x', 'x@gmail.com');

let r = call({ action: 'state', idToken: dad });
assert(r.ok, r.error); assert.equal(r.data.me.role, 'admin'); assert.equal(r.data.tasks.length, 3);

r = call({ action: 'state', idToken: stranger }); assert.equal(r.code, 'NOT_MEMBER');
r = call({ action: 'state', idToken: 'bogus' }); assert.equal(r.code, 'AUTH');
r = call({ action: 'state' }); assert.equal(r.code, 'AUTH');

r = call({ action: 'saveUser', idToken: dad, email: 'Grandpa@gmail.com', name: '爺爺', role: 'giver', active: true }); assert(r.ok, r.error);
r = call({ action: 'saveUser', idToken: gpa, email: 'a@b.com', name: 'x', role: 'admin' }); assert.equal(r.code, 'FORBIDDEN');

r = call({ action: 'deposit', idToken: gpa, amount: 1500, note: '紅包' }); assert(r.ok, r.error);
assert.equal(r.data.balance, 1500); assert.equal(r.data.me.name, '爺爺');
r = call({ action: 'spend', idToken: gpa, amount: 10, note: 'x' }); assert.equal(r.code, 'FORBIDDEN');
r = call({ action: 'spend', idToken: dad, amount: 50 }); assert.equal(r.code, 'BAD_INPUT');
r = call({ action: 'spend', idToken: dad, amount: 50, note: '文具' }); assert(r.ok); assert.equal(r.data.balance, 1450);
r = call({ action: 'deduct', idToken: gpa, amount: 99999 }); assert.equal(r.code, 'INSUFFICIENT');
r = call({ action: 'deposit', idToken: gpa, amount: -5 }); assert.equal(r.code, 'BAD_INPUT');

// stamps: task 整理書桌 = 20 each, 5 stamps -> +30
r = call({ action: 'state', idToken: dad });
const desk = r.data.tasks.find(t => t.name === '整理書桌');
for (let i = 0; i < 5; i++) r = call({ action: 'task', idToken: gpa, taskId: desk.id });
assert(r.ok, r.error);
let st = r.data.stamps.find(s => s.id === desk.id);
assert.equal(st.total, 5); assert.equal(st.cards, 1); assert.equal(st.inCard, 0);
assert.equal(r.data.balance, 1450 + 100 + 30);

// reversing the 5th task reverses the bonus too
const fifth = r.data.ledger.find(x => x.type === 'task');
const dadEntry = r.data.ledger.find(x => x.type === 'spend');
let g = call({ action: 'state', idToken: gpa });
assert.equal(g.data.ledger.find(x => x.id === dadEntry.id).canReverse, false);
r = call({ action: 'reverse', idToken: gpa, id: dadEntry.id }); assert.equal(r.code, 'FORBIDDEN');
r = call({ action: 'reverse', idToken: gpa, id: fifth.id }); assert(r.ok, r.error);
st = r.data.stamps.find(s => s.id === desk.id);
assert.equal(st.total, 4); assert.equal(st.cards, 0);
assert.equal(r.data.balance, 1450 + 80);
r = call({ action: 'reverse', idToken: gpa, id: fifth.id }); assert.equal(r.code, 'BAD_INPUT');
// redo the 5th -> bonus fires again
r = call({ action: 'task', idToken: gpa, taskId: desk.id });
assert.equal(r.data.balance, 1450 + 130);

// by source
const src = Object.fromEntries(r.data.bySource.income.map(x => [x.source, x.amount]));
assert.equal(src['爺爺'], 1500 + 130); assert.equal(r.data.bySource.outgoing, -50);

// goals & rewards
r = call({ action: 'saveGoal', idToken: dad, name: '戰鬥陀螺', target: 2000 }); assert(r.ok, r.error);
r = call({ action: 'saveReward', idToken: dad, name: '寶可夢卡', price: 100 }); assert(r.ok);
r = call({ action: 'redeem', idToken: dad, rewardId: r.data.rewards[0].id }); assert(r.ok, r.error);
assert.equal(r.data.balance, 1480);
const goal = r.data.goals[0]; assert.equal(goal.remain, 520);
r = call({ action: 'goalAchieve', idToken: dad, id: goal.id }); assert.equal(r.code, 'INSUFFICIENT');

// interest: 1480 -> floor(1.48)*10 = 10
assert.equal(r.code, 'INSUFFICIENT');
ctx.monthlyInterest(); ctx.monthlyInterest(); // second call must be no-op
r = call({ action: 'state', idToken: dad });
assert.equal(r.data.balance, 1490);
assert.equal(r.data.ledger.filter(x => x.type === 'interest').length, 1);

// kid view
r = call({ action: 'kid', token: 'wrong' }); assert.equal(r.code, 'BAD_TOKEN');
r = call({ action: 'kid', token: props.KID_TOKEN }); assert(r.ok);
assert.equal(r.data.balance, 1490); assert(!('ledger' in r.data)); assert(!('bySource' in r.data));
assert.equal(r.data.goals[0].name, '戰鬥陀螺');
// kid token cannot be used to write
r = call({ action: 'deposit', token: props.KID_TOKEN, amount: 100 }); assert.equal(r.code, 'AUTH');

// admin cannot demote self
r = call({ action: 'saveUser', idToken: dad, email: 'dad@gmail.com', name: '爸爸', role: 'giver' }); assert.equal(r.code, 'BAD_INPUT');
// settings
r = call({ action: 'saveSettings', idToken: dad, kid_name: '小寶', interest_step: 1000, interest_per_step: 10 }); assert(r.ok, r.error);
assert.equal(r.data.kidName, '小寶');
r = call({ action: 'resetKidToken', idToken: dad }); assert(r.ok); assert.notEqual(r.data.admin.kidToken, undefined);

console.log('ALL BACKEND TESTS PASSED');
console.log(JSON.stringify(sheets.Ledger.rows.slice(0, 3)));
