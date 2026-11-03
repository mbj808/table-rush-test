/**
 * Mad Facts class leaderboard – Google Apps Script web app.
 * Bind to a Google Sheet ("Mad Facts Leaderboard"), run setup() once, then
 * Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 *
 * Tabs:
 *   Roster   – A: name (teacher edits), B: optional note
 *   Scores   – timestamp | name | game | streak   (appended by doPost)
 *   Points   – timestamp | name | game | right | wrong   (made on first use; the board totals these)
 *   Settings – B1: class code (default MAD)
 *   Roster C – Class (v8): A to J (A Eagles, B Hawks, C Wolves, D Bears, E Foxes, F Deer, G Owls, H Falcons, I Turtles, J Coyotes). Blank = A. Moving a kid to another class moves their scores too.
 *   Settings A3:C12 – Classes A..J: id | name | join code (v8; defaults are the animal names; kids may type the name, its singular, or the letter; codes are case-insensitive; teacher can change them)
 *   Passwords – Name | Password   (v6; teacher fills in passwords; made with every Roster name on first use or by running setupPasswords)
 *
 * v9 (TEST): Sessions tab – one row per finished game (homework report).
 *   POST {type:'session', name, code, game, start, end, q, right}  (start/end = epoch ms from the device)
 *   Sessions – Start (HST) | Week of (Mon) | Date | Day | Class | Name | Game | Questions | Right | % correct | Minutes | Homework? | Start ms
 *   Homework – formulas (made by setupSessions): per week + student, Homework?=Yes rows only.
 * GET  ?action=roster  -> {ok:true, names:[...]}
 * v7 classes:
 * GET  ?action=roster&code=X -> {ok:true, cls:'A', clsName:'Class A', names:[only that class]} | {ok:false, error:'class'}
 * GET  ?action=classes -> {ok, classes:[{id,name}]}  (codes are never sent back)
 * board adds classOf:{name:'A'} and classes:[{id,name}]; the page shows one class at a time.
 * pwCheck adds cls to its reply.
 * GET  ?action=avatar&name=X -> {ok, avatar}
 * POST {type:'avatar', name, code, avatar:{...}} saves to the Avatars tab
 * v6 sign in (POST, text/plain JSON; passwords are never sent back):
 *   {type:'pwCheck', name, password} -> {ok, name, token} | {ok:false, error:'pw'|'none'|'tries'}
 *   Compare is trimmed and case-insensitive. 'none' = no password in the sheet yet.
 *   Every other POST must send code = that token (the class code no longer works for saves).
 *   Changing a kid's password in the sheet signs that kid out on the next save.
 * GET  ?action=version -> {ok, version:6, (passwords:true) points:true, timed:true, avatarDelete:true, rooms:true}
 * Head to Head rooms (v5, kept in CacheService for 2 hours, no sheet writes, no points):
 * GET  ?action=roomNew&grade=1..6&op=mul|div|add|sub|mix -> room state with new 4-letter room
 * GET  ?action=roomJoin&room=ABCD  -> room state (player 2), error 'room' or 'full'
 * GET  ?action=roomGet&room=ABCD   -> room state
 * GET  ?action=roomHit&room=ABCD&p=1|2&q=N -> first right answer for problem N scores
 * GET  ?action=roomRematch&room=ABCD -> scores to 0, new seed
 * Room state: {ok, room, seed, grade, op, joined, q, s:[p1,p2], last, done}
 * POST {type:'avatarDelete', name, code} deletes every Avatars row for that name (v4)
 * POST {type:'timed', name, code, mode:'sixty'|'all', score} adds a row to the Timed tab
 *      sixty: score = right minus wrong (0..200). all: finish time in ms incl. 5s per wrong.
 *   Timed    – timestamp | name | mode | score   (made on first use; board keeps each kid's best)
 * POST {type:'points', name, code, game, right, wrong} adds a row to the Points tab
 * GET  ?action=check&code=X -> {ok:true|false}  (class code check)
 * GET  ?action=board   -> {ok:true, today:{overall:[...], games:{game:[...]}}, week:{...}, all:{...}}
 *      each list = [{name, streak}] best streak per name, highest first, top 20.
 * POST text/plain body {"name","code","game","streak"} -> {ok:true} or {ok:false,error}
 */

var SHEET_ID = '1_x-dVyYlQ6SbKfaFvnZoY6KtXGpLuiGDBkdzwEFvNp8'; // optional: set when used as a standalone script; blank = the bound sheet
var TOP_N = 300; // v7: lists hold every class; the page filters one class at a time
var MAX_STREAK = 500;
var MIN_GAP_MS = 3000;
var MAX_BATCH = 200; // most answers in one points batch
var MAX_SIXTY = 200;
var MIN_ALL_MS = 20000, MAX_ALL_MS = 3 * 3600 * 1000;

function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function setup() {
  var ss = ss_();
  var roster = ss.getSheetByName('Roster') || ss.insertSheet('Roster');
  if (roster.getLastRow() === 0) {
    roster.getRange(1, 1, 1, 2).setValues([['Name', 'Note (optional)']]).setFontWeight('bold');
    roster.getRange(2, 1, 6, 1).setValues([['Student A'], ['Student B'], ['Student C'], ['Student D'], ['Student E'], ['Student F']]);
    roster.setFrozenRows(1);
  }
  var scores = ss.getSheetByName('Scores') || ss.insertSheet('Scores');
  if (scores.getLastRow() === 0) {
    scores.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Name', 'Game', 'Streak']]).setFontWeight('bold');
    scores.setFrozenRows(1);
  }
  var settings = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  if (!settings.getRange('B1').getValue()) {
    settings.getRange('A1:B1').setValues([['Class code', 'MAD']]);
    settings.getRange('A1').setFontWeight('bold');
  }
  var points = ss.getSheetByName('Points') || ss.insertSheet('Points');
  if (points.getLastRow() === 0) {
    points.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Name', 'Game', 'Right', 'Wrong']]).setFontWeight('bold');
    points.setFrozenRows(1);
  }
  var avatars = ss.getSheetByName('Avatars') || ss.insertSheet('Avatars');
  if (avatars.getLastRow() === 0) {
    avatars.getRange(1, 1, 1, 3).setValues([['Name', 'Avatar JSON', 'Updated']]).setFontWeight('bold');
    avatars.setFrozenRows(1);
  }
  timedSheet_();
  setupClasses();
  var first = ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 5 && first.getLastRow() === 0) ss.deleteSheet(first);
}

function rosterNames_() {
  var sh = ss_().getSheetByName('Roster');
  if (!sh || sh.getLastRow() < 2) return [];
  var seen = {};
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(function (r) { return String(r[0] || '').trim(); })
    .filter(function (n) { if (!n || seen[n.toLowerCase()]) return false; seen[n.toLowerCase()] = true; return true; });
}

// ===== v8: ten animal classes, ids A-J. Roster column C = class id (blank = A). Settings rows 3-12 = id | name | join code. =====
var CLASS_DEFAULTS = [['A', 'Eagles'], ['B', 'Hawks'], ['C', 'Wolves'], ['D', 'Bears'], ['E', 'Foxes'], ['F', 'Deer'], ['G', 'Owls'], ['H', 'Falcons'], ['I', 'Turtles'], ['J', 'Coyotes']]
  .map(function (c) { return [c[0], c[1], c[1]]; });
var CLASS_IDS = 'ABCDEFGHIJ';
function classList_() {
  var sh = ss_().getSheetByName('Settings');
  var rows = sh && sh.getLastRow() >= 3 ? sh.getRange(3, 1, CLASS_DEFAULTS.length, 3).getValues() : [];
  return CLASS_DEFAULTS.map(function (d, i) {
    var r = rows[i] || [];
    return { id: d[0], name: String(r[1] || '').trim() || d[1], code: String(r[2] || '').trim() || d[2] };
  });
}
function setupClasses() {
  var sh = ss_().getSheetByName('Settings') || ss_().insertSheet('Settings');
  setupClassRows_(sh);
  var ro = ss_().getSheetByName('Roster');
  if (ro && !String(ro.getRange('C1').getValue() || '').trim()) ro.getRange('C1').setValue('Class (A to J)').setFontWeight('bold');
}
// Fills empty class rows and upgrades rows still holding the old v7 defaults (X | Class X | X).
// Rows the teacher renamed are left alone. Returns how many rows were written.
function setupClassRows_(sh) {
  if (!String(sh.getRange('A2').getValue() || '').trim()) sh.getRange('A2:C2').setValues([['Class id', 'Class name', 'Join code']]).setFontWeight('bold');
  var n = 0;
  CLASS_DEFAULTS.forEach(function (d, i) {
    var row = sh.getRange(3 + i, 1, 1, 3), v = row.getValues()[0].map(function (x) { return String(x || '').trim(); });
    var empty = !v[0] && !v[1] && !v[2];
    var oldDefault = v[0] === d[0] && (!v[1] || v[1] === 'Class ' + d[0]) && (!v[2] || v[2].toUpperCase() === d[0]);
    if (empty || oldDefault) { row.setValues([d]); n++; }
  });
  return n;
}
function codeKey_(c) { return String(c || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function codeForms_(name) {
  var k = codeKey_(name), out = [k];
  if (/VES$/.test(k)) out.push(k.slice(0, -3) + 'F');
  if (/ES$/.test(k)) out.push(k.slice(0, -2));
  if (/S$/.test(k)) out.push(k.slice(0, -1)); else out.push(k + 'S');
  return out;
}
function classForCode_(code) {
  var k = codeKey_(code), list = classList_();
  for (var i = 0; i < list.length; i++) {
    var forms = codeForms_(list[i].code).concat(codeForms_(list[i].name), [list[i].id]);
    if (forms.indexOf(k) >= 0 || forms.indexOf(k.replace(/^CLASS/, '')) >= 0) return list[i];
  }
  return null;
}
function classOf_() {
  var sh = ss_().getSheetByName('Roster'), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) {
    var n = String(r[0] || '').trim(); if (!n || out[n] ) return;
    var c = String(r[2] || '').trim().toUpperCase().replace(/^CLASS\s*/, '');
    out[n] = (c.length === 1 && CLASS_IDS.indexOf(c) >= 0) ? c : 'A';
  });
  return out;
}
function rosterForCode_(code) {
  var c = classForCode_(code);
  if (!c) return { ok: false, error: 'class' };
  var of = classOf_();
  return { ok: true, cls: c.id, clsName: c.name, names: rosterNames_().filter(function (n) { return of[n] === c.id; }) };
}

function classCode_() {
  var sh = ss_().getSheetByName('Settings');
  var v = sh ? String(sh.getRange('B1').getValue() || '').trim() : '';
  return v || 'MAD';
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'board';
    if (action === 'roster' && e.parameter.code) return json_(rosterForCode_(e.parameter.code));
    if (action === 'roster') return json_({ ok: true, names: rosterNames_() });
    if (action === 'classes') return json_({ ok: true, classes: classList_().map(function (c) { return { id: c.id, name: c.name }; }) });
    if (action === 'avatar') { var a = avatarRows_()[String(e.parameter.name || '').trim().toLowerCase()]; return json_({ ok: true, avatar: a ? a.data : null }); }
    if (action === 'version') return json_({ ok: true, version: 9, sessions: true, classes: true, points: true, timed: true, avatarDelete: true, rooms: true, passwords: true, locks: true });
    if (action.indexOf('room') === 0) return json_(room_(action, e.parameter));
    if (action === 'check') return json_({ ok: String(e.parameter.code || '').trim().toUpperCase() === classCode_().toUpperCase() });
    if (action === 'board') return json_(board_());
    if (action === 'locks') return json_(locksFor_(e.parameter.name));
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  }
}

function board_() {
  var tz = ss_().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  var now = new Date();
  var todayKey = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  // Week starts Monday (local to the sheet's time zone).
  var dow = Number(Utilities.formatDate(now, tz, 'u')); // 1 = Monday .. 7 = Sunday
  var weekStartKey = Utilities.formatDate(new Date(now.getTime() - (dow - 1) * 86400000), tz, 'yyyy-MM-dd');
  var sh = ss_().getSheetByName('Scores');
  var rows = sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues() : [];
  var buckets = { today: {}, week: {}, all: {} };
  function add(bucket, game, name, streak) {
    [['__all', name], [game, name]].forEach(function (k) {
      var g = bucket[k[0]] || (bucket[k[0]] = {});
      if (!g[name] || streak > g[name]) g[name] = streak;
    });
  }
  rows.forEach(function (r) {
    var ts = r[0] instanceof Date ? r[0] : new Date(r[0]);
    var name = String(r[1] || '').trim(), game = String(r[2] || '').trim(), streak = Number(r[3]);
    if (!name || !game || !(streak > 0) || isNaN(ts.getTime())) return;
    var key = Utilities.formatDate(ts, tz, 'yyyy-MM-dd');
    add(buckets.all, game, name, streak);
    if (key >= weekStartKey) add(buckets.week, game, name, streak);
    if (key === todayKey) add(buckets.today, game, name, streak);
  });
  function list(map) {
    return Object.keys(map || {}).map(function (n) { return { name: n, streak: map[n] }; })
      .sort(function (a, b) { return b.streak - a.streak || a.name.localeCompare(b.name); })
      .slice(0, TOP_N);
  }
  function shape(bucket) {
    var games = {};
    Object.keys(bucket).forEach(function (g) { if (g !== '__all') games[g] = list(bucket[g]); });
    return { overall: list(bucket.__all), games: games };
  }
  var av = avatarRows_(), avatars = {};
  Object.keys(buckets.all.__all || {}).forEach(function (n) { var a = av[n.toLowerCase()]; if (a) avatars[n] = a.data; });
  // Points: 1 for each right answer, minus 1 for each wrong one. A total never goes below 0.
  var ps = ss_().getSheetByName('Points');
  var prow = ps && ps.getLastRow() > 1 ? ps.getRange(2, 1, ps.getLastRow() - 1, 5).getValues() : [];
  var pb = { today: {}, week: {}, all: {} };
  function padd(bucket, game, name, d) {
    [['__all', name], [game, name]].forEach(function (k) {
      var m = bucket[k[0]] = bucket[k[0]] || {};
      m[k[1]] = (m[k[1]] || 0) + d;
    });
  }
  prow.forEach(function (r) {
    if (!(r[0] instanceof Date) || !r[1]) return;
    var day = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd');
    var d = (Number(r[3]) || 0) - (Number(r[4]) || 0);
    var name = String(r[1]), game = String(r[2] || '');
    padd(pb.all, game, name, d);
    if (day >= weekStartKey) padd(pb.week, game, name, d);
    if (day === todayKey) padd(pb.today, game, name, d);
  });
  function plist(map) {
    return Object.keys(map || {}).map(function (n) { return { name: n, points: Math.max(0, map[n]) }; })
      .filter(function (x) { return x.points > 0; })
      .sort(function (a, b) { return b.points - a.points || a.name.localeCompare(b.name); })
      .slice(0, TOP_N);
  }
  function pshape(bucket) {
    var games = {};
    Object.keys(bucket).forEach(function (g) { if (g !== '__all') games[g] = plist(bucket[g]); });
    return { overall: plist(bucket.__all), games: games };
  }
  Object.keys(pb.all.__all || {}).forEach(function (n) { var a = av[n.toLowerCase()]; if (a) avatars[n] = a.data; });
  var timed = timedBoard_();
  timed.sixty.concat(timed.all).forEach(function (x) { var a = av[x.name.toLowerCase()]; if (a) avatars[x.name] = a.data; });
  return { ok: true, version: 3, classOf: classOf_(), classes: classList_().map(function (c) { return { id: c.id, name: c.name }; }), today: shape(buckets.today), week: shape(buckets.week), all: shape(buckets.all), avatars: avatars,
    points: { today: pshape(pb.today), week: pshape(pb.week), all: pshape(pb.all) }, timed: timed };
}

// name(lowercase) -> { row, data }
function avatarRows_() {
  var sh = ss_().getSheetByName('Avatars'), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r, i) {
    var n = String(r[0] || '').trim().toLowerCase();
    if (!n) return;
    try { out[n] = { row: i + 2, data: JSON.parse(r[1]) }; } catch (err) {}
  });
  return out;
}

// v4: remove every Avatars row whose name matches (exact, case-insensitive). Nothing else changes.
function deleteAvatar_(match) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return json_({ ok: false, error: 'busy' });
  var removed = 0;
  try {
    var sh = ss_().getSheetByName('Avatars');
    if (sh && sh.getLastRow() > 1) {
      var key = String(match).trim().toLowerCase();
      var names = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
      for (var i = names.length - 1; i >= 0; i--) {
        if (String(names[i][0] || '').trim().toLowerCase() === key) { sh.deleteRow(i + 2); removed++; }
      }
    }
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true, removed: removed });
}

function saveAvatar_(match, avatar) {
  if (!avatar || typeof avatar !== 'object') return json_({ ok: false, error: 'avatar' });
  var text = JSON.stringify(avatar);
  if (text.length > 4000) return json_({ ok: false, error: 'too big' });
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return json_({ ok: false, error: 'busy' });
  try {
    var sh = ss_().getSheetByName('Avatars');
    var hit = avatarRows_()[match.toLowerCase()];
    if (hit) sh.getRange(hit.row, 2, 1, 2).setValues([[text, new Date()]]);
    else sh.appendRow([match, text, new Date()]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

function savePoints_(match, body) {
  var game = String(body.game || '').trim().slice(0, 60);
  var right = Number(body.right), wrong = Number(body.wrong);
  if (!game) return json_({ ok: false, error: 'game' });
  if (right !== Math.floor(right) || wrong !== Math.floor(wrong) || right < 0 || wrong < 0 || right > MAX_BATCH || wrong > MAX_BATCH || right + wrong < 1) return json_({ ok: false, error: 'points' });
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return json_({ ok: false, error: 'busy' });
  try {
    var ps = ss_().getSheetByName('Points');
    if (!ps) {
      ps = ss_().insertSheet('Points');
      ps.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Name', 'Game', 'Right', 'Wrong']]).setFontWeight('bold');
      ps.setFrozenRows(1);
    }
    ps.appendRow([new Date(), match, game, right, wrong]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

function timedSheet_() {
  var sh = ss_().getSheetByName('Timed');
  if (!sh) {
    sh = ss_().insertSheet('Timed');
    sh.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Name', 'Mode', 'Score']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Best per kid. sixty: highest score first. all: fastest time (ms) first.
function timedBoard_() {
  var sh = ss_().getSheetByName('Timed');
  var best = { sixty: {}, all: {} };
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (r) {
      var name = String(r[1] || '').trim(), mode = String(r[2] || ''), v = Number(r[3]);
      if (!name || !best[mode] || !isFinite(v)) return;
      var cur = best[mode][name];
      if (mode === 'sixty' ? (cur == null || v > cur) : (v > 0 && (cur == null || v < cur))) best[mode][name] = v;
    });
  }
  var sixty = Object.keys(best.sixty).map(function (n) { return { name: n, score: best.sixty[n] }; })
    .sort(function (a, b) { return b.score - a.score || a.name.localeCompare(b.name); }).slice(0, TOP_N);
  var all = Object.keys(best.all).map(function (n) { return { name: n, ms: best.all[n] }; })
    .sort(function (a, b) { return a.ms - b.ms || a.name.localeCompare(b.name); }).slice(0, TOP_N);
  return { sixty: sixty, all: all };
}

function saveTimed_(match, body) {
  var mode = String(body.mode || ''), score = Number(body.score);
  if (score !== Math.floor(score)) return json_({ ok: false, error: 'score' });
  if (mode === 'sixty') { if (score < 0 || score > MAX_SIXTY) return json_({ ok: false, error: 'score' }); }
  else if (mode === 'all') { if (score < MIN_ALL_MS || score > MAX_ALL_MS) return json_({ ok: false, error: 'score' }); }
  else return json_({ ok: false, error: 'mode' });
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return json_({ ok: false, error: 'busy' });
  try {
    timedSheet_().appendRow([new Date(), match, mode, score]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

function doPost(e) {
  var body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return json_({ ok: false, error: 'bad json' }); }
  var name = String(body.name || '').trim();
  var code = String(body.code || '').trim();
  var game = String(body.game || '').trim().slice(0, 60);
  var streak = Number(body.streak);
  var names = rosterNames_();
  var match = names.filter(function (n) { return n.toLowerCase() === name.toLowerCase(); })[0];
  if (String(body.type || '').indexOf('pw') === 0) return json_(pw_(body, match));
  if (!match) return json_({ ok: false, error: 'name' });
  if (!code || code !== pwToken_(match)) return json_({ ok: false, error: 'code' });
  if (body.type === 'avatar') return saveAvatar_(match, body.avatar);
  if (body.type === 'avatarDelete') return deleteAvatar_(match);
  if (body.type === 'points') return savePoints_(match, body);
  if (body.type === 'timed') return saveTimed_(match, body);
  if (body.type === 'session') return saveSession_(match, body);
  if (body.type === 'complete') return saveComplete_(match, body);
  if (!game) return json_({ ok: false, error: 'game' });
  if (!(streak === Math.floor(streak)) || streak < 1 || streak > MAX_STREAK) return json_({ ok: false, error: 'streak' });
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return json_({ ok: false, error: 'busy' });
  try {
    var props = PropertiesService.getScriptProperties();
    var key = 'last_' + match.toLowerCase();
    var last = Number(props.getProperty(key) || 0);
    var now = Date.now();
    if (now - last < MIN_GAP_MS) return json_({ ok: false, error: 'slow down' });
    props.setProperty(key, String(now));
    ss_().getSheetByName('Scores').appendRow([new Date(), match, game, streak]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

// ===== v5: Head to Head rooms. Practice only: nothing is written to the sheet. =====
var ROOM_WIN = 10;
var ROOM_TTL = 7200;
var ROOM_ABC = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
function roomSeed_() { return Math.floor(Math.random() * 1000000); }
function roomOut_(r) { return { ok: true, room: r.room, seed: r.seed, grade: r.grade, op: r.op, joined: r.joined, q: r.q, s: r.s, last: r.last, done: r.done }; }
function room_(action, p) {
  var cache = CacheService.getScriptCache();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, error: 'busy' };
  try {
    if (action === 'roomNew') {
      var grade = Math.max(1, Math.min(6, Math.floor(Number(p.grade)) || 3));
      var op = ['mul', 'div', 'add', 'sub', 'mix'].indexOf(String(p.op)) >= 0 ? String(p.op) : 'add';
      var code = '';
      for (var t = 0; t < 30; t++) {
        code = '';
        for (var i = 0; i < 4; i++) code += ROOM_ABC.charAt(Math.floor(Math.random() * ROOM_ABC.length));
        if (!cache.get('room_' + code)) break;
      }
      var r0 = { room: code, seed: roomSeed_(), grade: grade, op: op, joined: false, q: 0, s: [0, 0], last: 0, done: 0 };
      cache.put('room_' + code, JSON.stringify(r0), ROOM_TTL);
      return roomOut_(r0);
    }
    var key = 'room_' + String(p.room || '').trim().toUpperCase().replace(/[^A-Z]/g, '').slice(0, 4);
    var raw = cache.get(key);
    if (!raw) return { ok: false, error: 'room' };
    var r = JSON.parse(raw);
    if (action === 'roomJoin') {
      if (r.joined) return { ok: false, error: 'full' };
      r.joined = true;
    } else if (action === 'roomHit') {
      var who = Number(p.p), q = Number(p.q);
      if ((who === 1 || who === 2) && q === r.q && !r.done && r.joined) {
        r.s[who - 1]++;
        r.last = who;
        if (r.s[who - 1] >= ROOM_WIN) r.done = who; else r.q++;
      }
    } else if (action === 'roomRematch') {
      if (r.done) { r.seed = roomSeed_(); r.q = 0; r.s = [0, 0]; r.last = 0; r.done = 0; }
    } else if (action !== 'roomGet') {
      return { ok: false, error: 'unknown action' };
    }
    cache.put(key, JSON.stringify(r), ROOM_TTL);
    return roomOut_(r);
  } finally {
    lock.releaseLock();
  }
}

// ===== v6: kid passwords. The teacher types them in the Passwords tab. Never returned to the page. =====
// Makes the Passwords tab if missing and adds any Roster name that is not there yet (blank password).
function pwSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName('Passwords');
  if (!sh) { sh = ss.insertSheet('Passwords'); sh.appendRow(['Name', 'Password']); sh.setFrozenRows(1); sh.getRange('B:B').setNumberFormat('@'); }
  var have = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0] || '').trim().toLowerCase()] = 1; });
  rosterNames_().forEach(function (n) { if (!have[n.toLowerCase()]) sh.appendRow([n, '']); });
  return sh;
}
// Run once from the editor to make the tab right away (it is also made on the first sign in).
function setupPasswords() { pwSheet_(); }
function pwGet_(sh, match) {
  if (sh.getLastRow() < 2) return '';
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0] || '').trim().toLowerCase() === match.toLowerCase()) return String(v[i][1] || '').trim();
  return '';
}
// The save token comes from the sheet password, so changing a password signs that kid out on the next save.
function pwTok_(match, saved) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, match.toLowerCase() + '|' + saved.toLowerCase() + '|' + pwSalt_())).slice(0, 32); }
function pwToken_(match) { var saved = pwGet_(pwSheet_(), match); return saved ? pwTok_(match, saved) : ''; }
function pw_(body, match) {
  if (!match) return { ok: false, error: 'name' };
  if (String(body.type) !== 'pwCheck') return { ok: false, error: 'unknown' };
  var cache = CacheService.getScriptCache(), triesKey = 'pwtries_' + match.toLowerCase();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, error: 'busy' };
  try {
    var tries = Number(cache.get(triesKey) || 0);
    if (tries >= 10) return { ok: false, error: 'tries' };
    var saved = pwGet_(pwSheet_(), match);
    if (!saved) return { ok: false, error: 'none' };
    if (String(body.password || '').trim().toLowerCase() !== saved.toLowerCase()) { cache.put(triesKey, String(tries + 1), 300); return { ok: false, error: 'pw' }; }
    cache.remove(triesKey);
    return { ok: true, name: match, token: pwTok_(match, saved), cls: classOf_()[match] || 'A' };
  } finally {
    lock.releaseLock();
  }
}
function pwSalt_() {
  var p = PropertiesService.getScriptProperties(), s = p.getProperty('pw_salt');
  if (!s) { s = Utilities.getUuid(); p.setProperty('pw_salt', s); }
  return s;
}


// ===== v9 (TEST): per-game session log. The teacher reads the Sessions and Homework tabs in the sheet. =====
var TZ = 'Pacific/Honolulu';
var SCHOOL_START = 7 * 60 + 45, SCHOOL_END = 14 * 60 + 5; // 7:45 AM - 2:05 PM Hawaii time, Mon-Fri
var SESS_HEAD = ['Start (HST)', 'Week of (Mon)', 'Date', 'Day', 'Class', 'Name', 'Game', 'Questions', 'Right', '% correct', 'Minutes', 'Homework?', 'Start ms'];
function sessSheet_() {
  var sh = ss_().getSheetByName('Sessions');
  if (!sh) {
    sh = ss_().insertSheet('Sessions');
    sh.getRange(1, 1, 1, SESS_HEAD.length).setValues([SESS_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:D').setNumberFormat('@');
    sh.getRange('J:J').setNumberFormat('0%');
  }
  return sh;
}
// Run once from the editor: makes the Sessions tab and the Homework summary tab.
function setupSessions() {
  sessSheet_();
  var ss = ss_(), hw = ss.getSheetByName('Homework') || ss.insertSheet('Homework');
  hw.clear();
  hw.getRange(1, 1, 1, 9).setValues([['Week of (Mon)', 'Class', 'Name', 'Nights played', 'Minutes', 'Questions', '% correct', 'Games', 'Games played']]).setFontWeight('bold');
  hw.setFrozenRows(1);
  var S = 'Sessions!', HW = S + 'L2:L,"Yes"';
  var crit = function (w, n) { return S + 'B2:B,' + w + ',' + S + 'F2:F,' + n + ',' + HW; };
  hw.getRange('A2').setFormula('=IFERROR(SORT(UNIQUE(FILTER({' + S + 'B2:B,' + S + 'E2:E,' + S + 'F2:F},' + S + 'L2:L="Yes")),1,FALSE,2,TRUE,3,TRUE),)');
  hw.getRange('D2').setFormula('=MAP(A2:A,C2:C,LAMBDA(w,n,IF(n="","",COUNTUNIQUEIFS(' + S + 'C2:C,' + crit('w', 'n') + '))))');
  hw.getRange('E2').setFormula('=MAP(A2:A,C2:C,LAMBDA(w,n,IF(n="","",ROUND(SUMIFS(' + S + 'K2:K,' + crit('w', 'n') + ')))))');
  hw.getRange('F2').setFormula('=MAP(A2:A,C2:C,LAMBDA(w,n,IF(n="","",SUMIFS(' + S + 'H2:H,' + crit('w', 'n') + '))))');
  hw.getRange('G2').setFormula('=MAP(A2:A,C2:C,F2:F,LAMBDA(w,n,q,IF(OR(n="",q=0),"",SUMIFS(' + S + 'I2:I,' + crit('w', 'n') + ')/q)))');
  hw.getRange('H2').setFormula('=MAP(A2:A,C2:C,LAMBDA(w,n,IF(n="","",COUNTIFS(' + crit('w', 'n') + '))))');
  hw.getRange('I2').setFormula('=MAP(A2:A,C2:C,LAMBDA(w,n,IF(n="","",LET(g,UNIQUE(FILTER(' + S + 'G2:G,' + S + 'B2:B=w,' + S + 'F2:F=n,' + S + 'L2:L="Yes")),' +
    'TEXTJOIN(", ",TRUE,MAP(g,LAMBDA(x,x&" \u00D7"&COUNTIFS(' + S + 'G2:G,x,' + crit('w', 'n') + '))))))))');
  hw.getRange('G:G').setNumberFormat('0%');
  hw.setColumnWidth(9, 380);
}
function isHomework_(ms) {
  var dow = Number(Utilities.formatDate(new Date(ms), TZ, 'u')); // 1 = Mon .. 7 = Sun
  var hm = Utilities.formatDate(new Date(ms), TZ, 'HH:mm').split(':'), m = Number(hm[0]) * 60 + Number(hm[1]);
  return dow >= 6 || m < SCHOOL_START || m >= SCHOOL_END;
}
function saveSession_(match, body) {
  var now = Date.now();
  var start = Math.floor(Number(body.start)), end = Math.floor(Number(body.end));
  var q = Math.floor(Number(body.q)), right = Math.floor(Number(body.right));
  var game = String(body.game || '').trim().slice(0, 60);
  if (!game) return json_({ ok: false, error: 'game' });
  if (!(start > now - 14 * 864e5) || start > now + 600000 || !(end >= start) || end - start > 4 * 3600000) return json_({ ok: false, error: 'time' });
  if (!(q >= 0 && q <= 2000) || !(right >= 0 && right <= q)) return json_({ ok: false, error: 'points' });
  var d = new Date(start), dow = Number(Utilities.formatDate(d, TZ, 'u'));
  var monday = Utilities.formatDate(new Date(start - (dow - 1) * 864e5), TZ, 'yyyy-MM-dd') + ' wk'; // text, so sheet formulas never turn it into a date
  var mins = Math.round((end - start) / 6000) / 10;
  var row = [Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm'), monday, Utilities.formatDate(d, TZ, 'yyyy-MM-dd'), Utilities.formatDate(d, TZ, 'EEE'),
    classOf_()[match] || 'A', match, game, q, right, q ? right / q : '', mins, isHomework_(start) ? 'Yes' : 'No', start];
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return json_({ ok: false, error: 'busy' });
  try {
    sessSheet_().appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

// ===== v10 (locks): topic complete & lock, controlled by the sheet. =====
// Locks tab:     Game id | Game | On (ON/OFF). Teacher sets ON to let a game lock. Multiplication is never listed.
// Completed tab: Kid | Class | Game id | Game | Date. The game adds a row when a kid finishes (4 streaks of 5).
//                Delete a row to unlock that kid (all devices). Clear the tab to reset everyone.
var LOCK_GAMES = [["multiples", "Multiples"], ["divide", "Division"], ["add", "Addition"], ["subtract", "Subtraction"], ["maketen", "Make Ten"], ["decimals", "Order Decimals"], ["rounding", "Rounding"], ["fraction", "Fractions"], ["placevalue", "Place Value"], ["evenodd", "Even or Odd"], ["prime", "Prime or Composite"], ["clock", "Telling Time"], ["money", "Counting Money"], ["oceanmap", "Ocean map"], ["ocean", "Oceans"], ["sport", "Sports"], ["hawaii", "Hawaiian Islands"], ["capital", "Capital trivia"], ["state", "Capital map"], ["usstate", "US States"], ["element", "Elements"], ["flag", "Flags"], ["flagtrivia", "Flag trivia"], ["rhyme", "Rhyming"], ["synonym", "Synonyms & Antonyms"], ["piano", "Piano Notes"], ["instrument", "Instruments"], ["habitat", "Animal Habitats"], ["vehicle", "Vehicles"], ["vehicletrivia", "Vehicle trivia"], ["wonder", "Wonders of the World"], ["athlete", "Famous Sports People"], ["nfl", "NFL Team Logos"], ["mlb", "MLB Team Logos"], ["bor", "Bill of Rights"], ["nounverb", "Noun or Verb"], ["punct", "Punctuation"], ["fragrun", "Fragment or Run-on"], ["propcom", "Proper or Common Noun"], ["concab", "Concrete or Abstract Noun"], ["cartoon", "Cartoon Characters"], ["measure", "Measurement"], ["vocabfill", "Vocabulary"], ["analogy", "Analogies"], ["sci-animals", "Animals & Habitats"], ["sci-plants", "Plants"], ["sci-body", "Human Body"], ["sci-earth", "Earth & Weather"], ["sci-space", "Oceans & Space"], ["sci-matter", "Matter & Energy"], ["sci-forces", "Forces & Magnets"], ["sci-elec", "Electricity & Simple Machines"], ["sci-scientists", "Scientists & Inventions"], ["sci-hawaii", "Hawaii Science"]];
function lockSheet_() {
  var ss = ss_(), sh = ss.getSheetByName('Locks');
  if (!sh) {
    sh = ss.insertSheet('Locks');
    sh.getRange(1, 1, 1, 3).setValues([['Game id', 'Game', 'On (ON/OFF)']]).setFontWeight('bold');
    sh.getRange(2, 1, LOCK_GAMES.length, 3).setValues(LOCK_GAMES.map(function (g) { return [g[0], g[1], 'OFF']; }));
    sh.setFrozenRows(1);
  }
  return sh;
}
function doneSheet_() {
  var ss = ss_(), sh = ss.getSheetByName('Completed');
  if (!sh) {
    sh = ss.insertSheet('Completed');
    sh.getRange(1, 1, 1, 5).setValues([['Kid', 'Class', 'Game id', 'Game', 'Date']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
// Run once from the editor (optional; the tabs are also made on first use). Adds any missing games to Locks.
function setupLocks() {
  var sh = lockSheet_(); doneSheet_();
  var have = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim().toLowerCase()] = 1; });
  var add = LOCK_GAMES.filter(function (g) { return !have[g[0].toLowerCase()]; }).map(function (g) { return [g[0], g[1], 'OFF']; });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 3).setValues(add);
}
// game id or game name (as typed in the sheet) -> game id
function lockId_(v) {
  var t = String(v || '').trim().toLowerCase();
  if (!t) return '';
  for (var i = 0; i < LOCK_GAMES.length; i++) if (LOCK_GAMES[i][0].toLowerCase() === t || LOCK_GAMES[i][1].toLowerCase() === t) return LOCK_GAMES[i][0];
  return '';
}
function lockOn_() {
  var sh = lockSheet_(), on = {};
  if (sh.getLastRow() < 2) return on;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) {
    var id = lockId_(r[0]) || lockId_(r[1]);
    var v = String(r[2] === true ? 'ON' : r[2] || '').trim().toUpperCase();
    if (id && (v === 'ON' || v === 'YES' || v === 'TRUE' || v === 'Y' || v === '1')) on[id] = 1;
  });
  return on;
}
function doneFor_(kid) {
  var sh = doneSheet_(), out = {}, k = String(kid || '').trim().toLowerCase();
  if (!k || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (r) {
    if (String(r[0] || '').trim().toLowerCase() !== k) return;
    var id = lockId_(r[2]) || lockId_(r[3]);
    if (id) out[id] = 1;
  });
  return out;
}
function locksFor_(name) {
  var on = lockOn_(), done = name ? Object.keys(doneFor_(name)) : [];
  return { ok: true, on: on, done: done };
}
function saveComplete_(match, body) {
  var id = lockId_(body.game);
  if (!id) return json_({ ok: false, error: 'game' });
  if (!lockOn_()[id]) return json_({ ok: false, error: 'off' });
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return json_({ ok: false, error: 'busy' });
  try {
    if (doneFor_(match)[id]) return json_({ ok: true, dup: true });
    var name = id; LOCK_GAMES.forEach(function (g) { if (g[0] === id) name = g[1]; });
    doneSheet_().appendRow([match, classOf_()[match] || 'A', id, name, new Date()]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}
