/**
 * Chili Cook-Off — Google Sheet backend (Apps Script web app)
 *
 * Setup (one time, ~3 minutes):
 *  1. Make a new Google Sheet (name it whatever, e.g. "Chili Cook-Off").
 *  2. Extensions → Apps Script. Delete what's there, paste this whole file, Save.
 *  3. Change ADMIN_PIN below if you want a different master PIN.
 *  4. Deploy → New deployment → type "Web app".
 *       Execute as: Me    Who has access: Anyone
 *     Click Deploy, approve the permissions, copy the Web app URL (ends in /exec).
 *  5. Paste that URL into config.js in the GitHub repo.
 *
 * The sheet gets three tabs automatically: Config, Judges, Scores.
 * If you edit this code later: Deploy → Manage deployments → pencil → Version: New version → Deploy
 * (keeps the same URL).
 */

const ADMIN_PIN = '2026';

const JUDGE_HEADERS = ['id', 'name', 'pin', 'group', 'joinedAt', 'pick'];
const SCORE_HEADERS = ['key', 'round', 'judgeId', 'judgeName', 'chili',
  'taste', 'texture', 'appearance', 'aroma', 'total', 'notes', 'updatedAt', 'spice'];

function doGet(e) { return respond_(e && e.parameter ? e.parameter : {}); }

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { p = {}; }
  return respond_(p);
}

function respond_(p) {
  let out;
  try {
    const data = typeof p.payload === 'string' ? JSON.parse(p.payload) : (p.payload || {});
    out = route_(p.action, data);
    out.ok = true;
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function route_(action, d) {
  if (action === 'state') return state_();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    switch (action) {
      case 'join': return join_(d);
      case 'score': return score_(d);
      case 'predict': return predict_(d);
      case 'admin': return admin_(d);
      default: throw new Error('Unknown action: ' + action);
    }
  } finally {
    lock.releaseLock();
  }
}

/* ---------- storage helpers ---------- */

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  } else if (headers) {
    // keep header row current when new columns are added
    const cur = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    if (cur.join('|') !== headers.join('|')) sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sh;
}

function rows_(name, headers) {
  const sh = sheet_(name, headers);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, headers.length).getValues();
  return vals.map(function (r, i) {
    const o = { _row: i + 2 };
    headers.forEach(function (h, j) { o[h] = r[j]; });
    return o;
  });
}

function getConfig_() {
  const sh = sheet_('Config');
  const raw = sh.getRange('A1').getValue();
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function setConfig_(cfg) {
  const sh = sheet_('Config');
  sh.getRange('A1').setValue(JSON.stringify(cfg));
  sh.getRange('A2').setValue('↑ app settings (edited from the admin screen — no need to touch)');
}

function publicJudge_(j) {
  return { id: String(j.id), name: String(j.name), group: Number(j.group) || 0, joinedAt: j.joinedAt, pick: String(j.pick || '') };
}

function publicScore_(s) {
  return {
    round: String(s.round), judgeId: String(s.judgeId), judgeName: String(s.judgeName),
    chili: String(s.chili), taste: Number(s.taste), texture: Number(s.texture),
    appearance: Number(s.appearance), aroma: Number(s.aroma), total: Number(s.total),
    notes: String(s.notes || ''), updatedAt: s.updatedAt, spice: Number(s.spice) || 0
  };
}

/* ---------- actions ---------- */

function state_() {
  return {
    config: getConfig_(),
    judges: rows_('Judges', JUDGE_HEADERS).map(publicJudge_),
    scores: rows_('Scores', SCORE_HEADERS).map(publicScore_)
  };
}

function join_(d) {
  const name = String(d.name || '').trim();
  const pin = String(d.pin || '').trim();
  if (!name) throw new Error('Enter your name');
  if (!/^\d{4}$/.test(pin)) throw new Error('PIN must be 4 digits');
  const judges = rows_('Judges', JUDGE_HEADERS);
  const existing = judges.filter(function (j) {
    return String(j.name).trim().toLowerCase() === name.toLowerCase();
  })[0];
  if (existing) {
    if (String(existing.pin) !== pin) throw new Error('That name is taken. Wrong PIN, or pick a different name.');
    return { judge: publicJudge_(existing) };
  }
  // auto-balance into the group with the fewest judges
  const cfg = getConfig_() || {};
  const groupCount = Math.max(1, Number(cfg.groupCount) || 4);
  const counts = [];
  for (let g = 0; g < groupCount; g++) counts.push(0);
  judges.forEach(function (j) { const g = Number(j.group); if (g >= 0 && g < groupCount) counts[g]++; });
  let group = 0;
  counts.forEach(function (c, i) { if (c < counts[group]) group = i; });
  const judge = { id: Utilities.getUuid().slice(0, 8), name: name, pin: pin, group: group, joinedAt: new Date().toISOString() };
  // write the PIN as text so leading zeros survive
  sheet_('Judges', JUDGE_HEADERS).appendRow([judge.id, judge.name, "'" + pin, group, judge.joinedAt, '']);
  return { judge: publicJudge_(judge) };
}

function score_(d) {
  const judges = rows_('Judges', JUDGE_HEADERS);
  const j = judges.filter(function (x) { return String(x.id) === String(d.judgeId); })[0];
  if (!j || String(j.pin) !== String(d.pin)) throw new Error('Please sign in again');
  const cfg = getConfig_() || {};
  if (cfg.phase !== d.round) throw new Error('That round is closed');
  const cats = ['taste', 'texture', 'appearance', 'aroma'];
  const vals = cats.map(function (c) {
    const v = Number(d[c]);
    if (!(v >= 1 && v <= 10)) throw new Error('Scores must be 1–10');
    return Math.round(v);
  });
  const total = vals.reduce(function (a, b) { return a + b; }, 0);
  const sp = Math.round(Number(d.spice));
  const spice = (sp >= 1 && sp <= 5) ? sp : '';
  const key = d.round + '|' + j.id + '|' + d.chili;
  const row = [key, d.round, String(j.id), String(j.name), String(d.chili)]
    .concat(vals).concat([total, String(d.notes || '').slice(0, 300), new Date().toISOString(), spice]);
  const sh = sheet_('Scores', SCORE_HEADERS);
  const existing = rows_('Scores', SCORE_HEADERS).filter(function (s) { return s.key === key; })[0];
  if (existing) sh.getRange(existing._row, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return { saved: key };
}

function predict_(d) {
  const j = rows_('Judges', JUDGE_HEADERS).filter(function (x) { return String(x.id) === String(d.judgeId); })[0];
  if (!j || String(j.pin) !== String(d.pin)) throw new Error('Please sign in again');
  const cfg = getConfig_() || {};
  if (cfg.phase !== 'final') throw new Error('Picks are closed');
  if ((cfg.finalists || []).indexOf(String(d.chili)) < 0) throw new Error('Not a finalist');
  sheet_('Judges', JUDGE_HEADERS).getRange(j._row, 6).setValue(String(d.chili));
  return {};
}

function admin_(d) {
  if (String(d.pin) !== ADMIN_PIN) throw new Error('Wrong master PIN');
  switch (d.op) {
    case 'check':
      return {};
    case 'saveConfig':
      setConfig_(d.config);
      return {};
    case 'setGroup': {
      const j = rows_('Judges', JUDGE_HEADERS).filter(function (x) { return String(x.id) === String(d.judgeId); })[0];
      if (!j) throw new Error('Judge not found');
      sheet_('Judges', JUDGE_HEADERS).getRange(j._row, 4).setValue(Number(d.group));
      return {};
    }
    case 'removeJudge': {
      const j = rows_('Judges', JUDGE_HEADERS).filter(function (x) { return String(x.id) === String(d.judgeId); })[0];
      if (j) sheet_('Judges', JUDGE_HEADERS).deleteRow(j._row);
      return {};
    }
    case 'resetScores': {
      const sh = sheet_('Scores', SCORE_HEADERS);
      if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
      const jh = sheet_('Judges', JUDGE_HEADERS);
      if (jh.getLastRow() > 1) jh.getRange(2, 6, jh.getLastRow() - 1, 1).clearContent();
      return {};
    }
    case 'resetAll': {
      ['Scores', 'Judges'].forEach(function (n) {
        const sh = sheet_(n, n === 'Scores' ? SCORE_HEADERS : JUDGE_HEADERS);
        if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
      });
      if (d.config) setConfig_(d.config);
      return {};
    }
    default:
      throw new Error('Unknown admin op');
  }
}
