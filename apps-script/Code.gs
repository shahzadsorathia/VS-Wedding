/**
 * Virani & Sorathia — wedding RSVP
 * Google Apps Script backend. Replies are stored in the Google Sheet this script is attached to.
 *
 * Guest page:      <your web app URL>
 * Your dashboard:  tap "Couple's dashboard" at the bottom of the form and enter your password
 *
 * SETTINGS YOU MAY CHANGE LATER live in Project Settings (gear icon) > Script Properties, not in this code:
 *   DASHBOARD_PASSWORD  your dashboard password
 *   SITE_URL            the wedding website address (used by the "Back to the wedding website" links)
 *   ALERT_EMAILS        who gets the RSVP emails, separated by commas
 *   RSVP_OPEN           yes = the form takes replies; no = the form is closed (you switch it yourself)
 * A change there works straight away: no pasting code, no new deployment.
 * The values below are only used when a Script Property is missing.
 */

// ---------- settings ----------
// Fallbacks only: set the real values in Project Settings > Script Properties (see the note at the top).
const DEFAULTS = {
  DASHBOARD_PASSWORD: '',
  SITE_URL: '',
  ALERT_EMAILS: '',
  RSVP_OPEN: 'yes',
};
function setting_(name) { const v = PropertiesService.getScriptProperties().getProperty(name); return (v != null && String(v).trim() !== '') ? String(v).trim() : DEFAULTS[name]; }
function alertEmails_() { return setting_('ALERT_EMAILS').split(/[,;\s]+/).map(s => s.trim()).filter(s => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s)); }
function siteUrl_() { const u = setting_('SITE_URL'); return /^https?:\/\/\S+$/.test(u) ? u : ''; }
const EVENTS = [
  { id: 'sangeet',   name: 'Pithi / Sangeet' },
  { id: 'nikkah',    name: 'Nikkah' },
  { id: 'reception', name: 'Reception' },
];
const FAMILY_SHEET = 'Replies';
const GUEST_SHEET = 'Guests';
const SUMMARY_SHEET = 'Summary';
const KID_LABEL = 'Child (under 10)';

// ---------- web pages ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.view === 'dashboard') {
    const t = HtmlService.createTemplateFromFile('Dashboard');
    t.key = p.key || '';
    t.url = ScriptApp.getService().getUrl();
    t.site = siteUrl_();
    return t.evaluate()
      .setTitle('RSVP Dashboard · Virani & Sorathia')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  const ti = HtmlService.createTemplateFromFile('Index');
  ti.url = ScriptApp.getService().getUrl();
  ti.site = siteUrl_();
  return ti.evaluate()
    .setTitle('RSVP · Virani weds Sorathia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- called from the guest page ----------
function getConfig() {
  return { open: isOpen_() };
}

function findReply(phone) {
  const d = digits_(phone);
  if (d.length !== 10) return null;
  const row = findFamilyRow_(d);
  return row ? rowToFamily_(row.values) : null;
}

function saveReply(payload) {
  if (!isOpen_()) throw new Error('RSVPs are closed. Please message us directly.');
  const clean = validate_(payload);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(FAMILY_SHEET);
    const now = new Date();
    const existing = findFamilyRow_(clean.phone);
    // A brand-new reply must never silently replace another family's reply that used the same phone.
    if (existing && payload.mode === 'new') return { ok: false, conflict: true, existingFamily: String(existing.values[4]) };
    let id, created, isUpdate = false;
    if (existing) {
      id = existing.values[0]; created = existing.values[1]; isUpdate = true;
      sh.getRange(existing.row, 1, 1, 9).setValues([[id, created, now, clean.team, text_(clean.family), "'" + clean.phone, text_(clean.note), JSON.stringify(clean.people), clean.people.length]]);
    } else {
      id = Utilities.getUuid().slice(0, 8);
      created = now;
      sh.appendRow([id, created, now, clean.team, text_(clean.family), "'" + clean.phone, text_(clean.note), JSON.stringify(clean.people), clean.people.length]);
    }
    rewriteGuests_(id, clean, now);
    try { sendInstantAlert_(clean, isUpdate); } catch (err) { console.error('Alert email failed: ' + err); }
    return { ok: true, updated: isUpdate, family: rowToFamily_([id, created, now, clean.team, clean.family, clean.phone, clean.note, JSON.stringify(clean.people)]) };
  } finally {
    lock.releaseLock();
  }
}

// ---------- called from the dashboard ----------
function getDashboard(secret) {
  const real = PropertiesService.getScriptProperties().getProperty('DASHBOARD_KEY');
  const s = String(secret || '');
  const pass = setting_('DASHBOARD_PASSWORD');
  const okPass = pass && pass !== 'CHANGE-ME' && s === pass;
  const okKey = real && s === real;
  if (!okPass && !okKey) { Utilities.sleep(1200); throw new Error('That password isn\'t right.'); }
  const sh = sheet_(FAMILY_SHEET);
  const n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, 8).getValues() : [];
  return { families: rows.filter(r => r[0]).map(rowToFamily_), open: isOpen_() };
}

// ---------- one-time setup: run this once from the editor ----------
function setup() {
  const ss = SpreadsheetApp.getActive();
  const fam = ss.getSheetByName(FAMILY_SHEET) || ss.insertSheet(FAMILY_SHEET);
  if (fam.getLastRow() === 0) {
    fam.appendRow(['ID', 'First replied', 'Last updated', 'Team', 'Family name', 'Phone', 'Message / dietary', 'People (data)', 'Party size']);
    fam.setFrozenRows(1);
    fam.hideColumns(8);
  }
  const g = ss.getSheetByName(GUEST_SHEET) || ss.insertSheet(GUEST_SHEET);
  if (g.getLastRow() === 0) {
    g.appendRow(['Family ID', 'Team', 'Family name', 'First name', 'Last name'].concat(EVENTS.map(e => e.name)).concat(['Phone', 'Last updated', KID_LABEL]));
    g.setFrozenRows(1);
  } else if (!g.getRange(1, 11).getValue()) {
    g.getRange(1, 11).setValue(KID_LABEL);
  }
  const s = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET, 0);
  s.clear();
  s.getRange('A1:E1').setValues([['Event', 'Attending', 'Team Bride', 'Team Groom', 'Kids (under 10)']]).setFontWeight('bold');
  EVENTS.forEach((e, i) => {
    const col = String.fromCharCode('F'.charCodeAt(0) + i); // F, G, H in Guests
    const r = i + 2;
    s.getRange(r, 1, 1, 5).setFormulas([[
      '="' + e.name + '"',
      '=COUNTIF(' + GUEST_SHEET + '!' + col + ':' + col + ',"Attending")',
      '=COUNTIFS(' + GUEST_SHEET + '!' + col + ':' + col + ',"Attending",' + GUEST_SHEET + '!B:B,"Team Bride")',
      '=COUNTIFS(' + GUEST_SHEET + '!' + col + ':' + col + ',"Attending",' + GUEST_SHEET + '!B:B,"Team Groom")',
      '=COUNTIFS(' + GUEST_SHEET + '!' + col + ':' + col + ',"Attending",' + GUEST_SHEET + '!K:K,"Yes")',
    ]]);
  });
  s.getRange('A6:B7').setFormulas([
    ['="Families replied"', '=COUNTA(' + FAMILY_SHEET + '!A2:A)'],
    ['="Guests listed"', '=COUNTA(' + GUEST_SHEET + '!A2:A)'],
  ]);
  s.autoResizeColumns(1, 5);
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('DASHBOARD_KEY');
  if (!key) { key = Utilities.getUuid().replace(/-/g, '').slice(0, 16); props.setProperty('DASHBOARD_KEY', key); }
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailySummary').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailySummary').timeBased().everyDays(1).atHour(20).inTimezone('America/Denver').create();
  const def = ss.getSheetByName('Sheet1'); if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  Logger.log('Setup done. Your private dashboard key is: ' + key);
  Logger.log('Daily summary email scheduled for about 8 PM Denver time, to: ' + alertEmails_().join(', '));
  Logger.log('Dashboard password ' + (setting_('DASHBOARD_PASSWORD') ? 'is set.' : 'is NOT set yet: add DASHBOARD_PASSWORD in Project Settings > Script Properties.'));
  Logger.log('Website link: ' + (siteUrl_() || 'NOT set yet: add SITE_URL in Project Settings > Script Properties (until then the back-to-website links are hidden).'));
  if (!alertEmails_().length) Logger.log('Alert emails: NOT set yet: add ALERT_EMAILS in Project Settings > Script Properties (until then no RSVP emails are sent).');
}

// ---------- emails ----------
function sendInstantAlert_(c, isUpdate) {
  const to = alertEmails_(); if (!to.length) return;
  const lines = EVENTS.map(e => {
    const y = c.people.filter(x => x.a[e.id]);
    const yk = y.filter(x => x.k).length;
    return '<li>' + e.name + ': ' + (y.length ? '<b>' + y.length + ' coming</b>' + (yk ? ' incl. ' + kids_(yk) : '') + ' (' + esc_(y.map(x => x.f + (x.k ? ' (child)' : '')).join(', ')) + ')' : 'not coming') + '</li>';
  }).join('');
  const t = totals_();
  const subject = (isUpdate ? 'Updated RSVP: ' : 'New RSVP: ') + c.family + ' Family (' + c.team + ')';
  const html = mailWrap_(
    '<p style="margin:0 0 10px">' + esc_(c.family) + ' Family ' + (isUpdate ? 'just updated their reply.' : 'just sent their reply.') + '</p>' +
    '<ul style="margin:0 0 12px;padding-left:18px">' + lines + '</ul>' +
    '<p style="margin:0 0 4px;color:#84586B;font-size:13px">Phone: ' + fmtPhone_(c.phone) + '</p>' +
    '<p style="margin:0 0 14px;color:#84586B;font-size:13px">Running totals: ' + EVENTS.map(e => e.name + ' ' + t.byEvent[e.id] + ' (' + kids_(t.kidsByEvent[e.id]) + ')').join(' · ') + '</p>');
  MailApp.sendEmail({ to: to.join(','), subject: subject, htmlBody: html, name: 'Wedding RSVP' });
}

// Runs every evening (set up by setup()). Stops by itself the day after RSVPs close.
function dailySummary() {
  const to = alertEmails_(); if (!to.length) return;
  if (!isOpen_()) return; // no daily emails once you close the RSVPs
  const sh = sheet_(FAMILY_SHEET);
  const n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, 8).getValues().filter(r => r[0]) : [];
  const day = d => d instanceof Date ? Utilities.formatDate(d, 'America/Denver', 'yyyy-MM-dd') : '';
  const today = Utilities.formatDate(new Date(), 'America/Denver', 'yyyy-MM-dd');
  const todays = rows.filter(r => day(r[2]) === today).map(r => String(r[4]));
  const t = totals_();
  const label = Utilities.formatDate(new Date(), 'America/Denver', 'MMM d');
  const html = mailWrap_(
    '<p style="margin:0 0 10px">' + (todays.length ? todays.length + ' ' + (todays.length === 1 ? 'reply' : 'replies') + ' new or updated today: ' + esc_(todays.join(', ')) + '.' : 'No new or updated replies today.') + '</p>' +
    '<ul style="margin:0 0 12px;padding-left:18px">' + EVENTS.map(e => '<li>' + e.name + ': <b>' + t.byEvent[e.id] + '</b> coming (' + (t.byEvent[e.id] - t.kidsByEvent[e.id]) + ' adults, ' + kids_(t.kidsByEvent[e.id]) + ')</li>').join('') + '</ul>' +
    '<p style="margin:0 0 14px;color:#84586B;font-size:13px">So far: ' + t.families + ' ' + (t.families === 1 ? 'family' : 'families') + ' · ' + t.guests + ' ' + (t.guests === 1 ? 'guest' : 'guests') + '</p>');
  MailApp.sendEmail({ to: to.join(','), subject: 'Daily RSVP summary · ' + label, htmlBody: html, name: 'Wedding RSVP' });
}

function totals_() {
  const g = sheet_(GUEST_SHEET);
  const n = g.getLastRow() - 1;
  const rows = n > 0 ? g.getRange(2, 1, n, 8 + EVENTS.length).getValues().filter(r => r[0]) : [];
  const byEvent = {}, kidsByEvent = {};
  EVENTS.forEach((e, i) => {
    byEvent[e.id] = rows.filter(r => r[5 + i] === 'Attending').length;
    kidsByEvent[e.id] = rows.filter(r => r[5 + i] === 'Attending' && r[7 + EVENTS.length] === 'Yes').length;
  });
  const fam = sheet_(FAMILY_SHEET); const fn = fam.getLastRow() - 1;
  const families = fn > 0 ? fam.getRange(2, 1, fn, 1).getValues().filter(r => r[0]).length : 0;
  return { byEvent: byEvent, kidsByEvent: kidsByEvent, families: families, guests: rows.length };
}
function kids_(n) { return n + ' ' + (n === 1 ? 'child' : 'kids'); }
function mailWrap_(inner) {
  let url = ''; try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;color:#4E1A31;line-height:1.5;max-width:520px">' + inner +
    (url ? '<a href="' + url + '?view=dashboard" style="display:inline-block;padding:9px 18px;border-radius:999px;background:#8E2450;color:#ffffff;text-decoration:none;font-size:13px;letter-spacing:1px">OPEN YOUR DASHBOARD</a>' : '') +
    '<p style="margin:18px 0 0;color:#B8873B;font-size:12px">Virani &amp; Sorathia · Wedding RSVP</p></div>';
}
function esc_(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function fmtPhone_(d) { return d.length === 10 ? '(' + d.slice(0, 3) + ') ' + d.slice(3, 6) + '-' + d.slice(6) : d; }

// ---------- helpers ----------
// Store guest-typed text as plain text, so a name starting with = + - @ is never read as a formula.
function text_(s) { s = String(s || ''); return /^[=+\-@]/.test(s) ? "'" + s : s; }
// Open until you set RSVP_OPEN to no (or closed / off / false) in Script Properties.
function isOpen_() { return !/^(no|n|closed|close|off|false|0)$/i.test(setting_('RSVP_OPEN')); }
function digits_(s) { let d = String(s || '').replace(/\D/g, ''); if (d.length === 11 && d[0] === '1') d = d.slice(1); return d; }
// A real US or Canadian number: 10 digits, and neither the area code nor the exchange starts with 0 or 1.
function phoneOk_(d) { return /^[2-9]\d{2}[2-9]\d{6}$/.test(d); }
// Guest-typed names: line breaks and runs of spaces become single spaces.
function clean_(s, max) { return String(s || '').replace(/\s+/g, ' ').trim().slice(0, max); }
function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('The sheet "' + name + '" is missing. Run setup() once from the script editor.');
  return sh;
}
function findFamilyRow_(phone) {
  const sh = sheet_(FAMILY_SHEET);
  const n = sh.getLastRow() - 1;
  if (n < 1) return null;
  const vals = sh.getRange(2, 1, n, 8).getValues();
  for (let i = 0; i < vals.length; i++) {
    const p = digits_(vals[i][5]);
    if (p && p === phone) return { row: i + 2, values: vals[i] };
  }
  return null;
}
function rowToFamily_(r) {
  let people = [];
  try { people = JSON.parse(r[7] || '[]'); } catch (e) {}
  const fmt = d => d instanceof Date ? Utilities.formatDate(d, 'America/Denver', 'MMM d') : String(d || '');
  const when = r[2] || r[1];
  return { id: String(r[0]), at: fmt(when), ts: when instanceof Date ? when.getTime() : 0, side: r[3] === 'Team Bride' ? 'bride' : 'groom', family: String(r[4]), phone: digits_(r[5]), note: String(r[6] || ''), people: people };
}
function validate_(p) {
  if (!p || typeof p !== 'object') throw new Error('Something went wrong. Please try again.');
  const team = p.side === 'bride' ? 'Team Bride' : p.side === 'groom' ? 'Team Groom' : null;
  if (!team) throw new Error('Please choose Team Bride or Team Groom.');
  const family = clean_(p.family, 60);
  if (!family) throw new Error('Please add your family name.');
  const phone = digits_(p.phone);
  if (!phoneOk_(phone)) throw new Error('Please enter a US or Canadian 10-digit number.');
  if (!Array.isArray(p.people) || !p.people.length || p.people.length > 25) throw new Error('Please list who is coming.');
  const people = p.people.map(x => {
    const f = clean_(x.f, 40), l = clean_(x.l, 40);
    if (!f) throw new Error('Please add a first name for every guest.');
    const a = {};
    EVENTS.forEach(e => {
      if (x.a == null || typeof x.a[e.id] !== 'boolean') throw new Error('Please choose for ' + f + ' at every celebration.');
      a[e.id] = x.a[e.id];
    });
    return { f: f, l: l, k: x.k === true, a: a };
  });
  return { team: team, family: family, phone: phone, note: String(p.note || '').trim().slice(0, 500), people: people };
}
function rewriteGuests_(id, c, now) {
  const g = sheet_(GUEST_SHEET);
  const n = g.getLastRow() - 1;
  if (n > 0) {
    const ids = g.getRange(2, 1, n, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(id)) g.deleteRow(i + 2);
  }
  const rows = c.people.map(x => [id, c.team, text_(c.family), text_(x.f), text_(x.l)].concat(EVENTS.map(e => x.a[e.id] ? 'Attending' : "Can't make it")).concat(["'" + c.phone, now, x.k ? 'Yes' : '']));
  if (rows.length) g.getRange(g.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}
