/**
 * RTU GRADE PORTAL — Google Apps Script Web App
 * -------------------------------------------------------------
 *  Student page    : <web app URL>
 *  Instructor page : <web app URL>?page=instructor
 *
 *  Deploy as: Execute as "Me", Who has access "Anyone within rtu.edu.ph".
 *  See README.md for full setup steps.
 *
 *  One record book (Google Sheets) can hold many subjects/sections — one per tab.
 *  Every tab that has a "STUDENT NUMBER" header is read automatically.
 */

const DOMAIN = 'rtu.edu.ph';
// Old versions wrote these tabs into the record book; they are ignored if present.
// Logs now live in a separate "Grade Portal — Activity Log" spreadsheet (see getLogBook_).
const LOG_TABS = ['Portal Log', 'Consultation Requests', 'Email Log'];
const TZ = 'Asia/Manila';

// Logo images hosted on GitHub Pages (docs/ folder). Used for the browser-tab icon and email header.
const LOGO_BASE = 'https://mrvitalez.github.io/Grade-Distribution/';
const FAVICON_URL = LOGO_BASE + 'favicon.png';
const EMAIL_LOGO_URL = LOGO_BASE + 'logo.png';
const CACHE_SECONDS = 120; // students see sheet edits within ~2 minutes

/** Every setting the portal stores (Script Properties), with defaults. */
const DEFAULTS = {
  portalTitle: 'Grade Portal',
  instructorName: '',          // blank = read from the record book ("INSTRUCTOR:")
  instructorEmails: '',        // extra co-instructor emails (comma separated)
  sheetUrl: '',
  sheetId: '',
  tabs: '{}',                  // JSON: { "<tab name>": { include: true, released: false } }
  consultLink: 'https://calendar.app.google/7zMwLhfQeC39F5Ea6',
  consultNote: 'Consultations are held only during my academic consultation hours. Please choose an open slot on the booking page.',
  consultNotify: 'false',      // also email the instructor when a student opens the booking page
  consultEmail: '',
  msgSubject: 'Your Semestral Grade in {{course}}',
  msgBody:
    'Dear {{name}},\n\n' +
    'Your grades for {{course}} ({{section}}) have been released.\n\n' +
    'Initial Grade: {{initialGrade}}\n' +
    'Semestral Grade: {{semestralGrade}}\n' +
    'Remarks: {{remarks}}\n\n' +
    '{{breakdown}}\n\n' +
    'You may also view your grades anytime on the Grade Portal: {{portalLink}}\n\n' +
    'If you have questions about your grade, you may book a consultation during my consultation hours.\n\n' +
    'Respectfully,\n{{instructor}}'
};

/* =============================================================
 *  ROUTING
 * ============================================================= */

function doGet(e) {
  const page = String((e && e.parameter && e.parameter.page) || 'student').toLowerCase();
  const email = getActiveEmail_();
  const s = getSettings_();
  const appUrl = ScriptApp.getService().getUrl();
  let t;

  if (page === 'instructor') {
    if (!isInstructor_(email)) {
      t = HtmlService.createTemplate(
        '<!DOCTYPE html><html><head><base target="_top"><?!= include("Styles"); ?></head><body>' +
        '<div class="center-screen"><div class="card narrow"><h2>Instructor access only</h2>' +
        '<p class="muted">You are signed in as <b><?= email || "an unknown account" ?></b>, which is not an instructor of this portal.</p>' +
        '<a class="btn" href="<?= appUrl ?>">Go to the Student Portal</a></div></div></body></html>');
    } else {
      t = HtmlService.createTemplateFromFile('Instructor');
    }
  } else {
    t = HtmlService.createTemplateFromFile('Student');
  }

  t.email = email;
  t.appUrl = appUrl;
  t.bootJson = safeJson_({
    email: email,
    appUrl: appUrl,
    portalTitle: s.portalTitle,
    instructorName: s.instructorName,
    isInstructor: isInstructor_(email),
    domain: DOMAIN
  });

  const out = t.evaluate()
    .setTitle(page === 'instructor' ? 'Grade Portal · Instructor' : 'Grade Portal · Student')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // allows embedding in Google Sites
  try { out.setFaviconUrl(FAVICON_URL); } catch (err) { /* keep the default icon if the image is unavailable */ }
  return out;
}

/** Used in HTML templates: <?!= include('Styles'); ?> */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* =============================================================
 *  INSTRUCTOR API  (all guarded by assertInstructor_)
 * ============================================================= */

/** Overview of every tab (subject/section) in the linked record book. Always reads fresh. */
function getInstructorDashboard() {
  assertInstructor_();
  const s = getSettings_();
  const appUrl = ScriptApp.getService().getUrl();
  const out = {
    settings: s,
    studentUrl: appUrl,
    instructorUrl: appUrl + '?page=instructor',
    quota: MailApp.getRemainingDailyQuota(),
    linked: !!s.sheetId,
    spreadsheetName: '',
    tabs: [],
    error: '',
    logUrl: logBookUrl_()
  };
  if (!s.sheetId) return out;

  try {
    const ss = SpreadsheetApp.openById(s.sheetId);
    out.spreadsheetName = ss.getName();
    const cfg = getTabConfig_(s);
    const views = viewedMap_();
    ss.getSheets().forEach(function (sh) {
      const name = sh.getName();
      if (LOG_TABS.indexOf(name) >= 0) return;
      const c = cfg[name] || { include: true, released: false };
      const row = { name: name, include: c.include !== false, released: !!c.released, isRecord: false, error: '' };
      try {
        const t = readTab_(ss, name, false);
        if (t) {
          row.isRecord = true;
          row.course = t.course;
          row.section = t.section;
          row.instructor = t.instructor;
          row.stats = t.stats;
          const v = views[name] || {};
          row.stats.viewed = t.records.filter(function (r) { return v[norm_(r.studentNo)]; }).length;
        }
      } catch (err) { row.error = err.message; }
      out.tabs.push(row);
    });
  } catch (err) {
    out.error = err.message;
  }
  return out;
}

/** Full parsed data (all students) of one tab, for Class Records and previews. */
function getTabRecords(tabName) {
  assertInstructor_();
  const s = getSettings_();
  const t = readTab_(SpreadsheetApp.openById(s.sheetId), tabName, false);
  if (!t) throw new Error('"' + tabName + '" does not look like a class record (no "Student Number" header).');
  const v = viewedMap_()[tabName] || {};
  t.records.forEach(function (r) {
    const x = v[norm_(r.studentNo)];
    r.lastViewed = x ? Utilities.formatDate(new Date(x.last), TZ, 'MMM d, yyyy h:mm a') : '';
    r.views = x ? x.count : 0;
  });
  return t;
}

/** Recent portal activity (newest first) and consultation requests, for the Activity tab. */
function getActivity() {
  assertInstructor_();
  const out = { logUrl: logBookUrl_(), visits: [], consultations: [] };
  const ss = existingLogBook_();
  if (!ss) return out;
  const fmt = function (d) { return d instanceof Date ? Utilities.formatDate(d, TZ, 'MMM d, yyyy h:mm a') : String(d); };
  const read = function (name, limit) {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return [];
    const n = Math.min(limit, sh.getLastRow() - 1);
    return sh.getRange(sh.getLastRow() - n + 1, 1, n, sh.getLastColumn()).getValues().reverse();
  };
  out.visits = read('Portal Log', 1000).map(function (r) {
    return { time: fmt(r[0]), email: String(r[1]), studentNo: String(r[2]), name: String(r[3]), result: String(r[4]), tabs: String(r[5] || '') };
  });
  out.consultations = read('Consultation Requests', 300).map(function (r) {
    return { time: fmt(r[0]), tab: String(r[1]), course: String(r[2]), studentNo: String(r[3]), name: String(r[4]),
      email: String(r[5]), grade: String(r[6]), concern: String(r[8]) };
  });
  return out;
}

/** Link (or re-link) the Google Sheets record book. */
function linkRecordBook(form) {
  assertInstructor_();
  form = form || {};
  const id = parseSheetId_(form.sheetUrl);
  if (!id) throw new Error('Please paste a valid Google Sheets link.');

  let ss;
  try { ss = SpreadsheetApp.openById(id); }
  catch (e) { throw new Error('Cannot open that spreadsheet. Make sure you own it or have edit access.'); }

  const s = getSettings_();
  const update = { sheetUrl: String(form.sheetUrl).trim(), sheetId: id };
  if (id !== s.sheetId) update.tabs = '{}';               // new book: start fresh (all tabs unreleased)
  if (s.portalTitle === DEFAULTS.portalTitle) update.portalTitle = ss.getName();
  saveSettings_(update);
  return getInstructorDashboard();
}

function unlinkRecordBook() {
  assertInstructor_();
  saveSettings_({ sheetUrl: '', sheetId: '', tabs: '{}' });
  return getInstructorDashboard();
}

/** Save include / release switches for each tab: { "<tab>": {include, released} } */
function saveTabs(config) {
  assertInstructor_();
  const clean = {};
  Object.keys(config || {}).forEach(function (k) {
    clean[k] = { include: config[k].include !== false, released: !!config[k].released };
  });
  saveSettings_({ tabs: JSON.stringify(clean) });
  return getInstructorDashboard();
}

function saveMessage(form) {
  assertInstructor_();
  saveSettings_({ msgSubject: String(form.msgSubject || ''), msgBody: String(form.msgBody || '') });
  return true;
}

function savePortalSettings(form) {
  assertInstructor_();
  form = form || {};
  saveSettings_({
    portalTitle: String(form.portalTitle || '').trim() || DEFAULTS.portalTitle,
    instructorName: String(form.instructorName || '').trim(),
    instructorEmails: String(form.instructorEmails || '').trim(),
    consultLink: String(form.consultLink || '').trim(),
    consultNote: String(form.consultNote || '').trim(),
    consultNotify: form.consultNotify ? 'true' : 'false',
    consultEmail: String(form.consultEmail || '').trim()
  });
  return getInstructorDashboard();
}

/** Renders the message for one student of one tab (live preview). */
function previewMessage(tabName, studentNo, subject, body) {
  assertInstructor_();
  const s = getSettings_();
  const t = readTab_(SpreadsheetApp.openById(s.sheetId), tabName, true);
  if (!t || !t.records.length) throw new Error('No student records found in ' + tabName + '.');
  const rec = t.records.filter(function (r) { return norm_(r.studentNo) === norm_(studentNo); })[0] || t.records[0];
  return {
    to: recipientFor_(rec),
    subject: fillTemplate_(subject, rec, t, s, false),
    html: fillTemplate_(body, rec, t, s, true)
  };
}

/** Sends one sample email (first student of the tab) to the instructor. */
function sendTestEmail(tabName, subject, body) {
  assertInstructor_();
  const s = getSettings_();
  const t = readTab_(SpreadsheetApp.openById(s.sheetId), tabName, false);
  if (!t || !t.records.length) throw new Error('No student records found in ' + tabName + '.');
  const rec = t.records[0];
  const me = getActiveEmail_();
  MailApp.sendEmail({
    to: me,
    subject: '[TEST] ' + fillTemplate_(subject, rec, t, s, false),
    body: fillTemplate_(body, rec, t, s, false),
    htmlBody: emailShell_(fillTemplate_(body, rec, t, s, true), t.course || s.portalTitle),
    name: instructorName_(s, t) || undefined
  });
  return { to: me, quota: MailApp.getRemainingDailyQuota() };
}

/** Emails the students of one tab (all / failed only / all except failed). */
function sendGradeEmails(opts) {
  assertInstructor_();
  opts = opts || {};
  const s = getSettings_();
  const subject = opts.subject != null ? String(opts.subject) : s.msgSubject;
  const body = opts.body != null ? String(opts.body) : s.msgBody;
  saveSettings_({ msgSubject: subject, msgBody: body });

  const ss = SpreadsheetApp.openById(s.sheetId);
  const t = readTab_(ss, opts.tab, false);
  if (!t) throw new Error('Choose a subject first.');
  let list = t.records.filter(function (r) { return r.semestralGrade; }); // skip students without a grade yet
  if (opts.scope === 'failed') list = list.filter(function (r) { return r.failed; });
  if (opts.scope === 'passed') list = list.filter(function (r) { return !r.failed; });
  if (!list.length) throw new Error('No students with grades match that selection.');

  const quota = MailApp.getRemainingDailyQuota();
  if (list.length > quota) {
    throw new Error('Only ' + quota + ' emails remain in today\'s Gmail quota, but ' + list.length +
      ' are needed. Try "Failed only" first, or send the rest tomorrow.');
  }

  const log = [];
  let sent = 0;
  const errors = [];
  list.forEach(function (rec) {
    const to = recipientFor_(rec);
    try {
      MailApp.sendEmail({
        to: to,
        subject: fillTemplate_(subject, rec, t, s, false),
        body: fillTemplate_(body, rec, t, s, false),
        htmlBody: emailShell_(fillTemplate_(body, rec, t, s, true), t.course || s.portalTitle),
        name: instructorName_(s, t) || undefined
      });
      sent++;
      log.push([new Date(), t.tab, rec.studentNo, rec.name, to, 'Sent', '']);
    } catch (e) {
      errors.push(rec.studentNo + ' (' + to + '): ' + e.message);
      log.push([new Date(), t.tab, rec.studentNo, rec.name, to, 'Error', e.message]);
    }
  });

  try {
    const sh = getOrCreateSheet_(getLogBook_(), 'Email Log',
      ['Timestamp', 'Subject Tab', 'Student Number', 'Name', 'Email', 'Status', 'Error']);
    if (log.length) sh.getRange(sh.getLastRow() + 1, 1, log.length, log[0].length).setValues(log);
  } catch (e) { /* logging is optional */ }

  return { sent: sent, errors: errors, remaining: MailApp.getRemainingDailyQuota() };
}

/* =============================================================
 *  STUDENT API
 * ============================================================= */

function getStudentGrades(form) {
  form = form || {};
  const email = getActiveEmail_();
  const instructor = isInstructor_(email);
  const typed = String(form.studentNo || '').trim();
  const typedName = [form.lastName, form.firstName, form.middleName]
    .map(function (x) { return String(x || '').trim(); }).join(' | ');
  const s = getSettings_();

  if (!typed) throw new Error('Please enter your student number.');

  // 1. Must be a signed-in institutional account whose email matches the typed student number.
  if (!instructor) {
    const local = studentLocalPart_(email);
    if (norm_(typed) !== norm_(local)) {
      logAccess_(email, typed, typedName, 'Denied: student number does not match email', '');
      throw new Error('The student number you entered does not match your institutional email (' + email +
        '). Please check it and try again.');
    }
  }
  if (!s.sheetId) throw new Error('The record book has not been set up yet. Please contact your instructor.');

  // 2. Look for the student in every included (and, for students, released) tab.
  const ss = SpreadsheetApp.openById(s.sheetId);
  const cfg = getTabConfig_(s);
  const visibleTabs = ss.getSheets().map(function (sh) { return sh.getName(); }).filter(function (n) {
    if (LOG_TABS.indexOf(n) >= 0) return false;
    const c = cfg[n] || { include: true, released: false };
    return c.include !== false && (instructor || c.released);
  });
  if (!visibleTabs.length) throw new Error('Grades have not been released yet. Please check back later.');

  const subjects = [];
  visibleTabs.forEach(function (name) {
    let t;
    try { t = readTab_(ss, name, true); } catch (e) { return; }
    if (!t) return;
    const rec = t.records.filter(function (r) { return norm_(r.studentNo) === norm_(typed); })[0];
    if (!rec) return;
    subjects.push({
      tab: name,
      course: t.course || name,
      section: t.section,
      instructor: instructorName_(s, t),
      released: instructor ? !!(cfg[name] && cfg[name].released) : true,
      record: rec,
      // The page already shows the breakdown, so leave {{breakdown}} out of the on-page message.
      messageHtml: fillTemplate_(s.msgBody.replace(/\{\{\s*breakdown\s*\}\}\n*/g, ''), rec, t, s, true),
      consult: rec.failed && s.consultLink ? { note: s.consultNote, link: s.consultLink } : null
    });
  });

  if (!subjects.length) {
    logAccess_(email, typed, typedName, 'Not found in any released class', '');
    throw new Error('Your student number was not found in any class with released grades. ' +
      'If you think this is a mistake, please contact your instructor.');
  }

  logAccess_(email, typed, typedName, instructor ? 'Instructor preview' : 'Viewed',
    subjects.map(function (x) { return x.tab; }).join(', '));

  return { preview: instructor, portalTitle: s.portalTitle, subjects: subjects };
}

/**
 * Called when a student with "Failed" remarks opens the consultation booking page.
 * Logs the request (and optionally emails the instructor). The booking itself happens
 * on the Google Calendar appointment page, which only offers the instructor's open slots.
 */
function logConsultation(form) {
  form = form || {};
  const email = getActiveEmail_();
  const instructor = isInstructor_(email);
  const s = getSettings_();
  const typed = String(form.studentNo || '').trim();

  if (!instructor && norm_(typed) !== norm_(studentLocalPart_(email))) {
    throw new Error('Your student number does not match your institutional email.');
  }
  const ss = SpreadsheetApp.openById(s.sheetId);
  const t = readTab_(ss, form.tab, true);
  const rec = t && t.records.filter(function (r) { return norm_(r.studentNo) === norm_(typed); })[0];
  if (!rec) throw new Error('Student record not found.');
  if (!rec.failed) throw new Error('Consultation is only available for students with "Failed" remarks.');

  const concern = String(form.concern || '').trim().slice(0, 2000);
  try {
    getOrCreateSheet_(getLogBook_(), 'Consultation Requests',
      ['Timestamp', 'Subject Tab', 'Course', 'Student Number', 'Name', 'Email', 'Semestral Grade', 'Remarks', 'Concern', 'Status'])
      .appendRow([new Date(), t.tab, t.course, rec.studentNo, rec.name, email, rec.semestralGrade, rec.remarks, concern,
        'Opened booking page']);
  } catch (e) { /* optional */ }

  if (s.consultNotify === 'true' && !instructor) {
    const cache = CacheService.getScriptCache();
    const key = 'consult_' + norm_(typed) + '_' + norm_(t.tab);
    if (!cache.get(key)) {
      cache.put(key, '1', 1800);
      const rows = [['Subject', t.course + (t.section ? ' (' + t.section + ')' : '')], ['Student Number', rec.studentNo],
        ['Name', rec.name], ['Email', email], ['Semestral Grade', rec.semestralGrade], ['Concern', concern || '—']];
      try {
        MailApp.sendEmail({
          to: s.consultEmail || Session.getEffectiveUser().getEmail(),
          replyTo: email,
          subject: 'Consultation: ' + rec.name + ' (' + rec.studentNo + ') · ' + t.course,
          body: rows.map(function (r) { return r[0] + ': ' + r[1]; }).join('\n') +
            '\n\nThe student was sent to your consultation booking page.',
          htmlBody: emailShell_('<p>A student with <b>Failed</b> remarks opened your consultation booking page.</p>' +
            kvTableHtml_(rows) + '<p style="color:#5E6170">Check your Google Calendar for the booked slot.</p>', t.course)
        });
      } catch (e) { /* optional */ }
    }
  }
  return { ok: true, link: s.consultLink };
}

/* =============================================================
 *  RECORD BOOK PARSER
 * ============================================================= */

const RE = {
  studentNo: /student\s*(no|num|#|id)|stud\.?\s*no|^id\s*(no\.?|number)?$/i,
  last: /last\s*name|surname|family\s*name/i,
  first: /first\s*name|given\s*name/i,
  middle: /middle\s*(name|initial)|^m\.?\s*i\.?$/i,
  name: /learner|full\s*name|student'?s?\s*name|^names?$|name\s*of/i,
  initial: /initial/i,
  semestral: /semestral|semester\s*grade|sem\.?\s*grade/i,
  semestralAlt: /final\s*(semestral\s*)?grade|final\s*rating|^grade$/i,
  adjusted: /adjusted/i,
  remarks: /remark/i,
  email: /e-?mail/i,
  scoresRow: /score|hps|highest|perfect|max/i,
  skipSub: /#$|^date$/i,
  summary: /total|average|\bave\.?\b|\bavg\b|weighted|percent|%|summary|class\s*standing|equivalent|term\s*grade/i
};

/** Reads one tab (with a short cache for student lookups). Returns null if it isn't a class record. */
function readTab_(ss, name, useCache) {
  const cache = CacheService.getScriptCache();
  const key = 'tab_' + ss.getId().slice(0, 20) + '_' + Utilities.base64EncodeWebSafe(name).slice(0, 80);
  if (useCache) {
    const hit = getCached_(cache, key);
    if (hit) return hit.none ? null : hit;
  }
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Tab "' + name + '" was not found in the record book.');
  const parsed = parseSheet_(sh);
  putCached_(cache, key, parsed || { none: true });
  return parsed;
}

function parseSheet_(sh) {
  const range = sh.getDataRange();
  const values = range.getDisplayValues();
  const merges = function (row, numRows) {
    return sh.getRange(row, 1, numRows, values[0].length).getMergedRanges().map(function (m) {
      return { r1: m.getRow(), r2: m.getLastRow(), c1: m.getColumn(), c2: m.getLastColumn(), v: m.getDisplayValue() };
    });
  };
  const t = parseValues_(values, merges);
  if (t) t.tab = sh.getName();
  return t;
}

/**
 * Core parser (pure function, testable outside Apps Script).
 * values: 2-D display values; getMerges(row1, numRows) -> [{r1,r2,c1,c2,v}] (1-based).
 */
function parseValues_(values, getMerges) {
  // 1. Header row = first row with a "Student Number" cell.
  let hr = -1, noCol = -1;
  for (let r = 0; r < Math.min(values.length, 40) && hr < 0; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (RE.studentNo.test(clean_(values[r][c]))) { hr = r; noCol = c; break; }
    }
  }
  if (hr < 0) return null;
  const width = values[0].length;

  // 2. First student row.
  let first = -1;
  for (let r = hr + 1; r < Math.min(values.length, hr + 20); r++) {
    if (isStudentNo_(values[r][noCol])) { first = r; break; }
  }
  if (first < 0) first = Math.min(values.length, hr + 1);

  // 3. Header block (may span several rows) with merged cells spread across their area.
  const block = [];
  for (let r = hr; r < first; r++) block.push(values[r].map(clean_));
  if (getMerges && first > hr) {
    getMerges(hr + 1, first - hr).forEach(function (m) {
      const v = clean_(m.v);
      for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) {
        const bi = r - 1 - hr;
        if (bi >= 0 && bi < block.length && c - 1 < width) block[bi][c - 1] = v;
      }
    });
  }

  // 4. "SCORES" / highest-possible-score row inside the header block.
  let scoresIdx = -1;
  for (let i = 1; i < block.length; i++) if (RE.scoresRow.test(block[i][noCol] || '')) { scoresIdx = i; break; }
  const headRows = block.filter(function (_, i) { return i !== scoresIdx; });

  const parts = [];
  for (let c = 0; c < width; c++) {
    const p = [];
    headRows.forEach(function (row) { const v = row[c]; if (v && p.indexOf(v) < 0) p.push(v); });
    parts.push(p);
  }
  const full = parts.map(function (p) { return p.join(' '); });
  const find = function (re, exclude) {
    exclude = exclude || [];
    for (let i = 0; i < width; i++) if (full[i] && re.test(full[i]) && exclude.indexOf(i) < 0) return i;
    return -1;
  };

  // 5. Key columns.
  const col = { studentNo: noCol };
  col.last = find(RE.last);
  col.first = find(RE.first);
  col.middle = find(RE.middle);
  col.name = find(RE.name, [noCol, col.last, col.first, col.middle]);
  col.initial = find(RE.initial);
  col.semestral = find(RE.semestral);
  if (col.semestral < 0) col.semestral = find(RE.semestralAlt, [col.initial]);
  col.adjusted = find(RE.adjusted);
  col.remarks = find(RE.remarks);
  col.email = find(RE.email);
  if (col.name < 0 && col.last < 0) throw new Error('No "Learner\'s Name" column found (header row ' + (hr + 1) + ').');
  if (col.semestral < 0) throw new Error('No "Semestral Grade" column found (header row ' + (hr + 1) + ').');

  const keyCols = Object.keys(col).map(function (k) { return col[k]; }).filter(function (i) { return i >= 0; });
  const keyLabels = keyCols.map(function (i) { return full[i]; });

  // 6. Student rows.
  const rows = [];
  for (let r = first; r < values.length; r++) {
    if (isStudentNo_(values[r][noCol])) rows.push(r);
  }
  const hasData = [];
  for (let c = 0; c < width; c++) {
    hasData[c] = rows.some(function (r) { return clean_(values[r][c]) !== ''; });
  }

  // 7. Which columns are Summary and which are Breakdown.
  const usable = function (c) {
    if (keyCols.indexOf(c) >= 0 || !hasData[c] || !parts[c].length) return false;
    if (keyLabels.indexOf(full[c]) >= 0) return false;                  // part of a merged key header
    const sub = parts[c].slice(1).join(' ');
    if (RE.skipSub.test(sub) || RE.skipSub.test(parts[c][0])) return false;
    return true;
  };
  const nameEnd = Math.max(noCol, col.name, col.last, col.first, col.middle);
  const gradeEnd = Math.max(col.initial, col.semestral, col.adjusted, col.remarks);
  let summaryCols = [], breakdownCols = [];
  if (col.initial > nameEnd) {
    // Class-record layout: [ID][Name] | weighted components | Initial / Semestral | raw scores ...
    for (let c = nameEnd + 1; c < col.initial; c++) if (usable(c)) summaryCols.push(c);
    for (let c = gradeEnd + 1; c < width; c++) if (usable(c)) breakdownCols.push(c);
  } else {
    for (let c = 0; c < width; c++) if (usable(c)) (RE.summary.test(full[c]) ? summaryCols : breakdownCols).push(c);
  }

  const colInfo = function (c) {
    const group = parts[c][0];
    // Item name first, date last: "E1 · 9/17/2026"
    const subs = parts[c].slice(1);
    const dates = subs.filter(isDateLike_), others = subs.filter(function (x) { return !isDateLike_(x); });
    let label = others.concat(dates).join(' · ') || group;
    if (ABBR[label.toUpperCase()]) label = ABBR[label.toUpperCase()];
    if (/^\d+$/.test(label)) label = (/exam/i.test(group) ? 'Exam ' : 'No. ') + label;
    let max = scoresIdx >= 0 ? clean_(block[scoresIdx][c]) : '';
    if (!/^-?\d+(\.\d+)?$/.test(max)) max = '';
    if (/^wa$/i.test(label)) {
      label = 'Weighted Score';
      if (max && Number(max) <= 1) max = String(Math.round(Number(max) * 10000) / 100);
    }
    const attendance = /attendance/i.test(group) && !max && dates.length > 0 && !others.length;
    return { c: c, group: group, label: label, max: max ? fmtNum_(max) : '', attendance: attendance };
  };
  const groupsOf = function (cols) {
    const groups = [];
    cols.map(colInfo).forEach(function (ci) {
      const last = groups[groups.length - 1];
      if (last && last.title === ci.group) last.cols.push(ci);
      else groups.push({ title: ci.group, cols: [ci] });
    });
    return groups;
  };
  const notZeroMax = function (c) { return !(scoresIdx >= 0 && /^0+(\.0+)?$/.test(clean_(block[scoresIdx][c]))); };
  summaryCols = summaryCols.filter(notZeroMax);    // e.g. QUIZZES with a highest score of 0 (not yet given)
  breakdownCols = breakdownCols.filter(notZeroMax);
  const summaryGroups = groupsOf(summaryCols);
  const breakdownGroups = groupsOf(breakdownCols);

  // 8. Build records.
  const cell = function (row, c) { return c >= 0 ? clean_(row[c]) : ''; };
  const pick = function (row, groups) {
    return groups.map(function (g) {
      return {
        title: g.title,
        items: g.cols.map(function (ci) {
          let v = fmtNum_(cell(row, ci.c));
          if (ci.attendance) v = v === '1' ? 'Present' : v === '0' ? 'Absent' : v;
          return { label: ci.label, value: v, max: ci.max };
        })
      };
    });
  };
  const seen = {};
  const records = [];
  rows.forEach(function (r) {
    const row = values[r];
    const no = cell(row, noCol);
    if (seen[norm_(no)]) return; // ignore duplicate rows of the same student
    seen[norm_(no)] = true;
    let name = cell(row, col.name);
    if (!name && col.last >= 0) {
      const ln = cell(row, col.last), fn = cell(row, col.first), mn = cell(row, col.middle);
      name = ln + (fn ? ', ' + fn : '') + (mn ? ' ' + mn : '');
    }
    const sem = cell(row, col.semestral);
    const adj = cell(row, col.adjusted);
    const rm = computeRemarks_(adj || sem, cell(row, col.remarks));
    records.push({
      row: r + 1,
      studentNo: no,
      name: name.replace(/\s+,/g, ','),
      email: cell(row, col.email),
      initialGrade: fmtNum_(cell(row, col.initial)),
      semestralGrade: sem,
      adjustedGrade: adj,
      remarks: rm.remarks,
      failed: rm.failed,
      summary: pick(row, summaryGroups),
      breakdown: pick(row, breakdownGroups)
    });
  });

  const meta = function (re) {
    for (let r = 0; r < hr; r++) {
      for (let c = 0; c < values[r].length; c++) {
        if (re.test(clean_(values[r][c]))) {
          for (let k = c + 1; k < values[r].length; k++) if (clean_(values[r][k])) return clean_(values[r][k]);
        }
      }
    }
    return '';
  };

  const colLabel = {};
  Object.keys(col).forEach(function (k) { colLabel[k] = col[k] >= 0 ? full[col[k]] : ''; });
  const withGrade = records.filter(function (x) { return x.semestralGrade; });

  return {
    course: meta(/^(course|subject)(\s*(title|code))?\s*:?$/i),
    section: meta(/^(year\s*(and|&)\s*section|section)\s*:?$/i),
    instructor: meta(/^instructor\s*:?$/i),
    headerRow: hr + 1,
    scoresRow: scoresIdx >= 0 ? hr + 1 + scoresIdx : 0,
    firstDataRow: first + 1,
    columns: colLabel,
    summaryGroups: summaryGroups.map(function (g) { return { title: g.title, labels: g.cols.map(function (x) { return x.label; }) }; }),
    breakdownGroups: breakdownGroups.map(function (g) { return { title: g.title, labels: g.cols.map(function (x) { return x.label; }) }; }),
    records: records,
    stats: {
      total: records.length,
      graded: withGrade.length,
      passed: withGrade.filter(function (x) { return /pass/i.test(x.remarks); }).length,
      failed: records.filter(function (x) { return x.failed; }).length
    }
  };
}

/** Philippine college scale: 1.00–3.00 Passed, 5.00 Failed; also supports percentages (< 75 Failed). */
function computeRemarks_(sem, given) {
  const g = String(given || '').trim();
  const s = String(sem || '').trim();
  const n = parseFloat(s.replace(/[^0-9.]/g, ''));
  const isNum = /\d/.test(s) && !isNaN(n);

  if (g) return { remarks: g, failed: /fail/i.test(g) };
  if (!s) return { remarks: 'No grade yet', failed: false };
  if (/fail/i.test(s) || (isNum && ((n >= 5 && n <= 5.0001) || (n > 5 && n < 75)))) return { remarks: 'Failed', failed: true };
  if (isNum && ((n >= 1 && n <= 3) || n >= 75)) return { remarks: 'Passed', failed: false };
  if (/^inc/i.test(s)) return { remarks: 'Incomplete', failed: false };
  if (/dr(o)?p/i.test(s)) return { remarks: 'Dropped', failed: false };
  if (isNum && n === 4) return { remarks: 'Conditional', failed: false };
  return { remarks: '', failed: false };
}

/* =============================================================
 *  MESSAGE TEMPLATES
 * ============================================================= */

function fillTemplate_(tpl, rec, tab, s, html) {
  const map = {
    name: rec.name,
    studentNo: rec.studentNo,
    initialGrade: rec.initialGrade || '—',
    semestralGrade: rec.semestralGrade || '—',
    adjustedGrade: rec.adjustedGrade || '—',
    remarks: rec.remarks || '—',
    course: tab.course || tab.tab,
    section: tab.section || '—',
    instructor: instructorName_(s, tab) || 'Your Instructor',
    portalLink: ScriptApp.getService().getUrl()
  };
  let out = html ? esc_(tpl || '') : String(tpl || '');
  out = out.replace(/\{\{\s*(\w+)\s*\}\}/g, function (m, k) {
    if (k === 'breakdown') return html ? breakdownHtml_(rec) : breakdownText_(rec);
    if (!(k in map)) return m;
    if (html && k === 'portalLink') return '<a href="' + esc_(map[k]) + '" style="color:#1F3E9C">' + esc_(map[k]) + '</a>';
    return html ? esc_(map[k]) : map[k];
  });
  if (html) out = out.replace(/\n/g, '<br>');
  return out;
}

function scoreText_(it) { return (it.value || '—') + (it.max && it.value ? ' / ' + it.max : ''); }

function breakdownHtml_(rec) {
  const groups = rec.summary.concat(rec.breakdown);
  if (!groups.length) return '';
  let h = '<table cellpadding="6" style="border-collapse:collapse;font-size:13px;min-width:300px">';
  groups.forEach(function (g) {
    h += '<tr><th colspan="2" style="background:#1F3E9C;color:#EFE4CF;text-align:left;padding:7px 8px">' + esc_(g.title) + '</th></tr>';
    g.items.forEach(function (it, i) {
      h += '<tr style="background:' + (i % 2 ? '#FFFFFF' : '#FAF6EE') + '"><td style="border-bottom:1px solid #E3D8C3">' +
        esc_(it.label) + '</td><td style="border-bottom:1px solid #E3D8C3;text-align:right;font-weight:600">' +
        esc_(scoreText_(it)) + '</td></tr>';
    });
  });
  return h + '</table>';
}

function breakdownText_(rec) {
  const groups = rec.summary.concat(rec.breakdown);
  if (!groups.length) return '';
  return 'GRADE BREAKDOWN\n' + groups.map(function (g) {
    return g.title + '\n' + g.items.map(function (it) { return '  ' + it.label + ': ' + scoreText_(it); }).join('\n');
  }).join('\n');
}

function kvTableHtml_(rows) {
  return '<table cellpadding="8" style="border-collapse:collapse;font-size:14px">' + rows.map(function (r) {
    return '<tr><td style="border:1px solid #E3D8C3;background:#FAF6EE;font-weight:600">' + esc_(r[0]) +
      '</td><td style="border:1px solid #E3D8C3">' + esc_(r[1]).replace(/\n/g, '<br>') + '</td></tr>';
  }).join('') + '</table>';
}

function emailShell_(inner, heading) {
  return '<div style="font-family:Arial,Helvetica,sans-serif;background:#EFE4CF;padding:24px">' +
    '<div style="max-width:640px;margin:0 auto;background:#FFFFFF;border-radius:12px;overflow:hidden;border:1px solid #E3D8C3">' +
    '<div style="background:#1F3E9C;color:#EFE4CF;padding:18px 24px;border-bottom:4px solid #7A1F35">' +
    '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse"><tr>' +
    '<td style="padding-right:14px;vertical-align:middle"><img src="' + EMAIL_LOGO_URL + '" width="44" height="44" alt="Grade Portal" style="display:block;border-radius:10px"></td>' +
    '<td style="vertical-align:middle;color:#EFE4CF"><div style="font-size:12px;letter-spacing:2px;text-transform:uppercase;opacity:.85">Grade Portal</div>' +
    '<div style="font-size:18px;font-weight:bold;margin-top:4px">' + esc_(heading) + '</div></td></tr></table></div>' +
    '<div style="padding:24px;color:#1D2033;font-size:14px;line-height:1.6">' + inner + '</div>' +
    '<div style="background:#FAF6EE;color:#5E6170;font-size:11px;padding:12px 24px">This message was sent through the course Grade Portal.</div>' +
    '</div></div>';
}

/* =============================================================
 *  HELPERS
 * ============================================================= */

function getSettings_() {
  const p = PropertiesService.getScriptProperties().getProperties();
  const s = {};
  Object.keys(DEFAULTS).forEach(function (k) { s[k] = p[k] !== undefined ? p[k] : DEFAULTS[k]; });
  return s;
}

function saveSettings_(obj) {
  const clean = {};
  Object.keys(obj).forEach(function (k) { if (k in DEFAULTS) clean[k] = String(obj[k] == null ? '' : obj[k]); });
  PropertiesService.getScriptProperties().setProperties(clean, false);
}

function getTabConfig_(s) {
  try { return JSON.parse(s.tabs || '{}') || {}; } catch (e) { return {}; }
}

function instructorName_(s, tab) { return s.instructorName || (tab && tab.instructor) || ''; }

function getActiveEmail_() {
  try { return (Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) { return ''; }
}

function isInstructor_(email) {
  if (!email) return false;
  email = email.toLowerCase();
  const owner = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const extra = (PropertiesService.getScriptProperties().getProperty('instructorEmails') || '')
    .toLowerCase().split(/[\s,;]+/).filter(String);
  return email === owner || extra.indexOf(email) >= 0;
}

function assertInstructor_() {
  if (!isInstructor_(getActiveEmail_())) throw new Error('Instructor access only.');
}

/** Returns the part before @ for a valid institutional email, or throws. */
function studentLocalPart_(email) {
  if (!email) throw new Error('We could not detect your email. Please sign in with your institutional account (studentnumber@' + DOMAIN + ').');
  const parts = email.split('@');
  if (parts[1] !== DOMAIN) throw new Error('Please sign in with your institutional email (studentnumber@' + DOMAIN + '). You are signed in as ' + email + '.');
  return parts[0];
}

function recipientFor_(rec) {
  if (rec.email && /@/.test(rec.email)) return rec.email;
  return String(rec.studentNo).replace(/\s+/g, '').toLowerCase() + '@' + DOMAIN;
}

function logAccess_(email, typedNo, typedName, result, tabs) {
  try {
    getOrCreateSheet_(getLogBook_(), 'Portal Log',
      ['Timestamp', 'Signed-in Email', 'Student Number Entered', 'Name Entered (Last | First | Middle)', 'Result', 'Subjects Viewed'])
      .appendRow([new Date(), email, typedNo, typedName, result, tabs]);
    CacheService.getScriptCache().remove('viewedMap_n'); // invalidate the chunked cache
  } catch (e) { /* optional */ }
}

/**
 * The portal's own log spreadsheet ("Grade Portal — Activity Log" in your Drive).
 * Created on first use. The record book itself is NEVER written to.
 */
function getLogBook_() {
  const ss = existingLogBook_();
  if (ss) return ss;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const again = existingLogBook_();
    if (again) return again;
    const created = SpreadsheetApp.create('Grade Portal — Activity Log');
    PropertiesService.getScriptProperties().setProperty('logSheetId', created.getId());
    return created;
  } finally { lock.releaseLock(); }
}

function existingLogBook_() {
  const id = PropertiesService.getScriptProperties().getProperty('logSheetId');
  if (!id) return null;
  try { return SpreadsheetApp.openById(id); } catch (e) { return null; }
}

function logBookUrl_() { const ss = existingLogBook_(); return ss ? ss.getUrl() : ''; }

/** { "<tab>": { "<normalized student no>": { last: ms, count: n } } } from the Portal Log. */
function viewedMap_() {
  const cache = CacheService.getScriptCache();
  const hit = getCached_(cache, 'viewedMap');
  if (hit) return hit;
  const map = {};
  const ss = existingLogBook_();
  const sh = ss && ss.getSheetByName('Portal Log');
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(function (r) {
      if (String(r[4]) !== 'Viewed' || !(r[0] instanceof Date)) return;
      const no = norm_(r[2]), t = r[0].getTime();
      String(r[5] || '').split(/\s*,\s*/).filter(String).forEach(function (tab) {
        const m = map[tab] = map[tab] || {};
        const x = m[no] = m[no] || { last: 0, count: 0 };
        x.count++;
        if (t > x.last) x.last = t;
      });
    });
  }
  putCached_(cache, 'viewedMap', map);
  return map;
}

function getOrCreateSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    // A brand-new log spreadsheet has one empty "Sheet1": reuse it for the first log tab.
    const blank = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 && /^Sheet1$/.test(ss.getSheets()[0].getName());
    sh = blank ? ss.getSheets()[0].setName(name) : ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1F3E9C').setFontColor('#EFE4CF');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Chunked CacheService helpers (each cache value must stay under 100 KB). */
function putCached_(cache, key, obj) {
  try {
    const str = JSON.stringify(obj), size = 24000, map = {};
    let n = 0;
    for (let i = 0; i < str.length; i += size) map[key + '_' + (n++)] = str.slice(i, i + size);
    map[key + '_n'] = String(n);
    cache.putAll(map, CACHE_SECONDS);
  } catch (e) { /* cache is optional */ }
}
function getCached_(cache, key) {
  try {
    const n = parseInt(cache.get(key + '_n'), 10);
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(key + '_' + i);
    const got = cache.getAll(keys);
    let str = '';
    for (let i = 0; i < keys.length; i++) { if (got[keys[i]] == null) return null; str += got[keys[i]]; }
    return JSON.parse(str);
  } catch (e) { return null; }
}

const ABBR = { Q: 'Quizzes', E: 'Exercises', PS: 'Problem Sets', AS: 'Assignments', LR: 'Laboratory Reports', PR: 'Experiment Worksheets' };

function isDateLike_(v) {
  return /^\d{1,4}[\/-]\d{1,2}([\/-]\d{2,4})?$/.test(v) || /^[A-Z][a-z]{2,8}\.? \d{1,2}(, ?\d{4})?$/.test(v);
}

/** Student-number-like value: has at least 3 digits and is not a header word. */
function isStudentNo_(v) { v = clean_(v); return /\d{3,}/.test(v) && !/[a-z]{4,}/i.test(v); }

/** Compare student numbers ignoring case, spaces and dashes: "2021-101234" == "2021101234". */
function norm_(x) { return String(x || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }

function clean_(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }

/** Rounds long decimals (e.g. 18.81656805 -> 18.82); leaves everything else as typed. */
function fmtNum_(v) {
  v = clean_(v);
  if (/^-?\d+\.\d{3,}$/.test(v)) return String(Math.round(Number(v) * 100) / 100);
  return v;
}

function parseSheetId_(url) {
  url = String(url || '').trim();
  const m = url.match(/\/d\/([a-zA-Z0-9-_]{20,})/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9-_]{25,}$/.test(url)) return url;
  return '';
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function safeJson_(o) { return JSON.stringify(o).replace(/</g, '\\u003c'); }
