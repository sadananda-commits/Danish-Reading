/**
 * Reading Club – Danish Newspaper  |  Google Apps Script backend
 *
 * Creates (only if missing) three tabs in your spreadsheet. Your existing tabs are not touched:
 *   Articles : ID | Date | Name | URL | Headline | Updated
 *   Names    : Name
 *   Settings : Key | Value      (meetingTime, specialQuiz)
 *
 * Deploy as Web app:  Execute as = Me,  Who has access = Anyone.
 */

const SPREADSHEET_ID = '1P6p2rTSuxYhJ18CjTYDTB6Uyj_s1UTcbl-R-n3vSJZM';
const DEFAULT_MEETING_TIME = '06:00';
const MAX_HEADLINE_WORDS = 20;
const SETTING_KEYS = ['meetingTime', 'specialQuiz'];

const TABS = {
  articles: { name: 'Articles', headers: ['ID', 'Date', 'Name', 'URL', 'Headline', 'Updated'] },
  names:    { name: 'Names',    headers: ['Name'] },
  settings: { name: 'Settings', headers: ['Key', 'Value'] }
};

/* ---------- HTTP entry points ---------- */

function doGet(e) {
  try {
    const action = (e.parameter.action || 'init');
    if (action === 'init') {
      return json_({
        ok: true,
        rows: readRows_(),
        names: readNames_(),
        settings: readSettings_()
      });
    }
    if (action === 'headline') {
      return json_({ ok: true, headline: getHeadline_(e.parameter.url || '') });
    }
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case 'saveRow':     return json_({ ok: true, id: saveRow_(body.row) });
      case 'addName':     return json_({ ok: true, names: addName_(body.name) });
      case 'saveSetting': saveSetting_(body.key, body.value); return json_({ ok: true });
      default:            return json_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Sheet helpers ---------- */

function getSheet_(key) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const t = TABS[key];
  let sh = ss.getSheetByName(t.name);
  if (!sh) {
    sh = ss.insertSheet(t.name);
    sh.getRange(1, 1, 1, t.headers.length).setValues([t.headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // Keep dates / times as plain text so Sheets does not reformat them
    if (key === 'articles') sh.getRange('B:B').setNumberFormat('@');
    if (key === 'settings') sh.getRange('B:B').setNumberFormat('@');
  }
  return sh;
}

function readRows_() {
  const values = getSheet_('articles').getDataRange().getDisplayValues();
  return values.slice(1)
    .filter(r => r[0])
    .map(r => ({ id: r[0], date: r[1], name: r[2], url: r[3], headline: r[4] }));
}

function readNames_() {
  const values = getSheet_('names').getDataRange().getDisplayValues();
  const names = values.slice(1).map(r => r[0]).filter(Boolean);
  return names.sort((a, b) => a.localeCompare(b));
}

function readSettings_() {
  const values = getSheet_('settings').getDataRange().getDisplayValues().slice(1);
  const out = { meetingTime: DEFAULT_MEETING_TIME, specialQuiz: '' };
  values.forEach(r => { if (SETTING_KEYS.indexOf(r[0]) > -1) out[r[0]] = r[1]; });
  if (!out.meetingTime) out.meetingTime = DEFAULT_MEETING_TIME;
  return out;
}

/* ---------- Writes ---------- */

function saveRow_(row) {
  if (!row || !row.id) throw new Error('Missing row id');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(row.date || '')) throw new Error('Invalid date');
  const name = String(row.name || '').trim();
  if (!name) throw new Error('Name is required');
  if (!/^https?:\/\//i.test(row.url || '')) throw new Error('URL must start with http:// or https://');

  const sh = getSheet_('articles');
  const record = [row.id, row.date, name, row.url, limitWords_(row.headline || ''), new Date()];

  const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  let target = -1;
  for (let i = 1; i < ids.length; i++) {
    if (ids[i][0] === row.id) { target = i + 1; break; }
  }
  if (target === -1) target = sh.getLastRow() + 1;

  sh.getRange(target, 2).setNumberFormat('@'); // keep date as text
  sh.getRange(target, 1, 1, record.length).setValues([record]);

  addName_(name); // make sure the name exists in the dropdown list
  return row.id;
}

function addName_(name) {
  name = String(name || '').trim();
  if (!name) throw new Error('Name is empty');
  const sh = getSheet_('names');
  const existing = readNames_();
  const found = existing.some(n => n.toLowerCase() === name.toLowerCase());
  if (!found) sh.appendRow([name]);
  return readNames_();
}

function saveSetting_(key, value) {
  if (SETTING_KEYS.indexOf(key) === -1) throw new Error('Unknown setting');
  value = String(value == null ? '' : value).trim();
  if (key === 'meetingTime' && !/^\d{2}:\d{2}$/.test(value)) throw new Error('Invalid time');
  if (key === 'specialQuiz' && value && !/^https?:\/\//i.test(value)) throw new Error('Quiz link must start with http:// or https://');

  const sh = getSheet_('settings');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sh.getRange(i + 1, 2).setNumberFormat('@').setValue(value);
      return;
    }
  }
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 2).setNumberFormat('@');
  sh.getRange(row, 1, 1, 2).setValues([[key, value]]);
}

/* ---------- Headline reader ---------- */

function getHeadline_(url) {
  if (!/^https?:\/\//i.test(url)) throw new Error('Invalid URL');
  let res;
  try {
    res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; ReadingClubBot/1.0)',
        'Accept-Language': 'da,en;q=0.8'
      }
    });
  } catch (err) {
    return '';
  }
  if (res.getResponseCode() >= 400) return '';

  const html = res.getContentText('UTF-8').slice(0, 400000);

  const og =
    matchOne_(html, /<meta[^>]+property=["']og:title["'][^>]*content=["']([^"']+)["']/i) ||
    matchOne_(html, /<meta[^>]+content=["']([^"']+)["'][^>]*property=["']og:title["']/i);
  const h1 = matchOne_(html, /<h1[^>]*>([\s\S]*?)<\/h1>/i);
  const title = matchOne_(html, /<title[^>]*>([\s\S]*?)<\/title>/i);

  const raw = og || h1 || title || '';
  const clean = decodeEntities_(raw.replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim();
  return limitWords_(clean);
}

function matchOne_(text, re) {
  const m = text.match(re);
  return m ? m[1] : '';
}

function limitWords_(text) {
  const words = String(text).trim().split(/\s+/).filter(Boolean);
  return words.slice(0, MAX_HEADLINE_WORDS).join(' ');
}

function decodeEntities_(s) {
  const named = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ',
                  aelig: 'æ', oslash: 'ø', aring: 'å', AElig: 'Æ', Oslash: 'Ø', Aring: 'Å' };
  return s
    .replace(/&#x([0-9a-f]+);/gi, (m, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (m, d) => String.fromCharCode(parseInt(d, 10)))
    .replace(/&([a-zA-Z]+);/g, (m, n) => (named[n] !== undefined ? named[n] : m));
}
