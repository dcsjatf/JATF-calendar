/**
 * JATF Mission Booking — Google Apps Script backend
 * Stores bookings in a "Bookings" tab of this spreadsheet.
 * Setup: see the steps at the bottom of this file.
 */
const SHEET = 'Bookings';
const COLS = ['date', 'discord', 'title', 'type', 'era', 'map', 'time', 'notes', 'status', 'createdBy', 'createdAt'];

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.appendRow(COLS);
    s.setFrozenRows(1);
    s.getRange('A:A').setNumberFormat('@');
    s.getRange('G:G').setNumberFormat('@');
  }
  return s;
}

function readAll_() {
  const v = sheet_().getDataRange().getDisplayValues();
  const out = {};
  for (let i = 1; i < v.length; i++) {
    const r = {};
    COLS.forEach((c, j) => (r[c] = v[i][j]));
    if (/^\d{4}-\d{2}-\d{2}$/.test(r.date)) { r.row = i + 1; out[r.date] = r; }
  }
  return out;
}

// What the browser sees: no tokens, only "mine"
function public_(all, token) {
  const o = {};
  Object.keys(all).forEach(k => {
    const r = all[k];
    o[k] = {
      discord: r.discord, title: r.title, type: r.type, era: r.era, map: r.map,
      time: r.time, notes: r.notes, status: r.status || 'pending',
      createdAt: Number(r.createdAt) || 0,
      mine: !!token && r.createdBy === token
    };
  });
  return o;
}

function clean_(x, n) {
  let s = String(x == null ? '' : x).slice(0, n || 120);
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // block formula injection
  return s;
}

function rowOf_(date, d) {
  return COLS.map(c => {
    if (c === 'date') return "'" + date;
    if (c === 'time') return "'" + clean_(d.time, 5);
    if (c === 'notes' || c === 'title') return clean_(d[c], 300);
    return clean_(d[c]);
  });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.ics) return ics_(String(p.ics));          // ?ics=all (subscription) or ?ics=YYYY-MM-DD (one mission)
  return json_({ ok: true, bookings: public_(readAll_(), p.token) });
}

/* ===== CALENDAR FILES (.ics) for Apple Calendar, Outlook and calendar subscriptions ===== */
function ics_(which) {
  const all = readAll_();
  let keys = Object.keys(all).sort().filter(k => all[k].status === 'confirmed' || all[k].status === 'cancelled');
  if (which !== 'all') keys = keys.filter(k => k === which);
  const esc = t => String(t || '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
  const fmt = ms => Utilities.formatDate(new Date(ms), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const L = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//JATF//Event Calendar//EN', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH',
    'X-WR-CALNAME:JATF Events', 'X-WR-CALDESC:Joint Air Task Force mission schedule',
    'REFRESH-INTERVAL;VALUE=DURATION:PT1H', 'X-PUBLISHED-TTL:PT1H'];
  keys.forEach(k => {
    const r = all[k], d = k.split('-').map(Number), t = (r.time || '20:00').split(':').map(Number);
    const st = Date.UTC(d[0], d[1] - 1, d[2], t[0] || 0, t[1] || 0);
    const name = r.title || r.map || (r.type + ' mission');
    const canc = r.status === 'cancelled';
    L.push('BEGIN:VEVENT', 'UID:jatf-' + k + '@dcsjatf.github.io', 'DTSTAMP:' + fmt(Date.now()),
      'DTSTART:' + fmt(st), 'DTEND:' + fmt(st + 3 * 3600e3),
      'SUMMARY:' + esc((canc ? 'CANCELLED – ' : '') + 'JATF – ' + name),
      'DESCRIPTION:' + esc(r.type + ' ' + r.era + (r.map ? ' · ' + r.map : '') + '\nMission editor: ' + r.discord +
        (r.notes ? '\n' + r.notes : '') + '\n\n' + CAL_URL),
      'URL:' + CAL_URL, 'STATUS:' + (canc ? 'CANCELLED' : 'CONFIRMED'), 'END:VEVENT');
  });
  L.push('END:VCALENDAR');
  // fold lines longer than 74 characters (iCalendar rule)
  const out = L.map(l => { let o = ''; while (l.length > 74) { o += l.slice(0, 74) + '\r\n '; l = l.slice(74); } return o + l; }).join('\r\n');
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.ICAL);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let body = {};
  try {
    body = JSON.parse(e.postData.contents || '{}');
    const token = String(body.token || '');
    const pinSet = PropertiesService.getScriptProperties().getProperty('STAFF_PIN');
    const isStaff = !!pinSet && body.pin === pinSet;
    const s = sheet_();
    let all = readAll_();
    const fail = err => json_({ ok: false, error: err, bookings: public_(all, token) });

    if (body.action === 'checkPin') return isStaff ? json_({ ok: true }) : fail('badpin');

    // Called by the GitHub "Discord poster" workflow after a new poster.png is online
    if (body.action === 'discord') {
      if (body.v) PropertiesService.getScriptProperties().setProperty('DISCORD_IMG_V', String(body.v).replace(/[^a-z0-9]/gi, '').slice(0, 40));
      updateDiscord_(all);
      return json_({ ok: true });
    }

    const date = String(body.date || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return fail('bad');

    if (body.action === 'save') {
      const editing = body.editing ? String(body.editing) : null;
      const prev = editing ? all[editing] : null;
      if (editing && !prev) return fail('gone');
      if (prev && !(isStaff || prev.createdBy === token)) return fail('forbidden');
      if (body.blank && !isStaff) return fail('forbidden');
      if (all[date] && date !== editing && all[date].status !== 'cancelled') return fail('taken');
      const d = body.data || {};
      if (!body.blank && !String(d.discord || '').trim()) return fail('bad');
      const rec = body.blank
        ? { status: 'blocked', title: d.title || d.notes || '', createdBy: token, createdAt: Date.now() }
        : {
            discord: d.discord, title: d.title, type: d.type, era: d.era, map: d.map, time: d.time, notes: d.notes,
            status: prev && editing === date ? prev.status : (isStaff ? 'confirmed' : 'pending'),
            createdBy: prev ? prev.createdBy : token,
            createdAt: prev ? prev.createdAt : Date.now()
          };
      const row = rowOf_(date, rec);
      const target = all[date];
      if (prev && editing === date) {
        s.getRange(prev.row, 1, 1, COLS.length).setValues([row]);
      } else if (target) {                       // replaces a cancelled mission
        s.getRange(target.row, 1, 1, COLS.length).setValues([row]);
        if (prev) s.deleteRow(prev.row);
      } else {
        s.appendRow(row);
        if (prev) s.deleteRow(prev.row);
      }
    } else if (body.action === 'status') {
      if (!isStaff) return fail('forbidden');
      const r = all[date]; if (!r) return fail('gone');
      const st = ['pending', 'confirmed', 'cancelled'].indexOf(body.status) >= 0 ? body.status : 'pending';
      s.getRange(r.row, COLS.indexOf('status') + 1).setValue(st);
    } else if (body.action === 'delete') {
      const r = all[date]; if (!r) return fail('gone');
      if (!(isStaff || r.createdBy === token)) return fail('forbidden');
      s.deleteRow(r.row);
    } else {
      return fail('bad');
    }
    SpreadsheetApp.flush();
    all = readAll_();
    try { updateDiscord_(all); } catch (err) { Logger.log('Discord: ' + err); }
    try { triggerPoster_(); } catch (err) { Logger.log('GitHub: ' + err); }
    return json_({ ok: true, bookings: public_(all, token) });
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ===== IMPORT FROM THE GOOGLE FORM RESPONSES ===== */

// Run ONCE from the editor: copies every form response into the Bookings tab (as confirmed).
function importFromForm() { return import_('confirmed'); }

// Installed as a "From spreadsheet – On form submit" trigger: new form requests arrive as pending.
function onFormSubmit(e) { import_('pending'); try { updateDiscord_(); } catch (err) { Logger.log('Discord: ' + err); } try { triggerPoster_(); } catch (err) { Logger.log('GitHub: ' + err); } }

function import_(status) {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheets().find(sh => sh.getName() !== SHEET &&
    sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0]
      .some(h => /DATE REQUESTED/i.test(h)));
  if (!src) throw new Error('Form responses tab not found (no "DATE REQUESTED" column).');
  const vals = src.getDataRange().getValues();
  const disp = src.getDataRange().getDisplayValues();
  const head = disp[0].map(h => String(h).toUpperCase().trim());
  const col = re => head.findIndex(h => re.test(h));
  const C = {
    ts: 0, discord: col(/DISCORD/), title: col(/MISSION TITLE/), type: col(/MISSION TYPE/),
    era: col(/PLAYSET|ERA/), date: col(/DATE REQUESTED/), time: col(/^TIME/), notes: col(/^NOTES?$/)
  };
  const tz = ss.getSpreadsheetTimeZone();
  const s = sheet_();
  const have = readAll_();
  let added = 0, skipped = [];
  for (let i = 1; i < vals.length; i++) {
    const v = vals[i], d = disp[i];
    const date = toDate_(v[C.date], d[C.date], tz);
    if (!date) continue;
    const discord = String(d[C.discord] || '').trim();
    if (!discord) continue;
    if (have[date]) { if (have[date].discord !== discord) skipped.push(date + ' (' + discord + ')'); continue; }
    const tm = (String(d[C.time]).match(/(\d{1,2})[.:](\d{2})/) || []);
    const time = tm.length ? ('0' + tm[1]).slice(-2) + ':' + tm[2] : '20:00';
    let title = C.title >= 0 ? String(d[C.title]).trim() : '';
    let notes = C.notes >= 0 ? String(d[C.notes]).replace(/\*/g, '').trim() : '';
    if (/^tbd$/i.test(title)) title = '';
    if (/^not available/i.test(title)) { notes = title + (notes ? '; ' + notes : ''); title = ''; }
    const rec = {
      discord: discord, title: title, type: String(d[C.type] || 'PvE').trim(),
      era: era_(d[C.era]), map: '', time: time, notes: notes, status: status, createdBy: '',
      createdAt: v[C.ts] instanceof Date ? v[C.ts].getTime() : Date.now()
    };
    s.appendRow(rowOf_(date, rec));
    have[date] = { discord: discord };
    added++;
  }
  const msg = 'Imported ' + added + ' booking(s).' + (skipped.length ? ' Date already taken, skipped: ' + skipped.join(', ') : '');
  Logger.log(msg);
  return msg;
}

function toDate_(v, shown, tz) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const m = String(shown || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);   // dd/mm/yyyy
  return m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : null;
}

function era_(x) {
  const t = String(x || '');
  const cw = /cold war/i.test(t), mod = /modern/i.test(t), ww2 = /ww2|wwii/i.test(t);
  return cw && mod ? 'Cold War / Modern' : mod ? 'Modern' : cw ? 'Cold War' : ww2 ? 'WW2' : t.trim();
}

/* ===== DISCORD: one pinned message with the upcoming missions, kept up to date ===== */
// Needs Script Property DISCORD_WEBHOOK (never put the webhook URL in public code).
// For several Discord servers, put several webhook URLs in DISCORD_WEBHOOK, separated by a space or comma.
const CAL_URL = 'https://dcsjatf.github.io/JATF-calendar/';
const ICON = { PvE: '🟡', PvP: '🔴', PvEvP: '🔵', Training: '🟢' };

// Run ONCE from the editor: posts the message and refreshes it every hour.
function setupDiscord() {
  if (!hooks_().length) throw new Error('Add the Script Property DISCORD_WEBHOOK first (one or more webhook URLs).');
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'refreshDiscord').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshDiscord').timeBased().everyHours(1).create();
  updateDiscord_();
  Logger.log('Discord message posted. Pin it in the channel.');
}

// Hourly trigger: drops missions that are over and updates the countdowns.
function refreshDiscord() { updateDiscord_(); }

// Run if you deleted the Discord message and want a fresh one.
function resetDiscordMessage() {
  const props = PropertiesService.getScriptProperties();
  Object.keys(props.getProperties()).filter(k => k.indexOf('DISCORD_MSG_ID') === 0).forEach(k => props.deleteProperty(k));
  updateDiscord_();
}

function hooks_() {
  return String(PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK') || '')
    .split(/[\s,]+/).filter(h => /^https:\/\/(discord|discordapp)\.com\/api\/webhooks\//.test(h));
}

function msgKey_(hook) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, hook);
  return 'DISCORD_MSG_ID_' + Utilities.base64EncodeWebSafe(d).replace(/=+$/, '');
}

function updateDiscord_(all) {
  const props = PropertiesService.getScriptProperties();
  const hooks = hooks_();
  if (!hooks.length) return;
  all = all || readAll_();
  const now = Date.now();
  const list = Object.keys(all).sort().map(k => {
    const r = all[k], t = (r.time || '20:00').split(':');
    const p = k.split('-').map(Number);
    return { k: k, r: r, ts: Date.UTC(p[0], p[1] - 1, p[2], Number(t[0]) || 0, Number(t[1]) || 0) };
  }).filter(x => x.r.status !== 'blocked' && x.ts > now - 3 * 3600e3);
  const conf = list.filter(x => x.r.status === 'confirmed' || x.r.status === 'cancelled');
  const pend = list.filter(x => x.r.status === 'pending').length;
  const desc = conf.slice(0, 10).map(x => {
    const r = x.r, u = Math.floor(x.ts / 1000);
    const name = r.title || r.map || (r.type + ' mission');
    if (r.status === 'cancelled') return '❌ ~~<t:' + u + ':F>~~ · **CANCELLED**\n~~' + name + '~~ — ' + r.type + ' · by ' + r.discord;
    return (ICON[r.type] || '⚪') + ' **<t:' + u + ':F>** · <t:' + u + ':R>\n' +
      '**' + name + '** — ' + r.type + ' ' + r.era + (r.map && r.title ? ' · ' + r.map : '') +
      ' · ' + r.time + 'Z · by ' + r.discord;
  }).join('\n\n') || '_No missions scheduled yet._';
  const v = props.getProperty('DISCORD_IMG_V');
  const month = Utilities.formatDate(new Date(), 'UTC', 'MMMM yyyy');
  const embed = {
    title: '📅 JATF Event Schedule — ' + month,
    url: CAL_URL,
    color: 0xffc000,
    footer: { text: 'Updated' },
    timestamp: new Date().toISOString()
  };
  if (v) {
    // Poster of the current month, rendered by the GitHub workflow (same as the portal's poster)
    embed.image = { url: CAL_URL + 'posters/poster-' + v + '.png' };
    embed.description = '[Open the full calendar](' + CAL_URL + ')';
  } else {
    // Fallback until the first poster exists: text list
    embed.description = (desc + (conf.length > 10 ? '\n\n…and ' + (conf.length - 10) + ' more this season.' : '') +
      '\n\n[Open the full calendar](' + CAL_URL + ')').slice(0, 4000);
  }
  const payload = { username: 'JATF Schedule', embeds: [embed] };
  const errors = [];
  hooks.forEach(hook => {
    try { send_(hook, payload, props); } catch (err) { errors.push(String(err)); }
  });
  if (errors.length) throw new Error(errors.join(' | '));
}

function send_(hook, payload, props) {
  const key = msgKey_(hook);
  // message posted by the older single-server version of this script
  const legacy = props.getProperty('DISCORD_MSG_ID');
  if (legacy) {
    if (!props.getProperty(key)) props.setProperty(key, legacy);
    else if (props.getProperty(key) !== legacy)   // a duplicate: remove it, keep one message only
      UrlFetchApp.fetch(hook + '/messages/' + legacy, { method: 'delete', muteHttpExceptions: true });
    props.deleteProperty('DISCORD_MSG_ID');
  }
  const opt = { method: 'patch', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
  const id = props.getProperty(key);
  if (id) {
    const res = UrlFetchApp.fetch(hook + '/messages/' + id, opt);
    if (res.getResponseCode() < 300) return;
    if (res.getResponseCode() !== 404) throw new Error('Discord edit failed: ' + res.getResponseCode() + ' ' + res.getContentText());
  }
  opt.method = 'post';
  const res = UrlFetchApp.fetch(hook + '?wait=true', opt);
  if (res.getResponseCode() >= 300) throw new Error('Discord post failed: ' + res.getResponseCode() + ' ' + res.getContentText());
  props.setProperty(key, JSON.parse(res.getContentText()).id);
}

/* ===== GITHUB: redraw the poster right after a change ===== */
// Needs Script Property GITHUB_TOKEN: a fine-grained token for dcsjatf/JATF-calendar with "Actions: Read and write".
// Without it the poster still updates, but only when GitHub runs its 15-minute schedule (which can lag a lot).
const GH_REPO = 'dcsjatf/JATF-calendar';

function triggerPoster_() {
  const tok = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!tok) return;
  const res = UrlFetchApp.fetch('https://api.github.com/repos/' + GH_REPO + '/actions/workflows/poster.yml/dispatches', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + tok, Accept: 'application/vnd.github+json' },
    payload: JSON.stringify({ ref: 'main' })
  });
  if (res.getResponseCode() >= 300) throw new Error(res.getResponseCode() + ' ' + res.getContentText());
}

// Run from the editor to test the token: the "Discord poster" action should start on GitHub.
function testPosterTrigger() { triggerPoster_(); Logger.log('Poster update requested on GitHub.'); }

/*
SETUP (10 minutes)
1. Open the JATF spreadsheet → Extensions → Apps Script. Paste this whole file as Code.gs and save.
2. Project Settings (gear) → Script Properties → Add: STAFF_PIN = a PIN you share only with staff.
3. Select the function "importFromForm" and press Run (authorize when asked).
   A "Bookings" tab appears with every mission from the form responses.
4. Optional, to keep the Google Form working too: Triggers (alarm clock) → Add trigger →
   function onFormSubmit · event source "From spreadsheet" · event type "On form submit" → Save.
   New form requests then show up in the portal as pending.
5. Deploy → New deployment → type "Web app" → Execute as: Me → Who has access: Anyone → Deploy.
6. Copy the Web app URL (ends with /exec) into API_URL in index.html and booking.html.
7. Discord (optional): Script Properties → add DISCORD_WEBHOOK = the channel webhook URL
   (several servers: several URLs separated by a space).
   Run "setupDiscord" once (authorize). It posts the schedule message; pin it.
   The message updates itself on every booking change and every hour.
   The poster image comes from the GitHub workflow "Discord poster" (runs every 15 minutes).
8. Instant poster updates (optional): Script Properties → add GITHUB_TOKEN (fine-grained token,
   repository dcsjatf/JATF-calendar, permission Actions: Read and write). Run testPosterTrigger once.
After changing this code, use Deploy → Manage deployments → Edit → Version: New version, so the URL stays the same.
*/
