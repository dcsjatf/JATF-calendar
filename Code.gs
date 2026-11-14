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
  return json_({ ok: true, bookings: public_(readAll_(), e.parameter.token) });
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
function onFormSubmit(e) { import_('pending'); try { updateDiscord_(); } catch (err) { Logger.log('Discord: ' + err); } }

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
const CAL_URL = 'https://dcsjatf.github.io/JATF-calendar/';
const ICON = { PvE: '🟡', PvP: '🔴', PvEvP: '🔵', Training: '🟢' };

// Run ONCE from the editor: posts the message and refreshes it every hour.
function setupDiscord() {
  if (!PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK')) throw new Error('Add the Script Property DISCORD_WEBHOOK first.');
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'refreshDiscord').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshDiscord').timeBased().everyHours(1).create();
  updateDiscord_();
  Logger.log('Discord message posted. Pin it in the channel.');
}

// Hourly trigger: drops missions that are over and updates the countdowns.
function refreshDiscord() { updateDiscord_(); }

// Run if you deleted the Discord message and want a fresh one.
function resetDiscordMessage() { PropertiesService.getScriptProperties().deleteProperty('DISCORD_MSG_ID'); updateDiscord_(); }

function updateDiscord_(all) {
  const props = PropertiesService.getScriptProperties();
  const hook = props.getProperty('DISCORD_WEBHOOK');
  if (!hook) return;
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
  const extra = (conf.length > 10 ? '\n\n…and ' + (conf.length - 10) + ' more this season.' : '') +
    (pend ? '\n\n⏳ ' + pend + ' booking request' + (pend > 1 ? 's' : '') + ' waiting for staff approval.' : '');
  const payload = {
    username: 'JATF Schedule',
    embeds: [{
      title: '📅 JATF — Upcoming missions',
      url: CAL_URL,
      color: 0xffc000,
      description: (desc + extra + '\n\n[Open the full calendar](' + CAL_URL + ') · [Book a mission](' + CAL_URL + 'booking.html)').slice(0, 4000),
      footer: { text: 'Times shown in your local time zone · updated' },
      timestamp: new Date().toISOString()
    }]
  };
  const opt = { method: 'patch', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
  const id = props.getProperty('DISCORD_MSG_ID');
  if (id) {
    const res = UrlFetchApp.fetch(hook + '/messages/' + id, opt);
    if (res.getResponseCode() < 300) return;
    if (res.getResponseCode() !== 404) throw new Error('Discord edit failed: ' + res.getResponseCode() + ' ' + res.getContentText());
  }
  opt.method = 'post';
  const res = UrlFetchApp.fetch(hook + '?wait=true', opt);
  if (res.getResponseCode() >= 300) throw new Error('Discord post failed: ' + res.getResponseCode() + ' ' + res.getContentText());
  props.setProperty('DISCORD_MSG_ID', JSON.parse(res.getContentText()).id);
}

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
7. Discord (optional): Script Properties → add DISCORD_WEBHOOK = the channel webhook URL.
   Run "setupDiscord" once (authorize). It posts the schedule message; pin it.
   The message updates itself on every booking change and every hour.
After changing this code, use Deploy → Manage deployments → Edit → Version: New version, so the URL stays the same.
*/
