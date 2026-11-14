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
      if (all[date] && date !== editing) return fail('taken');
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
      if (prev && editing === date) {
        s.getRange(prev.row, 1, 1, COLS.length).setValues([row]);
      } else {
        s.appendRow(row);
        if (prev) s.deleteRow(prev.row);
      }
    } else if (body.action === 'status') {
      if (!isStaff) return fail('forbidden');
      const r = all[date]; if (!r) return fail('gone');
      const st = ['pending', 'confirmed'].indexOf(body.status) >= 0 ? body.status : 'pending';
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
    return json_({ ok: true, bookings: public_(all, token) });
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Run ONCE from the editor: loads the 2026–27 bookings already in the form responses. */
function setupSeed() {
  const s = sheet_();
  if (s.getLastRow() > 1) throw new Error('Bookings tab is not empty — seed skipped.');
  const seed = [
 {
  "date": "2026-09-14",
  "discord": "ienatom",
  "title": "",
  "type": "PvEvP",
  "era": "Modern",
  "map": "",
  "time": "19:00",
  "notes": "",
  "createdAt": 1787148433000
 },
 {
  "date": "2026-10-11",
  "discord": "NightFlier",
  "title": "VIETNAM WAR Rolling Thunder 2",
  "type": "PvP",
  "era": "Cold War",
  "map": "",
  "time": "18:45",
  "notes": "",
  "createdAt": 1789404793000
 },
 {
  "date": "2026-10-19",
  "discord": "Raccoon",
  "title": "Damascus Shield Campaign",
  "type": "PvE",
  "era": "Cold War / Modern",
  "map": "",
  "time": "20:00",
  "notes": "",
  "createdAt": 1784389281000
 },
 {
  "date": "2026-11-02",
  "discord": "Raccoon",
  "title": "Damascus Shield Campaign",
  "type": "PvE",
  "era": "Cold War / Modern",
  "map": "",
  "time": "19:00",
  "notes": "",
  "createdAt": 1784389369000
 },
 {
  "date": "2026-11-16",
  "discord": "Raccoon",
  "title": "Damascus Shield Campaign",
  "type": "PvE",
  "era": "Cold War / Modern",
  "map": "",
  "time": "20:00",
  "notes": "",
  "createdAt": 1784389405000
 },
 {
  "date": "2026-11-18",
  "discord": "Night Flier",
  "title": "VIETNAM WAR Rolling Thunder 3",
  "type": "PvP",
  "era": "Cold War",
  "map": "",
  "time": "19:45",
  "notes": "",
  "createdAt": 1792416792000
 },
 {
  "date": "2026-11-23",
  "discord": "Grampa",
  "title": "",
  "type": "PvE",
  "era": "Cold War",
  "map": "",
  "time": "20:00",
  "notes": "Not available yet (two part Mission)",
  "createdAt": 1784178543000
 },
 {
  "date": "2026-11-30",
  "discord": "Raccoon",
  "title": "Damascus Shield Campaign",
  "type": "PvE",
  "era": "Cold War / Modern",
  "map": "",
  "time": "20:00",
  "notes": "",
  "createdAt": 1784389443000
 },
 {
  "date": "2026-12-14",
  "discord": "Raccoon",
  "title": "Damascus Shield Campaign",
  "type": "PvE",
  "era": "Cold War / Modern",
  "map": "",
  "time": "20:00",
  "notes": "",
  "createdAt": 1784389470000
 },
 {
  "date": "2026-12-16",
  "discord": "I/JG7_Killerfliege",
  "title": "To Paris and back - Hell above the Clouds",
  "type": "PvEvP",
  "era": "WW2",
  "map": "",
  "time": "19:30",
  "notes": "changed 14 to 16",
  "createdAt": 1786957130000
 },
 {
  "date": "2027-01-18",
  "discord": "Grampa",
  "title": "",
  "type": "PvE",
  "era": "Cold War",
  "map": "",
  "time": "20:00",
  "notes": "Not available yet (second part of Mission)",
  "createdAt": 1784178586000
 },
 {
  "date": "2027-02-01",
  "discord": "fuke3853",
  "title": "Operation Spider Web",
  "type": "PvE",
  "era": "Modern",
  "map": "",
  "time": "20:00",
  "notes": "",
  "createdAt": 1784464296000
 },
 {
  "date": "2027-02-08",
  "discord": "Accipiter",
  "title": "Levantine Ghost 2.0",
  "type": "PvP",
  "era": "Cold War",
  "map": "",
  "time": "20:00",
  "notes": "",
  "createdAt": 1788350316000
 },
 {
  "date": "2027-02-15",
  "discord": "George “Gonzo”Cerrone",
  "title": "Iron Tide",
  "type": "PvE",
  "era": "Modern",
  "map": "",
  "time": "19:30",
  "notes": "",
  "createdAt": 1787218555000
 },
 {
  "date": "2027-02-22",
  "discord": "Accipiter",
  "title": "Levantine Ghost 2.0",
  "type": "PvP",
  "era": "Cold War",
  "map": "",
  "time": "20:00",
  "notes": "",
  "createdAt": 1788350347000
 },
 {
  "date": "2027-03-01",
  "discord": "Grampa",
  "title": "",
  "type": "PvE",
  "era": "Modern",
  "map": "",
  "time": "19:15",
  "notes": "changed from 22/02",
  "createdAt": 1794416546000
 },
 {
  "date": "2027-03-08",
  "discord": "Accipiter",
  "title": "Levantine Ghost 2.0",
  "type": "PvP",
  "era": "Cold War",
  "map": "",
  "time": "20:00",
  "notes": "",
  "createdAt": 1788350384000
 },
 {
  "date": "2027-03-22",
  "discord": "Accipiter",
  "title": "Levantine Ghost 2.0",
  "type": "PvP",
  "era": "Cold War",
  "map": "",
  "time": "20:00",
  "notes": "",
  "createdAt": 1788350484000
 },
 {
  "date": "2027-04-05",
  "discord": "Accipiter",
  "title": "Levantine Ghost 2.0",
  "type": "PvP",
  "era": "Cold War",
  "map": "",
  "time": "20:00",
  "notes": "",
  "createdAt": 1788350525000
 },
 {
  "date": "2027-04-19",
  "discord": "fuke3853",
  "title": "Operation Mongoose",
  "type": "PvE",
  "era": "Modern",
  "map": "",
  "time": "19:00",
  "notes": "",
  "createdAt": 1789479968000
 }
];
  seed.forEach(r => s.appendRow(rowOf_(r.date, Object.assign({ status: 'confirmed', createdBy: '', createdAt: r.createdAt }, r))));
}

/*
SETUP (10 minutes)
1. Open the JATF spreadsheet → Extensions → Apps Script. Paste this whole file as Code.gs and save.
2. Project Settings (gear) → Script Properties → Add: STAFF_PIN = a PIN you share only with staff.
3. Select the function "setupSeed" and press Run (authorize when asked). A "Bookings" tab appears with the season's missions.
4. Deploy → New deployment → type "Web app" → Execute as: Me → Who has access: Anyone → Deploy.
5. Copy the Web app URL (ends with /exec) and paste it into API_URL near the top of the script in JATF-mission-booking.html.
6. Host the HTML anywhere (GitHub Pages, your website, Discord file) — everyone sees the same bookings.
After changing this code, use Deploy → Manage deployments → Edit → Version: New version, so the URL stays the same.
*/
