# JATF Event Calendar & Mission Booking

- `index.html` — public event calendar (read-only, for pilots)
- `booking.html` — mission booking portal for mission editors (staff log in with a PIN)
- `Code.gs` — Google Apps Script backend; bookings live in the "Bookings" tab of the JATF Google Sheet

## Setup
1. In the JATF Google Sheet: Extensions → Apps Script → paste `Code.gs`, save.
2. Project Settings → Script Properties → add `STAFF_PIN`.
3. Run `importFromForm` once (authorize when asked): it copies the Google Form responses into the "Bookings" tab.
   Optional: add a trigger `onFormSubmit` (From spreadsheet → On form submit) so new form responses arrive as pending.
4. Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone.
5. Copy the `/exec` URL into `API_URL` near the top of the script in **both** `index.html` and `booking.html`.
6. Repo Settings → Pages → Source: Deploy from a branch → `main` / root.

Site: `https://dcsjatf.github.io/JATF-calendar/` (calendar) and `.../booking.html` (booking).

After editing `Code.gs`: Deploy → Manage deployments → Edit → New version (the URL stays the same).
