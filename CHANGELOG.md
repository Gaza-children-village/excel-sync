# Changelog - gcv/sync (GCV Sync add-in)

## 0.1.0

- New add-in "GCV Sync.xlam": ribbon "sync now" sends the values of every sheet of a GCV workbook (very hidden sheets excepted) to the school's Google Sheet; "sync settings" stores the web-app link, school code and key per PC.
- One-way: Excel is the source; the Google Sheet is a mirror (tabs replaced on every sync).
- Google side: one Apps Script web app for all schools (apps-script/), one key per school (only its hash is stored), a file of another school is refused, sharing reset on every sync (admins + the school's secretary only, no link sharing, editors cannot share).
- The target sits behind ISyncTarget, so a GCV server can replace Google later.
