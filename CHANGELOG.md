# Changelog - gcv/sync (GCV Sync add-in)

## 0.1.0

- New add-in "GCV Sync.xlam": ribbon "sync now" sends the values of every sheet of a GCV workbook (very hidden sheets excepted) to the school's Google Sheet; "sync settings" stores the school number (5 -> GCV05) per PC.
- The web-app link is built into the add-in; each PC makes its own random key, never shown or typed.
- One-way: Excel is the source; the Google Sheet is a mirror (tabs replaced on every sync).
- Google side: one Apps Script web app for all schools (apps-script/). The first sync of a school creates its Sheet and links the school to that PC (only the key's hash is stored); other PCs are refused. New schools only while REGISTRATION_OPEN is on; unlinkSchool lets a new PC take over. A file of another school is refused; sharing is reset on every sync (no link sharing, editors cannot share).
- Clear Arabic messages for: linked, done, offline, school linked to another PC, school not registered, file of another school, service busy.
- The target sits behind ISyncTarget, so a GCV server can replace Google later.
