# GCV Sync

![Release](https://img.shields.io/github/v/release/Gaza-children-village/excel-sync?filter=sync-v*&label=release)
![Excel](https://img.shields.io/badge/Excel-2016%2B%20Windows-217346)
![Google Apps Script](https://img.shields.io/badge/Google-Apps%20Script-4285F4)
![Status](https://img.shields.io/badge/status-preview-orange)

Excel add-in (`GCV Sync.xlam`) that copies a school's GCV Administration workbook to that school's Google Sheet, plus the Google Apps Script web app that receives it ([`apps-script/`](apps-script)).

## Ribbon

| Button | What it does |
|---|---|
| **مزامنة الملف الآن** (sync now) | Sends the values of every sheet of the active GCV workbook to the school's Google Sheet, then offers to open it |
| **إعدادات المزامنة** (sync settings) | Web-app link, school code and key, entered once per PC (given by the technical manager) |

## How it works

- **One-way, values only.** Excel is the source; the Google Sheet is a mirror. Every sync replaces its tabs with the workbook's sheets (results of formulas, not formulas). Anything typed in the Google Sheet is overwritten by the next sync.
- IDs stored as text keep their leading zeros, dates stay dates, right-to-left sheets stay right-to-left. Very hidden sheets (Settings) are not sent. Only GCV workbooks (hidden name `GCV_TEMPLATE`) are synced.
- **On demand.** Nothing happens without the button. Offline: a clear message, nothing changes.
- A tab `_sync` in the Google Sheet shows the last sync (time, file, add-in version, rows).
- The sending part (`ISyncTarget`) is separate: a GCV server target can replace Google later without changing the rest.

## Access and privacy

- One Google Sheet per school, created and owned by the GCV Google account.
- Every request needs the school's own key; the script stores only its SHA-256 hash. The key is kept per PC in the registry, never in the workbook.
- A file of another school is refused: the first sync records the workbook's `SCHOOL_NAME`; a different name is not accepted.
- On every sync the sharing is reset: only the admins (`ADMIN_EMAILS`) and that school's secretary can edit, link sharing is off, and editors cannot share further.
- The add-in's logs contain counts and results, never data.

## Google setup (once, in the GCV Google account)

1. **Folder:** create a Drive folder for the school Sheets and copy its ID (the part after `folders/` in its link).
2. **Project:** [script.google.com](https://script.google.com) › New project, name it *GCV Sync*. Create the files `lib.gs` and `Code.gs` with the contents of [`apps-script/lib.js`](apps-script/lib.js) and [`apps-script/Code.js`](apps-script/Code.js). In Project Settings tick *Show "appsscript.json"* and replace it with [`apps-script/appsscript.json`](apps-script/appsscript.json). (Or push the folder with [clasp](https://github.com/google/clasp).)
3. **Properties:** Project Settings › Script properties:
   - `FOLDER_ID`: the folder ID
   - `ADMIN_EMAILS`: admin addresses, comma-separated
4. **Deploy:** Deploy › New deployment › *Web app*, Execute as **Me**, Who has access **Anyone**. Authorise, then copy the link ending in `/exec`.
5. **Each school:** in `Code.gs` edit `addSchoolFromEditor` (code, name, secretary e-mail), select it and press **Run**. The execution log shows the Sheet's link and the school's **KEY** (shown once).
6. **Each school PC:** GCV tab › **إعدادات المزامنة**: the `/exec` link, the school code and the key. Then **مزامنة الملف الآن**.

Maintenance (run from the editor): `setEditors("S03", "new.secretary@…")` when the secretary changes, `resetKey("S03")` when a PC is lost, `clearSchoolName("S03")` when a school is renamed, `listSchools()` for an overview. After changing the script: Deploy › Manage deployments › edit › *New version* (the link stays the same).

## Requirements

- Internet when syncing; GCV Logger (recommended).
- A Google account owned by GCV for the Sheets and the script.

## Building from source

Excel cannot load a VBA project created outside Excel, so the add-in is assembled in Excel from the files in this repository:

1. Copy `GCV Sync - shell.xlam` and rename the copy to `GCV Sync.xlam` (the shell already contains the add-in file type and the ribbon).
2. Open it, press **Alt+F11** and select its VBAProject.
3. **File › Import File**: every file in `src/`.
4. Double-click **ThisWorkbook** and paste the contents of `ThisWorkbook.code`.
5. **Debug › Compile VBAProject**, then save with **Ctrl+S**. Saving records the add-in's id, version and file name in the file.

Source files are ASCII with CRLF line endings. Arabic text is stored as Unicode code points (`FromCodePoints("0627 ...")`), because the VBA editor cannot keep Arabic characters.

## Tests

Run `RunSyncTests` from the VBA editor (click inside the procedure and press **F5**). Results appear in a message box and in the Immediate window (**Ctrl+G**). The tests use fakes and temporary files: no network, no real data.

Before a release, remove the development modules (`TestSync`, `FakeSyncTarget`), compile and save.

## Releasing

Releases are checked and published by GitHub Actions ([`.github/workflows/ci.yml`](.github/workflows/ci.yml)); only the build happens by hand, because it needs Excel.

1. Raise `ADDIN_VERSION` in `src/GcvAddin.bas` and add a `## <version>` section to `CHANGELOG.md`.
2. Build the add-in in Excel (see above) **without** the test modules, and save it.
3. Upload it as a **draft** release:
   ```
   gh release create sync-v<version> "GCV Sync.xlam" --draft --repo Gaza-children-village/excel-sync
   ```
4. Merge to `main` and push the tag `sync-v<version>` (through Gitea; the mirror carries it to GitHub).
5. The workflow checks the uploaded file (same id and version as the tag, add-in type, ribbon, no test modules, CHANGELOG section), adds `manifest.json` (SHA-256 and size), renames the file to its asset name, sets the release notes from `CHANGELOG.md` and **publishes**. If anything is wrong, the release stays a draft (the updater ignores drafts) and the job shows why.

Uploaded the draft after pushing the tag? Re-run the failed **release** job in the Actions tab, or `gh workflow run ci.yml -f tag=sync-v<version> --repo Gaza-children-village/excel-sync`.

Every push also runs the source checks: VBA files ASCII with CRLF, module names equal to file names, `Option Explicit`, the shared ribbon tab, a valid identity in `GcvAddin.bas` with a CHANGELOG section. Run them locally with `python .github/scripts/gcv_ci.py check`.

The Apps Script helpers are tested with Node: `cd apps-script && node --test` (also in CI).

## Files

| File | Purpose |
|---|---|
| `src/SyncUi.bas` | Ribbon, settings, messages |
| `src/SyncCore.bas` | Building the request (JSON, values, dates, school name), answers, settings |
| `src/ISyncTarget.cls`, `src/GoogleScriptTarget.cls`, `src/SyncResult.cls` | Where a sync goes (Google now, GCV server later) |
| `src/UpdJson.bas`, `src/GcvLogClient.bas`, `src/GcvAddin.bas` | JSON reader; logging through GCV Logger; identity of the add-in |
| `src/TestSync.bas`, `src/FakeSyncTarget.cls` | Tests (development only) |
| `apps-script/Code.js`, `apps-script/lib.js` | The web app (Google services) and its pure helpers |
| `apps-script/lib.test.js` | Tests of the helpers (Node) |

## Part of GCV Tools

Excel add-ins for the schools of Gaza Children's Village. They share one ribbon tab, **Gaza Children Village Tools**, and are installed and updated together.

| Add-in | Repository | What it does |
|---|---|---|
| GCV Logger | [excel-logger](https://github.com/Gaza-children-village/excel-logger) | Logs of all GCV add-ins; show logs, report an issue |
| GCV Updater | [excel-updater](https://github.com/Gaza-children-village/excel-updater) | Installs updates and new tools; home of the catalog, loader and installer |
| GCV Student Lookup | [excel-student-lookup](https://github.com/Gaza-children-village/excel-student-lookup) | Fills student rows from the Civil Registry, on demand |
| GCV English Names | [excel-english-names](https://github.com/Gaza-children-village/excel-english-names) | `=TO_ENGLISH(cell)`: English spelling of Arabic names |
| GCV Sync | [excel-sync](https://github.com/Gaza-children-village/excel-sync) | Copies the school workbook to its Google Sheet |

See [CHANGELOG.md](CHANGELOG.md).
