/**
 * GCV Sync - web app that receives a school's Excel file from the GCV Sync
 * add-in and mirrors it into that school's Google Sheet.
 *
 * ONE standalone Apps Script project under the GCV Google account serves
 * every school. Deploy: Deploy > New deployment > Web app,
 *   Execute as: Me (the GCV account)    Who has access: Anyone
 * ("Anyone" because Excel cannot sign in to Google; every request needs the
 * school's own key, and only a SHA-256 hash of each key is stored).
 *
 * Script properties (Project Settings > Script properties):
 *   FOLDER_ID     Drive folder where the school Sheets are created
 *   ADMIN_EMAILS  admins who may edit every school Sheet, e.g. "admin@gcv.org"
 * Schools are added from the editor (Run > addSchoolFromEditor after editing
 * its arguments, or call addSchool from another function); see README.
 *
 * Every sync:
 *   - checks the key and that the file belongs to this school (school name)
 *   - replaces the tabs with the Excel sheets' values (tabs that no longer
 *     exist in Excel are removed); writes the tab "_sync" (time, file, version)
 *   - resets sharing: only ADMIN_EMAILS and the school's secretary may edit,
 *     no link sharing, editors cannot share further
 */

var PROPERTY_PREFIX = "school:";

// ---------------------------------------------------------------------------
// Web app
// ---------------------------------------------------------------------------

function doPost(e) {
  var payload;
  try {
    payload = JSON.parse(e && e.postData ? e.postData.contents : "");
  } catch (err) {
    return reply({ ok: false, error: "The request is not valid JSON." });
  }
  var problem = checkPayload(payload);
  if (problem) return reply({ ok: false, error: problem });

  var school = readSchool(payload.school);
  if (!school || !sameText(school.keyHash, hashKey(payload.key))) {
    console.warn("Refused sync for school " + payload.school + ": unknown school or wrong key.");
    return reply({ ok: false, error: "Unknown school or wrong key." });
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return reply({ ok: false, error: "Another sync is running. Try again in a minute." });
  try {
    var nameProblem = schoolNameProblem(school.schoolName, payload.schoolName);
    if (nameProblem) return reply({ ok: false, error: nameProblem });
    if (!school.schoolName && payload.schoolName) {
      school.schoolName = String(payload.schoolName).trim();
      saveSchool(payload.school, school);
    }

    var ss = SpreadsheetApp.openById(school.sheetId);
    var rows = writeSheets(ss, payload.sheets);
    writeSyncTab(ss, payload, rows);
    enforceSharing(ss, school);
    console.log("Synced " + payload.school + ": " + payload.sheets.length + " sheet(s), " + rows + " row(s).");
    return reply({ ok: true, sheets: payload.sheets.length, rows: rows, url: ss.getUrl() });
  } catch (err) {
    console.error("Sync of " + payload.school + " failed: " + err);
    return reply({ ok: false, error: "The sync failed on the server: " + err.message });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return reply({ ok: true, service: "GCV Sync" });
}

function reply(object) {
  return ContentService.createTextOutput(JSON.stringify(object)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// Writing the sheets
// ---------------------------------------------------------------------------

/** Mirrors the sent sheets (in their order) and returns the number of rows written. */
function writeSheets(ss, sheets) {
  var total = 0;
  var keep = [SYNC_TAB];
  sheets.forEach(function (sent, index) {
    var tab = ss.getSheetByName(sent.name) || ss.insertSheet(sent.name);
    var data = toValues(sent.rows);
    tab.clear();
    if (tab.getMaxRows() < data.values.length) tab.insertRowsAfter(tab.getMaxRows(), data.values.length - tab.getMaxRows());
    if (tab.getMaxColumns() < data.width) tab.insertColumnsAfter(tab.getMaxColumns(), data.width - tab.getMaxColumns());
    tab.getRange(1, 1, data.values.length, data.width).setValues(data.values);
    ss.setActiveSheet(tab);
    ss.moveActiveSheet(index + 1);
    if (sent.rightToLeft) tab.setRightToLeft(true);
    keep.push(sent.name);
    total += sent.rows.length;
  });
  ss.getSheets().forEach(function (tab) {
    if (keep.indexOf(tab.getName()) < 0) ss.deleteSheet(tab);
  });
  return total;
}

function writeSyncTab(ss, payload, rows) {
  var tab = ss.getSheetByName(SYNC_TAB) || ss.insertSheet(SYNC_TAB);
  tab.clear();
  tab.getRange(1, 1, 6, 2).setValues([
    ["Last sync", new Date()],
    ["Workbook", toCell(payload.workbook || "")],
    ["School name", toCell(payload.schoolName || "")],
    ["Sent by add-in", toCell(payload.addinVersion || "")],
    ["Sheets", payload.sheets.length],
    ["Rows", rows]
  ]);
  ss.setActiveSheet(tab);
  ss.moveActiveSheet(ss.getSheets().length);
}

// ---------------------------------------------------------------------------
// Sharing
// ---------------------------------------------------------------------------

function enforceSharing(ss, school) {
  var file = DriveApp.getFileById(ss.getId());
  var allowed = parseEmails(PropertiesService.getScriptProperties().getProperty("ADMIN_EMAILS")).concat(school.editors || []);
  var owner = file.getOwner() ? file.getOwner().getEmail() : "";
  try {
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  } catch (err) {
    console.warn("Link sharing could not be changed: " + err);
  }
  file.setShareableByEditors(false);
  file.getViewers().forEach(function (user) {
    if (allowed.indexOf(user.getEmail().toLowerCase()) < 0) file.removeViewer(user.getEmail());
  });
  var plan = planEditors(file.getEditors().map(function (u) { return u.getEmail(); }), allowed, owner);
  plan.remove.forEach(function (email) { file.removeEditor(email); });
  plan.add.forEach(function (email) { file.addEditor(email); });
}

// ---------------------------------------------------------------------------
// Schools (run by the admin in the script editor)
// ---------------------------------------------------------------------------

/**
 * Creates the school's Google Sheet in FOLDER_ID, stores its settings and
 * prints the key for the GCV Sync settings on the school's PC.
 * Edit the three values and press Run.
 */
function addSchoolFromEditor() {
  addSchool("S01", "School name", "secretary@example.org");
}

function addSchool(code, title, secretaryEmail) {
  if (!isSchoolCode(code)) throw new Error("Invalid school code: " + code);
  if (readSchool(code)) throw new Error("School " + code + " exists already; use resetKey or setEditors.");
  var folderId = PropertiesService.getScriptProperties().getProperty("FOLDER_ID");
  if (!folderId) throw new Error("Set the script property FOLDER_ID first.");
  var ss = SpreadsheetApp.create("GCV - " + code + " - " + title);
  DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(folderId));
  var key = newKey();
  var school = { sheetId: ss.getId(), keyHash: hashKey(key), editors: parseEmails(secretaryEmail), schoolName: "" };
  saveSchool(code, school);
  enforceSharing(ss, school);
  console.log("School " + code + " added. Sheet: " + ss.getUrl());
  console.log("KEY for " + code + " (shown once, enter it in GCV Sync settings): " + key);
  return key;
}

/** New key for a school (e.g. after a PC was lost); the old key stops working. */
function resetKey(code) {
  var school = mustReadSchool(code);
  var key = newKey();
  school.keyHash = hashKey(key);
  saveSchool(code, school);
  console.log("New KEY for " + code + ": " + key);
  return key;
}

/** Replaces the school's editors (secretary emails), e.g. setEditors("S01", "new@school.org"). */
function setEditors(code, emails) {
  var school = mustReadSchool(code);
  school.editors = parseEmails(emails);
  saveSchool(code, school);
  enforceSharing(SpreadsheetApp.openById(school.sheetId), school);
  console.log(code + " editors: " + school.editors.join(", "));
}

/** Forget the recorded school name (if a school was renamed). */
function clearSchoolName(code) {
  var school = mustReadSchool(code);
  school.schoolName = "";
  saveSchool(code, school);
}

function listSchools() {
  var props = PropertiesService.getScriptProperties().getProperties();
  Object.keys(props).filter(function (k) { return k.indexOf(PROPERTY_PREFIX) === 0; }).sort().forEach(function (k) {
    var s = JSON.parse(props[k]);
    console.log(k.substring(PROPERTY_PREFIX.length) + "  " + (s.schoolName || "(no sync yet)") + "  editors: " +
      (s.editors || []).join(", ") + "  sheet: https://docs.google.com/spreadsheets/d/" + s.sheetId);
  });
}

function readSchool(code) {
  var text = PropertiesService.getScriptProperties().getProperty(PROPERTY_PREFIX + code);
  return text ? JSON.parse(text) : null;
}

function mustReadSchool(code) {
  var school = readSchool(code);
  if (!school) throw new Error("Unknown school: " + code);
  return school;
}

function saveSchool(code, school) {
  PropertiesService.getScriptProperties().setProperty(PROPERTY_PREFIX + code, JSON.stringify(school));
}

function newKey() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "");
}

function hashKey(key) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(key), Utilities.Charset.UTF_8)
    .map(function (b) { return ("0" + (b & 0xff).toString(16)).slice(-2); })
    .join("");
}
