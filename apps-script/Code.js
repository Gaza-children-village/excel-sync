/**
 * GCV Sync - web app that receives a school's Excel file from the GCV Sync
 * add-in and mirrors it into that school's Google Sheet.
 *
 * ONE standalone Apps Script project under the GCV Google account serves
 * every school. Deploy: Deploy > New deployment > Web app,
 *   Execute as: Me (the GCV account)    Who has access: Anyone
 * ("Anyone" because Excel cannot sign in to Google). The /exec link is built
 * into the add-in: after code changes use Deploy > Manage deployments > Edit
 * > Version: New version, so the link stays the same.
 *
 * Schools link themselves: the add-in makes a random key on the PC, and the
 * first sync of a school code (GCV01 ... GCV99) creates the school's Sheet
 * and links the school to that key (only its SHA-256 hash is stored). Other
 * PCs are then refused for that school.
 *
 * Script properties (Project Settings > Script properties):
 *   FOLDER_ID          Drive folder where the school Sheets are created
 *   REGISTRATION_OPEN  "true" while schools set up their PCs; "false" after.
 *                      New schools are refused while it is off.
 *   ADMIN_EMAILS       optional: accounts that may edit every school Sheet
 *
 * Every sync:
 *   - checks the school and its PC, and that the file belongs to this school
 *     (school name recorded at the first sync)
 *   - replaces the tabs with the Excel sheets' values (tabs that no longer
 *     exist in Excel are removed); writes the tab "_sync" (time, file, version)
 *   - resets sharing: only ADMIN_EMAILS and the school's editors (setEditors)
 *     may edit, no link sharing, editors cannot share further
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
    return refuse("bad-request", "The request is not valid JSON.");
  }
  var problem = checkPayload(payload);
  if (problem) return refuse("bad-request", problem);

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return refuse("busy", "Another sync is running. Try again in a minute.");
  try {
    var code = payload.school;
    var school = readSchool(code);
    var keyHash = hashKey(payload.key);
    var open = isOn(PropertiesService.getScriptProperties().getProperty("REGISTRATION_OPEN"));
    var access = accessDecision(school, keyHash, open);
    if (access === "closed") {
      console.warn("Refused " + code + ": not registered and registration is closed.");
      return refuse("registration-closed", "School " + code + " is not registered and registration is closed.");
    }
    if (access === "taken") {
      console.warn("Refused " + code + ": the school is linked to another computer.");
      return refuse("school-taken", "School " + code + " is linked to another computer.");
    }
    var nameProblem = schoolNameProblem(school && school.schoolName, payload.schoolName);
    if (nameProblem) return refuse("wrong-school", nameProblem);

    if (access === "register") school = createSchool(code, payload.schoolName, keyHash);
    if (access === "claim") school.keyHash = keyHash;
    if (!school.schoolName && payload.schoolName) school.schoolName = String(payload.schoolName).trim();
    saveSchool(code, school);
    if (access !== "ok") console.log("School " + code + " linked to a computer (" + access + ").");

    var ss = SpreadsheetApp.openById(school.sheetId);
    var rows = writeSheets(ss, payload.sheets);
    writeSyncTab(ss, payload, rows);
    enforceSharing(ss, school);
    console.log("Synced " + code + ": " + payload.sheets.length + " sheet(s), " + rows + " row(s).");
    return reply({ ok: true, registered: access !== "ok", sheets: payload.sheets.length, rows: rows, url: ss.getUrl() });
  } catch (err) {
    console.error("Sync of " + payload.school + " failed: " + err);
    return refuse("server-error", "The sync failed on the server: " + err.message);
  } finally {
    lock.releaseLock();
  }
}

function refuse(code, message) {
  return reply({ ok: false, code: code, error: message });
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
// Schools
// ---------------------------------------------------------------------------

/** First sync of a school: its Sheet in FOLDER_ID, linked to the sending PC. */
function createSchool(code, schoolName, keyHash) {
  var folderId = PropertiesService.getScriptProperties().getProperty("FOLDER_ID");
  if (!folderId) throw new Error("The script property FOLDER_ID is not set.");
  var name = String(schoolName || "").trim();
  var ss = SpreadsheetApp.create("GCV Sync - " + code + (name ? " - " + name : ""));
  DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(folderId));
  console.log("School " + code + " created. Sheet: " + ss.getUrl());
  return { sheetId: ss.getId(), keyHash: keyHash, editors: [], schoolName: name };
}

// The functions below are run by the admin in the script editor. The Run
// button cannot pass values: edit the code in the ...FromEditor function,
// save, choose it in the list and press Run.

/**
 * Unlinks a school from its PC (new or replaced PC, or a secretary typed the
 * wrong school number). The next sync for this school links the PC that sends
 * it, even with registration closed. The recorded school name is forgotten
 * too; the Sheet and its data stay until that sync replaces them.
 */
function unlinkSchoolFromEditor() {
  unlinkSchool("GCV01");
}

function unlinkSchool(code) {
  var school = mustReadSchool(code);
  school.keyHash = "";
  school.schoolName = "";
  saveSchool(code, school);
  console.log(code + " unlinked. The next sync for " + code + " links its computer.");
}

/** Optional: accounts that may edit this school's Sheet, e.g. setEditors("GCV05", "sec@gmail.com"). */
function setEditors(code, emails) {
  var school = mustReadSchool(code);
  school.editors = parseEmails(emails);
  saveSchool(code, school);
  enforceSharing(SpreadsheetApp.openById(school.sheetId), school);
  console.log(code + " editors: " + (school.editors.join(", ") || "(none)"));
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
    console.log(k.substring(PROPERTY_PREFIX.length) + "  " + (s.schoolName || "(no name yet)") +
      (s.keyHash ? "" : "  (UNLINKED)") + "  editors: " + ((s.editors || []).join(", ") || "-") +
      "  sheet: https://docs.google.com/spreadsheets/d/" + s.sheetId);
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

function hashKey(key) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(key), Utilities.Charset.UTF_8)
    .map(function (b) { return ("0" + (b & 0xff).toString(16)).slice(-2); })
    .join("");
}
