/**
 * GCV Sync - pure helpers (no Google services; tested with Node: node --test).
 * In Apps Script every file shares one global scope, so Code.js calls these
 * directly.
 */

var SYNC_TAB = "_sync";
var MAX_SCHOOL_CODE = 20;

/**
 * One value from Excel -> what setValues() must receive.
 * Text that Sheets would turn into a number, date, formula, boolean or error
 * gets a leading apostrophe, so IDs keep their leading zeros.
 * Dates arrive as {"$d": "yyyy-mm-dd" | "yyyy-mm-ddThh:nn:ss"}.
 */
function toCell(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "number" || typeof value === "boolean") return value;
  if (typeof value === "object" && typeof value.$d === "string") return parseLocalDate(value.$d);
  var text = String(value);
  if (text === "") return "";
  if (/^[=+\-@#'.0-9]/.test(text) || /^(true|false)$/i.test(text) || /^\s/.test(text)) return "'" + text;
  return text;
}

/** "2015-03-15" or "2015-03-15T10:30:00" as a local date (no time zone shift). */
function parseLocalDate(text) {
  var m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2}))?$/.exec(text);
  if (!m) return text;
  return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
}

/** Rows from Excel -> a rectangular array for setValues(). */
function toValues(rows) {
  var width = 1;
  for (var i = 0; i < rows.length; i++) width = Math.max(width, rows[i].length);
  var values = [];
  for (var r = 0; r < Math.max(rows.length, 1); r++) {
    var source = rows[r] || [];
    var row = [];
    for (var c = 0; c < width; c++) row.push(c < source.length ? toCell(source[c]) : "");
    values.push(row);
  }
  return { values: values, width: width };
}

/** Which editors to add and remove so that exactly `allowed` (plus the owner) can edit. */
function planEditors(current, allowed, owner) {
  var cur = current.map(lower);
  var want = allowed.map(lower);
  var own = lower(owner || "");
  return {
    add: want.filter(function (e) { return cur.indexOf(e) < 0 && e !== own; }),
    remove: cur.filter(function (e) { return want.indexOf(e) < 0 && e !== own; })
  };
}

function parseEmails(text) {
  var seen = [];
  String(text || "").split(/[\s,;]+/).forEach(function (part) {
    var e = lower(part);
    if (e && seen.indexOf(e) < 0) seen.push(e);
  });
  return seen;
}

function isSchoolCode(code) {
  return typeof code === "string" && code.length > 0 && code.length <= MAX_SCHOOL_CODE && /^[A-Za-z0-9-]+$/.test(code);
}

/** "" when the payload can be written, otherwise the reason. */
function checkPayload(p) {
  if (!p || typeof p !== "object") return "The request is not a JSON object.";
  if (!isSchoolCode(p.school)) return "Invalid school code.";
  if (typeof p.key !== "string" || p.key === "") return "The key is missing.";
  if (!Array.isArray(p.sheets) || p.sheets.length === 0) return "No sheets were sent.";
  var names = [];
  for (var i = 0; i < p.sheets.length; i++) {
    var s = p.sheets[i];
    if (!s || typeof s.name !== "string" || s.name === "") return "A sheet has no name.";
    if (s.name.toLowerCase() === SYNC_TAB) return "A sheet may not be called " + SYNC_TAB + ".";
    if (names.indexOf(s.name.toLowerCase()) >= 0) return "Sheet " + s.name + " was sent twice.";
    names.push(s.name.toLowerCase());
    if (!Array.isArray(s.rows)) return "Sheet " + s.name + " has no rows.";
  }
  return "";
}

/** The first sync records the workbook's school name; a file of another school is refused. */
function schoolNameProblem(recorded, sent) {
  var a = String(recorded || "").trim();
  var b = String(sent || "").trim();
  if (a === "" || b === "" || a === b) return "";
  return "This file belongs to another school (" + b + "), not to " + a + ".";
}

/** Compares two strings without stopping at the first difference. */
function sameText(a, b) {
  a = String(a);
  b = String(b);
  var diff = a.length ^ b.length;
  for (var i = 0; i < Math.max(a.length, b.length); i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

function lower(text) {
  return String(text || "").trim().toLowerCase();
}
