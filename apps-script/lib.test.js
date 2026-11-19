// Tests for lib.js (pure functions of the GCV Sync Apps Script). Run: node --test
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const context = {};
vm.createContext(context);
vm.runInContext(fs.readFileSync(__dirname + "/lib.js", "utf8"), context);
const lib = context;

test("text that Sheets would turn into a number, date, formula or error is kept as text", () => {
  assert.equal(lib.toCell("0123456789"), "'0123456789");
  assert.equal(lib.toCell("=SUM(A1)"), "'=SUM(A1)");
  assert.equal(lib.toCell("+970"), "'+970");
  assert.equal(lib.toCell("#N/A"), "'#N/A");
  assert.equal(lib.toCell("TRUE"), "'TRUE");
  assert.equal(lib.toCell("1/2/2020"), "'1/2/2020");
});

test("ordinary text, numbers and booleans are written as they are", () => {
  assert.equal(lib.toCell("محمد"), "محمد");
  assert.equal(lib.toCell("Grade A"), "Grade A");
  assert.equal(lib.toCell(42.5), 42.5);
  assert.equal(lib.toCell(true), true);
  assert.equal(lib.toCell(null), "");
  assert.equal(lib.toCell(""), "");
});

test("dates sent as {$d: ...} become dates", () => {
  const d = lib.toCell({ $d: "2015-03-15" });
  assert.equal(Object.prototype.toString.call(d), "[object Date]");
  assert.equal(d.getFullYear(), 2015);
  assert.equal(d.getMonth(), 2);
  assert.equal(d.getDate(), 15);
  const t = lib.toCell({ $d: "2026-11-19T10:30:00" });
  assert.equal(t.getHours(), 10);
  assert.equal(t.getMinutes(), 30);
});

test("rows are padded to one width", () => {
  const out = lib.toValues([["a"], ["b", 1, "0123"], []]);
  assert.equal(out.width, 3);
  assert.deepEqual(JSON.parse(JSON.stringify(out.values)), [["a", "", ""], ["b", 1, "'0123"], ["", "", ""]]);
});

test("an empty sheet still gives one cell", () => {
  const out = lib.toValues([]);
  assert.equal(out.width, 1);
  assert.equal(out.values.length, 1);
});

test("editors: add the allowed ones, remove all others, never the owner", () => {
  const plan = lib.planEditors(["Owner@gcv.org", "old@x.org", "Sec@school.org"], ["sec@school.org", "admin@gcv.org"], "owner@gcv.org");
  assert.deepEqual(JSON.parse(JSON.stringify(plan)), { add: ["admin@gcv.org"], remove: ["old@x.org"] });
});

test("emails are read from a comma or space separated list, lower case, no duplicates", () => {
  assert.deepEqual(JSON.parse(JSON.stringify(lib.parseEmails(" A@x.org, b@y.org;a@x.org  c@z.org "))), ["a@x.org", "b@y.org", "c@z.org"]);
  assert.deepEqual(JSON.parse(JSON.stringify(lib.parseEmails(""))), []);
});

test("school codes are short letters, digits and dashes", () => {
  assert.equal(lib.isSchoolCode("S03"), true);
  assert.equal(lib.isSchoolCode("gaza-north-1"), true);
  assert.equal(lib.isSchoolCode(""), false);
  assert.equal(lib.isSchoolCode("S 03"), false);
  assert.equal(lib.isSchoolCode("x".repeat(21)), false);
});

test("a payload is checked before anything is written", () => {
  const good = { school: "S03", key: "k", workbook: "f.xlsx", sheets: [{ name: "Students", rows: [["a"]] }] };
  assert.equal(lib.checkPayload(good), "");
  assert.match(lib.checkPayload(null), /JSON/);
  assert.match(lib.checkPayload({ ...good, school: "S 03" }), /school/);
  assert.match(lib.checkPayload({ ...good, key: "" }), /key/);
  assert.match(lib.checkPayload({ ...good, sheets: [] }), /sheets/);
  assert.match(lib.checkPayload({ ...good, sheets: [{ name: "", rows: [] }] }), /name/);
  assert.match(lib.checkPayload({ ...good, sheets: [{ name: "_sync", rows: [] }] }), /_sync/);
  assert.match(lib.checkPayload({ ...good, sheets: [{ name: "A", rows: [] }, { name: "a", rows: [] }] }), /twice/);
  assert.match(lib.checkPayload({ ...good, sheets: [{ name: "A", rows: "x" }] }), /rows/);
});

test("the school name check accepts the first name and refuses another one", () => {
  assert.equal(lib.schoolNameProblem("", "مدرسة أ"), "");
  assert.equal(lib.schoolNameProblem("مدرسة أ", " مدرسة أ "), "");
  assert.equal(lib.schoolNameProblem("مدرسة أ", ""), "");
  assert.match(lib.schoolNameProblem("مدرسة أ", "مدرسة ب"), /another school/);
});

test("keys are compared in constant time", () => {
  assert.equal(lib.sameText("abc", "abc"), true);
  assert.equal(lib.sameText("abc", "abd"), false);
  assert.equal(lib.sameText("abc", "abcd"), false);
});
