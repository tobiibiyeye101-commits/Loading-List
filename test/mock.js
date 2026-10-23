// A small in-memory stand-in for the parts of SpreadsheetApp the script uses.
// Formulas are stored as text and are not calculated.

function colToNumber(letters) {
  return letters.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

function parseA1(a1) {
  const [start, end] = a1.split(":");
  const cell = ref => {
    const m = /^\$?([A-Z]+)\$?(\d*)$/.exec(ref);
    return { col: colToNumber(m[1]), row: m[2] ? Number(m[2]) : null };
  };
  const s = cell(start);
  const e = end ? cell(end) : s;
  return { row: s.row, col: s.col, numRows: e.row - s.row + 1, numCols: e.col - s.col + 1 };
}

class Range {
  constructor(sheet, row, col, numRows, numCols) {
    Object.assign(this, { sheet, row, col, numRows: numRows || 1, numCols: numCols || 1 });
  }
  getSheet() { return this.sheet; }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  getNumRows() { return this.numRows; }
  getNumColumns() { return this.numCols; }
  getA1Notation() { return `R${this.row}C${this.col}`; }

  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const row = [];
      for (let c = 0; c < this.numCols; c++) row.push(this.sheet.get(this.row + r, this.col + c));
      out.push(row);
    }
    return out;
  }
  getValue() { return this.sheet.get(this.row, this.col); }

  setValues(values) {
    if (values.length !== this.numRows || values.some(r => r.length !== this.numCols)) {
      throw new Error(`setValues size mismatch on ${this.sheet.name} R${this.row}C${this.col}`);
    }
    values.forEach((r, i) => r.forEach((v, j) => this.sheet.set(this.row + i, this.col + j, v)));
    return this;
  }
  setFormulas(values) { return this.setValues(values); }
  setValue(v) { this.sheet.set(this.row, this.col, v); return this; }
  setFormula(v) { return this.setValue(v); }
  clearContent() {
    for (let r = 0; r < this.numRows; r++)
      for (let c = 0; c < this.numCols; c++) this.sheet.set(this.row + r, this.col + c, "");
    return this;
  }
  setBackgrounds(colours) {
    colours.forEach((r, i) => r.forEach((v, j) => { this.sheet.bg[`${this.row + i},${this.col + j}`] = v; }));
    return this;
  }
  setBackground(colour) {
    for (let r = 0; r < this.numRows; r++)
      for (let c = 0; c < this.numCols; c++) this.sheet.bg[`${this.row + r},${this.col + c}`] = colour;
    return this;
  }
  insertCheckboxes() { return this.setValue(false); }
  protect() { return new Protection(); }
}
["setFontColor", "setFontWeight", "setBorder", "setNumberFormat", "clearDataValidations"].forEach(name => {
  Range.prototype[name] = function () { return this; };
});

class Protection {
  setDescription() { return this; }
  setWarningOnly() { return this; }
}

class Sheet {
  constructor(name, maxRows) {
    this.name = name;
    this.rows = [];
    this.bg = {};
    this.maxRows = maxRows || 200;
  }
  getName() { return this.name; }
  get(row, col) { return (this.rows[row - 1] || [])[col - 1] ?? ""; }
  set(row, col, v) {
    if (row > this.maxRows) throw new Error(`Row ${row} is past the end of ${this.name}`);
    while (this.rows.length < row) this.rows.push([]);
    this.rows[row - 1][col - 1] = v;
  }
  getRange(a, b, c, d) {
    if (typeof a === "string") {
      const p = parseA1(a);
      return new Range(this, p.row, p.col, p.numRows, p.numCols);
    }
    return new Range(this, a, b, c, d);
  }
  getMaxRows() { return this.maxRows; }
  getLastRow() {
    for (let r = this.rows.length; r >= 1; r--) {
      if ((this.rows[r - 1] || []).some(v => v !== "" && v !== undefined && v !== null)) return r;
    }
    return 0;
  }
  insertRowsBefore(row, n) {
    while (this.rows.length < row - 1) this.rows.push([]);
    this.rows.splice(row - 1, 0, ...Array.from({ length: n }, () => []));
    this.maxRows += n;
  }
  insertRowsAfter(row, n) { this.insertRowsBefore(row + 1, n); }
  deleteRows(row, n) {
    this.rows.splice(row - 1, n);
    this.maxRows -= n;
  }
  getProtections() { return []; }
  protect() { return new Protection(); }
}

class Spreadsheet {
  constructor() {
    this.sheets = {};
    ["Owner_List", "Loading_Board", "Leftover"].forEach(n => { this.sheets[n] = new Sheet(n); });
  }
  getSheetByName(n) { return this.sheets[n] || null; }
  getSheets() { return Object.values(this.sheets); }
  getId() { return "test-sheet"; }
  getSpreadsheetTimeZone() { return "Africa/Lagos"; }
}

function formatDate(date, timeZone, pattern) {
  const parts = {};
  new Intl.DateTimeFormat("en-GB", {
    timeZone, year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23", weekday: "short"
  }).formatToParts(date).forEach(p => { parts[p.type] = p.value; });
  const monthName = new Intl.DateTimeFormat("en-US", { timeZone, month: "long" }).format(date);
  const dayNumber = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"].indexOf(parts.weekday) + 1;

  return pattern
    .replace("yyyy", parts.year).replace("MMMM", monthName).replace("MM", parts.month)
    .replace("dd", parts.day).replace("HH", parts.hour).replace("mm", parts.minute)
    .replace("ss", parts.second).replace("EEE", parts.weekday).replace(/^u$/, String(dayNumber));
}

function parseDate(text, timeZone) {
  // Midnight of yyyy-MM-dd in the given time zone.
  const [y, m, d] = text.split("-").map(Number);
  const guess = new Date(Date.UTC(y, m - 1, d));
  const shown = formatDate(guess, timeZone, "yyyy-MM-dd HH:mm");
  const [sd, st] = shown.split(" ");
  const [hh, mm] = st.split(":").map(Number);
  let offsetMinutes = hh * 60 + mm;
  if (sd !== text) offsetMinutes -= 24 * 60;
  return new Date(guess.getTime() - offsetMinutes * 60000);
}

function makeGlobals(ss, alerts) {
  return {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getUi: () => ({
        alert: (a) => { alerts.push(a); return "YES"; },
        ButtonSet: { YES_NO: "YES_NO" },
        Button: { YES: "YES" }
      }),
      BorderStyle: { SOLID_MEDIUM: "SOLID_MEDIUM" },
      ProtectionType: { SHEET: "SHEET", RANGE: "RANGE" }
    },
    Utilities: { formatDate, parseDate },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    ScriptApp: { getProjectTriggers: () => [] },
    Session: { getEffectiveUser: () => ({}) },
    console
  };
}

module.exports = { Spreadsheet, makeGlobals, formatDate };
