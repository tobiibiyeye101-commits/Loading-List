// Run with: node test/run.js
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");
const { Spreadsheet, makeGlobals } = require("./mock");

const code = fs.readFileSync(path.join(__dirname, "..", "Code.gs"), "utf8");

function load() {
  const ss = new Spreadsheet();
  const alerts = [];
  const context = vm.createContext(makeGlobals(ss, alerts));
  vm.runInContext(code + "\n;this.__api = { START_ROW, STATUS, onEdit, reconcileBoards, syncAllStatuses, resetAllStatusesAndLeftovers, collectCarryForward, getNextShiftInfo, rebuildLoadingBoard };", context);
  return { ss, alerts, api: context.__api };
}

// [id, name, 12 quantities..., status, timestamp, source]
function order(id, name, qty, status, source) {
  const q = new Array(12).fill("");
  Object.entries(qty).forEach(([i, v]) => { q[i] = v; });
  return [id, name, ...q, status || "", "", source || "Manual"];
}

function fillOwner(ss, rows) {
  const sheet = ss.getSheetByName("Owner_List");
  rows.forEach((r, i) => sheet.getRange(8 + i, 1, 1, 17).setValues([r]));
}

function boardRows(ss) {
  return ss.getSheetByName("Loading_Board").getRange(8, 1, 99, 17).getValues().filter(r => r[1] !== "");
}

function edit(api, ss, sheetName, row, col, value, numRows) {
  const sheet = ss.getSheetByName(sheetName);
  if (value !== undefined) sheet.getRange(row, col).setValue(value);
  api.onEdit({ source: ss, range: sheet.getRange(row, col, numRows || 1, 1) });
}

function leftoverRows(ss) {
  const sheet = ss.getSheetByName("Leftover");
  const last = sheet.getLastRow();
  return last < 10 ? [] : sheet.getRange(10, 1, last - 9, 16).getValues();
}

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

test("orders with the same customer name are added together on the board", () => {
  const { ss, api } = load();
  fillOwner(ss, [
    order(1, "Alpha", { 8: 78 }, "Not Started", "Leftover"),
    order(2, "Bravo", { 1: 1300, 2: 1500 }, "Not Started", "Leftover"),
    order(3, "alpha ", { 8: 5, 0: 10 }, "Not Started", "Leftover"),
    order(4, "Charlie", { 2: -36, 6: 36 }, "Not Started", "Leftover"),
    order(5, "Bravo", { 2: 100 }, "", "Manual")
  ]);
  api.reconcileBoards();

  const board = boardRows(ss);
  assert.deepStrictEqual(board.map(r => r[1]), ["Alpha", "Bravo", "Charlie"]);
  assert.deepStrictEqual(board[0].slice(0, 14), [1, "Alpha", 10, "", "", "", "", "", "", "", 83, "", "", ""]);
  assert.strictEqual(board[1][4], 1600);
  assert.strictEqual(board[2][4], -36, "negatives are kept");
  assert.strictEqual(board[1][16], "Leftover + Manual");
  assert.strictEqual(board[0][14], "Not started", "status matches the dropdown spelling");

  const owner = ss.getSheetByName("Owner_List");
  assert.strictEqual(owner.getRange(12, 15).getValue(), "Not started", "blank owner status filled in");

  // Totals cover every data row
  assert.strictEqual(owner.getRange("C7").getValue(), "=SUM(C8:C106)");
  assert.match(ss.getSheetByName("Loading_Board").getRange("C7").getValue(), /^=SUMIFS\(C8:C106,\$O\$8:\$O\$106,/);
});

test("status set on the board applies to every order for that customer", () => {
  const { ss, api } = load();
  fillOwner(ss, [
    order(1, "Alpha", { 0: 10 }),
    order(2, "Bravo", { 0: 20 }),
    order(3, "Alpha", { 0: 5 })
  ]);
  api.reconcileBoards();

  edit(api, ss, "Loading_Board", 8, 15, "In progress");
  const owner = ss.getSheetByName("Owner_List");
  assert.strictEqual(owner.getRange(8, 15).getValue(), "In progress");
  assert.strictEqual(owner.getRange(10, 15).getValue(), "In progress");
  assert.strictEqual(owner.getRange(9, 15).getValue(), "Not started");
  assert.notStrictEqual(owner.getRange(10, 16).getValue(), "", "timestamp copied");
  assert.strictEqual(owner.bg["10,1"], "#fff2cc", "lower-case 'In progress' still turns yellow");
});

test("Leftover blocks follow the merged board and keep Taken amounts", () => {
  const { ss, api } = load();
  fillOwner(ss, [
    order(1, "Alpha", { 0: 10 }),
    order(2, "Bravo", { 0: 20 }),
    order(3, "Alpha", { 0: 5 })
  ]);
  api.reconcileBoards();

  edit(api, ss, "Loading_Board", 9, 15, "Leftover");   // Bravo
  edit(api, ss, "Loading_Board", 8, 15, "Leftover");   // Alpha
  let rows = leftoverRows(ss);
  assert.deepStrictEqual(rows.map(r => [r[1], r[2]]), [
    ["Alpha", "Order"], ["", "Taken"], ["", "Leftover"],
    ["Bravo", "Order"], ["", "Taken"], ["", "Leftover"]
  ], "blocks are in board order");
  assert.strictEqual(rows[0][3], 15, "Order row holds the merged quantity");
  assert.strictEqual(rows[2][3], 15, "Leftover = Order minus Taken (Taken is blank)");
  assert.strictEqual(rows[2][15], false, "carry-forward checkbox");

  // Type a Taken amount, then add another Alpha order
  ss.getSheetByName("Leftover").getRange(11, 4).setValue(12);
  fillOwner(ss, [order(1, "Alpha", { 0: 10 }), order(2, "Bravo", { 0: 20 }), order(3, "Alpha", { 0: 5 }), ["", "Alpha", 7, ...new Array(14).fill("")]]);
  edit(api, ss, "Owner_List", 11, 2);
  assert.strictEqual(ss.getSheetByName("Owner_List").getRange(11, 1).getValue(), 4, "new order gets the next ID");
  assert.strictEqual(ss.getSheetByName("Owner_List").getRange(11, 17).getValue(), "Manual");
  rows = leftoverRows(ss);
  assert.strictEqual(rows[0][3], 22);
  assert.strictEqual(rows[1][3], 12, "Taken amount kept");
  assert.strictEqual(ss.getSheetByName("Owner_List").getRange(11, 15).getValue(), "Leftover", "new order picks up the customer's status");

  // Marking Alpha Done removes only Alpha's block
  edit(api, ss, "Loading_Board", 8, 15, "Done");
  rows = leftoverRows(ss);
  assert.deepStrictEqual(rows.filter(r => r[2] === "Order").map(r => r[1]), ["Bravo"]);
});

test("deleting an owner row shifts the board without mixing up statuses", () => {
  const { ss, api } = load();
  fillOwner(ss, [order(1, "Alpha", { 0: 1 }), order(2, "Bravo", { 0: 2 }), order(3, "Charlie", { 0: 3 })]);
  api.reconcileBoards();
  edit(api, ss, "Loading_Board", 10, 15, "Done");      // Charlie

  const owner = ss.getSheetByName("Owner_List");
  owner.deleteRows(9, 1);                               // remove Bravo
  owner.maxRows += 1;
  api.rebuildLoadingBoard(ss);

  const board = boardRows(ss);
  assert.deepStrictEqual(board.map(r => [r[1], r[14]]), [["Alpha", "Not started"], ["Charlie", "Done"]]);
});

test("Sync All turns Not started customers into Leftovers", () => {
  const { ss, api } = load();
  fillOwner(ss, [order(1, "Alpha", { 0: 1 }, "Done"), order(2, "Bravo", { 0: 2 }, "Not Started"), order(3, "Bravo", { 1: 4 })]);
  api.syncAllStatuses();

  const board = boardRows(ss);
  assert.deepStrictEqual(board.map(r => r[14]), ["Done", "Leftover"]);
  assert.deepStrictEqual(leftoverRows(ss).filter(r => r[2] === "Order").map(r => [r[1], r[3], r[4]]), [["Bravo", 2, 4]]);
});

test("Reset clears statuses and the Leftover tab", () => {
  const { ss, api } = load();
  fillOwner(ss, [order(1, "Alpha", { 0: 1 }), order(2, "Bravo", { 0: 2 })]);
  api.syncAllStatuses();
  api.resetAllStatusesAndLeftovers();

  assert.deepStrictEqual(boardRows(ss).map(r => r[14]), ["Not started", "Not started"]);
  assert.strictEqual(leftoverRows(ss).length, 0);
});

test("carry-forward keeps negatives and skips ticked or empty leftovers", () => {
  const { api } = load();
  const blank = new Array(16).fill("");
  const row = (b, c, qty, skip) => { const r = blank.slice(); r[1] = b; r[2] = c; qty.forEach((v, i) => { r[3 + i] = v; }); r[15] = skip; return r; };

  const result = JSON.parse(JSON.stringify(api.collectCarryForward([
    row("Alpha", "Order", [10, 5]), row("", "Taken", [4]), row("", "Leftover", [6, 5], false),
    row("Bravo", "Order", [3]), row("", "Taken", [3]), row("", "Leftover", [0], false),
    row("Charlie", "Order", [1]), row("", "Taken", [3]), row("", "Leftover", [-2], false),
    row("Delta", "Order", [9]), row("", "Taken", []), row("", "Leftover", [9], true)
  ])));

  assert.deepStrictEqual(result.map(r => [r[0], r[1], r[2], r[3]]), [[1, "Alpha", 6, 5], [2, "Charlie", -2, ""]]);
  assert.strictEqual(result[0][14], "Not started");
  assert.strictEqual(result[0][16], "Leftover");
});

test("next shift dates", () => {
  const { api } = load();
  const tz = "Africa/Lagos";
  const lagosMidnight = d => new Date(`${d}T00:00:00+01:00`);

  let next = api.getNextShiftInfo(lagosMidnight("2026-10-23"), "Morning", tz);
  assert.deepStrictEqual([next.dayShort, next.dateText, next.shift], ["Fri", "2026-10-23", "Night"]);

  next = api.getNextShiftInfo(lagosMidnight("2026-10-23"), "Night", tz);
  assert.deepStrictEqual([next.dayShort, next.dateText, next.shift, next.monthName, next.year], ["Sat", "2026-10-24", "Morning", "October", "2026"]);
  assert.strictEqual(next.date.getTime(), lagosMidnight("2026-10-24").getTime());

  next = api.getNextShiftInfo(lagosMidnight("2026-10-24"), "Night", tz);
  assert.deepStrictEqual([next.dayShort, next.dateText, next.shift], ["Sun", "2026-10-25", "Night"], "Sunday Morning is skipped");
});

// ---- Generate Next Shift Sheet ----
function fakeDrive(existingNames) {
  const created = [];
  const iter = items => ({ hasNext: () => items.length > 0, next: () => items.shift() });
  const copies = {};
  return {
    created,
    copies,
    DriveApp: {
      getFoldersByName: () => iter([]),
      createFolder: name => {
        const folder = { getFoldersByName: () => iter([]), createFolder: n => ({ name: n, getFoldersByName: () => iter([]), createFolder() {} }), name, getFilesByName: () => iter([]) };
        created.push(name);
        return folder;
      },
      getFilesByName: name => iter(existingNames.includes(name) ? [{ getName: () => name }] : []),
      getFileById: () => ({
        makeCopy: (name, folder) => {
          const id = `copy-${Object.keys(copies).length + 1}`;
          copies[id] = { name, folder, ss: null };
          return { getId: () => id, getUrl: () => `https://example.test/${id}` };
        }
      })
    }
  };
}

function cloneSpreadsheet(src) {
  const dst = new Spreadsheet();
  Object.keys(src.sheets).forEach(name => {
    const from = src.sheets[name];
    const to = dst.sheets[name];
    to.rows = JSON.parse(JSON.stringify(from.rows));
    to.maxRows = from.maxRows;
    to.bg = { ...from.bg };
  });
  return dst;
}

function loadForGenerate(answers) {
  const ss = new Spreadsheet();
  const alerts = [];
  const drive = fakeDrive(answers.existing || []);
  const g = makeGlobals(ss, alerts);
  g.SpreadsheetApp.getUi = () => ({
    alert: (...args) => { alerts.push(args[0]); return answers.button || "YES"; },
    ButtonSet: { YES_NO: "YES_NO" },
    Button: { YES: "YES", NO: "NO" }
  });
  g.SpreadsheetApp.openById = id => drive.copies[id].ss;
  g.DriveApp = drive.DriveApp;
  g.MailApp = { sendEmail() {} };
  const ctx = vm.createContext(g);
  vm.runInContext(code + "\n;this.api={processLeftovers,reconcileBoards,onEdit}", ctx);
  // Dates must come from the script's own realm, or `instanceof Date` fails
  const setShift = () => {
    const o = ss.getSheetByName("Owner_List");
    o.getRange("A1").setValue(vm.runInContext("new Date('2026-10-23T00:00:00+01:00')", ctx));
    o.getRange("B1").setValue("Morning");
  };
  // Make copies usable: the copy becomes a clone of the source when made
  const origCopy = drive.DriveApp.getFileById;
  drive.DriveApp.getFileById = id => {
    const f = origCopy(id);
    return {
      makeCopy: (name, folder) => {
        const copyFile = f.makeCopy(name, folder);
        drive.copies[copyFile.getId()].ss = cloneSpreadsheet(ss);
        return copyFile;
      }
    };
  };
  return { ss, alerts, drive, api: ctx.api, setShift };
}

test("Generate marks Not started customers as Leftover before carrying forward", () => {
  const { ss, alerts, drive, api, setShift } = loadForGenerate({});
  fillOwner(ss, [order(1, "Alpha", { 0: 10 }), order(2, "Bravo", { 0: 20 })]);
  api.reconcileBoards();
  edit(api, ss, "Loading_Board", 8, 15, "Done");        // Alpha is done
  // Bravo stays Not started; the owner date/shift are set
  setShift();

  api.processLeftovers();

  const copy = Object.values(drive.copies)[0].ss;
  const next = copy.getSheetByName("Owner_List").getRange(8, 1, 1, 17).getValues()[0];
  assert.strictEqual(next[1], "Bravo", "Not started customer carried forward");
  assert.strictEqual(next[2], 20);
  assert.strictEqual(ss.getSheetByName("Owner_List").getRange(9, 15).getValue(), "Leftover", "source file marked Leftover");
  assert.match(alerts[alerts.length - 1], /1 'Not started' customer\(s\) were marked Leftover/);
});

test("Generate asks before creating a second file with the same name", () => {
  const name = "Loading List - Fri - 2026-10-23 - Night";
  const cancelled = loadForGenerate({ existing: [name], button: "NO" });
  fillOwner(cancelled.ss, [order(1, "Alpha", { 0: 10 })]);
  cancelled.setShift();
  cancelled.api.processLeftovers();
  assert.strictEqual(cancelled.alerts.length, 2, "asked, then said cancelled");
  assert.match(cancelled.alerts[1], /Cancelled/);
  assert.strictEqual(Object.keys(cancelled.drive.copies).length, 0, "no copy made when cancelled");
  assert.strictEqual(cancelled.ss.getSheetByName("Owner_List").getRange(8, 15).getValue(), "", "source untouched on cancel");

  const confirmed = loadForGenerate({ existing: [name], button: "YES" });
  fillOwner(confirmed.ss, [order(1, "Alpha", { 0: 10 })]);
  confirmed.setShift();
  confirmed.api.processLeftovers();
  assert.strictEqual(Object.keys(confirmed.drive.copies).length, 1, "copy made when confirmed");
});

let failed = 0;
tests.forEach(t => {
  try {
    t.fn();
    console.log(`ok   ${t.name}`);
  } catch (err) {
    failed++;
    console.log(`FAIL ${t.name}\n     ${err.stack.split("\n").slice(0, 3).join("\n     ")}`);
  }
});
console.log(`\n${tests.length - failed}/${tests.length} passed`);
process.exit(failed ? 1 : 0);
