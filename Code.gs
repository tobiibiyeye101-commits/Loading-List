// @ts-nocheck
// Bakery loading list.
//   Owner_List    - orders are typed in here (one row per order).
//   Loading_Board - one row per customer: orders with the same customer name
//                   are added together. Loaders set the Status here.
//   Leftover      - an Order / Taken / Leftover block for every customer
//                   marked Leftover. Unticked leftovers carry into the next
//                   shift's file.

const START_ROW = 8;              // Owner_List and Loading_Board data start row
const END_ROW = 106;              // Owner_List and Loading_Board data end row
const NUM_ROWS = END_ROW - START_ROW + 1;
const TOTALS_ROW = 7;             // Owner_List and Loading_Board totals row

const LEFTOVER_START_ROW = 10;    // Leftover data starts row 10
const LEFTOVER_TOTAL_ORDER_ROW = 7;
const LEFTOVER_TOTAL_TAKEN_ROW = 8;
const LEFTOVER_TOTAL_LEFTOVER_ROW = 9;
const LEFTOVER_COLS = 15;         // A:O = Order ID, Customer Name, Row Type, bread quantities
const SKIP_COL = 16;              // P = "do not carry forward" checkbox
const LEFTOVER_SPARE_ROWS = 100;  // blank rows kept below the data when the tab is cleared

const ORDER_COLS = 14;            // A:N = Order ID, Customer Name, bread quantities
const QTY_COUNT = 12;             // C:N = bread quantities
const STATUS_COL = 15;            // O = Status
const TIMESTAMP_COL = 16;         // P = Status Changed At
const SOURCE_COL = 17;            // Q = Source

const TIME_ZONE = "Africa/Lagos";
const ROOT_FOLDER_NAME = "Loading List";

// Spelled exactly as in the Status dropdown. Comparisons ignore case, so
// "Not Started" and "Not started" are the same status.
const STATUS = {
  NOT_STARTED: "Not started",
  IN_PROGRESS: "In progress",
  RESCHEDULED: "Rescheduled",
  DONE: "Done",
  CANCELLED: "Cancelled",
  LEFTOVER: "Leftover"
};

const STATUS_COLOURS = {
  "done": "#d9ead3",        // Green
  "in progress": "#fff2cc", // Yellow
  "rescheduled": "#cfe2f3", // Blue
  "cancelled": "#f4cccc",   // Red
  "leftover": "#eadcf8"     // Purple
};

// ==========================================
// MENU AND TRIGGERS
// ==========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🍞 Bakery Tools')
    .addItem('Sync All Statuses & Colors', 'syncAllStatuses')
    .addItem('🔄 Auto-Correct Missing Orders', 'reconcileBoards')
    .addSeparator()
    .addItem('⚠️ Reset Board & Leftovers', 'resetAllStatusesAndLeftovers')
    .addSeparator()
    .addItem('Generate Next Shift Sheet', 'processLeftovers')
    .addToUi();

  // Inserting, deleting or sorting rows doesn't fire onEdit, so bring the
  // board up to date whenever the file is opened.
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    withLock(() => {
      setupTotalsRows(ss);
      const customers = rebuildLoadingBoard(ss);
      reconcileLeftoverBlocks(ss, customers);
    });
  } catch (err) {
    console.error(err);
  }
}

// Simple trigger: Google runs this on every edit, in every copy of the file,
// as the person editing. Nothing needs installing.
function onEdit(e) {
  if (!e || !e.range) return;

  const ss = e.source;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const startRow = e.range.getRow();
  const startCol = e.range.getColumn();
  const endRow = startRow + e.range.getNumRows() - 1;
  const endCol = startCol + e.range.getNumColumns() - 1;

  if (endRow < START_ROW || startRow > END_ROW) return;
  if (sheetName !== "Owner_List" && sheetName !== "Loading_Board") return;

  withLock(() => {
    if (sheetName === "Owner_List") {
      handleOwnerEdit(ss, sheet, startRow, endRow, startCol, endCol);
    } else {
      handleBoardEdit(ss, sheet, startRow, endRow, startCol, endCol);
    }
  });
}

// ==========================================
// EDIT HANDLERS
// ==========================================

function handleOwnerEdit(ss, sheet, startRow, endRow, startCol, endCol) {
  const first = Math.max(startRow, START_ROW);
  const last = Math.min(endRow, END_ROW);
  const data = sheet.getRange(START_ROW, 1, NUM_ROWS, SOURCE_COL).getValues();

  let highestId = 0;
  data.forEach(row => {
    if (!isBlank(row[0]) && !isNaN(row[0])) {
      highestId = Math.max(highestId, Number(row[0]));
    }
  });

  const statusEdited = startCol <= STATUS_COL && STATUS_COL <= endCol;
  const stamp = nowStamp();
  const overrides = {};

  for (let row = first; row <= last; row++) {
    const rowData = data[row - START_ROW];
    const key = customerKey(rowData[1]);
    if (!key) continue;

    // Generate ID for new customers
    if (isBlank(rowData[0])) {
      highestId++;
      sheet.getRange(row, 1).setValue(highestId);
      if (isBlank(rowData[SOURCE_COL - 1])) {
        sheet.getRange(row, SOURCE_COL).setValue("Manual");
      }
    }

    // A status typed on Owner_List applies to every order for that customer
    if (statusEdited) {
      overrides[key] = {
        status: canonicalStatus(rowData[STATUS_COL - 1]) || STATUS.NOT_STARTED,
        timestamp: stamp
      };
    }
  }

  const customers = rebuildLoadingBoard(ss, overrides);
  reconcileLeftoverBlocks(ss, customers);
}

function handleBoardEdit(ss, sheet, startRow, endRow, startCol, endCol) {
  // Only run if the Status column was edited
  if (STATUS_COL < startCol || STATUS_COL > endCol) return;

  const first = Math.max(startRow, START_ROW);
  const last = Math.min(endRow, END_ROW);
  const rows = sheet.getRange(first, 1, last - first + 1, STATUS_COL).getValues();
  const stamp = nowStamp();
  const overrides = {};

  rows.forEach(row => {
    const key = customerKey(row[1]);
    if (key) {
      overrides[key] = {
        status: canonicalStatus(row[STATUS_COL - 1]) || STATUS.NOT_STARTED,
        timestamp: stamp
      };
    }
  });

  // Rebuild even without a status change, so a cleared board row comes back
  const customers = rebuildLoadingBoard(ss, overrides);
  reconcileLeftoverBlocks(ss, customers);
}

// ==========================================
// LOADING BOARD
// ==========================================

/**
 * Merges Owner_List rows into one Loading_Board row per customer.
 * Names match ignoring case and extra spaces. Customers keep the order they
 * first appear in on Owner_List, and take the first order's ID and spelling.
 * Quantities are added up (negatives included).
 *
 * Statuses set on the board survive a rebuild. `overrides` maps a customer
 * key to {status, timestamp} and replaces the board's status for that
 * customer. With `ignoreBoard`, statuses come from Owner_List only.
 */
function buildBoardRows(ownerData, boardData, overrides, ignoreBoard) {
  overrides = overrides || {};

  const previous = {};
  if (!ignoreBoard) {
    boardData.forEach(row => {
      const key = customerKey(row[1]);
      if (key && !previous[key]) {
        previous[key] = {
          status: row[STATUS_COL - 1],
          timestamp: row[TIMESTAMP_COL - 1]
        };
      }
    });
  }

  const customers = [];
  const byKey = {};

  ownerData.forEach(row => {
    const key = customerKey(row[1]);
    if (isBlank(row[0]) || !key) return;

    let customer = byKey[key];
    if (!customer) {
      customer = {
        key,
        id: row[0],
        name: String(row[1]).trim().replace(/\s+/g, " "),
        totals: new Array(QTY_COUNT).fill(0),
        orderCount: 0,
        statuses: [],
        timestamps: [],
        sources: []
      };
      byKey[key] = customer;
      customers.push(customer);
    }

    customer.orderCount++;
    row.slice(2, ORDER_COLS).forEach((value, i) => {
      if (!isBlank(value) && !isNaN(value)) {
        customer.totals[i] += Number(value);
      }
    });

    const status = canonicalStatus(row[STATUS_COL - 1]);
    if (status) customer.statuses.push(status);
    if (!isBlank(row[TIMESTAMP_COL - 1])) customer.timestamps.push(row[TIMESTAMP_COL - 1]);

    const source = String(row[SOURCE_COL - 1]).trim();
    if (source && customer.sources.indexOf(source) === -1) customer.sources.push(source);
  });

  customers.forEach(customer => {
    const override = overrides[customer.key];
    const prev = previous[customer.key];

    if (override) {
      customer.status = override.status;
      customer.timestamp = override.timestamp;
    } else if (prev) {
      customer.status = canonicalStatus(prev.status) || STATUS.NOT_STARTED;
      customer.timestamp = isBlank(prev.timestamp) ? "" : prev.timestamp;
    } else {
      customer.status = customer.statuses[0] || STATUS.NOT_STARTED;
      customer.timestamp = customer.timestamps[0] || "";
    }

    customer.row = [
      customer.id,
      customer.name,
      ...customer.totals.map(total => (total === 0 ? "" : total)),
      customer.status,
      customer.timestamp,
      customer.sources.join(" + ") || "Manual"
    ];
  });

  return customers;
}

/**
 * Rewrites Loading_Board from Owner_List, then copies each customer's status
 * and colour back onto all of their Owner_List rows. Returns the customers.
 */
function rebuildLoadingBoard(ss, overrides, options) {
  options = options || {};

  const ownerSheet = ss.getSheetByName("Owner_List");
  const boardSheet = ss.getSheetByName("Loading_Board");

  const ownerRange = ownerSheet.getRange(START_ROW, 1, NUM_ROWS, SOURCE_COL);
  const ownerData = ownerRange.getValues();
  const boardRange = boardSheet.getRange(START_ROW, 1, NUM_ROWS, SOURCE_COL);

  const customers = buildBoardRows(
    ownerData,
    options.ignoreBoard ? [] : boardRange.getValues(),
    overrides,
    options.ignoreBoard
  );

  const boardValues = [];
  const boardColours = [];
  for (let i = 0; i < NUM_ROWS; i++) {
    const customer = customers[i];
    boardValues.push(customer ? customer.row : new Array(SOURCE_COL).fill(""));
    boardColours.push(new Array(SOURCE_COL).fill(customer ? statusColour(customer.status) : null));
  }
  boardRange.setValues(boardValues).setBackgrounds(boardColours).setFontColor("#000000");

  const byKey = {};
  customers.forEach(customer => { byKey[customer.key] = customer; });

  const ownerStatus = [];
  const ownerColours = [];
  let statusChanged = false;

  ownerData.forEach(row => {
    const customer = isBlank(row[0]) ? null : byKey[customerKey(row[1])];
    let status = row[STATUS_COL - 1];
    let timestamp = row[TIMESTAMP_COL - 1];

    if (customer && (String(status) !== String(customer.status) || String(timestamp) !== String(customer.timestamp))) {
      status = customer.status;
      timestamp = customer.timestamp;
      statusChanged = true;
    }

    ownerStatus.push([status, timestamp]);
    ownerColours.push(new Array(SOURCE_COL).fill(customer ? statusColour(customer.status) : null));
  });

  if (statusChanged) {
    ownerSheet.getRange(START_ROW, STATUS_COL, NUM_ROWS, 2).setValues(ownerStatus);
  }
  ownerRange.setBackgrounds(ownerColours).setFontColor("#000000");

  return customers;
}

function setupTotalsRows(ss) {
  const statusLetter = columnToLetter(STATUS_COL);
  const statusRange = `$${statusLetter}$${START_ROW}:$${statusLetter}$${END_ROW}`;
  const ownerFormulas = [];
  const boardFormulas = [];

  for (let col = 3; col <= ORDER_COLS; col++) {
    const letter = columnToLetter(col);
    const range = `${letter}${START_ROW}:${letter}${END_ROW}`;

    ownerFormulas.push(`=SUM(${range})`);
    boardFormulas.push(
      `=SUMIFS(${range},${statusRange},"<>Cancelled",${statusRange},"<>Rescheduled",${statusRange},"<>Leftover")`
    );
  }

  ss.getSheetByName("Owner_List").getRange(TOTALS_ROW, 3, 1, QTY_COUNT).setFormulas([ownerFormulas]);
  ss.getSheetByName("Loading_Board").getRange(TOTALS_ROW, 3, 1, QTY_COUNT).setFormulas([boardFormulas]);
}

// ==========================================
// LEFTOVER TAB
// ==========================================

/**
 * Makes the Leftover tab hold exactly one block per customer whose status is
 * Leftover, in Loading_Board order, with each block's Order row matching the
 * board. Taken amounts already typed in are kept.
 */
function reconcileLeftoverBlocks(ss, customers) {
  const sheet = ss.getSheetByName("Leftover");

  const expected = {};
  const boardIndex = {};
  customers.forEach((customer, i) => {
    boardIndex[customer.key] = i;
    if (customer.status === STATUS.LEFTOVER) expected[customer.key] = customer;
  });

  // Delete blocks for customers that are no longer Leftover, and duplicates
  let blocks = readLeftoverBlocks(sheet);
  const seen = {};
  blocks
    .filter(block => {
      if (!expected[block.key] || seen[block.key]) return true;
      seen[block.key] = true;
      return false;
    })
    .sort((a, b) => b.row - a.row)
    .forEach(block => sheet.deleteRows(block.row, block.size));

  blocks = readLeftoverBlocks(sheet);

  // Keep the Order rows in step with the board
  blocks.forEach(block => {
    const wanted = leftoverOrderRow(expected[block.key]);
    const current = block.values;
    if (wanted.some((value, i) => String(value) !== String(current[i]))) {
      sheet.getRange(block.row, 1, 1, LEFTOVER_COLS).setValues([wanted]);
    }
  });

  // Add missing blocks in board order
  let created = false;
  customers.forEach(customer => {
    if (!expected[customer.key]) return;
    if (blocks.some(block => block.key === customer.key)) return;

    const nextBlock = blocks.find(block => boardIndex[block.key] > boardIndex[customer.key]);
    const insertRow = nextBlock
      ? nextBlock.row
      : Math.max(sheet.getLastRow() + 1, LEFTOVER_START_ROW);

    insertLeftoverBlock(sheet, insertRow, customer, Boolean(nextBlock));
    blocks = readLeftoverBlocks(sheet);
    created = true;
  });

  if (created) setupLeftoverSummaryRows(sheet);
}

function readLeftoverBlocks(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < LEFTOVER_START_ROW) return [];

  const values = sheet
    .getRange(LEFTOVER_START_ROW, 1, lastRow - LEFTOVER_START_ROW + 1, LEFTOVER_COLS)
    .getValues();

  const blocks = [];
  values.forEach((row, i) => {
    if (String(row[2]).trim() !== "Order") return;

    let size = 1;
    while (
      size < 3 &&
      i + size < values.length &&
      ["Taken", "Leftover"].indexOf(String(values[i + size][2]).trim()) !== -1
    ) {
      size++;
    }

    blocks.push({
      row: LEFTOVER_START_ROW + i,
      size,
      key: customerKey(row[1]),
      values: row
    });
  });

  return blocks;
}

function leftoverOrderRow(customer) {
  return [customer.id, customer.name, "Order", ...customer.row.slice(2, ORDER_COLS)];
}

function insertLeftoverBlock(sheet, insertRow, customer, beforeExistingBlock) {
  if (beforeExistingBlock) {
    sheet.insertRowsBefore(insertRow, 3);
  } else if (insertRow + 2 > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), insertRow + 2 - sheet.getMaxRows());
  }

  // Inserted rows copy the formatting around them, so start clean
  sheet
    .getRange(insertRow, 1, 3, SKIP_COL)
    .clearContent()
    .clearDataValidations()
    .setBorder(false, false, false, false, false, false)
    .setBackground(null)
    .setFontColor("#000000")
    .setFontWeight("normal");

  // Leftover = Order - Taken
  const leftoverRow = ["", "", "Leftover"];
  for (let col = 4; col <= LEFTOVER_COLS; col++) {
    const letter = columnToLetter(col);
    const orderCell = `${letter}${insertRow}`;
    const takenCell = `${letter}${insertRow + 1}`;
    leftoverRow.push(`=IF(${orderCell}="","",${orderCell}-IF(${takenCell}="",0,${takenCell}))`);
  }

  sheet.getRange(insertRow, 1, 3, 1).setNumberFormat("0");
  sheet.getRange(insertRow, 1, 3, LEFTOVER_COLS).setValues([
    leftoverOrderRow(customer),
    ["", "", "Taken", ...new Array(QTY_COUNT).fill("")],
    leftoverRow
  ]);

  sheet
    .getRange(insertRow + 2, 3)
    .setFontWeight("bold");

  // Make Leftover label and quantities red.
  sheet
    .getRange(insertRow + 2, 3, 1, QTY_COUNT + 1)
    .setFontColor("red");

  // Borders & Colors
  sheet
    .getRange(insertRow, 1, 1, LEFTOVER_COLS)
    .setBorder(true, null, null, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setBackground("#f8f8f8");

  sheet
    .getRange(insertRow + 1, 1, 1, LEFTOVER_COLS)
    .setBackground("#e8f4ff");

  sheet
    .getRange(insertRow + 2, 1, 1, LEFTOVER_COLS)
    .setBorder(null, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setBackground("#fff0f0");

  // Column P checkbox: tick to NOT carry this leftover into the next shift
  sheet.getRange(insertRow + 2, SKIP_COL).insertCheckboxes();
}

function setupLeftoverSummaryRows(sheet) {
  sheet.getRange("C7:C9").setValues([["Total Order"], ["Total Taken"], ["Total Leftover"]]);

  const formulas = [[], [], []];
  for (let col = 4; col <= LEFTOVER_COLS; col++) {
    const letter = columnToLetter(col);
    const sumRange = `${letter}$${LEFTOVER_START_ROW}:${letter}`;
    const typeRange = `$C$${LEFTOVER_START_ROW}:$C`;

    formulas[0].push(`=SUMIF(${typeRange},"Order",${sumRange})`);
    formulas[1].push(`=SUMIF(${typeRange},"Taken",${sumRange})`);
    formulas[2].push(`=SUMIF(${typeRange},"Leftover",${sumRange})`);
  }

  sheet.getRange(LEFTOVER_TOTAL_ORDER_ROW, 4, 3, QTY_COUNT).setFormulas(formulas);
  sheet.getRange("C7:O9").setFontWeight("bold");
  sheet.getRange(`C${LEFTOVER_TOTAL_LEFTOVER_ROW}:O${LEFTOVER_TOTAL_LEFTOVER_ROW}`).setFontColor("red");
}

function clearLeftoverTab(sheet) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < LEFTOVER_START_ROW) return;

  sheet
    .getRange(LEFTOVER_START_ROW, 1, maxRows - LEFTOVER_START_ROW + 1, SKIP_COL)
    .clearContent()
    .clearDataValidations()
    .setBackground(null)
    .setBorder(false, false, false, false, false, false)
    .setFontColor("#000000")
    .setFontWeight("normal");

  // Drop the empty rows past the spare ones (keeps anything beyond column P)
  const keepRows = Math.max(sheet.getLastRow(), LEFTOVER_START_ROW + LEFTOVER_SPARE_ROWS - 1);
  if (maxRows > keepRows) {
    sheet.deleteRows(keepRows + 1, maxRows - keepRows);
  }
}

// ==========================================
// MENU ACTIONS
// ==========================================

function syncAllStatuses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let converted = 0;

  withLock(() => {
    converted = convertNotStartedToLeftover(ss);
  });

  showAlert(`Sync Complete! ${converted} 'Not started' customer(s) were converted to Leftovers, and all tabs are fully updated.`);
}

// Turns every "Not started" customer into "Leftover" and brings the Leftover
// tab up to date. Returns how many customers were converted.
// Callers must hold the lock.
function convertNotStartedToLeftover(ss) {
  let customers = rebuildLoadingBoard(ss);

  const stamp = nowStamp();
  const overrides = {};
  let converted = 0;
  customers.forEach(customer => {
    if (customer.status === STATUS.NOT_STARTED) {
      overrides[customer.key] = { status: STATUS.LEFTOVER, timestamp: stamp };
      converted++;
    }
  });

  if (converted > 0) customers = rebuildLoadingBoard(ss, overrides);
  reconcileLeftoverBlocks(ss, customers);
  return converted;
}

function reconcileBoards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let customers = [];

  withLock(() => {
    setupTotalsRows(ss);
    customers = rebuildLoadingBoard(ss);
    reconcileLeftoverBlocks(ss, customers);
  });

  const orders = customers.reduce((sum, customer) => sum + customer.orderCount, 0);
  showAlert(`Loading Board rebuilt: ${orders} order(s) merged into ${customers.length} customer(s). Totals and Leftover tab checked.`);
}

function resetAllStatusesAndLeftovers() {
  const ui = SpreadsheetApp.getUi();

  // 1. Safety Check
  const response = ui.alert(
    "⚠️ WARNING: Reset All Orders?",
    "This will clear every order's status, erase all timestamps, and permanently delete everything on the Leftover tab.\n\nYour original orders and bread quantities will be kept.\n\nAre you sure you want to proceed?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  withLock(() => {
    const ownerSheet = ss.getSheetByName("Owner_List");
    const leftoverSheet = ss.getSheetByName("Leftover");

    // 2. Wipe Statuses and Timestamps on the Owner List
    const ids = ownerSheet.getRange(START_ROW, 1, NUM_ROWS, 1).getValues();
    const statusRange = ownerSheet.getRange(START_ROW, STATUS_COL, NUM_ROWS, 2);
    const statusData = statusRange.getValues().map((row, i) =>
      isBlank(ids[i][0]) ? row : [STATUS.NOT_STARTED, ""]
    );
    statusRange.setValues(statusData);

    // 3. Erase the Leftover Tab Completely
    clearLeftoverTab(leftoverSheet);
    setupLeftoverSummaryRows(leftoverSheet);

    // 4. Rebuild the board from the Owner List (also resets the colours)
    rebuildLoadingBoard(ss, {}, { ignoreBoard: true });
  });

  ui.alert("✅ Reset Complete! All statuses are cleared, colors are reset, and leftovers have been removed.");
}

function processLeftovers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeZone = ss.getSpreadsheetTimeZone();
  const ownerSheet = ss.getSheetByName("Owner_List");
  const leftoverSheet = ss.getSheetByName("Leftover");

  const currentDate = ownerSheet.getRange("A1").getValue();
  const currentShift = ownerSheet.getRange("B1").getValue();
  const emailList = String(ownerSheet.getRange("H2").getValue()).trim();

  const overrideDate = leftoverSheet.getRange("N2").getValue();
  const overrideShift = String(leftoverSheet.getRange("O2").getValue()).trim();

  let nextInfo;
  if (overrideDate && overrideShift) {
    nextInfo = describeShift(new Date(overrideDate), overrideShift, timeZone);
  } else if (currentDate instanceof Date) {
    nextInfo = getNextShiftInfo(currentDate, currentShift, timeZone);
  } else {
    showAlert("Owner_List A1 must hold this sheet's date before the next shift sheet can be generated.");
    return;
  }

  const newFileName = `Loading List - ${nextInfo.dayShort} - ${nextInfo.dateText} - ${nextInfo.shift}`;

  // Ask before making a second file with the same name
  if (DriveApp.getFilesByName(newFileName).hasNext()) {
    const ui = SpreadsheetApp.getUi();
    const answer = ui.alert(
      "A loading list with this name already exists",
      `"${newFileName}" already exists. Create another copy with the same name?`,
      ui.ButtonSet.YES_NO
    );
    if (answer !== ui.Button.YES) {
      showAlert("Cancelled. No new loading list was created.");
      return;
    }
  }

  // Mark "Not started" customers as Leftover first, so none are missed
  let converted = 0;
  withLock(() => {
    converted = convertNotStartedToLeftover(ss);
  });

  // Collect the leftovers before copying
  const lastLeftoverRow = leftoverSheet.getLastRow();
  const leftoverValues = lastLeftoverRow >= LEFTOVER_START_ROW
    ? leftoverSheet.getRange(LEFTOVER_START_ROW, 1, lastLeftoverRow - LEFTOVER_START_ROW + 1, SKIP_COL).getValues()
    : [];
  let carryRows = collectCarryForward(leftoverValues);

  const warnings = [];
  if (converted > 0) {
    warnings.push(`${converted} 'Not started' customer(s) were marked Leftover before the copy.`);
  }
  if (carryRows.length > NUM_ROWS) {
    warnings.push(`Only the first ${NUM_ROWS} of ${carryRows.length} leftovers fit on the new Owner_List.`);
    carryRows = carryRows.slice(0, NUM_ROWS);
  }

  const rootFolders = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  const rootFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(ROOT_FOLDER_NAME);
  const yearFolder = getOrCreateFolder(rootFolder, nextInfo.year);
  const monthFolder = getOrCreateFolder(yearFolder, nextInfo.monthName);

  const newFile = DriveApp.getFileById(ss.getId()).makeCopy(newFileName, monthFolder);
  const newSS = SpreadsheetApp.openById(newFile.getId());

  const newOwnerSheet = newSS.getSheetByName("Owner_List");
  const newBoardSheet = newSS.getSheetByName("Loading_Board");
  const newLeftoverSheet = newSS.getSheetByName("Leftover");

  // Wipe the new sheets completely clean to avoid ghost data
  newLeftoverSheet.getRange("N2:O2").clearContent();
  newOwnerSheet.getRange("A1").setValue(nextInfo.date);
  newOwnerSheet.getRange("B1").setValue(nextInfo.shift);
  newOwnerSheet.getRange("H2").setValue(emailList);

  [newOwnerSheet, newBoardSheet].forEach(sheet => {
    const maxRows = sheet.getMaxRows();
    if (maxRows >= START_ROW) {
      sheet.getRange(START_ROW, 1, maxRows - START_ROW + 1, SOURCE_COL)
        .clearContent().setBackground(null).setFontColor("#000000");
    }
  });

  clearLeftoverTab(newLeftoverSheet);
  setupLeftoverSummaryRows(newLeftoverSheet);

  if (carryRows.length > 0) {
    newOwnerSheet.getRange(START_ROW, 1, carryRows.length, SOURCE_COL).setValues(carryRows);
  }

  setupTotalsRows(newSS);
  rebuildLoadingBoard(newSS, {}, { ignoreBoard: true });

  try {
    copyProtections(ss, newSS);
  } catch (err) {
    warnings.push(`Protections were not copied: ${err.message}`);
  }

  if (emailList) {
    try {
      MailApp.sendEmail({
        to: emailList,
        subject: newFileName,
        body: `Hello,\n\nThe next loading list has been created.\n\nDate: ${nextInfo.dayShort}, ${nextInfo.dateText}\nShift: ${nextInfo.shift}\n\nLink: ${newFile.getUrl()}`
      });
    } catch (err) {
      warnings.push(`Email was not sent: ${err.message}`);
    }
  } else {
    warnings.push("No email sent: the email list in Owner_List H2 is empty.");
  }

  showAlert(
    `Leftovers processed. "${newFileName}" was created with ${carryRows.length} leftover order(s).` +
    (warnings.length ? `\n\n${warnings.join("\n")}` : "")
  );
}

/**
 * Turns Leftover tab rows (from row LEFTOVER_START_ROW, columns A:P) into
 * Owner_List rows for the next shift. Skips ticked rows and rows with
 * nothing left. Negative amounts are carried over.
 */
function collectCarryForward(leftoverValues) {
  const rows = [];
  let customerName = "";

  leftoverValues.forEach(row => {
    const rowType = String(row[2]).trim();

    if (rowType === "Order") {
      customerName = String(row[1]).trim();
      return;
    }
    if (rowType !== "Leftover") return;

    const skip = row[SKIP_COL - 1];
    if (skip === true || String(skip).toUpperCase() === "TRUE") return;

    const quantities = row.slice(3, 3 + QTY_COUNT).map(value => {
      const num = Number(value);
      return isBlank(value) || isNaN(num) || num === 0 ? "" : num;
    });

    if (customerName && quantities.some(value => value !== "")) {
      rows.push([rows.length + 1, customerName, ...quantities, STATUS.NOT_STARTED, "", "Leftover"]);
    }
  });

  return rows;
}

function getNextShiftInfo(currentDate, currentShift, timeZone) {
  let date = new Date(currentDate);
  let nextShift;

  if (String(currentShift).trim().toLowerCase() === "morning") {
    nextShift = "Night";
  } else {
    nextShift = "Morning";
    date = startOfDay(new Date(date.getTime() + 36 * 60 * 60 * 1000), timeZone);
  }

  // Skip Sunday Morning ("u" is the day of the week, 7 = Sunday).
  if (nextShift === "Morning" && Utilities.formatDate(date, timeZone, "u") === "7") {
    nextShift = "Night";
  }

  return describeShift(date, nextShift, timeZone);
}

function describeShift(date, shift, timeZone) {
  return {
    date,
    shift,
    year: Utilities.formatDate(date, timeZone, "yyyy"),
    monthName: Utilities.formatDate(date, timeZone, "MMMM"),
    dateText: Utilities.formatDate(date, timeZone, "yyyy-MM-dd"),
    dayShort: Utilities.formatDate(date, timeZone, "EEE")
  };
}

function startOfDay(date, timeZone) {
  return Utilities.parseDate(Utilities.formatDate(date, timeZone, "yyyy-MM-dd"), timeZone, "yyyy-MM-dd");
}

function copyProtections(sourceSS, newSS) {
  const me = Session.getEffectiveUser();
  const sourceSheets = sourceSS.getSheets();

  sourceSheets.forEach(sourceSheet => {
    const destSheet = newSS.getSheetByName(sourceSheet.getName());
    if (!destSheet) return;

    // Start from none, in case the file copy already brought some across
    destSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    destSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

    // 1. Copy Sheet-Level Protections (Tab Locks)
    const sheetProtections = sourceSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    sheetProtections.forEach(p => {
      const newProtection = destSheet.protect().setDescription(p.getDescription());
      newProtection.setWarningOnly(p.isWarningOnly());

      if (!p.isWarningOnly()) {
        // Copy exact editors
        const editors = p.getEditors().map(user => user.getEmail());
        newProtection.addEditor(me);
        newProtection.removeEditors(newProtection.getEditors()); // Strip script defaults
        if (editors.length > 0) newProtection.addEditors(editors);

        // Copy any unprotected ranges (e.g., "Protect whole sheet EXCEPT column B")
        const unprotected = p.getUnprotectedRanges().map(r => destSheet.getRange(r.getA1Notation()));
        newProtection.setUnprotectedRanges(unprotected);
      }
    });

    // 2. Copy Range-Level Protections (Specific Cell Locks)
    const rangeProtections = sourceSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    rangeProtections.forEach(p => {
      const newProtection = destSheet.getRange(p.getRange().getA1Notation()).protect().setDescription(p.getDescription());
      newProtection.setWarningOnly(p.isWarningOnly());

      if (!p.isWarningOnly()) {
        const editors = p.getEditors().map(user => user.getEmail());
        newProtection.addEditor(me);
        newProtection.removeEditors(newProtection.getEditors());
        if (editors.length > 0) newProtection.addEditors(editors);
      }
    });
  });
}

// ==========================================
// HELPERS
// ==========================================

function canonicalStatus(value) {
  const text = String(isBlank(value) ? "" : value).trim();
  if (!text) return "";

  const lower = text.toLowerCase();
  for (const name in STATUS) {
    if (STATUS[name].toLowerCase() === lower) return STATUS[name];
  }
  return text;
}

function statusColour(status) {
  return STATUS_COLOURS[String(status).trim().toLowerCase()] || null;
}

// Customer names match ignoring case and extra spaces
function customerKey(name) {
  return String(isBlank(name) ? "" : name).trim().replace(/\s+/g, " ").toLowerCase();
}

function isBlank(value) {
  return value === "" || value === null || value === undefined;
}

function nowStamp() {
  return Utilities.formatDate(new Date(), TIME_ZONE, "yyyy-MM-dd HH:mm:ss");
}

// Stops two people's edits from rebuilding the board at the same time.
// Simple triggers only get 30 seconds, so if the lock is busy for 10, carry on.
function withLock(fn) {
  let lock = null;
  try {
    lock = LockService.getDocumentLock();
    if (lock && !lock.tryLock(10000)) lock = null;
  } catch (err) {
    lock = null;
  }

  try {
    return fn();
  } finally {
    if (lock) lock.releaseLock();
  }
}

function showAlert(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (err) {
    console.log(message);
  }
}

function columnToLetter(column) {
  let temp;
  let letter = "";

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return parentFolder.createFolder(folderName);
}
