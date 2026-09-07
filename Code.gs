// @ts-nocheck
const START_ROW = 8;              // Owner_List and Loading_Board data start row
const END_ROW = 106;

const LEFTOVER_START_ROW = 10;    // Leftover data starts row 10
const LEFTOVER_TOTAL_ORDER_ROW = 7;
const LEFTOVER_TOTAL_TAKEN_ROW = 8;
const LEFTOVER_TOTAL_LEFTOVER_ROW = 9;

const ORDER_COLS = 14;            // A:M = Order ID, Customer Name, bread quantities
const STATUS_COL = 15;            // N = Status
const TIMESTAMP_COL = 16;         // O = Status Changed At
const SOURCE_COL = 17;            // P = Source

function syncLoadingBoard() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ownerSheet =
    ss.getSheetByName("Owner_List");

  const boardSheet =
    ss.getSheetByName("Loading_Board");

  const ownerData =
    ownerSheet
      .getRange(
        START_ROW,
        1,
        END_ROW - START_ROW + 1,
        SOURCE_COL
      )
      .getValues();

  const outputRows = [];

  ownerData.forEach(row => {

    const orderId = row[0];
    const customer = row[1];

    if (!orderId || !customer)
      return;

    const cleanQuantities =
      row
      .slice(2,14)
      .map(value => {

        if (
          value === "" ||
          Number(value) === 0
        ) {
          return "";
        }

        return Number(value);

      });

    outputRows.push([

      orderId,
      customer,

      ...cleanQuantities,

      row[STATUS_COL-1] || "Not Started",

      row[TIMESTAMP_COL-1] || "",

      row[SOURCE_COL-1] || ""

    ]);

  });

  boardSheet
    .getRange(
      START_ROW,
      1,
      END_ROW - START_ROW + 1,
      SOURCE_COL
    )
    .clearContent()
    .setBackground(null)
    .setFontColor("#000000");

  if (outputRows.length > 0) {

    boardSheet
      .getRange(
        START_ROW,
        1,
        outputRows.length,
        SOURCE_COL
      )
      .setValues(outputRows);

    for (
      let i = 0;
      i < outputRows.length;
      i++
    ) {

      updateRowStyle(
        boardSheet,
        START_ROW + i
      );

    }

  }

}
function installedOnEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const startRow = e.range.getRow();
  const startCol = e.range.getColumn();
  const endRow = startRow + e.range.getNumRows() - 1;
  const endCol = startCol + e.range.getNumColumns() - 1;

  // ==========================================
  // 1. IF THE EDIT WAS ON THE OWNER LIST
  // ==========================================
  if (sheetName === "Owner_List") {
    if (endRow < START_ROW) return;
    
    // (Email alerts removed to prevent Google Security blocks)
    
    for (let row = Math.max(startRow, START_ROW); row <= endRow; row++) {
      const orderIdCell = sheet.getRange(row, 1);
      const customerName = sheet.getRange(row, 2).getValue();
      
      // Generate ID for new customers
      if (customerName && !orderIdCell.getValue()) {
        orderIdCell.setValue(generateOrderId());
        sheet.getRange(row, SOURCE_COL).setValue("Manual");
      }

      // LIGHTNING FAST: SYNC ONLY THIS SPECIFIC ROW
      syncSingleRowToBoard(row);
    }
    
    syncLeftoverOrderRows();
    return;
  }

  // ==========================================
  // 2. IF THE EDIT WAS ON THE LOADING BOARD
  // ==========================================
  if (sheetName === "Loading_Board") {
    if (endRow < START_ROW) return;
    
    // Only run if the Status column was the specific cell that got edited
    if (STATUS_COL >= startCol && STATUS_COL <= endCol) {
      for (let row = Math.max(startRow, START_ROW); row <= endRow; row++) {
        const orderId = sheet.getRange(row, 1).getValue();
        const newStatus = sheet.getRange(row, STATUS_COL).getValue();
        
        // Stamp the time
        sheet.getRange(row, TIMESTAMP_COL).setValue(
          Utilities.formatDate(new Date(), "Africa/Lagos", "yyyy-MM-dd HH:mm:ss")
        );
        
        // Update the color instantly
        updateRowStyle(sheet, row);
        
        if (orderId) {
          // Push the new status back to the Owner List
          updateOwnerStatus(orderId, newStatus);
          
          // Generate or delete the red Leftover rows
          if (newStatus === "Leftover") {
            createLeftoverRows(orderId);
          } else {
            deleteLeftoverRows(orderId);
          }
        }
      }
    }
  }
}
function generateOrderId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ownerSheet = ss.getSheetByName("Owner_List");

  let highestId = 0;

  for (let row = START_ROW; row <= END_ROW; row++) {
    const value = ownerSheet.getRange(row, 1).getValue();

    if (value && !isNaN(value)) {
      highestId = Math.max(highestId, Number(value));
    }
  }

  return highestId + 1;
}

function updateOwnerStatus(orderIdToFind, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ownerSheet = ss.getSheetByName("Owner_List");

  for (let row = START_ROW; row <= END_ROW; row++) {
    const orderId = ownerSheet.getRange(row, 1).getValue();

    if (String(orderId).trim() === String(orderIdToFind).trim()) {
      ownerSheet.getRange(row, STATUS_COL).setValue(newStatus);      // N
      ownerSheet.getRange(row, TIMESTAMP_COL).setValue(
  Utilities.formatDate(
    new Date(),
    
    "Africa/Lagos",
    "yyyy-MM-dd HH:mm:ss"
  )
); 
updateRowStyle(ownerSheet, row); // O
      return;
    }
  }
}

function createLeftoverRows(orderIdToFind) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const boardSheet = ss.getSheetByName("Loading_Board");
  const leftoverSheet = ss.getSheetByName("Leftover");

  setupLeftoverSummaryRows();

  // Prevent duplicate leftover sections.
  for (let row = LEFTOVER_START_ROW; row <= leftoverSheet.getLastRow(); row++) {
    const existingOrderId = String(leftoverSheet.getRange(row, 1).getValue()).trim();
    const rowType = String(leftoverSheet.getRange(row, 3).getValue()).trim();

    if (normalizeOrderId(existingOrderId) === normalizeOrderId(orderIdToFind) && rowType === "Order") {
      return;
    }
  }

  // Find order in Loading_Board.
  let sourceRow = null;
  for (let row = START_ROW; row <= END_ROW; row++) {
    const orderId = boardSheet.getRange(row, 1).getValue();
    if (String(orderId).trim() === String(orderIdToFind).trim()) {
      sourceRow = row;
      break;
    }
  }

  if (!sourceRow) {
    return;
  }

  const orderData = boardSheet.getRange(sourceRow, 1, 1, ORDER_COLS).getValues()[0];

  // --- NEW SORTING LOGIC STARTS HERE ---
  let insertRow = leftoverSheet.getLastRow() + 1;
  if (insertRow < LEFTOVER_START_ROW) insertRow = LEFTOVER_START_ROW;

  for (let r = LEFTOVER_START_ROW; r <= leftoverSheet.getLastRow(); r++) {
    const rowType = String(leftoverSheet.getRange(r, 3).getValue()).trim();
    if (rowType === "Order") {
      const existingOrderId = leftoverSheet.getRange(r, 1).getValue();
      let existingSourceRow = 9999;
      
      for (let bRow = START_ROW; bRow <= END_ROW; bRow++) {
        if (String(boardSheet.getRange(bRow, 1).getValue()).trim() === String(existingOrderId).trim()) {
          existingSourceRow = bRow;
          break;
        }
      }

      if (sourceRow < existingSourceRow) {
        insertRow = r;
        break;
      }
    }
  }

  if (insertRow > leftoverSheet.getMaxRows()) {
    leftoverSheet.insertRowsAfter(leftoverSheet.getMaxRows(), 3);
  } else {
    leftoverSheet.insertRowsBefore(insertRow, 3);
  }
  // --- NEW SORTING LOGIC ENDS HERE ---

  leftoverSheet
    .getRange(insertRow, 1, 3, 1)
    .setNumberFormat("@");

  // Row 1: Order
  leftoverSheet.getRange(insertRow, 1, 1, 15).setValues([[
    orderData[0],
    orderData[1],
    "Order",
    ...orderData.slice(2, 14)
  ]]);

  // Row 2: Taken
  leftoverSheet.getRange(insertRow + 1, 1, 1, 15).setValues([[
    "",
    "",
    "Taken",
    "", "", "", "", "", "", "", "", "", "", "", ""
  ]]);

  // Row 3: Leftover
  leftoverSheet.getRange(insertRow + 2, 1, 1, 3).setValues([[
    "",
    "",
    "Leftover"
  ]]);

  leftoverSheet
    .getRange(insertRow + 2, 3)
    .setFontColor("red")
    .setFontWeight("bold");

  // Formulas: Leftover = Order - Taken
  for (let col = 4; col <= 15; col++) {
    const orderCell = leftoverSheet.getRange(insertRow, col).getA1Notation();
    const takenCell = leftoverSheet.getRange(insertRow + 1, col).getA1Notation();

    leftoverSheet
      .getRange(insertRow + 2, col)
      .setFormula(
        `=IF(${orderCell}="","",${orderCell}-IF(${takenCell}="",0,${takenCell}))`
      );
  }

  // Make Leftover quantities red.
  leftoverSheet
    .getRange(insertRow + 2, 4, 1, 12)
    .setFontColor("red");

  // Borders & Colors
  leftoverSheet
    .getRange(insertRow, 1, 1, 15)
    .setBorder(true, null, null, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setBackground("#f8f8f8");

  leftoverSheet
    .getRange(insertRow + 1, 1, 1, 15)
    .setBackground("#e8f4ff");

  leftoverSheet
    .getRange(insertRow + 2, 1, 1, 15)
    .setBorder(null, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setBackground("#fff0f0");

  leftoverSheet
    .getRange(insertRow, 1, 3, 1)
    .setNumberFormat("0");

  // --- NEW: ADD CHECKBOX TO COLUMN P (16) FOR "DO NOT CARRY FORWARD" ---
  leftoverSheet.getRange(insertRow + 2, 16).insertCheckboxes();
}

function setupLeftoverSummaryRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leftoverSheet = ss.getSheetByName("Leftover");

  leftoverSheet.getRange("C7").setValue("Total Order");
  leftoverSheet.getRange("C8").setValue("Total Taken");
  leftoverSheet.getRange("C9").setValue("Total Leftover");

  for (let col = 4; col <= 15; col++) {
    const colLetter = columnToLetter(col);

    leftoverSheet
      .getRange(LEFTOVER_TOTAL_ORDER_ROW, col)
      .setFormula(`=SUMIF($C$${LEFTOVER_START_ROW}:$C,"Order",${colLetter}$${LEFTOVER_START_ROW}:${colLetter})`);

    leftoverSheet
      .getRange(LEFTOVER_TOTAL_TAKEN_ROW, col)
      .setFormula(`=SUMIF($C$${LEFTOVER_START_ROW}:$C,"Taken",${colLetter}$${LEFTOVER_START_ROW}:${colLetter})`);

    leftoverSheet
      .getRange(LEFTOVER_TOTAL_LEFTOVER_ROW, col)
      .setFormula(`=SUMIF($C$${LEFTOVER_START_ROW}:$C,"Leftover",${colLetter}$${LEFTOVER_START_ROW}:${colLetter})`);
  }

  leftoverSheet.getRange("C7:O9").setFontWeight("bold");
leftoverSheet.getRange("C9:O9").setFontColor("red");

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

function syncLeftoverOrderRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Redirected: Look at Loading_Board instead of Owner_List for the merged truth
  const boardSheet = ss.getSheetByName("Loading_Board");
  const leftoverSheet = ss.getSheetByName("Leftover");

  for (let boardRow = START_ROW; boardRow <= END_ROW; boardRow++) {
    const boardOrderId = String(boardSheet.getRange(boardRow, 1).getValue()).trim();
    
    if (!boardOrderId) {
      continue;
    }

    const boardData = boardSheet
      .getRange(boardRow, 1, 1, ORDER_COLS)
      .getValues()[0];

    for (let leftRow = LEFTOVER_START_ROW; leftRow <= leftoverSheet.getLastRow(); leftRow++) {
      const leftoverOrderId = String(leftoverSheet.getRange(leftRow, 1).getValue()).trim();
      const rowType = String(leftoverSheet.getRange(leftRow, 3).getValue()).trim();

      if (leftoverOrderId === boardOrderId && rowType === "Order") {
        leftoverSheet.getRange(leftRow, 1, 1, 15).setValues([[
          boardData[0],
          boardData[1],
          "Order",
          ...boardData.slice(2, 14)
        ]]);
      }
    }
  }
}

function processLeftovers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ownerSheet = ss.getSheetByName("Owner_List");
  const leftoverSheet = ss.getSheetByName("Leftover");

  setupLeftoverSummaryRows();

  const currentDate = ownerSheet.getRange("A1").getValue();
  const currentShift = ownerSheet.getRange("B1").getValue();
  const emailList = ownerSheet.getRange("H2").getValue();
  const nextInfo = getNextShiftInfo(currentDate, currentShift);

  const overrideDate = leftoverSheet.getRange("N2").getValue();
  const overrideShift = leftoverSheet.getRange("O2").getValue();

  if (overrideDate && overrideShift) {
    nextInfo.date = new Date(overrideDate);
    nextInfo.shift = overrideShift;
    nextInfo.year = nextInfo.date.getFullYear();
    nextInfo.monthName = Utilities.formatDate(nextInfo.date, Session.getScriptTimeZone(), "MMMM");
    nextInfo.dateText = Utilities.formatDate(nextInfo.date, Session.getScriptTimeZone(), "yyyy-MM-dd");
    nextInfo.dayShort = Utilities.formatDate(nextInfo.date, Session.getScriptTimeZone(), "EEE");
  }

  const rootFolder = DriveApp.getFoldersByName("Loading List").next();
  const yearFolder = getOrCreateFolder(rootFolder, String(nextInfo.year));
  const monthFolder = getOrCreateFolder(yearFolder, nextInfo.monthName);

  const newFileName = `Loading List - ${nextInfo.dayShort} - ${nextInfo.dateText} - ${nextInfo.shift}`;
  const currentFile = DriveApp.getFileById(ss.getId());
  
  const newFile = currentFile.makeCopy(newFileName, monthFolder);
  const newSS = SpreadsheetApp.openById(newFile.getId());

  // Automatically wire up the script permissions for the new file
  ScriptApp.newTrigger("installedOnEdit")
    .forSpreadsheet(newSS)
    .onEdit()
    .create();

  const newOwnerSheet = newSS.getSheetByName("Owner_List");
  const newBoardSheet = newSS.getSheetByName("Loading_Board");
  const newLeftoverSheet = newSS.getSheetByName("Leftover");

  // Wipe the new sheets completely clean to avoid ghost data
  newLeftoverSheet.getRange("M2:N2").clearContent();
  newOwnerSheet.getRange("A1").setValue(nextInfo.date);
  newOwnerSheet.getRange("B1").setValue(nextInfo.shift);
  newOwnerSheet.getRange("H2").setValue(emailList);

  const maxOwnerRows = newOwnerSheet.getMaxRows();
  if (maxOwnerRows >= START_ROW) {
    newOwnerSheet.getRange(START_ROW, 1, maxOwnerRows - START_ROW + 1, SOURCE_COL)
      .clearContent().setBackground(null).setFontColor("#000000");
  }

  const maxBoardRows = newBoardSheet.getMaxRows();
  if (maxBoardRows >= START_ROW) {
    newBoardSheet.getRange(START_ROW, 1, maxBoardRows - START_ROW + 1, SOURCE_COL)
      .clearContent().setBackground(null).setFontColor("#000000");
  }
  
  const maxLeftoverRows = newLeftoverSheet.getMaxRows();
  if (maxLeftoverRows >= LEFTOVER_START_ROW) {
    newLeftoverSheet.getRange(LEFTOVER_START_ROW, 1, maxLeftoverRows - LEFTOVER_START_ROW + 1, 16)
      .clearContent().clearDataValidations();
  }

  let targetRow = START_ROW;
  let newOrderId = 1;

  // Process the Leftovers
  const lastLeftoverRow = leftoverSheet.getLastRow();
  for (let row = LEFTOVER_START_ROW; row <= lastLeftoverRow; row++) {
    const rowType = String(leftoverSheet.getRange(row, 3).getValue()).trim();

    if (rowType === "Leftover") {
      // Check the Column P Checkbox
      const skipCheckbox = leftoverSheet.getRange(row, 16).getValue();
      if (skipCheckbox === true || String(skipCheckbox).toUpperCase() === "TRUE") continue;

      const customerName = findCustomerAbove(leftoverSheet, row);
      
      // Grab the bread arrays and format them
      const quantities = leftoverSheet.getRange(row, 4, 1, 12).getValues()[0].map(value => {
        const num = Number(value);
        return isNaN(num) || num === 0 ? "" : num;
      });

      // Valid if any bread column has a non-zero number
      const hasLeftover = quantities.some(value => value !== "");

      if (customerName && hasLeftover) {
        newOwnerSheet.getRange(targetRow, 1, 1, SOURCE_COL).setValues([[
          newOrderId,
          customerName,
          ...quantities,
          "Not Started",
          "",
          "Leftover"
        ]]);
        targetRow++;
        newOrderId++;
      }
    }
  }
  if (emailList) {
    MailApp.sendEmail({
      to: emailList,
      subject: `Loading List - ${nextInfo.dayShort} - ${nextInfo.dateText} - ${nextInfo.shift}`,
      body: `Hello,\n\nThe next loading list has been created.\n\nDate: ${nextInfo.dayShort}, ${nextInfo.dateText}\nShift: ${nextInfo.shift}\n\nLink: ${newFile.getUrl()}`
    });
  }

  SpreadsheetApp.getUi().alert("Leftovers processed. New loading list created and emailed.");
}
function getNextShiftInfo(currentDate, currentShift) {
  const date = new Date(currentDate);
  let nextShift = "";

  if (String(currentShift).toLowerCase() === "morning") {
    nextShift = "Night";
  } else {
    nextShift = "Morning";
    date.setDate(date.getDate() + 1);
  }

  // Skip Sunday Morning.
  if (date.getDay() === 0 && nextShift === "Morning") {
    nextShift = "Night";
  }

  const year = date.getFullYear();

  const monthName = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "MMMM"
  );

  const dateText = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  const dayShort = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "EEE"
  );

  return {
    date,
    shift: nextShift,
    year,
    monthName,
    dateText,
    dayShort
  };
}

function findCustomerAbove(sheet, leftoverRow) {
  for (let row = leftoverRow; row >= LEFTOVER_START_ROW; row--) {
    const customerName = sheet.getRange(row, 2).getValue();
    const rowType = sheet.getRange(row, 3).getValue();

    if (customerName && rowType === "Order") {
      return customerName;
    }
  }

  return "";
}

function syncLoadingBoardForFile(ss) {
  const ownerSheet = ss.getSheetByName("Owner_List");
  const boardSheet = ss.getSheetByName("Loading_Board");

  boardSheet
    .getRange(START_ROW, 1, END_ROW - START_ROW + 1, SOURCE_COL)
    .clearContent()
    .setFontColor("#000000")
    .setBackground(null);

  let boardRow = START_ROW;

  for (let row = START_ROW; row <= END_ROW; row++) {
    const ownerData = ownerSheet
      .getRange(row, 1, 1, SOURCE_COL)
      .getValues()[0];

    const orderId = ownerData[0];
    const customerName = ownerData[1];

    if (orderId && customerName) {
      boardSheet
        .getRange(boardRow, 1, 1, SOURCE_COL)
        .setValues([ownerData]);

      updateRowStyle(boardSheet, boardRow);
      boardRow++;
    }
  }
}

function updateRowStyle(sheet, row) {
  const status = String(
    sheet.getRange(row, STATUS_COL).getValue()
  ).trim();

  const rowRange = sheet.getRange(
    row,
    1,
    1,
    SOURCE_COL
  ); // A:Q

  rowRange.setFontColor("#000000");

  switch (status) {

    case "Done":
      rowRange.setBackground("#d9ead3"); // Green
      break;

    case "In Progress":
      rowRange.setBackground("#fff2cc"); // Yellow
      break;

    case "Rescheduled":
      rowRange.setBackground("#cfe2f3"); // Blue
      break;

    case "Cancelled":
      rowRange.setBackground("#f4cccc"); // Red
      break;

    case "Leftover":
      rowRange.setBackground("#eadcf8"); // Purple
      break;

    default:
      rowRange.setBackground(null);
      break;
  }
}

function deleteLeftoverRows(orderIdToFind) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leftoverSheet = ss.getSheetByName("Leftover");

  const targetId = normalizeOrderId(orderIdToFind);

  for (let row = leftoverSheet.getLastRow(); row >= LEFTOVER_START_ROW; row--) {
    const orderId = normalizeOrderId(leftoverSheet.getRange(row, 1).getValue());
    const rowType = String(leftoverSheet.getRange(row, 3).getValue()).trim();

    if (orderId === targetId && rowType === "Order") {
      leftoverSheet.deleteRows(row, 3);
      return;
    }
  }
}
function normalizeOrderId(value) {
  if (value instanceof Date) {
    const sheetEpoch = new Date(1899, 11, 30);
    return String(Math.round((value - sheetEpoch) / (1000 * 60 * 60 * 24)));
  }

  return String(value).trim();
}
function refreshAllRowColours() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ["Owner_List", "Loading_Board"].forEach(name => {

    const sheet = ss.getSheetByName(name);

    for (let row = START_ROW; row <= END_ROW; row++) {

      const orderId = sheet.getRange(row, 1).getValue();

      if (orderId) {
        updateRowStyle(sheet, row);
      }
    }
  });
}
function refreshAllRowColoursForFile(ss) {
  ["Owner_List", "Loading_Board"].forEach(name => {
    const sheet = ss.getSheetByName(name);

    for (let row = START_ROW; row <= END_ROW; row++) {
      const orderId = sheet.getRange(row, 1).getValue();

      if (orderId) {
        updateRowStyle(sheet, row);
      }
    }
  });
}
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

  // If the Loading Board is empty, do the full initial build
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const boardSheet = ss.getSheetByName("Loading_Board");
  if (boardSheet) {
    const firstCell = boardSheet.getRange(START_ROW, 1).getValue();
    if (!firstCell) {
      syncLoadingBoard();
    }
  }
}
function syncAllStatuses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ownerSheet = ss.getSheetByName("Owner_List");
  const boardSheet = ss.getSheetByName("Loading_Board");
  const leftoverSheet = ss.getSheetByName("Leftover");

  // Format a single timestamp to use for all automated changes
  const now = Utilities.formatDate(new Date(), "Africa/Lagos", "yyyy-MM-dd HH:mm:ss");

  // --- STEP 1: READ & UPDATE THE TRUTH FROM LOADING BOARD ---
  // We grab up to the Timestamp column (16) so we can update it in one batch [cite: 374, 375]
  const boardRange = boardSheet.getRange(START_ROW, 1, END_ROW - START_ROW + 1, TIMESTAMP_COL);
  const boardData = boardRange.getValues();
  const boardStatusMap = {};
  const expectedLeftovers = [];
  let boardNeedsUpdate = false;

  boardData.forEach(row => {
    const id = String(row[0]).trim();
    let status = String(row[STATUS_COL - 1]).trim(); 
    
    if (id) {
      // Auto-convert "Not Started" (or completely blank cells) to "Leftover"
      if (status === "Not Started" || status === "") {
        status = "Leftover";
        row[STATUS_COL - 1] = status; // Update Status column array [cite: 374]
        row[TIMESTAMP_COL - 1] = now; // Update Timestamp column array [cite: 375]
        boardNeedsUpdate = true;
      }

      boardStatusMap[id] = status; 
      if (status === "Leftover") expectedLeftovers.push(id);
    }
  });

  // Batch write the new "Leftover" statuses back to the Loading Board instantly
  if (boardNeedsUpdate) {
    boardRange.setValues(boardData);
  }

  // --- STEP 2: UPDATE OWNER LIST TO MATCH ---
  for (let r = START_ROW; r <= END_ROW; r++) {
    const id = String(ownerSheet.getRange(r, 1).getValue()).trim();
    
    if (id && boardStatusMap[id]) {
      const currentStatus = String(ownerSheet.getRange(r, STATUS_COL).getValue()).trim();
      
      if (currentStatus !== boardStatusMap[id]) {
        ownerSheet.getRange(r, STATUS_COL).setValue(boardStatusMap[id]);
        ownerSheet.getRange(r, TIMESTAMP_COL).setValue(now); // [cite: 509]
      }
    }
  }

  // --- STEP 3: CLEAN UP LEFTOVER TAB ---
  const existingLeftovers = [];
  const leftoverLastRow = leftoverSheet.getLastRow();

  if (leftoverLastRow >= LEFTOVER_START_ROW) {
    const leftoverData = leftoverSheet.getRange(LEFTOVER_START_ROW, 1, leftoverLastRow - LEFTOVER_START_ROW + 1, 3).getValues(); // [cite: 511]
    leftoverData.forEach(row => {
      const id = String(row[0]).trim();
      const rowType = String(row[2]).trim();
      if (id && rowType === "Order") existingLeftovers.push(id); // [cite: 512]
    });
  }

  // Create missing ones [cite: 513]
  expectedLeftovers.forEach(id => {
    if (!existingLeftovers.includes(id)) createLeftoverRows(id);
  });

  // Delete orphaned ones [cite: 514]
  existingLeftovers.forEach(id => {
    if (!expectedLeftovers.includes(id)) deleteLeftoverRows(id);
  });

  // --- STEP 4: INSTANT COLOR REFRESH ---
  refreshAllRowColours(); // [cite: 515]
  
try {
    SpreadsheetApp.getUi().alert("Sync Complete! 'Not Started' orders were converted to Leftovers, and all tabs are fully updated.");
  } catch (error) {}
}
function refreshAllRowColours() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ["Owner_List", "Loading_Board"].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    const dataRange = sheet.getRange(START_ROW, 1, END_ROW - START_ROW + 1, SOURCE_COL);
    const data = dataRange.getValues();
    
    // Calculate all colors in memory instantly
    const backgrounds = data.map(row => {
      const id = row[0];
      const status = String(row[STATUS_COL - 1]).trim();
      let color = null;
      
      if (id) {
        switch (status) {
          case "Done": color = "#d9ead3"; break;
          case "In Progress": color = "#fff2cc"; break;
          case "Rescheduled": color = "#cfe2f3"; break;
          case "Cancelled": color = "#f4cccc"; break;
          case "Leftover": color = "#eadcf8"; break;
        }
      }
      
      return Array(SOURCE_COL).fill(color);
    });
    
    // Apply all colors in one single server call
    dataRange.setBackgrounds(backgrounds).setFontColor("#000000");
  });
}
function copyProtections(sourceSS, newSS) {
  const sourceSheets = sourceSS.getSheets();
  
  sourceSheets.forEach(sourceSheet => {
    const destSheet = newSS.getSheetByName(sourceSheet.getName());
    if (!destSheet) return;

    // 1. Copy Sheet-Level Protections (Tab Locks)
    const sheetProtections = sourceSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    sheetProtections.forEach(p => {
      const newProtection = destSheet.protect().setDescription(p.getDescription());
      newProtection.setWarningOnly(p.isWarningOnly());
      
      if (!p.isWarningOnly()) {
        // Copy exact editors
        const editors = p.getEditors().map(user => user.getEmail());
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
        newProtection.removeEditors(newProtection.getEditors());
        if (editors.length > 0) newProtection.addEditors(editors);
      }
    });
  });
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
  const ownerSheet = ss.getSheetByName("Owner_List");
  const leftoverSheet = ss.getSheetByName("Leftover");

  // 2. Wipe Statuses and Timestamps on the Owner List
  const ownerRange = ownerSheet.getRange(START_ROW, 1, END_ROW - START_ROW + 1, TIMESTAMP_COL);
  const ownerData = ownerRange.getValues();
  
  for (let i = 0; i < ownerData.length; i++) {
    const orderId = String(ownerData[i][0]).trim();
    if (orderId) { 
      ownerData[i][STATUS_COL - 1] = "Not Started"; // Reset Status
      ownerData[i][TIMESTAMP_COL - 1] = "";         // Erase Timestamp
    }
  }
  ownerRange.setValues(ownerData); 
  
  // --- FIX: VISUALLY RESET THE OWNER LIST COLORS ---
  ownerSheet.getRange(START_ROW, 1, END_ROW - START_ROW + 1, SOURCE_COL)
    .setBackground(null)
    .setFontColor("#000000");

  // 3. Erase the Leftover Tab Completely
  const maxLeftoverRows = leftoverSheet.getMaxRows();
  if (maxLeftoverRows >= LEFTOVER_START_ROW) {
    const leftoverRange = leftoverSheet.getRange(LEFTOVER_START_ROW, 1, maxLeftoverRows - LEFTOVER_START_ROW + 1, 16);
    
    leftoverRange
      .clearContent()
      .clearDataValidations()
      .setBackground(null)
      .setBorder(false, false, false, false, false, false)
      .setFontColor("#000000")
      .setFontWeight("normal");
  }

  // 4. Run the Universal Merge to Repaint the Loading Board
  syncLoadingBoard();

  ui.alert("✅ Reset Complete! All statuses are cleared, colors are reset, and leftovers have been removed.");
}
function syncSingleRowToBoard(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ownerSheet = ss.getSheetByName("Owner_List");
  const boardSheet = ss.getSheetByName("Loading_Board");

  // Grab just the single edited row
  const rowData = ownerSheet.getRange(rowIndex, 1, 1, SOURCE_COL).getValues()[0];
  
  const orderId = rowData[0];
  const customerName = rowData[1];
  const status = rowData[STATUS_COL - 1];
  const timestamp = rowData[TIMESTAMP_COL - 1];
  const source = rowData[SOURCE_COL - 1];

  // If the row was deleted/cleared, clear it on the board too
  if (!orderId || !String(customerName).trim()) {
     boardSheet.getRange(rowIndex, 1, 1, SOURCE_COL).clearContent().setBackground(null);
     return;
  }

  // Clean zeroes
  const cleanQuantities = rowData.slice(2, 14).map(v => 
    (Number(v) === 0 || v === "") ? "" : Number(v)
  );

  const finalRow = [
    [
      Number(orderId),
      customerName,
      ...cleanQuantities,
      status || "Not Started",
      timestamp || "",
      source || "Manual"
    ]
  ];

  // Pick color
  let color = null;
  switch (status) {
    case "Done": color = "#d9ead3"; break;
    case "In Progress": color = "#fff2cc"; break;
    case "Rescheduled": color = "#cfe2f3"; break;
    case "Cancelled": color = "#f4cccc"; break;
    case "Leftover": color = "#eadcf8"; break;
  }

  // Paste to the EXACT same row on the Loading Board
  const targetRange = boardSheet.getRange(rowIndex, 1, 1, SOURCE_COL);
  targetRange.setValues(finalRow);
  targetRange.setBackgrounds([Array(SOURCE_COL).fill(color)]);
}
