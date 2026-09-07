// @ts-nocheck
// ============================================================
// Web App: New Order entry form for Owner_List
//
// Deploy via Deploy > New deployment > Web app. The doGet()
// entry point below serves OrderForm.html, which lets staff add
// a new order to the Owner_List sheet (the first tab) from a
// phone or browser without opening the spreadsheet itself.
//
// Reuses START_ROW, END_ROW, STATUS_COL, TIMESTAMP_COL,
// SOURCE_COL, ORDER_COLS, generateOrderId() and
// syncSingleRowToBoard() from Code.gs so a web-app submission
// behaves exactly like typing a new row into Owner_List by hand.
// ============================================================

const OWNER_SHEET_NAME = "Owner_List";

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('OrderForm')
    .setTitle('Loading List — New Order')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Reads the bread-quantity columns (C:N) straight from the
 * Owner_List header rows, so the form always matches whatever
 * brands/codes are currently on the sheet.
 * Returns [{ code, group, column }, ...] in sheet column order.
 */
function getBreadColumns_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(OWNER_SHEET_NAME);

  const numBreadCols = ORDER_COLS - 2; // A=Order ID, B=Customer, rest are bread codes
  const groupRow = sheet.getRange(5, 3, 1, numBreadCols).getValues()[0]; // C5:N5
  const codeRow = sheet.getRange(6, 3, 1, numBreadCols).getValues()[0]; // C6:N6

  let lastGroup = "";
  return codeRow.map((code, i) => {
    const groupValue = String(groupRow[i] || "").trim();
    if (groupValue) lastGroup = groupValue;

    return {
      code: String(code).trim(),
      group: lastGroup,
      column: 3 + i // C = 3
    };
  });
}

/** Pulls the customer dropdown's allowed values straight off B8's data validation. */
function getCustomerList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(OWNER_SHEET_NAME);

  const rule = sheet.getRange(START_ROW, 2).getDataValidation();
  if (!rule) return [];

  const criteria = rule.getCriteriaType();
  const values = rule.getCriteriaValues();

  if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return values[0];
  }

  if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    return values[0].getValues().flat().filter(String);
  }

  return [];
}

/** Called from OrderForm.html on load to build the form. */
function getFormMeta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(OWNER_SHEET_NAME);

  const date = sheet.getRange("A1").getValue();
  const shift = sheet.getRange("B1").getValue();

  return {
    date: date ? Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy") : "",
    shift: shift || "",
    customers: getCustomerList_(),
    breadColumns: getBreadColumns_()
  };
}

/**
 * Called from OrderForm.html on submit.
 * payload = { customerName: string, quantities: { <code>: number, ... } }
 */
function submitOrder(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(OWNER_SHEET_NAME);

    const customerName = String((payload && payload.customerName) || "").trim();
    if (!customerName) {
      throw new Error("Please choose a customer.");
    }

    const columns = getBreadColumns_();
    const rawQuantities = (payload && payload.quantities) || {};

    const quantities = columns.map(col => {
      const raw = rawQuantities[col.code];
      const num = Number(raw);
      return raw === "" || raw === undefined || raw === null || isNaN(num) || num === 0 ? "" : num;
    });

    if (!quantities.some(q => q !== "")) {
      throw new Error("Enter at least one bread quantity.");
    }

    // Find the next blank row in the data range (matches how a
    // person typing a new order into the sheet finds a free row).
    let targetRow = -1;
    for (let row = START_ROW; row <= END_ROW; row++) {
      const orderId = sheet.getRange(row, 1).getValue();
      const existingCustomer = sheet.getRange(row, 2).getValue();
      if (!orderId && !existingCustomer) {
        targetRow = row;
        break;
      }
    }

    if (targetRow === -1) {
      throw new Error("Owner_List is full — there are no empty rows left for a new order.");
    }

    // Customer name + bread quantities (B:N)
    sheet.getRange(targetRow, 2, 1, quantities.length + 1).setValues([[customerName, ...quantities]]);
    sheet.getRange(targetRow, SOURCE_COL).setValue("Manual");

    // Assign the Order ID last, same as installedOnEdit does for a new row.
    const newOrderId = generateOrderId();
    sheet.getRange(targetRow, 1).setValue(newOrderId);

    syncSingleRowToBoard(targetRow);

    return {
      success: true,
      orderId: newOrderId,
      row: targetRow
    };
  } finally {
    lock.releaseLock();
  }
}
