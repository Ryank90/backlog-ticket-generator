

function getTemplateArea() {
    return "A1:F10";
}

/**
 * ==================================================
 * SHEET METHODS
 * ==================================================
 */

function getSheetInstance() {
    return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetTabByName(name) {
    return getSheetInstance().getSheetByName(name)
}

function getPreparedItemSheet(template, itemCount, rowCount) {
    var neededRows = itemCount * rowCount;
    var sheet = getSheetTabByName("Tickets");

    sheet.clear();

    setColWidthTo(sheet, "Template", template);

    var rows = sheet.getMaxRows();

    if (rows < neededRows) {
        sheet.insertRows(1, (neededRows - rows));
    }

    setRowHeightTo(sheet, "Template", rowCount, itemCount);

    return sheet;
}

/**
 * ==================================================
 * HELPER METHODS
 * ==================================================
 */

function setColWidthTo(sheet, name, range) {
    var template = getSheetTabByName(name);
    var max = range.getLastColumn() + 1;
    for (var i = 1; i < max; i++) {
        var currentWidth = template.getColumnWidth(i);
        sheet.setColumnWidth(i, currentWidth);
    }
}

function setRowHeightTo(sheet, name, rowCount, itemCount) {
    var template = getSheetTabByName(name);
    for (var i = 0; i < rowCount; i++) {
        for (var j = 1; j < (rowCount + 1); j++) {
            var currentRow = (i * rowCount) + j;
            var currentHeight = template.getRowHeight(j);

            sheet.setRowHeight(currentRow, currentHeight);
        }
    }
}

/**
 * ==================================================
 * RANGE METHODS
 * ==================================================
 */

function getTemplateRange(name) {
    return getSheetTabByName(name).getRange(getTemplateArea());
}

function getHeaderRange(items) {
    return items.getRange(1, 1, 1, items.getLastColumn());
}

function getItemsRange(items) {
    var rowCount = items.getLastRow() - 1;
    return items.getRange(2, 1, rowCount, items.getLastColumn());
}

function getSelectedItemRange(items) {
    var range = getSheetInstance().getActiveRange();
    var rowStart = range.getRowIndex();
    var rowCount = range.getNumRows();

    if (rowStart < 2) {
        rowStart = 2;
        rowCount = (rowCount > 1 ? rowCount - 1 : rowCount);
    }

    return items.getRange(rowStart, 1, rowCount, items.getLastColumn());
}

/**
 * ==================================================
 * TEMPLATE METHODS
 * ==================================================
 */

/**
 * Template: Set "ID" of the item.
 */
function setItemId(backlog, item) {

}

/**
 * Template: Set "Name" of the item.
 */
function setItemName(backlog, item) {

}

/**
 * Template: Set "Story" of the item.
 */
function setItemStory(backlog, item) {

}

/**
 * Template: Set "How To Demo" of the item.
 */
function setItemHowToDemo(backlog, item) {

}

/**
 * Template: Set "Priority" of the item.
 */
function setItemPriority(backlog, item) {

}

/**
 * Template: Set "Estimate" of the item.
 */
function setItemEstimate(backlog, item) {

}

/**
 *
 */
function getItemStartCol() {

}

/**
 *
 */
function getItemStartRow() {

}

/**
 *
 */
function getItemLastCol() {

}

/**
 *
 */
function getItemLastRow() {

}


/**
 * Runs when the sheet is loaded.
 */
function onOpen() {
    var ui = SpreadsheetApp.getUi();

    // Include the Card options menu to the Google Sheet.
    ui.createMenu('Card Generator')
        .addItem('Generate Cards', 'genItemsFromBacklog')
        .addItem('Generate Specific Cards', 'genSpecificItemsFromBacklog')
        .addToUi();

    // Include the JIRA options menu to the Google Sheet.
    ui.createMenu('JIRA Options')
        .addItem('Push All Items to JIRA', '')
        .addItem('Push Specific Items to JIRA', '')
        .addToUi();
}

/**
 *
 */
function genItemsFromBacklog() {
  if (!validateTabExists('Items', 1)) {
    return;
  }

  Browser.msgBox("We look good to process");
}

/**
 *
 */
function genSpecificItemsFromBacklog() {
  if (!validateTabExists('Items', 1)) {
    return;
  }

  Browser.msgBox("We look good to process");
}

/**
 * Validating if a specific tab exists in the Google Spreadsheet.
 */
function validateTabExists(name, position) {
  if (getSheetTabByName(name) == null) {
    getSheetInstance().insertSheet(name, position);
    Browser.msgBox('The (' + name + ') sheet was missing and now has been included below. Please try again.');
    return false;
  }

  return true;
}
