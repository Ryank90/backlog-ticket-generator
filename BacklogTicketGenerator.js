/**
 * -----------------------------------------------------------------------------
 * Backlog Ticket Generator Script.
 *
 * The following script is included into the Template: Product Backlog Google
 * spreadsheet to provide new menu items. These menu items allow for the
 * printing of physical cards to create a board of items and pushing of items
 * into JIRA for high-level epic management.
 * -----------------------------------------------------------------------------
 */

/**
 * Runs when the sheet is loaded.
 */
function onOpen() {
    var ui = SpreadsheetApp.getUi();

    // Include the Card options menu to the Google Sheet.
    ui.createMenu('Card Generator')
        .addItem('Generate Items', 'genItemsFromBacklog')
        .addItem('Generate Specific Items', 'genSpecificItemsFromBacklog')
        .addToUi();

    // Include the JIRA options menu to the Google Sheet.
    ui.createMenu('JIRA Options')
        .addItem('Generate Items in JIRA', 'genItemsInJira')
        .addItem('Generate Specific Items in JIRA', 'genSpecificItemsInJira')
        .addToUi();
}

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
 * Generate items from the backlog within the document.
 */
function genItemsFromBacklog() {
  if (!validateTabExists('Items', 1)) {
    return;
  }

  Browser.msgBox("We look good to process");
}

/**
 * Generate specific items from the backlog within the document.
 */
function genSpecificItemsFromBacklog() {
  if (!validateTabExists('Items', 1)) {
    return;
  }

  Browser.msgBox("We look good to process");
}

/**
 * Generate items from the backlog in JIRA.
 */
function genItemsInJira() {

}

/**
 * Generate specific items from the backlog in JIRA.
 */
function genSpecificItemsInJira() {

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

/**
 *
 */
function generateCards(items) {

}
