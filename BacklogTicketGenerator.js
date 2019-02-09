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

    ui.createMenu('Card Generator')
        .addItem('Generate Items', 'genItemsFromBacklog')
        .addItem('Generate Specific Items', 'genSpecificItemsFromBacklog')
        .addToUi();
}

/**
 * Get the area in which the template is set.
 */
function getTemplateArea() {
    return "A1:F10";
}

/**
 * Get an instance of the Google sheet.
 */
function getSheetInstance() {
    return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Get a tab/sheet within the Google sheet.
 */
function getSheetTabByName(name) {
    return getSheetInstance().getSheetByName(name)
}

/**
 *
 */
function getPreparedItemSheet(template, itemCount, rowCount) {
    var neededRows = itemCount * rowCount;
    var sheet = getSheetTabByName("Items");

    sheet.clear();

    setColWidthTo(sheet, "_Template", template);

    var rows = sheet.getMaxRows();

    if (rows < neededRows) {
        sheet.insertRows(1, (neededRows - rows));
    }

    setRowHeightTo(sheet, "_Template", rowCount, itemCount);

    return sheet;
}

/**
 * Set the sheets column width.
 */
function setColWidthTo(sheet, name, range) {
    var template = getSheetTabByName(name);
    var max = range.getLastColumn() + 1;
    for (var i = 1; i < max; i++) {
        var currentWidth = template.getColumnWidth(i);
        sheet.setColumnWidth(i, currentWidth);
    }
}

/**
 * Set the sheets row height.
 */
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
 * Get the range of a specific sheet.
 */
function getTemplateRange(name) {
    return getSheetTabByName(name).getRange(getTemplateArea());
}

/**
 * Get the range of the header.
 */
function getHeaderRange(items) {
    return items.getRange(1, 1, 1, items.getLastColumn());
}

/**
 * Get the range of the items.
 */
function getItemsRange(items) {
    var rowCount = items.getLastRow() - 1;
    return items.getRange(2, 1, rowCount, items.getLastColumn());
}

/**
 * Get the range of the selected items within the Google sheet.
 */
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
 * Template: Set "Name" of the item.
 */
function setItemName(backlogItem, item) {
    var maxLength = 30;
    var name = backlogItem['Name (Epic)'];

    if (name && name.length > maxLength) {
        name = name.substring(0, maxLength) + '...';
    }

    item.getCell(3, 3).setValue(name);
}

/**
 * Template: Set "ID" of the item.
 */
function setItemId(backlogItem, item) {
    item.getCell(2, 3).setValue(backlogItem['ID']);
}

/**
 * Template: Set "Theme" of the item.
 */
function setItemTheme(backlogItem, item) {
    var maxLength = 12;
    var theme = backlogItem['Theme'];

    if (theme && theme.length > maxLength) {
        theme = theme.substring(0, maxLength) + '...';
    }

    item.getCell(2, 5).setValue(theme);
}

/**
 * Template: Set "Story" of the item.
 */
function setItemStory(backlogItem, item) {
    item.getCell(5, 3).setValue(backlogItem['User Story']);
}

/**
 * Template: Set "How To Demo" of the item.
 */
function setItemHowToDemo(backlogItem, item) {
    item.getCell(8, 3).setValue(backlogItem['How to Demo']);
}

/**
 * Template: Set "Priority" of the item.
 */
function setItemPriority(backlogItem, item) {
    if (backlogItem['Priority'] == '' || backlogItem['Priority'] == 'undefined') {
         backlogItem['Priority'] = '';
    }
    item.getCell(5, 5).setValue(backlogItem['Priority']);
}

/**
 * Template: Set "Estimate" of the item.
 */
function setItemEstimate(backlogItem, item) {
    if (backlogItem['Estimate'] == '' || backlogItem['Estimate'] == 'undefined') {
         backlogItem['Estimate'] = '';
    }
    item.getCell(8, 5).setValue(backlogItem['Estimate']);
}

/**
 * Get the start column of an item.
 */
function getItemStartCol() {
    return getTemplateArea().substring(0, 1);
}

/**
 * Get the start row of an item.
 */
function getItemStartRow() {
    return parseInt(getTemplateArea().substring(1, 2), 10);
}

/**
 * Get the last column of an item.
 */
function getItemLastCol() {
    return getTemplateArea().substring(3, 4);
}

/**
 * Get the last row of an item.
 */
function getItemLastRow() {
    return parseInt(getTemplateArea().substring(4), 10);
}

/**
 * Get the product backlog items in the correct format.
 */
function getProductBacklogItems(selectedItems) {
    var productBacklog = getSheetTabByName("Product Backlog");
    var rangeRows = (selectedItems ? getSelectedItemRange(productBacklog) : getItemsRange(productBacklog));
    var rows = rangeRows.getValues();
    var headers = getHeaderRange(productBacklog).getValues()[0];

    var productBacklogItems = [];
    for (var i = 0; i < rows.length; i++) {
        var productBacklogItem = {};
        for (var j = 0; j < rows[i].length; j++) {
            productBacklogItem[headers[j]] = rows[i][j];
        }
        productBacklogItems.push(productBacklogItem);
    }

    return productBacklogItems;
}

/**
 * Generate cards that can be printed from the backlog.
 */
function generateCards(items) {
    var rowsCount = getItemLastRow();
    var template = getTemplateRange("_Template");
    var tab = getPreparedItemSheet(template, items.length, rowsCount);

    var rowStart = getItemStartRow();
    var rowLast = getItemLastRow();

    var colStart = getItemStartCol();
    var colLast = getItemLastCol();

    for (var i = 0; i < items.length; i++) {
        var rangeVal = colStart + rowStart + ':' + colLast + rowLast;
        var card = tab.getRange(rangeVal);

        template.copyTo(card);

        setItemId(items[i], card);
        setItemTheme(items[i], card);
        setItemName(items[i], card);
        setItemStory(items[i], card);
        setItemHowToDemo(items[i], card);
        setItemEstimate(items[i], card);
        setItemPriority(items[i], card);

        rowStart += rowsCount;
        rowLast += rowsCount;
    }

    Browser.msgBox("Completed!");
}

/**
 * Generate items from the backlog within the document.
 */
function genItemsFromBacklog() {
  if (!validateTabExists('Items', 1)) {
    return;
  }

  var items = getProductBacklogItems(false);
  generateCards(items);

}

/**
 * Generate specific items from the backlog within the document.
 */
function genSpecificItemsFromBacklog() {
  if (!validateTabExists('Items', 1)) {
    return;
  }

  if (getSheetTabByName("Product Backlog").getName() != SpreadsheetApp.getActiveSheet().getName()) {
      Browser.msgBox('The Backlog sheet need to be active when creating cards from selected rows. Please try again.');
      return;
  }

  var items = getProductBacklogItems(true);
  generateCards(items);
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
