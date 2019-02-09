/**
 * -----------------------------------------------------------------------------
 * Backlog JIRA Generator Script.
 *
 * Description in here...
 * -----------------------------------------------------------------------------
 */

/**
 *
 */
function onOpen() {
    var ui = SpreadsheetApp.getUi();

    ui.createMenu('JIRA Options')
        .addItem('Generate Items in JIRA', 'genItemsInJira')
        .addItem('Generate Specific Items in JIRA', 'genSpecificItemsInJira')
        .addToUi();
}

/**
 *
 */
function genItemsInJira() {

}

/**
 *
 */
function genSpecificItemsInJira() {

}
