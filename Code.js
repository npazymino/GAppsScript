// Global variable to store admin status
var isAdmin = false;
const ADMIN_CODE = "PROSPR123"; // This should be declared ONLY ONCE in your entire project

/**
 * Runs when the spreadsheet is opened.
 * Creates the custom "Admin" menu.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('Admin');

  if (isAdmin) {
    // If already authenticated, show full menu
    addAdminMenuItems(menu);
  } else {
    // Show only the "Unlock Admin" option initially
    menu.addItem('Unlock Admin', 'showAdminPrompt');
  }
  menu.addToUi();
}

/**
 * Prompts the user for the admin code.
 */
function showAdminPrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Admin Access',
    'Please enter the admin code:',
    ui.ButtonSet.OK_CANCEL
  );

  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.OK) {
    if (response.getResponseText() === ADMIN_CODE) {
      isAdmin = true;
      SpreadsheetApp.getUi().alert('Success', 'Admin access granted!', ui.ButtonSet.OK);
      // Rebuild the menu to show admin options
      onOpen();
    } else {
      SpreadsheetApp.getUi().alert('Error', 'Incorrect admin code.', ui.ButtonSet.OK);
    }
  }
}

/**
 * Adds the admin-only menu items.
 * @param {GoogleAppsScript.Ui.Menu} menu The menu to add items to.
 */
function addAdminMenuItems(menu) {
  menu.addItem('Recalculate all values', 'recalculateValues');
  // Add other admin-only functions here if needed
}
