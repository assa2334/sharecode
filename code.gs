
const TARGET_SPREADSHEET_ID = '1YvwAWM8TPK2g4-5aAjAgx6v3AY90pK_y6KJemnUz2tY';


/**
 * Adds custom menu when spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Manage')
    .addItem('Manage User', 'showManageDialog')
    .addToUi();
}


/**
 * Open Manage User dialog.
 */
function showManageDialog() {
  const html = HtmlService
    .createHtmlOutputFromFile('Manage')
    .setWidth(750)
    .setHeight(550);

  SpreadsheetApp.getUi().showModalDialog(html, 'Manage User');
}


/**
 * Get all sheets from target spreadsheet.
 */
function getActiveTabs() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);

  return ss.getSheets()
    .filter(sheet => !sheet.isSheetHidden())
    .map(sheet => ({
      id: sheet.getSheetId(),
      name: sheet.getName()
    }));
}


/**
 * Create new tab and give user access.
 */
function createUserTab(tabName, email) {

  if (!tabName || !email) {
    throw new Error('Tab name and email are required.');
  }

  email = email.trim().toLowerCase();

  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);

  // Prevent duplicate tab name
  if (ss.getSheetByName(tabName)) {
    throw new Error('A tab with this name already exists.');
  }

  // Create tab
  const sheet = ss.insertSheet(tabName);

  /*
   * Protect the tab.
   */
  const protection = sheet.protect();

  protection.setDescription(
    'Managed tab for ' + email
  );

  /*
   * Remove existing editors from protection.
   */
  protection.removeEditors(
    protection.getEditors()
  );

  /*
   * Add only requested user to protected range.
   */
  protection.addEditor(email);

  /*
   * Keep domain restrictions if applicable.
   */
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }

  /*
   * Give user spreadsheet-level edit access.
   *
   * IMPORTANT:
   * Google Sheets sharing is spreadsheet-level.
   */
  DriveApp.getFileById(TARGET_SPREADSHEET_ID)
    .addEditor(email);

  /*
   * Optional header in the new tab.
   */
  sheet.getRange('A1').setValue('User');
  sheet.getRange('B1').setValue(email);

  sheet.getRange('A2').setValue('Created');
  sheet.getRange('B2').setValue(new Date());

  /*
   * Send email.
   */
  MailApp.sendEmail({
    to: email,
    subject: 'Your Google Sheet tab has been created',
    htmlBody:
      '<p>Hello,</p>' +
      '<p>A new tab <b>' + escapeHtml_(tabName) + '</b> has been created for you.</p>' +
      '<p>You can open the spreadsheet here:</p>' +
      '<p><a href="' + ss.getUrl() + '">' +
      ss.getUrl() +
      '</a></p>' +
      '<p>You have been given edit access to your protected tab.</p>'
  });

  return {
    success: true,
    tabName: tabName,
    email: email,
    url: ss.getUrl()
  };
}


/**
 * Delete a tab and remove the user's spreadsheet access.
 */
function deleteUserTab(tabName, email) {

  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(tabName);

  if (!sheet) {
    throw new Error('Tab not found.');
  }

  /*
   * Get email from tab if email isn't supplied.
   */
  if (!email) {
    email = getProtectedUserEmail_(sheet);
  }

  /*
   * Don't allow deleting the final sheet.
   */
  if (ss.getSheets().length <= 1) {
    throw new Error('Cannot delete the only sheet.');
  }

  /*
   * Delete tab.
   */
  ss.deleteSheet(sheet);

  /*
   * Remove spreadsheet access.
   */
  if (email) {
    try {
      DriveApp.getFileById(TARGET_SPREADSHEET_ID)
        .removeEditor(email);
    } catch (err) {
      console.log('Could not remove editor: ' + err);
    }
  }

  return {
    success: true,
    tabName: tabName,
    email: email
  };
}


/**
 * Try to find the user assigned to a protected sheet.
 */
function getProtectedUserEmail_(sheet) {

  const protections = sheet.getProtections(
    SpreadsheetApp.ProtectionType.SHEET
  );

  if (!protections.length) {
    return '';
  }

  const editors = protections[0].getEditors();

  if (!editors.length) {
    return '';
  }

  return editors[0].getEmail();
}


/**
 * Escape HTML.
 */
function escapeHtml_(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
