/********************************************************************
 *
 * Author: Robin O'Connell
 * Date:   22 July 2020
 *
 * Provides a utility for quickly opening students' Message pages.
 *
 ********************************************************************/

/**
 * Adds Open Message Pages to Add-ons menu when document is opened.
 * If onOpen() is used by another script, delete this function and
 * add the call to openMessagePages_addMenuItem the other script.
 */
function onOpen() {
  openMessagePages_addMenuItem();
}


/**
 * Adds Open Message Pages to Add-ons menu.
 */
function openMessagePages_addMenuItem() {
  const ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addItem('Open Message Pages', 'openMessagePages')
    .addToUi();
}

/**
 * Finds all strings in the first column of the selected range that look like
 * Hackingtons usernames, constructs the URLs of the message pages for those
 * users, and opens those pages in new tabs/windows.
 *
 * @return {Object[]} list of values of all selected cells in the active sheet
 */
function getAllSelectedCellValues() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const selection = sheet.getSelection();
  const ranges = selection.getActiveRangeList().getRanges();
  const values = [];
  
  for(let range of ranges) {
    const rangeValues = range.getValues().flat();
    values.push.apply(values, rangeValues); // In place array concatenation
  }
  
  return values;
}

/**
 * Finds all strings in the selected cells that look like Hackingtons usernames,
 * constructs the URLs of the message pages for those users, and opens those
 * pages in new tabs/windows.
 */
function openMessagePages() {
  // Most usernames are of the form First or FirstL, perhaps with a number suffix.
  // At least one username contains a dash (Kal-El). Some have last names with more
  // than one letter (e.g. Le) and some have other capitalized letters (CeAnnaB).
  // Allows for a number suffix from 1-99. Also allows leading/trailing spaces.
  const usernamePattern = /^ *([A-Z][a-z]*-?)+([1-9]0|[1-9]?[1-9])? *$/;
  const baseUrl = "https://hackingtons.io/teacherCommunication/";
  
  const cellValues = getAllSelectedCellValues();
  const usernames = cellValues.filter(
    val => (typeof val === "string" && usernamePattern.test(val.trim()))
  );
  const urls = usernames.map(val => baseUrl + val);
  
  openURLs(urls);
}

/**
 * Opens one or more URLs in new tabs/windows by simulating ctrl + left click events
 * in a modal dialog box.
 * Adapted from code by Stephen M. Harris. See https://stackoverflow.com/a/47098533
 *
 * @param {string[]} urls - a list of urls to open
 */
function openURLs(urls){
  if(urls.length == 0) return;
  
  const html = HtmlService.createHtmlOutput(
      '<html><head><script>'
    + 'const urls = ["' + urls.join('","') + '"];'
    + 'for(let url of urls) {'
    + '  var a = document.createElement("a");'
    + '  a.href = url;'
    + '  a.target = "_blank";'
    + '  var event = new MouseEvent("click", {"ctrlKey": true});'
    + '  a.dispatchEvent(event);'
    + '}'
    + 'google.script.host.close();'
    + '</script></head>'
    // Offer URL as clickable link in case above code fails.
    + '<body style="word-break: break-word; font-family: sans-serif;">'
    + '  Failed to open automatically. <a href="' + urls[0] + '" target="_blank" onclick="google.script.host.close()">Click here to proceed</a>.'
    + '  <script>google.script.host.setHeight(40); google.script.host.setWidth(410)</script>'
    + '</body></html>')
    .setWidth(90)
    .setHeight(1);
  
  SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
}
