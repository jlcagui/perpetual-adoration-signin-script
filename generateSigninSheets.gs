/**
 * generateAdorationSignInPages
 * 
 * This function automates the creation of daily sign-in sheets for Perpetual Adoration
 * in Google Sheets. It duplicates a template block multiple times with a one-row gap
 * and sets the correct date for each day.
 */
function generateAdorationSignInPages() {
  // Get the currently active sheet in the active spreadsheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // --- Template configuration ---
  var templateStartRow = 1;   // Row where the template starts
  var templateRows = 41;      // Number of rows in the template
  var templateCols = 8;       // Number of columns in the template
  var numPages = 59;          // Total number of days/pages to generate
  var gapRows = 1;            // Number of blank rows between templates

  // Get the range of the template (A1:H41)
  var templateRange = sheet.getRange(
    templateStartRow,  // start row
    1,                 // start column (A)
    templateRows,      // number of rows
    templateCols       // number of columns
  );

  // --- Loop to create each page after the first ---
  // Start from 1 because the first template is already in place
  for (var i = 1; i < numPages; i++) {
    // Calculate the starting row for this new block including the gap
    var pasteRow = i * (templateRows + gapRows) + 1;

    // Check if there are enough rows in the sheet; if not, add more
    var requiredRows = pasteRow + templateRows - sheet.getMaxRows();
    if (requiredRows > 0) {
      // Insert extra rows at the bottom of the sheet as needed
      sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows);
    }

    // Get the target range where the template will be copied
    var targetRange = sheet.getRange(pasteRow, 1, templateRows, templateCols);

    // Copy the template into the target range
    // contentsOnly: false → copies values, formatting, merged cells, and formulas
    templateRange.copyTo(targetRange, { contentsOnly: false });

    // Set the date in the second cell of the block (B2 relative to template)
    var dateCell = sheet.getRange(pasteRow + 1, 2); 
    dateCell.setValue(new Date(2026, 1, 1 + i)); // Months are 0-indexed: 1 = February
  }
}
