/**
 * Custom function to automatically resize the selected column(s) 
 * to fit data plus an extra margin percentage.
 */
function autoResizeWithMargin() {
  const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  const SHEET = SPREADSHEET.getActiveSheet();
  const UI = SpreadsheetApp.getUi();

  // --- CONFIGURATION ---
  // Define the extra margin you want to add (e.g., 0.10 for 10%)
  const MARGIN_PERCENTAGE = 0.15; 
  // ---------------------
  
  // Get the selected column index(es)
  const selection = SPREADSHEET.getSelection();
  const activeRange = selection.getActiveRange();
  
  // Check if a range is selected
  if (!activeRange) {
    UI.alert('Error', 'Please select the column(s) you wish to resize first.', UI.ButtonSet.OK);
    return;
  }
  
  // Get the start and end column numbers
  const firstCol = activeRange.getColumn();
  const lastCol = activeRange.getLastColumn();
  
  // Loop through each selected column
  for (let colIndex = firstCol; colIndex <= lastCol; colIndex++) {
    
    // 1. Temporarily auto-resize the column to find the minimum 'fit to data' width.
    SHEET.autoResizeColumn(colIndex); 
    SpreadsheetApp.flush(); // Forces the UI updates to be calculated
    
    // 2. Get the calculated 'fit to data' width.
    const fitWidth = SHEET.getColumnWidth(colIndex);
    
    // 3. Calculate the new width with the margin.
    // New Width = Fit Width * (1 + Margin Percentage)
    const newWidth = Math.ceil(fitWidth * (1 + MARGIN_PERCENTAGE));
    
    // 4. Set the final width.
    SHEET.setColumnWidth(colIndex, newWidth);
  }

  UI.alert('Success', `Column(s) ${SHEET.getRange(activeRange.getRow(), firstCol).getA1Notation().replace(/\d/g,'')}:${SHEET.getRange(activeRange.getRow(), lastCol).getA1Notation().replace(/\d/g,'')} resized with a ${MARGIN_PERCENTAGE * 100}% margin.`, UI.ButtonSet.OK);
}

// Add a custom menu to make the script easy to run
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Custom Resize')
      .addItem('Resize with 10% Margin', 'autoResizeWithMargin')
      .addToUi();
}
