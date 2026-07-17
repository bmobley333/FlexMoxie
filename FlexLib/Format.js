/* global SpreadsheetApp */
/* exported fApplyVisualIsolation */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Visual Formatting and Isolation Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fApplyVisualIsolation
   Purpose: Programmatically formats and freezes Row 1 and Column A to isolate Jodar's coordinate registers.
   Assumptions: None.
   Notes: Applies the Nord-dark slate theme and adds conditional formatting for empty tag cells.
   @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to isolate.
   @returns {void}
*/
function fApplyVisualIsolation(sheet) {
  if (!sheet) return;

  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();

  // 1. Freeze Row 1 and Column A
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  // 2. Style Row 1 (Header/Column tags)
  // Text format: bold, centered, Nord slate background, light gray text
  const row1Range = sheet.getRange(1, 1, 1, maxCols);
  row1Range.setBackground('#3B4252'); // Nord Slate Grey
  row1Range.setFontColor('#ECEFF4'); // Nord Snow White
  row1Range.setFontWeight('bold');
  row1Range.setHorizontalAlignment('center');

  // 3. Style Column A (Row tags)
  // Text format: bold, italic, left-aligned, Nord charcoal background
  const colARange = sheet.getRange(1, 1, maxRows, 1);
  colARange.setBackground('#2E3440'); // Darker Nord Charcoal
  colARange.setFontColor('#ECEFF4'); // Nord Snow White
  colARange.setFontWeight('bold');
  colARange.setFontStyle('italic');
  colARange.setHorizontalAlignment('left');

  // 4. Add Conditional Formatting to highlight empty cells in Column A and Row 1
  const emptyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground('#FFD2D2') // Muted warning red
    .setFontColor('#900000')
    .setRanges([sheet.getRange(1, 1, 1, maxCols), sheet.getRange(1, 1, maxRows, 1)])
    .build();

  const currentRules = sheet.getConditionalFormatRules();
  const newRules = [];

  for (let i = 0; i < currentRules.length; i++) {
    const r = currentRules[i];
    const cond = r.getBooleanCondition();
    if (cond) {
      const criteria = cond.getCriteriaType();
      const bg = cond.getBackground();
      // Skip/remove our existing empty-cell warning rule to prevent accumulation
      if (criteria === SpreadsheetApp.BooleanCriteria.CELL_EMPTY && bg && bg.toLowerCase() === '#ffd2d2') {
        continue;
      }
    }
    newRules.push(r);
  }

  newRules.push(emptyRule);
  sheet.setConditionalFormatRules(newRules);
} // End function fApplyVisualIsolation
