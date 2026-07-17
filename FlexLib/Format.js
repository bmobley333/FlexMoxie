/* global SpreadsheetApp, fNormalizeTags */
/* exported fApplyVisualIsolation */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Visual Formatting and Isolation Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fApplyVisualIsolation
   Purpose: Programmatically formats and freezes Row 1 / Header and Column A to isolate Jodar's coordinate registers.
   Assumptions: None.
   Notes: Applies the Nord-dark slate theme, auto-compiles A1 notes (Header: True/False), and adds conditional warning formatting.
   @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to isolate.
   @returns {void}
*/
function fApplyVisualIsolation(sheet) {
  if (!sheet) return;

  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();

  // 1. Read Cell A1 metadata note
  const metadata = fGetA1Metadata(sheet);
  
  // 2. Resolve Header Status
  let hasHeader = false;
  
  if (metadata.header) {
    hasHeader = (metadata.header.toLowerCase() === 'true');
  } else {
    // Fallback: Scan Column A to resolve header presence
    const headerRowIndex = fFindHeaderRowIndex(sheet);
    if (headerRowIndex !== -1) {
      hasHeader = true;
      metadata.header = 'True';
    } else {
      hasHeader = false;
      metadata.header = 'False';
    }
  }
  
  // Always ensure tags is set to True in the metadata note
  metadata.tags = 'True';
  
  // Write the compiled and standardized metadata back to the A1 note
  fWriteA1Metadata(sheet, metadata);

  // 3. Freeze Header Row(s) and Column A
  let rowsToFreeze = 1;
  if (hasHeader) {
    const headerRowIndex = fFindHeaderRowIndex(sheet);
    if (headerRowIndex !== -1) {
      rowsToFreeze = headerRowIndex + 1; // Freeze all rows from the header row up
    }
  }
  
  sheet.setFrozenRows(rowsToFreeze);
  sheet.setFrozenColumns(1);

  // 4. Style Column A (Row tags)
  // Text format: bold, italic, left-aligned, Nord charcoal background
  const colARange = sheet.getRange(1, 1, maxRows, 1);
  colARange.setBackground('#2E3440'); // Darker Nord Charcoal
  colARange.setFontColor('#ECEFF4'); // Nord Snow White
  colARange.setFontWeight('bold');
  colARange.setFontStyle('italic');
  colARange.setHorizontalAlignment('left');

  // 5. Style Row 1 (Column tags)
  // Text format: bold, centered, Nord slate background, light gray text
  const row1Range = sheet.getRange(1, 1, 1, maxCols);
  row1Range.setBackground('#3B4252'); // Nord Slate Grey
  row1Range.setFontColor('#ECEFF4'); // Nord Snow White
  row1Range.setFontWeight('bold');
  row1Range.setFontStyle('normal'); // Reset italic for header row
  row1Range.setHorizontalAlignment('center');

  // 6. Add Conditional Formatting to highlight empty cells in Column A and Row 1
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


/* function fGetA1Metadata
   Purpose: Parses Cell A1 note into an object of key-value pairs (lowercased keys).
   Assumptions: Note format is "Key: Value" on separate lines.
   @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to read from.
   @returns {object} The parsed metadata object.
*/
function fGetA1Metadata(sheet) {
  const note = sheet.getRange(1, 1).getNote() || '';
  const metadata = {};
  
  note.split('\n').forEach(line => {
    const parts = line.split(':');
    if (parts.length >= 2) {
      const key = parts[0].trim().toLowerCase();
      const val = parts.slice(1).join(':').trim();
      metadata[key] = val;
    }
  });
  return metadata;
} // End function fGetA1Metadata


/* function fWriteA1Metadata
   Purpose: Standardizes and writes the metadata object back to the Cell A1 note.
   Assumptions: None.
   @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to write to.
   @param {object} metadata - The metadata object.
   @returns {void}
*/
function fWriteA1Metadata(sheet, metadata) {
  const lines = [];
  
  // Write keys back in a standardized order
  if (metadata.tags) lines.push(`Tags: ${metadata.tags}`);
  if (metadata.header) lines.push(`Header: ${metadata.header}`);
  if (metadata.hide) lines.push(`Hide: ${metadata.hide}`);
  
  // Also preserve any custom unrecognized metadata lines
  Object.keys(metadata).forEach(key => {
    if (key !== 'tags' && key !== 'header' && key !== 'hide') {
      lines.push(`${key}: ${metadata[key]}`);
    }
  });
  
  sheet.getRange(1, 1).setNote(lines.join('\n'));
} // End function fWriteA1Metadata


/* function fFindHeaderRowIndex
   Purpose: Scans Column A for the "header" tag.
   Assumptions: fNormalizeTags is globally available.
   @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to scan.
   @returns {number} The 0-based row index of the "Header" tag, or -1 if not found.
*/
function fFindHeaderRowIndex(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return -1;
  
  const colAValues = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let r = 0; r < colAValues.length; r++) {
    const val = colAValues[r][0];
    if (val && typeof val === 'string') {
      const normalizedTags = fNormalizeTags(val);
      if (normalizedTags.includes('header')) {
        return r;
      }
    }
  }
  return -1;
} // End function fFindHeaderRowIndex
