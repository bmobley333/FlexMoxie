/* global SpreadsheetApp, fNormalizeTags, fShowMessage */
/* exported fVerifyActiveSheetTags */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Sheet Verification Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function toA1_
   Purpose: Pure JavaScript helper to compute A1 notation from 1-based row and column indices.
   Assumptions: Row and col are 1-based positive integers.
   Notes: Eliminates the need for sheet.getRange().getA1Notation() RPC calls inside loops.
   @param {number} row - The 1-based row number.
   @param {number} col - The 1-based column number.
   @returns {string} The A1 notation string (e.g., "A1", "Z5", "AA10").
*/
function toA1_(row, col) {
  let letter = '';
  let c = col;
  while (c > 0) {
    c--;
    letter = String.fromCharCode(65 + (c % 26)) + letter;
    c = Math.floor(c / 26);
  }
  return letter + row;
}

/* function fVerifyActiveSheetTags
   Purpose: Verifies unique column and row tags on the currently active sheet by leveraging fGetSheetData.
   Assumptions: The function is triggered by a user on an active sheet.
   Notes: This version is faster and tests the same cached tag maps the rest of the system uses.
   @returns {void}
*/
function fVerifyActiveSheetTags() {
  fShowToast('⏳ Verifying all tags...', 'Tag Verification');
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();

  try {
    // Use fGetSheetData to get the canonical tag maps for this sheet, forcing a refresh
    const { arr } = fGetSheetData('Temp', sheetName, SpreadsheetApp.getActiveSpreadsheet(), true);

    // 1. Verify Column Tags
    const seenColTags = {};
    const colTagRow = arr[0] || [];
    for (let c = 0; c < colTagRow.length; c++) {
      const normalizedTags = fNormalizeTags(colTagRow[c]);
      for (const tag of normalizedTags) {
        if (seenColTags[tag]) {
          const message = `Duplicate column tag found: "${tag}"\n\nOriginal in cell: ${seenColTags[tag]}\nDuplicate in cell: ${toA1_(1, c + 1)}`;
          fEndToast();
          fShowMessage('⚠️ Tag Verification Failed', message);
          return;
        }
        seenColTags[tag] = toA1_(1, c + 1);
      }
    }

    // 2. Verify Row Tags
    const seenRowTags = {};
    for (let r = 0; r < arr.length; r++) {
      // Ensure the row and cell exist before trying to read from it
      if (arr[r] && arr[r][0]) {
        const normalizedTags = fNormalizeTags(arr[r][0]);
        for (const tag of normalizedTags) {
          if (seenRowTags[tag]) {
            const message = `Duplicate row tag found: "${tag}"\n\nOriginal in cell: ${seenRowTags[tag]}\nDuplicate in cell: ${toA1_(r + 1, 1)}`;
            fEndToast();
            fShowMessage('⚠️ Tag Verification Failed', message);
            return;
          }
          seenRowTags[tag] = toA1_(r + 1, 1);
        }
      }
    }

    fEndToast();
    fShowMessage('✅ Tag Verification', '✅ Success! All column and row tags are unique.');

  } catch (e) {
    fEndToast();
    fShowMessage('❌ Error', `An error occurred during verification: ${e.message}`);
  }
} // End function fVerifyActiveSheetTags