/* global fShowMessage, fVerifyActiveSheetTags, fApplyVisualIsolation, SpreadsheetApp, fActivateSheetByName, fNormalizeTags, fShowToast, fEndToast */
/* exported run */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Command Dispatcher
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function run
   Purpose: Acts as the central dispatcher for all commands initiated from a local sheet script.
   Assumptions: The command string passed matches a key in the commandMap.
   Notes: This provides a single entry point and a master try/catch for robust error handling.
   @param {string} command - The unique identifier for the command to execute.
   @param {string} [sheetToActivate] - The optional name of the sheet to activate before running the command.
   @returns {void}
*/
function run(command, sheetToActivate) {
  try {
    if (sheetToActivate) {
      fActivateSheetByName(sheetToActivate);
    }

    const commandMap = {
      // --- THIS IS THE FIX ---
      CharacterOnboarding: fCharacterOnboarding,
      ClearPowerFilters: () => {
        fClearAllFilterCheckboxes('Filter Powers');
        fFilterPowers();
      },
      ClearMagicItemFilters: () => {
        fClearAllFilterCheckboxes('Filter Magic Items');
        fFilterMagicItems();
      },
      ClearSkillSetFilters: () => {
        fClearAllFilterCheckboxes('Filter Skill Sets');
        fFilterSkillSets();
      },
      InvalidateGameCache: () => fInvalidateSheetCache('CS', 'Game'),
      ShareCustomLists: fShareCustomLists,
      RenameCustomList: fRenameCustomList,
      DeleteCustomList: fDeleteCustomList,
      DeleteSelectedPowers: fDeleteSelectedPowers,
      VerifyAndPublish: fVerifyAndPublish,
      VerifyAndPublishMagicItems: fVerifyAndPublishMagicItems,
      DeleteSelectedMagicItems: fDeleteSelectedMagicItems,
      VerifyAndPublishSkillSets: fVerifyAndPublishSkillSets,
      DeleteSelectedSkillSets: fDeleteSelectedSkillSets,
      CreateCustomList: fCreateNewCustomList,
      SyncPowerChoices: () => fUpdatePowerTablesList(false),
      SyncMagicItemChoices: () => fUpdateMagicItemChoices(false),
      SyncSkillSetChoices: () => fUpdateSkillSetChoices(false),
      AddNewCustomSource: fAddNewCustomSource,
      InitialSetup: fInitialSetup,
      TagVerification: fVerifyActiveSheetTags,
      ApplyVisualIsolation: () => {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheets = ss.getSheets();
        let formattedCount = 0;

        fShowToast('⏳ Scanning and formatting tabs...', 'Visual Isolation');

        sheets.forEach(sheet => {
          if (fIsCoordinateSheet(sheet)) {
            fApplyVisualIsolation(sheet);
            formattedCount++;
          }
        });

        fEndToast();
        fShowMessage('✅ Success', `Successfully formatted ${formattedCount} tagged sheets in "${ss.getName()}". (Skipped ${sheets.length - formattedCount} non-metadata sheets).`);
      },
      ToggleVisibility: fToggleDesignerVisibility,
      TrimSheet: fTrimSheet,
      Test: fTestIdManagement,
      CreateLatestCharacter: fCreateLatestCharacter,
      CreateLegacyCharacter: fCreateLegacyCharacter,
      RenameCharacter: fRenameCharacter,
      DeleteCharacter: fDeleteCharacter,
      ShowPlaceholder: fShowPlaceholderMessage,
      BuildPowers: fBuildPowers,
      BuildMagicItems: fBuildMagicItems,
      FilterPowers: () => fFilterPowers(false),
      FilterMagicItems: () => fFilterMagicItems(false),
      FilterSkillSets: () => fFilterSkillSets(false),
      PrepGameForPaper: fPrepGameForPaper,
      VerifyIndividualSkills: fVerifyIndividualSkills,
      VerifySkillSetLists: fVerifySkillSetLists,
      BuildSkillSets: fBuildSkillSets,
    };

    if (commandMap[command]) {
      commandMap[command]();
    } else {
      throw new Error(`Unknown command received: ${command}`);
    }
  } catch (e) {
    console.error(e);
    fShowMessage('❌ Error', e.message);
  }
} // End function run

/* function fIsCoordinateSheet
   Purpose: Qualifies if a sheet tab uses the coordinate tag system.
   Checks Cell A1 note for "Tags: True" or Column A for the "Header" tag.
   Assumptions: None.
   @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object to check.
   @returns {boolean} True if the sheet qualifies.
*/
function fIsCoordinateSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return false;

  // 1. Check A1 Note for "Tags: True"
  const a1Note = sheet.getRange(1, 1).getNote();
  if (a1Note && /tags\s*:\s*true/i.test(a1Note)) {
    return true;
  }

  // 2. Fallback: Scan Column A for the "Header" tag
  const colAValues = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let r = 0; r < colAValues.length; r++) {
    const cellValue = colAValues[r][0];
    if (cellValue && typeof cellValue === 'string') {
      const normalizedTags = fNormalizeTags(cellValue);
      if (normalizedTags.includes('header')) {
        return true;
      }
    }
  }

  return false;
} // End function fIsCoordinateSheet