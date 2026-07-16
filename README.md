# Flex Moxie Development Engine

This repository houses the developer-focused Google Apps Script codebase, custom JSON tables, and database schemas for the **Flex Moxie** game setting.

---

## 🌐 Website & Rules Boundary Decision

To maintain clean repository separation and a unified player experience:
*   **Rules & Presentation Site:** The public rules, player guides, and landing portal are centralized in the [MetaScape VitePress Hub](https://github.com/bmobley333/MetaScape-VitePress-GitHub-Pages) repository. Do **not** commit markdown documentation or player-facing manual files to this repository.
*   **Engine & Scripts Code:** This repository contains strictly spreadsheet-processing backend code, validation modules, and custom Sheets formulas.
*   **Master Data Backend:** All calculation logic is bound to the [Flex Moxie Core Spreadsheet](https://docs.google.com/spreadsheets/d/1S4FqrXWDsNNIvWg4B9kum1EG_xtfs6DNFeKsdYFTLqs/edit#gid=1346804318).

---

## 🛠️ Development & Deployment Workflow

This project is structured as a collection of subprojects (modules) designed to sync to Google Apps Script projects using `@google/clasp`.

### Subproject Layout
*   `CS/` - Character Sheets forms and sheets processing logic.
*   `Codex/` - Reference catalog database lookup engine.
*   `Cust/` - Custom custom sheet formulas and modifiers.
*   `DB/` - Database CRUD endpoints bound to Sheets.
*   `FlexLib/` - Shared standard helper functions.
*   `Tables/` - Spreadsheet row-mapping and schema builders.
*   `Ver/` - Version tracking and schema upgrading scripts.

### Syncing Code (Google Apps Script Sync)
1. **Login to clasp:**
   Authenticate clasp with your Google account:
   ```bash
   npx @google/clasp login
   ```
   *Note: Select the `metascapegame@gmail.com` profile in the browser authentication screen.*

2. **Configure Script IDs:**
   Create a `.clasp.json` configuration file inside each subproject directory pointing to the corresponding Apps Script project ID:
   ```json
   {
     "scriptId": "YOUR_APPS_SCRIPT_PROJECT_ID_HERE",
     "rootDir": "."
   }
   ```

3. **Deploy Changes:**
   To push all local changes to the Google Apps Script cloud environment, run the PowerShell script helper:
   ```powershell
   powershell -ExecutionPolicy Bypass -File push.ps1
   ```

---

## ⚖️ License

**Copyright © 2026 High Tower Forge**

This work is licensed under the **Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License**. To view a copy of this license, visit: [http://creativecommons.org/licenses/by-nc-sa/4.0/](http://creativecommons.org/licenses/by-nc-sa/4.0/).
