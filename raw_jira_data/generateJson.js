// Google Drive Folder ID for JSON output
const JSON_FOLDER = "1DJDyuclwfrK_mdr2lIpssIHaxSknC04L"

function generateJsonFromSheets() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  const cyclesToFetch = '25.10'
  let cycles = cyclesToFetch.split(";").filter(c => c.trim() !== "");

  if (cycles.length === 0) {
    Logger.log("No cycles defined in Config sheet cell " + CYCLES_CELL);
    SpreadsheetApp.getUi().alert("Error: No cycles defined in 'Config' sheet cell " + CYCLES_CELL);
    return;
  }

  Logger.log("Starting JSON generation from existing sheets...");
  let filesGenerated = 0;

  for (const cycle of cycles) {
    const dataSheetName = cycle + "_data";
    const dataSheet = sheet.getSheetByName(dataSheetName);

    if (!dataSheet) {
      Logger.log(`Sheet '${dataSheetName}' not found. Skipping cycle '${cycle}'.`);
      continue; // Skip to the next cycle
    }

    const lastRow = dataSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log(`Sheet '${dataSheetName}' is empty (no data found). Skipping cycle '${cycle}'.`);
      continue; // Skip to the next cycle
    }

    try {
      // Read all data from row 2 to the end
      // Get the full data range (row 2, col 1, to last row, last col)
      const numCols = dataSheet.getLastColumn();
      if (numCols < 14) {
        Logger.log(`Sheet '${dataSheetName}' has fewer than 14 columns. Skipping to avoid errors.`);
        continue;
      }

      const allData = dataSheet.getRange(2, 1, lastRow - 1, numCols).getValues();

      if (allData.length > 0) {
        // Prepare JSON data from the sheet data
        const jsonData = prepareJsonData_sheet(allData);
        // Save data as a JSON file in the specified folder
        saveDataAsJson_sheet(cycle, jsonData, JSON_FOLDER);
        Logger.log(`Successfully generated JSON for cycle: ${cycle}`);
        filesGenerated++;
      } else {
        Logger.log(`No data rows found in sheet '${dataSheetName}'.`);
      }
    } catch (e) {
      Logger.log(`Error processing sheet '${dataSheetName}': ${e.message}`);
      SpreadsheetApp.getUi().alert(`Error processing sheet '${dataSheetName}'. Check logs.`);
    }
  }

  Logger.log("JSON generation from existing sheets complete.");
}


/**
 * Converts the 2D array from the sheet into an array of objects for JSON.
 * This assumes the sheet has at least 14 columns matching the original script's output.
 */
function prepareJsonData_sheet(allData) {
  return allData.map(row => ({
    project: row[0] || "",          // Project Key
    key: row[1] || "",              // Epic Key
    summary: row[2] || "",          // Summary
    status: row[3] || "",           // Current Status
    roadmapState: row[4] || "",     // Roadmap State
    labels: row[5] || "",           // Labels
    components: row[6] || "",       // Components
    parentKey: row[7] || "",        // Parent Key
    parentSummary: row[8] || "",    // Parent Summary
    epicLink: row[9] || "",         // Epic Link
    parentLink: row[10] || "",      // Parent Link
    issueRank: row[11] || "",       // Issue Rank
    parentRank: row[12] || "",      // Parent Rank
    team: row[13] || ""             // Team
  }));
}

/**
 * Saves the provided data as a JSON file in a specific Google Drive folder.
 */
function saveDataAsJson_sheet(cycle, data, folder_id) {
  try {
    const folder = DriveApp.getFolderById(folder_id);

    const cycle_folder = folder.getFoldersByName(cycle).hasNext()
      ? folder.getFoldersByName(cycle).next()
      : folder.createFolder(cycle);

    const jsonData = JSON.stringify(data, null, 2); // Add indentation for readability
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    const fileName = `jira_data_${cycle}__${timestamp}.json`;

    cycle_folder.createFile(fileName, jsonData, MimeType.PLAIN_TEXT); // Specify MIME type
    Logger.log(`JSON file saved: ${fileName} in folder ${cycle}`);
  } catch (e) {
    Logger.log(`Error saving JSON file for cycle ${cycle}: ${e.message}`);
    SpreadsheetApp.getUi().alert(`Error saving JSON file for cycle ${cycle}. Check logs and Drive folder permissions.`);
  }
}

