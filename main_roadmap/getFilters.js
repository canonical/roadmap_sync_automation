const SHEETS_TO_SCAN = ["cloud", "charming", "saas", "devices", "is", "product", "security", "excellence", "ubuntu", "web"];
const OUTPUT_FOLDER_ID = "1DJDyuclwfrK_mdr2lIpssIHaxSknC04L";
const OUTPUT_FILENAME = "roadmap_filters.json";


function generateFilterDictionary() {
  Logger.log("Starting filter dictionary generation...");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const finalDictionary = {};

  // Define which rows contain the team names and the filter strings
  const TEAM_NAME_ROW = 1;
  const FILTER_STRING_ROW = 2;

  for (const sheetName of SHEETS_TO_SCAN) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Sheet "${sheetName}" not found. Skipping.`);
      continue;
    }

    Logger.log(`Processing sheet: "${sheetName}"...`);
    finalDictionary[sheetName] = {};

    const lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
      continue; // Skip empty sheets
    }

    // Read the top two rows in a single batch call for efficiency
    const headerData = sheet.getRange(TEAM_NAME_ROW, 1, 2, lastColumn).getValues();
    const teamNames = headerData[0];     // This is the first row of data
    const filterStrings = headerData[1]; // This is the second row of data

    for (let i = 0; i < lastColumn; i++) {
      const teamName = teamNames[i];
      const filterString = filterStrings[i];

      // Only add an entry if both a team name and a filter string exist
      if (teamName && filterString) {
        finalDictionary[sheetName][teamName.toString()] = filterString.toString();
      }
    }
  }

  // Convert the JavaScript object into a nicely formatted JSON string
  const jsonOutput = JSON.stringify(finalDictionary, null, 2);

  // Save the JSON string to a file in Google Drive
  try {
    const folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);

    // To ensure we overwrite, we find and delete any existing file with the same name
    const existingFiles = folder.getFilesByName(OUTPUT_FILENAME);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    const newFile = folder.createFile(OUTPUT_FILENAME, jsonOutput, MimeType.PLAIN_TEXT);
    Logger.log(`Success! File created: ${newFile.getName()}`);
    Logger.log(`File URL: ${newFile.getUrl()}`);

  } catch (e) {
    Logger.log(`Error saving file to Google Drive. Please check that the OUTPUT_FOLDER_ID is correct and that you have permission to write to it.`);
    Logger.log(e);
  }
}
