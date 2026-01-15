function compareJsonFiles() {
  const FILE_NAME_1 = "jira_data_26.04__2025-12-04T00-52-08-225Z.json";
  //const FILE_NAME_2 = "jira_data_26.04__2025-11-19T11-51-35-318Z.json";
  const SHEET_NAME = "Delta";
  const JSON_FOLDER_ID = "1DJDyuclwfrK_mdr2lIpssIHaxSknC04L"
  const CYCLE = '26.04'

  const folder = DriveApp.getFolderById(JSON_FOLDER_ID)
  const cycle_folder = folder.getFoldersByName(CYCLE).next()

  const result = getLatestAndFirstFileInFolder(cycle_folder);
  //const file1 = result.firstCreatedFile;
  const file2 = result.latestFile;

  // Get the files from Google Drive
  const file1 = cycle_folder.getFilesByName(FILE_NAME_1).hasNext() ? cycle_folder.getFilesByName(FILE_NAME_1).next() : null;
  //const file2 = cycle_folder.getFilesByName(FILE_NAME_2).hasNext() ? cycle_folder.getFilesByName(FILE_NAME_2).next() : null;

  if (!file1 || !file2) {
    Logger.log("One or both files not found. Please check the file names.");
    return;
  }

  const json1 = JSON.parse(file1.getBlob().getDataAsString());
  const json2 = JSON.parse(file2.getBlob().getDataAsString());

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
  sheet.clear(); // Clear previous data

  const headers = ["Change Type", "Key", "Field", "Old Value", "New Value"];
  sheet.appendRow(headers);

  const diffs = findDiffs(json1, json2);


  if (diffs.length > 0) {
    sheet.getRange(2, 1, diffs.length, headers.length).setValues(diffs);
  } else {
    sheet.appendRow(["No differences found"]);
  }

  Logger.log("Comparison completed. Check the sheet for results.");
}

// Function to find differences between two JSON arrays
function findDiffs(json1, json2) {
  const diffs = [];

  // Create maps for quick lookup by key
  const map1 = createMapFromJson(json1);
  const map2 = createMapFromJson(json2);

  // Find removed items (present in json1 but not in json2)
  for (const key in map1) {
    if (!map2[key]) {
      diffs.push(["Removed", key, "-", JSON.stringify(map1[key]), "-"]);
    }
  }

  // Find added items (present in json2 but not in json1)
  for (const key in map2) {
    if (!map1[key]) {
      diffs.push(["Added", key, "-", "-", JSON.stringify(map2[key])]);
    }
  }

  // Find updated items (present in both but with differences)
  for (const key in map1) {
    if (map2[key]) {
      const updates = findFieldDiffs(map1[key], map2[key]);
      updates.forEach(update => {
        diffs.push(["Updated", key, update.field, update.oldValue, update.newValue]);
      });
    }
  }

  return diffs;
}

function createMapFromJson(jsonArray) {
  const map = {};
  jsonArray.forEach(item => {
    map[item.key] = item;
  });
  return map;
}

function findFieldDiffs(obj1, obj2) {
  const fieldDiffs = [];
  for (const field in obj1) {
    if (obj1[field] !== obj2[field]) {
      fieldDiffs.push({
        field: field,
        oldValue: obj1[field] !== undefined ? obj1[field] : "-",
        newValue: obj2[field] !== undefined ? obj2[field] : "-"
      });
    }
  }
  return fieldDiffs;
}

function getLatestAndFirstFileInFolder(folder) {

  if (!folder) {
    Logger.log(`Folder "${folderName}" not found.`);
    return;
  }

  // Get all files in the folder
  const files = folder.getFiles();
  const fileList = [];

  // Collect file metadata (creation date and file object)
  while (files.hasNext()) {
    const file = files.next();
    fileList.push({
      name: file.getName(),
      id: file.getId(),
      createdDate: file.getDateCreated(),
      fileObject: file
    });
  }

  if (fileList.length === 0) {
    Logger.log(`No files found in the folder "${folderName}".`);
    return;
  }

  // Sort the files by creation date
  fileList.sort((a, b) => a.createdDate - b.createdDate);

  // Get the first created file (oldest) and the latest file (newest)
  const firstCreatedFile = fileList[0];
  const latestFile = fileList[fileList.length - 1];

  // Log the results
  Logger.log(`First Created File: ${firstCreatedFile.name} (Created: ${firstCreatedFile.createdDate})`);
  Logger.log(`Latest File: ${latestFile.name} (Created: ${latestFile.createdDate})`);

  // Return the results
  return {
    firstCreatedFile: firstCreatedFile.fileObject,
    latestFile: latestFile.fileObject
  };
}