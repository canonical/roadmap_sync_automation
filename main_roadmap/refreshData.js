//TODO: transition process between cycles
//TODO: grouping of the cycles

const SHEETS = ["cloud", "charming"]; //sheet names that should be processed
const CURRENT_CYCLE = "25.04"; //current cycle
const JIRA_DATA_SPREADSHEET_ID = "1E_Qa5zCtI4JeiXKq0yW2KNzU9F1Q_Bt39FxVCxVMjZ4"; //Spreadsheet ID with Jira data for roadmap
const CYCLE_REGEX_PATTERN = /^\d{2}\.\d{2}$/; //regex pattern for cycles
const PROJECT_REGEX_PATTERN = /^(.*?)\((.*?)\)\[(.*?)\]$|^(.*?)\[(.*?)\]\((.*?)\)$|^(.*?)\((.*?)\)$|^(.*?)\[(.*?)\]$|^(.*?)$/; // regex pattern for project filter
const BACKUP_FOLDER_ID = "10TXVDrdGcvQjmvjf5u0m8Lzj2vYObvXP"; //from the URL e.g. https://drive.google.com/drive/folders/**FOLDER_ID**
const MAX_BACKUPS_COUNT = 720; // 1 month for run every hour 

const WHITE_STATUSES = ["Untriaged", "Triaged"] //not started statuses
const GREEN_STATUSES = ["In Progress", "In Review", "To Be Deployed", "BLOCKED"] // statuses that green by default, if roadmap state is empty
const RED_STATUSES = ["Rejected"] // statuses that red by default, if roadmap state is empty
const COMPLETED_STATUSES = ["Done"] // C value in the color state

const STATE_COLORS = {
  "At Risk": "orange",
  "Excluded": "red",
  "Added": "blue",
  "Dropped": "black"
};

const PROJECT_ROW_INDEX = 2; //row index with projects keys
const ALL_CYCLES_COLUMN_INDEX = 1; //column index with all cycles


//main fucntion
function main() {
  backupSpreadsheet(BACKUP_FOLDER_ID, MAX_BACKUPS_COUNT) //create backup of the spreadsheet

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const sheetName of SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Sheet ${sheetName} not found.`);
      continue;
    }

    Logger.log(`Fetching data for sheet: ${sheetName}`);
    processSheet(ss, sheet, CURRENT_CYCLE); //process the sheet
  }
}

function processSheet(ss, sheet, cycleNumber) {
  let cycles = findCyclesOnTheSheet(sheet); // get info about cycles numbers and their positions on the sheet
  

  if (cycles.has(cycleNumber)) {

    sheetName = sheet.getName();
    //create a copy of the sheet and hide it
    let tempSheet = copyAndHideSheet(ss, sheet, sheetName + "_temp", true);
    //copy original state of the sheet
    copyAndHideSheet(ss, sheet, sheetName + "_original");

    let cycleRowIndex = cycles.get(cycleNumber);
    let nextprevCycles = getNextPrevKeys(cycles, cycleNumber);

    let lastCycleRow = nextprevCycles.next ? cycles.get(nextprevCycles.next) : tempSheet.getLastRow() + 2;
    let maxRow = 0

    let projectColumnIndex = 4;
    let projectsStringCell = tempSheet.getRange(PROJECT_ROW_INDEX, projectColumnIndex);
    let projectValue = projectsStringCell.getValue();


    let currentRowIndex = cycleRowIndex + 1;
    let currentColumnIndex = 2;
    while (projectValue) {
      let projects = projectsStringCell.getValue().toString().split(";");

      let projectsRange = tempSheet.getRange(currentRowIndex, currentColumnIndex, lastCycleRow - currentRowIndex, 3);
      projectsRange.clear();
      projectsRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

      let row_index = currentRowIndex;

      projects.forEach(projectKey => {
        let projectFilter = parseProjectFilterValue(projectKey);

        if (projectFilter) {
          row_index = processProject(projectFilter, tempSheet, row_index, currentColumnIndex, lastCycleRow)
          if (row_index > maxRow) {
            maxRow = row_index;
          }
          if (row_index > lastCycleRow)
            lastCycleRow = row_index;
        }
        else
          Logger.log(`Project filter ${projectKey} is invalid. Please check the format. It can be <project key>(<labels>)[<components>], <project key>[<components>](<labels>), <project key>(<labels>), <project key>[<components>] or just <project key>`)
      })

      currentColumnIndex += 3;
      projectColumnIndex += 3;
      projectsStringCell = tempSheet.getRange(PROJECT_ROW_INDEX, projectColumnIndex);
      projectValue = projectsStringCell.getValue();
    }

    //delete empty rows

    for (var i = lastCycleRow - 1; i >= maxRow; i--) {
      tempSheet.deleteRow(i);
    }

    //grouping
    const group = tempSheet.getRowGroup(cycleRowIndex + 1, 1)
    group.remove();
    let cycleRange = tempSheet.getRange(cycleRowIndex + 1, 1, maxRow - currentRowIndex, 3)
    cycleRange.shiftRowGroupDepth(1);

    //switch the temp and original sheets
    let position = getSheetPosition(ss, sheetName);
    ss.deleteSheet(sheet);
    tempSheet.setName(sheetName);
    tempSheet.showSheet();
    moveSheetToPosition(ss, sheetName, position)
  }
  else {
    Logger.log(`Cycle ${cycleNumber} not found on the roadmap.`)
  }
}

function processProject(cycleNumber, projectFilter, sheet, row_index, projectColumnIndex, lastCycleRow) {

  Logger.log(`Fetching data for project: ${projectFilter.projectKey}. Components: ${projectFilter.components} Labels: ${projectFilter.labels}`);

  let projectData = getProjectIssuesInHierarchy(projectFilter, cycleNumber);

  for (let parent in projectData) {
    let issue = projectData[parent]

    if (parent != 'None') {
      //Carry over
      sheet.getRange(row_index, projectColumnIndex).setValue("")

      //State
      sheet.getRange(row_index, projectColumnIndex + 1).setValue("")

      //Summary
      let parentLinkURL = issue.parentLink;
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(issue.summary + " (" + issue.key + ")")
        .setLinkUrl(issue.summary.length + 2, issue.summary.length + 2 + issue.key.length, parentLinkURL)
        .build();

      sheet.getRange(row_index, projectColumnIndex + 2).setRichTextValue(richText);
      sheet.getRange(row_index, projectColumnIndex + 2).setFontWeight("bold");
      row_index++;
      if (row_index >= lastCycleRow - 1) {
        sheet.insertRowAfter(row_index)
      }
    }

    let childrens = issue.children

    for (let childIndex in childrens) {

      let child = childrens[childIndex]
      //Logger.log(`Row Index: ${row_index} Epic: ${child.key} Status: "${child.status}" State: "${child.state}" Labels: "${child.labels}"`)

      //carry over
      let carryOverCell = sheet.getRange(row_index, projectColumnIndex);
      let labelsarr = findCyclesInTheLabel(child.labels, cycleNumber);
      let labels_count_value = labelsarr.length > 1 ? labelsarr.length : ""
      carryOverCell.setValue(labels_count_value);
      if (labelsarr.length > 1) {
        carryOverCell.setFontColor('white').setFontWeight('bold').setHorizontalAlignment("center");
        carryOverCell.setBackground("purple");
      }

      //State
      let stateCell = sheet.getRange(row_index, projectColumnIndex + 1);

      stateCell.setFontColor('white').setFontWeight('bold').setHorizontalAlignment("center");
      let backgroundColor = "white"; // Default color

      if (COMPLETED_STATUSES.includes(child.status)) {
        stateCell.setValue("C");
        backgroundColor = "green"
      } else {
        if (child.state && STATE_COLORS[child.state]) {
          backgroundColor = STATE_COLORS[child.state];
        }
        else {
          if (GREEN_STATUSES.includes(child.status))
            backgroundColor = "green";
          else if (RED_STATUSES.includes(child.status))
            backgroundColor = "red";
        }
      }
      stateCell.setBackground(backgroundColor);

      //Summary
      let epicLinkURL = child.epicLink;
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(child.summary + " (" + child.key + ")")
        .setLinkUrl(child.summary.length + 2, child.summary.length + 2 + child.key.length, epicLinkURL)
        .build();

      sheet.getRange(row_index, projectColumnIndex + 2).setRichTextValue(richText);

      row_index++;
      if (row_index >= lastCycleRow - 1) {
        sheet.insertRowAfter(row_index);
      }
    }

    sheet.getRange(row_index, projectColumnIndex, 1, 3).setValues([["", "", ""]])
    row_index++;
    if (row_index >= lastCycleRow - 1) {
      sheet.insertRowAfter(row_index);
    }
  }
  return row_index;
}

function getProjectIssuesInHierarchy(projectFilter, cycleNumber) {
  let sheetName = cycleNumber + "_data";
  let externalSpreadsheet = SpreadsheetApp.openById(JIRA_DATA_SPREADSHEET_ID);
  let sheet = externalSpreadsheet.getSheetByName(sheetName);
  let data = sheet.getDataRange().getValues()
  let hierarchy = {};

  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let project = row[0];

    if (project === projectFilter.projectKey) {

      let components = row[6];
      if (projectFilter.components && projectFilter.components.length > 0) {
        let include = false;
        components.split(",").forEach(component => {
          if (projectFilter.components.includes(component.trim())) {
            include = true;
          }
        })

        if (!include)
          continue;
      }

      let labels = row[5];
      if (projectFilter.labels && projectFilter.labels.length > 0) {
        let include = false;
        labels.split(",").forEach(label => {
          if (projectFilter.labels.includes(label.trim())) {
            include = true;
          }
        })

        if (!include)
          continue;
      }

      let parentSummary = row[8];
      let parentKey = row[7];
      let parentLink = row[10]
      let summary = row[2];
      let key = row[1];
      let epicLink = row[9]
      let epicState = row[4]
      let epicStatus = row[3]

      if (!parentKey) {
        parentKey = "None"
      }
      if (!hierarchy[parentKey]) {
        hierarchy[parentKey] = {
          key: parentKey,
          summary: parentSummary,
          parentLink: parentLink,
          children: []
        };
      }

      hierarchy[parentKey].children.push({
        key: key,
        summary: summary,
        epicLink: epicLink,
        status: epicStatus,
        state: epicState,
        labels: labels
      });
    }
  }

  return hierarchy
}

function findCyclesOnTheSheet(sheet) {
  const range = sheet.getRange(1, ALL_CYCLES_COLUMN_INDEX, sheet.getLastRow());
  const values = range.getValues();
  const valueToCellMap = new Map();

  for (let i = 0; i < values.length; i++) {
    const cellValue = values[i][0];
    if (CYCLE_REGEX_PATTERN.test(cellValue)) {
      valueToCellMap.set(cellValue, i + 1);
    }
  }

  return valueToCellMap;
}

function findCyclesInTheLabel(labels, cycleNumber) {
  let labelsArray = []
  if (labels) {
    const allLabels = labels.split(",");

    allLabels.forEach(label => {
      label = label.toString().trim()
      if (CYCLE_REGEX_PATTERN.test(label)) {
        if (parseFloat(label) <= parseFloat(cycleNumber)) {
          labelsArray.push(label)
        }
      }
    })
  }

  return labelsArray;
}

function getNextPrevKeys(map, key) {
  let keys = Array.from(map.keys());
  let index = keys.indexOf(key);

  if (index === -1) {
    return { next: null, prev: null }; // Key not found
  }

  return {
    next: index < keys.length - 1 ? keys[index + 1] : null,
    prev: index > 0 ? keys[index - 1] : null
  };
}

function backupSpreadsheet(backupFolderID, maxBackups) {

  Logger.log("Backup creation. Backup folder: " + backupFolderID)

  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceId = sourceSpreadsheet.getId();
  const folder = DriveApp.getFolderById(backupFolderID);

  // Format timestamp: YYYY-MM-DD HH-MM-SS
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH-mm-ss");
  const backupName = `${sourceSpreadsheet.getName()} - Backup ${timestamp}`;

  // Make a backup copy and move it to the folder
  const backupFile = DriveApp.getFileById(sourceId).makeCopy(backupName, folder);
  Logger.log("Backup created: " + backupFile.getUrl());

  // Clean up old backups
  deleteOldBackups(folder, sourceSpreadsheet.getName(), maxBackups);
}

// Helper function to delete older backups beyond the limit
function deleteOldBackups(folder, baseName, maxBackups) {
  const files = folder.getFiles();
  const backups = [];

  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().startsWith(baseName + " - Backup")) {
      backups.push(file);
    }
  }

  // Sort backups by creation date (newest first)
  backups.sort((a, b) => b.getDateCreated() - a.getDateCreated());

  // Delete oldest ones beyond the allowed max
  for (let i = maxBackups; i < backups.length; i++) {
    backups[i].setTrashed(true);
    Logger.log("Deleted old backup: " + backups[i].getName());
  }
}

function parseProjectFilterValue(value) {
  const match = value.match(PROJECT_REGEX_PATTERN);

  if (!match) {
    Logger.log("Invalid format");
    return null;
  }

  let projectKey = "";
  let labels = [];
  let components = [];

  function parseValues(valueString) {
    if (!valueString) return [];
    return valueString
      .match(/"([^"]+)"|[^,]+/g) // Match quoted values or unquoted ones
      ?.map(v => v.replace(/^"|"$/g, '').trim()) // Remove surrounding quotes and trim
      .filter(v => v.length > 0) || [];
  }

  if (match[1]) { // Format: <project key>(<labels>)[<components>]
    projectKey = match[1].trim();
    labels = parseValues(match[2]);
    components = parseValues(match[3]);
  } else if (match[4]) { // Format: <project key>[<components>](<labels>)
    projectKey = match[4].trim();
    components = parseValues(match[5]);
    labels = parseValues(match[6]);
  } else if (match[7]) { // Format: <project key>(<labels>)
    projectKey = match[7].trim();
    labels = parseValues(match[8]);
  } else if (match[9]) { // Format: <project key>[<components>]
    projectKey = match[9].trim();
    components = parseValues(match[10]);
  } else { // Format: <project key> (Only project key)
    projectKey = match[11].trim();
  }

  return {
    projectKey,
    components,
    labels
  };
}

function getSheetPosition(ss, sheetName) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === sheetName) {
      return i + 1; // Positions are 1-based for human readability
    }
  }
  Logger.log("Sheet not found.");
  return -1;
}

function moveSheetToPosition(ss, sheetName, position) {
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log("Sheet not found.");
    return;
  }

  // Set the sheet as active (required to move it)
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(position);
}

function copyAndHideSheet(ss, source_sheet, target_sheetName, deleteIfExists = false) {
  const existSheet = ss.getSheetByName(target_sheetName);

  if (existSheet) {
    if (deleteIfExists) {
      let targetSheet = ss.getSheetByName(target_sheetName);
      ss.deleteSheet(targetSheet);
    }
    else
      return existSheet;
  }

  let tempSheet = source_sheet.copyTo(ss);
  tempSheet.setName(target_sheetName);
  tempSheet.hideSheet()
  return tempSheet;
}


function transitionToTheFutureCycle() {
  // step 1: update the current cycle, but the state color for each issue should be
  // - green with C if the status is in the COMPLETED_STATUSES
  // - black if Roadmap state = dropped
  // - red in all other statuses
  // step 2: add new cycle to the roadmap, if it doesn't exist
  // step 3: get issues for each project and add them to the new cycle
  // step 5: group new cycle
}