//sheet names that should be processed
const SHEETS = ["cloud", "charming", "saas", "devices", "is", "product", "security", "excellence", "ubuntu", "web"];
const RELOAD_INDEX_SHEET = true;
const BACKUP_NEEDED = true;
const EXECUTION_FREQUENCY_IN_HOURS = 4
//const SHEETS = ["devices"];
//const RELOAD_INDEX_SHEET = false;
//const BACKUP_NEEDED = false;

const CURRENT_CYCLE = "25.10"; //current cycle
const FUTURE_CYCLE = "26.04"; //can be left empty if not applicable.

const JIRA_DATA_SPREADSHEET_ID = "1E_Qa5zCtI4JeiXKq0yW2KNzU9F1Q_Bt39FxVCxVMjZ4"; //Spreadsheet ID with Jira data for roadmap
const CYCLE_REGEX_PATTERN = /^\d{2}\.\d{2}$/; //regex pattern for cycles
const PROJECT_REGEX_PATTERN = /^(.*?)\((.*?)\)\[(.*?)\]$|^(.*?)\[(.*?)\]\((.*?)\)$|^(.*?)\((.*?)\)$|^(.*?)\[(.*?)\]$|^(.*?)$/; // regex pattern for project filter

const BACKUP_FOLDER_ID = "1B37pAPfBXAsTlSD3azt4FY-mrNhaa3jT"; //from the URL e.g. https://drive.google.com/drive/folders/**FOLDER_ID**
const MAX_BACKUPS_COUNT = 360;

const WHITE_STATUSES = ["Untriaged", "Triaged"] //not started statuses
const GREEN_STATUSES = ["In Progress", "In Review", "To Be Deployed", "BLOCKED"] // statuses that green by default, if roadmap state is empty
const RED_STATUSES = ["Rejected"] // statuses that red by default, if roadmap state is empty
const COMPLETED_STATUSES = ["Done"] // C value in the color state

const STATE_COLORS = {
  "At Risk": "#e59138",        // Custom orange
  "Excluded": "#cc0000",       // Dark Red 1
  "Added": "#3d85c6",          // Dark Blue 1
  "Dropped": "#000000",        // Black
  "green": "#6aa84f",          // Dark Green 1
  "purple": "#741b47",         // Dark Magenta 2
  "red": "#cc0000",            // Dark Red 1
  "white": "#ffffff"           // White
};
const FONT_FAMILY = "Ubuntu";

const PROJECT_ROW_INDEX = 2; //row index with projects keys
const ALL_CYCLES_COLUMN_INDEX = 1; //column index with all cycles

const INDEX_SHEET_NAME = "index";

const CONTACT_LINK = "https://chat.canonical.com/canonical/channels/jira"
const DOC_LINK = "https://docs.google.com/document/d/1_Cjvi3gjZ6xsdkodkan3fMCoaSbDH1GgUdy5jzQGq5U/edit?tab=t.0"


let raw_data_cache = {};

//main function for manual run or for time based trigger
function main() {
  const propService = PropertiesService.getScriptProperties();
  const lastExecutionTime = propService.getProperty('last_execution');
  const currentTime = new Date();

  if (lastExecutionTime) {
    const hoursInMillis = EXECUTION_FREQUENCY_IN_HOURS * 60 * 60 * 1000;
    const lastExecutionDate = new Date(lastExecutionTime);
    const timeSinceLastExecution = currentTime.getTime() - (lastExecutionDate.getTime() - (5 * 60 * 1000));//fix the shift

    if (timeSinceLastExecution < hoursInMillis) {
      Logger.log(`Last execution was less than ${EXECUTION_FREQUENCY_IN_HOURS} hours ago.`);
      Logger.log(`Last run: ${lastExecutionDate.toISOString()}`);
      Logger.log(`Current time: ${currentTime.toISOString()}`);
      Logger.log(`TimeSinceLastExecution in ms: ${timeSinceLastExecution}`)
      return;
    }
  }

  //create backup of the spreadsheet
  if (BACKUP_NEEDED) {
    backupSpreadsheet(BACKUP_FOLDER_ID, MAX_BACKUPS_COUNT)
  }
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  //remove all unfinished temp sheets
  for (const sheetName of SHEETS) {
    let sheet = ss.getSheetByName(sheetName + "_temp");
    if (sheet) {
      ss.deleteSheet(sheet);
    }
  }

  processSheets(ss, CURRENT_CYCLE, true);
  if (FUTURE_CYCLE) {
    processSheets(ss, FUTURE_CYCLE, false);
  }
  if (RELOAD_INDEX_SHEET) {
    generateIndexSheet(ss, currentTime);
  }

  //switch originals to temps
  for (const sheetName of SHEETS) {
    switchSheets(ss, sheetName, sheetName + "_temp")
  }

  switchSheets(ss, INDEX_SHEET_NAME, INDEX_SHEET_NAME + "_temp")

  Logger.log("Running script at: " + currentTime.toISOString());
  propService.setProperty('last_execution', currentTime.toISOString());
}

function processSheets(ss, cycleNumber, withColorUpdate) {
  for (const sheetName of SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Sheet ${sheetName} not found.`);
      continue;
    }

    Logger.log(`Fetching data for sheet: ${sheetName}. Cycle Number: ${cycleNumber}`);
    processSheet(ss, sheet, cycleNumber, withColorUpdate); //process the sheet
  }
}

function processSheet(ss, sheet, cycleNumber, withColorUpdate) {

  sheetName = sheet.getName();
  //create a copy of the sheet and hide it
  let tempSheet = copyAndHideSheet(ss, sheet, sheetName + "_temp");

  let cycles = findCyclesOnTheSheet(tempSheet); // get info about cycles numbers and their positions on the sheet

  if (!cycles.has(cycleNumber)) {
    Logger.log(`Cycle ${cycleNumber} not found on the roadmap.`)
    let newReleaseCell = tempSheet.getRange(tempSheet.getLastRow() + 2, 1)
    newReleaseCell.setNumberFormat('@STRING@');
    newReleaseCell.setValue(cycleNumber);
    tempSheet.getRange(6, 1, 1, tempSheet.getLastColumn()).copyFormatToRange(tempSheet, newReleaseCell.getColumn(), newReleaseCell.getColumn(), newReleaseCell.getRow(), newReleaseCell.getRow());
    cycles = findCyclesOnTheSheet(tempSheet);
  }

  if (cycles.has(cycleNumber)) {

    let cycleRowIndex = cycles.get(cycleNumber);
    let nextprevCycles = getNextPrevKeys(cycles, cycleNumber);

    let lastCycleRow = nextprevCycles.next ? cycles.get(nextprevCycles.next) : tempSheet.getLastRow() + 2;
    let maxRow = 0

    let projectColumnIndex = 4;
    let projectsStringCell = tempSheet.getRange(PROJECT_ROW_INDEX, projectColumnIndex);
    let projectValue = projectsStringCell.getValue();

    let currentRowIndex = cycleRowIndex + 1;
    let currentColumnIndex = 2;

    for (let row = cycleRowIndex; row < lastCycleRow; row++) {
      let depth = tempSheet.getRowGroupDepth(row);
      if (depth < 1) continue;
      tempSheet.getRowGroup(row, depth).remove();
    }

    while (projectValue) {
      Logger.log(`Processing the ${projectValue}.`);
      //one time call. Needed for creation of the additional column
      //tempSheet.insertColumnAfter(currentColumnIndex+3)

      let projectsRange = tempSheet.getRange(currentRowIndex, currentColumnIndex, lastCycleRow - currentRowIndex, 5);
      projectsRange.clear();
      projectsRange.clearFormat();
      projectsRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

      let row_index = currentRowIndex;

      let projects = projectsStringCell.getValue().toString().split(";");
      let projectFilters = []
      projects.forEach(projectKey => {
        if (isNullOrWhitespace(projectKey)) {
          return
        }
        let projectFilter = parseProjectFilterValue(projectKey);
        if (projectFilter) {
          projectFilters.push(projectFilter);
        }
        else
          Logger.log(`Project filter ${projectKey} is invalid. Please check the format. It can be <project key>(<labels>)[<components>], <project key>[<components>](<labels>), <project key>(<labels>), <project key>[<components>] or just <project key>`)
      })

      row_index = processProject(cycleNumber, projectFilters, tempSheet, row_index, currentColumnIndex, lastCycleRow, withColorUpdate)
      if (row_index > maxRow) {
        maxRow = row_index;
      }
      if (row_index > lastCycleRow)
        lastCycleRow = row_index;

      //tempSheet.setColumnWidth(currentColumnIndex, 25)
      //tempSheet.setColumnWidth(currentColumnIndex + 1, 25)
      //tempSheet.autoResizeColumn(currentColumnIndex + 3)
      //const currentWidth = tempSheet.getColumnWidth(currentColumnIndex + 3);
      //tempSheet.setColumnWidth(currentColumnIndex + 3, currentWidth + 10);
      //tempSheet.setColumnWidth(currentColumnIndex + 4, 25)

      currentColumnIndex += 5;
      projectColumnIndex += 5;
      projectsStringCell = tempSheet.getRange(PROJECT_ROW_INDEX, projectColumnIndex);
      projectValue = projectsStringCell.getValue();
    }

    //delete empty rows
    if (lastCycleRow - 1 > maxRow)
      tempSheet.deleteRows(maxRow, (lastCycleRow - 1) - maxRow)

    //grouping
    try {
      const group = tempSheet.getRowGroup(cycleRowIndex + 1, 1)
      if (group) {
        group.remove();
      }
    }
    catch {
    }

    if (maxRow - (cycleRowIndex + 1) > 0) {
      let cycleRange = tempSheet.getRange(cycleRowIndex + 1, 1, maxRow - (cycleRowIndex + 1), 3)
      cycleRange.shiftRowGroupDepth(1);
    }
  }
}

function processProject(cycleNumber, projectFilters, sheet, row_index, projectColumnIndex, lastCycleRow, withColorUpdate) {
  const start = new Date();

  let projectData = getProjectIssuesInHierarchy(projectFilters, cycleNumber);

  if (!projectData.hierarchy || projectData.hierarchy.length == 0) {
    Logger.log("Jira Data is empty.")
  }

  let itemsRowsCount = projectData.count
  let currentRowsCount = lastCycleRow - row_index;
  let howMany = itemsRowsCount - currentRowsCount

  //Logger.log("Row index: " + row_index + " Last Cycle Row: " + lastCycleRow + " Total Rows: " + itemsRowsCount + " currentRowsCount " + currentRowsCount + " howMany " + howMany)
  if (itemsRowsCount > currentRowsCount) {
    sheet.insertRowsBefore(lastCycleRow, howMany);
  }

  for (let parent in projectData.hierarchy) {
    let issue = projectData.hierarchy[parent]

    if (parent != 'None') {

      //Summary
      let parentLinkURL = issue.parentLink;

      sheet.getRange(row_index, projectColumnIndex, 1, 4).setValues([["", "", issue.summary, issue.key]]).setFontWeight("bold").setFontFamily(FONT_FAMILY);

      const richText = SpreadsheetApp.newRichTextValue()
        .setText(issue.key)
        .setLinkUrl(parentLinkURL)
        .build();
      sheet.getRange(row_index, projectColumnIndex + 3).setRichTextValue(richText).setFontFamily(FONT_FAMILY);
      sheet.getRange(row_index, projectColumnIndex + 4).setValue("").setFontFamily(FONT_FAMILY);
      row_index++;
    }

    let childrens = issue.children

    for (let childIndex in childrens) {

      let child = childrens[childIndex]

      //carry over
      if (withColorUpdate) {
        let carryOverCell = sheet.getRange(row_index, projectColumnIndex);
        let labelsarr = findCyclesInTheLabel(child.labels, cycleNumber);
        let labels_count_value = labelsarr.length > 1 ? labelsarr.length - 1 : ""
        if (labels_count_value > 1)
          carryOverCell.setValue(labels_count_value);
        if (labelsarr.length > 1) {
          carryOverCell.setFontColor(STATE_COLORS['white']).setFontWeight('bold').setHorizontalAlignment("center").setFontFamily(FONT_FAMILY);
          carryOverCell.setBackground(STATE_COLORS["purple"]);
        }

        //State
        let stateCell = sheet.getRange(row_index, projectColumnIndex + 1);

        stateCell.setFontColor(STATE_COLORS['white']).setFontWeight('bold').setHorizontalAlignment("center").setFontFamily(FONT_FAMILY);
        let backgroundColor = STATE_COLORS["white"]; // Default color

        if (COMPLETED_STATUSES.includes(child.status)) {
          stateCell.setValue("C");
          backgroundColor = (child.state === "Added")
            ? STATE_COLORS["Added"]
            : STATE_COLORS["green"];
        } else if (child.state && STATE_COLORS[child.state]) {
          backgroundColor = STATE_COLORS[child.state];
        } else if (GREEN_STATUSES.includes(child.status)) {
          backgroundColor = STATE_COLORS["green"];
        } else if (RED_STATUSES.includes(child.status)) {
          backgroundColor = STATE_COLORS["red"];
        }
        stateCell.setBackground(backgroundColor);
      }

      //Summary
      let epicLinkURL = child.epicLink;

      sheet.getRange(row_index, projectColumnIndex + 2).setValue(child.summary).setFontFamily(FONT_FAMILY)
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(child.key)
        .setLinkUrl(epicLinkURL)
        .build();

      sheet.getRange(row_index, projectColumnIndex + 3).setRichTextValue(richText).setFontFamily(FONT_FAMILY);

      sheet.getRange(row_index, projectColumnIndex + 4).setValue("").setFontFamily(FONT_FAMILY);
      row_index++;
    }

    sheet.getRange(row_index, projectColumnIndex, 1, 5).setValues([["", "", "", "", ""]]).setFontFamily(FONT_FAMILY)
    row_index++;
  }
  const end = new Date();
  const executionTime = end - start
  Logger.log(`processProject execution time: ${executionTime} ms`);
  return row_index;
}

function getProjectIssuesInHierarchy(projectFilters, cycleNumber) {
  //const start = new Date();
  let data;
  let itemsCount = 0;

  if (cycleNumber in raw_data_cache && raw_data_cache[cycleNumber].length > 0) {
    data = raw_data_cache[cycleNumber]
  }
  else {
    let sheetName = cycleNumber + "_data";
    let externalSpreadsheet = SpreadsheetApp.openById(JIRA_DATA_SPREADSHEET_ID);
    let sheet = externalSpreadsheet.getSheetByName(sheetName);
    data = sheet.getDataRange().getValues()
    raw_data_cache[cycleNumber] = data
  }

  let hierarchy = {};

  projectFilters.forEach(projectFilter => {
    Logger.log(`Fetching data for project: ${projectFilter.projectKey}. Components: ${projectFilter.components} Labels: ${projectFilter.labels}`);
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
          itemsCount += 2 //parent plus whitespace
        }

        hierarchy[parentKey].children.push({
          key: key,
          summary: summary,
          epicLink: epicLink,
          status: epicStatus,
          state: epicState,
          labels: labels
        });
        itemsCount += 1;
      }
    }
  })

  //const end = new Date();
  //const executionTime = end - start
  //Logger.log(`getProjectIssuesInHierarchy() execution time: ${executionTime} ms`);

  return {
    hierarchy: hierarchy,
    count: itemsCount
  };
}

function findCyclesOnTheSheet(sheet) {
  const range = sheet.getRange(1, ALL_CYCLES_COLUMN_INDEX, sheet.getLastRow());
  const values = range.getValues();
  const valueToCellMap = new Map();

  for (let i = 0; i < values.length; i++) {
    const cellValue = values[i][0];
    if (CYCLE_REGEX_PATTERN.test(cellValue)) {
      valueToCellMap.set(cellValue.toString(), i + 1);
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
  const start = new Date();
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

  const end = new Date();
  const executionTime = (end - start);
  Logger.log(`backupSpreadsheet execution time: ${executionTime} ms`);
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
  Logger.log(`Temp sheet has been created: ${target_sheetName}`);
  return tempSheet;
}

function generateIndexSheet(ss, lastExecutionTime) {

  Logger.log("Index sheet generation started.");
  // Get or create the Index sheet
  let indexSheet = ss.getSheetByName(INDEX_SHEET_NAME);

  let tempIndexSheet = copyAndHideSheet(ss, indexSheet, indexSheet.getName() + "_temp", true)

  for (let i = 1; i <= tempIndexSheet.getLastColumn(); i++) {
    sheetName = tempIndexSheet.getRange(1, i).getValue().toLowerCase()
    //Logger.log("Sheet Name: " + sheetName)
    tempIndexSheet.getRange(3, i, tempIndexSheet.getLastRow(), 1).clear({
      contentsOnly: true,
      formatOnly: false
    });

    targetSheet = ss.getSheetByName(sheetName + "_temp");
    if (targetSheet) {
      const targetSheetId = targetSheet.getSheetId();

      valuesRange = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn());

      let rowIndex = 3

      for (let j = 1; j <= valuesRange.getNumColumns(); j++) {
        let teamName = targetSheet.getRange(1, j).getValue()
        let jiraProjectKey = targetSheet.getRange(2, j).getValue()
        if (teamName && jiraProjectKey) {
          //Logger.log("Team Name:" + teamName);
          let rangeid = targetSheet.getRange(1, j - 2, 1, 3).getA1Notation()
          const url = `${ss.getUrl()}#gid=${targetSheetId}&range=${rangeid}`;
          const linkText = `${teamName}`;

          const richText = SpreadsheetApp.newRichTextValue()
            .setText(linkText)
            .setLinkUrl(url)
            .build();

          tempIndexSheet.getRange(rowIndex, i).setRichTextValue(richText);
          rowIndex += 1;
        }
      }
    }
  }

  if (lastExecutionTime) {
    const lastRow = 16;
    const lastExecutionDate = new Date(lastExecutionTime);

    const infoText = `Updated every: ${EXECUTION_FREQUENCY_IN_HOURS} hours | Last updated on: ${lastExecutionDate.toISOString()} (UTC) | Link to documentation: PR030 | Contact: JIRA`;

    const richTextBuilder = SpreadsheetApp.newRichTextValue()
      .setText(infoText);

    const docLinkText = "PR030";
    const docStartIndex = infoText.indexOf(docLinkText);
    const docEndIndex = docStartIndex + docLinkText.length;
    richTextBuilder.setLinkUrl(docStartIndex, docEndIndex, DOC_LINK);

    const contactLinkText = "JIRA";
    const contactStartIndex = infoText.indexOf(contactLinkText);
    const contactEndIndex = contactStartIndex + contactLinkText.length;
    richTextBuilder.setLinkUrl(contactStartIndex, contactEndIndex, CONTACT_LINK);

    // Set the RichTextValue to the cell
    const range = tempIndexSheet.getRange(lastRow, 1);
    range.clearFormat();
    range.setWrap(false);
    range.setRichTextValue(richTextBuilder.build());
  }

  Logger.log("Index sheet updated.");
}

function switchSheets(ss, sheetName, tempSheetName) {
  Logger.log(`Sheet ${sheetName} will be switched to ${tempSheetName}.`);
  const sheet = ss.getSheetByName(sheetName);
  let tempSheet = ss.getSheetByName(tempSheetName)
  let position = getSheetPosition(ss, sheetName);
  ss.deleteSheet(sheet);
  tempSheet.setName(sheetName);
  tempSheet.showSheet();
  moveSheetToPosition(ss, sheetName, position)
}


function isNullOrWhitespace(str) {
  return (str === null || str === undefined || str.trim() === '');
}