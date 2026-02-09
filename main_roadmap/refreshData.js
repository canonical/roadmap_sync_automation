//sheet names that should be processed
const SHEETS = ["cloud", "charming", "saas", "devices", "cs", "sre", "product", "security", "excellence", "ubuntu", "web"];
const RELOAD_INDEX_SHEET = true;
const BACKUP_NEEDED = true;

const frequency = PropertiesService.getScriptProperties().getProperty("EXECUTION_FREQUENCY_IN_HOURS");
const EXECUTION_FREQUENCY_IN_HOURS = frequency ? parseInt(frequency, 10) : 12; // default to 12 if missing

const CURRENT_CYCLE = PropertiesService.getScriptProperties().getProperty("CURRENT_CYCLE"); //current cycle
const FUTURE_CYCLE = PropertiesService.getScriptProperties().getProperty("FUTURE_CYCLE"); //can be left empty if not applicable.
//const FUTURE_CYCLE = false; //can be left empty if not applicable.

const JIRA_DATA_SPREADSHEET_ID = "1E_Qa5zCtI4JeiXKq0yW2KNzU9F1Q_Bt39FxVCxVMjZ4"; //Spreadsheet ID with Jira data for roadmap
const CYCLE_REGEX_PATTERN = /^\d{2}\.\d{2}$/; //regex pattern for cycles
const PROJECT_REGEX_PATTERN = /^(.*?)\((.*?)\)\[(.*?)\]$|^(.*?)\[(.*?)\]\((.*?)\)$|^(.*?)\((.*?)\)$|^(.*?)\[(.*?)\]$|^(.*?)$/; // regex pattern for project filter

const BACKUP_FOLDER_ID = "1B37pAPfBXAsTlSD3azt4FY-mrNhaa3jT"; //from the URL e.g. https://drive.google.com/drive/folders/**FOLDER_ID**

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
let project_map_cache = {};

//main function for manual run or for time based trigger
function main() {
  const propService = PropertiesService.getScriptProperties();
  const lastExecutionTime = propService.getProperty('last_execution');
  var currentTime = new Date();

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

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  //create backup of the spreadsheet
  if (BACKUP_NEEDED) {
    backupBeforeRun(ss, BACKUP_FOLDER_ID)
  }

  //remove all unfinished temp sheets
  for (const sheetName of SHEETS) {
    let sheet = ss.getSheetByName(sheetName + "_temp");
    if (sheet) {
      ss.deleteSheet(sheet);
    }
  }

  processSheets(ss, CURRENT_CYCLE, true, lastExecutionTime);

  if (FUTURE_CYCLE) {
    processSheets(ss, FUTURE_CYCLE, false, lastExecutionTime);
  }
  if (RELOAD_INDEX_SHEET) {
    generateIndexSheet(ss);
  }

  //switch originals to temps
  for (const sheetName of SHEETS) {
    switchSheets(ss, sheetName, sheetName + "_temp")
  }

  if (RELOAD_INDEX_SHEET) {
    switchSheets(ss, INDEX_SHEET_NAME, INDEX_SHEET_NAME + "_temp")
  };

  Logger.log("Running script at: " + currentTime.toISOString());
  propService.setProperty('last_execution', currentTime.toISOString());
}

function processSheets(ss, cycleNumber, withColorUpdate, lastExecutionTime) {
  for (const sheetName of SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Sheet ${sheetName} not found.`);
      continue;
    }

    Logger.log(`Processing of the sheet: ${sheetName}. Cycle Number: ${cycleNumber}`);
    processSheet(ss, sheet, cycleNumber, withColorUpdate, lastExecutionTime); //process the sheet
  }
}

function processSheet(ss, sheet, cycleNumber, withColorUpdate, lastExecutionTime) {
  const start = new Date();
  Logger.log("processSheet start")
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

  // shift last_execution backwards to avoid skipping updates made during the previous run
  const lastExecutionDate = lastExecutionTime
    ? new Date(new Date(lastExecutionTime).getTime() - 60 * 60 * 1000)
    : null;

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

    Logger.log("Groups remove process start")
    for (let row = cycleRowIndex; row < lastCycleRow; row++) {
      let depth = tempSheet.getRowGroupDepth(row);
      if (depth < 1)
        continue;
      else
        tempSheet.getRowGroup(row, depth).remove();
      break;
    }
    Logger.log("Groups remove process end")

    while (projectValue) {
      Logger.log(`Processing the ${projectValue}.`);
      //one time call. Needed for creation of the additional column
      //tempSheet.insertColumnAfter(currentColumnIndex+3)

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

      let projectData = getProjectIssuesInHierarchy(projectFilters, cycleNumber);

      const expectedEndRow = currentRowIndex + projectData.count;

      if (projectData && projectData.latestUpdated && lastExecutionDate) {
        if (lastExecutionDate >= projectData.latestUpdated) {
          Logger.log(
            `Skip refresh for ${projectValue}: last_execution (${lastExecutionDate.toISOString()}) ` +
            `>= latest issue updated (${projectData.latestUpdated.toISOString()})`
          );
          if (expectedEndRow > maxRow) {
            maxRow = expectedEndRow;
          }

          // move to next project column
          currentColumnIndex += 5;
          projectColumnIndex += 5;
          projectsStringCell = tempSheet.getRange(PROJECT_ROW_INDEX, projectColumnIndex);
          projectValue = projectsStringCell.getValue();
          continue;
        }
      }

      if (!projectData || !projectData.hierarchy || projectData.hierarchy.length == 0) {
        Logger.log("Jira Data is empty. Project will not be synced.")
        return
      }

      let projectsRange = tempSheet.getRange(currentRowIndex, currentColumnIndex, lastCycleRow - currentRowIndex, 5);
      projectsRange.clear();
      projectsRange.clearFormat();
      normalize_range_format(projectsRange);

      let row_index = currentRowIndex;

      row_index = processProject(cycleNumber, projectData, tempSheet, row_index, currentColumnIndex, lastCycleRow, withColorUpdate)
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
  Logger.log(`processSheet execution time: ${Date.now() - start} ms`);
}

function normalize_range_format(range) {
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setHorizontalAlignment("left")
    .setFontWeight("normal")
    .setFontFamily(FONT_FAMILY)
    .setBackground("white")
    .setFontColor("black");
}

function processProject(cycleNumber, projectData, sheet, row_index, projectColumnIndex, lastCycleRow, withColorUpdate) {
  const start = new Date();
  Logger.log("processProject start")
  const itemsRowsCount = projectData.count;
  const currentRowsCount = lastCycleRow - row_index;
  const howMany = itemsRowsCount - currentRowsCount;

  if (howMany > 0) {
    sheet.insertRowsBefore(lastCycleRow > 1 ? lastCycleRow - 1 : 1, howMany);
  }

  if (itemsRowsCount === 0) {
    Logger.log(`processProject execution time: ${Date.now() - start} ms (No items to process)`);
    return row_index;
  }

  const totalRowsToProcess = itemsRowsCount;
  const targetRange = sheet.getRange(row_index, projectColumnIndex, totalRowsToProcess, 5);

  // BATCH ARRAYS
  const values = Array(totalRowsToProcess).fill(null).map(() => Array(5).fill(""));
  const backgrounds = Array(totalRowsToProcess).fill(null).map(() => Array(5).fill(STATE_COLORS["white"]));
  const fontColors = Array(totalRowsToProcess).fill(null).map(() => Array(5).fill("black"));
  const fontWeights = Array(totalRowsToProcess).fill(null).map(() => Array(5).fill("normal"));
  const richTextValues = Array(totalRowsToProcess).fill(null).map(() => Array(5).fill(null));
  const alignments = Array(totalRowsToProcess).fill(null).map(() => Array(5).fill("left"));

  let currentRow = 0;

  for (const parent in projectData.hierarchy) {
    const issue = projectData.hierarchy[parent];
    if (parent !== 'None') {
      values[currentRow] = ["", "", issue.summary, issue.key, ""];
      fontWeights[currentRow] = ["bold", "bold", "bold", "bold", "bold"];

      if (issue.key && issue.parentLink) {
        const parentRichText = SpreadsheetApp.newRichTextValue()
          .setText(issue.key)
          .setLinkUrl(issue.parentLink)
          .build();
        richTextValues[currentRow][3] = parentRichText;
      }
      currentRow++;
    }

    const childrens = issue.children;
    for (const child of childrens) {
      const labelsarr = findCyclesInTheLabel(child.labels, cycleNumber);
      if (labelsarr.length > 1) {
        if (labelsarr.length - 1 != 1) {
          values[currentRow][0] = labelsarr.length - 1;
        }
        else
          values[currentRow][0] = "";
        fontColors[currentRow][0] = STATE_COLORS['white'];
        fontWeights[currentRow][0] = 'bold';
        alignments[currentRow][0] = "center";
        backgrounds[currentRow][0] = STATE_COLORS["purple"];
      }
      if (withColorUpdate) {
        fontColors[currentRow][1] = STATE_COLORS['white'];
        fontWeights[currentRow][1] = 'bold';
        alignments[currentRow][1] = "center";
        let backgroundColor = STATE_COLORS["white"];
        if (COMPLETED_STATUSES.includes(child.status)) {
          values[currentRow][1] = "C";
          backgroundColor = (child.state === "Added") ? STATE_COLORS["Added"] : STATE_COLORS["green"];
        } else if (child.state && STATE_COLORS[child.state]) {
          backgroundColor = STATE_COLORS[child.state];
        } else if (GREEN_STATUSES.includes(child.status)) {
          backgroundColor = STATE_COLORS["green"];
        } else if (RED_STATUSES.includes(child.status)) {
          backgroundColor = STATE_COLORS["red"];
        }
        backgrounds[currentRow][1] = backgroundColor;
      }

      values[currentRow][2] = child.summary;
      values[currentRow][3] = child.key;

      // Only create linked rich text for specific cells
      if (child.key && child.epicLink) {
        const childRichText = SpreadsheetApp.newRichTextValue()
          .setText(child.key)
          .setLinkUrl(child.epicLink)
          .build();
        richTextValues[currentRow][3] = childRichText;
      }

      values[currentRow][4] = "";
      currentRow++;
    }

    values[currentRow] = ["", "", "", "", ""];
    currentRow++;
  }

  // --- BATCH WRITE 
  for (let r = 0; r < totalRowsToProcess; r++) {
    for (let c = 0; c < 5; c++) {
      if (richTextValues[r][c] === null) {
        const text = values[r][c] ? values[r][c].toString() : "";
        richTextValues[r][c] = SpreadsheetApp.newRichTextValue().setText(text).build();
      }
    }
  }

  // Apply all formatting at once
  targetRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setBackgrounds(backgrounds)
    .setFontColors(fontColors)
    .setFontWeights(fontWeights)
    .setHorizontalAlignments(alignments)
    .setRichTextValues(richTextValues);

  Logger.log(`processProject execution time: ${Date.now() - start} ms`);

  return row_index + currentRow;
}

function getProjectIssuesInHierarchy(projectFilters, cycleNumber) {
  let data;

  if (cycleNumber in raw_data_cache && raw_data_cache[cycleNumber].length > 0) {
    data = raw_data_cache[cycleNumber];
  } else {
    let sheetName = cycleNumber + "_data";
    let externalSpreadsheet = SpreadsheetApp.openById(JIRA_DATA_SPREADSHEET_ID);
    let sheet = externalSpreadsheet.getSheetByName(sheetName);
    data = sheet.getDataRange().getValues();
    raw_data_cache[cycleNumber] = data;
  }

  let projectMap = project_map_cache[cycleNumber];
  if (!projectMap) {
    projectMap = new Map();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const projectKey = row[0];
      if (!projectMap.has(projectKey)) projectMap.set(projectKey, []);
      projectMap.get(projectKey).push(row);
    }
    project_map_cache[cycleNumber] = projectMap;
  }

  let hierarchy = {};
  let itemsCount = 0;
  let latestUpdated = null;

  projectFilters.forEach(projectFilter => {
    Logger.log(`Processing the data for project: ${projectFilter.projectKey}. Components: ${projectFilter.components} Excluded Components: ${projectFilter.excludedComponents} Labels: ${projectFilter.labels} Teams: ${projectFilter.teams} Excluded Teams: ${projectFilter.excludedTeams}`);

    const projectRows = projectMap.get(projectFilter.projectKey);

    if (!projectRows || projectRows.length === 0) {
      return;
    }

    for (const row of projectRows) {
      let include = true;
      let componentsCell = (row[6] || "").toString();

      if (projectFilter.components && projectFilter.components.length > 0) {
        include = false;
        componentsCell.split(",").forEach(component => {
          if (projectFilter.components.includes(component.trim())) include = true;
        });
      }

      if (projectFilter.excludedComponents && projectFilter.excludedComponents.length > 0) {
        componentsCell.split(",").forEach(component => {
          if (projectFilter.excludedComponents.includes(component.trim())) include = false;
        });
      }
      if (!include) continue;

      let labelsCell = (row[5] || "").toString();
      if (projectFilter.labels && projectFilter.labels.length > 0) {
        let labelInclude = false;
        labelsCell.split(",").forEach(label => {
          if (projectFilter.labels.includes(label.trim())) labelInclude = true;
        });
        if (!labelInclude) continue;
      }

      let teamsCell = (row[13] || "").toString();
      if (projectFilter.teams && projectFilter.teams.length > 0) {
        let teamInclude = false;
        teamsCell.split(",").forEach(team => {
          if (projectFilter.teams.includes(team.trim())) teamInclude = true;
        });
        if (!teamInclude) continue;
      }

      if (projectFilter.excludedTeams && projectFilter.excludedTeams.length > 0) {
        let teamExcluded = false;
        teamsCell.split(",").forEach(team => {
          if (projectFilter.excludedTeams.includes(team.trim())) teamExcluded = true;
        });
        if (teamExcluded) continue;
      }

      const updatedCell = row[14];
      let updatedDate = null;
      if (updatedCell instanceof Date) {
        updatedDate = updatedCell;
      } else if (updatedCell) {
        const d = new Date(updatedCell);
        if (!isNaN(d.getTime())) updatedDate = d;
      }
      if (updatedDate && (!latestUpdated || updatedDate > latestUpdated)) {
        latestUpdated = updatedDate;
      }

      let parentSummary = row[8];
      let parentKey = row[7];
      let parentLink = row[10];
      let summary = row[2];
      let key = row[1];
      let epicLink = row[9];
      let epicState = row[4];
      let epicStatus = row[3];

      if (!parentKey) {
        parentKey = "None";
      }
      if (!hierarchy[parentKey]) {
        hierarchy[parentKey] = {
          key: parentKey,
          summary: parentSummary,
          parentLink: parentLink,
          children: []
        };
        itemsCount += 2;
      }

      hierarchy[parentKey].children.push({
        key: key,
        summary: summary,
        epicLink: epicLink,
        status: epicStatus,
        state: epicState,
        labels: labelsCell
      });
      itemsCount += 1;
    }
  });

  return {
    hierarchy: hierarchy,
    count: itemsCount,
    latestUpdated: latestUpdated
  };
}

function findCyclesOnTheSheet(sheet) {
  const range = sheet.getRange(1, ALL_CYCLES_COLUMN_INDEX, sheet.getLastRow());
  const values = range.getValues();
  const valueToCellMap = new Map();

  for (let i = 0; i < values.length; i++) {
    const cellValue = values[i][0];
    if (cellValue) {
      if (CYCLE_REGEX_PATTERN.test(cellValue)) {
        valueToCellMap.set(cellValue.toString(), i + 1);
      }
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

function backupBeforeRun(ss, backupFolderID) {
  const start = Date.now();
  const folder = DriveApp.getFolderById(backupFolderID);

  const now = new Date();
  const ts = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH-mm-ss");
  const name = `${ss.getName()} - Backup ${ts}`;

  DriveApp.getFileById(ss.getId()).makeCopy(name, folder);

  Logger.log(`Pre-run backup created: ${name}`);
  Logger.log(`backupBeforeRun execution time: ${Date.now() - start} ms`);
}

function parseProjectFilterValue(value) {
  if (!value || !value.trim()) return null;

  function parseValues(valueString) {
    if (!valueString) return [];
    return valueString
      .match(/"([^"]+)"|[^,]+/g)
      ?.map(v => v.replace(/^"|"$/g, '').trim())
      .filter(v => v.length > 0) || [];
  }

  let remainingValue = value;
  let labelsStr = "";
  let componentsStr = "";
  let teamsStr = "";


  const componentMatch = remainingValue.match(/\[([^\]]*)\]/);
  if (componentMatch) {
    componentsStr = componentMatch[1] || "";
    remainingValue = remainingValue.replace(componentMatch[0], "");
  }

  const teamMatch = remainingValue.match(/\{([^}]*)\}/);
  if (teamMatch) {
    teamsStr = teamMatch[1] || "";
    remainingValue = remainingValue.replace(teamMatch[0], "");
  }
  const labelMatch = remainingValue.match(/\(([^)]*)\)/);
  if (labelMatch) {
    labelsStr = labelMatch[1] || "";
    remainingValue = remainingValue.replace(labelMatch[0], "");
  }

  const projectKey = remainingValue.trim();

  const labels = parseValues(labelsStr);

  const allComponents = parseValues(componentsStr);
  const components = allComponents.filter(c => !c.startsWith("!"));
  const excludedComponents = allComponents
    .filter(c => c.startsWith("!"))
    .map(c => c.substring(1));

  const allTeams = parseValues(teamsStr);
  const teams = allTeams.filter(t => !t.startsWith("!"));
  const excludedTeams = allTeams
    .filter(t => t.startsWith("!"))
    .map(t => t.substring(1));

  return {
    projectKey,
    components,
    excludedComponents,
    labels,
    teams,
    excludedTeams
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

function generateIndexSheet(ss) {
  Logger.log("Index sheet generation started.");
  const indexSheet = ss.getSheetByName(INDEX_SHEET_NAME);
  if (!indexSheet) {
    Logger.log(`Sheet ${INDEX_SHEET_NAME} not found.`);
    return;
  }
  const tempIndexSheet = copyAndHideSheet(ss, indexSheet, indexSheet.getName() + "_temp", true);


  // 1. Read all necessary data into memory first
  const indexHeaderValues = tempIndexSheet.getRange(1, 1, 1, tempIndexSheet.getLastColumn()).getValues()[0];
  const sheetDataCache = {};
  let maxTeams = 0;

  for (const sheetName of indexHeaderValues) {
    if (!sheetName) continue;
    let targetSheet = ss.getSheetByName(sheetName.toString().toLowerCase() + "_temp");
    if (!targetSheet)
      targetSheet = ss.getSheetByName(sheetName.toString().toLowerCase());
    if (targetSheet) {
      const lastCol = targetSheet.getLastColumn();
      if (lastCol > 0) {
        // Cache the header rows (team name, project key) and sheet ID
        sheetDataCache[sheetName.toString().toLowerCase()] = {
          id: targetSheet.getSheetId(),
          headers: targetSheet.getRange(1, 1, 2, lastCol).getValues()
        };
      }
    }
  }

  // 2. Prepare the rich text array in memory
  const outputColumns = indexHeaderValues.length;
  // Determine the max number of rows needed for the output range
  Object.values(sheetDataCache).forEach(data => {
    let teamCount = 0;
    for (let i = 0; i < data.headers[0].length; i++) {
      if (data.headers[0][i] && data.headers[1][i]) { // Team name and project key exist
        teamCount++;
      }
    }
    if (teamCount > maxTeams) maxTeams = teamCount;
  });

  const outputRows = maxTeams > 0 ? maxTeams : 1;
  const richTextMatrix = Array(outputRows).fill(null).map(() => Array(outputColumns).fill(null));

  // 3. Populate the rich text array
  for (let c = 0; c < outputColumns; c++) {
    const sheetName = indexHeaderValues[c].toString().toLowerCase();
    const data = sheetDataCache[sheetName];
    if (data) {
      let currentRow = 0;
      const teamNames = data.headers[0];
      const projectKeys = data.headers[1];

      for (let j = 0; j < teamNames.length; j++) {
        if (teamNames[j] && projectKeys[j]) {
          let targetSheet = ss.getSheetByName(sheetName + "_temp")
          if (!targetSheet)
            targetSheet = ss.getSheetByName(sheetName)

          const rangeA1 = targetSheet.getRange(1, Math.max(1, j - 1), 1, 3).getA1Notation();
          const url = `${ss.getUrl()}#gid=${data.id}&range=${rangeA1}`;
          const richText = SpreadsheetApp.newRichTextValue()
            .setText(teamNames[j].toString())
            .setLinkUrl(url)
            .build();
          if (currentRow < outputRows) {
            richTextMatrix[currentRow][c] = richText;
          }
          currentRow++;
        }
      }
    }
  }

  // 4. Fill any remaining nulls with blank rich text objects and clear old content
  const indexContentRange = tempIndexSheet.getRange(3, 1, tempIndexSheet.getLastRow() - 2, outputColumns);
  indexContentRange.clearContent();

  for (let r = 0; r < outputRows; r++) {
    for (let c = 0; c < outputColumns; c++) {
      if (richTextMatrix[r][c] === null) {
        richTextMatrix[r][c] = SpreadsheetApp.newRichTextValue().setText("").build();
      }
    }
  }

  // 5. Write everything back to the sheet in one go
  if (outputRows > 0) {
    tempIndexSheet.getRange(3, 1, outputRows, outputColumns).setRichTextValues(richTextMatrix);
  }

  const lastRow = 17;
  const lastExecutionDate = new Date();
  const infoText = `Updated every: ${EXECUTION_FREQUENCY_IN_HOURS} hours | Last updated on: ${lastExecutionDate.toISOString()} (UTC) | Link to documentation: PR030 | Contact: JIRA`;
  //const infoText = `Updated during breaks and overnight | Last updated on: ${lastExecutionDate.toISOString()} (UTC) | Link to documentation: PR030 | Contact: JIRA`;
  const richTextBuilder = SpreadsheetApp.newRichTextValue().setText(infoText);

  const docLinkText = "PR030";
  const docStartIndex = infoText.indexOf(docLinkText);
  richTextBuilder.setLinkUrl(docStartIndex, docStartIndex + docLinkText.length, DOC_LINK);

  const contactLinkText = "JIRA";
  const contactStartIndex = infoText.indexOf(contactLinkText);
  richTextBuilder.setLinkUrl(contactStartIndex, contactStartIndex + contactLinkText.length, CONTACT_LINK);

  const range = tempIndexSheet.getRange(lastRow, 1);
  range.clearFormat();
  range.setWrap(false);
  range.setRichTextValue(richTextBuilder.build());

  Logger.log("Index sheet updated.");
}

function switchSheets(ss, sheetName, tempSheetName) {
  const sheet = ss.getSheetByName(sheetName);
  let tempSheet = ss.getSheetByName(tempSheetName)
  let position = getSheetPosition(ss, sheetName);
  ss.deleteSheet(sheet);
  tempSheet.setName(sheetName);
  tempSheet.showSheet();
  moveSheetToPosition(ss, sheetName, position)
  Logger.log(`Sheet ${sheetName} have been switched to ${tempSheetName}.`);
}


function isNullOrWhitespace(str) {
  return (str === null || str === undefined || str.trim() === '');
}