//Sheet names
const CONFIG_SHEET_NAME = "Config"
const PROJECTS_SHEET_NAME = "Projects"
const DATA_SHEET_NAME = "Jira Data";

//Config cells
const JIRA_BASE_URL_CELL = "B5" // Jira's base url e.g. https://warthogs-sandbox.atlassian.net
const JIRA_EMAIL_CELL = "B6" //Jira's system username or email
const CYCLES_CELL = "B3" //name of the cycles that need to be  fetched. Will be used in the jql search e.g. labels = 24.04;24.10;25.04;25.10;
const LAST_SYNC_DATE_CELL = "B1" //cell on the config sheet to store the lst update datetime
const JIRA_API_TOKEN_PROPERTY_NAME = "JIRA_API_TOKEN" //Api token saved in the script properties

//!!! check data processing section in case these ids need to be changed
const ROADMAP_STATE_FIELD_ID = "customfield_10968" //id of the Roadmap State customfield in Jira. 
const RANK_FIELD_ID = "customfield_10019"
const TEAM_FIELD_ID = "customfield_10001"

const JIRA_API_BATCH_SIZE = 200; // For Jira API limits
const JIRA_API_SLEEP = 100; // For Jira API limits

const JSON_FOLDER_ID = "1DJDyuclwfrK_mdr2lIpssIHaxSknC04L"

parentRanks = new Map()

function main() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = sheet.getSheetByName(CONFIG_SHEET_NAME);
  const projectsSheet = sheet.getSheetByName(PROJECTS_SHEET_NAME);

  // Read Jira Base URL and Email from Configuration sheet
  if (!configSheet || !projectsSheet) {
    Logger.log("Configuration or Projects sheet not found.");
    return;
  }

  const jiraBaseUrl = configSheet.getRange(JIRA_BASE_URL_CELL).getValue();
  const jiraEmail = configSheet.getRange(JIRA_EMAIL_CELL).getValue();
  const cyclesToFetch = configSheet.getRange(CYCLES_CELL).getValue().trim();


  // Get Jira API token from Script Properties
  const jiraToken = PropertiesService.getScriptProperties().getProperty(JIRA_API_TOKEN_PROPERTY_NAME);
  if (!jiraBaseUrl || !jiraEmail || !jiraToken) {
    Logger.log(`Missing Jira credentials in Configuration sheet ("${JIRA_BASE_URL_CELL}" and "${JIRA_EMAIL_CELL}" cells) or Script Properties ("${JIRA_API_TOKEN_PROPERTY_NAME}").`);
    return;
  }

  // Read project keys from "Projects" sheet
  const projects = projectsSheet.getRange("A:A").getValues().flat().filter(p => p); // Remove empty rows
  if (projects.length === 0) {
    Logger.log("No projects found in Projects sheet.");
    return;
  }

  var cycles = cyclesToFetch.split(";");

  cycles = cycles.filter(function (part) {
    return part !== "";
  });

  for (let i = 0; i < cycles.length; i++) {

    cycle = cycles[i]
    Logger.log(`Fetching data for cycle: ${cycle}`);

    const dataSheetName = cycle + "_data"
    let dataSheet = sheet.getSheetByName(dataSheetName);
    if (!dataSheet) {
      dataSheet = sheet.insertSheet();
      dataSheet.setName(dataSheetName);
    }

    const projectsJql = projects.map(p => `"${p}"`).join(",");

    const jql =
      `"Properties[Checkboxes]" = "Roadmap Item"` +
      ` AND project IN (${projectsJql})` +
      ` AND issuetype = Epic` +
      ` AND labels = "${cycle}"` +
      ` ORDER BY Parent ASC, Rank`;

    Logger.log(`Fetching data for all projects in one call for cycle: ${cycle}`);

    const totalIssues = fetchAllIssuesByJql_(jiraBaseUrl, jiraEmail, jiraToken, jql);

    const parentKeys = [...new Set(
      totalIssues.map(i => i.fields.parent?.key).filter(Boolean)
    )];

    Logger.log(`Getting info about parent's rank in a bulk`);
    fetchParentRanksBulk_(jiraBaseUrl, jiraEmail, jiraToken, parentKeys, parentRanks);


    let allData = totalIssues.map(issue => {
      const projectKey = issue.fields.project?.key || "";

      return [
        projectKey,
        issue.key,
        issue.fields.summary,
        issue.fields.status.name,
        issue.fields.customfield_10968
          ? issue.fields.customfield_10968.value.replace(/[^\p{L}\p{N}\p{P}\p{Z}^$\n]/gu, '').trim()
          : "",
        `'` + (issue.fields.labels || []).join(", "),
        getComponentsString(issue.fields.components),
        issue.fields.parent ? issue.fields.parent.key : "",
        issue.fields.parent ? issue.fields.parent.fields.summary : "",
        jiraBaseUrl + '/browse/' + issue.key,
        issue.fields.parent ? jiraBaseUrl + '/browse/' + issue.fields.parent.key : "",
        issue.fields.customfield_10019 ? issue.fields.customfield_10019 : "",
        issue.fields.parent ? parentRanks.get(issue.fields.parent.key) : "",
        issue.fields.customfield_10001 ? issue.fields.customfield_10001.name : "",
        (issue.fields.updated ? new Date(issue.fields.updated) : "")
      ];
    });

    allData.sort((a, b) => {
      const projectA = a[0] || "";
      const projectB = b[0] || "";

      // 1) Group by project first
      if (projectA < projectB) return -1;
      if (projectA > projectB) return 1;

      // 2) Then by parent rank
      const parentKeyA = a[7];
      const parentKeyB = b[7];
      const parentRankA = parentRanks.get(parentKeyA) || "";
      const parentRankB = parentRanks.get(parentKeyB) || "";

      if (parentRankA < parentRankB) return -1;
      if (parentRankA > parentRankB) return 1;

      // 3) Then by issue rank
      const issueRankA = a[11] || "";
      const issueRankB = b[11] || "";

      if (issueRankA < issueRankB) return -1;
      if (issueRankA > issueRankB) return 1;

      return 0;
    });



    dataSheet.clear(); // Clear previous data

    // Define headers and write them to the sheet
    const headers = ["Project Key", "Epic Key", "Summary", "Current Status", "Roadmap State", "Labels", "Components", "Parent Key", "Parent Summary", "Epic Link", "Parent Link", "Issue Rank", "Parent Rank", "Team", "Updated"];
    dataSheet.appendRow(headers);

    // Write all data at once for better performance
    if (allData.length > 0) {
      dataSheet.getRange(2, 1, allData.length, headers.length).setValues(allData);

      // Prepare JSON data from allData
      const jsonData = prepareJsonData(allData);
      //Save data as a json in folder
      saveDataAsJson(cycle, jsonData, JSON_FOLDER_ID);
    }

  }

  // Save last sync date in Configuration sheet
  const lastSyncDate = new Date();
  configSheet.getRange(LAST_SYNC_DATE_CELL).setValue(lastSyncDate);
  Logger.log(`Data updated successfully. Last Sync: ${lastSyncDate}`);
}

function fetchParentRank(parentKey, jiraBaseUrl, jiraEmail, jiraToken) {
  const url = `${jiraBaseUrl}/rest/api/3/issue/${parentKey}?fields=customfield_10019`;
  const options = {
    method: "GET",
    headers: {
      Authorization: `Basic ${Utilities.base64Encode(jiraEmail + ":" + jiraToken)}`,
      Accept: "application/json"
    },
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200 || !data.fields || !data.fields.customfield_10019) {
      Logger.log(`Parent rank not found for ${parentKey}`);
      return "";
    }

    Logger.log(`Parent call to check the rank ${parentKey} Rank ${data.fields.customfield_10019}`);
    return data.fields.customfield_10019;

  } catch (e) {
    Logger.log(`Error fetching rank for ${parentKey}: ${e.message}`);
    return "";
  }
}

function getComponentsString(components) {
  var result = []
  if (components) {
    for (let i = 0; i < components.length; i++) {
      result.push(components[i].name);
    }
  }

  return result.join(", ");
}

function prepareJsonData(allData) {
  return allData.map(row => ({
    project: row[0],
    key: row[1],
    summary: row[2],
    status: row[3],
    roadmapState: row[4],
    labels: row[5],
    components: row[6],
    parentKey: row[7],
    parentSummary: row[8],
    epicLink: row[9],
    parentLink: row[10],
    issueRank: row[11],
    parentRank: row[12],
    team: row[13],
    updated: row[14]
  }));
}

// Save data as JSON file in Google Drive
function saveDataAsJson(cycle, data, folder_id) {
  const folder = DriveApp.getFolderById(folder_id);

  const cycle_folder = folder.getFoldersByName(cycle).hasNext()
    ? folder.getFoldersByName(cycle).next()
    : folder.createFolder(cycle);

  const jsonData = JSON.stringify(data);
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  const fileName = `jira_data_${cycle}__${timestamp}.json`;

  cycle_folder.createFile(fileName, jsonData);
  Logger.log(`JSON file saved: ${fileName}`);
}

function fetchAllIssuesByJql_(jiraBaseUrl, jiraEmail, jiraToken, jql) {
  Logger.log(`JQL: ${jql}`);
  let issues = [];
  let nextPageToken = null;

  const url = `${jiraBaseUrl}/rest/api/3/search/jql`;

  const headers = {
    Authorization: `Basic ${Utilities.base64Encode(jiraEmail + ":" + jiraToken)}`,
    Accept: "application/json",
    "Content-Type": "application/json"
  };

  while (true) {
    const payload = {
      jql: jql,
      maxResults: JIRA_API_BATCH_SIZE,
      fields: [
        "key",
        "summary",
        "status",
        "project",
        ROADMAP_STATE_FIELD_ID,
        "labels",
        "parent",
        "components",
        RANK_FIELD_ID,
        TEAM_FIELD_ID,
        "updated"
      ]
    };

    if (nextPageToken) payload.nextPageToken = nextPageToken;

    const options = {
      method: "POST",
      headers: headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const text = response.getContentText();
    const data = JSON.parse(text);

    if (response.getResponseCode() !== 200) {
      Logger.log(`Error fetching issues. Code: ${response.getResponseCode()}`);
      Logger.log(`Response: ${text}`);
      throw new Error(`Jira search/jql failed: ${response.getResponseCode()}`);
    }

    const pageIssues = data.issues || [];
    if (!pageIssues.length) break;

    issues = issues.concat(pageIssues);
    nextPageToken = data.nextPageToken || null;

    if (!nextPageToken) break;

    if (JIRA_API_SLEEP > 0) Utilities.sleep(JIRA_API_SLEEP);
  }

  return issues;
}

function fetchParentRanksBulk_(jiraBaseUrl, jiraEmail, jiraToken, parentKeys, parentRanksMap) {
  // parentKeys: array of strings
  if (!parentKeys || parentKeys.length === 0) return;

  // Only fetch missing keys
  const missing = parentKeys.filter(k => k && !parentRanksMap.has(k));
  if (missing.length === 0) return;

  const url = `${jiraBaseUrl}/rest/api/3/search/jql`;

  const headers = {
    Authorization: `Basic ${Utilities.base64Encode(jiraEmail + ":" + jiraToken)}`,
    Accept: "application/json",
    "Content-Type": "application/json"
  };

  // Chunk size: keep safe for JQL size + maxResults.
  const CHUNK_SIZE = 100;

  for (let i = 0; i < missing.length; i += CHUNK_SIZE) {
    const chunk = missing.slice(i, i + CHUNK_SIZE);

    const keysJql = chunk.map(k => `"${k}"`).join(",");
    const jql = `key in (${keysJql})`;

    let nextPageToken = null;
    while (true) {
      const payload = {
        jql: jql,
        maxResults: Math.min(JIRA_API_BATCH_SIZE, chunk.length),
        fields: ["key", RANK_FIELD_ID]
      };
      if (nextPageToken) payload.nextPageToken = nextPageToken;

      const options = {
        method: "POST",
        headers: headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const text = response.getContentText();
      const data = JSON.parse(text);

      if (response.getResponseCode() !== 200) {
        Logger.log(`Bulk parent ranks failed. Code: ${response.getResponseCode()}`);
        Logger.log(`Response: ${text}`);
        break;
      }

      (data.issues || []).forEach(issue => {
        const key = issue.key;
        const rank = issue.fields ? (issue.fields[RANK_FIELD_ID] || "") : "";
        parentRanksMap.set(key, rank);
      });

      nextPageToken = data.nextPageToken || null;
      if (!nextPageToken) break;

      if (JIRA_API_SLEEP > 0) Utilities.sleep(JIRA_API_SLEEP);
    }

    // Cache missing keys as empty to avoid re-fetching them again later
    chunk.forEach(k => {
      if (!parentRanksMap.has(k)) parentRanksMap.set(k, "");
    });

    if (JIRA_API_SLEEP > 0) Utilities.sleep(JIRA_API_SLEEP);
  }
}

