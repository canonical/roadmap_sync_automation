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

const ROADMAP_STATE_FIELD_ID = "customfield_10968" //id of the Roadmap State customfield in Jira. !!! check data processing section in case this id need to be changed
const RANK_FIELD_ID = "customfield_10019"
const JIRA_API_BATCH_SIZE = 100; // For Jira API limits
const JIRA_API_SLEEP = 0; // For Jira API limits

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

    let allData = [];

    // Loop through each project
    for (let project of projects) {
      Logger.log(`Fetching data for project: ${project}`);

      let totalIssues = [];
      let response, jsonResponse;
      let nextPageToken = null;

      do {
        const jqlQuery = encodeURIComponent(`"Properties[Checkboxes]" = "Roadmap Item" AND project = "${project}" AND issuetype = Epic AND labels = "${cycle}" ORDER BY Parent ASC, Rank`);
        let jiraUrl = `${jiraBaseUrl}/rest/api/3/search/jql?jql=${jqlQuery}&fields=key,summary,status,${ROADMAP_STATE_FIELD_ID},labels,parent,components,${RANK_FIELD_ID}&maxResults=${JIRA_API_BATCH_SIZE}`;

        if (nextPageToken) {
          jiraUrl += "&nextPageToken=" + nextPageToken;
        }

        // API Request
        const options = {
          method: "GET",
          headers: {
            Authorization: `Basic ${Utilities.base64Encode(jiraEmail + ":" + jiraToken)}`,
            Accept: "application/json"
          }
        };

        try {
          response = UrlFetchApp.fetch(jiraUrl, options);
          jsonResponse = JSON.parse(response.getContentText());

          if (response.getResponseCode() !== 200) {
            // If the response code is not 200 (OK), log the error
            Logger.log(`Error fetching data for project: ${project}. Response code: ${response.getResponseCode()}`);
            Logger.log(`Error message: ${jsonResponse.errorMessages || jsonResponse.errors}`);
            return;
          }

          if (!jsonResponse.issues || jsonResponse.issues.length === 0) {
            Logger.log(`No issues found for project: ${project} with label: ${cycle}`);
            break;
          }

          totalIssues = totalIssues.concat(jsonResponse.issues);
          nextPageToken = jsonResponse.nextPageToken || null;

          // Prevent hitting Jira rate limits
          if (JIRA_API_SLEEP > 0)
            Utilities.sleep(JIRA_API_SLEEP);

        } catch (e) {
          Logger.log(`Error during API request for project: ${project}. Error: ${e.message}`);
          break;
        }

      } while (nextPageToken);

      totalIssues.forEach(issue => {
        const parentKey = issue.fields.parent?.key;
        if (parentKey && !parentRanks.has(parentKey)) {
          const parentRank = fetchParentRank(parentKey, jiraBaseUrl, jiraEmail, jiraToken);
          parentRanks.set(parentKey, parentRank);
        }
      });

      // Process and append data
      const projectData = totalIssues.map(issue => [
        project,
        issue.key,
        issue.fields.summary,
        issue.fields.status.name,
        issue.fields.customfield_10968 ? issue.fields.customfield_10968.value.replace(/[^\p{L}\p{N}\p{P}\p{Z}^$\n]/gu, '').trim() : "",
        `'` + (issue.fields.labels || []).join(", "),
        getComponentsString(issue.fields.components),
        issue.fields.parent ? issue.fields.parent.key : "",
        issue.fields.parent ? issue.fields.parent.fields.summary : "",
        jiraBaseUrl + '/browse/' + issue.key,
        issue.fields.parent ? jiraBaseUrl + '/browse/' + issue.fields.parent.key : "",
        issue.fields.customfield_10019 ? issue.fields.customfield_10019 : "",
        issue.fields.parent ? parentRanks.get(issue.fields.parent.key) : ""
      ]);

      projectData.sort((a, b) => {
        const parentKeyA = a[7]; // Parent Key
        const parentKeyB = b[7];
        const issueRankA = a[11]; // Issue Rank
        const issueRankB = b[11];

        const parentRankA = parentRanks.get(parentKeyA) || "";
        const parentRankB = parentRanks.get(parentKeyB) || "";

        // Compare parent rank first
        if (parentRankA < parentRankB) return -1;
        if (parentRankA > parentRankB) return 1;

        // If parent ranks are equal, compare issue rank
        if (issueRankA < issueRankB) return -1;
        if (issueRankA > issueRankB) return 1;

        return 0;
      });

      allData = allData.concat(projectData);
    }

    dataSheet.clear(); // Clear previous data

    // Define headers and write them to the sheet
    const headers = ["Project Key", "Epic Key", "Summary", "Current Status", "Roadmap State", "Labels", "Components", "Parent Key", "Parent Summary", "Epic Link", "Parent Link", "Issue Rank", "Parent Rank"];
    dataSheet.appendRow(headers);

    // Write all data at once for better performance
    if (allData.length > 0) {
      dataSheet.getRange(2, 1, allData.length, headers.length).setValues(allData);
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