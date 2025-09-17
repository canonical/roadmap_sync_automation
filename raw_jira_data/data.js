function readSheetToNestedDict() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("25.10_data");
  const data = sheet.getDataRange().getValues();
  
  const headers = data[0];
  const result = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const projectKey = row[0];
    const epicKey = row[1];

    if (!result[projectKey]) {
      result[projectKey] = {};
    }

    result[projectKey][epicKey] = {
      [headers[2]]: row[2] || "",
      [headers[3]]: row[3] || "",
      [headers[4]]: row[4] || ""
    };
  }

  const json = JSON.stringify(result, null, 2);
  DriveApp.createFile('nested_dict_new.json', json);
  return result;
}