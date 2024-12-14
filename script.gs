var ui = SpreadsheetApp.getUi();

ui.createMenu('LSR New')
  .addItem('⬇ Strava data', 'getStravaActivityData')
  .addItem('⬇ Intervals data', 'fillWellnessData')
  .addToUi();


function fillWellnessData() {
  const till = getDateString();
  const from = getDateString(21);
  const userId = readSettings("intervals_userId");
  const apiKey = readSettings("intervals_apikey");

  writeToSheet(
    fetchIntervalsWellnessData(
      userId, apiKey,
      from, till
    )
  );
}

function readSettings(key) {
  const SETTINGS_SHEET = "Settings";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET);

  if (!settingsSheet) {
    ss.insertSheet().setName(SETTINGS_SHEET);
  }

  for(var i = 1; i < 50; i++) {
    const settingKey = settingsSheet.getRange(i, 1).getValue();
    if (settingKey === key) {
      return settingsSheet.getRange(i, 2).getValue();
    }
  }
  throw Error("Setting " + key + " not found!");
}

function fetchIntervalsWellnessData(userId, apiKey, dateFrom, dateTo) {
      var endpoint = 'https://intervals.icu/api/v1/athlete/' + userId + '/wellness';
      var params = '?oldest='+dateFrom+'&newest='+dateTo;
    
      var options = {
        method : 'GET',
        muteHttpExceptions: true,
        headers: {
          Authorization: 'Basic ' + Utilities.base64Encode("API_KEY:" + apiKey),
        }
      };
      
      var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options));
      Logger.log(response);

      return response; 
}

function writeToSheet(intervalsData) {
      var byDate = {};
      intervalsData.map(data => {
        byDate[data.id.slice(-2)] = data;
      });

      var ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet = ss.getActiveSheet();

      var row = sheet.getActiveCell().getRow();
      var col = sheet.getActiveCell().getColumn();
      
      const dateRow = findDateRow(row, col);
      const wellnessRow = dateRow + 1;
      sheet.insertRowAfter(dateRow);

      for(colIdx = col; colIdx < 20; colIdx++) {
          const day = ("0" + sheet.getRange(dateRow, colIdx).getValue()).slice(-2);
          const wellness = byDate[day];
          if (wellness) {
            sheet.getRange(wellnessRow, colIdx).setBackgroundRGB(255,255,0).setFontColor("red");
            sheet.getRange(wellnessRow, colIdx).setValue("hrv: " + wellness.hrv + " rhr: " + wellness.restingHR);  
          }
      } 
}
