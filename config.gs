function init_config() {
  SpreadsheetApp.getUi().createMenu('Config')
    .addItem('load scripts', 'readScriptFromURL')
    .addItem('load from cell', 'readScriptFromSheetCell')
    .addToUi();
}

init_config();

function readScriptFromSheetCell() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet = ss.getActiveSheet();

      const script = sheet.getActiveCell().getValue();
      
      eval(script);    
}

function readScriptFromURL() {
      const endpoint = 'https://raw.githubusercontent.com/lawreenas/lsr-sheet-helper/refs/heads/main/script.gs';
    
      const options = {
        method : 'GET',
        muteHttpExceptions: true,
      };
      
      const script = UrlFetchApp.fetch(endpoint, options);

      eval(script.getContentText());
}
