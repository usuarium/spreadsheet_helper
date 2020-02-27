var System3 = {
  values: {
    types: ['MISS', 'OFF', 'RIT']
  },
  headers: [],
  data: [],
  
  resetTemporary: function () {
    System3.sheetVersion = 1;
    System3.headers = [];
    System3.data = [];
  },
    
  loadSheetData: function () {
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getRange('A1:V1');
    System3.headers = range.getValues()[0];
    
    var dataRange = sheet.getDataRange();
    System3.data = sheet.getRange(2, 1, dataRange.getLastRow()-1, 22).getValues();
  },
  
  getHeaderTemplate: function () {
    var headers = [
      'ID',
      'TYPE',
      'PART',
      'SEASON/MONTH',
      'WEEK',
      'DAY',
      'FEAST',
      'COMMUNE/VOTIVE',
      'MASS/HOUR',
      'CEREMONY',
      'MODULE',
      'SEQUENCE',
      'RUBRICS',
      'LAYER',
      'GENRE',
      'SERIES',
      'ITEM',
      'PAGE NUMBER (DIGITAL)',
      'PAGE NUMBER (ORIGINAL)',
      'REMARK',
      'MADE BY'
    ];

    if (System3.sheetVersion === 1) {
      return headers;
    }
      
    return headers.splice(8, 0, 'TOPICS');
  },
  
  migrateSheet: function () {
    if (System3.sheetVersion !== 1) {
      return
    }

    var result = System3Validator.validateFeastCommuneToTopics()
    if (result.length > 0) {
      ValidatorResultUI.addResults(result)
      throw new Error('migrate failed due errors')
    }
    
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.insertColumnAfter(8)
    sheet.getRange('I1').setValue("TOPICS")
    
    System3.resetTemporary();
    System3.loadSheetData();
    System3.determinateSheetVersion();
  }
}


