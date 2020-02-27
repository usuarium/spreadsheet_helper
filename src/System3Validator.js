var System3Validator = {
  
  validateFeastCommuneToTopics: function () {
    var invalids = [];
    
    for (var i = 0; i < System3.data.length; i++) {
      var row = System3.data[i]
  
      // feast
      var feastCellPointer = {row: i+1, column: 7}
      var communeCellPointer = {row: i+1, column: 8}
      
      if (row[feastCellPointer.column].length > 0 && row[communeCellPointer.column].length) {
        invalids.push({
          coord: feastCellPointer,
          message: 'You can\'t fill feast and commune column at the same time.'
        })
        invalids.push({
          coord: communeCellPointer,
          message: 'You can\'t fill feast and commune column at the same time.'
        })
      }
    }
    
    return invalids;
  },
  
  validateColumns: function () {
    
    System3Validator.validateTopics();
    
    var invalids = [];
    
    for (var i = 0; i < System3.data.length; i++) {
      var row = System3.data[i]
  
      // item
      var cellPointer = {row: i+1, column: 1}
      
      // type col
      var cellPointer = {row: i+1, column: 1}
      var type = row[cellPointer.column]
      if (type.length === 0) {
        // find value
      }
      else if (System3.values.types[type] === undefined) {
        invalids.push({
          coord: cellPointer,
          message: 'Invalid value for TYPE column: '+ type
        })
      }
      
      // 
    }
    
    return invalids;
  },

  validateHeaders: function () {
    var expectedHeaders = System3.getHeaderTemplate();
    
    var invalids = []
    
    for (var i = 0; i < expectedHeaders.length; i++) {
      if (expectedHeaders[i] !== System3.headers[i]) {
        invalids.push({
          row: 1,
          column: i,
          message: 'invalid header: "'+ System3.headers[i] + '" expected: "'+ expectedHeaders[i] +'"'
        })
      }
    }
    
    return invalids
  },
  
  determinateSheetVersion: function () {
    System3.sheetVersion = System3.headers[8] === 'TOPICS' ? 2 : 1;
  }
}

var System3ValidatorStartValidation = function () {
  ValidatorResultUI.showValidatorSidebar('System3 Validation Result');
  
  System3.resetTemporary();
  System3.loadSheetData();
  System3.determinateSheetVersion();
  
  var result = System3Validator.validateHeaders();
  
  if (result.length > 0) {
    ValidatorResultUI.addResults(result);
    return;
  }
  
  try {
    System3.migrateSheet();
  }
  catch (e) {
    return
  }
  
  System3Validator.validateColumns();
  
  // https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app.html#fetch(String)
  // UrlFetchApp.fetch("")
}