function calendarAddRubrum() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange('A1:R1');
  var values = range.getValues();
  
  if (values[0].indexOf('RUBRUM') === -1) {
    var indexOfTitle = values[0].indexOf('ORIGINAL TITLE');
    
    sheet.insertColumnBefore(indexOfTitle+1);
    
    sheet.getRange(1, indexOfTitle+1).setValue('RUBRUM')
  }
  
  var dataRange = sheet.getDataRange();
  
  var originalTitleRange = sheet.getRange(2, 10, dataRange.getLastRow()-1, 1);
  var rubrumRange = sheet.getRange(2, 9, dataRange.getLastRow()-1, 1);
  
  var fontColorData = originalTitleRange.getFontColors();
  var originalTitleData = originalTitleRange.getValues();

  var contents = [];
  for (var i in fontColorData) {
    if (originalTitleData[i][0].length === 0 || fontColorData[i][0] === '#000000') {
      contents.push(['']);
      continue;
    }
    
    if (fontColorData[i][0] !== '#000000') {
      contents.push(["RUBRUM"]);
    }
  }
  
  rubrumRange.setValues(contents);
}

function validateCalendar() {
  showSidebar('Calendar Validation Result');
}
