class ValidatorResultUI
{
  costructor() {
    this.errors = []
  }
  
  showValidatorSidebar(title, errors) {
    this.errors = errors

    let template = HtmlService.createTemplateFromFile('Sidebar')
    template.errors = errors
    template.title = title
    SpreadsheetApp.getUi().showSidebar(template.evaluate());
  }
}

const validatorResult = new ValidatorResultUI()

function goToError(col, row) {
  col *= 1
  row *= 1
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(row, col).activateAsCurrentCell();
}

var ValidatorResultUI_old = {
  errorTextStyle: null,
  normalTextStyle: null,
  
  initTextStyles: function () {
    if (ValidatorResultUI.errorTextStyle !== null) {
      return;
    }
    
    var styleBuilder = SpreadsheetApp.newTextStyle().setForegroundColor('white').setFontName('Ariel').setFontSize(10)
    ValidatorResultUI.errorTextStyle = styleBuilder.build()

    styleBuilder = SpreadsheetApp.newTextStyle().setForegroundColor('black').setFontName('Ariel').setFontSize(10)
    ValidatorResultUI.normalTextStyle = styleBuilder.build()
  },
  
  initDatabase: function () {
    if (ValidatorResultUI.database !== null) {
      return
    }
  },
  
  showValidatorSidebar: function (title, result) {
    ValidatorResultUI.errors = result

    let template = HtmlService.createTemplateFromFile('Sidebar')
    template.errors = result
    template.title = title
    SpreadsheetApp.getUi().showSidebar(template.evaluate());
  },
  
  clear: function () {
    ValidatorResultUI.resetHighlighting();
    ValidatorResultUI.errors = [];
    ValidatorResultUI.resetUI();
  },
  
  addResults: function (validationErrors) {
    validationErrors.forEach(function (error) {
      ValidatorResultUI.addResult(error)
    })
  },
  
  addResult: function (validationError) {
    if (validationError.coord === undefined) {
      throw new Error('message is coord')
    }
    
    if (validationError.message === undefined) {
      throw new Error('message is missing')
    }
    
    var range = sheet.getRange(validationError.coord.row, validationError.coord.column)
    ValidatorResultUI.highlightCell(range);
    ValidatorResultUI.errors.push(validationError);
  },
  
  highlightCell: function (range) {
    range.setTextStyle(ValidatorResultUI.errorTextStyle)
  },
  
  resetHighlighting: function () {
    var error = ValidatorResultUI.errors[errorIndex]    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    ValidatorResultUI.errors.forEach(function (error) {
      var range = sheet.getRange(error.coord.row, error.coord.column)
      ValidatorResultUI.resetHighlight(range)
    })
  },
  
  resetHighlight: function (range) {
    range.setTextStyle(ValidatorResultUI.normalTextStyle)
  },
}



