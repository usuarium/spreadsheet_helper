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
