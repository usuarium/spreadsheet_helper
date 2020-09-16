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
