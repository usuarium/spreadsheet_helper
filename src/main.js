var usuariumMenu;

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Usuarium')
    // .addSubMenu(SpreadsheetApp.getUi().createMenu('Calendar')
    //   .addItem('Validate Calendar', 'validateCalendar')
    //   .addItem('Add Rubrum column and fill it', 'calendarAddRubrum'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Index')
      .addItem('Validate', 'IndexValidatorStartValidation'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('System 3')
      .addItem('Validate', 'System3ValidatorStartValidation'))
    .addToUi();
}


var System3ValidatorStartValidation = function () {
  let system3 = new System3()
  let result = system3.validator.validateHeaders()

  if (result.length > 0) {
    validatorResult.showValidatorSidebar('System3 Validator Result', result)
    return;
  }
}