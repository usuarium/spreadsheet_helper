var usuariumMenu;

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Usuarium')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Calendar')
      .addItem('Validate Calendar', 'validateCalendar')
      .addItem('Add Rubrum column and fill it', 'calendarAddRubrum'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Index')
      .addItem('Validate', 'IndexValidatorStartValidation'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('System 3')
      .addItem('Validate and Fill', 'System3ValidatorStartValidation'))
    .addToUi();
}
