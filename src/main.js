var usuariumMenu;

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Usuarium')
    // .addSubMenu(SpreadsheetApp.getUi().createMenu('Calendar')
    //   .addItem('Validate Calendar', 'validateCalendar')
    //   .addItem('Add Rubrum column and fill it', 'calendarAddRubrum'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Index')
      .addItem('Validate', 'IndexValidatorStartValidation_')
    )
    .addSubMenu(SpreadsheetApp.getUi().createMenu('System 3')
      .addItem('Migrate Sheet', 'System3MigrateSheet_')
      .addItem('Fill Empty Values', 'System3FillEmptyValues_')
      .addItem('Validate', 'System3ValidatorStartValidation_')
    )
    .addToUi();
}

function System3MigrateSheet_() {
    let system3 = new System3()
    system3.loadSheetHeaders()
    system3.determinateSheetState()

    let result = system3.validator.validateHeaders()

    if (result.length > 0) {
        validatorResult.showValidatorSidebar('System3 Validator Result', result)
        return;
    }

    system3.migrateSheet()
}


function System3FillEmptyValues_() {
    let system3 = new System3()
    system3.loadSheetHeaders()
    system3.determinateSheetState()

    let result = system3.validator.validateHeaders()

    if (result.length > 0) {
        validatorResult.showValidatorSidebar('System3 Validator Result', result)
        return;
    }
    
    system3.fillEmptyFields()
}

function System3ValidatorStartValidation_() {
    let system3 = new System3()
    system3.loadSheetHeaders()
    system3.determinateSheetState()

    let result = system3.validator.validateHeaders()

    if (result.length > 0) {
        validatorResult.showValidatorSidebar('System3 Validator Result', result)
        return;
    }

    
}

function IndexValidatorStartValidation_() {
  const index = new IndexValidator()
  
  index.loadSheetData()
  index.loadIndexLabelData()
  let errors = index.validate()
  
  validatorResult.showValidatorSidebar('Index Validation Result', errors)
}
