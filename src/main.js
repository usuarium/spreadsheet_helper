var usuariumMenu;

let version = 'DEV' // current version

function onOpen() {
    SpreadsheetApp.getUi().createMenu('Usuarium')
        .addSubMenu(
            SpreadsheetApp.getUi().createMenu('Index')
                .addItem('Validate', 'IndexValidatorStartValidation_')
        )
        .addSubMenu(
            SpreadsheetApp.getUi().createMenu('System 3')
                .addItem('Migrate Sheet', 'System3MigrateSheet_')
                .addItem('Fill Empty Values', 'System3FillEmptyValues_')
                .addItem('Spanish Inquisition (validate)', 'System3ValidatorStartValidation_')
                .addItem(`Version: ${version}`, 'System3Version_')
        )
    .addToUi();
}

function System3Version_() {
    let system3 = new System3()
    
    Logger.log({rows: system3.sheetWrapper.getActiveRowsCount(), first: system3.sheetWrapper.getFirstSelectedRow() })
}

function System3MigrateSheet_() {
    let system3 = new System3()
    system3.loadSheetHeaders()

    let result = system3.validator.validateHeaders()

    if (result.length > 0) {
        validatorResult.showValidatorSidebar('System3 Validator Result', result)
        return;
    }

    if (system3.hasMissingColumn()) {
        system3.addMissingColumns()
    }
}


function System3FillEmptyValues_() {
    let system3 = new System3()
    system3.loadSheetHeaders()

    let result = system3.validator.validateHeaders()

    if (result.length > 0) {
        validatorResult.showValidatorSidebar('System3 Validator Result', result)
        return;
    }

    if (system3.hasMissingColumn()) {
        system3.addMissingColumns()
    }
    
    system3.loadLists()
    system3.fillEmptyFields()
}

function System3ValidatorStartValidation_() {
    let system3 = new System3()
    system3.loadSheetHeaders()

    let result = system3.validator.validateHeaders()

    if (result.length > 0) {
        validatorResult.showValidatorSidebar('System3 Validator Result', result)
        return;
    }
    
    if (system3.hasMissingColumn()) {
        // todo feedback
        return
    }
    
    system3.loadLists()
    let errors = system3.validate()
    
    if (errors.length > 0) {
        validatorResult.showValidatorSidebar('Conspectus Validation Result', errors)
    }
    else {
        var ui = SpreadsheetApp.getUi();
        var response = ui.alert('Validation result', 'The given data range is valid.', ui.ButtonSet.OK);
    }
}

function IndexValidatorStartValidation_() {
    const index = new IndexValidator()

    index.loadSheetData()
    let errors = index.validate()

    if (errors.length > 0) {
        validatorResult.showValidatorSidebar('Index Validation Result', errors)
    }
    else {
        var ui = SpreadsheetApp.getUi();
        var response = ui.alert('Validation result', 'The given data range is valid.', ui.ButtonSet.OK);
    }
}
