var usuariumMenu;

let version = 'b3a1ef5' // current version

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
                .addItem('Validate', 'System3ValidatorStartValidation_')
                .addItem(`Version: ${version}`, 'System3Version_')
        )
    .addToUi();
}

function System3Version_() {
    
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
}

function IndexValidatorStartValidation_() {
    const index = new IndexValidator()

    index.loadSheetData()
    index.loadIndexLabelData()
    let errors = index.validate()

    validatorResult.showValidatorSidebar('Index Validation Result', errors)
}
