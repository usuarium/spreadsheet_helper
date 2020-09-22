class ValidatorResultUI
{
    costructor() {
        this.errors = []
    }

    showValidatorSidebar(title, errors) {
        this.errors = errors

        let template = HtmlService.createTemplateFromFile('Sidebar')
        template.errors = ValidatorResultUI.toErrorObject(errors)
        template.title = title
        SpreadsheetApp.getUi().showSidebar(template.evaluate());
    }
    
    static toErrorObject(errors) {
        let buffer = []
        
        if (errors.length === 0) {
            return []
        }
        
        if (errors[0] instanceof CellValidationError) {
            for (let error of errors) {
                buffer.push(error.errorObject)
            }
            
            return buffer
        }
        
        return errors
    }
}

const validatorResult = new ValidatorResultUI()
