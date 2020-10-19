class IndexValidator
{
    constructor() {
        this.indexLabelCache = {}
        this.sheetWrapper = new SheetWrapper()

        this.data = []
        this.indexLabels = []
        this.books = []

        this.activeSheet = null
    }
	
    // active sheet data
    getActiveSheet() {
        if (!this.activeSheet) {
            this.activeSheet = SpreadsheetApp.getActiveSheet()
        }
        
        return this.activeSheet
    }

    loadSheetData() {
        this.data = this.sheetWrapper.getDataRows(12)
    }

    // index label data
    getIndexLabelData() {
        return UsuariumAPIClient.fetchIndexLabels()
    }
    
    loadIndexLabelData() {
        this.indexLabels = this.getIndexLabelData()
    }
	
    // books data
    getBooksData() {
        return UsuariumAPIClient.fetchBooks()
    }

    loadBooksData() {
        this.books = this.getBooksData()
    }
	
    getBookWithId(bookId) {
        if (!Number.isInteger(bookId)) {
            bookId = bookId * 1
        }
		
        for (let book of this.books) {
            if (book.id === bookId) {
                return book
            }
        }
		
        return null
    }
	
    // data and coord getters
    getRowAtCoord(rowCoord) {
        return this.sheetData[rowCoord-2]
    }
	
    getLastRowCoord() {
        return this.sheetData.length + 1
    }
	
    splitIndexLabels(indexLabelRaw) {
        if (indexLabelRaw === null || indexLabelRaw === undefined) {
            return []
        }
		
        if (Number.isInteger(indexLabelRaw)) {
            indexLabelRaw = indexLabelRaw.toString()
        }
        
        if (indexLabelRaw.length === 0) {
            return []
        }
		
        return indexLabelRaw.split(',').map(label => label.trim())
    }
	
    findRightCaseLabel(labelName) {
        for (let label of this.indexLabels) {
            if (label.name.toUpperCase() == labelName.toUpperCase()) {
                return label.name
            }
        }
    
        return null
    }
	
    setIndexLabel(indexLabelRaw, rowCoord) {
        this.getActiveSheet().getRange(rowCord, 6).setValue(caseRightLabel)
    }
	
    // validater methods
    validate() {
        this.loadSheetData()
        this.loadBooksData()
        this.loadIndexLabelData()
		
        let errors = this.validateRows(this.data)
        
        if (errors.length > 0) {
            let rows = []
            for (let error of errors) {
                rows.push(error.errorObject.row)
            }
            this.sheetWrapper.highlightRowsAt(rows)
        }
        
        return errors
    }
    
    validateRows(rows) {
        let checkSubset = this.sheetWrapper.getActiveRowsCount() > 1
        let firstRowIndex = 0
        let lastRowIndex = 0
        
        if (checkSubset === true) {
            firstRowIndex = this.sheetWrapper.getFirstSelectedRow() - 2
            lastRowIndex = this.sheetWrapper.getLastSelectedRow() - 2
        }
        
        let errors = []
        for (let rowIndex in rows) {
            if (checkSubset && (rowIndex < firstRowIndex || rowIndex > lastRowIndex)) {
                continue
            }
            
            let row = rows[rowIndex]
            let rowErrors = this.validateRow(row)
            
            if (rowErrors !== true) {
                for (let error of rowErrors) {
                    error.rowIndex = rowIndex
                    errors.push(error)
                }
            }
        }
        
        return errors
    }
    
    validateRow(row) {
        let errors = []
        
        if (row.length != 12) {
            let error = new CellValidationError('Missing fields?')
            error.cellIndex = 0
            return [error]
        }
  
        let bookId = this.getRowDataWithName('shelfmark', row)
        if (!this.isValidBookIdValue(bookId)) {
            let error = new CellValidationError('Missing book id or invalid format')
            error.cellIndex = this.getIndexWithFieldName('shelfmark')
            errors.push(error)
        }
        else if (!this.isValidBookId(bookId)) {
            let error = new CellValidationError(`Unknown book id: ${bookId}`)
            error.cellIndex = this.getIndexWithFieldName('shelfmark')
            errors.push(error)
        }
		
        let digitalPageNumber = this.getRowDataWithName('digital_page_number', row)
        Logger.log(digitalPageNumber)
        if (!this.isValidDigitalPageNumberValue(digitalPageNumber)) {
            let error = new CellValidationError(`Missing digital page number or invalid format`)
            error.cellIndex = this.getIndexWithFieldName('digital_page_number')
            errors.push(error)
        }
        else if (!this.isValidDigitalPageNumber(digitalPageNumber, bookId)) {
            let error = new CellValidationError(`Invalid digital page number ${digitalPageNumber} in book: ${bookId}`)
            error.cellIndex = this.getIndexWithFieldName('digital_page_number')
            errors.push(error)
        }
		
        let indexLabelRaw = this.getRowDataWithName('labels', row)
        
        if (indexLabelRaw === '-') {
            indexLabelRaw = ''
        }
        
        let indexLabels = this.splitIndexLabels(indexLabelRaw)
        Logger.log({l: indexLabels.length})
        if (indexLabels.length !== 0) {
            for (let label of indexLabels) {
                if (this.isValidIndexLabel(label)) {
                    continue
                }
    
                let castRightLabel = this.findRightCaseLabel(label)
                if (castRightLabel !== null) {
                    this.setRowDataWithName('labels', indexLabelRaw.replace(label, castRightLabel), row)
                    continue
                }
    
                let error = new CellValidationError(`Unknown index label: ${label}`)
                error.cellIndex = this.getIndexWithFieldName('labels')
                errors.push(error)
            }
        }
        
        return errors.length === 0 ? true : errors
    }
    
    
    getIndexWithFieldName(fieldName) {
        switch (fieldName) {
            case 'shelfmark':
                return 0;           // SHELFMARK
                
            case 'source':
                return 1;           // SOURCE

            case 'digital_page_number':
                return 2;          // PAGE NUMBER (DIGITAL)

            case 'original_page_number':
                return 3;          // PAGE NUMBER (ORIGINAL)

            case 'title':
                return 4;          // PAGE NUMBER (ORIGINAL)

            case 'labels':
                return 5;          // PAGE NUMBER (ORIGINAL)

            case 'illustration':
                return 6;          // PAGE NUMBER (ORIGINAL)

            case 'vernacular':
                return 7;          // PAGE NUMBER (ORIGINAL)

            case 'musical_notation':
                return 8;          // PAGE NUMBER (ORIGINAL)

            case 'made_by':
                return 9;          // MADE BY

            case 'status':
                return 10;          // MADE BY

            case 'sequence':
                return 11;          // MADE BY

            case 'page_link':
                return 12;          // PAGE NUMBER (ORIGINAL)
        }
        
        throw new Error(`invalid field name: ${fieldName}`)
    }
    
    getRowDataWithName(fieldName, row) {
        let index = this.getIndexWithFieldName(fieldName)
        
        if (index === null) {
            throw new Error('invalid index')
        }
        
        let data = row[index] === undefined ? '' : row[index]
        
        return data === null ? '' : data
    }
    
    setRowDataWithName(fieldName, value, row) {
        let index = this.getIndexWithFieldName(fieldName)
        
        if (index === null) {
            throw new Error('invalid index')
        }
        
        row[index] = value
    }

    isValidBookIdValue(bookId) {
        if (bookId === null || bookId === undefined) {
            return false
        }
		
        // string and trimmed and 0-9
        if (typeof bookId === 'string' && bookId.trim().match(/^[0-9]+$/)) {
            return true
        }
		
        // numeric
        if (Number.isInteger(bookId)) {
            return true
        }
		
        return false
    }

    isValidBookId(bookId) {
        return this.getBookWithId(bookId) !== null
    }

    isValidDigitalPageNumberValue(digitalPageNumber) {
        if (digitalPageNumber === null || digitalPageNumber === undefined) {
            return false
        }
		
        // string and trimmed and 0-9
        if (typeof digitalPageNumber === 'string' && digitalPageNumber.trim().match(/^[0-9]+$/)) {
            return true
        }
		
        // numeric
        if (Number.isInteger(digitalPageNumber)) {
            return true
        }
		
        return false
    }
	
    isValidDigitalPageNumber(digitalPageNumber, bookId) {
        let bookData = this.getBookWithId(bookId)
		
        if (bookData === null) {
            return false
        }
		
        digitalPageNumber *= 1
		
        return digitalPageNumber <= bookData.number_of_pages
    }

    isValidIndexLabel(labelName) {
        for (let label of this.indexLabels) {
            if (label.name == labelName) {
                return true
            }
        }
    
        return false
    }
}

