class IndexValidator
{
    constructor() {
        this.indexLabelCache = {}

        this.sheetData = []
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

    // sheet data
    getSheetData() {
        var dataRange = this.getActiveSheet().getDataRange();
        return this.activeSheet.getRange(2, 1, dataRange.getLastRow()-1, 12).getValues();
    }

    loadSheetData() {
        this.sheetData = this.getSheetData()
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
		
        let labels = indexLabelRaw.split(',')
        return labels.map((label) => {
            return label.trim()
        })
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
        let errors = []
    
        this.loadSheetData()
        this.loadBooksData()
        this.loadIndexLabelData()
		
        let lastRowCoord = this.getLastRowCoord()
		
        if (lastRowCoord < 2) {
            errors.push({
                row: 1,
                col: 1,
                message: `Missing rows?`
            })
            return errors
        }
		
        for (let rowCoord = 2; rowCoord <= lastRowCoord; ++rowCoord) {
            let row = this.getRowAtCoord(rowCoord)
			
            if (row.length != 12) {
                errors.push({
                    row: rowCoord,
                    col: 1,
                    message: `Missing fields?`
                })
                continue
            }
      
            let bookId = row[0]
            if (!this.isValidBookIdValue(bookId)) {
                errors.push({
                    row: rowCoord,
                    col: 1,
                    message: `Missing book id or invalid format`
                })
            }
            else if (!this.isValidBookId(bookId)) {
                errors.push({
                    row: rowCoord,
                    col: 1,
                    message: `Unknown book id: ${bookId}`
                })
            }
			
            let digitalPageNumber = row[2]
            if (!this.isValidDigitalPageNumberValue(digitalPageNumber)) {
                errors.push({
                    row: rowCoord,
                    col: 3,
                    message: `Missing digital page number or invalid format`
                })
            }
            else if (!this.isValidDigitalPageNumber(digitalPageNumber, bookId)) {
                errors.push({
                    row: rowCoord,
                    col: 3,
                    message: `Invalid digital page number in book: ${bookId}`
                })
            }
			
            let indexLabelRaw = row[5]
            let indexLabels = this.splitIndexLabels(indexLabelRaw)
			
            if (indexLabels.length === 0) {
                continue
            }
      
            for (let label of indexLabels) {
                if (this.isValidIndexLabel(label)) {
                    continue
                }
        
                let castRightLabel = this.findRightCaseLabel(label)
                if (castRightLabel !== null) {
                    this.setIndexLabel(indexLabelRaw.replace(label, castRightLabel), rowCoord)
                    continue
                }
        
                errors.push({
                    row: rowCoord,
                    col: 6,
                    message: `Unknown index label: ${label}`
                })
            }
        }
    
        return errors
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

