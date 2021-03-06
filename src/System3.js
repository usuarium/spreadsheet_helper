class System3
{
    constructor() {
        this.headers = []
        this.data = []
        this.hasPageLinkColumn = false
        this.hasTopicsColumn = false
        this.hasShelfmarkColumn = false
        this.hasSourceColumn = false
    
        this.colIndexes = {}
        this.rowPointer = 0
        this.validator = new System3Validator(this)
        this.sheetWrapper = new SheetWrapper()
        
        this.indexLabels = []
        this.topics = []
        
        this.knownHeaders = [
            'ID',
            'SHELFMARK',
            'SOURCE',
            'TYPE',
            'PART',
            'SEASON/MONTH',
            'WEEK',
            'DAY',
            'FEAST',
            'COMMUNE/VOTIVE',
            'TOPICS',
            'MASS/HOUR',
            'CEREMONY',
            'MODULE',
            'SEQUENCE',
            'RUBRICS',
            'LAYER',
            'GENRE',
            'SERIES',
            'ITEM',
            'STANDARD ITEM',
            'PAGE NUMBER (DIGITAL)',
            'PAGE NUMBER (ORIGINAL)',
            'PAGE LINK',
            'REMARK',
            'MADE BY',
        ]
    }
    
    loadLists() {
        this.indexLabels = UsuariumAPIClient.fetchIndexLabels()
        this.topics =  UsuariumAPIClient.fetchTopics()
    }
  
    getDocumentName() {
        return SpreadsheetApp.getActiveSpreadsheet().getName()
    }
  
    getBookIdentifier(documentName) {
        if (documentName.indexOf(' - ') === -1) {
            return null
        }
        
        return documentName.split(' - ')[1] * 1
    }
  
    loadSheetHeaders() {
        this.headers = this.sheetWrapper.getHeaders()
    }
  
    loadSheetData() {
        this.data = this.sheetWrapper.getDataRows()
    }
    
    writeData() {
        let currentData = null
        let changedDataCoordinates = []
        
        currentData = this.sheetWrapper.getDataRows()
        
        for (let rowIndex in this.data) {
            for (let colIndex in this.data[rowIndex]) {
                if (this.data[rowIndex][colIndex] != currentData[rowIndex][colIndex]) {
                    changedDataCoordinates.push({row: rowIndex + 1, col: colIndex + 1})
                }
            }
        }
        
        this.doWriteData(this.data)
        
        // if (changedDataCoordinates.length > 0) {
        //     this.markBackgroundsAt(changedDataCoordinates)
        // }
    }
    
    doWriteData(data, indexToWrite) {
        if (indexToWrite === undefined) {
            indexToWrite = 0
        }
        
        let coords = {row: 2, col: indexToWrite + 1, rows: data.length, cols: data[0].length}
        this.sheetWrapper.setValuesAt(data, coords)
    }
    
    markBackgroundsAt(coordinates) {
        for (let coordinate of coordinates) {
            this.sheetWrapper.setBackgroundColorAt('#c9daf8', {row: 2 + coordinate.row, col: coordinate.col, rows: 1, cols: 1})
        }
    }
  
    getSheetHeaders() {
        let sheet = SpreadsheetApp.getActiveSheet();
        let dataRange = sheet.getDataRange()
        let range = sheet.getRange(1, 1, 1, dataRange.getLastColumn());
        return range.getValues()[0];
    }
  
    getHeadersTemplate() {
        // v1 cols: 
        // PART	SEASON/MONTH	WEEK	DAY	FEAST	MASS/HOUR	SEQUENCE	LAYER	GENRE	SERIES	ITEM	PAGE NUMBER (DIGITAL)	PAGE NUMBER (ORIGINAL)	REMARK
        let headers = [
            'ID',
            'TYPE',
            'PART',
            'SEASON/MONTH',
            'WEEK',
            'DAY',
            'FEAST',
            'COMMUNE/VOTIVE',
            'MASS/HOUR',
            'CEREMONY',
            'MODULE',
            'SEQUENCE',
            'RUBRICS',
            'LAYER',
            'GENRE',
            'SERIES',
            'ITEM',
            'STANDARD ITEM',
            'PAGE NUMBER (DIGITAL)',
            'PAGE NUMBER (ORIGINAL)',
            'REMARK',
            'MADE BY'
        ];
    
        if (this.hasShelfmarkColumn) {
            ArrayHelper.insertItemAfter(headers, 'SHELFMARK', 'ID')
        }

        if (this.hasSourceColumn) {
            ArrayHelper.insertItemAfter(headers, 'SOURCE', 'SHELFMARK')
        }

        if (this.hasTopicsColumn) {
            ArrayHelper.insertItemAfter(headers, 'TOPICS', 'COMMUNE/VOTIVE')
        }

        if (this.hasPageLinkColumn) {
            ArrayHelper.insertItemAfter(headers, 'PAGE LINK', 'PAGE NUMBER (ORIGINAL)')
        }
      
        return headers;
    }

    insertHeaderAfter(header, after) {
        let newHeaderIndex = ArrayHelper.insertItemAfter(this.headers, header, after)
      
        if (newHeaderIndex === false) {
            return false
        }
    
        this.sheetWrapper.insertColumnAfter(newHeaderIndex + 1)
        this.sheetWrapper.setHeaderValueAtPosition(newHeaderIndex + 2, header)
    }
    
    hasMissingColumn() {
        for (let header of this.knownHeaders) {
            if (this.headers.indexOf(header) === -1) {
                return true
            }
        }
        
        return false
    }
  
    addMissingColumns() {
        let newHeaders = this.headers
        for (let headerIndex in this.knownHeaders) {
            let knownHeader = this.knownHeaders[headerIndex]
            let currentHeaderAtPosition = this.headers[headerIndex]
            
            if (currentHeaderAtPosition === knownHeader) {
                continue
            }
            
            if (this.headers.indexOf(knownHeader) > -1) {
                // exists but wrong position
            }
            else {
                let insertPosition = this.knownHeaders.indexOf(knownHeader)

                // not exitsts, create it
                this.sheetWrapper.insertColumnAfter(insertPosition)
                this.sheetWrapper.setHeaderValueAtPosition(insertPosition+1, knownHeader)
                
                newHeaders.splice(insertPosition, 0, knownHeader)
            }
        }
        
        this.headers = newHeaders
    }
  
    loadDocumentName() {
        this.documentName = this.sheetWrapper.getDocumentName()
    }
  
    migrateShelfmark() {
        let documentName = this.sheetWrapper.getDocumentName()
        let shelfmark = this.getBookIdentifier(documentName)
        
        if (shelfmark === null || shelfmark <= 10000) {
            this.data = this.fillShelfmark(this.data, shelfmark)
        }
        else {
            let bookId = UsuariumAPIClient.fetchBookId(shelfmark)
            documentName = documentName.replace(` - ${shelfmark}`, ` - ${bookId}`)
            this.sheetWrapper.setDocumentName(documentName)
            this.data = this.fillShelfmark(this.data, bookId)
        }

        this.writeData()
    }
  
    fillShelfmark(rows, newShelfmark) {
        if (!this.hasShelfmarkColumn) {
            return rows
        }
        
        // data
        let shelfmarks = []
        for (let row of rows) {
            let shelfmark = this.getRowDataWithName('shelfmark', row)
            
            if (shelfmark.length === 0) {
                continue
            }
            
            shelfmark *= 1
            
            if (shelfmark < 10000) {
                continue
            }
            
            if (shelfmarks.indexOf(shelfmark) === -1) {
                shelfmarks.push(shelfmark)
            }
        }
        
        if (shelfmarks.length >= 1) {
            // multiple
            let shelfmarkTransformer = {}
            let hasOldShelfmark = false
            for (let shelfmark of shelfmarks) {
                if (shelfmark >= 10000) {
                    hasOldShelfmark = true
                    
                    shelfmarkTransformer[shelfmark] = UsuariumAPIClient.fetchBookId(shelfmark)
                }
            }
            
            if (hasOldShelfmark === false) {
                return rows
            }
            
            for (let row of rows) {
                let shelfmark = this.getRowDataWithName('shelfmark', row)
            
                if (shelfmark.length === 0) {
                    continue
                }
            
                shelfmark *= 1
                
                if (shelfmark < 10000) {
                    continue
                }
                
                if (shelfmarkTransformer[shelfmark] === undefined) {
                    continue
                }
                
                this.setRowDataWithName('shelfmark', shelfmarkTransformer[shelfmark], row)
            }
        }
        
        if (newShelfmark === null) {
            return rows
        }
        
        for (let row of rows) {
            this.setRowDataWithName('shelfmark', newShelfmark, row)
        }
        
        return rows
    }
    
    migrateGenre() {
        let indexOfGenre = this.headers.indexOf('GENRE')

        let genres = this.sheetWrapper.getValuesAtColumn(indexOfGenre+1)
        
        let genreChanged = false
        
        for (let rowIndex in genres) {
            let genre = genres[rowIndex][0]
            
            if (!genre || genre.length === 0) {
                continue
            }
            
            let newGenre = genre.replace(/[0-9]/, '')
            
            if (newGenre !== genre) {
                genreChanged = true
            }
            
            genres[rowIndex][0] = newGenre
        }
        
        if (genreChanged === false) {
            return
        }
        
        let coords = {row: 2, col: indexOfGenre+1, rows: genres.length, cols: 1}
        this.sheetWrapper.setValuesAt(genres, coords)
    }
  
    fillPageLinks() {
        let sheet = SpreadsheetApp.getActiveSheet();
        let firstPageLinkCellValue = sheet.getRange(2, 21).getValue()
        if (firstPageLinkCellValue && firstPageLinkCellValue.length > 0) {
            return
        }

        let dataRange = sheet.getDataRange()
        return sheet.getRange(2, 21, dataRange.getLastRow()-1, 1).setFormulaR1C1('=PAGELINK(R[0]C[-2])');
    }
    
    getValidValues(fieldName, dependentValues) {
        if (fieldName === 'type') {
            return ['MISS', 'OFF', 'RIT']
        }
        else if (fieldName === 'part') {
            let missValues = ['T', 'S', 'C', 'V']
            let offValues = ['PS', 'T', 'S', 'C', 'V']
            let ritValue = ['', 'O', 'T', 'S', 'V', 'C']
            
            let value = {
                'MISS': missValues,
                'OFF': offValues,
                'RIT': ritValue,
            }[dependentValues.type]
            
            return value === undefined ? null : value
        }
        else if (fieldName === 'season_month') {
            let temporalValues = ['Adv', 'Nat', 'Ep', 'LXX', 'Qu', 'Pasc', 'Pent', 'Trin', 'Gen']
            let sanctoralValues = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            let psValues = ['Gen', 'Ep', 'Trin']
            
            let value = {
                'T': temporalValues,
                'PS': psValues,
                'S': sanctoralValues,
                'C': '',
                'V': ['', ...temporalValues, ...sanctoralValues],
                'O': ['']
            }[dependentValues.part]

            return value === undefined ? null : value
        }
        else if (fieldName === 'week') {
            let temporalValues = ['H0', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'H7', 'H8', 'H9', 'H10', 'H11', 'H12', 'H13', 'H14', 'H15', 'H16', 'H17', 'H18', 'H19', 'H20', 'H21', 'H22', 'H23', 'H24', 'H25', 'H26', 'H27', 'H28', 'QuT', 'Gen']
            
            let value = {
                'T': temporalValues,
                'PS': 'Gen',
                'S': '',
                'C': '',
                'V': [...temporalValues, ''],
                'O': ['']
            }[dependentValues.part]

            return value === undefined ? null : value
        }
        else if (fieldName === 'day') {
            let temporalValues = ['D', 'ff', 'f2', 'f3', 'f4', 'f5', 'f6', 'S', 'Gen', 'F']
            let value = {
                'T': temporalValues,
                'PS': temporalValues,
                'S': new RegExp(/^([0-9]|[12]\d|3[01])$/),
                'C': '',
                'V': [...temporalValues, ''],
                'O': ['']
            }[dependentValues.part]

            return value === undefined ? null : value
        }
        else if (fieldName === 'mass_hour') {
            let value = {
                'MISS': ['', 'M1', 'M2', 'M3'],
                'OFF': ['V1', 'C1', 'M', 'N1', 'N2', 'N3', 'L', 'I', 'I+', 'III', 'VI', 'IX', 'V', 'V2', 'C', 'C2'],
                'RIT': '',
            }[dependentValues.type]
            
            return value === undefined ? null : value
        }
        else if (fieldName === 'ceremony') {
            let value = {
                'MISS': ['', 'Mass Propers'],
                'OFF': ['', 'Office Propers'],
                'RIT': '',
            }[dependentValues.type]
            
            return value === undefined ? null : value
        }
        else if (fieldName === 'sequence') {
            return /^\d+$/
        }
        else if (fieldName === 'layer') {
            return ['S/C', 'G/A', 'L', 'R']
        }
        else if (fieldName === 'genre') {
            let type = dependentValues.type
            let layer = dependentValues.layer
            
            let value = this.getAllGenres()[type][layer]
            
            return value === undefined ? null : value
        }
        else if (fieldName === 'series') {
            return ['', 'P', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14']
        }

        // https://mail.google.com/mail/u/0/#starred/FMfcgxwDsFgmnPzVSncRPBlxPMNRRksr
    }
    
    getAllGenres() {
        let missSC = [
            // romana
            'Or', 'Coll', 'Secr', 'Postcomm', 'Superpop', 'VD', 'Infracan', 'BenP',
            // ambrosiana
            'Litaniae', 'Super populum', 'Super sindonem', 'Super oblatam', 'Praefatio', 'Post communionem',
            // mozarabica
            'Oratio post officium', 'Missa', 'Alia oratio', 'Post nomina', 'Ad pacem', 
            'Illatio', 'Post Sanctus', 'Post Pridie', 'Ad orationem dominicam', 'Benedictio', 'Post communionem',
            // gallicana
            'Praefatio', 'Collectio', 'Post prophetiam', 'Post precem', 'Post Aius', 
            'Ante nomina', 'Super oblata', 'Post nomina', 'Ad pacem', 'Secreta', 'Contestatio', 
            'Super munera', 'Immolatio', 'Post Sanctus', 'Post secreta', 'Post mysterium', 
            'Ante orationem dominicam', 'Post orationem dominicam', 'Benedictio populi', 
            'Post communionem', 'Post Eucharistiam', 'Consummatio missae', 'Consecratio',
        ]
        
        let offSC = [
            'Cap', 'Or',
        ]
        
        let possibleValues = {
            'MISS': {
                'G/A': [
                    // romana
                    'Intr', 'IntrV', 'IntrTrop', 'Tr', 'Gr', 'GrV', 'H', 'All', 'AllV', 'Seq', 'Off', 'OffV', 
                    'OffTrop', 'Comm', 'CommV', 'CommTrop', 
                    // ambrosiana
                    'Ingressa ', 'Psalmellus', 'PsalmellusV', 'Cantus', 'CantusV', 'V. in Alleluia', 
                    'Post evangelium', 'Offerenda', 'OfferendaV', 'Confractorium', 'Transitorium', 
                    // mozarabica
                    'Officium', 'OfficiumV', 'Psallendo', 'PsallendoV', 'Tractus', 'TractusV', 
                    'Lauda', 'LaudaV', 'Sacrificium', 'SacrificiumV', 'Ad confractionem panis', 
                    'Ad accedentes', 'Ad accedentesV', 'Communio', 'Benedicamen',
                ],
                'L': [
                    // romana
                    'Proph', 'ProphTrop', 'Lc', 'Ev', 'EvTrop',
                    // ambrosiana
                    'Lectio', 'Epistola', 'Evangelium',
                    // mozarabica
                    'Prophetia', 'Sapientia', 'Lectio', 'Epistola', 'Evangelium',
                    // gallicana
                    'Lectio', 'Epistola', 'Evangelium',
                ],
                'S/C': missSC,
                'R': ['Rub']
            },
            'OFF': {
                'G/A': [
                    'Inv', 'Ant', 'AntV', 'AntB', 'AntQ', 'AntM', 'AntN', 'R', 'RV', 'RTrop', 
                    'Rb', 'RbV', 'H', 'Ps', 'CantVT', 'CantNT', 'W', 'WSac', 'Benedicamen',
                    'Gr', 'GrV', 'All', 'AllV', 'Seq',
                ],
                'L': [
                    'Lec', 'Ser', 'Leg', 'Ev', 'Hom', 
                ],
                'S/C': offSC,
                'R': ['Rub']
            },
            'RIT': {
                'G/A': [],
                'L': [],
                'S/C': [
                    ...missSC, ...offSC, ...['F', 'Ex', 'Pf', 'Alloc', 'Lit', 'Abs', 'Ben', 'Or', 'VD']
                ],
                'R': ['Rub']
            }
        }

        possibleValues['RIT']['G/A'] = [...possibleValues['MISS']['G/A'], ...possibleValues['OFF']['G/A']]
        possibleValues['RIT']['L'] = [...possibleValues['MISS']['L'], ...possibleValues['OFF']['L']]
        
        return possibleValues
    }
    
    getIndexWithFieldName(fieldName) {
        switch (fieldName) {
            case 'id':
                return 0;           // ID
                
            case 'shelfmark':
                return 1;           // SHELFMARK
                
            case 'source':
                return 2;           // SOURCE
                
            case 'type':
                return 3;           // TYPE

            case 'part':
                return 4;           // PART
                
            case 'season_month':
                return 5;           // SEASON/MONTH
                
            case 'week':
                return 6;           // WEEK

            case 'day':
                return 7;           // DAY

            case 'feast':
                return 8;           // FEAST

            case 'commune_votive':
                return 9;           // COMMUNE/VOTIVE

            case 'topics':
                return 10;          // TOPICS

            case 'mass_hour':
                return 11;          // MASS/HOUR

            case 'ceremony':
                return 12;          // CEREMONY

            case 'module':
                return 13;          // MODULE

            case 'sequence':
                return 14;          // SEQUENCE

            case 'rubrics':
                return 15;          // RUBRICS

            case 'layer':
                return 16;          // LAYER

            case 'genre':
                return 17;          // GENRE

            case 'series':
                return 18;          // SERIES

            case 'item':
                return 19;          // ITEM

            case 'standard_item':
                return 20;          // STANDARD ITEM

            case 'digital_page_number':
                return 21;          // PAGE NUMBER (DIGITAL)

            case 'original_page_number':
                return 22;          // PAGE NUMBER (ORIGINAL)

            case 'page_link':
                return 23;          // PAGE NUMBER (ORIGINAL)

            case 'remark':
                return 24;          // REMARK

            case 'made_by':
                return 25;          // MADE BY
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
    
    getTopicInfoWithLabel(topicToFind) {
        for (let topic of this.topics) {
            if (topic.name === topicToFind) {
                return topic
            }
        }
        
        return null
    }
    
    validate() {
        this.loadSheetData()
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
        let errors = []
        
        let checkSubset = this.sheetWrapper.getActiveRowsCount() > 1
        let firstRowIndex = 0
        let lastRowIndex = 0
        
        if (checkSubset === true) {
            firstRowIndex = this.sheetWrapper.getFirstSelectedRow() - 2
            lastRowIndex = this.sheetWrapper.getLastSelectedRow() - 2
        }
        
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
    
    /**
     * VALIDATION
     *
     * TYPE >                       DONE
     * PART >                       DONE
     * SEASON/MONTH >               DONE
     * WEEK >                       DONE
     * DAY >                        DONE
     * FEAST >                      DEPRECATED
     * COMMUNE/VOTIVE >             DEPRECATED
     * TOPICS >                     
     * MASS/HOUR >                  DONE
     * CEREMONY >                   DONE
     * MODULE > 
     * SEQUENCE >                   DONE
     * RUBRICS >                    DONE
     * LAYER >
     * GENRE >                      DONE
     * SERIES >
     * ITEM >                       DONE
     * PAGE NUMBER (DIGITAL) >      DONE
     * PAGE NUMBER (ORIGINAL) > 
     * REMARK > 
     * MADE BY >                    DONE
     */
    validateRow(row) {
        let errors = []
        let result = null
        
        let remark = this.getRowDataWithName('remark', row)
        if (remark === 'del') {
            return errors // skip
        }
        
        result = this.validateDigitalPageNumber(row)
        if (result !== true) {
            result.cellIndex = this.getIndexWithFieldName('digital_page_number')
            errors.push(result)
        }

        result = this.validateRubricsItem(row)
        if (result !== true) {
            result.cellIndex = this.getIndexWithFieldName('item')
            errors.push(result)
        }
        
        result = this.validateGenre(row)
        if (result !== true) {
            result.cellIndex = this.getIndexWithFieldName('genre')
            errors.push(result)
        }

        result = this.validateType(row)
        if (result !== true) {
            result.cellIndex = this.getIndexWithFieldName('type')
            errors.push(result)
        }
        
        result = this.validateLayer(row)
        if (result !== true) {
            result.cellIndex = this.getIndexWithFieldName('layer')
            errors.push(result)
        }
        
        result = this.validateCeremony(row)
        if (result !== true) {
            result.cellIndex = this.getIndexWithFieldName('ceremony')
            errors.push(result)
        }
        
        result = this.validateMassHour(row)
        if (result !== true) {
            result.cellIndex = this.getIndexWithFieldName('mass_hour')
            errors.push(result)
        }

        result = this.validatePart(row)
        if (result !== true) {
            result.cellIndex = this.getIndexWithFieldName('part')
            errors.push(result)
        }

        result = this.validateSeasonMonth(row)
        if (result !== true) {
            result.cellIndex = this.getIndexWithFieldName('season_month')
            errors.push(result)
        }

        result = this.validateWeek(row)
        if (result !== true) {
            result.cellIndex = this.getIndexWithFieldName('week')
            errors.push(result)
        }

        result = this.validateDay(row)
        if (result !== true) {
            result.cellIndex = this.getIndexWithFieldName('day')
            errors.push(result)
        }
        
        result = this.validateTopics(row)
        if (result !== true) {
            result.cellIndex = this.getIndexWithFieldName('topics')
            errors.push(result)
        }
        
        result = this.validateMadeBy(row)
        if (result !== true) {
            result.cellIndex = this.getIndexWithFieldName('made_by')
            errors.push(result)
        }
        
        result = this.validateSequence(row)
        if (result !== true) {
            result.cellIndex = this.getIndexWithFieldName('sequence')
            errors.push(result)
        }
        
        return errors.length === 0 ? true : errors
    }
    
    // ok
    validateDigitalPageNumber(row) {
        let digitalPageNumber = this.getRowDataWithName('digital_page_number', row)

        if (digitalPageNumber.length === 0) {
            return new CellValidationError('Digital Page Number missing')
        }
        
        if (/^\d+$/.exec(digitalPageNumber) === null) {
            return new CellValidationError('Digital Page Number format is invalid')
        }
        
        return true
    }
    
    // ok
    validateRubricsItem(row) {
        let item = this.getRowDataWithName('item', row)
        let rubrics = this.getRowDataWithName('rubrics', row)
        
        if (item === '-') {
            item = ''
        }
        
        if (item.length === 0 && rubrics.length === 0) {
            return new CellValidationError('Either Rubric or Item is required')
        }
        
        return true
    }
    
    // ok
    validateGenre(row) {
        let genre = this.getRowDataWithName('genre', row)
        let item = this.getRowDataWithName('item', row)
        
        if (genre.length === 0) {
            return new CellValidationError('Genre is required')
        }
        
        if (item.length > 0 && genre === 'Rub') {
            return new CellValidationError('Invalid Genre')
        }
        
        let type = this.getRowDataWithName('type', row)
        
        if (type.length === 0) {
            if (genre === 'Rub') {
                type = 'MISS'
            }
            else {
                return new CellValidationError('Type is required for Genre')
            }
        }

        let layer = this.getRowDataWithName('layer', row)
        if (layer.length === 0) {
            return new CellValidationError('Layer is required for Genre')
        }
        
        let genres = this.getValidValues('genre', {
            type: type,
            layer: layer
        })
        
        if (genres === null) {
            return new CellValidationError(`Validation system error, genres list is empty.`)
        }
        
        if (genres.indexOf(genre) === -1) {
            return new CellValidationError(`Genre is invalid: ${genre}`)
        }
        
        return true
    }
    
    validateCeremony(row) {
        let type = this.getRowDataWithName('type', row)
        let ceremony = this.getRowDataWithName('ceremony', row)
        
        if (type === 'MISS' && ceremony !== 'Mass Propers') {
            return new CellValidationError('Ceremony at MISS must be Mass Propers')
        }

        if (type === 'OFF' && ceremony !== 'Office Propers') {
            return new CellValidationError('Ceremony at OFF must be Office Propers')
        }
        
        if (type === 'RIT' && ceremony.length === 0) {
            return new CellValidationError('Ceremony is required')
        }
        
        if (ceremony.length !== 0) {
            for (let label of this.indexLabels) {
                if (label.name === ceremony) {
                    return true
                }
            }
        
            return new CellValidationError(`Ceremony is invalid: ${ceremony}`)
        }
        
        return true
    }
    
    validateMassHour(row) {
        let type = this.getRowDataWithName('type', row)
        let massHour = this.getRowDataWithName('mass_hour', row)
        let genre = this.getRowDataWithName('genre', row)
        
        if (type === 'RIT') {
            if (massHour.length > 0) {
                return new CellValidationError('Mass/Hour must empty on RIT')
            }
            
            return true
        }
        
        if (genre === 'Rub') {
            return true
        }
        
        if (massHour.length === 0) {
            return new CellValidationError('Mass/Hour is required on MISS / OFF')
        }
        
        let valids = this.getValidValues('mass_hour', {type: type})
        
        if (valids === null || valids.indexOf(massHour) === -1) {
            return new CellValidationError('Mass/Hour value is invalid')
        }
        
        return true
    }
    
    validateType(row) {
        let type = this.getRowDataWithName('type', row)
        let genre = this.getRowDataWithName('genre', row)
        
        if (genre === 'Rub' && type.length === 0) {
            return true
        }
        
        if (type.length === 0) {
            return new CellValidationError('Type is required')
        }
        
        let valids = this.getValidValues('type')
        
        if (valids.indexOf(type) === -1) {
            return new CellValidationError(`Type is invalid: ${type}`)
        }
        
        return true
    }
    
    validatePart(row) {
        let type = this.getRowDataWithName('type', row)
        let part = this.getRowDataWithName('part', row)
        let genre = this.getRowDataWithName('genre', row)
        
        if (part.length === 0) {
            if (genre === 'Rub') {
                return true
            }
            
            if (type !== 'RIT') {
                return new CellValidationError('Part is required')
            }
            
            return true
        }
        
        let valids = this.getValidValues('part', {type: type})
        
        if (valids === null || valids.indexOf(part) === -1) {
            return new CellValidationError(`Part is invalid: ${part}`)
        }
        
        return true
    }
    
    validateSeasonMonth(row) {
        let type = this.getRowDataWithName('type', row)
        let seasonMonth = this.getRowDataWithName('season_month', row)
        let part = this.getRowDataWithName('part', row)
        let genre = this.getRowDataWithName('genre', row)

        if (part.length === 0) {
            if (genre === 'Rub') {
                return true
            }

            if (type !== 'RIT') {
                return new CellValidationError('Part is missing for Season/Month')
            }
            
            return true
        }
        
        let validSeasonMonth = this.getValidValues('season_month', {part: part})
        
        if (validSeasonMonth === null) {
            return new CellValidationError('Invalid Part value for Season/Month')
        }
        
        if (validSeasonMonth.indexOf(seasonMonth) === -1) {
            return new CellValidationError('Invalid Part value for Season/Month')
        }
        
        return true
    }
    
    validateWeek(row) {
        let type = this.getRowDataWithName('type', row)
        let week = this.getRowDataWithName('week', row)
        let part = this.getRowDataWithName('part', row)
        let genre = this.getRowDataWithName('genre', row)
        
        if (part.length === 0) {
            if (genre === 'Rub') {
                return true
            }

            if (type !== 'RIT') {
                return new CellValidationError('Part is missing for Week')
            }
            
            return true
        }
        
        let validWeek = this.getValidValues('week', {part: part})
        
        if (validWeek === null) {
            return new CellValidationError('Invalid Part value for Week')
        }
        
        if (validWeek.indexOf(week) === -1) {
            return new CellValidationError(`Invalid Week value: ${week}`)
        }
        
        return true
    }
    
    validateDay(row) {
        let type = this.getRowDataWithName('type', row)
        let day = this.getRowDataWithName('day', row)
        let part = this.getRowDataWithName('part', row)
        let genre = this.getRowDataWithName('genre', row)
        
        if (part.length === 0) {
            if (genre === 'Rub') {
                return true
            }

            if (type !== 'RIT') {
                return new CellValidationError('Part is missing for Day')
            }
            
            return true
        }
        
        let validDay = this.getValidValues('day', {part: part})
        
        if (validDay === null) {
            return new CellValidationError('Invalid Part value for Day')
        }
        
        if (validDay instanceof RegExp) {
            if (validDay.exec(day)) {
                return true
            }
            
            return new CellValidationError('Invalid Part value for Day')
        }
        else if (validDay.indexOf(day) === -1) {
            return new CellValidationError('Invalid Part value for Day')
        }
        
        return true
    }
    
    validateTopics(row) {
        if (this.topics.length === 0) {
            return new CellValidationError('Lift of Topics not loaded :(')
        }
        
        let topics = this.getRowDataWithName('topics', row)
        
        if (topics.length === 0) {
            return true
        }
        
        for (let topic of System3.iterableTopics(topics)) {
            let topicsInfo = this.getTopicInfoWithLabel(topic)
        
            if (topicsInfo === null) {
                return new CellValidationError(`Topics not found: ${topic}`)
            }
        }
        
        return true
    }
    
    static iterableTopics(topics) {
        return topics.split(',').map(e => e.trim())
    }
    
    validateModule(row) {
        
    }
    
    validateSequence(row) {
        let sequence = this.getRowDataWithName('sequence', row)
        
        if (sequence.length === 0) {
            return new CellValidationError('Sequence is missing')
        }

        let validSeq = this.getValidValues('sequence')
        
        if (validSeq.exec(sequence)) {
            return true
        }

        return new CellValidationError('Invalid Sequence')
    }
    
    validateRubrics(row) {
    }
    
    validateLayer(row) {
        let layer = this.getRowDataWithName('layer', row)
        
        if (layer.length === 0) {
            return new CellValidationError('Missing Layer')
        }
        
        let layerValues = this.getValidValues('layer')
        
        if (layerValues.indexOf(layer) === -1) {
            return new CellValidationError(`Invalid Layer: ${layer}`)
        }
        
        return true
    }
    
    validateSeries(row) {
    }
    
    validateMadeBy(row) {
        return true
        let made = this.getRowDataWithName('made_by', row)
        
        if (made.length === 0) {
            return new CellValidationError('Made by is missing')
        }
        
        return true
    }
    
    /**
     * TYPE >                       DONE
     * PART >                       DONE
     * SEASON/MONTH >               -
     * WEEK >                       -
     * DAY > 
     * FEAST > 
     * COMMUNE/VOTIVE > 
     * TOPICS > 
     * MASS/HOUR > 
     * CEREMONY >                   DONE
     * MODULE > 
     * SEQUENCE > 
     * RUBRICS > 
     * LAYER >                      DONE
     * GENRE >                      DONE
     * SERIES >                     DONE
     * ITEM > 
     * PAGE NUMBER (DIGITAL) > 
     * PAGE NUMBER (ORIGINAL) > 
     * REMARK > 
     * MADE BY > 
     */
    fillEmptyFields() {
        this.loadSheetData()
        
        let checkSubset = this.sheetWrapper.getActiveRowsCount() > 1
        let firstRowIndex = 0
        let lastRowIndex = 0
        
        if (checkSubset === true) {
            firstRowIndex = this.sheetWrapper.getFirstSelectedRow() - 2
            lastRowIndex = this.sheetWrapper.getLastSelectedRow() - 2
        }

        for (let index in this.data) {
            if (checkSubset && (index < firstRowIndex || index > lastRowIndex)) {
                continue
            }

            let row = this.data[index]
            
            row = this.fillEmptyFieldsInRow(row)
            
            if (row === false) {
                continue
            }
            
            this.data[index] = row
        }
        
        if (checkSubset === false) {
            this.fillSequence(this.data)
            this.fillSeries(this.data)
        }
        this.writeData()
    }
    
    fillEmptyFieldsInRow(row) {
        let remark = this.getRowDataWithName('remark', row)
        
        if (remark === 'del') {
            return false
        }
        
        row = this.autofixes(row)
        
        // genre
        row = this.fillGenre(row)
        if (row === false) {
            return false
        }
        
        // layer
        row = this.fillLayer(row)
        if (row === false) {
            return false
        }
        
        // type
        row = this.fillType(row)
        if (row === false) {
            return false
        }
        
        // ceremony
        row = this.fillCeremony(row)
        if (row === false) {
            return false
        }
        
        // part
        row = this.fillPart(row)
        if (row === false) {
            return false
        }
        
        // mass hour
        row = this.fillMassHour(row)
        if (row === false) {
            return false
        }
        
        // topics
        row = this.fillTopics(row)
        if (row === false) {
            return false
        }
        
        return row
    }
    
    fillGenre(row) {
        let genre = this.getRowDataWithName('genre', row)
        if (genre.length > 0) {
            return row
        }
        
        let rubrics = this.getRowDataWithName('rubrics', row)
        if (rubrics.length === 0) {
            return false
        }

        let item = this.getRowDataWithName('item', row)
        if (item.length === 0) {
            this.setRowDataWithName('genre', 'Rub', row)
        }
        
        return row
    }
    
    fillLayer(row) {
        let genre = this.getRowDataWithName('genre', row)
        let layer = null
        
        let allGenres = this.getAllGenres()
        typeGenres:
        for (let _type in allGenres) {
            layerGenres:
            for (let _layer in allGenres[_type]) {
                if (allGenres[_type][_layer].indexOf(genre) != -1) {
                    layer = _layer
                    
                    break typeGenres
                }
            }
        }
        
        if (layer === null) {
            return false
        }
        
        row[this.getIndexWithFieldName('layer')] = layer
        
        return row
    }
    
    fillType(row) {
        let ceremony = this.getRowDataWithName('ceremony', row)
        
        if (ceremony === 'Mass Propers') {
            this.setRowDataWithName('type', 'MISS', row)
        }
        else if (ceremony === 'Office Propers') {
            this.setRowDataWithName('type', 'OFF', row)
        }
        else if (ceremony.length > 0) {
            this.setRowDataWithName('type', 'RIT', row)
        }

        return row
    }
    
    fillCeremony(row) {
        let type = this.getRowDataWithName('type', row)

        if (type === 'MISS') {
            let genre = this.getRowDataWithName('genre', row)
        
            if (genre === 'F') {
                this.setRowDataWithName('ceremony', 'Mass ordinary', row)
                this.setRowDataWithName('type', 'RIT', row)
                this.setRowDataWithName('part', '', row)
                this.setRowDataWithName('season_month', '', row)
                this.setRowDataWithName('week', '', row)
                this.setRowDataWithName('day', '', row)
                this.setRowDataWithName('topics', '', row)
                this.setRowDataWithName('mass_hour', '', row)
                this.setRowDataWithName('layer', 'S/C', row)
                return row
            }
            
            row[this.getIndexWithFieldName('ceremony')] = 'Mass Propers'
        }
        else if (type === 'OFF') {
            row[this.getIndexWithFieldName('ceremony')] = 'Office Propers'
        }

        let ceremony = this.getRowDataWithName('ceremony', row)
        
        let autoFill = {
            'Agatha': {
                season_month: 'Feb',
                week: '',
                day: '5',
                topics: 'Agatha',
            },
            'Ash Wednesday': {
                season_month: 'LXX',
                week: 'H3',
                day: 'f4',
                topics: 'Cin',
            },
            'Assumption Day': {
                season_month: 'Aug',
                week: '',
                day: '15',
                topics: 'Assumptio BMV',
            },
            'Baptismal Vespers': {
                season_month: 'Pasc',
                week: 'H0',
                day: '',
                topics: 'ff',
            },
            'Blasius': {
                season_month: 'Feb',
                week: '',
                day: '3',
                topics: 'Blasius',
            },
            'Candlemas': {
                season_month: 'Feb',
                week: '',
                day: '2',
                topics: 'Purificatio BMV',
            },
            'Chrism Mass': {
                season_month: 'Qu',
                week: 'H6',
                day: 'f5',
                topics: '',
            },
            'Christmas Eve': {
                season_month: 'Dec',
                week: '',
                day: '25',
                topics: 'Nat',
            },
            'Corpus Christi': {
                season_month: 'Trin',
                week: 'H1',
                day: 'f5',
                topics: 'Corp',
            },
            'Easter Sunday': {
                season_month: 'Pasc',
                week: 'H0',
                day: 'D',
                topics: 'Pasc',
            },
            'Epiphany': {
                season_month: 'Jan',
                week: '',
                day: '6',
                topics: 'Ep',
            },
            'Good Friday': {
                season_month: 'Qu',
                week: 'H6',
                day: 'f6',
                topics: '',
            },
            'Holy Saturday': {
                season_month: 'Qu',
                week: 'H6',
                day: 'S',
                topics: '',
            },
            'Maundy Thursday': {
                season_month: 'Qu',
                week: 'H6',
                day: 'f5',
                topics: '',
            },
            'Palm Sunday': {
                season_month: 'Qu',
                week: 'H6',
                day: 'D',
                topics: '',
            },
            'Paschal prayers': {
                season_month: 'Pasc',
                week: 'H0',
                day: 'ff',
                topics: '',
            },
            'reconciliation of penitents': {
                season_month: 'Qu',
                week: 'H6',
                day: 'f5',
                topics: '',
            },
            'Rogation Days': {
                season_month: 'Pasc',
                week: 'H5',
                topics: 'Rogat',
            },
            'stripping/washing of altars': {
                season_month: 'Qu',
                week: 'H6',
                day: 'f5',
                topics: '',
            },
            'Tenebrae': {
                season_month: 'Qu',
                week: 'H6',
                day: 'f5',
                topics: '',
            },
            'Vigil of Whitsun': {
                season_month: 'Pasc',
                week: 'H6',
                day: 'S',
                topics: 'Vig.Pent',
            },
            'washing of feet': {
                season_month: 'Qu',
                week: 'H6',
                day: 'f5',
                topics: '',
            },
        }
        
        if (autoFill[ceremony] !== undefined) {
            let data = autoFill[ceremony]
            
            for (let key of Object.keys(data)) {
                let value = data[key]
                
                row[this.getIndexWithFieldName(key)] = value
            }
        }
        
        return row
    }
    
    calcualtePart(type, seasonMonth, week, day, topics, ceremony) {
        if (typeof day === 'number') {
            day = day.toString()
        }
        
        if (type === 'RIT') {
            if (ceremony === 'Mass ordinary'
                || ceremony === 'preparation and vesting for Mass'
                || ceremony === 'thanksgiving after Mass') {
                    return 'O'
                }
        }
        
        if (seasonMonth === 'Gen') {
            seasonMonth = ''
        }
        if (week === 'Gen') {
            week = ''
        }
        
        if (day === 'Gen') {
            day = ''
        }
        
        let numberOfFilledFields = 0
        numberOfFilledFields += week.length > 0 ? 1 : 0
        numberOfFilledFields += day.length > 0 ? 1 : 0
        numberOfFilledFields += seasonMonth.length > 0 ? 1 : 0

        let part = ''

        if (numberOfFilledFields === 1) {
            part = 'V'
        }
        
        // PS mukodesei tde betenni
        if (week.length > 0) {
            part = 'T'
        }
        else if (week.length === 0 && seasonMonth.length > 0 && day.length > 0) {
            let temporalValues = this.getValidValues('season_month', {part: 'T'})
            
            if (temporalValues.indexOf(seasonMonth) > -1) {
                part = 'T'
            }
            else {
                part = 'S'
            }
        }
        else if (week.length === 0 && seasonMonth.length === 0 && day.length === 0) {
            if (topics.length > 0) {
                
                for (let topic of System3.iterableTopics(topics)) {
                    let topicInfo = this.getTopicInfoWithLabel(topic)
        
                    if (topicInfo !== null) {
                        // tobb topic info alapjan
                        if (topicInfo.kind === 1 && part !== 'V') {
                            part = 'C'
                        }
                        else if (topicInfo.kind === 2) {
                            part = 'V'
                        }
                        else if (topicInfo.votive) {
                            part = 'V'
                        }
                    }
                }
            }
            else {
                part = ''
            }
        }
        
        if (type === 'OFF') {
            let validCount = 0
            
            let psValidSeason = this.getValidValues('season_month', {part: 'PS'})
            if (seasonMonth === '' || psValidSeason.indexOf(seasonMonth) > -1) {
                validCount++
            }
            
            let psValidWeek = this.getValidValues('week', {part: 'PS'})
            if (week === '' || psValidWeek.indexOf(week) > -1) {
                validCount++
            }

            let psValidDay = this.getValidValues('day', {part: 'PS'})
            if (day === '' || psValidDay.indexOf(day) > -1) {
                validCount++
            }
            
            if (topics.length === 0) {
                validCount++
            }
            
            if (validCount === 4) {
                part = 'PS'
            }
        }
        
        return part
    }
    
    fillPart(row) {
        let type = this.getRowDataWithName('type', row)
        let seasonMonth = this.getRowDataWithName('season_month', row)
        let week = this.getRowDataWithName('week', row)
        let day = this.getRowDataWithName('day', row)
        let topic = this.getRowDataWithName('topics', row)
        let ceremony = this.getRowDataWithName('ceremony', row)
        
        let part = this.calcualtePart(type, seasonMonth, week, day, topic, ceremony)
        
        if (part === false) {
            return false
        }
        
        this.setRowDataWithName('part', part, row)
        
        return row
    }
    
    fillMassHour(row) {
        let type = this.getRowDataWithName('type', row)
        let massHour = this.getRowDataWithName('mass_hour', row)
        
        if (type === 'MISS' && massHour.length === 0) {
            row[this.getIndexWithFieldName('mass_hour')] = 'M2'
        }
        
        return row
    }
    
    counterKeyWithNames(row, fieldNames) {
        let buffer = []
        for (let fieldName of fieldNames) {
            buffer.push(this.getRowDataWithName(fieldName, row))
        }
        
        return buffer.join('')
    }
    
    fillSequence(rows) {
        let counterKey = '';
        let counter = 1;
        for (let index in rows) {
            let row = rows[index]
            let type = this.getRowDataWithName('type', row)
            let fieldNames = type === 'RIT' ? ['shelfmark', 'type', 'ceremony'] : ['shelfmark', 'ceremony', 'season_month', 'week', 'day', 'mass_hour']
            let newCounterKey = this.counterKeyWithNames(row, fieldNames)
            
            if (counterKey !== newCounterKey) {
                counter = 1
                counterKey = newCounterKey
            }
            
            rows[index][this.getIndexWithFieldName('sequence')] = `${counter++}`
        }
        
        return rows
    }
    
    fillTopics(row) {
        let feast = this.getRowDataWithName('feast', row)
        let commune_votive = this.getRowDataWithName('commune_votive', row)
        let topics = this.getRowDataWithName('topics', row)
        
        let newValue = ''
        if (feast.length > 0) {
            newValue = feast
        }

        if (commune_votive.length > 0) {
            newValue = commune_votive
        }

        if (newValue.length > 0 && topics.length === 0) {
            this.setRowDataWithName('topics', newValue, row)
        }
        
        return row
    }
    
    doFillSeries(rows, genreInfo) {
        for (let genre of Object.keys(genreInfo)) {
            let info = genreInfo[genre]
            if (info.counter <= 1) {
                continue
            }
            
            let counter = 1
            
            for (let index of info.rows) {
                this.setRowDataWithName('series', counter.toString(), rows[index])
                counter++
            }
        }
    }
    
    fillSeries(rows) {
        let genreCounter = {}
        let prevCounterKey = null

        for (let index in rows) {
            let row = rows[index]
            
            this.setRowDataWithName('series', '', row)

            let genre = this.getRowDataWithName('genre', row)
            let fieldNames = ['shelfmark', 'season_month', 'week', 'day', 'topics', 'mass_hour', 'ceremony']
            let counterKey = this.counterKeyWithNames(row, fieldNames)
            
            if (genre.length === 0) {
                continue
            }
            
            if (counterKey !== prevCounterKey) {
                this.doFillSeries(rows, genreCounter)
                genreCounter = {}
                prevCounterKey = counterKey
            }
            
            if (genreCounter[genre] === undefined) {
                genreCounter[genre] = {
                    counter: 0,
                    rows: []
                }
            }
            
            genreCounter[genre].counter++
            genreCounter[genre].rows.push(index)
        }
        
        return rows
    }
    
    
    autofixes(row) {
        let genre = this.getRowDataWithName('genre', row)
        if (genre === 'HV') {
            this.setRowDataWithName('genre', 'H', row)
        }
        if (genre === 'LitV') {
            this.setRowDataWithName('genre', 'Lit', row)
        }
        if (genre === 'Sec') {
            this.setRowDataWithName('genre', 'Secr', row)
        }

        let week = this.getRowDataWithName('week', row)
        if (week === 'HH') {
            this.setRowDataWithName('genre', 'Gen', row)
        }

        let seasonMonth = this.getRowDataWithName('season_month', row)
        if (seasonMonth === 'Ian') {
            this.setRowDataWithName('season_month', 'Jan', row)
        }

        let feast = this.getRowDataWithName('feast', row)
        if (feast === 'Innoc.mm') {
            this.setRowDataWithName('feast', 'Innocentes', row)
        }
        if (feast === 'Thomas.m') {
            this.setRowDataWithName('feast', 'Thomas martyr', row)
        }

        this.autofixCaseSensitiveInTopics(row)
        this.autofixCaseSensitiveInCeremony(row)
        this.autofixFeast(row)
        this.autofixTopics(row)
        this.autofixTrim(row)
        
        return row
    }
    
    static newFeastName(feast) {
        let feasts = {
            'Vig.Nat': 'Nativitas (Vigilia)',
            'Nat': 'Nativitas',
            'Die2.Nat': 'Nativitas (Die 2)',
            'Die3.Nat': 'Nativitas (Die 3)',
            'Die4.Nat': 'Nativitas (Die 4)',
            'Die5.Nat': 'Nativitas (Die 5)',
            'Die6.Nat': 'Nativitas (Die 6)',
            'Die7.Nat': 'Nativitas (Die 7)',
            'Steph.m': 'Stephanus martyr',
            'Oct.Steph.m': 'Stephanus martyr (Octava)',
            'Steph.m (Adventus costae)': 'Stephanus martyr (Adventus costae)',
            'Steph.m (Adventus reliquiarum)': 'Stephanus martyr (Adventus reliquiarum)',
            'Steph.m (Ded altaris)': 'Stephanus martyr (Dedicatio altaris)',
            'Steph.m (Inventio Octava)': 'Stephanus martyr (Inventio Octava)',
            'Steph.m (Inventio)': 'Stephanus martyr (Inventio)',
            'Steph.m (Translatio Octava)': 'Stephanus martyr (Translatio Octava)',
            'Steph.m (Translatio)': 'Stephanus martyr (Translatio)',
            'Ioan.ap': 'Ioannes apostolus',
            'Ioan.ap (ante Portam Latinam Octava)': 'Ioannes apostolus (ante Portam Latinam Octava)',
            'Ioan.ap (ante Portam Latinam)': 'Ioannes apostolus (ante Portam Latinam)',
            'Ioan.ap (Commemoratio)': 'Ioannes apostolus (Commemoratio)',
            'Ioan.ap (Octava)': 'Ioannes apostolus (Octava)',
            'Ioan.ap (Vigilia)': 'Ioannes apostolus (Vigilia)',
            'Oct.Ioan.ap': 'Ioannes apostolus (Octava)',
            'Innoc.mm': 'Innocentes',
            'Oct.Innoc.mm': 'Innocentes (Octava)',
            'Circ': 'Circumcisio',
            'Circ (Vigilia)': 'Circumcisio (Vigilia)',
            'Vig.Ep': 'Epiphania (Vigilia)',
            'Ep': 'Epiphania',
            'Die2.Ep': 'Epiphania (Die 2)',
            'Die3.Ep': 'Epiphania (Die 3)',
            'Die4.Ep': 'Epiphania (Die 4)',
            'Die5.Ep': 'Epiphania (Die 5)',
            'Die6.Ep': 'Epiphania (Die 6)',
            'Die7.Ep': 'Epiphania (Die 7)',
            'Oct.Ep (Vigilia)': 'Epiphania (Octava Vigilia)',
            'Oct.Ep': 'Epiphania (Octava)',
            'Thomas.m': 'Thomas martyr',
            'Transl.Iacobus.ap': 'Iacobus maior (Translatio)',
            'Silvester.cf': 'Silvester papa',
            'Columba.m': 'Columba',
            'D70': 'Septuagesima',
            'D60': 'Sexagesima',
            'D50': 'Quinquagesima',
            'Cin': 'Cinerum',
            'Pasc': 'Resurrectio Domini',
            'Pasc.annotinum': 'Pascha annotinum',
            'D.prop': 'Dominica propria',
            'Rogat': 'Rogationes',
            'Vig.Asc': 'Ascensio Domini (Vigilia)',
            'Asc': 'Ascensio Domini',
            'Oct.Asc': 'Ascensio Domini (Octava)',
            'Vig.Pent': 'Pentecostes (Vigilia)',
            'Pent': 'Pentecostes',
            'Oct.Pent': 'Pentecostes (Octava)',
            'Trin': 'Trinitas',
            'Oct.Trin': 'Trinitas (Octava)',
            'Vig.Corp': 'Corpus Christi (Vigilia)',
            'Corp': 'Corpus Christi',
            'Die2.Corp': 'Corpus Christi (Die 2)',
            'Die3.Corp': 'Corpus Christi (Die 3)',
            'Die4.Corp': 'Corpus Christi (Die 4)',
            'Die5.Corp': 'Corpus Christi (Die 5)',
            'Die6.Corp': 'Corpus Christi (Die 6)',
            'Die7.Corp': 'Corpus Christi (Die 7)',
            'Oct.Corp': 'Corpus Christi (Octava)',
        }
        
        if (feasts[feast] !== undefined && feasts[feast].length > 0) {
            return feasts[feast]
        }
        
        return null
    }
    
    autofixFeast(row) {
        let feast = this.getRowDataWithName('feast', row)
        
        if (feast.length === 0) {
            return
        }
        
        let newFeast = System3.newFeastName(feast)
        if (newFeast !== null) {
            this.setRowDataWithName('feast', newFeast, row)
        }
    }
    
    autofixTopics(row) {
        let changedTopicsBuffer = []
        
        let topics = this.getRowDataWithName('topics', row)
        if (topics.length === 0) {
            return
        }
        
        for (let topic of System3.iterableTopics(topics)) {
            let newFeast = System3.newFeastName(topic)
        
            if (newFeast !== null) {
                changedTopicsBuffer.push(newFeast)
            }
            else {
                changedTopicsBuffer.push(topic)
            }
        }
        
        let newTopic = changedTopicsBuffer.join(', ')
        
        if (newTopic === topics) {
            return
        }
        
        this.setRowDataWithName('topics', newTopic, row)
    }
    
    autofixCaseSensitiveInTopics(row) {
        let topics = this.getRowDataWithName('topics', row)
        if (topics.length === 0) {
            return
        }

        let topicBuffer = []

        for (let topic of System3.iterableTopics(topics)) {
            let lowerTopic = topic.toLowerCase()
            
            let found = false
            for (let topicInfo of this.topics) {
                if (topicInfo.name.toLowerCase() === lowerTopic) {
                    topicBuffer.push(topicInfo.name)
                    found = true
                    break
                }
            }
            
            if (found === false) {
                topicBuffer.push(topic)
            }
        }
        
        let newTopic = topicBuffer.join(', ')
        
        if (topics !== newTopic) {
            this.setRowDataWithName('topics', newTopic, row)
        }
    }
    
    autofixCaseSensitiveInCeremony(row) {
        let ceremony = this.getRowDataWithName('ceremony', row)
        if (ceremony.length === 0) {
            return
        }

        let lowerCeremony = ceremony.toLowerCase()
        
        let foundCeremony = ''
        let found = false
        for (let indexLabel of this.indexLabels) {
            if (indexLabel.name.toLowerCase() === lowerCeremony) {
                foundCeremony = indexLabel.name
                found = true
                break
            }
        }
        
        if (found === false) {
            return
        }

        if (ceremony !== foundCeremony && foundCeremony.length > 0) {
            this.setRowDataWithName('ceremony', foundCeremony, row)
        }
    }
    
    autofixTrim(row) {
        for (let index in row) {
            let value = row[index]
            if (value && value.trim) {
                row[index] = value.trim()
            }
        }
    }
    
    fixMergedCells() {
        for (let range of this.sheetWrapper.getMergedCellRanges()) {
            range.breakApart()
        }
    }
}
