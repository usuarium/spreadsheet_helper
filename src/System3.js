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
    }
    
    async loadLists() {
        this.indexLabels = await UsuariumAPIClient.fetchIndexLabels()
    }
  
    getDocumentName() {
        return SpreadsheetApp.getActiveSpreadsheet().getName()
    }
  
    static getBookIdentifier(documentName) {
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
        
        Logger.log(changedDataCoordinates)
        
        if (changedDataCoordinates.length > 0) {
            this.markBackgroundsAt(changedDataCoordinates)
        }
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
  
    determinateSheetState() {
        this.hasShelfmarkColumn = ArrayHelper.hasItem(this.headers, 'ID') && ArrayHelper.hasFieldAfterField(this.headers, 'SHELFMARK', 'ID')
        this.hasSourceColumn = ArrayHelper.hasItem(this.headers, 'SHELFMARK') && ArrayHelper.hasFieldAfterField(this.headers, 'SOURCE', 'SHELFMARK')
        this.hasTopicsColumn = ArrayHelper.hasItem(this.headers, 'COMMUNE/VOTIVE') && ArrayHelper.hasFieldAfterField(this.headers, 'TOPICS', 'COMMUNE/VOTIVE')
        this.hasPageLinkColumn = ArrayHelper.hasItem(this.headers, 'PAGE NUMBER (ORIGINAL)') && ArrayHelper.hasFieldAfterField(this.headers, 'PAGE LINK', 'PAGE NUMBER (ORIGINAL)')
    }
  
    insertHeaderAfter(header, after) {
        let newHeaderIndex = ArrayHelper.insertItemAfter(this.headers, header, after)
      
        if (newHeaderIndex === false) {
            return false
        }
    
        this.sheetWrapper.insertColumnAtPosition(newHeaderIndex + 1)
        this.sheetWrapper.setHeaderValueAtPosition(newHeaderIndex + 2, header)
    }
  
    migrateSheet() {
        if (!this.hasShelfmarkColumn) {
            if (this.insertHeaderAfter('SHELFMARK', 'ID') !== false) {
                this.hasShelfmarkColumn = true
            }
        }

        if (!this.hasSourceColumn) {
            if (this.insertHeaderAfter('SOURCE', 'SHELFMARK') !== false) {
                this.hasSourceColumn = true
            }
        }

        if (!this.hasTopicsColumn) {
            if (this.insertHeaderAfter('TOPICS', 'COMMUNE/VOTIVE') !== false) {
                this.hasTopicsColumn = true
            }
      
            this.migrateTopics()
        }

        if (!this.hasPageLinkColumn) {
            if (this.insertHeaderAfter('PAGE LINK', 'PAGE NUMBER (ORIGINAL)') !== false) {
                this.hasTopicsColumn = true
            }
        }

        this.loadSheetData()
        this.migrateShelfmark()
        this.migrateGenre()
    }
  
    loadDocumentName() {
        this.documentName = this.sheetWrapper.getDocumentName()
    }
  
    migrateShelfmark() {
        let documentName = this.sheetWrapper.getDocumentName()
    
        if (documentName.indexOf(' - ') === -1) {
            throw new Error('Can\'t find the shelfmark')
        }
    
        let shelfmark = System3.getBookIdentifier(documentName)
    
        if (shelfmark <= 10000) {
            this.fillShelfmark(shelfmark)
            return
        }
    
        let bookId = UsuariumAPIClient.fetchBookId(shelfmark)
        documentName = documentName.replace(` - ${shelfmark}`, ` - ${bookId}`)
        this.sheetWrapper.setDocumentName(documentName)
        this.fillShelfmark(bookId)
    }
  
    fillShelfmark(shelfmark) {
        if (!this.hasShelfmarkColumn) {
            return
        }
    
        let sheet = SpreadsheetApp.getActiveSheet();
        let firstPageLinkCellValue = sheet.getRange(2, 2).getValue()
        if (firstPageLinkCellValue && firstPageLinkCellValue.length > 0) {
            return
        }
    
        this.sheetWrapper.setValueAt(shelfmark, {row: 2, col: 2, rows: this.sheetWrapper.getLastRow() - 1, cols: 1})
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
            
            if (!genre.replace) {
                Logger.log({genre: genre})
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
  
    loadFeastCommuneTopicData() {
        let indexOfFeast = this.headers.indexOf('FEAST')
    
        let sheet = SpreadsheetApp.getActiveSheet();
        let dataRange = sheet.getDataRange()
        return sheet.getRange(2, indexOfFeast + 1, dataRange.getLastRow()-1, 3).getValues()
    }
  
    setTopicData(topic) {
        let indexOfTopics = this.headers.indexOf('TOPICS')
    
        let sheet = SpreadsheetApp.getActiveSheet();
        let dataRange = sheet.getDataRange()
    
        return sheet.getRange(2, indexOfTopics + 1, dataRange.getLastRow()-1, 1).setValues(topic)
    }
  
    migrateTopics() {
        let data = this.loadFeastCommuneTopicData()
        let topics = data.map(row => {
            if (row[0]) {
                return [row[0]];
            }
            else if (row[1]) {
                return [row[1]]
            }
      
            return ['']
        })
    
        this.setTopicData(topics)
    }
    
    getValidValues(fieldName, dependentValues) {
        if (fieldName === 'type') {
            return ['MISS', 'OFF', 'RIT']
        }
        else if (fieldName === 'part') {
            let missAndOffValues = ['T', 'S', 'C', 'V']
            let ritValue = ['O']
            
            let value = {
                'RIT': ritValue,
                'MISS': missAndOffValues,
                'OFF': missAndOffValues,
            }[dependentValues.type]
            
            return value === undefined ? null : value
        }
        else if (fieldName === 'season_month') {
            let temporalAndSanctoralValues = ['Adv', 'Nat', 'Ep', 'LXX', 'Qu', 'Pasc', 'Pent', 'Trin', 'Gen']
            
            let value = {
                'T': temporalAndSanctoralValues,
                'S': temporalAndSanctoralValues,
                'C': '',
                'V': ''
            }[dependentValues.part]

            return value === undefined ? null : value
        }
        else if (fieldName === 'week') {
            let temporalValues = ['H0', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'H7', 'H8', 'H9', 'H10', 'H11', 'H12', 'H13', 'H14', 'H15', 'H16', 'H17', 'H18', 'H19', 'H20', 'H21', 'H22', 'H23', 'H24', 'H25', 'H26', 'H27', 'QuT', 'Gen']
            
            let value = {
                'T': temporalValues,
                'S': '',
                'C': '',
                'V': ''
            }[dependentValues.part]

            return value === undefined ? null : value
        }
        else if (fieldName === 'day') {
            let value = {
                'T': new RegExp(/^([1-9]|[12]\d|3[01])$/),
                'S': ['D', 'ff', 'f2', 'f3', 'f4', 'f5', 'f6', 'S', 'Gen', 'F'],
                'C': '',
                'V': ''
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
                    'Ad accedentes', 'Ad accedentesV', 'Communio'
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
                'S/C': [
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
            },
            'OFF': {
                'G/A': [
                    'Inv', 'Ant', 'AntV', 'AntB', 'AntQ', 'AntM', 'AntN', 'R', 'RV', 'RTrop', 
                    'Rb', 'RbV', 'H', 'Ps', 'CantVT', 'CantNT', 'W', 'WSac',
                ],
                'L': [
                    'Lec', 'Ser', 'Leg', 'Ev', 'Hom',
                ],
                'S/C': [
                    'Cap', 'Or',
                ]
            },
            'RIT': {
                'G/A': [],
                'L': [],
                'S/C': [
                    'F', 'Ex', 'Pf', 'Alloc', 'Lit', 'Abs', 'Ben',
                ]
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

            case 'ribrics':
                return 15;          // RUBRICS

            case 'layer':
                return 16;          // LAYER

            case 'genre':
                return 17;          // GENRE

            case 'series':
                return 18;          // SERIES

            case 'item':
                return 19;          // ITEM

            case 'digital_page_number':
                return 20;          // PAGE NUMBER (DIGITAL)

            case 'original_page_number':
                return 21;          // PAGE NUMBER (ORIGINAL)

            case 'remark':
                return 23;          // REMARK

            case 'made_by':
                return 24;          // MADE BY
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
    
    getTopicInfoWithLabel(topicToFind) {
        for (let topic of this.topics) {
            if (topic.name === topicToFind) {
                return topic
            }
        }
        
        return null
    }
    
    validateRow(row) {
        let errors = []
        let result = null
        
        result = this.validateDigitalPageNumber(row)
        if (result !== true) {
            errors.push(result)
        }

        result = this.validateRubricsItem(row)
        if (result !== true) {
            errors.push(result)
        }
        
        result = this.validateGenre(row)
        if (result !== true) {
            errors.push(result)
        }

        result = this.validateType(row)
        if (result !== true) {
            errors.push(result)
        }
        
        // Genre, Module, Ceremony, Mass/hour
        
        return errors.length === 0 ? true : errors
    }
    
    validateDigitalPageNumber(row) {
        let digitalPageNumber = this.getRowDataWithName('digital_page_number', row)

        if (digitalPageNumber.length === 0) {
            return new Error('Digital Page Number missing')
        }
        
        if (digitalPageNumber.match(/^\d+$/) === null) {
            return new Error('Digital Page Number format is invalid')
        }
        
        return true
    }
    
    validateRubricsItem(row) {
        let item = this.getRowDataWithName('item', row)
        let rubrics = this.getRowDataWithName('ribrics', row)
        
        if (item.length === 0 && rubrics.length === 0) {
            return new Error('Either Rubric or Item is required')
        }
        
        return true
    }
    
    validateGenre(row) {
        let genre = this.getRowDataWithName('genre', row)
        
        if (genre.length === 0) {
            return new Error('Genre is required')
        }
        
        let allGenres = this.getAllGenres()
        
        if (allGenres['MISS']['G/A'].indexOf(genre) !== -1) {
            return true
        }

        if (allGenres['MISS']['L'].indexOf(genre) !== -1) {
            return true
        }

        if (allGenres['MISS']['S/C'].indexOf(genre) !== -1) {
            return true
        }
        
        if (allGenres['OFF']['G/A'].indexOf(genre) !== -1) {
            return true
        }

        if (allGenres['OFF']['L'].indexOf(genre) !== -1) {
            return true
        }

        if (allGenres['OFF']['S/C'].indexOf(genre) !== -1) {
            return true
        }
        
        if (allGenres['RIT']['G/A'].indexOf(genre) !== -1) {
            return true
        }

        if (allGenres['RIT']['L'].indexOf(genre) !== -1) {
            return true
        }

        if (allGenres['RIT']['S/C'].indexOf(genre) !== -1) {
            return true
        }
        
        return new Error('Genre is invalid')
    }
    
    validateCeremony(row) {
        let type = this.getRowDataWithName('type', row)
        let ceremony = this.getRowDataWithName('ceremony', row)
        
        if (type === 'MISS' && ceremony !== 'Mass Propers') {
            return new Error('Ceremony at MISS must be Mass Propers')
        }

        if (type === 'OFF' && ceremony !== 'Office Propers') {
            return new Error('Ceremony at MISS must be Mass Propers')
        }
        
        if (type === 'RIT' && ceremony.length === 0) {
            return new Error('Ceremony is required')
        }
        
        // validate from list
        for (let label of this.indexLabels) {
            if (label.name === ceremony) {
                return true
            }
        }
        
        return new Error('Ceremony is invalid')
    }
    
    validateMassHour(row) {
        let type = this.getRowDataWithName('type', row)
        let massHour = this.getRowDataWithName('mass_hour', row)
        
        if (type === 'RIT' && massHour.length > 0) {
            return new Error('Mass/Hour must empty on RIT')
        }
        
        if (massHour.length === 0) {
            return new Error('Mass/Hour is required on MISS / OFF')
        }
        
        let valids = this.getValidValues('mass_hour', {type: type})
        
        if (valids.indexOf(massHour) === -1) {
            return new Error('Mass/Hour value is invalid')
        }
        
        return true
    }
    
    validateType(row) {
        let type = this.getRowDataWithName('type', row)
        
        if (type.length === 0) {
            return new Error('Type is required')
        }
        
        let valids = this.getValidValues('type')
        
        if (valids.indexOf(type) === -1) {
            return new Error('This type is invalid')
        }
        
        return true
    }
    
    validatePart(row) {
        let type = this.getRowDataWithName('type', row)
        let part = this.getRowDataWithName('part', row)
        
        if (part.length === 0) {
            return new Error('Part is required')
        }
        
        let valids = this.getValidValues('part', {type: type})
        
        if (valids.indexOf(type) === -1) {
            return new Error('This type is invalid')
        }
        
        return true
    }
    
    validateSeasonMonth(row) {
        
    }
    
    validateWeek(row) {
        
    }
    
    validateDay(row) {
        
    }
    
    validateTopics(row) {
        
    }
    
    validateModule(row) {
        
    }
    
    validateSequence(row) {
        
    }
    
    validateRubrics(row) {
    }
    
    validateLayer(row) {
    }
    
    validateSeries(row) {
    }
    
    validateMadeBy(row) {
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

        for (let index in this.data) {
            let row = this.data[index]
            
            row = this.fillEmptyFieldsInRow(row)
            
            if (row === false) {
                continue
            }
            
            this.data[index] = row
        }
        
        this.fillSequence(this.data)
        this.fillSeries(this.data)
        this.writeData()
    }
    
    fillEmptyFieldsInRow(row) {
        // genre
        let genre = this.getRowDataWithName('genre', row)
        
        if (genre.length === 0) {
            return false
        }
        
        // layer
        row = this.fillLayer(row)
        
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
        
        row = this.fillMassHour(row)
        
        if (row === false) {
            return false
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
    
    fillCeremony(row) {
        let type = this.getRowDataWithName('type', row)

        if (type === 'MISS') {
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
                day: 'f2',
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
    
    fillPart(row) {
        let week = this.getRowDataWithName('week', row)
        let seasonMonth = this.getRowDataWithName('season_month', row)
        let day = this.getRowDataWithName('day', row)
        let topic = this.getRowDataWithName('topics', row)
        
        if (week.length > 0) {
            row[this.getIndexWithFieldName('part')] = 'T'
        }
        if (week.length === 0 && seasonMonth.length > 0 && day.length > 0) {
            row[this.getIndexWithFieldName('part')] = 'S'
        }
        if (week.length === 0 && seasonMonth.length === 0 && day.length === 0) {
            if (topic.length === 0) {
                return false
            }
            
            let topicInfo = this.getTopicInfoWithLabel(topic)
            
            if (topicInfo === null) {
                return false
            }
            
            if (topicInfo.kind === 1) {
                row[this.getIndexWithFieldName('part')] = 'C'
            }
            else if (topicInfo.votive) {
                row[this.getIndexWithFieldName('part')] = 'V'
            }
        }
        
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
            let fieldNames = type === 'RIT' ? ['type', 'ceremony'] : ['ceremony', 'season_month', 'week', 'day', 'mass_hour']
            let newCounterKey = this.counterKeyWithNames(row, fieldNames)
            
            if (counterKey !== newCounterKey) {
                counter = 1
                counterKey = newCounterKey
            }
            
            rows[index][this.getIndexWithFieldName('sequence')] = `${counter++}`
        }
        
        return rows
    }
    
    fillSeries(rows) {
        // let grouppedData = {}
        //
        // for (let index in rows) {
        //     let row = rows[index]
        //     let genre = this.getRowDataWithName('genre', row)
        //
        //     let fieldNames = ['type', 'part', 'season_month', 'week', 'day', 'topics', 'ceremony']
        //     let counterKey = this.counterKeyWithNames(row, fieldNames)
        // }
        //
        // season+week+day+feast+ceremony
    }
}
