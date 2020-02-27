function IndexValidatorStartValidation() {
  const index = new Index()
  
  index.loadSheetData()
  index.loadIndexLabelData()
  let errors = index.validate()
  
  validatorResult.showValidatorSidebar('Index Validation Result', errors)
}

class Index
{
  constructor() {
    this.indexLabelCache = {}
    this.bookIdCache = {}
    this.sheetData = []
    this.indexLabels = []
  }
  
  loadSheetData() {
    let sheet = SpreadsheetApp.getActiveSheet();
    
    var dataRange = sheet.getDataRange();
    this.sheetData = sheet.getRange(2, 1, dataRange.getLastRow()-1, 12).getValues();
  }
  
  loadIndexLabelData() {
    let contentText = UrlFetchApp.fetch("https://usuarium.elte.hu/api/index_labels").getContentText()
    this.indexLabels = JSON.parse(contentText)
  }
  
  validate() {
    let errors = []
    for (let index = 0; index < this.sheetData.length; ++index) {
      let row = this.sheetData[index]
      let rowCord = index + 2 // header + JS array index
      
      if (row[0].length == 0) {
        errors.push({
          row: rowCord,
          col: 1,
          message: `Missing book id`
        })
      }
      else if (!this.isValidBookId(row[0])) {
        errors.push({
          row: rowCord,
          col: 1,
          message: `Unknown book id: row[0]`
        })
      }
      
      if (row[2].length == 0) {
        errors.push({
          row: rowCord,
          col: 3,
          message: `Missing digital page number`
        })
      }
      
      if (row[5].length > 0) {
        let labels = row[5].split(',')
        labels = labels.map((label) => {
          return label.trim()
        })
        
        for (let label of labels) {
          if (!this.isValidLabel(label)) {
            errors.push({
              row: rowCord,
              col: 6,
              message: `Unknown index label: ${label}`
            })
          }
        }
      }
    }
    return errors
  }

  isValidBookId(bookId) {
    return true
    if (this.bookIdCache[bookId] !== undefined) {
      return this.bookIdCache[bookId]
    }
    
    let response = UrlFetchApp.fetch(`https://usuarium.elte.hu/api/book/${bookId}`)
    
    let valid = true
    if (response.getResponseCode() === 404) {
      valid = false
    }
    
    return this.bookIdCache[bookId] = valid
  }

  isValidLabel(labelName) {
    if (this.indexLabelCache[labelName] !== undefined) {
      return this.indexLabelCache[labelName]
    }
    
    let valid = false
    for (let label of this.indexLabels) {
      if (label.name == labelName) {
        valid = true
        break
      }
    }
    
    return this.indexLabelCache[labelName] = valid
  }
}
