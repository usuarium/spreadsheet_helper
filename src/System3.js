class System3
{
  constructor() {
    this.values = {
      types: ['MISS', 'OFF', 'RIT']
    }
    
    this.headers = []
    this.data = []
    this.hasPageLinkColumn = false
    this.hasTopicsColumn = false
    
    this.validator = new System3Validator(this)
    
    this.loadSheetHeaders()
    this.loadSheetData()
    
    this.documentName = SpreadsheetApp.getActiveSpreadsheet().getName()
    
    this.determinateSheetState()
  }
  
  static getBookIdentifier(documentName) {
    return documentName.split(' - ')[1] * 1
  }
  
  loadSheetHeaders() {
    this.headers = this.getSheetHeaders()
  }
  
  loadSheetData() {
    this.data = this.getSheetData()
  }
  
  getSheetHeaders() {
    let sheet = SpreadsheetApp.getActiveSheet();
    let range = sheet.getRange('A1:V1');
    return range.getValues()[0];
  }
  
  getSheetData() {
    let sheet = SpreadsheetApp.getActiveSheet();
    let dataRange = sheet.getDataRange()
    return sheet.getRange(2, 1, dataRange.getLastRow()-1, 22).getValues();
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

    if (this.hasTopicsColumn) {
      headers = headers.splice(8, 0, 'TOPICS');
    }
    
    if (this.hasPageLinkColumn) {
      let originalPageNumberPosition = headers.indexOf('PAGE NUMBER (ORIGINAL)')
      headers = headers.splice(originalPageNumberPosition+1, 0, 'PAGE LINK')
    }
      
    return headers;
  }
  
  determinateSheetState() {
    this.hasTopicsColumn = this.headers[8] === 'TOPICS'
    
    let pageLinkPosition = this.headers.indexOf('PAGE LINK')
    
    if (pageLinkPosition !== -1 && this.headers.indexOf('PAGE NUMBER (ORIGINAL)') === pageLinkPosition-1) {
      this.hasPageLinkColumn = true;
    }
    else {
      this.hasPageLinkColumn = false;
    }
  }
  
  addTopicsColumn() {
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.insertColumnAfter(8)
    sheet.getRange('I1').setValue("TOPICS")
  }
  
  addPageLinkColumn() {
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.insertColumnAfter(20)
    sheet.getRange('U1').setValue("PAGE LINK")
  }
  
  migrateSheet() {
    if (!this.hasTopicsColumn) {
      this.addTopicsColumn()
    }
    
    if (!this.hasPageLinkColumn) {
      this.addPageLinkColumn()
    }

    this.loadSheetData();
    this.determinateSheetState();
  }
  
  migrateShelfmark() {
    if (this.documentName.indexOf(' - ') === -1) {
      throw new Error('Can\'t find the shelfmark')
    }
    
    let shelfmark = System3.getBookIdentifier(this.documentName)
    
    if (shelfmark <= 10000) {
      return
    }
    
    let bookId = UsuariumAPIClient.fetchBookId(shelfmark)
    this.documentName = this.documentName.replace(` - ${shelfmark}`, ` - ${bookId}`)
    
    SpreadsheetApp.getActiveSpreadsheet().rename(this.documentName)
  }
  
  fillPageLinks() {
    let sheet = SpreadsheetApp.getActiveSheet();
    let dataRange = sheet.getDataRange()
    return sheet.getRange(2, 21, /*dataRange.getLastRow()-1*/ 5, 1).setFormulaR1C1('=PAGELINK(R[0]C[-2])');
  }
  
}

