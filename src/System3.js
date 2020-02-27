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
    
    this.determinateSheetState()
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
  
  migrateSheet() {
    if (this.hasTopicsColumn && this.hasPageLinkColumn) {
      return
    }

    var result = this.validator.validateFeastCommuneToTopics()
    if (result.length > 0) {
      
    }
    
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.insertColumnAfter(8)
    sheet.getRange('I1').setValue("TOPICS")
    
    this.loadSheetData();
    this.determinateSheetState();
  }
  
  determinateSheetState() {
    this.hasTopicsColumn = this.headers[8] === 'TOPICS' ? 2 : 1
    
    let pageLinkPosition = this.headers.indexOf('PAGE LINK')
    
    if (pageLinkPosition !== -1 && this.headers.indexOf('PAGE NUMBER (ORIGINAL)') === pageLinkPosition+1) {
      this.hasPageLinkColumn = true;
    }
    else {
      this.hasPageLinkColumn = false;
    }
  }
}
