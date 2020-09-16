class System3Validator
{
  constructor(system3) {
    this.system3 = system3
  }
  
  validateFeastCommuneToTopics() {
    let invalids = []
    
    for (let i = 0; i < System3.data.length; i++) {
      let row = this.system3.data[i]
  
      // feast
      let feastCellPointer = {row: i+1, col: 7}
      let communeCellPointer = {row: i+1, col: 8}
      
      if (row[feastCellPointer.column].length > 0 && row[communeCellPointer.column].length) {
        invalids.push({
          row: feastCellPointer.row,
          col: feastCellPointer.col,
          message: 'You can\'t fill feast and commune column at the same time.'
        })
        invalids.push({
          row: communeCellPointer.row,
          col: communeCellPointer.col,
          message: 'You can\'t fill feast and commune column at the same time.'
        })
      }
    }
    
    return invalids
  }
  
  validateHeaders() {
    let expectedHeaders = this.system3.getHeadersTemplate()

    var invalids = []

    for (var i = 0; i < expectedHeaders.length; i++) {
      if (expectedHeaders[i] !== this.system3.headers[i]) {
        invalids.push({
          row: 1,
          col: i,
          message: `invalid header: "${this.system3.headers[i]}" expected: "${expectedHeaders[i]}"`
        })
      }
    }

    return invalids
  }
}

