class SheetWrapper
{
    getLastRow() {
        let sheet = SpreadsheetApp.getActiveSheet();
        let dataRange = sheet.getDataRange()
        return dataRange.getLastRow()
    }
    
    setValuesAt(data, coords) {
        this.getRangeAt(coords).setValues(data)
    }
    
    setValueAt(data, coords) {
        this.getRangeAt(coords).setValue(data)
    }
    
    setBackgroundColorAt(color, coords) {
        this.getRangeAt(coords).setBackground(color)
    }
    
    getRangeAt(coords) {
        let sheet = SpreadsheetApp.getActiveSheet();
        return sheet.getRange(coords.row, coords.col, coords.rows, coords.cols)
    }
    
    getHeaders() {
        let sheet = SpreadsheetApp.getActiveSheet();
        let dataRange = sheet.getDataRange()
        let range = sheet.getRange(1, 1, 1, dataRange.getLastColumn());
        return range.getValues()[0];
    }
    
    getDataRows() {
        let sheet = SpreadsheetApp.getActiveSheet();
        let dataRange = sheet.getDataRange()
        return sheet.getRange(2, 1, dataRange.getLastRow()-1, 22).getValues();
    }
    
    insertColumnAtPosition(position) {
        SpreadsheetApp.getActiveSheet().insertColumnAfter(position)
    }
  
    setHeaderValueAtPosition(position, header) {
        SpreadsheetApp.getActiveSheet().getRange(1, position).setValue(header)
    }
    
    getDocumentName() {
        return SpreadsheetApp.getActiveSpreadsheet().getName()
    }
    
    setDocumentName(documentName) {
        SpreadsheetApp.getActiveSpreadsheet().rename(documentName)
    }
    
    getValuesAtColumn(column) {
        let sheet = SpreadsheetApp.getActiveSheet();
        let dataRange = sheet.getDataRange()

        return sheet.getRange(2, column, dataRange.getLastRow()-1, 1).getValues();
    }
}
