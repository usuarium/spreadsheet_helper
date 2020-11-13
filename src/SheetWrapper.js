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
    
    getDataRows(numberofcols) {
        numberofcols = numberofcols === undefined ? 24 : numberofcols
        
        let sheet = SpreadsheetApp.getActiveSheet();
        let dataRange = sheet.getDataRange()
        return sheet.getRange(2, 1, dataRange.getLastRow()-1, numberofcols).getValues();
    }
    
    insertColumnAfter(position) {
        if (position == 0) {
            SpreadsheetApp.getActiveSheet().insertColumnBefore(1)
        }
        else {
            SpreadsheetApp.getActiveSheet().insertColumnAfter(position)
        }
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
    
    getActiveRowsCount() {
        let sheet = SpreadsheetApp.getActiveSheet();
        let activeRange = sheet.getActiveRange()
        
        if (activeRange === null) {
            return 0
        }
        
        return activeRange.getNumRows()
    }
    
    getFirstSelectedRow() {
        let sheet = SpreadsheetApp.getActiveSheet();
        let activeRange = sheet.getActiveRange()

        if (activeRange === null) {
            return 0
        }
        
        return activeRange.getLastRow() - (activeRange.getNumRows() - 1)
    }

    getLastSelectedRow() {
        let sheet = SpreadsheetApp.getActiveSheet();
        let activeRange = sheet.getActiveRange()

        if (activeRange === null) {
            return 0
        }
        
        return activeRange.getLastRow()
    }
    
    highlightRowsAt(rows) {
        let sheet = SpreadsheetApp.getActiveSheet()
        let numberOfColumns = sheet.getDataRange().getLastColumn()
        for (let row of rows) {
            sheet.getRange(row, 1, 1, numberOfColumns).setBackground('#e6b8af')
        }
    }
    
    getMergedCellRanges() {
        return SpreadsheetApp.getActiveSheet().getDataRange().getMergedRanges()
    }
}
