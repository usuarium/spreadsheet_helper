class CellValidationError extends Error
{
    constructor(message) {
        super(message)
        
        this.row = null
        this.col = null
    }
    
    set rowIndex(index) {
        this.row = index + 1
    }
    
    set cellIndex(index) {
        this.col = index + 1
    }
    
    get errorObject() {
        return {
            message: this.message,
            row: this.row,
            col: this.col,
        }
    }
}

