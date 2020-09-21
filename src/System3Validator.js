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
        let invalids = []
        
        for (let headerIndex in this.system3.headers) {
            let header = this.system3.headers[headerIndex]
            
            headerIndex *= 1
            
            if (header === null || header === undefined || header.length === 0) {
                invalids.push({
                    row: 1,
                    col: headerIndex+1,
                    message: 'Header cell can not empty'
                })
            }
            else if (this.system3.knownHeaders.indexOf(header) === -1) {
                invalids.push({
                    row: 1,
                    col: headerIndex+1,
                    message: `Unknown header: "${header}"`
                })
            }
        }
        
        return invalids
    }
}

