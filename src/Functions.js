/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

var _bookId = undefined

/**
 * Generates link for the digital page
 *
 * @param {digitalPageNumber} input The digital page number
 * @return The link for the page
 * @customfunction
 */
function PAGELINK(digitalPageNumber, bookId)
{
  if (bookId === undefined && _bookId === undefined) {
    let documentName = SpreadsheetApp.getActiveSpreadsheet().getName()
    _bookId = System3.getBookIdentifier(documentName)
  }
  else {
    _bookId = bookId
  }
  return `https://usuarium.elte.hu/book/${_bookId}/pagelink?digital_page_number=${digitalPageNumber}`
}

function goToError(row, col) {
  col *= 1
  row *= 1
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(row, col).activateAsCurrentCell();
}
