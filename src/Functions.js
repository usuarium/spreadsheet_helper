/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

var bookId = undefined

/**
 * Generates link for the digital page
 *
 * @param {digitalPageNumber} input The digital page number
 * @return The link for the page
 * @customfunction
 */
function PAGELINK(digitalPageNumber)
{
  if (bookId === undefined) {
    let documentName = SpreadsheetApp.getActiveSpreadsheet().getName()
    bookId = System3.getBookIdentifier(documentName)
  }
  return `https://usuarium.elte.hu/book/${bookId}/pagelink?digital_page_number=${digitalPageNumber}`
}

