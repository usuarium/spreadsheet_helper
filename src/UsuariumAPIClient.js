class UsuariumAPIClient
{
  static fetchIndexLabels() {
    let contentText = UrlFetchApp.fetch("https://usuarium.elte.hu/api/index_labels").getContentText()
    return JSON.parse(contentText)
  }
  
  static fetchBookData(bookId) {
    
  }
  
  static fetchBookId(shelfmark) {
    let contentText = UrlFetchApp.fetch(`https://usuarium.elte.hu/api/shelfmark_convert/${shelfmark}`).getContentText()
    return JSON.parse(contentText)
  }
}
