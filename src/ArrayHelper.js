class ArrayHelper
{
    static insertItemAfter(arr, item, after) {
      let beforeItemPosition = arr.indexOf(after)
  
      if (beforeItemPosition === -1) {
        return false
      }
  
      arr.splice(beforeItemPosition + 1, 0, item)
  
      return beforeItemPosition
    }

    static hasItem(arr, needle) {
      return arr.indexOf(needle) !== -1
    }

    static hasFieldAfterField(arr, searchedField, afterField) {
      let afterFieldPosition = arr.indexOf(afterField)
  
      if (afterFieldPosition === -1) {
        return false
      }
  
      return arr.indexOf(searchedField) === afterFieldPosition+1
    }
    
}
