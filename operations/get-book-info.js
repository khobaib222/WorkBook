const getWorkBook = require('./get-book')

const getWorkBookInfo = (workBookName) => {
    const { SheetNames, Sheets } = getWorkBook(workBookName)
    const filledSheetNames = SheetNames.filter((sheetName) => !!Sheets[sheetName]['!ref'])
    const emptySheetNames = SheetNames.filter((sheetName) => !Sheets[sheetName]['!ref'])
    return {
        filledSheetNames,
        emptySheetNames,
    }
}

module.exports = getWorkBookInfo