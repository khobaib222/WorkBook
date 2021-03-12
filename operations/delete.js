const getWorkBook = require('./get-book')

const deleteSheets = (workBookName, sheetNames) => {
    const sheetNameList = sheetNames.split(',')
    const workBook = getWorkBook(workBookName)
    const { Sheets, SheetNames } = workBook
    sheetNameList.forEach(sheetName => {
        delete Sheets[sheetName]
        SheetNames.splice(SheetNames.indexOf(sheetName), 1)
    })
    return workBook
}

module.exports = deleteSheets