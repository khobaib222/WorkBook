const getWorkBook = require('./get-book')

const extractNames = (config) => {
    const names = config.split(',')
    return names.map((name) => name.trim().split('|'))
}
const rename = (workBookName, config) => {
    const workBook = getWorkBook(workBookName)
    const names = extractNames(config)
    names.forEach(([oldName, newName]) => {
        const oldSheetName = oldName.trim()
        const newSheetName = newName.trim()
        const {SheetNames, Sheets} = workBook
        SheetNames.splice(SheetNames.indexOf(oldSheetName), 1, newSheetName)
        Object.defineProperty(Sheets, newSheetName,
            Object.getOwnPropertyDescriptor(Sheets, oldSheetName));
        delete Sheets[oldSheetName];
    })
    return workBook
}

module.exports = rename