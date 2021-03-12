const add = (sourceWorkBook, sheetNames, destinationWorkBook) => {
    const sheetNamesList = sheetNames.split(',')
    const { Sheets: sourceSheets } = sourceWorkBook
    const { Sheets: destinationSheets, SheetNames: destinationSheetNames } = destinationWorkBook
    sheetNamesList.forEach((sheetName) => {
        for(let val=0;;val++){
            const tail = val? `-${val}` : ''
            let name = sheetName + tail
            if(destinationSheetNames.includes(name)) {
                continue
            } else {
                Object.defineProperty(destinationSheets, name,
                    Object.getOwnPropertyDescriptor(sourceSheets, sheetName))
                destinationSheetNames.push(name)
                break
            }
        }
    })
    return destinationWorkBook
}

module.exports = add
