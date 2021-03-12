const xlsx = require('xlsx')
const readAllBooks = require('./operations/read-all-books')
const readBook = require('./operations/read-book')
const rename = require('./operations/rename')
const add = require('./operations/add')
const deleteSheets = require('./operations/delete')
const aggregateWorkBooks = require('./operations/aggregate-books')
const aggregateWorkSheets = require('./operations/aggregate-sheets')
const getWorkBook = require('./operations/get-book')

const commandFolder = './commandBook/commandBook.xlsx'


const commandBook = xlsx.readFile(commandFolder)
const { Sheets } = commandBook
const commandSheet = Sheets['commandSheet']
const {s: { r: start }, e: { r: end }} = xlsx.utils.decode_range(commandSheet['!ref'])
for (let i=start+1; i<=end+1; i++) {
    const command = commandSheet[`A${i}`].v
    switch(command){
        case 'READ': {
            xlsx.utils.sheet_add_aoa(commandSheet, [readAllBooks()], {origin: `B${i}`})
            break
        }
        case 'READ_BOOK': {
            const workBookName = commandSheet[`B${i}`].v
            xlsx.utils.sheet_add_aoa(commandSheet, [readBook(workBookName)], {origin: `C${i}`})
            break
        }
        case 'RENAME': {
            const workBookName = commandSheet[`B${i}`].v
            const config = commandSheet[`C${i}`].v
            const workBook = rename(workBookName, config)
            xlsx.writeFile(workBook, `./resourceBook/${workBookName}`)
            break
        }
        case 'ADD': {
            const sourceWorkBookName = commandSheet[`B${i}`].v
            const sheetNames = commandSheet[`C${i}`].v
            const destinationWorkBookName = commandSheet[`D${i}`].v
            const destinationWorkBook = add(getWorkBook(sourceWorkBookName), sheetNames, getWorkBook(destinationWorkBookName))
            xlsx.writeFile(destinationWorkBook, `./resourceBook/${destinationWorkBookName}`)
            break
        }
        case 'DELETE': {
            const workBookName = commandSheet[`B${i}`].v
            const sheetNames = commandSheet[`C${i}`].v
            const workBook = deleteSheets(workBookName, sheetNames)
            xlsx.writeFile(workBook, `./resourceBook/${workBookName}`)
            break
        }
        case 'AGG-WB': {
            const workBookNames = commandSheet[`B${i}`].v
            const newWorkBookName = commandSheet[`C${i}`].v
            const workBook = aggregateWorkBooks(workBookNames)
            xlsx.writeFile(workBook, `./resourceBook/${newWorkBookName}`)
            break
        }
        case 'AGG-WS': {
            const workBookName = commandSheet[`B${i}`].v
            let sheetNames = commandSheet[`C${i}`]
            if(!sheetNames){
                sheetNames = getWorkBook(workBookName)['SheetNames'].join(',')
            } else {
                sheetNames = sheetNames.v
            }
            const workBook = aggregateWorkSheets(getWorkBook(workBookName), sheetNames)
            xlsx.writeFile(workBook, `./resourceBook/${workBookName}`)
            break
        }
    }
    xlsx.writeFile(commandBook, commandFolder)
}

