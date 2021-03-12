const add = require('./add')
const getWorkBook = require('./get-book')

const aggregateWorkBooks = (workBooks) => {
    const workBookList = workBooks.split(',')
    let workBook = {
        Sheets: {},
        SheetNames: []
    }
    workBookList.forEach((workBookName) => {
        const sourceWorkBook = getWorkBook(workBookName)
        workBook = add(sourceWorkBook, sourceWorkBook.SheetNames.join(','), workBook)
    })
    return workBook
}

module.exports = aggregateWorkBooks