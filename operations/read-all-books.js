const getWorkBookInfoAsString = require('./get-book-info-as-string')
const getAllWorkBooksNames = require('./get-all-books-names')

const readAllBooks = () => getAllWorkBooksNames().reduce((acc, workBookName) => acc.concat([workBookName, getWorkBookInfoAsString(workBookName)]), [])

module.exports = readAllBooks