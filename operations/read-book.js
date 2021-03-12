const getWorkBookInfo = require('./get-book-info')
const getWorkBookInfoAsString = require('./get-book-info-as-string')

const readBook = (workBookName) => [getWorkBookInfoAsString(workBookName)]

module.exports = readBook