const getWorkBookInfo = require('./get-book-info')
const getAllWorkBooksNames = require(getAllBooksNames)


const getAllBooks = () => {
    const workBookNames = getAllWorkBooksNames()
    return workBookNames.reduce((acc, workBookName) => {
        acc[workBookName] = getWorkBookInfo(workBookName)
        return acc
    }, {})
}

module.exports = getAllBooks