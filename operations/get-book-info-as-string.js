const getWorkBookInfo = require('./get-book-info')

const getWorkBookInfoAsArray = (workBookName) => {
    const { filledSheetNames, emptySheetNames } = getWorkBookInfo(workBookName)
    let val = ''
    if (filledSheetNames.length) {
        val = `${filledSheetNames.join(',')}`
    }
    if (emptySheetNames.length) {
        if(!val){
            val += '_'
        }
        val += `|${emptySheetNames.join(',')}`
    }
    return val
}

module.exports = getWorkBookInfoAsArray