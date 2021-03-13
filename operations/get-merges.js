const getWorkBook = require('./get-book')
const xlsx = require('xlsx')

const getMerges = (workBookName, workSheetName) => {
    const workBook = getWorkBook(workBookName)
    const { Sheets } = workBook
    const sheet = Sheets[workSheetName]
    const merges = sheet['!merges']
    const mergesList = merges.reduce((acc, range) => {
        const { s, e } = range
        const rangeString = xlsx.utils.encode_range(s, e)
        acc.push(rangeString)
        return acc
    }, [])
    return mergesList.join(',')
}

module.exports = getMerges