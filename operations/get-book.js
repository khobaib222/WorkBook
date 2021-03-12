const path = require('path')
const xlsx = require('xlsx')

const resourceFolder = path.join(__dirname, '..', 'resourceBook')

const getWorkBook = (workBookName) => xlsx.readFile(`${resourceFolder}/${workBookName}`)

module.exports = getWorkBook