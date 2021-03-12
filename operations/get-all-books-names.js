const fs = require('fs')
const path = require('path')

const resourceFolder = path.join(__dirname, '..', 'resourceBook')

const getAllWorkBooksNames = () => fs.readdirSync(resourceFolder)

module.exports = getAllWorkBooksNames