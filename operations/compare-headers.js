const xlsx = require('xlsx')

const isHeaderCell = (cell) => {
    return xlsx.utils.decode_cell(cell).r === 0
}

const getHeaderCells = (workSheet) => Object.keys(workSheet).filter(cell => /^[A-Z0-9]*$/.test(cell) && isHeaderCell(cell))

const compareHeaders = (workSheet, workSheet1) => {
    const cells = getHeaderCells(workSheet)
    const cells1 = getHeaderCells(workSheet1)
    if(cells.length!==cells1.length) return false
    cells.forEach(cell => {
        if(workSheet[cell].v !== workSheet1[cell].v) return false
    })
    return true
}

module.exports = compareHeaders