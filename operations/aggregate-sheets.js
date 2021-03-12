const compareHeaders = require('./compare-headers')
const xlsx = require('xlsx')

const isHeaderCell = (cell) => {
    return xlsx.utils.decode_cell(cell).r === 0
}

const aggregateWorkSheets = (workBook, sheetNames) => {
    const distinctSheets = []
    const { Sheets, SheetNames } = workBook
    const sheetNamesList = sheetNames.split(',')
    sheetNamesList.forEach(sheetName => {
        let dirty = false
        for(let i=0; i<distinctSheets.length; i++){
            if(compareHeaders(Sheets[distinctSheets[i][0]], Sheets[sheetName])) {
                distinctSheets[i].push(sheetName)
                dirty = true
                break
            }
        }
        if(!dirty) distinctSheets.push([sheetName])
    })
    const sheetSets = distinctSheets.filter(list => list.length > 1)
    sheetSets.forEach((list) => {
        var arr = []
        var header = []
        var name = ''
        list.forEach(sheetName => {
            xlsx.utils.sheet_to_json(Sheets[sheetName]).forEach(obj => arr.push(obj))
            if(!header.length){
                header = Object.keys(Sheets[sheetName]).reduce((acc,cell) => {
                    if(/^[A-Z0-9]*$/.test(cell) && isHeaderCell(cell)){
                        acc.push(Sheets[sheetName][cell].v)
                    }
                    return acc
                }, [])
            }
            if (!name) {
                name = sheetName
            }
        })
        const newSheet = xlsx.utils.json_to_sheet(arr, {header: header})
        Sheets[`${name}-${list.length}-aggr`] = newSheet
        SheetNames.push(`${name}-${list.length}-aggr`)
    })
    return workBook
}

module.exports = aggregateWorkSheets