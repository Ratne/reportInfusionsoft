const { GoogleSpreadsheet } = require('google-spreadsheet');
const data = {};
data.createDoc = (sheetId) => {
    return new Promise((resolve) => {
        const doc = new GoogleSpreadsheet(sheetId);
        doc.useServiceAccountAuth({
            client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
            private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),
        }).then(res =>{
            doc.loadInfo().then(res =>{
                resolve(doc)
            }).catch(err => {
                console.log(err, 'load info failed')
            })
        }).catch(err => {
            console.log(err, 'non carica doc');
        })
    })

}
const getColumnData  = (sheet, index,rowIndex, resolve) => {
    const cell = sheet.getCell(rowIndex, index);
    cell.value ? getColumnData(sheet, index +1 ,rowIndex, resolve) : resolve(index)
}
data.getColumnEmpty = (sheet, firstColumn = 0, rowIndex = 0) => {
    return new Promise(resolve => {
        getColumnData(sheet, firstColumn, rowIndex, resolve)
    } )
}
data.getInd= (row,col) => {
    return `INDIRETTO(INDIRIZZO(${row};${col}))`
}
data.addColumnsIfFinish = (columnIdEmpty, sheet) => {
    const isAddCell = (columnIdEmpty + 4)>= sheet.columnCount
    const cellFormula = sheet.getCell(1, columnIdEmpty+2);
    if (isAddCell) {
        cellFormula.value = '=addColumns(12)'
    }
}

data.getSheetIfExistOrCreate = (doc, title) => {
    return new Promise((resolve) => {
        const sheet = doc.sheetsByTitle[title];
        if(sheet){
            resolve(sheet)
        }else {
            doc.addSheet({title }).then(res => {
                resolve(res);
            })
        }
    })
}

module.exports = data;
