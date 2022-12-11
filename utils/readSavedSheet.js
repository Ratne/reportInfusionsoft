const data ={}


const { GoogleSpreadsheet } = require('google-spreadsheet');

const doc = new GoogleSpreadsheet('1rXQmljMXbpTGIR3q_Jeu4UP7271uqgLYpw0EEY_Qgj4');
data.getSavedSearch = () =>{

    return new Promise(function(resolve, reject){
        doc.useServiceAccountAuth({

            client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
            private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),

        }).then(res =>{
            doc.loadInfo().then(res =>{
                const sheet = doc.sheetsByIndex[0]
                sheet.loadCells('B:B').then(res =>{
                    listValue(sheet, resolve)
                });

            })
        });
    });
}
// legge i risultati delle celle finchè non ne trova una vuota
const listValue = (sheet, callback, index = 0, list=[]) =>{
    const cell = sheet.getCell(index, 1);
    if (cell.value){
        listValue(sheet, callback , index+1, [...list, cell.value])
    }
    else {
        callback(list)
    }
}



module.exports = data;
