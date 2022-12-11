const axios = require('axios');
const data ={};
const myCache = require("../myCache");
const {parseString} = require("xml2js");
const {createXml, formatJson, baseJsonXml} = require("../utils/XMLUtils");
const readSavedSheet = require("../utils/readSavedSheet");
const { GoogleSpreadsheet } = require('google-spreadsheet');
const doc = new GoogleSpreadsheet(process.env.GOOGLE_SAVED_REPORT);
const masterDoc = new GoogleSpreadsheet(process.env.GOOGLE_MASTER_REPORT);

const ind= (row,col) => {
    return `INDIRETTO(INDIRIZZO(${row};${col}))`
}
const singleSum = (col) => {
    return `SOMMA(${ind(45,col)}:${ind(47,col)})`
}
const sum12Months= (col, end = 12) => {
    const arr = []
    for (let i = 0; i < end ; i++) {
        arr.push(singleSum(col - (i*2)))
    }
    return `(${arr.join("+")})`
}

const getNumberSavedSearch = (formatData, index=0, values=[]) => {

    const requestXml= createXml(`
       
        <param>
            <value><int>${formatData[index].Id}</int></value>
        </param>
        <param>
            <value><int>2903</int></value>
        </param>
        <param>
            <value><int>0</int></value>
        </param> `
        , 'SearchService.getSavedSearchResultsAllFields')

    axios.post(`https://api.infusionsoft.com/crm/xmlrpc/v1?access_token=${myCache.get('tokens').accessToken}`, requestXml ,{
        headers: {
            'Content-Type': 'text/xml'
        }
    }).then(res => {
        parseString(res.data, function (err, result) {
            const data = baseJsonXml(result)
            const newValue = [...values, {...formatData[index], count: data?.length || 0}]
            if ( index < formatData.length-1 ) {
                getNumberSavedSearch(formatData, index +1 , newValue)
            }
            else {
                doc.useServiceAccountAuth({

                    client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
                    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),

                }).then(res =>{
                    doc.loadInfo().then(res =>{
                        const sheet = doc.sheetsByIndex[0]
                        sheet.loadCells('C:C').then(res =>{
                            newValue.forEach((ele, index) =>{
                                const cell = sheet.getCell(index, 2);
                                cell.value = ele.count;
                            })

                            sheet.saveUpdatedCells().then(res =>{
                                console.log('Saved Search Salvati')
                            });
                        });

                    })
                });

            }
        })
    })


}




data.getSavedSearch=() => {
        const requestXml= createXml(`<param>
      <value><string>SavedFilter</string></value>
    </param>
    <param>
      <value><int>10000</int></value>
    </param>
    <param>
      <value><int>0</int></value>
    </param>
    <param>
      <value><struct>
        <member><name>Id</name>
          <value><string>%</string></value>
        </member>
      </struct></value>
    </param>
    <param>
      <value><array>
        <data>
          <value><string>Id</string></value>
          <value><string>FilterName</string></value>
          <value><string>UserId</string></value>
          <value><string>ReportStoredName</string></value>
        </data>
      </array></value>
    </param>
`)


    axios.post(`https://api.infusionsoft.com/crm/xmlrpc/v1?access_token=${myCache.get('tokens').accessToken}`, requestXml ,{
        headers: {
            'Content-Type': 'text/xml'
        }
    }).then(res => {
        parseString(res.data, function (err, result) {
            readSavedSheet.getSavedSearch().then(savedSearch =>{
                const formatData = formatJson(result)
                const filterData = savedSearch.map(ele =>{
                    return formatData.find(el => el.FilterName === ele)
                })
                getNumberSavedSearch(filterData)
            })
        });
    }).catch(err =>{
        console.log(err)
    })
}


const listValueSearched = (sheet, callback, index = 0, list=[]) =>{
    const name = sheet.getCell(index, 1).value;
    const value = sheet.getCell(index, 2).value;
    const row = sheet.getCell(index, 4).value;
    if (name){
        listValueSearched(sheet, callback , index+1, [...list, {name, value, row}])
    }
    else {
        callback(list)
    }
}
// funzione trova la cella vuota, cicliamo fintanto che non si trova la colonna vuota
const getColumnData  = (sheet, index,rowIndex, resolve) => {
    const cell = sheet.getCell(rowIndex, index);
    cell.value ? getColumnData(sheet, index +1 ,rowIndex, resolve) : resolve(index)
}

// funzione trova la colonna vuota, cicliamo fintanto che si trova la colonna vuota e poi ci spostiamo
const getColumnEmpty = (sheet, firstColumn = 0, rowIndex = 0) => {


    return new Promise(resolve => {
        getColumnData(sheet, firstColumn, rowIndex, resolve)
    } )
}

const offices = {
    ssa : {
        index : 0,
        value : 'M2',
        clients: 'P2',
        payments: 'M3',
        leadsource: 'S:T',
        unique3m: 'P4',
        unique12m: 'P6'
    },
    pea : {
        index : 1,
        value : 'N2',
        clients: 'P3',
        payments: 'N3',
        leadsource: 'U:V',
        unique3m: 'P5',
        unique12m: 'P7'
    }
};

const operations = (doc,office, newList) => {
    const masterSheet = doc.sheetsByIndex[offices[office].index] // CAPELLI_SS
    masterSheet.loadCells('').then(res => {
        getColumnEmpty(masterSheet, 1).then(columnId => {
            let columnIdEmpty = columnId;
            const isAddCell = (columnIdEmpty + 4)>= masterSheet.columnCount
            const cellFormula = masterSheet.getCell(1, columnIdEmpty+2);
            if (isAddCell) {
                cellFormula.value = '=addColumns(12)'
            }
            const d = process.env.DATAESTRAZIONE ? new Date(process.env.DATAESTRAZIONE) : new Date();
            const today = new Date (d.setMonth(d.getMonth() - 1))
            const titleDate = `${today.toLocaleString('it-IT', { month: 'long' })}-${today.getFullYear()}`
            let cell2 = masterSheet.getCell(0, columnIdEmpty-2);
            if (cell2.value === titleDate){
                columnIdEmpty = columnIdEmpty - 2
                cell2 = masterSheet.getCell(0, columnIdEmpty-2);
            }
            const cellMonth = masterSheet.getCell(0,columnIdEmpty);
            const columnDiff= columnIdEmpty+1
            const cellMonthDifference = masterSheet.getCell(0,columnDiff);
            cellMonthDifference.value = 'D';
            cellMonth.value= titleDate
            newList[office].forEach(ele => {
                const cell = masterSheet.getCell(ele.row-1, columnIdEmpty);
                const cell2 = masterSheet.getCell(ele.row-1, columnIdEmpty-2);
                const isGray = cell2.effectiveFormat.backgroundColor.red ===1
                cell.value = ele.value;
                const cellDiff = masterSheet.getCell(ele.row-1,columnDiff)

                cellDiff.value = `=INDIRETTO(INDIRIZZO(${ele.row};${columnIdEmpty+1}))-INDIRETTO(INDIRIZZO(${ele.row};${columnIdEmpty-1}))`;

                if (isGray) {
                    cell.backgroundColor = {
                        "red": 80,
                        "green": 80,
                        "blue": 80,
                        "alpha": 1
                    }
                    cellDiff.backgroundColor = {
                        "red": 80,
                        "green": 80,
                        "blue": 80,
                        "alpha": 1
                    }
                }

            })
            const cellGuadagno =  masterSheet.getCell(40, columnIdEmpty);
            const totaleSpese =  masterSheet.getCell(47, columnIdEmpty);
            const roi = masterSheet.getCell(48, columnIdEmpty);
            const numClienti = masterSheet.getCell(49, columnIdEmpty);
            const totVendite = masterSheet.getCell(50, columnIdEmpty);
            const ltvMensile = masterSheet.getCell(51, columnIdEmpty);
            const guadagno3Mesi= masterSheet.getCell(52, columnIdEmpty)
            const clienti3Mesi= masterSheet.getCell(53, columnIdEmpty)
            const vendite3Mesi= masterSheet.getCell(54, columnIdEmpty)
            const ltv3Mesi= masterSheet.getCell(55, columnIdEmpty)
            const guadagno12Mesi= masterSheet.getCell(56, columnIdEmpty)
            const clienti12Mesi= masterSheet.getCell(57, columnIdEmpty)
            const vendite12Mesi= masterSheet.getCell(58, columnIdEmpty)
            const ltv12Mesi= masterSheet.getCell(59, columnIdEmpty)
            const leadsource=  masterSheet.getCell(63, columnIdEmpty)
            const roi3Month= masterSheet.getCell(61, columnIdEmpty);
            const roi12Month=  masterSheet.getCell(62, columnIdEmpty);;
            cellGuadagno.value=`=IMPORTRANGE("1DA3_6pkQ-35BIi5smULwyK405gh8Qzudrel37v_ohqU";"${today.toLocaleString('it-IT', { month: 'long' })}_${today.getFullYear()}!${offices[office].value}")`;
            totaleSpese.value= `=INDIRETTO(INDIRIZZO(41;${columnIdEmpty+1}))-SUM(INDIRETTO(INDIRIZZO(45;${columnIdEmpty+1})):INDIRETTO(INDIRIZZO(47;${columnIdEmpty+1})))`
            roi.value=`=((INDIRETTO(INDIRIZZO(41;${columnIdEmpty+1}))-SOMMA(INDIRETTO(INDIRIZZO(45;${columnIdEmpty+1})):INDIRETTO(INDIRIZZO(47;${columnIdEmpty+1}))))/SOMMA(INDIRETTO(INDIRIZZO(45;${columnIdEmpty+1})):INDIRETTO(INDIRIZZO(47;${columnIdEmpty+1}))))`
            numClienti.value=`=IMPORTRANGE("1DA3_6pkQ-35BIi5smULwyK405gh8Qzudrel37v_ohqU";"${today.toLocaleString('it-IT', { month: 'long' })}_${today.getFullYear()}!${offices[office].clients}")`;
            totVendite.value=`=IMPORTRANGE("1DA3_6pkQ-35BIi5smULwyK405gh8Qzudrel37v_ohqU";"${today.toLocaleString('it-IT', { month: 'long' })}_${today.getFullYear()}!${offices[office].payments}")`;
            ltvMensile.value= `=INDIRETTO(INDIRIZZO(50;${columnIdEmpty+1}))/INDIRETTO(INDIRIZZO(49;${columnIdEmpty+1}))`
            guadagno3Mesi.value=`=INDIRETTO(INDIRIZZO(48;${columnIdEmpty+1}))+INDIRETTO(INDIRIZZO(48;${columnIdEmpty-1}))+INDIRETTO(INDIRIZZO(48;${columnIdEmpty-3}))`
            vendite3Mesi.value=`=INDIRETTO(INDIRIZZO(51;${columnIdEmpty+1}))+INDIRETTO(INDIRIZZO(51;${columnIdEmpty-1}))+INDIRETTO(INDIRIZZO(51;${columnIdEmpty-3}))`
            clienti3Mesi.value=`=IMPORTRANGE("1DA3_6pkQ-35BIi5smULwyK405gh8Qzudrel37v_ohqU";"${today.toLocaleString('it-IT', { month: 'long' })}_${today.getFullYear()}!${offices[office].unique3m}")`;
            ltv3Mesi.value=`=INDIRETTO(INDIRIZZO(53;${columnIdEmpty+1}))/INDIRETTO(INDIRIZZO(54;${columnIdEmpty+1}))`
            guadagno12Mesi.value=`=INDIRETTO(INDIRIZZO(48;${columnIdEmpty+1}))+INDIRETTO(INDIRIZZO(48;${columnIdEmpty-1}))+INDIRETTO(INDIRIZZO(48;${columnIdEmpty-3}))+INDIRETTO(INDIRIZZO(48;${columnIdEmpty-5}))+INDIRETTO(INDIRIZZO(48;${columnIdEmpty-7}))
                                    +INDIRETTO(INDIRIZZO(48;${columnIdEmpty-9}))+INDIRETTO(INDIRIZZO(48;${columnIdEmpty-11}))+INDIRETTO(INDIRIZZO(48;${columnIdEmpty-13}))+INDIRETTO(INDIRIZZO(48;${columnIdEmpty-15}))+INDIRETTO(INDIRIZZO(48;${columnIdEmpty-17}))
                                    +INDIRETTO(INDIRIZZO(48;${columnIdEmpty-19}))+INDIRETTO(INDIRIZZO(48;${columnIdEmpty-21}))`
            vendite12Mesi.value=`=INDIRETTO(INDIRIZZO(51;${columnIdEmpty+1}))+INDIRETTO(INDIRIZZO(51;${columnIdEmpty-1}))+INDIRETTO(INDIRIZZO(51;${columnIdEmpty-3}))+INDIRETTO(INDIRIZZO(51;${columnIdEmpty-5}))+INDIRETTO(INDIRIZZO(51;${columnIdEmpty-7}))
                                    +INDIRETTO(INDIRIZZO(51;${columnIdEmpty-9}))+INDIRETTO(INDIRIZZO(51;${columnIdEmpty-11}))+INDIRETTO(INDIRIZZO(51;${columnIdEmpty-13}))+INDIRETTO(INDIRIZZO(51;${columnIdEmpty-15}))+INDIRETTO(INDIRIZZO(51;${columnIdEmpty-17}))
                                    +INDIRETTO(INDIRIZZO(51;${columnIdEmpty-19}))+INDIRETTO(INDIRIZZO(51;${columnIdEmpty-21}))`
            clienti12Mesi.value=`=IMPORTRANGE("1DA3_6pkQ-35BIi5smULwyK405gh8Qzudrel37v_ohqU";"${today.toLocaleString('it-IT', { month: 'long' })}_${today.getFullYear()}!${offices[office].unique12m}")`;
            ltv12Mesi.value=`=INDIRETTO(INDIRIZZO(57;${columnIdEmpty+1}))/INDIRETTO(INDIRIZZO(58;${columnIdEmpty+1}))`
            leadsource.value=`=IMPORTRANGE("1DA3_6pkQ-35BIi5smULwyK405gh8Qzudrel37v_ohqU";"${today.toLocaleString('it-IT', { month: 'long' })}_${today.getFullYear()}!${offices[office].leadsource}")`;
            roi12Month.value=`=((${ind(41,columnIdEmpty+1)}+${ind(41,columnIdEmpty-1)}+${ind(41,columnIdEmpty-3)}+${ind(41, columnIdEmpty-3)}+
                               ${ind(41,columnIdEmpty-5)}+${ind(41,columnIdEmpty-7)}+${ind(41,columnIdEmpty-9)}+${ind(41,columnIdEmpty-9)}+
                               ${ind(41,columnIdEmpty-11)}+${ind(41,columnIdEmpty-13)}+${ind(41,columnIdEmpty-15)}+${ind(41,columnIdEmpty-17)}+
                               ${ind(41,columnIdEmpty-19)}+${ind(41,columnIdEmpty-21)})-${sum12Months(columnIdEmpty+1)})/${sum12Months(columnIdEmpty+1)}
                                `

            roi3Month.value=`=((${ind(41,columnIdEmpty-3)}+INDIRETTO(INDIRIZZO(41;${columnIdEmpty-1}))+INDIRETTO(INDIRIZZO(41;${columnIdEmpty+1})))-(SOMMA(INDIRETTO(INDIRIZZO(45;${columnIdEmpty-3})):INDIRETTO(INDIRIZZO(47;${columnIdEmpty+1})))+SOMMA(INDIRETTO(INDIRIZZO(45;${columnIdEmpty-3})):INDIRETTO(INDIRIZZO(47;${columnIdEmpty+1})))+SOMMA(INDIRETTO(INDIRIZZO(45;${columnIdEmpty-3})):INDIRETTO(INDIRIZZO(47;${columnIdEmpty+1})))))/((SOMMA(INDIRETTO(INDIRIZZO(45;${columnIdEmpty-3})):INDIRETTO(INDIRIZZO(47;${columnIdEmpty+1})))+SOMMA(INDIRETTO(INDIRIZZO(45;${columnIdEmpty-3})):INDIRETTO(INDIRIZZO(47;${columnIdEmpty+1})))+SOMMA(INDIRETTO(INDIRIZZO(45;${columnIdEmpty-3})):INDIRETTO(INDIRIZZO(47;${columnIdEmpty+1})))))    `
            masterSheet.saveUpdatedCells().then(res =>{

                if (isAddCell){
                    cellFormula.value=''
                    masterSheet.saveUpdatedCells().then(res =>{ console.log('formula colonne tolte')})
                }
            });

        })
    })
}


data.createArraySync = () => {
    doc.useServiceAccountAuth({

        client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),

    }).then(res =>{
        doc.loadInfo().then(res =>{
            const sheet = doc.sheetsByIndex[0]
            sheet.loadCells('B:E').then(res =>{    // legge le righe da b a e dello sheet e fa un'array di SSA e PEA
                                                   // per poi poterlo mettere nell'altro foglio con row name e value
                const callback = (list) => {
                    const newList = list.reduce((acc, ele) =>{
                            if (ele.name.startsWith('SSA')){
                                acc.ssa.push(ele)
                            }
                            if (ele.name.startsWith('PEA')){
                                acc.pea.push(ele)
                            }
                            return acc
                         }, {
                                ssa: [],
                                pea: []
                    })
                    // prese il nuovo newList vado a leggere il masterDoc per andare
                    // a salvare i risultati nella cella corrispondente
                       masterDoc.useServiceAccountAuth({

                           client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
                           private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),

                       }).then(res => {
                           masterDoc.loadInfo().then(res =>{
                               operations(masterDoc,'ssa', newList)
                               operations(masterDoc,'pea', newList)
                           })
                       })


                }
                listValueSearched(sheet,callback)


            });

        })
    });
}

module.exports = data;
