const { GoogleSpreadsheet } = require('google-spreadsheet');
const data = {}
const doc = new GoogleSpreadsheet('1DA3_6pkQ-35BIi5smULwyK405gh8Qzudrel37v_ohqU');
const dataName = require('./getData')
const {formatJsonLeadsource, createXml} = require("../utils/XMLUtils");
const {calcMonthPrev, calcTitle} = require("../utils/dateUtils");
const axios = require("axios");
const myCache = require("../myCache");
const {parseString} = require("xml2js");

const titles = ['Nome prodotto','Ordine specifico infusionsoft','Prezzo Ordine Infusionsoft', 'Nome','Email','Sede',
'Guadagno','Prezzo Vendita','','ID PRODOTTO','COSTO PRODOTTO', 'IVA PRODOTTO','TOTALE SAN SALVO','TOTALE PESCARA','Clienti Univoci','Valore','Totale Pagato','Valore','Leadsource SS',
'Valore','Leadsource Pescara','Valore','Univoci San Salvo','Univoci Pescara']
let productInfo= []

const listValue = (sheet, callback, index = 0, list=[]) =>{
   const cell1 = sheet.getCell(index, 9);
   const cell2 = sheet.getCell(index, 10);
    const cell3 = sheet.getCell(index, 11);

    if (cell1.value){
       listValue(sheet, callback , index+1, [...list, {
           productId: cell1.value,
           productPrice: cell2.value,
           productTax: cell3.value
       }])
    }
   else {
       productInfo = list;
       callback()
   }
}

const getProductInfo = (sheet,callback) => {
     sheet.loadCells('J:X').then(res => {
        listValue(sheet, callback)

   })
}


const getContactLeadsource = (list=[], callback, index = 0, newList = []) =>{
    const contact = list[index]
    // chiamata xml per prendere il leadsource
    // todo mettere xmlutils e ciaone
    const requestXml= `<?xml version='1.0' encoding='UTF-8'?>
                <methodCall>
                <methodName>ContactService.load</methodName>
                  <params>
                    <param>
                      <value><string>${process.env.CLIENTSECRET}</string></value>
                    </param>
                    <param>
                      <value><int>${contact.id}</int></value>
                    </param>
                    <param>
                      <value><array>
                        <data>
                          <value><string>Leadsource</string></value>
                          <value><string>LeadSourceId</string></value>
                        </data>
                      </array></value>
                    </param>
                  </params>
                </methodCall>`
    axios.post(`https://api.infusionsoft.com/crm/xmlrpc/v1?access_token=${myCache.get('tokens').accessToken}`, requestXml ,{
        headers: {
            'Content-Type': 'text/xml'
        }
    }).then(res => {

        parseString(res.data, function (err, result) {


            const data = formatJsonLeadsource(result)
            const contactUpdate = {...contact, ...data}
            // fare chiamata a prender ele tags
            axios.get(`https://api.infusionsoft.com/crm/rest/contacts/${contact.id}/tags`,
                {
                    headers: {
                        Accept: 'application/json, */*',
                        Authorization: 'Bearer '+myCache.get('tokens').accessToken
                    }})
                .then (res => {
                    contactUpdate.tags = res.data.tags
                    if ( index < list.length-1 ) {
                        getContactLeadsource(list, callback, index + 1, [...newList, contactUpdate])
                    }
                    else {
                        callback([...newList, contactUpdate])
                    }
                })

        })
    })
}

// legge i risultati delle celle finchè non ne trova una vuota
const listBlankValue = (sheet, callback, index = 0, columnIndex = 1, list=[]) =>{
    const cell = sheet.getCell(index, columnIndex);
    if (cell.value){
        listBlankValue(sheet, callback , index+1, columnIndex, [...list, cell.value])
    }
    else {
        callback(list)
    }
}


const calcQuarterUnique = (date, number, clk, columnIndex, list=[])  => {
    if (number === 0) {
        clk(list)
    }
    else {

        const prevMonth = calcMonthPrev(date);
        const prevSheet = doc.sheetsByTitle[calcTitle(date)]
        prevSheet.loadCells("W:X").then(res => {
            const internalCallback = (l) => {
                calcQuarterUnique(prevMonth, number -1, clk, columnIndex,[...list, ...l]  )
            }
            listBlankValue(prevSheet,internalCallback, 1, columnIndex)

        })
    }

}



const orderWrite = (sheet, products, lastMonth) => {
    sheet.loadCells('A:X').then(res =>{
        let base = 1;
        titles.forEach((ele, index) => {
            const cell = sheet.getCell(0, index);
            cell.value = ele
        })
        const totaleSS=[];
        const totalePE=[];
        let uniqueSS=[];
        let uniquePE=[];
        products.forEach((prod, prodIndex) => {
            const cell10 = sheet.getCell(prodIndex + 1, 9);
            const cell11 = sheet.getCell(prodIndex + 1, 10);
            const cell12 = sheet.getCell(prodIndex + 1, 11);
            cell10.value= prod.productId;
            const singleProductInfo = productInfo.find(ele => prod.productId === ele.productId)
            cell11.value= singleProductInfo ? singleProductInfo.productPrice : ''
            cell12.value= singleProductInfo ? singleProductInfo.productTax : ''

            prod.orders.forEach((order, index) => {

                const cell = sheet.getCell(index+base, 0);
                const cell2 = sheet.getCell(index+base, 1);
                const cell3 = sheet.getCell(index+base, 2);
                const cell4 = sheet.getCell(index+base, 3);
                const cell5 = sheet.getCell(index+base, 4);
                const cell6 = sheet.getCell(index+base, 5);
                const cell7 = sheet.getCell(index+base, 6);
                const cell8 = sheet.getCell(index+base, 7);

                cell.value = order.item.name;
                cell2.value = order.item.id;
                cell3.value = order.item.price;
                cell4.value = order.contact.first_name
                cell5.value= order.contact.email
                cell6.value= order.office
                order.office === 'San Salvo' && totaleSS.push(index+base);
                order.office === 'San Salvo' && !uniqueSS.find(ele => ele.email === order.contact.email) && uniqueSS.push(order.contact);
                order.office === 'Pescara' && totalePE.push(index+base);
                order.office === 'Pescara' && !uniquePE.find(ele => ele.email === order.contact.email) && uniquePE.push(order.contact)
                cell7.value= `=(H${index +1 +base}-K${prodIndex + 2})`
                cell8.value= `=(C${index +1 +base}*(100-L${prodIndex + 2})/100)`

                if (index === (prod.orders.length-1)){
                    base = base + (index + 1)
                }
            } )

        })


        const cell8 = sheet.getCell(1, 12);
        const cell9 = sheet.getCell(1, 13);
        const cell8Bis = sheet.getCell(2, 12);
        const cell9Bis = sheet.getCell(2, 13);
        cell8.value= `=(${totaleSS.map(ele => 'G'+(ele+1)).join('+')})`
        cell9.value= `=(${totalePE.map(ele => 'G'+(ele+1)).join('+')})`
        cell8Bis.value= `=(${totaleSS.map(ele => 'H'+(ele+1)).join('+')})`
        cell9Bis.value= `=(${totalePE.map(ele => 'H'+(ele+1)).join('+')})`


        const cell10 = sheet.getCell(1, 17);
        const cell11 = sheet.getCell(2, 17);
        cell10.value= `=(${totaleSS.map(ele => 'C'+(ele+1)).join('+')})`
        cell11.value= `=(${totalePE.map(ele => 'C'+(ele+1)).join('+')})`
        const cell12 = sheet.getCell(1, 15);
        const cell13 = sheet.getCell(2, 15);
        const cell14 = sheet.getCell(1, 14);
        const cell15 = sheet.getCell(2, 14);
        const cell16 = sheet.getCell(1, 16);
        const cell17 = sheet.getCell(2, 16);
        const cell18 = sheet.getCell(3, 14);
        const cell19 = sheet.getCell(4, 14);
        const cell20 = sheet.getCell(5, 14);
        const cell21 = sheet.getCell(6, 14);

        cell14.value="San Salvo";
        cell15.value="Pescara";
        cell16.value="San Salvo";
        cell17.value="Pescara";
        cell12.value= `${uniqueSS.length}`
        cell13.value= `=(${uniquePE.length})`
        cell18.value="San Salvo 3M";
        cell19.value="Pescara 3M";
        cell20.value="San Salvo 12M";
        cell21.value="Pescara 12M";
        const callback = (list) => {
            uniqueSS = list;
            const call = (l) => {
                uniquePE = l;
                const leadListSS = uniqueSS.reduce((acc, ele) => {
                    const data = acc.find(el => el.LeadSourceId === ele.LeadSourceId);

                    return [...acc.filter(e => e.LeadSourceId !== ele.LeadSourceId), data ? {...data, quantity: data.quantity + 1} : {
                        LeadSourceId : ele.LeadSourceId,
                        Leadsource: ele.Leadsource,
                        quantity: 1,
                        tags: ele.tags
                    }]
                }, []);
                const leadListPE =  uniquePE.reduce((acc, ele) => {
                    const data = acc.find(el => el.LeadSourceId === ele.LeadSourceId);

                    return [...acc.filter(e => e.LeadSourceId !== ele.LeadSourceId), data ? {...data, quantity: data.quantity + 1} : {
                        LeadSourceId : ele.LeadSourceId,
                        Leadsource: ele.Leadsource || 'VUOTO',
                        quantity: 1,
                        tags: ele.tags
                    }]
                }, [])
                console.log("Leadsource estrapolati");

                let indexList = 0;
                leadListSS.forEach((ele , index) => {
                    const cellLeadSourceSS = sheet.getCell(index +1, 18);
                    const cellLeadSourceValueSS = sheet.getCell(index +1, 19);
                    cellLeadSourceSS.value = ele.Leadsource || 'VUOTO'
                    cellLeadSourceValueSS.value = ele.quantity
                    indexList = index
                })
                const cellLeadSourceSS = sheet.getCell(indexList + 5, 18);
                cellLeadSourceSS.value="Leadsource Inserzioni"
                cellLeadSourceSS.backgroundColor = {
                    "red": 80,
                    "green": 80,
                    "blue": 80,
                    "alpha": 1
                }
                indexList = indexList + 5
                const leadListAds = leadListSS.filter(ele => ele.Leadsource?.startsWith('[FB]') || ele.Leadsource?.startsWith('[GOO]') )

                leadListAds.forEach((ele , index) => {

                              const cellLeadSourceSS = sheet.getCell(index + indexList +1, 18);
                              const cellLeadSourceValueSS = sheet.getCell(index + indexList +1, 19);
                              cellLeadSourceSS.value = ele.Leadsource || 'VUOTO'
                              cellLeadSourceValueSS.value = ele.quantity

                          })


                indexList = indexList + leadListAds.length + 5
                const cellLeadSourceWalkInSS = sheet.getCell(indexList, 18);
                cellLeadSourceWalkInSS.value="Leadsource Walk-in"
                cellLeadSourceWalkInSS.backgroundColor = {
                    "red": 80,
                    "green": 80,
                    "blue": 80,
                    "alpha": 1
                }



                const leadListWalkInSS = leadListSS.filter(ele => ele.tags.find(tag => {

                    const d = calcMonthPrev(lastMonth)
                    const date = new Date(tag.date_applied)

                    return [409].includes(tag.tag.id) &&
                        d.getMonth() === date.getMonth() &&
                        d.getFullYear() === date.getFullYear()
                } ))

                leadListWalkInSS.forEach((ele , index) => {

                    const cellLeadSourceSS = sheet.getCell(index + indexList +1, 18);
                    const cellLeadSourceValueSS = sheet.getCell(index + indexList +1, 19);
                    cellLeadSourceSS.value = ele.Leadsource || 'VUOTO'
                    cellLeadSourceValueSS.value = ele.quantity
                })

                indexList = indexList + leadListWalkInSS.length + 5
                const leadListVariousTitleSS = sheet.getCell(indexList, 18);
                leadListVariousTitleSS.value="Leadsource Varie"
                leadListVariousTitleSS.backgroundColor = {
                    "red": 80,
                    "green": 80,
                    "blue": 80,
                    "alpha": 1
                }
                const leadListCheckUpSS = leadListSS.filter(ele => ele.tags.find(tag => {
                    const d = calcMonthPrev(lastMonth)
                    const date = new Date(tag.date_applied)
                    return [674,1094].includes(tag.tag.id) &&
                        d.getMonth() === date.getMonth() &&
                        d.getFullYear() === date.getFullYear()
                } ))

                // calcolo varie
               const leadListVariousSS = leadListSS.filter(ele => !leadListCheckUpSS.find(el => el.LeadSourceId === ele.LeadSourceId) &&
                   !leadListWalkInSS.find(el => el.LeadSourceId === ele.LeadSourceId) &&
                   !leadListAds.find(el => el.LeadSourceId === ele.LeadSourceId)    )
                leadListVariousSS.forEach((ele , index) => {

                    const cellLeadSourceSS = sheet.getCell(index + indexList +1, 18);
                    const cellLeadSourceValueSS = sheet.getCell(index + indexList +1, 19);
                    cellLeadSourceSS.value = ele.Leadsource || 'VUOTO'
                    cellLeadSourceValueSS.value = ele.quantity
                })
                indexList = indexList + leadListVariousSS.length + 5
                const cellLeadSourceTitleSS = sheet.getCell(indexList, 18);
                cellLeadSourceTitleSS.value="Leadsource Checkup"
                cellLeadSourceTitleSS.backgroundColor = {
                    "red": 80,
                    "green": 80,
                    "blue": 80,
                    "alpha": 1
                }

                leadListCheckUpSS.forEach((ele , index) => {

                        const cellLeadSourceSS = sheet.getCell(index + indexList +1, 18);
                        const cellLeadSourceValueSS = sheet.getCell(index + indexList +1, 19);
                        cellLeadSourceSS.value = ele.Leadsource || 'VUOTO'
                        cellLeadSourceValueSS.value = ele.quantity
                    })

                indexList = 0;

                leadListPE.forEach((ele , index) => {
                    const cellLeadSourcePE = sheet.getCell(index +1, 20);
                    const cellLeadSourceValuePE = sheet.getCell(index +1, 21);
                    cellLeadSourcePE.value = ele.Leadsource
                    cellLeadSourceValuePE.value = ele.quantity
                    indexList = index
                })
                const cellLeadSourcePE = sheet.getCell(indexList + 5, 20);
                cellLeadSourcePE.value="Leadsource Inserzioni"
                cellLeadSourcePE.backgroundColor = {
                    "red": 80,
                    "green": 80,
                    "blue": 80,
                    "alpha": 1
                }
                indexList = indexList + 5
                const loadListAdsPE = leadListPE.filter(ele => ele.Leadsource?.startsWith('[FB]') || ele.Leadsource?.startsWith('[GOO]') )
                loadListAdsPE.forEach((ele , index) => {
                        const cellLeadSourcePE = sheet.getCell(index + indexList +1, 20);
                        const cellLeadSourceValuePE = sheet.getCell(index + indexList +1, 21);
                        cellLeadSourcePE.value = ele.Leadsource || 'VUOTO'
                        cellLeadSourceValuePE.value = ele.quantity

                    })
                indexList = indexList + loadListAdsPE.length + 5

                const cellLeadSourceWalkInPE = sheet.getCell(indexList, 20);
                cellLeadSourceWalkInPE.value="Leadsource Walk-in"
                cellLeadSourceWalkInPE.backgroundColor = {
                    "red": 80,
                    "green": 80,
                    "blue": 80,
                    "alpha": 1
                }
                const leadListWalkInPE = leadListPE.filter(ele => ele.tags.find(tag => {
                    const d = calcMonthPrev(lastMonth)
                    const date = new Date(tag.date_applied)
                    return [409].includes(tag.tag.id) &&
                        d.getMonth() === date.getMonth() &&
                        d.getFullYear() === date.getFullYear()
                } ))

                leadListWalkInPE.forEach((ele , index) => {

                    const cellLeadSourcePE = sheet.getCell(index + indexList +1, 20);
                    const cellLeadSourceValuePE = sheet.getCell(index + indexList +1, 21);
                    cellLeadSourcePE.value = ele.Leadsource || 'VUOTO'
                    cellLeadSourceValuePE.value = ele.quantity
                })
                indexList = indexList + leadListWalkInPE.length + 5

                const leadListVariousTitlePE = sheet.getCell(indexList, 20);
                leadListVariousTitlePE.value="Leadsource Varie"
                leadListVariousTitlePE.backgroundColor = {
                    "red": 80,
                    "green": 80,
                    "blue": 80,
                    "alpha": 1
                }
                const leadListCheckUpPE =  leadListPE.filter(ele => ele.tags.find(tag =>{
                    const d = calcMonthPrev(lastMonth)
                    const date = new Date(tag.date_applied)
                    return [674,1094].includes(tag.tag.id) &&
                        d.getMonth() === date.getMonth() &&
                        d.getFullYear() === date.getFullYear()


                }))
                const leadListVariousPE = leadListPE.filter(ele => !leadListCheckUpPE.find(el => el.LeadSourceId === ele.LeadSourceId) &&
                    !leadListWalkInPE.find(el => el.LeadSourceId === ele.LeadSourceId) &&
                    !loadListAdsPE.find(el => el.LeadSourceId === ele.LeadSourceId)    )
                leadListVariousPE.forEach((ele , index) => {

                    const cellLeadSourcePE = sheet.getCell(index + indexList +1, 20);
                    const cellLeadSourceValuePE = sheet.getCell(index + indexList +1, 21);
                    cellLeadSourcePE.value = ele.Leadsource || 'VUOTO'
                    cellLeadSourceValuePE.value = ele.quantity
                })

                indexList = indexList + leadListVariousPE.length + 5
                const cellLeadSourceTitlePE = sheet.getCell(indexList, 20);
                cellLeadSourceTitlePE.value="Leadsource Checkup"
                cellLeadSourceTitlePE.backgroundColor = {
                    "red": 80,
                    "green": 80,
                    "blue": 80,
                    "alpha": 1
                }


                leadListCheckUpPE.forEach((ele , index) => {

                    const cellLeadSourcePE = sheet.getCell(index + indexList +1, 20);
                    const cellLeadSourceValuePE = sheet.getCell(index + indexList +1, 21);
                    cellLeadSourcePE.value = ele.Leadsource || 'VUOTO'
                    cellLeadSourceValuePE.value = ele.quantity

                })


                uniqueSS.forEach((ele, index) => {
                    const uniqueSSCell = sheet.getCell(index+1, 22);
                    uniqueSSCell.value = ele.email
                })
                uniquePE.forEach((ele, index) => {
                    const uniquePECell = sheet.getCell(index+1, 23);
                    uniquePECell.value = ele.email
                })
                const today = process.env.DATAESTRAZIONE ? new Date(process.env.DATAESTRAZIONE) : new Date();
                const d = calcMonthPrev(today)
                const callbackSS = (list) => {
                    const newList = [...uniqueSS.map(ele => ele.email), ...list].reduce((acc, ele) => {
                        return [...acc, ...(acc.includes(ele) ? [] : [ele])]
                    }, [])

                const cellUniqueSS = sheet.getCell(3, 15);
                cellUniqueSS.value=newList.length
                    const callbackPE = (list) => {
                        const newList = [...uniquePE.map(ele => ele.email), ...list].reduce((acc, ele) => {
                            return [...acc, ...(acc.includes(ele) ? [] : [ele])]
                        }, [])
                        const cellUniquePE = sheet.getCell(4, 15);
                        cellUniquePE.value=newList.length
                        // in mancanza dei 12 mesi commentare sotto e decommentare qui
                        //sheet.saveUpdatedCells().then(res =>{
                        //    console.log("FINITO")
                       //     dataName.createArraySync()
                       // });


                        const callback12MSS = (list12MSS) => {
                            const newList = [...uniqueSS.map(ele => ele.email), ...list12MSS].reduce((acc, ele) => {
                                return [...acc, ...(acc.includes(ele) ? [] : [ele])]
                            }, [])
                            const cellUniqueSS = sheet.getCell(5, 15);
                            cellUniqueSS.value=newList.length
                            const callback12MPE = (list12MPE) => {

                                const newList = [...uniquePE.map(ele => ele.email), ...list12MPE].reduce((acc, ele) => {
                                    return [...acc, ...(acc.includes(ele) ? [] : [ele])]
                                }, [])

                                const cellUniquePE = sheet.getCell(6, 15);
                                cellUniquePE.value=newList.length
                                sheet.saveUpdatedCells().then(res =>{
                                    console.log("FINITO")
                                    dataName.createArraySync()
                                });

                            }       // qui faccio l'annuale per pescara
                            calcQuarterUnique(d,11, callback12MPE, 23)
                        }

                      //  qui faccio l'annuale per san salvo
                       calcQuarterUnique(d,11, callback12MSS, 22)
                    }

                    // qui faccio il trimestre per il pescara
                   calcQuarterUnique(d,2, callbackPE, 23)

                }
                sheet.saveUpdatedCells().then(res =>{
                    console.log("SALVATAGGIO 1")

                    // qui faccio il trimestre per il san salvo
                   calcQuarterUnique(d,2, callbackSS, 22)


                });



            }
            getContactLeadsource(uniquePE, call)
        }
        // funzione callback
        getContactLeadsource(uniqueSS, callback)


    })
}


data.productSheet = (products) =>{
    doc.useServiceAccountAuth({
        client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
        private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),
    }).then(res =>{
        doc.loadInfo().then(res =>{

            const sheetCount= doc.sheetCount
            const sheet = doc.sheetsByIndex[sheetCount - 1]
            const sheetTitle = sheet.title;
            const lastMonth = process.env.DATAESTRAZIONE ? new Date(process.env.DATAESTRAZIONE) : new Date();
            const newSheet = calcTitle(lastMonth)
            const callback = () => {
                if (sheetTitle.toLowerCase() === newSheet ){

                    console.log('il mese esiste')
                    orderWrite(sheet, products, lastMonth)
                }
                else {
                    doc.addSheet({'title':newSheet }).then(res => {
                       orderWrite(res, products, lastMonth)
                    })

                }
            }
            getProductInfo(sheet, callback)
            console.log('Prodotti Salvati')


        })
    })
}
module.exports = data;
