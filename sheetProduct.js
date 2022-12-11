const {createDoc} = require("./utils/excelUtils");
const data = {}
const {calcMonthPrev, calcTitle} = require("./utils/dateUtils");
const {printTag} = require("./printTag");
const {getLeadList, getLeadListAds, getLeadListWalkIn, getLeadListCheckUp, getLeadListVarious} = require('./services/leadService')
const {printLeadList, printLeadListTitle} = require('./services/leadSheetService')
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
       console.log(list, 'listValue')
       callback(list)
   }
}

const getProductInfo = (sheet,callback) => {
     sheet.loadCells('J:X').then(res => {
        listValue(sheet, callback)
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



const calcQuarterUnique = (doc, date, number, clk, columnIndex, list=[])  => {
    if (number === 0) {
        clk(list)
    }
    else {

        const prevMonth = calcMonthPrev(date);
        if(doc.sheetsByTitle[calcTitle(date)]){
            const prevSheet = doc.sheetsByTitle[calcTitle(date)]
            prevSheet.loadCells("W:X").then(res => {
                const internalCallback = (l) => {
                    calcQuarterUnique(doc, prevMonth, number -1, clk, columnIndex,[...list, ...l]  )
                }
                listBlankValue(prevSheet,internalCallback, 1, columnIndex)

            })
        } else {
            clk(list)
        }

    }

}

const calcMonthUniquePromise = (doc, contacts, number) => {
    return new Promise((resolve) => {
        const callback = (list) => {
            const newList = [...contacts.map(ele => ele.email), ...list].reduce((acc, ele) => {
                return [...acc, ...(acc.includes(ele) ? [] : [ele])]
            }, [])
            resolve(newList);
        }
        const today = process.env.DATAESTRAZIONE ? new Date(process.env.DATAESTRAZIONE) : new Date();
        const d = calcMonthPrev(today)
        calcQuarterUnique(doc, d,number, callback, 22)
    })
}

const calcQuarterUniquePromise = (doc, contacts) => {
    return calcMonthUniquePromise(doc, contacts, 2);
}
const calcYearUniquePromise = (doc, contacts) => {
    return calcMonthUniquePromise(doc, contacts, 11);
}

const calcYearAndQuarterUniquePromise = (doc, contacts) => {
    return new Promise((resolve) => {
        Promise.all([calcQuarterUniquePromise(doc, contacts), calcYearUniquePromise(doc, contacts)]).then(([quarter, year]) => {
            resolve({
                quarter: {
                    contacts: quarter
                },
                year: {
                    contacts: year
                },
            })
        })
    })
    return calcMonthUniquePromise(doc, contacts, 11);
}



const printLead = (category, sheet, lastMonth) => {
    const leadList = getLeadList(category);
    const leadListAds = getLeadListAds(leadList);
    const leadListWalkIn = getLeadListWalkIn(leadList, lastMonth);
    const leadListCheckUp = getLeadListCheckUp(leadList, lastMonth);
    const leadListVarious = getLeadListVarious(leadList, leadListCheckUp, leadListWalkIn, leadListAds);
    let indexList = printLeadList(leadList, sheet);
    console.log('fine print lead list', indexList);
    indexList = printLeadListTitle(sheet, indexList, 'Leadsource Inserzioni')
    indexList = printLeadList(leadListAds, sheet, indexList)
    console.log('printLeadList', indexList);
    indexList = printLeadListTitle(sheet, indexList, 'Leadsource Walk-in')
    indexList = printLeadList(leadListWalkIn, sheet, indexList)
    indexList = printLeadListTitle(sheet, indexList, 'Leadsource Varie')
    indexList = printLeadList(leadListVarious, sheet, indexList)
    indexList = printLeadListTitle(sheet, indexList, 'Leadsource Checkup')
    indexList = printLeadList(leadListCheckUp, sheet, indexList)
}

const printProductsInfo = (category, sheet) => {
    let base = 1;
    category.products.forEach((prod, prodIndex) => {
        const cellIdProduct = sheet.getCell(prodIndex + 1, 9);
        const cellPriceProduct = sheet.getCell(prodIndex + 1, 10);
        const cellTaxProduct = sheet.getCell(prodIndex + 1, 11);
        cellIdProduct.value= prod.productId;
        const singleProductInfo = category.productInfo.find(ele => +prod.productId === +ele.productId)
        cellPriceProduct.value= singleProductInfo ? singleProductInfo.productPrice : ''
        cellTaxProduct.value= singleProductInfo ? singleProductInfo.productTax : ''

        prod.orders.forEach((order, index) => {
            const cellProductName = sheet.getCell(index+base, 0);
            const cellProductId = sheet.getCell(index+base, 1);
            const cellProductPrice = sheet.getCell(index+base, 2);
            const cellContactName = sheet.getCell(index+base, 3);
            const cellContactEmail = sheet.getCell(index+base, 4);
            const cellContactOffice = sheet.getCell(index+base, 5);
            const cellEarn = sheet.getCell(index+base, 6);
            const cellSell = sheet.getCell(index+base, 7);

            cellProductName.value = order.item.name;
            cellProductId.value = order.item.id;
            cellProductPrice.value = order.item.price;
            cellContactName.value = order.contact.first_name
            cellContactEmail.value= order.contact.email
            cellContactOffice.value= order.office
            cellEarn.value= `=(H${index +1 +base}-K${prodIndex + 2})`
            cellSell.value= `=(C${index +1 +base}*(100-L${prodIndex + 2})/100)`
            if (index === (prod.orders.length-1)){
                base = base + (index + 1)
            }
        });
    })
    return base;
}

const orderWrite = (doc, sheet, category, lastMonth, callbackEnd) => {
    calcYearAndQuarterUniquePromise(doc, category.contacts).then(res => {
        category.data = res;
        sheet.loadCells('A:X').then(res =>{
            category.contacts.forEach((ele, index) => {
                const uniqueCell = sheet.getCell(index+1, 22);
                uniqueCell.value = ele.email
            })
            const cellUniqueQuarter = sheet.getCell(3, 15);
            cellUniqueQuarter.value=category.data.quarter.contacts.length
            const cellUniqueYear = sheet.getCell(5, 15);
            cellUniqueYear.value=category.data.year.contacts.length
            titles.forEach((ele, index) => {
                const cell = sheet.getCell(0, index);
                cell.value = ele
            });
            const base = printProductsInfo(category, sheet);
            const totEarnCell = sheet.getCell(1, 12);
            const totWithoutExpenseCell = sheet.getCell(2, 12);
            const totExpenseCell = sheet.getCell(1, 17);
            const totUniqueContactCell = sheet.getCell(1, 15);
            const arrayByBase = Array(base).fill('').map((_, i) => i+1);
            totEarnCell.value= `=(${arrayByBase.map(ele => 'G'+(ele+1)).join('+')})`
            totWithoutExpenseCell.value= `=(${arrayByBase.map(ele => 'H'+(ele+1)).join('+')})`
            totExpenseCell.value= `=(${arrayByBase.map(ele => 'C'+(ele+1)).join('+')})`
            totUniqueContactCell.value=`${category.contacts.length}`
            printLead(category, sheet, lastMonth);
            sheet.saveUpdatedCells().then(res =>{
                printTag(category, lastMonth, callbackEnd);
                console.log("SALVATAGGIO excel")
            });
        })
    })

}


data.productSheet = (category, callbackEnd) =>{
    createDoc(category.sheetId).then(doc => {
            const sheetCount= doc.sheetCount
            const sheet = doc.sheetsByIndex[sheetCount - 1]
            const sheetTitle = sheet.title;
            const lastMonth = process.env.DATAESTRAZIONE ? new Date(process.env.DATAESTRAZIONE) : new Date();
            const newSheet = calcTitle(lastMonth)
            const callback = (list) => {
                category.productInfo = list;
                if (sheetTitle.toLowerCase() === newSheet ){

                    console.log('il mese esiste')
                    orderWrite(doc, sheet, category, lastMonth, callbackEnd)
                }
                else {
                    doc.addSheet({'title':newSheet }).then(res => {
                        orderWrite(doc, res, category, lastMonth, callbackEnd)
                    })

                }
            }
            const firstSheet = doc.sheetsByIndex[0]
            getProductInfo(firstSheet, callback)
    })
}
module.exports = data;
