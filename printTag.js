const {createDoc} = require("./utils/excelUtils");
const {calcTitle} = require("./utils/dateUtils");
const {printMaster} = require("./printMaster");
const data = {};
const titlePrint = [{label: 'tag name', column: 1}, {label: 'tag id', column: 2}, {label: 'conteggio', column: 5}]

const listValue = (sheet, callback, index = 0, list=[]) =>{
    const cellTagId = sheet.getCell(index, 2);
    const cellPriceTag = sheet.getCell(index, 3);

    if (cellTagId.value){
        listValue(sheet, callback , index+1, [...list, {
            id: cellTagId.value,
            price: cellPriceTag.value
        }])
    }
    else {
        callback(list)
    }
}

const getCostTagInfo = (sheet,callback) => {
    sheet.loadCells('B:F').then(res => {
        listValue(sheet, callback)
    })
}

const printTagInfo = (category, sheet, tag, index) => {
    const tagPrev = category.tagInfo.find(tagP => tagP.id === tag.id);
    tag.index = index;
    const tagNameCell = sheet.getCell(index+1, 1);
    tagNameCell.value = tag.name;
    const tagIdCell = sheet.getCell(index+1, 2);
    tagIdCell.value = tag.id;
    if(tagPrev){
        const tagPriceCell = sheet.getCell(index+1, 3);
        tagPriceCell.value = tagPrev.price;
    }
    const tagCountCell = sheet.getCell(index+1, 5);
    tagCountCell.value = tag.count;

    const totCountCell = sheet.getCell(index+1, 7);
    totCountCell.value = `=D${index+2}*F${index+2}`;
}

const printInfoTags = (category, sheet) => {
    category.tags.forEach((ele, index) => {
        printTagInfo(category, sheet, ele, index)
    })
}

const printTitles = (sheet) => {
    titlePrint.forEach(ele => {
        const titleCell = sheet.getCell(0, ele.column);
        titleCell.value = ele.label;
    })
}

const listTagPrint = (doc, sheet, category, lastMonth, callbackEnd) => {
    sheet.loadCells('A:L').then(res =>{
        printTitles(sheet);
        printInfoTags(category, sheet)
        const arrayByBase = Array(category.tags.length).fill('').map((_, i) => i+2);
        const totMonth = sheet.getCell(1, 8);
        totMonth.value= `=(${arrayByBase.map(ele => 'H'+(ele)).join('+')})`
        sheet.saveUpdatedCells().then(res =>{
            printMaster(category, lastMonth, callbackEnd)
            console.log("SALVATAGGIO tag print")
        });
    });
}

data.printTag = (category, lastMonth, callbackEnd) =>{
    createDoc(category.sheetTagId).then(doc => {
            const sheetCount= doc.sheetCount
            const sheet = doc.sheetsByIndex[sheetCount - 1]
            const sheetTitle = sheet.title;
            const newSheet = calcTitle(lastMonth);
            const callback = (list) => {
                category.tagInfo = list;
                if (sheetTitle.toLowerCase() === newSheet ){

                    console.log('il mese esiste')
                    listTagPrint(doc, sheet, category, lastMonth, callbackEnd)
                    // orderWrite(doc, sheet, category, lastMonth)
                }
                else {
                    doc.addSheet({'title':newSheet }).then(res => {
                        listTagPrint(doc, res, category, lastMonth, callbackEnd)
                    })

                }
            }

            getCostTagInfo(sheet, callback)
    })


}
module.exports = data;
