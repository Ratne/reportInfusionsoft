const data = {};


const listValue = (sheet, callback, index = 1, list=[]) =>{
    const cellTagId = sheet.getCell(index, 0);

    if (cellTagId.value){
        listValue(sheet, callback , index+1, [...list, {
            id: cellTagId.value,
            rowIndex: index
        }])
    }
    else {
        callback(list)
    }
}

const getTagList = (sheet,callback) => {
    sheet.loadCells('A:B').then(res => {
        listValue(sheet, callback)
    })
}

data.getTagListPromise = (sheet) => {
    return new Promise((resolve) => {
        const callB = (data) => resolve(data);
        getTagList(sheet, callB);
    })
}

module.exports = data;
