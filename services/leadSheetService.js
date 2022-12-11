const data ={};

data.printLeadList = (leadList = [], sheet, indexList = 0) => {
    console.log(indexList, 'enter index list');
    leadList.forEach((ele , index) => {
        const cellLeadSource = sheet.getCell(indexList +1, 18);
        const cellLeadSourceValue = sheet.getCell(indexList +1, 19);
        cellLeadSource.value = ele.Leadsource || 'VUOTO'
        cellLeadSourceValue.value = ele.quantity
        indexList = indexList + 1;
        console.log(indexList, 'foreach index list');
    })
    return indexList + 5;
}

data.printLeadListTitle = (sheet, indexList, title) => {
    const cellTitleLeadSource = sheet.getCell(indexList + 1, 18);
    cellTitleLeadSource.value=title
    cellTitleLeadSource.backgroundColor = {
        "red": 80,
        "green": 80,
        "blue": 80,
        "alpha": 1
    }
    return indexList + 1;
}

module.exports = data;
