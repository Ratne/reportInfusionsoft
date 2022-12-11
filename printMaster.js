const {createDoc, getColumnEmpty, addColumnsIfFinish, getSheetIfExistOrCreate, getInd} = require("./utils/excelUtils");
const {getTagListPromise} = require("./masterReport");
const {calcMonthPrev, calcTitle} = require("./utils/dateUtils");
const data = {};
const printSheetExpense = (sheetExpense, titleDate, lastMonth, category) => {
    return new Promise((resolve) => {
        sheetExpense.loadCells('').then(res => {
            getColumnEmpty(sheetExpense, 1).then(columnId => {
                let columnIdEmpty = columnId;
                let cell2 = sheetExpense.getCell(0, columnIdEmpty-1);
                if (cell2.value === titleDate){
                    columnIdEmpty = columnIdEmpty - 1
                }
                const cellMonth = sheetExpense.getCell(0,columnIdEmpty);
                cellMonth.value= titleDate
                sheetExpense.saveUpdatedCells().then(res =>{
                    console.log('salvataggio foglio spese')
                    resolve();
                })
            })
        })
    })
}

const printSheetTags = (sheetTags, titleDate, lastMonth, category) => {
    return new Promise((resolve) => {
        sheetTags.loadCells('').then(res =>{
            const idCell = sheetTags.getCell(0, 0)
            idCell.value = 'ID TAG';
            const nameCell = sheetTags.getCell(0, 1)
            nameCell.value = 'NOME TAG';
            sheetTags.saveUpdatedCells().then(res =>{
                getTagListPromise(sheetTags).then(tags => {
                    category.masterTags = tags;
                    getColumnEmpty(sheetTags, 1).then(columnId => {
                        let columnIdEmpty = columnId;
                        addColumnsIfFinish(columnIdEmpty, sheetTags);
                        let cell2 = sheetTags.getCell(0, columnIdEmpty-1);
                        if (cell2.value === titleDate){
                            columnIdEmpty = columnIdEmpty - 1
                        }
                        const cellMonth = sheetTags.getCell(0,columnIdEmpty);
                        cellMonth.value= titleDate
                        const tagIssetInDocument = category.tags.filter(ele => tags.find(el => el.id === ele.id));
                        const tagToAddDocument = category.tags.filter(ele => !tags.find(el => el.id === ele.id));
                        let indexNewValue = tags.length +1;
                        tagIssetInDocument.forEach(tag => {
                            const rowInfo = tags.find(el => el.id === tag.id)
                            const cell = sheetTags.getCell(rowInfo.rowIndex, columnIdEmpty);
                            cell.value=`=IMPORTRANGE("${category.sheetTagId}";"${calcTitle(lastMonth)}!F${tag.index+2}")`;
                        });
                        tagToAddDocument.forEach(tag => {
                            const idCell = sheetTags.getCell(indexNewValue, 0)
                            idCell.value = tag.id;
                            const nameCell = sheetTags.getCell(indexNewValue, 1)
                            nameCell.value = tag.name;
                            const cell = sheetTags.getCell(indexNewValue, columnIdEmpty);
                            cell.value=`=IMPORTRANGE("${category.sheetTagId}";"${calcTitle(lastMonth)}!F${tag.index+2}")`;
                            indexNewValue = indexNewValue+1;
                        })
                        console.log('before save master');
                        sheetTags.saveUpdatedCells().then(res =>{
                            console.log("SALVATAGGIO tag master")
                            resolve();
                        });
                    })
                })
            });

        })
    })
}

const copyColumn = (sheetTags, sheetMaster, callback,  rowIndex = 0, columnIndex = 0) => {
    const idCellControl = sheetTags.getCell(0, columnIndex);
    const idCell = sheetTags.getCell(rowIndex, columnIndex);
    if(idCellControl.value !== null){
        const idCellMaster = sheetMaster.getCell(rowIndex, columnIndex);
        idCellMaster.value = idCell.formula || idCell.value;
        copyColumn(sheetTags, sheetMaster, callback, rowIndex, columnIndex + 1)
    } else {
        callback();
    }

}

const copyColumnsRow = (sheetTags, sheetMaster, rowIndex) => {
    return new Promise((resolve) => {
        copyColumn(sheetTags, sheetMaster, resolve, rowIndex);
    })
}

const copyRow = (sheetTags, sheetMaster, callback,  rowIndex = 0) => {
    const idCell = sheetTags.getCell(rowIndex, 0);
    if(idCell.value){
        copyColumnsRow(sheetTags, sheetMaster, rowIndex).then(res => {
            copyRow(sheetTags, sheetMaster, callback, rowIndex +1)
        })
    } else {
        callback(rowIndex);
    }

}

const copyRows = (sheetTags, sheetMaster) => {
    return new Promise((resolve) => {
        copyRow(sheetTags, sheetMaster, resolve)
    })
}


const copyColumnLead = (category, sheetTags, sheetLead, callback, columnIndex = 2) => {
    const idCell = sheetTags.getCell(0, columnIndex);
    if(idCell.value !== null){
        const idCellLeadTitle = sheetLead.getCell(0, (columnIndex-2)*2);
        idCellLeadTitle.value = idCell.value;
        const importSheet = idCell.value.split('-').join('_');
        const idCellLead = sheetLead.getCell(1, (columnIndex-2)*2);
        idCellLead.value =`=IMPORTRANGE("${category.sheetId}";"${importSheet}!S:T")`;
        copyColumnLead(category, sheetTags, sheetLead, callback, columnIndex + 1)
    } else {
        callback();
    }

}
const generateLeadsForMonth = (category, sheetTags, sheetLead) => {
    return new Promise((resolve) => {
        copyColumnLead(category, sheetTags, sheetLead, resolve)
    })
}

const fixTitle = (sheetMaster, category, indexRow) => {
    const cellTitleEarn =  sheetMaster.getCell(indexRow + 1, 1);
    cellTitleEarn.value = 'Totale Guadagno'
    const cellTitleExpenseTag =  sheetMaster.getCell(indexRow + 2, 1);
    cellTitleExpenseTag.value = 'Tot Spese Tag'
}

const copyExpenseColumn = (sheetExpense, sheetMaster, rowIndexMaster, callback,  rowIndex = 0, columnIndex = 0) => {
    const idCell = sheetExpense.getCell(0, columnIndex);
    if(idCell.value !== null){
        const idCellMaster = sheetMaster.getCell(rowIndexMaster, columnIndex+1);
        idCellMaster.value = `=INDIRECT(ADDRESS(${rowIndex+1};${columnIndex+1};;;"spese"))`;
        idCellMaster.numberFormat = {
            type: 'CURRENCY'
        }
        copyExpenseColumn(sheetExpense, sheetMaster, rowIndexMaster, callback, rowIndex, columnIndex +1)
    } else {
        callback();
    }

}

const copyColumnsExpenseRow = (sheetExpense, sheetMaster, rowIndexMaster, rowIndex) => {
    return new Promise((resolve) => {
        copyExpenseColumn(sheetExpense, sheetMaster, rowIndexMaster, resolve, rowIndex);
    })
}
const copyExpenseRow = (sheetExpense, sheetMaster, rowIndexMaster, callback,  rowIndex = 0) => {
    const idCell = sheetExpense.getCell(rowIndex, 0);
    if(idCell.value !== null){
        copyColumnsExpenseRow(sheetExpense, sheetMaster, rowIndexMaster, rowIndex).then(res => {
            copyExpenseRow(sheetExpense, sheetMaster, rowIndexMaster+1, callback, rowIndex +1)
        })
    } else {
        callback({
            rowIndexMaster
        });
    }

}

const copyCellExpenseSheet = (sheetExpense, sheetMaster, rowIndexMaster) => {
    return new Promise((resolve) => {
        copyExpenseRow(sheetExpense, sheetMaster, rowIndexMaster, resolve, 1)
    })
}

const fixTitleAfterExpense = (sheetMaster, category, indexRow) => {
    const titles = ['Totale Spese', 'Tot Guadagno senza spese', 'Totale vendite', 'ROI', 'Numero clienti', 'Life Time', 'Numero clienti 3 mesi', 'Totale vendite 3mesi', 'Totale margine 3 mesi', 'Numero clienti 12 mesi', 'Totale vendite 12mesi', 'Totale margine 12 mesi', 'totale spese 3 mesi', 'totale spese 12 mesi', 'roi 3m', 'roi 12m']
    titles.forEach((value, index) => {
        const totExpanse =  sheetMaster.getCell(indexRow + index, 1);
        totExpanse.value = value
    })
}


const sumPrevMonth = (columnId, rowIndex, num) => {
    return Array(columnId).fill('').map((_, i) => i).reduce((arr, index) => {
        return [...arr, ...(index > 0 && arr.length < num)? [getInd(rowIndex, (columnId + 2 - (index)))] : []]
    }, []).join('+')
}


const printInfoByMonth = (callback, indexRow, rowIndexMaster, sheetMaster, category, columnId = 2) => {
    const listIndexExpense = Array(rowIndexMaster - (indexRow+2)).fill('').map((_, i) => i+(indexRow+3));
    const totalEarn3mValue = '=' + sumPrevMonth(columnId, indexRow + 2, 3);
    const totalEarn12mValue = '=' + sumPrevMonth(columnId, indexRow + 2, 12);
    const totalMargin3mValue = '=' + sumPrevMonth(columnId, rowIndexMaster + 2, 3);
    const totalMargin12mValue = '=' + sumPrevMonth(columnId, rowIndexMaster + 2, 12);
    const totalSales3mValue = '=' + sumPrevMonth(columnId, rowIndexMaster + 3, 3);
    const totalSales12mValue = '=' + sumPrevMonth(columnId, rowIndexMaster + 3, 12);
    const totalExpanse3mValue = '=' + sumPrevMonth(columnId, rowIndexMaster + 1, 3);
    const totalExpanse12mValue = '=' + sumPrevMonth(columnId, rowIndexMaster + 1, 12);
    const roi3m = `=${getInd(rowIndexMaster+9, columnId+1)}/${getInd(rowIndexMaster +13, columnId+1)}`;
    const roi12m = `=${getInd(rowIndexMaster+12, columnId+1)}/${getInd(rowIndexMaster +14, columnId+1)}`;
    const cellTitle = sheetMaster.getCell(0, columnId)
    if(cellTitle.value){
        const importSheet = cellTitle.value.split('-').join('_');
        const objRow = {
            earnValue : `=IMPORTRANGE("${category.sheetId}";"${importSheet}!M2")`,
            totExpensesTagValue : `=IMPORTRANGE("${category.sheetTagId}";"${importSheet}!I2")`,
            totExpensesValue : `=${listIndexExpense.length ? listIndexExpense.map(ele => getInd(ele, columnId+1)).join('+') : 0}`,
            earnWithoutExpenseValue : `=${getInd(indexRow + 2, columnId+1)}-${getInd(rowIndexMaster+1, columnId+1)}`,
            totSalesValue : `=IMPORTRANGE("${category.sheetId}";"${importSheet}!M3")`,
            roiValue : `=${getInd(rowIndexMaster + 2, columnId+1)}/${getInd(rowIndexMaster+1, columnId+1)}`,
            numCustomerValue : `=IMPORTRANGE("${category.sheetId}";"${importSheet}!P2")`,
            lifeTimeValue : `=${getInd(rowIndexMaster + 5, columnId+1)}/${getInd(rowIndexMaster+4, columnId+1)}`,
            numCustomer3mValue : `=IMPORTRANGE("${category.sheetId}";"${importSheet}!P4")`,
            totalSales3mValue,
            totalMargin3mValue,
            numCustomer12mValue : `=IMPORTRANGE("${category.sheetId}";"${importSheet}!P6")`,
            totalSales12mValue,
            totalMargin12mValue,
            totalExpanse3mValue,
            totalExpanse12mValue,
            roi3m,
            roi12m
        }
        const formatRow  = {
            earnValue: {
                type: "CURRENCY"
            },
            totExpensesTagValue: {
                type: "CURRENCY"
            },
            totExpensesValue: {
                type: "CURRENCY"
            },
            earnWithoutExpenseValue: {
                type: "CURRENCY"
            },
            totSalesValue: {
                type: "CURRENCY"
            },
            roiValue: {
                type: "PERCENT"
            },
            roi3m: {
                type: "PERCENT"
            },
            roi12m: {
                type: "PERCENT"
            }
        }
        Object.keys(objRow).forEach((prop, index) => {
            const cell =  sheetMaster.getCell((index < 2 ? (indexRow + 1): (rowIndexMaster -2)) + index, columnId);
            if(formatRow[prop]){
                cell.numberFormat = {
                    ...formatRow[prop]
                }
            } else {
                cell.numberFormat = {
                    type: 'NUMBER'
                }
            }
            cell.value= objRow[prop];
        })

        printInfoByMonth(callback, indexRow, rowIndexMaster, sheetMaster, category, columnId +1)
    }else {
        callback && callback();
    }

}

const printInfoByMonthPromise = (indexRow, rowIndexMaster, sheetMaster, category) => {
    return new Promise((resolve) => {
        printInfoByMonth(resolve, indexRow, rowIndexMaster, sheetMaster, category)
    })
}

const printCalcMaster = (sheetExpense, sheetMaster, category, indexRow, callbackEnd) => {
    fixTitle(sheetMaster, category, indexRow);
    copyCellExpenseSheet(sheetExpense, sheetMaster, indexRow+3).then(({rowIndexMaster}) => {
        fixTitleAfterExpense(sheetMaster, category, rowIndexMaster);
        printInfoByMonthPromise(indexRow, rowIndexMaster, sheetMaster, category).then(res => {
            sheetMaster.saveUpdatedCells().then(res =>{
                callbackEnd()
                console.log('save row calc')
            })
        })
    })

}
data.printMaster = (category, lastMonth, callbackEnd) =>{
    createDoc(category.sheetMasterId).then(doc => {
        const titleDate = `${calcTitle(lastMonth, '-')}`
        const arrayPromise = [getSheetIfExistOrCreate(doc, 'tags'), getSheetIfExistOrCreate(doc, 'spese'), getSheetIfExistOrCreate(doc, 'master'), getSheetIfExistOrCreate(doc, 'lead')];
        Promise.all(arrayPromise).then(([sheetTags, sheetExpense, sheetMaster, sheetLead]) => {
            Promise.all([printSheetExpense(sheetExpense, titleDate, lastMonth, category), printSheetTags(sheetTags, titleDate, lastMonth, category)]).then(res => {
                console.log('finito print calcoli e tag, ora inizia la copia')
                sheetLead.loadCells('').then(res => {
                    generateLeadsForMonth(category, sheetTags, sheetLead).then(r => {
                        sheetLead.saveUpdatedCells().then(res =>{
                            console.log('sheetLead copiato correttamente')
                        })
                    })
                })
                sheetMaster.loadCells('').then(res =>{
                    copyRows(sheetTags, sheetMaster).then(indexRow => {
                        sheetMaster.saveUpdatedCells().then(res =>{
                            printCalcMaster(sheetExpense, sheetMaster, category, indexRow, callbackEnd);
                            console.log("SALVATAGGIO finale sheet master")
                        });
                    })
                })
            })

        })
    })


}
module.exports = data;
