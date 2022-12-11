const {getContactOffice, getContacts} = require('../api/contactAPI');
const {createDoc, getSheetIfExistOrCreate} = require("../utils/excelUtils");
const {calcMonthPrev, calcTitle} = require("../utils/dateUtils")
const database = require("../database/getDatabase");
const data ={};

const addOfficeToContact = (oldContacts = [], contacts, callback, index=0, contactsUpdate = []) => {
    const contact = contacts[index]
    if(contact){
        console.log('prelevo office da contatto con id ' +contact.id)
        const existContact = oldContacts.find(ele => ele.id === contact.id);
        if(existContact){
            addOfficeToContact(oldContacts, contacts, callback, index+1, [...contactsUpdate, existContact])
        } else {
            getContactOffice(contact.id).then(res => {
                addOfficeToContact(oldContacts, contacts, callback, index+1, [...contactsUpdate, {...contact, email: getEmailParameter(contact), office: res['custom_fields'].find(ele => ele.id === +process.env.CUSTOMSEAT)?.content}])
            })
        }
    } else {
        callback && callback(contactsUpdate.filter(ele => ele.office))
    }
}

data.getContactByInfusionsoft = (oldContacts = []) => {
    return new Promise((resolve) => {
        getContacts().then(contacts => {
            console.log('contatti totali' +contacts.length)
            addOfficeToContact(oldContacts, contacts, resolve)
        })
    })
}

const fields = [
    {label: 'ID', field: 'id'},
    {label: 'email', field: 'email'},
    {label: 'Nome', field: 'given_name'},
    {label: 'Cognome', field: 'family_name'},
    {label: 'Sede', field: 'office'}
]

const getEmailParameter = (contact) => {
        return (contact.email_addresses || []).map(ele => ele.email).join(',')
    }

const writeRowExcel = (contacts = [], sheet, callback) => {
    contacts.forEach((contact, rowIndex) => {
        fields.forEach((field, indexField) => {
            const cellValue = sheet.getCell(rowIndex +1, indexField);
            cellValue.value =  contact[field.field]
        })
    })
    sheet.saveUpdatedCells().then(res => {
        callback && callback(contacts)
    })

}

data.getSheetContacts = () => {
    return new Promise((resolve, reject) => {
        createDoc(database.contactsInfo.sheetId).then(doc => {
            getSheetIfExistOrCreate(doc, 'contatti').then(sheet => {
                sheet.loadCells('A:G').then(res => {
                    resolve(sheet)
                })
            })
        })
    })
}

const readExcelRow = (sheet, callback, rowIndex = 1, contacts = []) => {
    const cellFirstColumn = sheet.getCell(rowIndex, 0);
    if(cellFirstColumn.value){
        const contactData = fields.reduce((contact, field, columnIndex) => {
            const cellColumn = sheet.getCell(rowIndex, columnIndex);
            return {...contact, [field]: cellColumn.value}
        }, {})
        readExcelRow(sheet, callback, rowIndex+1, [...contacts, contactData])
    }else {
        callback && callback(contacts)
    }
}

data.readContactsExcel = (sheet) => {
    return new Promise((resolve, reject) => {
        readExcelRow(sheet, resolve)
    })
}



data.updateContactExcel = () => {
    return new Promise((resolve) => {
        console.log('inizio lettura foglio')
        data.getSheetContacts().then(sheet => {
            console.log('letto foglio');
            data.readContactsExcel(sheet).then(oldContacts => {
                data.getContactByInfusionsoft(oldContacts).then(contacts => {
                    fields.forEach((field, index) => {
                        const cellTitle = sheet.getCell(0, index);
                        cellTitle.value = field.label
                    })
                    writeRowExcel(contacts, sheet, resolve)
                })
            })
        })

    })
}


const writeRowTagsExcel = (tags = [], sheet, callback) => {
    let rowIndex = 0;
    tags.forEach((tag) => {
        rowIndex = rowIndex + 1
        const cellTagId = sheet.getCell(rowIndex, 0);
        cellTagId.value = tag.id;
        const cellTagName = sheet.getCell(rowIndex, 1);
        cellTagName.value = tag.name;
        tag.contacts.forEach((contact, indexContact) => {
            rowIndex = rowIndex+1;
            fields.forEach((field, indexField) => {
                const cellValue = sheet.getCell(rowIndex, indexField+2);
                cellValue.value =  contact[field.field]
            })
        })

    })
    sheet.saveUpdatedCells().then(res => {
        callback && callback()
    })

}

data.updateTagsExcel = () => {
    return new Promise((resolve) => {
        data.getSheetTags().then(sheet => {
            writeRowTagsExcel(database.tags, sheet, resolve)
        })

    })
}


data.getSheetTags = () => {
    return new Promise((resolve, reject) => {
        createDoc(database.contactsInfo.sheetId).then(doc => {
            getSheetIfExistOrCreate(doc, 'tags_'+ calcTitle(new Date())).then(sheet => {
                sheet.loadCells('A:G').then(res => {
                    resolve(sheet)
                })
            })
        })
    })
}

module.exports = data;
