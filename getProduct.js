const data ={};
const database = require('./database/getDatabase');
const {getOfficeData} = require('./getProductByOffice')
const {productSheet} = require('./sheetProduct')
const {getTags, getContTags, deleteTagsForContacts} = require("./api/tagAPI");
const {updateContactExcel, updateTagsExcel} = require("./services/excelContacts");

data.initOffice = (offices, callbackInit, index = 0) => {
    if (offices.length > index){
        getOfficeData(offices[index]).then(res => {
            data.initOffice(offices,callbackInit, index+1);
        })
    } else {
        callbackInit && callbackInit()
    }
}
const createSheetForOfficePromise = (office) =>{
    return new Promise((resolve) => {
        const callback = () => {
            console.log('terminato excel save')
            resolve()
        }
        createSheetForOffice(office, callback)
    })
}

const deleteTagPromises = (tags, callback, index=0) => {
    const tag = tags[index];
    if(tag){
        deleteTagsForContacts(tag.id, tag.contacts.map(ele => ele.id)).then(res => {
            deleteTagPromises(tags, callback, index +1)
        })
    } else {
        callback && callback()
    }
}

const createSheetByOffices = (offices) => {
    const promises = offices.map(office => createSheetForOfficePromise(office));
    Promise.all(promises).then(res => {
        updateTagsExcel().then(res => {
            // process.env.ENVINRONMENT === 'PROD' && deleteTagPromises(database.tags.filter(ele => !ele.name.startsWith('212 - REPORT-TOT')), () => console.log('terminato delete tag'))

        })
    })
}
const createSheetForOfficeAndCategoryPromise = (cat) => {
    return new Promise((resolve) => {
        createSheetForOfficeAndCategory(cat, resolve)
    })

}
const createSheetForOffice = (office, callback) => {
    const promises = office.categories.map(cat => createSheetForOfficeAndCategoryPromise(cat))
    Promise.all(promises).then(res => {
        callback();
    })
}
const createSheetForOfficeAndCategory = (category, callback) => {
    productSheet(category, callback)
}

const calcTags = (tags, callback, index = 0, updateList = []) => {
    if(tags.length > index){
        const tag = tags[index];
        getContTags(tag.id).then(res => {
             calcTags(tags, callback, index +1, [...updateList, {...tag, contacts: res.contacts.map(ele => ele.contact)}])
        })
    } else {
        callback(updateList);
    }
}

data.createDatabaseOffice = () => {
    console.log('partenza script')
    const callBackTags = (list) => {
        database.tags = list;
        console.log('prelevati dati tags')
        const callback = () => {
            console.log('inizio creazione sedi')
            createSheetByOffices(database.offices);
        }
        data.initOffice(database.offices, callback)
    }
    updateContactExcel().then(res => {
        database.contactsInfo.contacts = res;
        getTags().then(res => {
            const tagsAll = res.tags.filter(ele => ele.name.startsWith('212 - REPORT'))
            const tagsFilter = tagsAll;
            const tagsTot = tagsAll.filter(ele => ele.name.startsWith('212 - REPORT-TOT'));
            calcTags(tagsFilter, callBackTags);
        })
    })
}



module.exports = data;
