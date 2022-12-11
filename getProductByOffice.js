
const data ={};
const database = require('./database/getDatabase');
const {getProduct, getOrder, getContactLeadsourceCall} = require('./api/prouctAPI');
const {getContactOffice} = require('./api/contactAPI');

const getContactLeadsource = (list=[], callback, index = 0, newList = []) =>{
    const contact = list[index]
    getContactLeadsourceCall(contact).then(contactUpdate => {
        if ( index < list.length-1 ) {
            getContactLeadsource(list, callback, index + 1, [...newList, contactUpdate])
        }
        else {
            callback([...newList, contactUpdate])
        }
    })
}

const getContactLeadsourcePromise = (contacts) => new Promise((resolve) => {
    const callB = (updateContacts) => {
        resolve(updateContacts)
    };
    if(contacts.length){
        getContactLeadsource(contacts, callB);
    } else {
        resolve(contacts)
    }
})
const contactByOffice = (tag, office) => {
    return tag.contacts.filter(contact => {
        const con = database.contactsInfo.contacts.find(ele => ele.email.toLocaleLowerCase().trim() === contact.email.toLocaleLowerCase().trim())
        return con && con.office === office ;
    })
}
data.dataCategory = (office, callbackInit, indexCategory = 0) => {
    if (office.categories.length > indexCategory){
        const category = office.categories[indexCategory];

        getProduct(category.ProductCategoryId).then(res => {
            const callback = d => {
                category.products = d.map(ele => ({...ele, orders: ele.orders.filter(ord => ord.office === office.name)})).filter(ele => ele.orders.length);
                category.orders = category.products.reduce((acc, ele) => [...acc, ...ele.orders], []);
                category.contacts = category.orders.reduce((acc, ele) => {
                    return [...acc, ...(acc.find(e => e.email === ele.contact.email)? [] : [ele.contact])]
                }, [])
                getContactLeadsourcePromise(category.contacts).then(res => {
                    category.contacts = res;
                    category.tags = database.tags.map(tag => {
                        const contacts = contactByOffice(tag, office.name)
                        return ({...tag, contacts: contacts, count: contacts.length})
                    })
                    data.dataCategory(office,callbackInit, indexCategory+1);
                })
            }
            console.log('processo la sede di ' + office.name);
            getOrderByProduct(res.map(ele => ele.ProductId),callback);
        })
    } else {
        console.log('terminate le chiamate per ' + office.name)
        callbackInit && callbackInit()
    }
}
data.getOfficeData = (office) => new Promise((resolve) => {
    const callB = () => {
        resolve()
    };
    data.dataCategory(office, callB);
})



const getOrderByProduct=(products =[],callbackInit, index = 0, results = []) => {
    const today = process.env.DATAESTRAZIONE ? new Date(process.env.DATAESTRAZIONE) : new Date();
    const startDate = new Date(today.getFullYear(), today.getMonth()-1, 1).toISOString()
    const endDate = new Date(today.getFullYear(), today.getMonth(), 0).toISOString()
    console.log('i prodotti sono: ' + products.length);
    if (products.length>index){
        getOrder(startDate, endDate, products, index).then(res => {
            const callback = (listOrders) => getOrderByProduct(products, callbackInit, index+1, [...results, {productId: products[index], orders : listOrders}])
            getCustomFields(res.data.orders,products[index], callback)
            console.log('prodotto con index ' + index + ' mancano ' + (products.length - index) + 'prodotti');
        })
    }
    else {
        console.log('terminate chiamate una sede')
        callbackInit && callbackInit(results)

    }
};

const getCustomFields = (orders = [], productId, callback, index = 0, finalOrders=[]) => {
    if (orders.length > index){
        const ele = orders[index];
        getContactOffice(ele.contact.id).then(res => {
            getCustomFields(orders,productId,callback, index+1, [...finalOrders, {
                contact: ele.contact,
                item: ele.order_items.filter(el => el.product).find(el => el.product.id === (+productId)),
                office: res['custom_fields'].find(ele => ele.id === +process.env.CUSTOMSEAT)?.content
            }])
        })
    }
    else {
        callback(finalOrders)
    }
}


module.exports = data;
