const axios = require('axios');
const data ={};
const myCache = require("../myCache");
const getContactCall = (callback, list = [], url = "https://api.infusionsoft.com/crm/rest/v1/contacts") => {
    axios.get(url ,{
        headers: {
            Accept: 'application/json, */*',
            Authorization: 'Bearer '+ myCache.get('tokens').accessToken
        }}).then(res =>{
            console.log(res.data.next)
            if(res.data.contacts.length){
                getContactCall(callback, [...list, ...res.data.contacts], res.data.next)
            }else {
                console.log('non trovato')
                callback([...list, ...res.data.contacts]);
            }
    }).catch(err =>{
        console.log('errore',err)
    })
}

data.getContacts = () => new Promise((resolve) => {
    getContactCall(resolve)
})

data.getContactOffice = (contactId) => new Promise((resolve) => {
    axios.get(`https://api.infusionsoft.com/crm/rest/contacts/${contactId}?optional_properties=custom_fields&access_token=${myCache.get('tokens').accessToken}` ,{
        headers: {
            Accept: 'application/json, */*',
        }}).then(res => {
            resolve(res.data)
        }).catch(err =>{
            console.log('errore',err)
        })
})



module.exports = data;
