const axios = require('axios');
const data ={};
const myCache = require("../myCache");



data.getTags = () => new Promise((resolve) => {
    axios.get(`https://api.infusionsoft.com/crm/rest/tags
` ,{
        headers: {
            Accept: 'application/json, */*',
            Authorization: 'Bearer '+ myCache.get('tokens').accessToken
        }}).then(res =>{
            resolve(res.data)
    }).catch(err =>{
        console.log('errore',err)
    })
})

data.getContacts = () => new Promise((resolve) => {
    axios.get(`https://api.infusionsoft.com/crm/rest/v1/contacts
` ,{
        headers: {
            Accept: 'application/json, */*',
            Authorization: 'Bearer '+ myCache.get('tokens').accessToken
        }}).then(res =>{
        resolve(res.data)
    }).catch(err =>{
        console.log('errore',err)
    })
})

data.getContTags = (idTag) => new Promise((resolve) => {
    axios.get(`https://api.infusionsoft.com/crm/rest/tags/${idTag}/contacts` ,{
        headers: {
            Accept: 'application/json, */*',
            Authorization: 'Bearer '+ myCache.get('tokens').accessToken
        }}).then(res =>{
        resolve(res.data)
    }).catch(err =>{
        console.log('errore',err)
    })
})

data.deleteTagsForContacts = (idTag, ids = []) => new Promise((resolve) => {
    axios.delete(`https://api.infusionsoft.com/crm/rest/tags/${idTag}/contacts?ids=${ids.join(',')}` ,{
        headers: {
            Accept: 'application/json, */*',
            Authorization: 'Bearer '+ myCache.get('tokens').accessToken
        }}).then(res => {
        resolve(res)
    }).catch(err =>{
        console.log('errore',err)
    })
})

module.exports = data;
