const axios = require('axios');
const data ={};
const myCache = require("../myCache");
const {formatJson,formatJsonLeadsource} = require("../utils/XMLUtils");
const parseString = require('xml2js').parseString;


data.getProduct=(productCategoryId) => {
   return new Promise((resolve) => {
       const requestXml=`<?xml version='1.0' encoding='UTF-8'?>
<methodCall>
  <methodName>DataService.query</methodName>
  <params>
    <param>
      <value><string>${process.env.CLIENTSECRET}</string></value>
    </param>
    <param>
      <value><string>ProductCategoryAssign</string></value>
    </param>
    <param>
      <value><int>1000</int></value>
    </param>
    <param>
      <value><int>0</int></value>
    </param>
    <param>
      <value><struct>
        <member><name>ProductCategoryId</name>
          <value><string>${productCategoryId}</string></value>
        </member>
      </struct></value>
    </param>
    <param>
      <value><array>
        <data>
          <value><string>ProductId</string></value>
        </data>
      </array></value>
    </param>
  </params>
</methodCall>`
       axios.post(`https://api.infusionsoft.com/crm/xmlrpc/v1?access_token=${myCache.get('tokens').accessToken}`, requestXml ,{
           headers: {
               'Content-Type': 'text/xml'
           }
       }).then(res => {
           parseString(res.data, function (err, result) {
               const formatData = formatJson(result)
                resolve(formatData)
           });

       }).catch(err =>{
           console.log(err)
       })
   })
}

data.getOrder = (startDate, endDate, products, index) => new Promise((resolve) => {
    axios.get(`https://api.infusionsoft.com/crm/rest/orders?since=${startDate}&until=${endDate}&paid=true&product_id=${products[index]}&access_token=${myCache.get('tokens').accessToken}
` ,{
        headers: {
            Accept: 'application/json, */*',
        }}).then(res =>{
            resolve(res)
    }).catch(err =>{
        console.log('errore',err)
    })
})

data.getContactLeadsourceCall = (contact) => new Promise((resolve) => {
    const requestXml= `<?xml version='1.0' encoding='UTF-8'?>
                <methodCall>
                <methodName>ContactService.load</methodName>
                  <params>
                    <param>
                      <value><string>${process.env.CLIENTSECRET}</string></value>
                    </param>
                    <param>
                      <value><int>${contact.id}</int></value>
                    </param>
                    <param>
                      <value><array>
                        <data>
                          <value><string>Leadsource</string></value>
                          <value><string>LeadSourceId</string></value>
                        </data>
                      </array></value>
                    </param>
                  </params>
                </methodCall>`
    axios.post(`https://api.infusionsoft.com/crm/xmlrpc/v1?access_token=${myCache.get('tokens').accessToken}`, requestXml ,{
        headers: {
            'Content-Type': 'text/xml'
        }
    }).then(res => {

        parseString(res.data, function (err, result) {


            const data = formatJsonLeadsource(result)
            const contactUpdate = {...contact, ...data}
            // fare chiamata a prender ele tags
            axios.get(`https://api.infusionsoft.com/crm/rest/contacts/${contact.id}/tags`,
                {
                    headers: {
                        Accept: 'application/json, */*',
                        Authorization: 'Bearer '+myCache.get('tokens').accessToken
                    }})
                .then (res => {
                    contactUpdate.tags = res.data.tags
                    resolve(contactUpdate)
                })

        })
    })
})
module.exports = data;
