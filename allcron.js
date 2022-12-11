const cron = require("node-cron");
const axios = require('axios');
const myCache = require('./myCache');
const nodeBase64 = require('nodejs-base64-converter');
// const dataName = require("./getData");
const dataProduct = require("./getProduct");
require('dotenv').config()

// refresh token @ 6H
cron.schedule("7 */6 * * *", function() {

       const refreshToken = myCache.get('tokens').refreshToken;
       let data = process.env.CLIENTID+':'+process.env.CLIENTSECRET;
       console.log(data);
       let base64data = nodeBase64.encode(data)
    console.log(base64data)
    const params = new URLSearchParams()
    params.append('grant_type', 'refresh_token')
    params.append('refresh_token', refreshToken)
       axios.post('https://api.infusionsoft.com/token',
           params,
           {headers: {
               'Content-Type': 'application/x-www-form-urlencoded',
               Authorization : 'Basic ' + base64data
                   }}).then(res =>{
           let tokens = {
               accessToken: res.data.access_token,
               refreshToken: res.data.refresh_token
           };

           myCache.set( "tokens", tokens, 0 );

       }).catch(err =>{
           console.log(err)
       })
});


// cron.schedule("0 0 2 * *", function() {
//     dataName.getSavedSearch()
// });
cron.schedule("0 0 3 * *", function() {
    dataProduct.getProduct()
});

