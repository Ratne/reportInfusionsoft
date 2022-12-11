const data ={};
data.contactsInfo = {
    sheetId: "1pDM3PzMe0597nKzjxznr6A3Ol7RRpjsI7l3TZ0TpEGo"
}
if(process.env.ENVINRONMENT !== 'PROD'){
    //test
data.offices = [
    {
        name: 'PESCARA',
        categories: [{
            ProductCategoryId: 3,
            sheetId: '1t13jbvneZQl6CVb2yBoenUdOYa_rBFpwYEr0sT1ttZY',
            sheetTagId: '13LX5hy-GkohqjCBQRz3hNmGjxrVWikutrFU-jB6pYjI',
            sheetMasterId: '1qV6pKXhBtRroRjrlqxH4X-1Oog2ZX0J_GO9eIeF9W7Q',
            nameTag: '212 - REPORT',
            products: []
        }]
    },
    {
        name: 'SAN SALVO',
        categories: [{
            ProductCategoryId: 3,
            sheetId: '1-6WbHOHx9vJIbccJnKqWqvQvDKIH3cR8sej4rFnfxPQ',
            sheetTagId: '1gZb1YHf-jOGy1fBiJ2lMj3z3cIuD3QBh_WPoroY4fLc',
            sheetMasterId: '129KWUFu6zaHVD8DVD-Utz0PtysvAe_FK453fHV1GwNo',
            nameTag: '212 - REPORT',
            products: []
        },
            {
            ProductCategoryId: 3,
            sheetId: '1JFw626GMkKegKX6lkaVR1T6tTwbD0C5BztmhgHI9Izg',
            sheetTagId: '1mdXDtAegRatJvt_XlxNXdsRJejqBn31is8oAczx3P1Y',
            sheetMasterId: '1NJXoIJpvrUzZCF5T00pjw1xaY2_3CY0zsA7cW4S5FCY',
            nameTag: '212 - REPORT',
            products: []
        }]
    },
    ]
} else {
    data.offices = [{
        name: 'SAN SALVO',
        contacts: [],
        categories: [{
            ProductCategoryId: 3,
            sheetId: '1jS4-pog95Qf8yRYA_HKK_7tEx0ysZZz8ldlOroeHFOo',
            sheetTagId: '1pGdQ83GQ-MlZV1M1f-zki5hUcjSa8vKtmKL40jq69dI',
            sheetMasterId: '1cXVQqmBnEk42V9pQ-Occ7MEjLC3gW_x_9Nj9BhDaltY',
            nameTag: '212 - REPORT',
            products: []
        }]
    },
        {
            name: 'PESCARA',
            categories: [{
                ProductCategoryId: 3,
                sheetId: '1K4-KLUBPqJBZN4d003b-T0BSysZwWeRh5RfATsXOzJg',
                sheetTagId: '1vsUKHpefur7Upmkz-CHTNfIDw55BWt2ERZ6VMoWK4Hc',
                sheetMasterId: '1RE2Qpzo7noHBoeZuqHb9yO2ONypYP9qpdsFxJb-7hw4',
                nameTag: '212 - REPORT',
                products: []
            }]
        }]
}




data.tags = [];


module.exports = data;
