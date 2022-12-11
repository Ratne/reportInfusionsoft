const {calcMonthPrev} = require("../utils/dateUtils");
const data ={};

data.getLeadList = (category) => category.contacts.reduce((acc, ele) => {
    const data = acc.find(el => el.LeadSourceId === ele.LeadSourceId);

    return [...acc.filter(e => e.LeadSourceId !== ele.LeadSourceId), data ? {...data, quantity: data.quantity + 1} : {
            LeadSourceId : ele.LeadSourceId,
            Leadsource: ele.Leadsource,
            quantity: 1,
            tags: ele.tags
         }]
    }, []);

data.getLeadListAds = (leadList) => leadList.filter(ele => ele.Leadsource?.startsWith('[FB]') || ele.Leadsource?.startsWith('[GOO]') )
data.getLeadListWalkIn = (leadList, lastMonth) => leadList.filter(ele => ele.tags.find(tag => {

    const d = calcMonthPrev(lastMonth)
    const date = new Date(tag.date_applied)

    return [409].includes(tag.tag.id) &&
        d.getMonth() === date.getMonth() &&
        d.getFullYear() === date.getFullYear()
} ))

data.getLeadListCheckUp = (leadList, lastMonth) => leadList.filter(ele => ele.tags.find(tag => {
    const d = calcMonthPrev(lastMonth)
    const date = new Date(tag.date_applied)
    return [674,1094].includes(tag.tag.id) &&
        d.getMonth() === date.getMonth() &&
        d.getFullYear() === date.getFullYear()
} ))

data.getLeadListVarious = (leadList, leadListCheckUp, leadListWalkIn, leadListAds) => leadList.filter(ele => !leadListCheckUp.find(el => el.LeadSourceId === ele.LeadSourceId) &&
    !leadListWalkIn.find(el => el.LeadSourceId === ele.LeadSourceId) &&
    !leadListAds.find(el => el.LeadSourceId === ele.LeadSourceId)    );


module.exports = data;
