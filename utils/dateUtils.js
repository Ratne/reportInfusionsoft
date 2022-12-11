const data ={}

data.calcMonthPrev = (lastMonth) => {
    const date = new Date (lastMonth)
    return new Date (date.setMonth(date.getMonth() - 1));
}
data.calcTitle = (lastMonth, separator = '_') => {
    const d = data.calcMonthPrev(lastMonth)
    const today = new Date(d).toLocaleString('it-it',{month:'long', year:'numeric'})
    return today.split(' ').join(separator)
}



module.exports = data;
