export const generateTempId = () => generateDateNow() * Math.round(Math.random() * 100)

export const generateDateNow = () => Date.now();
