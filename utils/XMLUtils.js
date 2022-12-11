const createXml= (dataXml, methodName) => {

  return  `<?xml version='1.0' encoding='UTF-8'?>
            <methodCall>
              <methodName>${methodName || 'DataService.query'}</methodName>
              <params>
                <param>
                  <value><string>${process.env.CLIENTSECRET}</string></value>
                </param>
                ${dataXml}
              </params>
            </methodCall>`
}

const baseJsonXml = (json) =>{
    return json.methodResponse.params[0].param[0].value[0].array[0].data[0].value;
}


const formatJson = (json) =>{
    let data = baseJsonXml(json)
    return data.map(ele =>{
        return ele.struct[0].member.reduce((acc, ele) =>{
            acc[ele.name[0]]=ele.value[0]?.i4? ele.value[0].i4[0] : ele.value[0];
            return acc
        }, {})

    })

}
const formatJsonLeadsource = (json) =>{

    let data = json.methodResponse.params[0].param[0].value[0];

    return data.struct[0].member.reduce((acc, ele) =>{
        acc[ele.name[0]]=ele.value[0]?.i4? ele.value[0].i4[0] : ele.value[0];
        return acc
    }, {})

}

module.exports = {
    createXml,
    formatJson,
    baseJsonXml,
    formatJsonLeadsource
};
