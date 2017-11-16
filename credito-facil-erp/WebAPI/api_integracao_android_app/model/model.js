var Sequelize = require('sequelize');


exports.GetORM = function (banco, tabela, colunas){
    const sequelize = new Sequelize('[credito-facil-homologacao]', 'sa', '288744cla', {
        host: '187.114.153.212,81',
        dialect: 'mssql',
        dialectOptions: {
            appName: 'api_integracao_android_app'
        },  
        pool: {
            max: 300,
            min: 0,
            idle: 10000
        }
    });

    var obj = `{ `
    obj = obj + colunas.map((coluna)=>{
        var retorno = ``
        retorno = `"${coluna.nome}": { "type": "Sequelize.${coluna.tipo}", "primaryKey": ${coluna.pk} }`
        return retorno
    })
    obj = obj + ` }`

    const orm = sequelize.define(tabela, JSON.parse(obj), {timestamps: false, freezeTableName: true, operatorAliases: false});

    return orm;
}

exports.GetConection = function (){
    const sequelize = new Sequelize('[credito_facil]', 'sa', '288744cla', {
        host: 'credito-facil.cptqzj1ct8cm.sa-east-1.rds.amazonaws.com',
        // port: 81,
        dialect: 'mssql',
        dialectOptions: {
            appName: 'api_integracao_android_app'
        },  
        pool: {
            max: 1,
            min: 0,
            idle: 10000
        }
    });

    return sequelize
}

exports.GetRawConection = function (){
    const sequelize = new Sequelize('mssql://sa:288744cla@creditofacilapi.sytes.net:81/credito-facil-homologacao');

    return sequelize
}