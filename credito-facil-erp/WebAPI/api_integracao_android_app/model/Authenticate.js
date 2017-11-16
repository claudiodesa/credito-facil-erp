var path = require('path')
var model = require(path.join(__dirname, 'model'))
var Login = require(path.join(__dirname, 'Login'))
var ResumoRota = require(path.join(__dirname, 'ResumoRota'))
var GetCobrancasAgente = require(path.join(__dirname, 'GetCobrancasAgente'))
var RegistrarBaixa = require(path.join(__dirname, 'RegistrarBaixa'))
var sequelize = model.GetRawConection()

var Authenticate = function(req, res, cb){
    user = req.body.user
    pass = req.body.pass
    var usuario = {
        nome: '',
        senha: pass,
        rota: 0
    }

    sequelize.query(`exec [credito-facil-homologacao].[dbo].sp_CREDITO_FACIL_LOGIN :user,:pass`,
        { replacements: { user: user, pass: pass } }
    ).then(rotas => {
        // console.log(rotas[0][0].LOGIN_VALIDO)
        if (rotas[0][0].LOGIN_VALIDO == 'true') {
            // console.log(rotas[0][0].LOGIN_VALIDO)
            usuario.nome = rotas[0][0].AGENTE
            usuario.rota = rotas[0][0].ROTA
            cb(usuario, req, res)        
        } else {
            res.status(401).send()
        }
        // var response = {}
    })    

    // return false
}

module.exports = Authenticate