var path = require('path')
var model = require(path.join(__dirname, 'model'))
var sequelize = model.GetRawConection()


var RegistrarBaixa = function(usuario, req, res){
  sequelize.query(`exec [credito-facil-homologacao].[dbo].[sp_INFORMAR_VALOR_RECEBIDO] :id, :valor`,
      { replacements: { id: req.body.id, valor: req.body.valor } }
  ).then(retorno => {
        res.status(200).send('Ok')
  })
}

module.exports = RegistrarBaixa