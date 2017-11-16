var path = require('path')
var model = require(path.join(__dirname, 'model'))
var sequelize = model.GetRawConection()


var ResumoRota = function(usuario, req, res){
  var response = {}
  response.pendentes = []
  response.recebidos = []

  var resumo = {
      qtdAReceber: 0,
      valAReceber: 0,
      qtdRecebida: 0,
      valRecebido: 0
  }

  sequelize.query(`exec [credito-facil-homologacao].[dbo].[sp_CONSULTAR_LISTA_COBRANCA_POR_ROTA] :id`,
    { replacements: { id: usuario.rota } }
  ).then(cobrancas => {
    cobrancas[0].map(cobranca=>{
      if (cobranca.PENDENTE == 'S') {
        response.pendentes.push({id: cobranca.ID_PARCELAMENTO, nome: cobranca.CLIENTE})
      } else {
        response.recebidos.push({id: cobranca.ID_PARCELAMENTO, nome: cobranca.CLIENTE})
      }
    })

    sequelize.query(`exec [credito-facil-homologacao].[dbo].[sp_CONSULTAR_RESUMO_POR_ROTA] :id`,
      { replacements: { id: usuario.rota } }
    ).then(resumos => {
      console.log(resumos[0][0])
          resumo.qtdAReceber = resumos[0][0].QT_RECEBER
          resumo.valAReceber = resumos[0][0].TOTAL_RECEBER
          resumo.qtdRecebida = resumos[0][0].QTD_RECEBIDO
          resumo.valRecebido = resumos[0][0].TOTAL_RECEBIDO
          response.resumo = resumo
          res.status(200).send(response)
      // var response = {}
    })  
  })
}

module.exports = ResumoRota