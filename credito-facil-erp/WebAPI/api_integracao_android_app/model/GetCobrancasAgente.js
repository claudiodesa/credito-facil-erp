var path = require('path')
var model = require(path.join(__dirname, 'model'))
var sequelize = model.GetRawConection()


var GetCobrancasAgente = function(usuario, req, res){
    var recebimento = {}
    sequelize.query(`exec [credito-facil-homologacao].[dbo].[sp_CONSULTAR_COBRANCA] :id`,
        { replacements: { id: req.query.id } }
    ).then(rotas => {
        rotas[0].map(rota=>{
          recebimento.id = rota.ID_PARCELAMENTO
          recebimento.nomeCliente = rota.CLIENTE
          recebimento.fone = rota.TELEFONE1
          recebimento.saldoDevedor = rota.SALDO_DEVEDOR
          recebimento.parcela = rota.PARCELAS_PAGAS
          recebimento.valor = rota.VALOR_DEVIDO
          recebimento.valorRecebido = rota.VL_RECEBIDO_INFORMADO
          res.status(200).send(recebimento)
        })
    })
}

module.exports = GetCobrancasAgente