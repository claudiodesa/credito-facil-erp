var express = require('express');
var router = express.Router();

/* GET home page. */
router.get('/', function(req, res, next) {
  res.render('index', { title: 'Express' });
});

/* GET home page. */
router.get('/Login', function(req, res, next) {
  res.render('index', { title: 'Express' });
});

/* GET home page. */
router.get('/GetBaixasAgente', function(req, res, next) {
  // res.status(200).send(['Oi', 'Ola'])
  var agente = {
    nome: 'Natan',
    totalPendencias: 10,
    totalRecebimentos: 15,
    baixas: [
      {
        nomeCliente: 'Claudio',
        fone: '85999998888',
        saldoDevedor: '20,00',
        parcela: 9/10,
        valor: '2,00',
        valorRecebido: '18,00',        
      }
    ]
  }
  res.status(200).send(agente)
});

/* GET home page. */
router.post('/RegistrarBaixa', function(req, res, next) {
  res.status(200).send(['Oi', 'Ola'])
});

module.exports = router;