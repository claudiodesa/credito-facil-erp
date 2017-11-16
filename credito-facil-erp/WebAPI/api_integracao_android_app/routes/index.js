var express = require('express');
var router = express.Router();
var path = require('path');
var GetCobrancasAgente = require(path.resolve('./model/','GetCobrancasAgente.js'))
var Authenticate = require(path.resolve('./model/','Authenticate.js'))
var Login = require(path.resolve('./model/','Login.js'))
var ResumoRota = require(path.resolve('./model/','ResumoRota.js'))
var RegistrarBaixa = require(path.resolve('./model/','RegistrarBaixa.js'))
var jwt = require('jwt-simple')
const secret = `cr3d170-f4c12`;

/* GET home page. */
router.get('/', function(req, res, next) {
  res.render('index', { title: 'Express' });
});

/* POST Login endpoint. */
router.post('/Login', function(req, res, next) {
  // var login = Authenticate('evandro','3247')  
  // res.status(200).send()
  Authenticate(req,res,(usuario, req, res)=>{
    Login(usuario, req, res)
  })
});

/* GET Resumo endpoint. */
router.get('/ResumoRota', function(req, res, next) {
  var token = req.headers.authorization
  // res.status(200).send('token:'+JSON.stringify(token))
  try {
    var decoded = jwt.decode(token, secret);
    // res.status(200).send(JSON.stringify(decoded))
    req.body.user = decoded.nome
    req.body.pass = decoded.senha
    Authenticate(req,res,(usuario, req, res)=>{
      ResumoRota(usuario, req, res)
    })
  } catch (error) {
    res.status(401).send()
  }
});

/* GET Cobrancas endpoint. */
router.get('/GetCobrancasAgente', function(req, res, next) {
  var token = req.headers.authorization
  try {
    var decoded = jwt.decode(token, secret);
    req.body.user = decoded.nome
    req.body.pass = decoded.senha
    Authenticate(req,res,(usuario, req, res)=>{
      GetCobrancasAgente(usuario, req, res)
    })
  } catch (error) {
    res.status(401).send()
  }
});

/* GET home endpoint. */
router.post('/RegistrarBaixa', function(req, res, next) {
  var token = req.headers.authorization
  try {
    var decoded = jwt.decode(token, secret);
    req.body.user = decoded.nome
    req.body.pass = decoded.senha
    Authenticate(req,res,(usuario, req, res)=>{
      RegistrarBaixa(usuario, req, res)
    })
  } catch (error) {
    res.status(401).send()
  }
});

module.exports = router;