var path = require('path')
var model = require(path.join(__dirname, 'model'))
// var sequelize = model.GetConection()
var Sequelize = require('sequelize');
const sequelize = new Sequelize('mssql://sa:288744cla@177.42.171.242:81/credito-facil-homologacao');
var Connection = require('tedious').Connection;
var Request = require('tedious').Request;

  var config = {
    userName: 'sa',
    password: '288744cla',
    server: 'credito-facil.cptqzj1ct8cm.sa-east-1.rds.amazonaws.com',
    
    // If you're on Windows Azure, you will need this:
    // options: {encrypt: true}
  };

  // var connection = new Connection(config);

  // connection.on('connect', function(err) {
  //   // If no error, then good to go...
  //     // executeStatement();
  //     console.log('VTNC')
  //   }
  // );

var GetBaixasAgente = function(id_agente, req, res){
  var connection = new Connection(config);

  connection.on('connect', function(err) {
    
  function executeStatement() {
    request = new Request("select 42, 'hello world'", function(err, rowCount) {
      if (err) {
        console.log(err);
      } else {
        console.log(rowCount + ' rows');
      }
    });

    request.on('row', function(columns) {
      columns.forEach(function(column) {
        console.log(column.value);
      });
    });

    connection.execSql(request);
  }
    // If no error, then good to go...
      // executeStatement();
      console.log('VTNC')
    }
  );

  sequelize.query(`use [credito-facil-homologacao]
    exec SP_CONSULTAR_COBRANCA_POR_AGENTE :id`,
    { replacements: { id: id_agente } }
  ).then(rotas => {
    console.log(rotas)
    // var response = {}
  })
  res.status(200).send()
}

module.exports = GetBaixasAgente