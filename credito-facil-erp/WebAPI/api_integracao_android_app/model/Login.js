var jwt = require('jwt-simple')
const secret = `cr3d170-f4c12`;

var Login = function(usuario, req, res){
    var token = jwt.encode(usuario, secret);
    res.setHeader('Authorization', token)
    res.status(200).send({nome: usuario.nome, rota: usuario.rota, token: token})
}

module.exports = Login