const knex = require('../lib/knex')('mysql')
const dayjs = require('dayjs')

const config = require("../config");
var jwt = require("jsonwebtoken");
var bcrypt = require("bcryptjs");

const ctrl = {}
module.exports = ctrl

const roleLookup = {
  'user': 1,
  'moderator': 2,
  'admin': 3,
  '1': 'user',
  '2': 'moderator',
  '3': 'admin',
}

ctrl.signin = async (req, res) => {
  try {
    let userCheck = await req.$db('users').where('username', req.body.username)
    if (!userCheck.length) {
      return res.status(404).send({ status: false, message: "User Not found." })

    }
    let user = userCheck[0]
    var passwordIsValid = bcrypt.compareSync(
      req.body.password,
      user.password
    );
    if (!passwordIsValid) {
      return res.status(401).send({
        status: false,
        accessToken: null,
        message: "Invalid Password!"
      });
    }
    var token = jwt.sign({ id: user.id }, config.secret, {
      expiresIn: 86400 // 24 hours
    });
    var authorities = [];
    let userRoles = await req.$db('userRoles').where('userid', user.id)
    // roleLookup
    for (let i = 0; i < userRoles.length; i++) {
      let roleName = roleLookup[userRoles[i].roleid || 1]
      authorities.push("ROLE_" + roleName.toUpperCase());
    }
    return res.status(200).send({
      status: true,
      id: user.id,
      username: user.username,
      email: user.email,
      roles: authorities,
      accessToken: token
    });
  } catch (err) {
    console.error(err)
    res.send({
      status: false,
      message: err
    })
  }
}

ctrl.signup = async (req, res) => {
  try {
    console.log('yay req', req.body)
    // insert USER
    await req.$db('users').insert({
      username: req.body.username,
      email: req.body.email,
      password: bcrypt.hashSync(req.body.password, 8)
    })
    let row = await req.$db('users').where('username', req.body.username).select('id')
    // insert ROLE
    if (req.body.roles) {
      for (let i = 0; i < req.body.roles.length; i++) {
        let item = req.body.roles[i]
        let rid = roleLookup[item]
        await req.$db('userRoles').insert({
          userid: row[0].id,
          roleid: rid,
        })
      }
    } else {
      await req.$db('userRoles').insert({
        userid: row[0].id,
        roleid: 1,
      })
    }
    res.send({
      status: true,
      message: "User was registered successfully!"
    })

  } catch (err) {
    console.error(err)
    res.send({
      status: false,
      message: err
    })
  }
}
