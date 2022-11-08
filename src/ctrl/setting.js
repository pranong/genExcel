const knex = require('../lib/knex')('mysql')
const path = require('path')
const dayjs = require('dayjs')
const ctrl = {}
module.exports = ctrl

ctrl.getSetting = async (req, res) => {
    try {
      let rows = await knex('settings')
      res.send({
        status: true,
        rows,
      })
    } catch (err) {
      console.error(err)
    }
}
