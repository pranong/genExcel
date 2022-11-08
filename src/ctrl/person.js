const knex = require('../lib/knex')('mysql')
const path = require('path')
const dayjs = require('dayjs')
const createReport = require('../lib/create-report')
const ctrl = {}
module.exports = ctrl

ctrl.createPerson = async (req, res) => {
    try {
      console.log('yay req', req.body)
      // let rows = await knex('site') // import and use
      let staff = await req.$db('staff').where('code', 'naaNut    ')
      let date = dayjs().format('DDMMYY')
      let outFilename = `reportTest${date}.docx`
      let outputFile = path.resolve('./public/' + outFilename)
      await createReport({
        cmdDelimiter: '`',
        template: './reports/template.docx',
        output: outputFile,
        data: {
          list: staff,
          name: 'NutJa',
          age: 69
        },
        })
      console.log('data', staff)

      res.send({
        status: true,
        staff
      })
      //util
      await req.$util.promiseForEach(staff, async row => {
        let { id, idx, productPrice, ...data } = row
        console.log('=>', row.id)
      })
      // insert update delete
    } catch (err) {
      console.error(err)
    }
  }
