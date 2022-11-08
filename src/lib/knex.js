'use strict'

const dbPool = {}

const durations = {}

module.exports = (dbName, config) => {
  if (!dbPool[dbName]) {
    const opt = {
      client: config.driver,
      connection: config.param,
      debug: process.env.DEBUG === '1',
    }
    Object.keys(config).forEach(key => {
      if (['driver', 'param', 'init'].indexOf(key) >= 0) {
        return
      }
      opt[key] = config[key]
    })
    console.log('knex opt=', opt)
    const knex = require('knex')(opt)
    dbPool[dbName] = knex
    if (process.env.DEBUG === '1') {
      knex.on('query', querySpec => {
        durations[querySpec.__knexQueryUid] = new Date().getTime()
      })

      knex.on('query-response', (res, querySpec) => {
        let startTime = durations[querySpec.__knexQueryUid]
        if (!startTime) {
          return console.log('no start time')
        }
        if (process.env.DEBUG === '1') {
          console.log(querySpec.__knexQueryUid, 'DURATION=', new Date().getTime() - startTime)
        }
        delete durations[querySpec.__knexQueryUid]
      })
    }
  }
  return dbPool[dbName]
}
