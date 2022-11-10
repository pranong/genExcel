const path = require('path')
const util = require('./lib/util')
require('dotenv').config();

module.exports = {
  secret: process.env.ACCESS_SECRET || "secret-key",
  db: process.env.SELECTED_DB || 'precaldev02',

  precaldev02: {
    driver: 'oracledb',
    param: {
      user: 'pmr',
      password: 'pmr123',
      connectString: '10.210.192.23:1521/PRECALDEV02',
      acquireConnectionTimeout: 60 * 1000,
    },
    fetchAsString: [ 'date', 'clob' ],
    test: 1,
    pool: {
      min: 10,
      max: 10,
      createTimeoutMillis: 30 * 1000,
      acquireTimeoutMillis: 30 * 1000,
      idleTimeoutMillis: 30 * 1000,
      reapIntervalMillis: 1000,
      createRetryIntervalMillis: 100,
      propagateCreateError: false,
      validate: async res => {
        console.log('knex validate...')
        try {
          await res.raw('select 1 from dual')
          console.log('knex validate true')
          return true
        } catch (e) {
          console.log('knex validate false')
          return false
        }
      },
    },
    postProcessResponse(result, isRaw) {
      if (isRaw) {
        return result
      }
      if (Array.isArray(result)) {
        return result.map(row => snakeToCamel(row))
      } else {
        return snakeToCamel(result)
      }
    },
    wrapIdentifier(value) {
      if (value === '*') {
        return value
      }
      return '"' + camelToSnake(value) + '"'
    },
  },

  dbPlanBudget: {
    driver: 'oracledb',
    param: {
      user: process.env.DB_USER,
      password: process.env.DB_PASS,
      connectString: `${process.env.DB_HOST}:${process.env.DB_PORT}/${process.env.DB_SID}`,
      acquireConnectionTimeout: 60 * 1000,
    },
    fetchAsString: [ 'date', 'clob' ],
    test: 1,
    pool: {
      min: 10,
      max: 10,
      createTimeoutMillis: 30 * 1000,
      acquireTimeoutMillis: 30 * 1000,
      idleTimeoutMillis: 30 * 1000,
      reapIntervalMillis: 1000,
      createRetryIntervalMillis: 100,
      propagateCreateError: false,
      validate: async res => {
        console.log('knex validate...')
        try {
          await res.raw('select 1 from dual')
          console.log('knex validate true')
          return true
        } catch (e) {
          console.log('knex validate false')
          return false
        }
      },
    },
    postProcessResponse(result, isRaw) {
      if (isRaw) {
        return result
      }
      if (Array.isArray(result)) {
        return result.map(row => snakeToCamel(row))
      } else {
        return snakeToCamel(result)
      }
    },
    wrapIdentifier(value) {
      if (value === '*') {
        return value
      }
      return '"' + camelToSnake(value) + '"'
    },
  },
}

function snakeToCamel(s) {
  let newObj = {}
  Object.keys(s).forEach(k => {
    newObj[k.toLowerCase().replace(/(\w)(_\w)/g, m => m[0] + m[2].toUpperCase())] = s[k]
  })
  return newObj
}

function camelToSnake(s) {
  return s.replace(/[A-Z]/g, m => '_' + m[0]).toUpperCase()
}
