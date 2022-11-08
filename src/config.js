const path = require('path')
const util = require('./lib/util')
require('dotenv').config();

if (process.env.TELECAR && !process.env.TELECAR_CODE) {
  console.log('ERROR, Environment variable TELECAR_CODE is missing')
  process.exit(1)
}

module.exports = {
  secret: process.env.ACCESS_SECRET || "bezkoder-secret-key",
  db: process.env.CLARET_DB || 'mysql',

  mysql: {
    driver: 'mysql',
    param: {
      host: process.env.DB_HOST,
      port: process.env.DB_PORT,
      user: process.env.DB_USER,
      password: process.env.DB_PASS,
      database: process.env.DB_NAME,
      supportBigNumber: true,
      timezone: '+7:00',
      dateStrings: true,
      charset: 'utf8mb4_unicode_ci',
    },
    fetchAsString: ['date', 'clob'],
    // test: 1,
    pool: {
      min: 10,
      max: 100,
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
        return result.map(row => util.snakeToCamel(row))
      } else {
        return util.snakeToCamel(result)
      }
    },
    wrapIdentifier(value) {
      if (value === '*') {
        return value
      }
      return util.camelToSnake(value)
    },
  },

  /* config for telecar */
  telecar: {
    code: process.env.TELECAR_CODE,
    uploadsDir: path.resolve(__dirname, '../public/files/telecar'),
  },

  ssl: process.env.SSL || 0,

  app: {
    title: 'TBIS',
  },

  server: {
    port: parseInt(process.env.SERVER_PORT || 7000),
    portSSL: parseInt(process.env.SERVER_PORT_SSL || 7001),
    port_dev: 9001,
  },

  monitor: {
    host: '0.0.0.0',
    port: parseInt(process.env.MONITOR_PORT || '9991'),
    normalizePath: true,
  },

  session: {
    secret: '[CLARET]',
    name: process.env.SESSION_NAME || 'CLARET-DEV',
    timeout: parseInt(process.env.SESSION_TIMEOUT || '30'),
    mongodb: {
      uri: process.env.SESSION_MONGO_URI || 'mongodb://127.0.0.1:27017/claret-session',
      collection: process.env.SESSION_MONGO_COLLECTION || 'claret-session',
    },
    redis: {
      host: process.env.SESSION_REDIS_HOST || '10.0.2.226', // redis
      port: parseInt(process.env.SESSIONI_REDIS_PORT || 6379),
      password: process.env.SESSION_REDIS_PASS || 'redis!@1234',
    },
  },

  smw: {
    smwFolder: process.env.SMW_BASE || './public/files/smw',
    importResultFolder: process.env.SMW_RESULT || 'res',
    exportRequestFolder: process.env.SMW_REQUEST || 'req',
    exportBackupFolder: process.env.SMW_REQUEST_BACKUP || 'req/backup',
    logResultFolder: process.env.SMW_RESULT_LOG || 'res/log',
    logResultBySampleFolder: process.env.SMW_LOG_SAMPLE || 'res/log-sample',
    importResultNewFolder: process.env.SMW_RESULT_NEW || 'res/NEW',
    backupResultFolder: process.env.SMW_REsuLT_SAVE || 'res/sav',
    importSite: process.env.SMW_SITE || 'unknown',
  },

  socket: {
    url: process.env.SOCKET_URL || 'http://localhost:7000', // http://socket:8999,
    urlSSL: process.env.SOCKET_URL_SSL || 'http://localhost:7001', // http://socket:8998,
    user: process.env.SOCKET_USER || 'claret',
    pass: process.env.SOCKET_PASS || 'g8iupf',
  },


  bloodMixer: {
    baseFolder: process.env.BLOODMIXER_BASE || './public/files/blood-mixer',
    reqFolder: process.env.BLOODMIXER_REQ || 'req',
    logFolder: process.env.BLOODMIXER_LOG || 'log',
    bakFolder: process.env.BLOODMIXER_BAK || 'bak',
  },

  export: {
    baseFolder: process.env.EXPORT_BASE_FOLDER || './public/files',
    resultFolder: process.env.EXPORT_RESULT_FOLDER || './public/files/exportResult',
    deliveryFolder: process.env.EXPORT_DELIVERY_FOLDER || './public/files/exportDelivery',
    invoiceFolder: process.env.EXPORT_INVOICE_FOLDER || './public/files/invoice',
    smsFolder: process.env.EXPORT_SMS_FOLDER || './public/files/sms',
    stockDisplayFolder: process.env.EXPORT_STOCK_DISPLAY_FOLDER || './public/files/stockDisplay',
    lookbackFolder: process.env.EXPORT_LOOKBACK_FOLDER || './public/files/lookback',
    searchExcelFolder: process.env.EXPORT_SEARCH_LIST || './public/files/searchExcel',
    donorPicturesFolder: process.env.EXPORT_DONOR_PICTURES_FOLDER || './public/files/donorPictures',
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
