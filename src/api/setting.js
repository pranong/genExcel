const express = require('express')
const settingCtrl = require('../ctrl/setting')

const router = express.Router()
module.exports = router

router.post('/get-setting', settingCtrl.getSetting)