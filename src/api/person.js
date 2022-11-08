const express = require('express')
const personCtrl = require('../ctrl/person')

const router = express.Router()
module.exports = router

router.get('/create-person', personCtrl.createPerson)