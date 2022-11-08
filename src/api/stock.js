const express = require('express')
const stockCtrl = require('../ctrl/stock')

const router = express.Router()
module.exports = router

router.post('/get-stock', stockCtrl.getStock)
router.post('/get-stock-item', stockCtrl.getStockItem)