const express = require("express");

const router = express.Router()
module.exports = router

router.use('/person', require('./person'))
router.use('/stock', require('./stock'))
router.use('/setting', require('./setting'))
router.use('/auth', require('./auth'))

router.get('/', (req, res) => {
  res.send({
    status: true,
  })
})
