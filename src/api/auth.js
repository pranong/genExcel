const express = require('express')
const { verifySignUp, authJwt } = require("../middleware");
const authCtrl = require('../ctrl/auth')
const accessCtrl = require('../ctrl/access')

const router = express.Router()
module.exports = router

// ------------------- AUTH -------------------
router.post('/signup', [verifySignUp.checkDuplicateUsernameOrEmail, verifySignUp.checkRolesExisted], authCtrl.signup)
router.post("/signin", authCtrl.signin);
// ------------------ ACCESS ------------------
router.get("/all", accessCtrl.allAccess);
router.get("/user", [authJwt.verifyToken], accessCtrl.userBoard);
router.get("/mod", [authJwt.verifyToken, authJwt.isModerator], accessCtrl.moderatorBoard);
router.get("/admin", [authJwt.verifyToken, authJwt.isAdmin], accessCtrl.adminBoard);