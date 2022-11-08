// const db = require("../models");
// const ROLES = db.ROLES;
// const User = db.user;

checkDuplicateUsernameOrEmail = async (req, res, next) => {
    try {
        let userCheck = await req.$db('users').where('username', req.body.username)
        console.log('userCheck', userCheck)
        if (userCheck.length) {
            res.status(400).send({
                message: "Failed! Username is already in use!"
            });
            return;
        }
        let emailCheck = await req.$db('users').where('email', req.body.email)
        if (emailCheck.length) {
            res.status(400).send({
                message: "Failed! Email is already in use!"
            });
            return;
        }
        next();
    } catch (error) {
        res.status(400).send({
            message: error
        });
    }
};

checkRolesExisted = (req, res, next) => {
    if (req.body.roles) {
        let role = ["user", "admin", "moderator"]
        for (let i = 0; i < req.body.roles.length; i++) {
            console.log('ROLEEEEE', role)
            if (!role.includes(req.body.roles[i])) {
                res.status(400).send({
                    message: "Failed! Role does not exist = " + req.body.roles[i]
                });
                return;
            }
        }
    }

    next();
};

const verifySignUp = {
    checkDuplicateUsernameOrEmail: checkDuplicateUsernameOrEmail,
    checkRolesExisted: checkRolesExisted
};

module.exports = verifySignUp;