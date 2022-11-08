const jwt = require("jsonwebtoken");
const config = require("../config");

const roleLookup = {
    'user': 1,
    'moderator': 2,
    'admin': 3,
    '1': 'user',
    '2': 'moderator',
    '3': 'admin',
}

verifyToken = (req, res, next) => {
    let token = req.headers["x-access-token"];

    if (!token) {
        return res.status(403).send({
            message: "No token provided!"
        });
    }

    jwt.verify(token, config.secret, (err, decoded) => {
        if (err) {
            return res.status(401).send({
                message: "Unauthorized!"
            });
        }
        req.userId = decoded.id;
        next();
    });
};

isAdmin = async (req, res, next) => {
    // query userRole by 'req.userId'
    let roles = []
    await req.$db('userRoles').where('userid', req.userId).then(item => {
        console.log('item', item)
        for (let i = 0; i < item.length; i++) {
            const e = item[i];
            roles.push(e.roleid)
        }
    })
    // check admin included
    let checkAdminRole = roles.includes(3)
    if (checkAdminRole) {
        next();
        return;
    }
    res.status(403).send({
        status: false,
        message: "Require Admin Role!"
    });
    return;
};

isModerator = async (req, res, next) => {
    // query userRole by 'req.userId'
    let roles = []
    await req.$db('userRoles').where('userid', req.userId).then(item => {
        console.log('item', item)
        for (let i = 0; i < item.length; i++) {
            const e = item[i];
            roles.push(e.roleid)
        }
    })
    // check admin included
    let checkModRole = roles.includes(2)
    if (checkModRole) {
        next();
        return;
    }
    res.status(403).send({
        status: false,
        message: "Require Moderator Role!"
    });
    return;
};

isModeratorOrAdmin = async (req, res, next) => {
    // query userRole by 'req.userId'
    let roles = []
    await req.$db('userRoles').where('userid', req.userId).then(item => {
        console.log('item', item)
        for (let i = 0; i < item.length; i++) {
            const e = item[i];
            roles.push(e.roleid)
        }
    })
    // check admin included
    let checkAdminRole = roles.includes(3)
    let checkModRole = roles.includes(2)
    if (checkAdminRole || checkModRole) {
        next();
        return;
    }
    res.status(403).send({
        status: false,
        message: "Require Admin/Moderator Role!"
    });
    return;
};

const authJwt = {
    verifyToken: verifyToken,
    isAdmin: isAdmin,
    isModerator: isModerator,
    isModeratorOrAdmin: isModeratorOrAdmin
};
module.exports = authJwt;