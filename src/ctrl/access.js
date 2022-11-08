const ctrl = {}
module.exports = ctrl

ctrl.allAccess = async (req, res) => {
  res.status(200).send({ status: true, message: "Public Content." });
}

ctrl.userBoard = async (req, res) => {
  res.status(200).send({ status: true, message: "User Content." });
}

ctrl.moderatorBoard = async (req, res) => {
  res.status(200).send({ status: true, message: "Moderator Content." });
}

ctrl.adminBoard = async (req, res) => {
  res.status(200).send({ status: true, message: "Admin Content." });
}
