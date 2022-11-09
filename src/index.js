const express = require('express');
const app = express();
const cors = require('cors');
// const fsp = require('fs').promises;
// const bodyParser = require('body-parser');
// const fs = require('fs')
// const path = require('path')
// const http = require('http').createServer(app);
// const path = require('path');
// const util = require('./lib/util');
// const config = require('./config');
// const Excel = require('exceljs')
// const schedule = require('node-schedule');
// const moment = require('dayjs');
// const knex = require('./lib/knex')('mysql', config[config.db]);
const inquirer = require('inquirer')
const chalk = require('chalk')
const capex = require('./generate-budget-plan/capex');
const opex = require('./generate-budget-plan/opex');
const menuList = [
  new inquirer.Separator(),
  'CAPEX',
  'OPEX',
  new inquirer.Separator(),
  'EXIT',
]


// ---------------------------- SETUP ----------------------------
// use cors
app.use(cors());
// app.use(cors(corsOptions));
// parse requests of content-type - application/json
// app.use(express.json());
// parse requests of content-type - application/x-www-form-urlencoded
// app.use(express.urlencoded({ extended: true }));
// setup header and global use
// app.use((req, res, next) => {
//   // res.header(
//   //   "Access-Control-Allow-Headers",
//   //   "x-access-token, Origin, Content-Type, Accept"
//   // );
//   req.$db = knex;
//   req.$util = util;
//   req.$config = config;
//   // req.$redis = asyncRedisClient
//   next();
// });

run = async () => {
  // console.log(chalk)
  console.clear();
  let menu = await inquirer.prompt({
    type: 'list',
    name: 'menu',
    message: `GENERATE BUDGET PLAN: `,
    choices: menuList,
  }).then(res => res.menu);

  // console.log(menu);

  if (menu === 'EXIT') {
    process.exit(0);
  }

  let budgetYear = '';
  let dateTime = '';
  let contactPointDepartment = '';
  let context = {}
  let conditionsList = []

  conditionsList.push({ // donorNumber
    type: 'input',
    name: 'context',
    message: `ENTER CONDITIONS:`,
  })

  await inquirer.prompt(conditionsList).then(res => {
    let arr = [];
    if (res.context) {
      arr = res.context.split(' ')
      if (arr.length) {
        context = {
          budgetYear: arr[0],
          dateTime: arr[1],
          contactPointDepartment: arr[2],
        }
      }
    }
    // context = JSON.parse(res.context);
  })

  console.log('context=', context);


  let resp = true;
  budgetYear = context.budgetYear;
  dateTime = context.dateTime;
  contactPointDepartment = context.contactPointDepartment;
  if (menu === 'CAPEX') {
    resp = await capex.generateBudgetPlanCAPEX(budgetYear, dateTime, contactPointDepartment)
  } else if (menu === 'OPEX') {
    resp = await opex.generateBudgetPlanOPEX(budgetYear, dateTime, contactPointDepartment)
  }

  if (!resp) {
    let con = await inquirer.prompt({
      type: 'confirm',
      name: 'continue',
      message: `continue:`,
      default: true,
    }).then(res => res.continue)
  
    if (con) {
      run();
    } else {
      process.exit(0);
    }
  } else {
    process.exit(0);
  }
}


run()
