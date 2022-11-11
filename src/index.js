const express = require('express');
const app = express();
const cors = require('cors');
const path = require('path');
const inquirer = require('inquirer');
const capex = require('./generate-budget-plan/capex');
const opex = require('./generate-budget-plan/opex');
const cp = require("child_process");
const menuList = [
  new inquirer.Separator(),
  'CAPEX',
  'OPEX',
  new inquirer.Separator(),
  'SET->.env',
  'EXIT',
];
let nextList = ['CAPEX', 'OPEX'];

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

  if (menu === 'SET->.env') {
    await executeCommand('C:\\windows\\notepad.exe', [path.resolve('.env')])
  } else if (menu === 'EXIT') {
    process.exit(0);
  }

  if (!nextList.includes(menu)) {
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
      arr = res.context.split(/\s+/g)
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
  } else {
    process.exit(0);
  }

  console.log('');
  let con = await inquirer.prompt({
    type: 'confirm',
    name: 'continue',
    message: `continue:`,
    default: false,
  }).then(res => res.continue)

  if (con) {
    run();
  } else {
    process.exit(0);
  }
  // if (!resp) {
  //   let con = await inquirer.prompt({
  //     type: 'confirm',
  //     name: 'continue',
  //     message: `continue:`,
  //     default: true,
  //   }).then(res => res.continue)
  
  //   if (con) {
  //     run();
  //   } else {
  //     process.exit(0);
  //   }
  // } else {
  //   process.exit(0);
  // }
}

const executeCommand = (textToExecute, arrayList) => new Promise(resolve => {
  const command = cp.spawn(textToExecute, arrayList, { shell: true })
  command.on('close', () => resolve())
})


run()
