const express = require('express');
const fsp = require('fs').promises;
const bodyParser = require('body-parser');
const cors = require('cors');
const app = express();
const fs = require('fs')
// const path = require('path')
const http = require('http').createServer(app);
const path = require('path');
// const util = require('./lib/util');
// const config = require('./config');
const Excel = require('exceljs')
const schedule = require('node-schedule');
const moment = require('dayjs');
// const knex = require('./lib/knex')('mysql', config[config.db]);

module.exports = {
    generateBudgetPlanCAPEX
}

async function generateBudgetPlanCAPEX(budgetYear, dateTime, contactPointDepartment) {
        console.log('generateBudgetPlanCAPEX')
      // import template for header
      let fileName = 'src/generate-budget-plan/templates/TemplateCAPEX.xlsx'
      let wb = new Excel.Workbook();
      let ws
      await wb.xlsx.readFile(fileName)
          .then(function() {
              ws = wb.getWorksheet('Sheet1');
          });
      
      // set row data row 4 ++
      const r4 = ws.getRow(4);
      r4.values = [1, 2, 3, 4, 5, 6];
      
      // write file
      let writeFileName = 'Template Budget CAPEX 2023 - ALL.xlsx'
      wb.xlsx
      .writeFile(writeFileName)
      .then(() => {
        console.log('file created');
      })
      .catch(err => {
        console.log(err.message);
      });

}