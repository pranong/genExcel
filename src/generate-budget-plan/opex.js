const Excel = require('exceljs')
const config = require('../config');
const knex = require('../lib/knex')('mysql', config[config.db]);
module.exports = {
    generateBudgetPlanOPEX
}

async function generateBudgetPlanOPEX(budgetYear, dateTime, contactPointDepartment) {
    console.log(budgetYear, dateTime, contactPointDepartment);
    let fileName = 'src/generate-budget-plan/templates/TemplateOPEX.xlsx';
    let wb = new Excel.Workbook();
    let ws = await wb.xlsx.readFile(fileName).then(() => wb.getWorksheet('Sheet1'));

    let writeFileName = `Template Budget OPEX ${budgetYear} - ${contactPointDepartment}.xlsx`;
    wb.xlsx.writeFile(writeFileName)
    .then(() => {
        console.log('file created');
    })
    .catch(err => {
        console.log(err.message);
    });
}

