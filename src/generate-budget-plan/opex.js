const Excel = require('exceljs')
const config = require('../config');
const knex = require('../lib/knex')('mysql', config[config.db]);
const dayjs = require('dayjs');
const chalk = require('chalk');
module.exports = {
    generateBudgetPlanOPEX
}

async function generateBudgetPlanOPEX(budgetYear, dateTime, contactPointDepartment) {
    // console.log(budgetYear, dateTime, contactPointDepartment);
    // console.log(dayjs(dayjs(dateTime).format('YYYYMMDD')).format('YYYY-MM-DD'));

    // setup
    if (!budgetYear) {
        console.log('budgetYear required');
        return false;
    }
    
    if (!dateTime) {
        console.log('dateTime required');
        return false;
    }

    if (!contactPointDepartment) {
        console.log('contactPointDepartment required');
        return false;
    }

    if (dateTime && dateTime.indexOf('-') === -1) {
        dateTime = dayjs(dayjs(dateTime).format('YYYYMMDD')).format('YYYY-MM-DD');
    }

    let vendorNameList = [
        'AIS (1100)',
        'AWN (1200)',
        'WDS (1400)',
        'MMT (1500)',
        'FXL (1600)',
        'AMP (1700)',
        'SBN (1800)',
        'AIN (1900)',
        'ACC (2000)',
        'ABN (2500)'
    ]
    let fileName = 'src/generate-budget-plan/templates/TemplateOPEX.xlsx';
    let wb = new Excel.Workbook();
    let ws = await wb.xlsx.readFile(fileName).then(() => wb.getWorksheet('Sheet1'));
    let query = knex('pmr.planBudgetOpex as pbo')
    .where('pbo.budgetYear', (('' + budgetYear).trim()) || '')
    .where('pbo.dateTime', knex.raw(`to_date('${dateTime}', 'YYYY-MM-DD')`))
    .where('pbo.deapartmentOwner', (('' + contactPointDepartment).trim()) || '')
    .select([
        'budgetYear',
        'budgetType',
        'companyCode',
        'glCode',
        'costCenter',
        'titleActivityProject',
        'businessProduct',
        'businessFunction',
        'deapartmentOwner',
        'budgetOwnerName',
        'contactPointMobile',
        'budgetAssumption',
        'vendorName',
        'poValueAmountUsd',
        'poValueAmountThb',
        'poValueEquivalentToThb',
        'actualSap',
        knex.raw(`to_char(contract_period_start_date, 'DD/MM/YY') || ' - ' || to_char(contract_period_end_date, 'DD/MM/YY') contract_period`),
        'curJan',
        'curFeb',
        'curMar',
        'curApr',
        'curMay',
        'curJun',
        'curJul',
        'curAug',
        'curSep',
        'curOct',
        'curNov',
        'curDec',
        'nextJan',
        'nextFeb',
        'nextMar',
        'nextApr',
        'nextMay',
        'nextJun',
        'nextJul',
        'nextAug',
        'nextSep',
        'nextOct',
        'nextNov',
        'nextDec',
    ])
    
    // set header
    ws.getCell('A1').value = 'IT OPEX ' + (dayjs().add(1, 'year').format('YYYY')).toString();
    // ws.getRow(1).values[1] = 
    // console.log(ws.getRow(1).values[1]);

    // table 1 external
    let externalRows = await query.whereNotIn('pbo.vendorName', vendorNameList);
    // console.log('externalRows=', externalRows);
    let rowNum = 5;
    let ok = true;
    for (const row of externalRows) {
        const excelRow = ws.getRow(rowNum);
        // console.log('setupRow=', await setupRow(row));
        excelRow.values = await setupRow(row);
        ws.getCell('S' + rowNum).numFmt = '#,##0.00;[Red]-#,##0.00';

        rowNum++ // next row excel
    }

    // ws.getCell('S5').numFmt = '#,##0.00;[Red]-#,##0.00';
    // ws.getCell('S6').numFmt = '#,##0.00;[Red]-#,##0.00';
    // table 2 internal

    if (ok) {
        let deapartmentOwner = contactPointDepartment.replaceAll(/[<>:"\/\\|?*]+/g, '-')
        let writeFileName = `Template Budget OPEX ${budgetYear} - ${deapartmentOwner}.xlsx`;
        // console.log('Write file:', writeFileName);
        const writeFilePromise = new Promise((resolve, reject) => {
            wb.xlsx.writeFile(writeFileName)
            .then(() => {
                resolve({status: true, message: 'File created'});
            })
            .catch(err => {
                resolve({status: false, message: err.message});
                
            });
        });
        
        let res = await writeFilePromise;
        if (res.status) {
            console.log('');
            console.log(chalk.white.bgGreen.bold(' ' + res.message + ' '));
            console.log('');
        }

        if (!res.status) {
            console.log(res.message);
            knex.destroy();
            process.exit();
        }
    }

    // return false;
    knex.destroy();

    return true;
}

async function setupRow(row) {
    // console.log('setupRow', row);
    let res = [];
    if (row) {
        res.push(row.budgetYear);
        res.push(row.budgetType);
        res.push(row.companyCode);
        res.push(row.glCode);
        res.push(row.costCenter);
        res.push(row.titleActivityProject);
        res.push(row.businessProduct);
        res.push(row.businessFunction);
        res.push(row.deapartmentOwner);
        res.push(row.budgetOwnerName);
        res.push(row.contactPointMobile);
        res.push(row.budgetAssumption);
        res.push(row.vendorName);
        res.push(row.poValueAmountUsd);
        res.push(row.poValueAmountThb);
        res.push(row.poValueEquivalentToThb);
        res.push(row.actualSap);
        res.push(row.contractPeriod);
        res.push(row.curJan);
        res.push(row.curFeb);
        res.push(row.curMar);
        res.push(row.curApr);
        res.push(row.curMay);
        res.push(row.curJun);
        res.push(row.curJul);
        res.push(row.curAug);
        res.push(row.curSep);
        res.push(row.curOct);
        res.push(row.curNov);
        res.push(row.curDec);
        res.push(row.nextJan);
        res.push(row.nextFeb);
        res.push(row.nextMar);
        res.push(row.nextApr);
        res.push(row.nextMay);
        res.push(row.nextJun);
        res.push(row.nextJul);
        res.push(row.nextAug);
        res.push(row.nextSep);
        res.push(row.nextOct);
        res.push(row.nextNov);
        res.push(row.nextDec);
    }
    

    return res;
}
