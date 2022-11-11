const Excel = require('exceljs')
const config = require('../config');
const knex = require('../lib/knex')('dbPlanBudget', config[config.db]);
const dayjs = require('dayjs');
const chalk = require('chalk');
const path = require('path')
const mkdirp = require('mkdirp')

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

    if (contactPointDepartment.toUpperCase() === 'TEAM' || contactPointDepartment.toUpperCase() === 'TEAMS') {
        let queryTeam = knex('pmr.planBudgetOpex as pbo')
        .where('pbo.budgetYear', (('' + budgetYear).trim()) || '')
        .where('pbo.dateTime', knex.raw(`to_date('${dateTime}', 'YYYY-MM-DD')`))
        .distinct(knex.raw(`nvl(pbo.deapartment_owner, 'NULL') team`)); // NVL("PBO"."DEAPARTMENT_OWNER", 'NULL')
        let rows = await queryTeam;
        for (const row of rows) {
            await generateBudgetPlanOPEX(budgetYear, dateTime, row.team)
        }

        return true;
    }

    let vendorNameList = [
        'AIS', // 'AIS (1100)',
        'AWN', // 'AWN (1200)',
        'WDS', // 'WDS (1400)',
        'MMT', // 'MMT (1500)',
        'FXL', // 'FXL (1600)',
        'AMP', // 'AMP (1700)',
        'SBN', // 'SBN (1800)',
        'AIN', // 'AIN (1900)',
        'ACC', // 'ACC (2000)',
        'ABN', // 'ABN (2500)',
    ]
    let headerCols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'
        ,'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'
        ,'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL'
        ,'AM', 'AN', 'AO', 'AP'];
    let numFmtCols = ['N', 'O', 'P', 'Q', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC'
        ,'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP']
    let fileName = 'src/generate-budget-plan/templates/TemplateOPEX.xlsx';
    let wb = new Excel.Workbook();
    let ws = await wb.xlsx.readFile(fileName).then(() => wb.getWorksheet('Sheet1'));
    let query = knex('pmr.planBudgetOpex as pbo')
    .where('pbo.budgetYear', (('' + budgetYear).trim()) || '')
    .where('pbo.dateTime', knex.raw(`to_date('${dateTime}', 'YYYY-MM-DD')`))
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
    ]);
    
    // set header
    ws.getCell('A1').value = 'IT OPEX ' + (dayjs().add(1, 'year').format('YYYY')).toString();

    // table 1 external
    if (!['ALL', 'NULL'].includes(contactPointDepartment.toUpperCase())) {
        query.where(knex.raw('upper(pbo.deapartment_owner)'), (('' + contactPointDepartment).trim()).toUpperCase() || '');
    } else if (['NULL'].includes(contactPointDepartment.toUpperCase())) {
        query.where((q) => {
            q.whereNull('pbo.deapartmentOwner')
        })
    }

    let queryInternal = query.clone();
    let queryExternal = query.clone();
    let externalData = await queryExternal
    .where((vd) =>{
        vd.whereNotIn('pbo.vendorName', vendorNameList)
        vd.orWhereNull('pbo.vendorName')
    });
    // console.log('externalData=', externalData);
    let ok = true;
    let latestRowEx = 0;
    if (externalData.length) {
        let exRowLatest = 5;
        for (const [index, exRow] of externalData.entries()) {
            const excelRow = ws.getRow(exRowLatest);
            excelRow.values = await setupRow(exRow);
            for (const f of numFmtCols) {
                ws.getCell(f + exRowLatest).numFmt = '#,##0.00;[Red]-#,##0.00';
            }

            for (const h of headerCols) {
                if ((index + 1) === externalData.length) {
                    ws.getCell(h + exRowLatest).border = { 
                        top: {style:'dotted'},
                        left: {style:'thin'},
                        bottom: {style:'thin'},
                        right: {style:'thin'}
                    };
                } else {
                    ws.getCell(h + exRowLatest).border = { 
                        top: {style:'dotted'},
                        left: {style:'thin'},
                        bottom: {style:'dotted'},
                        right: {style:'thin'}
                    };
                }
            }
            
            exRowLatest++; // next row excel
            ws.insertRow(exRowLatest, []);
        }

        latestRowEx = exRowLatest;
    }

    // table 2 internal
    let internalData = await queryInternal.whereIn('pbo.vendorName', vendorNameList);
    if (internalData.length) {
        if (!externalData.length) {
            latestRowEx = 5;
        }

        let inRowLatest = latestRowEx + 11;
        for (const [index, inRow] of internalData.entries()) {
            const excelRow = ws.getRow(inRowLatest);
            excelRow.values = await setupRow(inRow);
            for (const f of numFmtCols) {
                ws.getCell(f + inRowLatest).numFmt = '#,##0.00;[Red]-#,##0.00';
            }

            for (const h of headerCols) {
                if ((index + 1) === internalData.length) {
                    ws.getCell(h + inRowLatest).border = { 
                        top: {style:'dotted'},
                        left: {style:'thin'},
                        bottom: {style:'thin'},
                        right: {style:'thin'}
                    };
                } else {
                    ws.getCell(h + inRowLatest).border = { 
                        top: {style:'dotted'},
                        left: {style:'thin'},
                        bottom: {style:'dotted'},
                        right: {style:'thin'}
                    };
                }
            }

            inRowLatest++;
        }
    }

    if (!externalData.length && !internalData.length) {
        ok = false;
    }

    if (ok) {
        let deapartmentOwner = contactPointDepartment.replaceAll(/[<>:"\/\\|?*]+/g, '_')
        let xlsxName = `Template Budget OPEX ${budgetYear} - ${deapartmentOwner.toUpperCase()}.xlsx`;
        // console.log(path.resolve(process.env.OPEX_FOLDER_PATH + '/' + xlsxName));
        await mkdirp(path.resolve(process.env.OPEX_FOLDER_PATH));
        await new Promise((resolve, reject) => setTimeout(resolve, 555));
        let writeFileName = path.resolve(process.env.OPEX_FOLDER_PATH + '/' + xlsxName);
        ws.name = 'OPEX_' + budgetYear;
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
            console.log(chalk.white.bgGreen.bold(' ' + res.message + ' ') + ' ' + xlsxName);
            console.log('');
        }

        if (!res.status) {
            console.log(res.message);
            knex.destroy();
            process.exit();
        }
    }

    if (!ok) {
        console.log('');
        console.log(chalk.white.bgMagenta.bold(' Data not found... '));
        console.log('');
    }

    // return false;
    // knex.destroy();

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
        if (row.contactPointMobile) {
            let contactNumber = row.contactPointMobile.replaceAll(/(\-+|\s+|\_+)/g, '');
            if (Number(contactNumber) > 0) {
                let strContactNumber = contactNumber.substring(0, 3)+ '-' + contactNumber.substring(3, 6)+ '-' + contactNumber.substring(6, contactNumber.length);
                res.push(strContactNumber);
            } else {
                res.push(null);
            }
        } else {
            res.push(row.contactPointMobile);
        }
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
