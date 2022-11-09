const Excel = require('exceljs')
const moment = require('dayjs');
const chalk = require('chalk');

// const util = require('./lib/util');
const config = require('../config');
const knex = require('../lib/knex')('precaldev02', config['precaldev02']);

module.exports = {
    generateBudgetPlanCAPEX
}

async function generateBudgetPlanCAPEX(budgetYear, dateTime, contactPointDepartment) {
    console.log('generating Budget Plan CAPEX...')

    // verrify data
    console.log('Verrify data...')
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
        dateTime = moment(moment(dateTime).format('YYYYMMDD')).format('YYYY-MM-DD');
    }

    // import template for header
    console.log('Importing template...')
    let fileName = './src/generate-budget-plan/templates/TemplateCAPEX.xlsx'
    let wb = new Excel.Workbook();
    let ws
    await wb.xlsx.readFile(fileName).then(function() {
        ws = wb.getWorksheet('Sheet1');
    });

    // get data
    console.log('Fetching data...')
    // let select = `null mock,null mock,budget_name,busines_function,business_group,business_owner_name,sub_project_product,project_priority,budget_company,cost_center,contact_point_budget_own_name,contact_point_department,contact_point_mobile,budget_amount_usd,budget_amount_thb,equivalent_to_thb,equivalent_to_usd,assumption_budget_calculation,project_description,hardware_total,software_license_total,software_dev_turnkey_total,software_dev_dbp_total,software_dev_automate_total,external_outsourcing_total,outsource_total_total,outsource_pm_si_total,outsource_sa_total,outsource_pa_total,outsource_tester_total,outsource_ts_total,outsource_tc_total,null mock,plan_use_jan,plan_use_feb,plan_use_mar,plan_use_apr,plan_use_may,plan_use_jun,plan_use_jul,plan_use_aug,plan_use_sep,plan_use_oct,plan_use_nov,plan_use_dec,plan_use_total,null mock,forecast_inv_cur_jan,forecast_inv_cur_feb,forecast_inv_cur_mar,forecast_inv_cur_apr,forecast_inv_cur_may,forecast_inv_cur_jun,forecast_inv_cur_jul,forecast_inv_cur_aug,forecast_inv_cur_sep,forecast_inv_cur_oct,forecast_inv_cur_nov,forecast_inv_cur_dec,forecast_inv_cur_total,null mock,forecast_inv_next_jan,forecast_inv_next_feb,forecast_inv_next_mar,forecast_inv_next_apr,forecast_inv_next_may,forecast_inv_next_jun,forecast_inv_next_jul,forecast_inv_next_aug,forecast_inv_next_sep,forecast_inv_next_oct,forecast_inv_next_nov,forecast_inv_next_dec,forecast_inv_next_total,investment_Level_1 lvl1,investment_Level_2 lvl2,investment_Level_3 lvl3,investment_Level_4 lvl4`
    // rows = await knex('planBudgetCapex').select(knex.raw(select)).where('budgetYear', '2023').orderBy('investmentLevel_1').orderBy('investmentLevel_2').orderBy('investmentLevel_3').orderBy('investmentLevel_4')
    budgetYear = '2023'
    let rows = await knex.raw(`
    with main as (
        select
        (REPLACE(SUBSTR(investment_Level_1, 0, Instr(investment_Level_1, ' ')), '')) orderl1,
        (REPLACE(SUBSTR(investment_Level_2, 0, Instr(investment_Level_2, ' ')), '')) orderl2,
        (REPLACE(SUBSTR(investment_Level_3, 0, Instr(investment_Level_3, ' ')), '')) orderl3,
        (REPLACE(SUBSTR(investment_Level_4, 0, Instr(investment_Level_4, ' ')), '')) orderl4,
        p.* from plan_Budget_Capex p
    ), gen as (
    select 
            (LPAD(REPLACE(orderl1, '.'), 8, '0')) neworder1,
            (LPAD(REPLACE(orderl2, '.'), 8, '0')) neworder2,
            (LPAD(REPLACE(orderl3, '.'), 8, '0')) neworder3,
            (LPAD(REPLACE(orderl4, '.'), 8, '0')) neworder4,
            m.*
    from main m
    )
    select 
        null mock,null mock,budget_name,busines_function,business_group,business_owner_name,sub_project_product,project_priority,budget_company,cost_center,contact_point_budget_own_name,contact_point_department,contact_point_mobile,budget_amount_usd,budget_amount_thb,equivalent_to_thb,equivalent_to_usd,assumption_budget_calculation,project_description,hardware_total,software_license_total,software_dev_turnkey_total,software_dev_dbp_total,software_dev_automate_total,external_outsourcing_total,outsource_total_total,outsource_pm_si_total,outsource_sa_total,outsource_pa_total,outsource_tester_total,outsource_ts_total,outsource_tc_total,null mock,plan_use_jan,plan_use_feb,plan_use_mar,plan_use_apr,plan_use_may,plan_use_jun,plan_use_jul,plan_use_aug,plan_use_sep,plan_use_oct,plan_use_nov,plan_use_dec,plan_use_total,null mock,forecast_inv_cur_jan,forecast_inv_cur_feb,forecast_inv_cur_mar,forecast_inv_cur_apr,forecast_inv_cur_may,forecast_inv_cur_jun,forecast_inv_cur_jul,forecast_inv_cur_aug,forecast_inv_cur_sep,forecast_inv_cur_oct,forecast_inv_cur_nov,forecast_inv_cur_dec,forecast_inv_cur_total,null mock,forecast_inv_next_jan,forecast_inv_next_feb,forecast_inv_next_mar,forecast_inv_next_apr,forecast_inv_next_may,forecast_inv_next_jun,forecast_inv_next_jul,forecast_inv_next_aug,forecast_inv_next_sep,forecast_inv_next_oct,forecast_inv_next_nov,forecast_inv_next_dec,forecast_inv_next_total,CASE WHEN investment_Level_1 IS NOT NULL THEN investment_Level_1 ELSE '6. Other' END lvl1,CASE WHEN investment_Level_2 IS NOT NULL THEN investment_Level_2 ELSE '6. Other' END lvl2,investment_Level_3 lvl3,investment_Level_4 lvl4, date_time
    from gen
    -- where budget_year = '${budgetYear}' and date_time = to_date('${dateTime}', 'YYYY-MM-DD') ${contactPointDepartment === 'ALL' ? '' : `and CONTACT_POINT_DEPARTMENT = '${contactPointDepartment}'`}
    order by neworder1, neworder2, neworder3, neworder4
    `)

    if (!rows.length) {
        console.log('No DATA')
        return false
    }

    // Re-arrange data
    // Set split No. and Name
    rows = rows.map(x => {
        x.lvl1 = x.lvl1 && x.lvl1.replace('  ', ' ')
        x.lvl2 = x.lvl2 && x.lvl2.replace('  ', ' ')
        x.lvl3 = x.lvl3 && x.lvl3.replace('  ', ' ')
        x.lvl4 = x.lvl4 && x.lvl4.replace('  ', ' ')

        x.lvl1No = x.lvl1 && x.lvl1.substring(0, x.lvl1.indexOf(' ')) || null
        x.lvl2No = x.lvl2 && x.lvl2.substring(0, x.lvl2.indexOf(' ')) || null
        x.lvl3No = x.lvl3 && x.lvl3.substring(0, x.lvl3.indexOf(' ')) || null
        x.lvl4No = x.lvl4 && x.lvl4.substring(0, x.lvl4.indexOf(' ')) || null

        x.lvl1Name = x.lvl1 && x.lvl1.substring(x.lvl1.indexOf(' ') + 1).trim() || null
        x.lvl2Name = x.lvl2 && x.lvl2.substring(x.lvl2.indexOf(' ') + 1).trim() || null
        x.lvl3Name = x.lvl3 && x.lvl3.substring(x.lvl3.indexOf(' ') + 1).trim() || null
        x.lvl4Name = x.lvl4 && x.lvl4.substring(x.lvl4.indexOf(' ') + 1).trim() || null
        return x
    })
    
    let arranged = []
    let level = ['', '', '', '']
    let levelStyle = {
        lvl1: [],
        lvl2: [],
        lvl3: [],
        lvl4: [],
        lvl5: [],
        lvl1Total: [],
        lvl2Total: [],
        lvl3Total: [],
        lvl4Total: [],
        lvlGTotal: [],
    }
    let pos = 0
    let nowLvl
    let nowLvlTotal = 0
    let countLvl = {
        lvl1: 0,
        lvl2: 0,
        lvl3: 0,
        lvl4: 0,
        g: 0,
    }
    let stack1 = 0
    let stack2 = 0
    let stack3 = 0
    let lastSubRow = false
    for (let i = 0; i < rows.length; i++) {
        const item = rows[i];
        if ((!level[0] || (level[0] && level[0] !== item.lvl1)) && item.lvl1) {
            level[0] = item.lvl1
            arranged.push({lvl1: item.lvl1})
            // arranged.push({lvl1No: item.lvl1No, lvl1Name: item.lvl1Name})
            levelStyle.lvl1.push(pos)
            nowLvlTotal = 1
            pos++
        }
        if ((!level[1] || (level[1] && level[1] !== item.lvl2)) && item.lvl2) {
            level[1] = item.lvl2
            arranged.push({lvl2No: item.lvl2No, lvl2Name: item.lvl2Name})
            levelStyle.lvl2.push(pos)
            nowLvlTotal = 2
            pos++
        }
        if ((!level[2] || (level[2] && level[2] !== item.lvl3)) && item.lvl3) {
            level[2] = item.lvl3
            arranged.push({lvl3No: item.lvl3No, lvl3Name: item.lvl3Name})
            levelStyle.lvl3.push(pos)
            nowLvlTotal = 3
            pos++
        }
        if ((!level[3] || (level[3] && level[3] !== item.lvl4)) && item.lvl4) {
            level[3] = item.lvl4
            arranged.push({lvl4No: item.lvl4No, lvl4Name: item.lvl4Name})
            levelStyle.lvl4.push(pos)
            nowLvlTotal = 4
            pos++
        }

        // push data
        arranged.push(item)
        lastSubRow = true
        levelStyle.lvl5.push(pos)
        countLvl.lvl1 += 1
        stack1 += 1
        countLvl.g += 1
        if (nowLvlTotal === 2) {
            countLvl.lvl2 += 1
        }
        if (nowLvlTotal === 3) {
            countLvl.lvl3 += 1
            stack2 += 1
        }
        if (nowLvlTotal === 4) {
            countLvl.lvl4 += 1
            stack3 += 1
        }
        pos++

        // check next if over push total
        let tmpItem = rows[i + 1]
        if (tmpItem) {
            if ((!level[3] || (level[3] && level[3] !== tmpItem.lvl4)) && item.lvl4) {
                if (lastSubRow) {
                    arranged.push({mock: null})
                    pos++
                }
                arranged.push({mock: null, lvl4: 'TOTAL FOR ' + item.lvl4Name + ' ' + countLvl.lvl4})
                countLvl.lvl4 = 0
                levelStyle.lvl4Total.push(pos)
                pos++
                arranged.push({mock: null})
                pos++
                nowLvl = 4
                lastSubRow = false
            }
            if ((!level[2] || (level[2] && level[2] !== tmpItem.lvl3)) && item.lvl3) {
                if (lastSubRow) {
                    arranged.push({mock: null})
                    pos++
                }
                arranged.push({mock: null, lvl3: 'TOTAL FOR ' + item.lvl3Name + ' ' + (countLvl.lvl3 === 0 ? stack3 : countLvl.lvl3)})
                countLvl.lvl3 = 0
                stack3 = 0
                levelStyle.lvl3Total.push(pos)
                pos++
                arranged.push({mock: null})
                pos++
                nowLvl = 3
                lastSubRow = false
            }
            if ((!level[1] || (level[1] && level[1] !== tmpItem.lvl2)) && item.lvl2) {
                if (lastSubRow) {
                    arranged.push({mock: null})
                    pos++
                }
                arranged.push({mock: null, lvl2: 'TOTAL FOR ' + item.lvl2Name + ' ' + (countLvl.lvl2 === 0 ? stack2 : countLvl.lvl2)})
                countLvl.lvl2 = 0
                stack2 = 0
                levelStyle.lvl2Total.push(pos)
                pos++
                arranged.push({mock: null})
                pos++
                nowLvl = 2
                lastSubRow = false
            }
            if ((!level[0] || (level[0] && level[0] !== tmpItem.lvl1)) && item.lvl1) {
                if (lastSubRow) {
                    arranged.push({mock: null})
                    pos++
                }
                arranged.push({mock: null, lvl1: 'TOTAL FOR ' + item.lvl1Name+ ' ' + countLvl.lvl1})
                countLvl.lvl1 = 0
                stack1 = 0
                levelStyle.lvl1Total.push(pos)
                pos++
                arranged.push({mock: null})
                pos++
                nowLvl = 1
                lastSubRow = false
            }
        } else {
            // last item
            arranged.push({mock: null})
            pos++
            arranged.push({mock: null, lvlx: 'TOTAL FOR ' + item[`lvl${nowLvl}Name`] + ' ' + countLvl[`lvl${nowLvl}`]})
            levelStyle[`lvl${nowLvl}Total`].push(pos)
            pos++
            arranged.push({mock: null})
            pos++
            
            if (nowLvl === 4) {
                // ถ้า 4 สรุป 3,2,1
                arranged.push({mock: null, lvlx: 'TOTAL FOR ' + item[`lvl${nowLvl - 1}Name`] + ' ' + stack3})
                levelStyle[`lvl${nowLvl - 1}Total`].push(pos)
                pos++
                arranged.push({mock: null})
                pos++

                arranged.push({mock: null, lvlx: 'TOTAL FOR ' + item[`lvl${nowLvl - 2}Name`] + ' ' + stack2})
                levelStyle[`lvl${nowLvl - 2}Total`].push(pos)
                pos++
                arranged.push({mock: null})
                pos++

                arranged.push({mock: null, lvlx: 'TOTAL FOR ' + item[`lvl${nowLvl - 3}Name`] + ' ' + stack1})
                levelStyle[`lvl${nowLvl - 3}Total`].push(pos)
                pos++
                arranged.push({mock: null})
                pos++
            }
            if (nowLvl === 3) {
                // ถ้า 3 สรุป 2, 1
                arranged.push({mock: null, lvlx: 'TOTAL FOR ' + item[`lvl${nowLvl - 1}Name`] + ' ' + stack2})
                levelStyle[`lvl${nowLvl - 1}Total`].push(pos)
                pos++
                arranged.push({mock: null})
                pos++

                arranged.push({mock: null, lvlx: 'TOTAL FOR ' + item[`lvl${nowLvl - 2}Name`] + ' ' + stack1})
                levelStyle[`lvl${nowLvl - 2}Total`].push(pos)
                pos++
                arranged.push({mock: null})
                pos++
            }
            if (nowLvl === 2) {
                // ถ้า 2 สรุป 1
                arranged.push({mock: null, lvlx: 'TOTAL FOR ' + item[`lvl${nowLvl - 1}Name`] + ' ' + stack1})
                levelStyle[`lvl${nowLvl - 1}Total`].push(pos)
                pos++
                arranged.push({mock: null})
                pos++
            }
        }
    }
    // Set Grand total
    // arranged.push({mock: null})
    // pos++
    arranged.push({mock: null, lvlx: `GRAND TOTAL FOR CAPEX INVESTMENT Y${budgetYear} ` + countLvl.g })
    levelStyle[`lvlGTotal`].push(pos)

    // Set row data
    r = 4
    for (let i = 0; i < arranged.length; i++) {
        const rawData = arranged[i];
        let rowArray = []
        let columnCount = 0
        for (const key in rawData) {
            if (columnCount < 74) {
                rowArray.push(rawData[key])
            }
            columnCount++
        }
        const row = ws.getRow(r);
        row.values = rowArray;
        r++
    }

    // Set style
    ws.eachRow((row, rowNumber) => {
        if (levelStyle.lvl1.includes(rowNumber - 4)) {
            row.font = {
                name: 'Arial',
                bold: true,
                size: 9,
            };
            row.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFF00' },
            }
        } else if (levelStyle.lvl2.includes(rowNumber - 4)) {
            row.font = {
                name: 'Arial',
                bold: true,
                size: 8,
                color: {argb: 'FF0066 '}
            };
            row.eachCell(function(cell, colNumber) {
                if (colNumber == 1) {
                    cell.alignment = {
                      horizontal: 'right'
                    }
                }
            });
        } else if (levelStyle.lvl3.includes(rowNumber - 4)) {
            row.font = {
                name: 'Arial',
                bold: true,
                size: 8,
            };
            row.eachCell(function(cell, colNumber) {
                if (colNumber == 1) {
                    cell.alignment = {
                      horizontal: 'right'
                    }
                }
            });
        } else if (levelStyle.lvl4.includes(rowNumber - 4)) {
            row.font = {
                name: 'Arial',
                bold: true,
                size: 8,
                color: {argb: '0000FF '}
            };
            row.eachCell(function(cell, colNumber) {
                if (colNumber == 1) {
                    cell.alignment = {
                      horizontal: 'right'
                    }
                }
            });
        } else if (levelStyle.lvl1Total.includes(rowNumber - 4)) {
            row.font = {
                name: 'Arial',
                bold: true,
                size: 9,
                color: {argb: 'C00000'}
            };
        } else if (levelStyle.lvl2Total.includes(rowNumber - 4)) {
            row.font = {
                name: 'Arial',
                bold: true,
                size: 9,
                color: {argb: '00B050'}
            };
        } else if (levelStyle.lvl3Total.includes(rowNumber - 4)) {
            row.font = {
                name: 'Arial',
                bold: true,
                size: 8,
                color: {argb: 'C00000'}
            };
        } else if (levelStyle.lvl4Total.includes(rowNumber - 4)) {
            row.font = {
                name: 'Arial',
                bold: true,
                size: 8,
                color: {argb: '0000FF '}
            };
        } else if (levelStyle.lvlGTotal.includes(rowNumber - 4)) {
            row.font = {
                name: 'Arial',
                bold: true,
                size: 9,
                color: {argb: '000000'}
            };
            row.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '92D050' },
            }
        } else {
            if (rowNumber > 3) {
                row.font = {
                    name: 'Arial',
                    bold: false,
                    size: 8,
                };
            }
            
        }
    });
    
    // write file
    let deapartmentOwner = contactPointDepartment.replaceAll(/[<>:"\/\\|?*]+/g, '_')
    let writeFileName = `Template Budget CAPEX ${budgetYear} - ${deapartmentOwner}.xlsx`

    const writeFilePromise = new Promise((resolve, reject) => {
        wb.xlsx.writeFile(writeFileName)
        .then(() => {
            resolve({status: true, message: 'File Created Successfully'});
        })
        .catch(err => {
            resolve({status: false, message: err.message});
            
        });
    });
    console.log('Saving file...')
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