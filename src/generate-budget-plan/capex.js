const Excel = require('exceljs')
const moment = require('dayjs');
const chalk = require('chalk');
const path = require('path')
const mkdirp = require('mkdirp')


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
        console.log('');
        console.log(chalk.white.bgRed.bold(' budgetYear required '));
        console.log('');
        return false;
    }
    
    if (!dateTime) {
        console.log('');
        console.log(chalk.white.bgRed.bold(' dateTime required '));
        console.log('');
        return false;
    }

    if (!contactPointDepartment) {
        console.log('');
        console.log(chalk.white.bgRed.bold(' contactPointDepartment required '));
        console.log('');
        return false;
    }

    const isbudgetYear = new RegExp('^[0-9]{4,4}$');
    if (!isbudgetYear.test(budgetYear)) {
        console.log('');
        console.log(chalk.white.bgRed.bold(' [budgetYear] - Should be in year format - Ex: 2022 '));
        console.log('');
        console.log('Should be in datetime format: 2022');
        return false
    }
    const isdateTime = new RegExp('^[0-9]{8,8}$');
    if (!isdateTime.test(dateTime)) {
        console.log('');
        console.log(chalk.white.bgRed.bold(' [dateTime] - Should be in datetime format - Ex: 20220101 '));
        console.log('');
        return false
    }

    if (dateTime && dateTime.indexOf('-') === -1) {
        dateTime = moment(moment(dateTime).format('YYYYMMDD')).format('YYYY-MM-DD');
    }

    // get team data
    let teamList = []
    if (contactPointDepartment === 'TEAM') {
        teamList = await knex('planBudgetCapex').select(knex.raw(`distinct(nvl(contact_Point_Department, 'NULL')) contact_Point_Department`)).where('dateTime', knex.raw(`to_date('${dateTime}', 'YYYY-MM-DD')`)).where('budgetYear', budgetYear).then(item => item = item.map(e => e = e.contactPointDepartment))
        console.log('teamList', teamList)
    } else {
        teamList.push(contactPointDepartment)
    }

    let conclusion = []
    for (let i = 0; i < teamList.length; i++) {
        const team = teamList[i];
        let res = await generateMaster(budgetYear, dateTime, team)
        conclusion.push(res)
    }

    if (conclusion.filter(e => !e.status).length) {
        console.log('');
        console.log('');
        console.log(chalk.white.bgBlue.bold(' FAILED CONCLUSION '));
        console.log('');
    }
    for (let i = 0; i < conclusion.length; i++) {
        const item = conclusion[i];
        if (!item.status) {
            console.log(chalk.white.bgRed.bold(' FAILED '), `[${item.team}] -  ERROR: ${item.message}`);
        }
        // else {
        //     console.log(chalk.white.bgGreen.bold(' Generate ' + item.team + ' Successfully '));
        // }
    }
}

async function generateMaster(budgetYear, dateTime, contactPointDepartment) {
    console.log('');
    console.log('TEAM:', contactPointDepartment)

    // import template for header
    let fileName = './src/generate-budget-plan/templates/TemplateCAPEX.xlsx'
    let wb = new Excel.Workbook();
    let ws
    await wb.xlsx.readFile(fileName).then(function() {
        ws = wb.getWorksheet('Sheet1');
    });

    
    // get data
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
        NULL,NULL,budget_name,busines_function,business_group,business_owner_name,sub_project_product,project_priority,budget_company,cost_center,contact_point_budget_own_name,contact_point_department,contact_point_mobile,
        NVL(budget_amount_usd, 0) budget_amount_usd,NVL(budget_amount_thb, 0) budget_amount_thb,NVL(equivalent_to_thb, 0) equivalent_to_thb,NVL(equivalent_to_usd, 0) equivalent_to_usd,
        assumption_budget_calculation,project_description,
        NVL(hardware_total, 0) hardware_total,NVL(software_license_total, 0) software_license_total,NVL(software_dev_turnkey_total, 0) software_dev_turnkey_total,NVL(software_dev_dbp_total, 0) software_dev_dbp_total,
        NVL(software_dev_automate_total, 0) software_dev_automate_total,NVL(external_outsourcing_total, 0) external_outsourcing_total,NVL(outsource_total_total, 0) outsource_total_total,
        NVL(outsource_pm_si_total, 0) outsource_pm_si_total,NVL(outsource_sa_total, 0) outsource_sa_total,NVL(outsource_pa_total, 0) outsource_pa_total,NVL(outsource_tester_total, 0) outsource_tester_total,
        NVL(outsource_ts_total, 0) outsource_ts_total,NVL(outsource_tc_total, 0) outsource_tc_total,NULL,NVL(plan_use_jan, 0) plan_use_jan,NVL(plan_use_feb, 0) plan_use_feb,
        NVL(plan_use_mar, 0) plan_use_mar,NVL(plan_use_apr, 0) plan_use_apr,NVL(plan_use_may, 0) plan_use_may,NVL(plan_use_jun, 0) plan_use_jun,NVL(plan_use_jul, 0) plan_use_jul,NVL(plan_use_aug, 0) plan_use_aug,
        NVL(plan_use_sep, 0) plan_use_sep,NVL(plan_use_oct, 0) plan_use_oct,NVL(plan_use_nov, 0) plan_use_nov,NVL(plan_use_dec, 0) plan_use_dec,NVL(plan_use_total, 0) plan_use_total,NULL,
        NVL(forecast_inv_cur_jan, 0) forecast_inv_cur_jan,NVL(forecast_inv_cur_feb, 0) forecast_inv_cur_feb,NVL(forecast_inv_cur_mar, 0) forecast_inv_cur_mar,NVL(forecast_inv_cur_apr, 0) forecast_inv_cur_apr,
        NVL(forecast_inv_cur_may, 0) forecast_inv_cur_may,NVL(forecast_inv_cur_jun, 0) forecast_inv_cur_jun,NVL(forecast_inv_cur_jul, 0) forecast_inv_cur_jul,NVL(forecast_inv_cur_aug, 0) forecast_inv_cur_aug,
        NVL(forecast_inv_cur_sep, 0) forecast_inv_cur_sep,NVL(forecast_inv_cur_oct, 0) forecast_inv_cur_oct,NVL(forecast_inv_cur_nov, 0) forecast_inv_cur_nov,NVL(forecast_inv_cur_dec, 0) forecast_inv_cur_dec,
        NVL(forecast_inv_cur_total, 0) forecast_inv_cur_total,NULL,NVL(forecast_inv_next_jan, 0) forecast_inv_next_jan,NVL(forecast_inv_next_feb, 0) forecast_inv_next_feb,
        NVL(forecast_inv_next_mar, 0) forecast_inv_next_mar,NVL(forecast_inv_next_apr, 0) forecast_inv_next_apr,NVL(forecast_inv_next_may, 0) forecast_inv_next_may,NVL(forecast_inv_next_jun, 0) forecast_inv_next_jun,
        NVL(forecast_inv_next_jul, 0) forecast_inv_next_jul,NVL(forecast_inv_next_aug, 0) forecast_inv_next_aug,NVL(forecast_inv_next_sep, 0) forecast_inv_next_sep,NVL(forecast_inv_next_oct, 0) forecast_inv_next_oct,
        NVL(forecast_inv_next_nov, 0) forecast_inv_next_nov,NVL(forecast_inv_next_dec, 0) forecast_inv_next_dec,NVL(forecast_inv_next_total, 0) forecast_inv_next_total,
        CASE WHEN investment_Level_1 IS NOT NULL THEN investment_Level_1 ELSE '99. NULL' END lvl1,CASE WHEN investment_Level_2 IS NOT NULL THEN investment_Level_2 ELSE '99. NULL' END lvl2,
        investment_Level_3 lvl3,investment_Level_4 lvl4, date_time
    from gen
    where budget_year = '${budgetYear}' and date_time = to_date('${dateTime}', 'YYYY-MM-DD')
            ${contactPointDepartment === 'ALL' ? '' : `and CONTACT_POINT_DEPARTMENT ${contactPointDepartment === 'NULL' ? 'IS NULL' : `= '${contactPointDepartment}'`}`}
    order by neworder1, neworder2, neworder3, neworder4
    `)

    // get total data
    let totalRows = await knex.raw(`
        with main as (
            select
                TRANS_DATE, DATE_TIME, FILE_NAME, BUDGET_YEAR, BUSINESS_OBJECTIVE, BUSINESS_STRATEGIC, BUDGET_NAME, NVL(BUDGET_AMOUNT_THB, 0) BUDGET_AMOUNT_THB, NVL(BUDGET_AMOUNT_USD, 0) BUDGET_AMOUNT_USD, NVL(EQUIVALENT_TO_THB, 0) EQUIVALENT_TO_THB, NVL(EQUIVALENT_TO_USD, 0) EQUIVALENT_TO_USD, C_LEVEL_TEAM, BUDGET_FLAG,
                CASE WHEN investment_Level_1 IS NOT NULL THEN investment_Level_1 ELSE '99. NULL' END investment_Level_1,CASE WHEN investment_Level_2 IS NOT NULL THEN investment_Level_2 ELSE '99. NULL' END investment_Level_2,
                INVESTMENT_LEVEL_3, INVESTMENT_LEVEL_4, BUSINES_FUNCTION, BUSINESS_GROUP, BUSINESS_OWNER_NAME, CATEGORY, SUB_PROJECT_PRODUCT, PROJECT_PRIORITY, BUDGET_COMPANY, COST_CENTER, CONTACT_POINT_BUDGET_OWN_NAME,
                CONTACT_POINT_DEPARTMENT, CONTACT_POINT_MOBILE, TYPE_ABC, ASSUMPTION_BUDGET_CALCULATION, PROJECT_DESCRIPTION, HARDWARE_USD, HARDWARE_THB, SOFTWARE_LICENSE_USD, SOFTWARE_LICENSE_THB, SOFTWARE_DEV_TURNKEY_USD,
                SOFTWARE_DEV_TURNKEY_THB, SOFTWARE_DEV_DBP_USD, SOFTWARE_DEV_DBP_THB, SOFTWARE_DEV_AUTOMATE_USD, SOFTWARE_DEV_AUTOMATE_THB, EXTERNAL_OUTSOURCING_USD, EXTERNAL_OUTSOURCING_THB, TOTAL_HW_SW_DBP_TURNKEY_USD,
                TOTAL_HW_SW_DBP_TURNKEY_THB, OUTSOURCE_TOTAL_USD, OUTSOURCE_TOTAL_THB, OUTSOURCE_PM_SI_USD, OUTSOURCE_PM_SI_THB, OUTSOURCE_SA_USD, OUTSOURCE_SA_THB, OUTSOURCE_PA_USD, OUTSOURCE_PA_THB, OUTSOURCE_TESTER_USD,
                OUTSOURCE_TESTER_THB, OUTSOURCE_TS_USD, OUTSOURCE_TS_THB, OUTSOURCE_TC_USD, OUTSOURCE_TC_THB, TOTAL_OUTSOURCE, NVL(PLAN_USE_JAN, 0) PLAN_USE_JAN, NVL(PLAN_USE_FEB, 0) PLAN_USE_FEB, NVL(PLAN_USE_MAR, 0) PLAN_USE_MAR, NVL(PLAN_USE_APR, 0) PLAN_USE_APR, NVL(PLAN_USE_MAY, 0) PLAN_USE_MAY, NVL(PLAN_USE_JUN, 0) PLAN_USE_JUN, NVL(PLAN_USE_JUL, 0) PLAN_USE_JUL,
                NVL(PLAN_USE_AUG, 0) PLAN_USE_AUG, NVL(PLAN_USE_SEP, 0) PLAN_USE_SEP, NVL(PLAN_USE_OCT, 0) PLAN_USE_OCT, NVL(PLAN_USE_NOV, 0) PLAN_USE_NOV, NVL(PLAN_USE_DEC, 0) PLAN_USE_DEC, NVL(PLAN_USE_TOTAL, 0) PLAN_USE_TOTAL, NVL(FORECAST_INV_CUR_JAN, 0) FORECAST_INV_CUR_JAN, NVL(FORECAST_INV_CUR_FEB, 0) FORECAST_INV_CUR_FEB, NVL(FORECAST_INV_CUR_MAR, 0) FORECAST_INV_CUR_MAR, NVL(FORECAST_INV_CUR_APR, 0) FORECAST_INV_CUR_APR, NVL(FORECAST_INV_CUR_MAY, 0) FORECAST_INV_CUR_MAY,
                NVL(FORECAST_INV_CUR_JUN, 0) FORECAST_INV_CUR_JUN, NVL(FORECAST_INV_CUR_JUL, 0) FORECAST_INV_CUR_JUL, NVL(FORECAST_INV_CUR_AUG, 0) FORECAST_INV_CUR_AUG, NVL(FORECAST_INV_CUR_SEP, 0) FORECAST_INV_CUR_SEP, NVL(FORECAST_INV_CUR_OCT, 0) FORECAST_INV_CUR_OCT, NVL(FORECAST_INV_CUR_NOV, 0) FORECAST_INV_CUR_NOV, NVL(FORECAST_INV_CUR_DEC, 0) FORECAST_INV_CUR_DEC, NVL(FORECAST_INV_CUR_TOTAL, 0) FORECAST_INV_CUR_TOTAL, NVL(FORECAST_INV_NEXT_JAN, 0) FORECAST_INV_NEXT_JAN,
                NVL(FORECAST_INV_NEXT_FEB, 0) FORECAST_INV_NEXT_FEB, NVL(FORECAST_INV_NEXT_MAR, 0) FORECAST_INV_NEXT_MAR, NVL(FORECAST_INV_NEXT_APR, 0) FORECAST_INV_NEXT_APR, NVL(FORECAST_INV_NEXT_MAY, 0) FORECAST_INV_NEXT_MAY, NVL(FORECAST_INV_NEXT_JUN, 0) FORECAST_INV_NEXT_JUN, NVL(FORECAST_INV_NEXT_JUL, 0) FORECAST_INV_NEXT_JUL, NVL(FORECAST_INV_NEXT_AUG, 0) FORECAST_INV_NEXT_AUG, NVL(FORECAST_INV_NEXT_SEP, 0) FORECAST_INV_NEXT_SEP, NVL(FORECAST_INV_NEXT_OCT, 0) FORECAST_INV_NEXT_OCT,
                NVL(FORECAST_INV_NEXT_NOV, 0) FORECAST_INV_NEXT_NOV, NVL(FORECAST_INV_NEXT_DEC, 0) FORECAST_INV_NEXT_DEC, NVL(FORECAST_INV_NEXT_TOTAL, 0) FORECAST_INV_NEXT_TOTAL, NVL(FORECAST_INV_TOTAL, 0) FORECAST_INV_TOTAL, INVOICE_FORECAST_2024, ASSUMPTION_2024, INVOICE_FORECAST_2025, ASSUMPTION_2025, CARRY_NEXT_YEAR_FLAG,
                CARRY_NEXT_YEAR_THB, CARRY_NEXT_YEAR_USD, CARRY_NEXT_YEAR_DATE, CARRY_NEXT_YEAR_BY, REVIEW_STATUS, REVIEW_DATE, REVIEW_TIME_NO, REVISE00_PO_AMOUNT_THB, REVISE00_PO_AMOUNT_USD, REVISE01_PO_AMOUNT_DATE,
                REVISE01_PO_AMOUNT_BY, REVISE01_PO_AMOUNT_THB, REVISE01_PO_AMOUNT_USD, REVISE02_PO_AMOUNT_DATE, REVISE02_PO_AMOUNT_BY, REVISE02_PO_AMOUNT_THB, REVISE02_PO_AMOUNT_USD, REVISE03_PO_AMOUNT_DATE,
                REVISE03_PO_AMOUNT_BY, REVISE03_PO_AMOUNT_THB, REVISE03_PO_AMOUNT_USD, REVISE04_PO_AMOUNT_DATE, REVISE04_PO_AMOUNT_BY, REVISE04_PO_AMOUNT_THB, REVISE04_PO_AMOUNT_USD, REVISE05_PO_AMOUNT_DATE,
                REVISE05_PO_AMOUNT_BY, REVISE05_PO_AMOUNT_THB, REVISE05_PO_AMOUNT_USD, BUDGET_CATEGORY, PO_VALUE_UNIT, PO_VALUE_AMOUNT_THB, PO_VALUE_AMOUNT_USD, BUSINESS_OWNER, BUSINESS_OBJECTIVE_OLD, BUSINESS_STRATEGIC_OLD,
                CREATED, CREATED_BY, MODIFIED, MODIFIED_BY, ID, ITEM_TYPE, Path, NVL(HARDWARE_TOTAL, 0) HARDWARE_TOTAL, NVL(SOFTWARE_LICENSE_TOTAL, 0) SOFTWARE_LICENSE_TOTAL, NVL(SOFTWARE_DEV_TURNKEY_TOTAL, 0) SOFTWARE_DEV_TURNKEY_TOTAL, NVL(SOFTWARE_DEV_DBP_TOTAL, 0) SOFTWARE_DEV_DBP_TOTAL, NVL(SOFTWARE_DEV_AUTOMATE_TOTAL, 0) SOFTWARE_DEV_AUTOMATE_TOTAL,
                NVL(EXTERNAL_OUTSOURCING_TOTAL, 0) EXTERNAL_OUTSOURCING_TOTAL, NVL(TOTAL_HW_SW_DBP_TURNKEY_TOTAL, 0) TOTAL_HW_SW_DBP_TURNKEY_TOTAL, NVL(OUTSOURCE_TOTAL_TOTAL, 0) OUTSOURCE_TOTAL_TOTAL, NVL(OUTSOURCE_PM_SI_TOTAL, 0) OUTSOURCE_PM_SI_TOTAL, NVL(OUTSOURCE_SA_TOTAL, 0) OUTSOURCE_SA_TOTAL, NVL(OUTSOURCE_PA_TOTAL, 0) OUTSOURCE_PA_TOTAL, NVL(OUTSOURCE_TESTER_TOTAL, 0) OUTSOURCE_TESTER_TOTAL, NVL(OUTSOURCE_TS_TOTAL, 0) OUTSOURCE_TS_TOTAL, NVL(OUTSOURCE_TC_TOTAL, 0) OUTSOURCE_TC_TOTAL
            from plan_Budget_Capex
            where budget_year = '${budgetYear}' and date_time = to_date('${dateTime}', 'YYYY-MM-DD')
            ${contactPointDepartment === 'ALL' ? '' : `and CONTACT_POINT_DEPARTMENT ${contactPointDepartment === 'NULL' ? 'IS NULL' : `= '${contactPointDepartment}'`}`}
        )
        SELECT 
            null,
            'TOTAL FOR' || REPLACE(SUBSTR(investment_Level_4, Instr(investment_Level_4, ' ')), '  ',' ') lvl_name, null, null,null,null,null,null,null,null,null,null,null,
            SUM(BUDGET_AMOUNT_USD) BUDGET_AMOUNT_USD, SUM(BUDGET_AMOUNT_THB) BUDGET_AMOUNT_THB, SUM(EQUIVALENT_TO_THB) EQUIVALENT_TO_THB, SUM(EQUIVALENT_TO_THB) / 34.78 / 1000000 AS EQUIVALENT_TO_USD, null,null,
            SUM(hardware_total) hardware_total, SUM(software_license_total) software_license_total, SUM(software_dev_turnkey_total) software_dev_turnkey_total, SUM(software_dev_dbp_total) software_dev_dbp_total, 
            SUM(software_dev_automate_total) software_dev_automate_total, SUM(external_outsourcing_total) external_outsourcing_total, SUM(outsource_total_total) outsource_total_total,SUM(outsource_pm_si_total) outsource_pm_si_total, SUM(outsource_sa_total) outsource_sa_total, SUM(outsource_pa_total) outsource_pa_total, SUM(outsource_tester_total) outsource_tester_total, SUM(outsource_ts_total) outsource_ts_total, SUM(outsource_tc_total) outsource_tc_total, null,SUM(plan_use_jan) plan_use_jan,SUM(plan_use_feb) plan_use_feb,SUM(plan_use_mar) plan_use_mar,SUM(plan_use_apr) plan_use_apr,SUM(plan_use_may) plan_use_may,SUM(plan_use_jun) plan_use_jun,SUM(plan_use_jul) plan_use_jul,SUM(plan_use_aug) plan_use_aug,SUM(plan_use_sep) plan_use_sep,SUM(plan_use_oct) plan_use_oct,SUM(plan_use_nov) plan_use_nov,SUM(plan_use_dec) plan_use_dec,SUM(plan_use_total) plan_use_total, null, SUM(forecast_inv_cur_jan) forecast_inv_cur_jan, SUM(forecast_inv_cur_feb) forecast_inv_cur_feb, SUM(forecast_inv_cur_mar) forecast_inv_cur_mar, SUM(forecast_inv_cur_apr) forecast_inv_cur_apr, SUM(forecast_inv_cur_may) forecast_inv_cur_may, SUM(forecast_inv_cur_jun) forecast_inv_cur_jun, SUM(forecast_inv_cur_jul) forecast_inv_cur_jul, SUM(forecast_inv_cur_aug) forecast_inv_cur_aug, SUM(forecast_inv_cur_sep) forecast_inv_cur_sep, SUM(forecast_inv_cur_oct) forecast_inv_cur_oct, SUM(forecast_inv_cur_nov) forecast_inv_cur_nov, SUM(forecast_inv_cur_dec) forecast_inv_cur_dec, SUM(forecast_inv_cur_total) forecast_inv_cur_total, null, SUM(forecast_inv_next_jan) forecast_inv_next_jan, SUM(forecast_inv_next_feb) forecast_inv_next_feb, SUM(forecast_inv_next_mar) forecast_inv_next_mar, SUM(forecast_inv_next_apr) forecast_inv_next_apr, SUM(forecast_inv_next_may) forecast_inv_next_may, SUM(forecast_inv_next_jun) forecast_inv_next_jun, SUM(forecast_inv_next_jul) forecast_inv_next_jul, SUM(forecast_inv_next_aug) forecast_inv_next_aug, SUM(forecast_inv_next_sep) forecast_inv_next_sep, SUM(forecast_inv_next_oct) forecast_inv_next_oct, SUM(forecast_inv_next_nov) forecast_inv_next_nov, SUM(forecast_inv_next_dec) forecast_inv_next_dec, SUM(forecast_inv_next_total) forecast_inv_next_total, (REPLACE(SUBSTR(investment_Level_4, 0, Instr(investment_Level_4, ' ')), '')) lvl_No
        FROM main
        WHERE investment_Level_1 IS NOT NULL AND investment_Level_2 IS NOT NULL AND investment_Level_3 IS NOT NULL AND investment_Level_4 IS NOT NULL
        GROUP BY investment_Level_1, investment_Level_2, investment_Level_3, investment_Level_4
        
        UNION
        
        SELECT
            null,
            'TOTAL FOR' || REPLACE(SUBSTR(investment_Level_3, Instr(investment_Level_3, ' ')), '  ',' ') lvl_name, null, null,null,null,null,null,null,null,null,null,null,
            SUM(BUDGET_AMOUNT_USD) BUDGET_AMOUNT_USD, SUM(BUDGET_AMOUNT_THB) BUDGET_AMOUNT_THB, SUM(EQUIVALENT_TO_THB) EQUIVALENT_TO_THB, SUM(EQUIVALENT_TO_THB) / 34.78 / 1000000 AS EQUIVALENT_TO_USD, null,null,
            SUM(hardware_total) hardware_total, SUM(software_license_total) software_license_total, SUM(software_dev_turnkey_total) software_dev_turnkey_total, SUM(software_dev_dbp_total) software_dev_dbp_total, 
            SUM(software_dev_automate_total) software_dev_automate_total, SUM(external_outsourcing_total) external_outsourcing_total, SUM(outsource_total_total) outsource_total_total, SUM(outsource_pm_si_total) outsource_pm_si_total, SUM(outsource_sa_total) outsource_sa_total, SUM(outsource_pa_total) outsource_pa_total, SUM(outsource_tester_total) outsource_tester_total, SUM(outsource_ts_total) outsource_ts_total, SUM(outsource_tc_total) outsource_tc_total, null,SUM(plan_use_jan) plan_use_jan,SUM(plan_use_feb) plan_use_feb,SUM(plan_use_mar) plan_use_mar,SUM(plan_use_apr) plan_use_apr,SUM(plan_use_may) plan_use_may,SUM(plan_use_jun) plan_use_jun,SUM(plan_use_jul) plan_use_jul,SUM(plan_use_aug) plan_use_aug,SUM(plan_use_sep) plan_use_sep,SUM(plan_use_oct) plan_use_oct,SUM(plan_use_nov) plan_use_nov,SUM(plan_use_dec) plan_use_dec,SUM(plan_use_total) plan_use_total, null, SUM(forecast_inv_cur_jan) forecast_inv_cur_jan, SUM(forecast_inv_cur_feb) forecast_inv_cur_feb, SUM(forecast_inv_cur_mar) forecast_inv_cur_mar, SUM(forecast_inv_cur_apr) forecast_inv_cur_apr, SUM(forecast_inv_cur_may) forecast_inv_cur_may, SUM(forecast_inv_cur_jun) forecast_inv_cur_jun, SUM(forecast_inv_cur_jul) forecast_inv_cur_jul, SUM(forecast_inv_cur_aug) forecast_inv_cur_aug, SUM(forecast_inv_cur_sep) forecast_inv_cur_sep, SUM(forecast_inv_cur_oct) forecast_inv_cur_oct, SUM(forecast_inv_cur_nov) forecast_inv_cur_nov, SUM(forecast_inv_cur_dec) forecast_inv_cur_dec, SUM(forecast_inv_cur_total) forecast_inv_cur_total, null, SUM(forecast_inv_next_jan) forecast_inv_next_jan, SUM(forecast_inv_next_feb) forecast_inv_next_feb, SUM(forecast_inv_next_mar) forecast_inv_next_mar, SUM(forecast_inv_next_apr) forecast_inv_next_apr, SUM(forecast_inv_next_may) forecast_inv_next_may, SUM(forecast_inv_next_jun) forecast_inv_next_jun, SUM(forecast_inv_next_jul) forecast_inv_next_jul, SUM(forecast_inv_next_aug) forecast_inv_next_aug, SUM(forecast_inv_next_sep) forecast_inv_next_sep, SUM(forecast_inv_next_oct) forecast_inv_next_oct, SUM(forecast_inv_next_nov) forecast_inv_next_nov, SUM(forecast_inv_next_dec) forecast_inv_next_dec, SUM(forecast_inv_next_total) forecast_inv_next_total, (REPLACE(SUBSTR(investment_Level_3, 0, Instr(investment_Level_3, ' ')), '')) lvl_No
        FROM main
        WHERE investment_Level_1 IS NOT NULL AND investment_Level_2 IS NOT NULL AND investment_Level_3 IS NOT NULL
        GROUP BY investment_Level_1, investment_Level_2, investment_Level_3
        
        UNION
        
        SELECT
            null,
            'TOTAL FOR' || REPLACE(SUBSTR(investment_Level_2,Instr(investment_Level_2, ' ')), '  ',' ') lvl_name, null,null,null,null,null,null,null,null,null,null,null,
            SUM(BUDGET_AMOUNT_USD) BUDGET_AMOUNT_USD,SUM(BUDGET_AMOUNT_THB) BUDGET_AMOUNT_THB,SUM(EQUIVALENT_TO_THB) EQUIVALENT_TO_THB,SUM(EQUIVALENT_TO_THB) / 34.78 / 1000000 AS EQUIVALENT_TO_USD,null,null,
            SUM(hardware_total) hardware_total,SUM(software_license_total) software_license_total,SUM(software_dev_turnkey_total) software_dev_turnkey_total,SUM(software_dev_dbp_total) software_dev_dbp_total,SUM(software_dev_automate_total) software_dev_automate_total,SUM(external_outsourcing_total) external_outsourcing_total,SUM(outsource_total_total) outsource_total_total,SUM(outsource_pm_si_total) outsource_pm_si_total,SUM(outsource_sa_total) outsource_sa_total,SUM(outsource_pa_total) outsource_pa_total,SUM(outsource_tester_total) outsource_tester_total,SUM(outsource_ts_total) outsource_ts_total,SUM(outsource_tc_total) outsource_tc_total,null,SUM(plan_use_jan) plan_use_jan,SUM(plan_use_feb) plan_use_feb,SUM(plan_use_mar) plan_use_mar,SUM(plan_use_apr) plan_use_apr,SUM(plan_use_may) plan_use_may,SUM(plan_use_jun) plan_use_jun,SUM(plan_use_jul) plan_use_jul,SUM(plan_use_aug) plan_use_aug,SUM(plan_use_sep) plan_use_sep,SUM(plan_use_oct) plan_use_oct,SUM(plan_use_nov) plan_use_nov,SUM(plan_use_dec) plan_use_dec,SUM(plan_use_total) plan_use_total, null,SUM(forecast_inv_cur_jan) forecast_inv_cur_jan,SUM(forecast_inv_cur_feb) forecast_inv_cur_feb,SUM(forecast_inv_cur_mar) forecast_inv_cur_mar,SUM(forecast_inv_cur_apr) forecast_inv_cur_apr,SUM(forecast_inv_cur_may) forecast_inv_cur_may,SUM(forecast_inv_cur_jun) forecast_inv_cur_jun,SUM(forecast_inv_cur_jul) forecast_inv_cur_jul,SUM(forecast_inv_cur_aug) forecast_inv_cur_aug,SUM(forecast_inv_cur_sep) forecast_inv_cur_sep,SUM(forecast_inv_cur_oct) forecast_inv_cur_oct,SUM(forecast_inv_cur_nov) forecast_inv_cur_nov,SUM(forecast_inv_cur_dec) forecast_inv_cur_dec,SUM(forecast_inv_cur_total) forecast_inv_cur_total,null,SUM(forecast_inv_next_jan) forecast_inv_next_jan,SUM(forecast_inv_next_feb) forecast_inv_next_feb,SUM(forecast_inv_next_mar) forecast_inv_next_mar,SUM(forecast_inv_next_apr) forecast_inv_next_apr,SUM(forecast_inv_next_may) forecast_inv_next_may,SUM(forecast_inv_next_jun) forecast_inv_next_jun,SUM(forecast_inv_next_jul) forecast_inv_next_jul,SUM(forecast_inv_next_aug) forecast_inv_next_aug,SUM(forecast_inv_next_sep) forecast_inv_next_sep,SUM(forecast_inv_next_oct) forecast_inv_next_oct,SUM(forecast_inv_next_nov) forecast_inv_next_nov,SUM(forecast_inv_next_dec) forecast_inv_next_dec,SUM(forecast_inv_next_total) forecast_inv_next_total, (REPLACE(SUBSTR(investment_Level_2, 0, Instr(investment_Level_2, ' ')), '')) lvl_No
        FROM main
        WHERE investment_Level_1 IS NOT NULL AND investment_Level_2 IS NOT NULL
        GROUP BY investment_Level_1, investment_Level_2
        
        UNION
        
        SELECT
            null,
            'TOTAL FOR' || REPLACE(SUBSTR(investment_Level_1,Instr(investment_Level_1, ' ')), '  ',' ') lvl_name, null, null,null,null,null,null,null,null,null,null,null,
            SUM(BUDGET_AMOUNT_USD) BUDGET_AMOUNT_USD, SUM(BUDGET_AMOUNT_THB) BUDGET_AMOUNT_THB, SUM(EQUIVALENT_TO_THB) EQUIVALENT_TO_THB, SUM(EQUIVALENT_TO_THB) / 34.78 / 1000000 AS EQUIVALENT_TO_USD, null,null,
            SUM(hardware_total) hardware_total, SUM(software_license_total) software_license_total, SUM(software_dev_turnkey_total) software_dev_turnkey_total, SUM(software_dev_dbp_total) software_dev_dbp_total, SUM(software_dev_automate_total) software_dev_automate_total, SUM(external_outsourcing_total) external_outsourcing_total, SUM(outsource_total_total) outsource_total_total, SUM(outsource_pm_si_total) outsource_pm_si_total, SUM(outsource_sa_total) outsource_sa_total, SUM(outsource_pa_total) outsource_pa_total, SUM(outsource_tester_total) outsource_tester_total, SUM(outsource_ts_total) outsource_ts_total, SUM(outsource_tc_total) outsource_tc_total, null,SUM(plan_use_jan) plan_use_jan,SUM(plan_use_feb) plan_use_feb,SUM(plan_use_mar) plan_use_mar,SUM(plan_use_apr) plan_use_apr,SUM(plan_use_may) plan_use_may,SUM(plan_use_jun) plan_use_jun,SUM(plan_use_jul) plan_use_jul,SUM(plan_use_aug) plan_use_aug,SUM(plan_use_sep) plan_use_sep,SUM(plan_use_oct) plan_use_oct,SUM(plan_use_nov) plan_use_nov,SUM(plan_use_dec) plan_use_dec,SUM(plan_use_total) plan_use_total, null, SUM(forecast_inv_cur_jan) forecast_inv_cur_jan, SUM(forecast_inv_cur_feb) forecast_inv_cur_feb, SUM(forecast_inv_cur_mar) forecast_inv_cur_mar, SUM(forecast_inv_cur_apr) forecast_inv_cur_apr, SUM(forecast_inv_cur_may) forecast_inv_cur_may, SUM(forecast_inv_cur_jun) forecast_inv_cur_jun, SUM(forecast_inv_cur_jul) forecast_inv_cur_jul, SUM(forecast_inv_cur_aug) forecast_inv_cur_aug, SUM(forecast_inv_cur_sep) forecast_inv_cur_sep, SUM(forecast_inv_cur_oct) forecast_inv_cur_oct, SUM(forecast_inv_cur_nov) forecast_inv_cur_nov, SUM(forecast_inv_cur_dec) forecast_inv_cur_dec, SUM(forecast_inv_cur_total) forecast_inv_cur_total, null, SUM(forecast_inv_next_jan) forecast_inv_next_jan, SUM(forecast_inv_next_feb) forecast_inv_next_feb, SUM(forecast_inv_next_mar) forecast_inv_next_mar, SUM(forecast_inv_next_apr) forecast_inv_next_apr, SUM(forecast_inv_next_may) forecast_inv_next_may, SUM(forecast_inv_next_jun) forecast_inv_next_jun, SUM(forecast_inv_next_jul) forecast_inv_next_jul, SUM(forecast_inv_next_aug) forecast_inv_next_aug, SUM(forecast_inv_next_sep) forecast_inv_next_sep, SUM(forecast_inv_next_oct) forecast_inv_next_oct, SUM(forecast_inv_next_nov) forecast_inv_next_nov, SUM(forecast_inv_next_dec) forecast_inv_next_dec, SUM(forecast_inv_next_total) forecast_inv_next_total, (REPLACE(SUBSTR(investment_Level_1, 0, Instr(investment_Level_1, ' ')), '')) lvl_No
        FROM main
        WHERE investment_Level_1 IS NOT NULL
        GROUP BY investment_Level_1
        
        UNION
        
        SELECT
            null, 'GRAND TOTAL FOR CAPEX INVESTMENT Y${budgetYear}' lvl_name, null, null,null,null,null,null,null,null,null,null,null, SUM(BUDGET_AMOUNT_USD) BUDGET_AMOUNT_USD, SUM(BUDGET_AMOUNT_THB) BUDGET_AMOUNT_THB,
            SUM(EQUIVALENT_TO_THB) EQUIVALENT_TO_THB, SUM(EQUIVALENT_TO_THB) / 34.78 / 1000000 AS EQUIVALENT_TO_USD, null,null, SUM(hardware_total) hardware_total, SUM(software_license_total) software_license_total, SUM(software_dev_turnkey_total) software_dev_turnkey_total, SUM(software_dev_dbp_total) software_dev_dbp_total, SUM(software_dev_automate_total) software_dev_automate_total, SUM(external_outsourcing_total) external_outsourcing_total, SUM(outsource_total_total) outsource_total_total, SUM(outsource_pm_si_total) outsource_pm_si_total, SUM(outsource_sa_total) outsource_sa_total, SUM(outsource_pa_total) outsource_pa_total, SUM(outsource_tester_total) outsource_tester_total, SUM(outsource_ts_total) outsource_ts_total, SUM(outsource_tc_total) outsource_tc_total, null,SUM(plan_use_jan) plan_use_jan,SUM(plan_use_feb) plan_use_feb,SUM(plan_use_mar) plan_use_mar,SUM(plan_use_apr) plan_use_apr,SUM(plan_use_may) plan_use_may,SUM(plan_use_jun) plan_use_jun,SUM(plan_use_jul) plan_use_jul,SUM(plan_use_aug) plan_use_aug,SUM(plan_use_sep) plan_use_sep,SUM(plan_use_oct) plan_use_oct,SUM(plan_use_nov) plan_use_nov,SUM(plan_use_dec) plan_use_dec,SUM(plan_use_total) plan_use_total, null, SUM(forecast_inv_cur_jan) forecast_inv_cur_jan, SUM(forecast_inv_cur_feb) forecast_inv_cur_feb, SUM(forecast_inv_cur_mar) forecast_inv_cur_mar, SUM(forecast_inv_cur_apr) forecast_inv_cur_apr, SUM(forecast_inv_cur_may) forecast_inv_cur_may, SUM(forecast_inv_cur_jun) forecast_inv_cur_jun, SUM(forecast_inv_cur_jul) forecast_inv_cur_jul, SUM(forecast_inv_cur_aug) forecast_inv_cur_aug, SUM(forecast_inv_cur_sep) forecast_inv_cur_sep, SUM(forecast_inv_cur_oct) forecast_inv_cur_oct, SUM(forecast_inv_cur_nov) forecast_inv_cur_nov, SUM(forecast_inv_cur_dec) forecast_inv_cur_dec, SUM(forecast_inv_cur_total) forecast_inv_cur_total, null, SUM(forecast_inv_next_jan) forecast_inv_next_jan, SUM(forecast_inv_next_feb) forecast_inv_next_feb, SUM(forecast_inv_next_mar) forecast_inv_next_mar, SUM(forecast_inv_next_apr) forecast_inv_next_apr, SUM(forecast_inv_next_may) forecast_inv_next_may, SUM(forecast_inv_next_jun) forecast_inv_next_jun, SUM(forecast_inv_next_jul) forecast_inv_next_jul, SUM(forecast_inv_next_aug) forecast_inv_next_aug, SUM(forecast_inv_next_sep) forecast_inv_next_sep, SUM(forecast_inv_next_oct) forecast_inv_next_oct, SUM(forecast_inv_next_nov) forecast_inv_next_nov, SUM(forecast_inv_next_dec) forecast_inv_next_dec, SUM(forecast_inv_next_total) forecast_inv_next_total, 'Grand' lvl_No
        FROM main
    `)

    if (!rows.length) {
        console.log(chalk.white.bgRed.bold(' No DATA '));
        return {status: false, message: 'No DATA', team: contactPointDepartment}
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
    let lastSubRow = false
    let nullRow = {
        null1: null, null2: null, null3: null, null4: null, null5: null, null6: null, null7: null, null8: null, null9: null, null10: null, null11: '', null12: '', null13: '', null14: '', null15: '', null16: '', null17: '', null18: '', null19: '', null20: '', null21: '', null22: '', null23: '', null24: '', null25: '', null26: '', null27: '', null28: '', null29: '', null30: '', null31: '', null32: '', null33: '', null34: '', null35: '', null36: '', null37: '', null38: '', null39: '', null40: '', null41: '', null42: '', null43: '', null44: '', null45: '', null46: '', null47: '', null48: '', null49: '', null50: '', null51: '', null52: '', null53: '', null54: '', null55: '', null56: '', null57: '', null58: '', null59: '', null60: '', null61: '', null62: '', null63: '', null64: '', null65: '', null66: '', null67: '', null68: '', null69: '', null70: '', null71: '', null72: '', null73: '', null74: '', null75: ''
    }
    for (let i = 0; i < rows.length; i++) {
        const item = rows[i];
        // Header
        // Set ครั้งแรก และ ทุกครั้งที่เปลี่ยน levelx ใหม่ โดยที่ data query จะต้องมี lvlx นั้นๆ ด้วย ไม้งั้นมันจะเช็คกับ null
        // == เมื่อตรงเงื่อนไข ==
        // 1. จะนำ ค่าจาก data query ที่ lvlx มาเก็บไว้ เพื่อไว้ check ในครั้งหน้าว่ายังเป็น level เดิมอยู่หรือไหม
        // 2. นำค่า Header ใส่ Excel
        // 3. เก็บตำแหน่ง ว่าตำแหน่งนี้ใช้ style ไหน
        if ((!level[0] || (level[0] && level[0] !== item.lvl1)) && item.lvl1) { 
            level[0] = item.lvl1
            arranged.push({lvl1: item.lvl1})
            levelStyle.lvl1.push(pos)
            pos++
        }
        if ((!level[1] || (level[1] && level[1] !== item.lvl2)) && item.lvl2) {
            level[1] = item.lvl2
            arranged.push({lvl2No: item.lvl2No, lvl2Name: item.lvl2Name, ...nullRow})
            levelStyle.lvl2.push(pos)
            pos++
        }
        if ((!level[2] || (level[2] && level[2] !== item.lvl3)) && item.lvl3) {
            level[2] = item.lvl3
            arranged.push({lvl3No: item.lvl3No, lvl3Name: item.lvl3Name, ...nullRow})
            levelStyle.lvl3.push(pos)
            pos++
        }
        if ((!level[3] || (level[3] && level[3] !== item.lvl4)) && item.lvl4) {
            level[3] = item.lvl4
            arranged.push({lvl4No: item.lvl4No, lvl4Name: item.lvl4Name, ...nullRow})
            levelStyle.lvl4.push(pos)
            pos++
        }

        // Data
        // 1. นำค่า data ใส่ Excel
        // 2. เก็บตำแหน่ง ว่าตำแหน่งนี้ใช้ style ไหน
        // 3. Set lastSubRow = true เพื่อให้รู้ว่าเป็นตัวสุดท้าย ถ้ามีการ สรุป Total ถ้ายังไม่สรุป ก็จะเป็น true เรื่อยๆไม่มีปัญหา
        arranged.push(item)
        levelStyle.lvl5.push(pos)
        lastSubRow = true
        pos++

        // Total
        // Check มีตัวต่อไปหรือไม่
        // ถ้าไม่มี แสดงว่าเป็นตัวสุดท้ายแล้ว ให้สรุป total จนครบ เช่น ตอนนี้อยู่ lvl3 ก็ต้องสรุป 3, 2, 1
        // ถ้ามี ก็ต้องเช็คว่าตัวต่อไปนั้นยังเป็น lvl เดิมหรือไม่
        // // ถ้ายังเป็น lvl เดิมอยู่ให้ข้ามไป
        // // ถ้าไม่ ให้ ลง total ของ lvl นั้นๆ เรียงจาก 4->3->2->1
        // // set lastSubRow = false เพราะมันเป็นตัวสุดท้ายของชุดแล้ว
        let tmpItem = rows[i + 1]
        if (tmpItem) {
            if ((!level[3] || (level[3] && level[3] !== tmpItem.lvl4)) && item.lvl4) {
                if (lastSubRow) {
                    arranged.push(nullRow)
                    pos++
                }
                let totalRow = totalRows.find(e => e.lvlNo.trim() === item.lvl4No)
                if (!totalRow) process.exit(0)
                arranged.push(totalRow)
                levelStyle.lvl4Total.push(pos)
                pos++
                arranged.push(nullRow)
                pos++
                lastSubRow = false
            }
            if ((!level[2] || (level[2] && level[2] !== tmpItem.lvl3)) && item.lvl3) {
                if (lastSubRow) {
                    arranged.push(nullRow)
                    pos++
                }
                let totalRow = totalRows.find(e => e.lvlNo.trim() === item.lvl3No)
                if (!totalRow) process.exit(0)
                arranged.push(totalRow)
                levelStyle.lvl3Total.push(pos)
                pos++
                arranged.push(nullRow)
                pos++
                lastSubRow = false
            }
            if ((!level[1] || (level[1] && level[1] !== tmpItem.lvl2)) && item.lvl2) {
                if (lastSubRow) {
                    arranged.push(nullRow)
                    pos++
                }
                let totalRow = totalRows.find(e => e.lvlNo.trim() === item.lvl2No)
                if (!totalRow) process.exit(0)
                arranged.push(totalRow)
                levelStyle.lvl2Total.push(pos)
                pos++
                arranged.push(nullRow)
                pos++
                lastSubRow = false
            }
            if ((!level[0] || (level[0] && level[0] !== tmpItem.lvl1)) && item.lvl1) {
                if (lastSubRow) {
                    arranged.push(nullRow)
                    pos++
                }
                let totalRow = totalRows.find(e => e.lvlNo.trim() === item.lvl1No)
                if (!totalRow) process.exit(0)
                arranged.push(totalRow)
                levelStyle.lvl1Total.push(pos)
                pos++
                arranged.push(nullRow)
                pos++
                lastSubRow = false
            }
        } else {
            // last item
            nowLvl = item.lvl4 ? 4 : item.lvl3 ? 3 : item.lvl2 ? 2 : 1
            arranged.push(nullRow)
            pos++
            let totalRow = totalRows.find(e => e.lvlNo.trim() === item[`lvl${nowLvl}No`])
            if (!totalRow) process.exit(0)
            arranged.push(totalRow)
            levelStyle[`lvl${nowLvl}Total`].push(pos)
            pos++
            arranged.push(nullRow)
            pos++
            
            if (nowLvl === 4) {
                // ถ้า 4 สรุป 3,2,1
                let totalRow3 = totalRows.find(e => e.lvlNo.trim() === item[`lvl${nowLvl-1}No`])
                if (!totalRow3) process.exit(0)
                arranged.push(totalRow3)
                levelStyle[`lvl${nowLvl - 1}Total`].push(pos)
                pos++
                arranged.push(nullRow)
                pos++

                let totalRow2 = totalRows.find(e => e.lvlNo.trim() === item[`lvl${nowLvl-2}No`])
                if (!totalRow2) process.exit(0)
                arranged.push(totalRow2)
                levelStyle[`lvl${nowLvl - 2}Total`].push(pos)
                pos++
                arranged.push(nullRow)
                pos++

                let totalRow1 = totalRows.find(e => e.lvlNo.trim() === item[`lvl${nowLvl-3}No`])
                if (!totalRow1) process.exit(0)
                arranged.push(totalRow1)
                levelStyle[`lvl${nowLvl - 3}Total`].push(pos)
                pos++
                arranged.push(nullRow)
                pos++
            }
            if (nowLvl === 3) {
                // ถ้า 3 สรุป 2, 1
                let totalRow2 = totalRows.find(e => e.lvlNo.trim() === item[`lvl${nowLvl-1}No`])
                if (!totalRow2) process.exit(0)
                arranged.push(totalRow2)
                levelStyle[`lvl${nowLvl - 1}Total`].push(pos)
                pos++
                arranged.push(nullRow)
                pos++

                let totalRow1 = totalRows.find(e => e.lvlNo.trim() === item[`lvl${nowLvl-2}No`])
                if (!totalRow1) process.exit(0)
                arranged.push(totalRow1)
                levelStyle[`lvl${nowLvl - 2}Total`].push(pos)
                pos++
                arranged.push(nullRow)
                pos++
            }
            if (nowLvl === 2) {
                // ถ้า 2 สรุป 1
                let totalRow1 = totalRows.find(e => e.lvlNo.trim() === item[`lvl${nowLvl-1}No`])
                if (!totalRow1) process.exit(0)
                arranged.push(totalRow1)
                levelStyle[`lvl${nowLvl - 1}Total`].push(pos)
                pos++
                arranged.push(nullRow)
                pos++
            }
            
        }
    }

    // Set Grand total
    let totalRow = totalRows.find(e => e.lvlNo.trim() === 'Grand')
    if (!totalRow) process.exit(0)
    arranged.push(totalRow)
    levelStyle[`lvlGTotal`].push(pos)

    // Set row data
    r = 4
    for (let i = 0; i < arranged.length; i++) {
        const rawData = arranged[i];
        let rowArray = []
        let columnCount = 0
        for (const key in rawData) {
            if (columnCount < 74) {
                if (['contactPointMobile'].includes(key)) {
                    if (rawData[key]) {
                        let contactNumber = rawData[key].replaceAll(/(\-+|\s+|\_+)/g, '');
                        if (Number(contactNumber) > 0) {
                            let strContactNumber = contactNumber.substring(0, 3)+ '-' + contactNumber.substring(3, 6)+ '-' + contactNumber.substring(6, 10);
                            rowArray.push(strContactNumber);
                        } else {
                            rowArray.push(rawData[key]);
                        }
                    } else {
                        rowArray.push(rawData[key]);
                    }
                } else {
                    rowArray.push(rawData[key])
                }
            }
            columnCount++
        }
        const row = ws.getRow(r);
        row.values = rowArray;
        r++
    }

    // Set style
    ws.eachRow({ includeEmpty: true },function (row, rowNumber) {
        if (rowNumber > 3) {
            row.eachCell(function(cell, colNumber) {
                if ([16, 17].includes(colNumber)) {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'D9D9D9' },
                    }
                }
            })
        }
    })

    let moneyColumn = [13,14,15,16,19,20,21,22,23,24,25,26,27,28,29,30,33,34,35,36,37,38,39,40,41,42,43,44,45,47,48,49,50,51,52,53,54,55,56,57,58,59,61,62,63,64,65,66,67,68,69,70,71,72,73]
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
        if (rowNumber > 3) {
            row.eachCell(function(cell, colNumber) {
                if (moneyColumn.includes(colNumber - 1)) {
                    ws.getCell(cell.address).alignment = { horizontal: 'right' };
                    cell.numFmt = '#,##0.00;[Red]-#,##0.00'
                }
            })
        }
    });
    
    // write file
    let deapartmentOwner = contactPointDepartment.replaceAll(/[<>:"\/\\|?*]+/g, '_')
    let xlsxName = `Template Budget CAPEX ${budgetYear} - ${deapartmentOwner.toUpperCase()}.xlsx`;
    await mkdirp(path.resolve(process.env.CAPEX_FOLDER_PATH));
    await new Promise((resolve, reject) => setTimeout(resolve, 555));
    let writeFileName = path.resolve(process.env.CAPEX_FOLDER_PATH + '/' + xlsxName);
    ws.name = 'CAPEX_' + deapartmentOwner
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
        console.log(chalk.white.bgGreen.bold(' SUCCESS '));
        return {...res, team: contactPointDepartment}
    }

    if (!res.status) {
        console.log(chalk.white.bgRed.bold(' ' + res.message + ' '));
        return {...res, team: contactPointDepartment}
        // knex.destroy();
        // process.exit();
    }
}