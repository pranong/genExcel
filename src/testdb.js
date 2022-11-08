const config = require('../src/config')
const knex = require('./lib/knex')('precaldev02', config['precaldev02']);

// console.log(knex);

run();
async function run() {
  console.log('Oracle Test Code');
  let rows = await knex('dual');
  console.log('rows=', rows);
}


