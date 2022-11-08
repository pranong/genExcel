const fs = require('fs')
const path = require('path')
const mkdirp = require('mkdirp')
const createReport = require('docx-templates').default

process.on('message', async param => {
  try {
    param.noSandbox = true
    console.log('param', param)
    // param.vm2Sandbox = true
    // require('fs').writeFileSync('./param.json', JSON.stringify(param))
    if (typeof param.template === 'string') {
      param.template = fs.readFileSync(param.template)
    }
    let buff = await createReport(param)
    if (param.output) {
      await mkdirp(path.dirname(param.output))
      console.log('mkdirp', path.dirname(param.output))
      fs.writeFileSync(param.output, buff)
    }
    console.log('done')
    process.send(buff)
    process.exit(0)
  } catch (e) {
    console.log('createReport error', e)
    process.send(false)
    process.exit(1)
  }
})
