const { fork } = require('child_process')
const path = require('path')
const fs = require('fs')
const fsp = fs.promises
const mkdirp = require('mkdirp')

module.exports = createReport

async function createReport(param) {
  let removeDuration = parseInt(process.env.REPORT_TEMP_REMOVE_DURATION || '60000') || 60000
  return new Promise((resolve, reject) => {
    if (process.pkg) {
      param.noSandbox = true
      if (typeof param.template === 'string') {
        param.template = fs.readFileSync(param.template)
      }
      return require('docx-templates').default(param).then(async buff => {
        if (param.output) {
          await mkdirp(path.dirname(param.output))
          fs.writeFileSync(param.output, buff)
          // unlinkTempFile(param.output, removeDuration)
        }
        resolve(buff)
      }).catch(reject)
    }
    const job = fork(path.resolve(__dirname, './create-report-job'))
    let t0 = new Date().getTime()
    job.on('message', buff => {
      console.log('job message', typeof buff)
      if (buff === false) {
        return reject(new Error('create report error'))
      }
      console.log('report done in', new Date().getTime() - t0)
      resolve(buff)
      if (param.output) {
        // unlinkTempFile(param.output, removeDuration)
      }
    })
    job.send(param)
    job.on('error', err => {
      console.log('job error', err)
      reject(new Error('create report error 2'))
    })
    job.on('exit', exitCode => {
      console.log('job end', exitCode)
      resolve(true)
    })
  })
}

async function unlinkTempFile(file, delay) {
  console.log(`delete '${file}' in ${delay}ms`)
  setTimeout(async () => {
    try {
      await fsp.unlink(file)
    } catch (e) {
      console.log(`delete '${file}' error`)
    } finally {
      console.log(`delete '${file}' done`)
    }
  }, delay)
}
