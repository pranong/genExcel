module.exports = {
  snakeToCamel,
  camelToSnake,
  camelToSnakeUpper,
  arrayToPack,
  packToArray,
  promiseWhile,
  promiseForEach,
  promiseSome,
  promiseEvery,
  delay,
}

function snakeToCamel(s) {
  let newObj = {}
  Object.keys(s).forEach((k) => {
    newObj[
      k.toLowerCase().replace(/(\w)(_\w)/g, (m) => m[0] + m[2].toUpperCase())
    ] = s[k]
  })
  return newObj
}

function camelToSnake(s) {
  return s.replace(/[A-Z]/g, (m) => '_' + m[0].toLowerCase())
}

function camelToSnakeUpper(s) {
  return s.replace(/[A-Z]/g, (m) => '_' + m[0]).toUpperCase()
}

function arrayToPack(rows) {
  if (rows.length === 0) {
    return { c: [], d: [] }
  }
  let c = Object.keys(rows[0])
  let d = rows.map((row) => c.map((col) => row[col]))
  return { c, d }
}

function packToArray(obj) {
  if (!obj.c || !obj.d) {
    return []
  }
  return obj.d.map((row) =>
    obj.c.reduce((prev, c, i) => {
      prev[c] = row[i]
      return prev
    }, {})
  )
}

function promiseWhile(condition, action) {
  function loop() {
    if (!condition()) {
      return Promise.resolve()
    }
    return Promise.resolve(action()).then(loop)
  }
  return loop()
}

async function promiseForEach(arr, fn) {
  for (let i = 0; i < arr.length; i++) {
    await Promise.resolve(fn(arr[i], i, arr))
  }
}

async function promiseSome(arr, fn) {
  for (let i = 0; i < arr.length; i++) {
    let ok = await Promise.resolve(fn(arr[i], i, arr))
    if (ok) {
      return Promise.resolve(true)
    }
  }
  return Promise.resolve(false)
}

async function promiseEvery(arr, fn) {
  for (let i = 0; i < arr.length; i++) {
    let ok = await Promise.resolve(fn(arr[i], i, arr))
    if (!ok) {
      return Promise.resolve(false)
    }
  }
  return Promise.resolve(true)
}

function delay(ms) {
  return new Promise((resolve, reject) => setTimeout(resolve, ms))
}
