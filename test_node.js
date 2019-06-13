const hello = require('./build/Release/test');
const fs = require('fs')

const result = hello('https://b2b.ccb.com')
for (const i in result) {
  fs.writeFileSync('C:\\work\\IE\\frame\\' + i, result[i], 'utf8')
  // console.log(result[i])
}
console.log('get: ', result.length)
console.log('中文')