const getAllFrame = require('./build/Release/frame_addon');
const fs = require('fs')

const result = getAllFrame('https://b2b.ccb.com')
for (const i in result) {
  fs.writeFileSync('C:\\work\\IE\\frame\\' + i, result[i], 'utf8')
  // console.log(result[i])
}
console.log('get: ', result.length)
