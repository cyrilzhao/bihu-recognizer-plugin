const getAllFrame = require('./build/Release/frame_addon');
const fs = require('fs')

const result = getAllFrame('https://b2b.ccb.com')
for (const i in result) {
  console.log(result[i].url)
  for (const j in result[i].frames) {
    fs.writeFileSync('C:\\work\\IE\\frame\\' + i + '_' + j, result[i].frames[j], 'utf8')
  }
  // fs.writeFileSync('C:\\work\\IE\\frame\\' + i, result[i], 'utf8')
  // console.log(result[i])
}
console.log('get: ', result.length)
