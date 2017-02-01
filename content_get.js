const xlsx = require('xlsx');
const utils = xlsx.utils;
function sheet_content_get(){
  let workbook = xlsx.readFile('test.xlsx');
  worksheet = workbook.Sheets['1月'];
  var range = worksheet['!ref'];
  var rangeVal = utils.decode_range(range);
  content = utils.sheet_to_json(worksheet);
  console.log(JSON.stringify(content));
  document.write(typeof(JSON.stringify(content)))
  a = JSON.stringify(content)
  for(i=0;i < a.length;i++){
    document.write(a[i])
    document.write('<br>')
  }
  return content
}