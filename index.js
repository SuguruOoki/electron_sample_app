const xlsx = require('xlsx');
const utils = xlsx.utils;
let workbook = xlsx.readFile('test.xlsx');
let sheetNames = workbook.SheetNames;
worksheet = workbook.Sheets['1月'];
var range = worksheet['!ref'];
var rangeVal = utils.decode_range(range);
var len = utils.sheet_to_json(worksheet).length;
var content = JSON.stringify(utils.sheet_to_json(worksheet));

for (var i = 0; i < utils.sheet_to_json(worksheet).length; i++) {

    document.write(JSON.stringify(utils.sheet_to_json(worksheet)[i]["①Pepperの発話内容"])+"<br><br>");

};


function textSave(name, text) {
    var blob = new Blob( [text], {type: 'text/plain'} );
    var link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = name + '.txt';
    link.click();
}

// textSave('myfile',content);
// document.write(content);

// var a = document.createElement('a');
// a.textContent = 'export';
// a.download = 'context.json';
// a.href = window.URL.createObjectURL(new Blob([content], { type: 'text/plain' }));
// a.dataset.downloadurl = ['text/plain', a.download, a.href].join(':');
 
// var exportLink = document.getElementById('export-link');
// exportLink.appendChild(a);
