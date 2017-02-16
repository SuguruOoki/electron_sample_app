const xlsx = require('xlsx');
const utils = xlsx.utils;
let workbook = xlsx.readFile('test.xlsx');
let sheetNames = workbook.SheetNames;
worksheet = workbook.Sheets['1月'];
var range = worksheet['!ref'];
var rangeVal = utils.decode_range(range);
var len = utils.sheet_to_json(worksheet).length;
var content = JSON.stringify(utils.sheet_to_json(worksheet));

var first = "①Pepperの発話内容"
var second = "②反応する言葉"
var third = "③お客さんの言葉に対するPepperの反応(Pepper言語で)"
var forth = "④反応する言葉"
var fifth = "⑤お客さんの言葉に対するPepperの反応"

for (var i = 0; i < utils.sheet_to_json(worksheet).length; i++) {

    //Pepperの発話内容をJSONから取り出すS
    if(typeof(utils.sheet_to_json(worksheet)[i][first]) !== "undefined"){
        document.write(JSON.stringify(utils.sheet_to_json(worksheet)[i][first])+"<br>");
        document.write(JSON.stringify(utils.sheet_to_json(worksheet)[i][second])+"<br>");
        document.write(JSON.stringify(utils.sheet_to_json(worksheet)[i][third])+"<br>");
        if(typeof(utils.sheet_to_json(worksheet)[i][forth]) !== "undefined"){
            document.write(JSON.stringify(utils.sheet_to_json(worksheet)[i][forth])+"<br>");
        }
        else if(typeof(utils.sheet_to_json(worksheet)[i+1][first]) === "undefined"){
            document.write(JSON.stringify(utils.sheet_to_json(worksheet)[i+1][second])+"<br>");
            document.write(JSON.stringify(utils.sheet_to_json(worksheet)[i+1][third])+"<br>");
        }
        if(typeof(utils.sheet_to_json(worksheet)[i][fifth]) !== "undefined"){
            document.write(JSON.stringify(utils.sheet_to_json(worksheet)[i][fifth])+"<br><br><br>");
        }

        document.write("<br><br>")
    }
    // document.write(JSON.stringify(utils.sheet_to_json(worksheet)[i])+"<br><br><br>");

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
