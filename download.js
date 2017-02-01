function handleDownload() {
    const xlsx = require('xlsx');
    const utils = xlsx.utils;
    let workbook = xlsx.readFile('test.xlsx');
    let sheetNames = workbook.SheetNames;
    //worksheet = workbook.Sheets['Sheet1'];
    worksheet = workbook.Sheets['1月'];
    var range = worksheet['!ref'];
    var rangeVal = utils.decode_range(range);
    var content = utils.sheet_to_json(worksheet);
    console.log(content);
    //var content
    var value
    /*for (let r=rangeVal.s.r ; r <= rangeVal.e.r ; r++) {
        for (let c=rangeVal.s.c ; c <= rangeVal.e.c ; c++) {
            let adr = utils.encode_cell({c:c, r:r});
            let cell = worksheet[adr];
            if(!(value == null | value == "" | value == undefined)){
             value = cell.v;
             content = value
             console.log(r)
             console.log(cell+":"+value)
         }
     }*/
 var blob = new Blob([ content ], { "type" : "text/json" });

 if (window.navigator.msSaveBlob) { 
    window.navigator.msSaveBlob(blob, "test.json"); 

    window.navigator.msSaveOrOpenBlob(blob, "test.json"); 
 } else {
    document.getElementById("download").href = window.URL.createObjectURL(blob);
 }
}