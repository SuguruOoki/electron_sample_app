function loadTextFile(){
httpObj = createXMLHttpRequest(displayData);
if (httpObj){
//ファイル名の指定
httpObj.open("GET","whatsnew.csv",true);
httpObj.send(null);
}
}
 
//WebKit（KHTML）におけるXMLHttpRequestのresponceText文字化け防止処理
var ajax_filter = function(t){return t};
if(navigator.appVersion.indexOf( "KHTML" ) > -1){
    ajax_filter = function(t){
        var esc = escape(t);
        return(esc.indexOf("%u") < 0 && esc.indexOf("%") > -1) ? decodeURIComponent(esc) : t
    }
}
 
function displayData(){
    if ((httpObj.readyState == 4) && (httpObj.status == 200)){
        var text = ajax_filter(httpObj.responseText);
        document.getElementById("result").innerHTML = parseText(text);
    }
}
 
//CSVデータを配列にし、HTMLへ入れこむ処理
function parseText(str){
    var resultText="<table border=1>";
    //改行コードを変数にし分割処理をする
    var CR = String.fromCharCode(13);
    var LF = String.fromCharCode(10);
    lineData = str.split(CR);
 
    for (var i=0; i<lineData.length; i++){
        strText = lineData[i].split(",");
        resultText += "<tr>";
 
        //HTMLのtableを繰り返し処理で作成し、セルにデータを入れていく
        //i < 4は列数に応じて変更する
        for (var j=0; j<strText.length; j++){
            if (i < 4){
                if(j == 0){
                    resultText += "<td>"+strText[j]+"</td>";
                }
                if(j == 1){
                    if (i == 0){
                        resultText += "<td>";
                    }else{
                        resultText += "<td><a href="+strText[j]+">";
                    }
                }
                if(j == 2){
                    if (i == 0){
                        resultText += strText[j]+"</td>";
                    }else{
                        resultText += strText[j] +"</a></td>";
                    }
                }
            }
        }
        resultText += "</tr>";
    }
    resultText += "</table>";
    return resultText;
}