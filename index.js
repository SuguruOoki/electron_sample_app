
function LoadWrite(sheet,filename,sheetname){
    var xlsx = require('xlsx');
    var utils = xlsx.utils;
    var workbook = xlsx.readFile(filename);
    var sheetNames = workbook.SheetNames;
    var worksheet = workbook.Sheets[sheetNames[1]];
    //document.write(sheetNames[1]);
    //document.write(worksheet);
    var range = worksheet['!ref'];
    //document.write(range);
    //var rangeVal = utils.decode_range(range);
    var len = utils.sheet_to_json(worksheet).length;
    var content = JSON.stringify(utils.sheet_to_json(worksheet));
    var first = "①Pepperの発話内容";
    var second = "②反応する言葉";
    var third = "③お客さんの言葉に対するPepperの反応(Pepper言語で)";
    var forth = "④反応する言葉";
    var fifth = "⑤お客さんの言葉に対するPepperの反応";
    var start_date = "開始日";
    var end_date = "終了日";

    // 重複を削除したリスト

    date_dic = {};
    var number = 0;//conversationのナンバーと一致させるための変数
    var text = "";

    for (var i = 0; i < len; i++) {
        if(typeof(utils.sheet_to_json(worksheet)[i][first]) !== "undefined"){
            number++;
            date_dic[number] = [utils.sheet_to_json(worksheet)[i][start_date],utils.sheet_to_json(worksheet)[i][end_date]];
        }
    }

    var arr = [];
    //連想配列のvalueを取得する
    for(key in date_dic){arr.push(date_dic[key]);}

    //配列同士の比較を行う関数
    //配列同士が等しいならtrueを返す
    var isArrayEqual = function(array1, array2) {
        if (JSON.stringify(array1) === JSON.stringify(array2)){return true;}
            return false;
        }

    var unique = [];
    var flag;
    unique.push(arr[0]);

    for(var i = 0; i < arr.length; i++){
        flag = true;
        for(var j = 0; j < unique.length; j++){
            if(isArrayEqual(arr[i],unique[j])){
            flag = false;
            }
        }
        if(flag){
            unique.push(arr[i]);
        }
    }

    document.write("<br>");
    for(key in unique){document.write("unique:"+unique[key]+"<br>");}

    //uniqueとdate_dicを元にconversationのナンバーで分岐文を記述する。
    text = text + "proposal:%conversation_season_topic\n\n";
    //console.log("unique:"+unique);
    for (var i = 0; i < unique.length; i++) {
        //日付の条件文を記述
        text = text + "\nu1:(\"e:Launchpad/LifeTime $Dialog/Month=="+sheet.toString()+" $Dialog/Day>"+(Number(unique[i][0])-1).toString()+" $Dialog/Day<"+(Number(unique[i][1])+1).toString()+"\")\n    ^rand[\n";
        for (var j = 1; j <= arr.length; j++) {
            //持っている日付が同じなら条件文の中にconversationの種類を追加
            //console.log(unique[i]+":"+date_dic[j]);
            if (isArrayEqual(unique[i],date_dic[j])){
                text = text +"    \"^gotoReactivate(conversation_"+('00'+sheet).slice(-2)+"_"+('000'+j).slice(-3)+")\"\n";
            }
        }
        text = text + "    ]\n"
    }
    // document.write("<br>");
    // for (var key in unique) {document.write(unique[key]+"<br>");}
    var count = 0;
    //text = text + "length:"+len+"\n\n");
    for (var i = 0; i < len; i++) {
        //Pepperの発話内容をJSONから取り出す
        //最初に判断するのは一つ目の会話があるかどういか
        //あるなら次の処理をする
        if(typeof(utils.sheet_to_json(worksheet)[i][first]) !== "undefined"){
            text = text + "\n\n";
            count += 1;
            text = text + "proposal:%conversation_"+('00'+sheet).slice(-2)+"_"+('000'+count).slice(-3)+"\n";   //数値は指定の桁数で固定する
            //text = text + "    "+"\\vct=140\\ \\rspd=100\\ $SekkyakuPepper/Scene=conversation\n    "+JSON.stringify(utils.sheet_to_json(worksheet)[i][first]).replace(/["']+/g, '')+"\n    \\vct=140\\ \\rspd=100\\ $SekkyakuPepper/Scene=conversation_end\n");
            text = text + "    "+"\\vct=140\\ \\rspd=100\\ $SekkyakuPepper/Scene=conversation\n       "+JSON.stringify(utils.sheet_to_json(worksheet)[i][first]).replace(/["']+/g, '')+"\n    \\vct=140\\ \\rspd=100\\ $SekkyakuPepper/Scene=conversation_end\n";
            //最初が空欄ではなく２、３番目は空欄の場合
            if(typeof(utils.sheet_to_json(worksheet)[i][second]) === "undefined" && typeof(utils.sheet_to_json(worksheet)[i][third]) === "undefined"){
                text = text + "\n\n";
            }
            //最初が空欄ではなく２、３番目がきっちりあった場合
            else if(typeof(utils.sheet_to_json(worksheet)[i][second]) !== "undefined" && typeof(utils.sheet_to_json(worksheet)[i][third]) !== "undefined"){
                text = text + "\n        u1:(\"{*}"+JSON.stringify(utils.sheet_to_json(worksheet)[i][second]).replace(/["']+/g, '')+" {*} $SekkyakuPepper/Scene<>conversation $SekkyakuPepper/Scene<>speech\""+")\n";
                text = text + "\n            \\vct=140\\ \\rspd=100\\ $SekkyakuPepper/Scene=conversation\n                "+JSON.stringify(utils.sheet_to_json(worksheet)[i][third]).replace(/["']+/g, '')+"\n            \\vct=140\\ \\rspd=100\\ $SekkyakuPepper/Scene=conversation_end\n";
            }
        }
        //最初が空欄かつ二番目が空欄でない場合
        else if(typeof(utils.sheet_to_json(worksheet)[i][first]) === "undefined" && typeof(utils.sheet_to_json(worksheet)[i][second]) !== "undefined"){
            text = text + "\n        u1:(\"{*}"+JSON.stringify(utils.sheet_to_json(worksheet)[i][second]).replace(/["']+/g, '')+" {*} $SekkyakuPepper/Scene<>conversation $SekkyakuPepper/Scene<>speech\""+")\n";
            text = text + "\n            \\vct=140\\ \\rspd=100\\ $SekkyakuPepper/Scene=conversation\n                "+JSON.stringify(utils.sheet_to_json(worksheet)[i][third]).replace(/["']+/g, '')+"\n            \\vct=140\\ \\rspd=100\\ $SekkyakuPepper/Scene=conversation_end\n";
        }
        //最初と２番目が空欄の場合
        else if(typeof(utils.sheet_to_json(worksheet)[i][first]) === "undefined" && typeof(utils.sheet_to_json(worksheet)[i][second]) === "undefined"　&& typeof(utils.sheet_to_json(worksheet)[i][third]) !== "undefined"){
            text = text + "\n                \\vct=140\\ \\rspd=100\\ $SekkyakuPepper/Scene=conversation\n                "+JSON.stringify(utils.sheet_to_json(worksheet)[i][third]).replace(/["']+/g, '')+"\n            \\vct=140\\ \\rspd=100\\ $SekkyakuPepper/Scene=conversation_end\n";
        }

    }
    return text;
};

function textSave(text,month) {
    var blob = new Blob( [text], {type: 'text/plain'} );
    var link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = month.toString()+'月のトピック.txt';
    link.click();
};

//textSave('myfile',text);
