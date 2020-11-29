/**
 * Return a list of sheet names in the Spreadsheet with the given ID.
 * @param {String} a Spreadsheet ID.
 * @return {Array} A list of sheet names.
 */


var sid="1CpwNLrurUVVLX2dmMgZHU-uQC7WQfyfWqLlaiooRaN8";
var sname="ゲーム誕生日";


function doGet() {
  var ss = SpreadsheetApp.openById(sid);
  var sheets = ss.getSheetByName(sname);
  
　var last_row =sheets.getRange(1,6).getValue();
　var last_col = 6;
  
  var values= sheets.getRange(2,1,last_row-1 ,last_col).getValues();
 var value = JSON.parse(JSON.stringify(values));
  
  
  var title="";
  var title2="";
  var title3="";
  var title4="";
  var stat= "";
  var end= "";
  var moment = Moment.load();
  for (var i=1;i<value.length;i++){
  value[i][5]=i;
  }
  
  value.sort(function (a, b) {
    if(a[1]=="稼働"){return 1;}
    if(b[1]=="稼働"){return -1;}
    var nowyear=moment().format("YYYY-");
    var at=a[1];
    var bt=b[1];
    if(a[1]=="--"){at=nowyear + moment().add("h",-25).format("MM-DDT00:00:00Z")}
    if(b[1]=="--"){bt=nowyear + moment().add("h",-25).format("MM-DDT00:00:00Z")}
    var aa=moment(nowyear + moment(at).add("h",24).format("MM-DDT00:00:00Z"))-moment();
    var bb=moment(nowyear + moment(bt).add("h",24).format("MM-DDT00:00:00Z"))-moment();
    if(aa<0 && bb<0){return aa-bb;}
    if(aa<0){return 1;}
    if(bb<0){return -1;}
    
  return aa-bb;
  });
  
  var magic_imas="http://sokudon.s17.xrea.com/gametimer.html#";
  var html="";
  var url="http://www.shurey.com/js/timer/countdown.html?C,"; //154626840,
  
  for(var k=0;k<value.length;k++){
  var titleraw=value[k][0] +"("+value[k][2]+")";
  title=value[k][0] +"("+value[k][2]+")開始";
  title2=value[k][0] +"("+value[k][2]+")サービス終了";
  title3=value[k][0] +"("+value[k][2]+")周年";
  title4=value[k][0] +"("+value[k][2]+")命日";
    
  var reg = new RegExp('[<>#"%]',"gm");
  title = title.replace(reg,"");
  title2 = title2.replace(reg,"");
  title3 = title3.replace(reg,"");
  title4 = title4.replace(reg,"");
  
  stat= value[k][1];
  end= value[k][3];
  
  
  if(stat!=""){
  stat = (moment(stat).valueOf()/10000).toFixed(0);
  var kotosi = (moment(moment().format("YYYY-")+moment(stat*10000).format("MM-DDT00:00:00Z"))/10000).toFixed(0);
    
    if(value[k][0].match(/ミリシタ海外版/)){//value[k][4]!="";
      kotosi = (moment(moment().format("YYYY-")+moment(value[k][4]).format("MM-DDTHH:mm:00Z"))/10000).toFixed(0);
    }
      
    if(value[k][1]=="--"){
    html += "<tr><td>"+hyperlink(magic_imas + value[k][5],titleraw)+"<td>--</td><td>--</td>";
    }
    else{
  html += "<tr><td>"+hyperlink(magic_imas + value[k][5],titleraw)+"</td><td>"+ hyperlink( url +stat +"," +geturl(title),moment(stat*10000).format()) ;
  html += "</td><td>"+ hyperlink( url + kotosi +"," +geturl(title3),moment(kotosi*10000).format()) +"</td>";
  html += "</td>";
    }
  }
  if(end!=""){
  end = (moment(end).valueOf()/10000).toFixed(0);
  var kotosi = (moment(moment().format("YYYY-")+moment(end*10000).format("MM-DDT00:00:00Z"))/10000).toFixed(0);
      
    
  html += "<td>"+ hyperlink( url +end +"," +geturl(title2),moment(end*10000).format())  ;
  html += "</td><td>"+ hyperlink( url + kotosi +"," +geturl(title4),moment(kotosi*10000).format()) +"</td>";
  html += "</td>";
  }
    html += "</tr>";
  }
    
  html+= "<tr><td><b>あいます関連</b></td></tr><tr><td>あいますげーむたいまー</td>"+"<td>"+ hyperlink("http://sokudon.s17.xrea.com/gametimer.html","http://sokudon.s17.xrea.com/gametimer.html")+"</td></tr>";
 
  
  var header= "<style>th,td{  border:solid 1px #aaaaaa;},.table-scroll{  overflow-x : auto}</style>";
  var h ="<table><thead><tr><th>ゲーム名</th><th>開始日から</th><th>今年の周年</th><th>サービス終了から</th><th>今年の命日</th></tr></thead>";
  
  
  
 html= h+"<tbody>"+ header  +html + "<tbody></table>";
  
  return HtmlService.createHtmlOutput(html);
  //return ContentService.createTextOutput(url).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

//<>#"%　"#"はURI参照として、"%"はエスケープ用文字として使われます。
//除外されている記号 (RFC2396 に定義がないもの)
//以下の文字は使用できません。
// {}|\^[]`<>#"%

function geturl(title){

   var i, len, arr;
        for(i=0,len=title.length,arr=[]; i<len; i++) {
          if(title.charCodeAt(i) < 0x80){
            arr += title.substr(i, 1);
          }
          else{
            arr +="%25u"+  ("00"+title.charCodeAt(i).toString(16)).slice(-4);
          }
        }

  return arr;
}

function hyperlink(link,name){
  link= "<a href='" + link +"' target=\"_blank\" rel=\"noopener\">" +name +"</a>";
  
  return link;
}

function wmap_getSheetsName(sheets){
  //var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var sheet_names = new Array();
  
  if (sheets.length >= 1) {  
    for(var i = 0;i < sheets.length; i++)
    {
      sheet_names.push(sheets[i].getName());
    }
  }
  return sheet_names;
}