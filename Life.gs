// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 変更(行挿入 / 削除)時の補正
function onEdit(e){

  // ---------- ---------- ----------  
  // 定数
  const DowCol = 1;
  const DateCol = 2;
  const PCol = 5;
  const EventCol = 6;
  
  const DefaultBGColor = "#ffffff";
  const ErrorFormula = "#REF!";
  // ---------- ---------- ----------  
  
  // 変更したシートを取得
  var sheet = e.source.getActiveSheet();
 
  // 変更行
  var rowStart = e.range.rowStart;
  var rowEnd = e.range.rowEnd;

  // 最終行  
  var lastRow = sheet.getLastRow();
  
  /* 変更列(必要に応じて)
  var columnStart = e.range.columnStart;
  var columnEnd = e.range.columnEnd;
  */
  
  //Browser.msgBox("rowStart: " + rowStart + "\\n rowEnd: "+ rowEnd);

  // シート名
  var sheetName = sheet.getSheetName();
  
  // シートが月次シートかをチェック(YYYY_MM)
  if (sheetName.match(/\d{4}_\d{1,2}/)) {

    // 変更範囲
    for (var row = rowStart; row <= rowEnd; row++) {

      // 1,2行目はスキップ
      if (row <= 2) {
        continue;
      }
                  
      // 日付列に数式が設定されてない場合
      if (!sheet.getRange(row, DateCol).getFormula()){
        sheet.getRange(row, DowCol).setFormula("=A" + String(row-1));
        sheet.getRange(row, DateCol).setFormula("=B" + String(row-1));        
      }
      
      // タスク作業時間列に数式が設定されていない場合
      if (sheet.getRange(row, PCol).isBlank()){ // 空白の場合のみ
        if (!sheet.getRange(row, PCol).getFormula()){
          // 前日の数式を取得
          var formarFormula = sheet.getRange(row-1, PCol).getFormulaR1C1();
          sheet.getRange(row, PCol).setFormulaR1C1(formarFormula);        
        }
      }

      // Eventが空欄にも関わらず背景色が設定されている場合
      if ((sheet.getRange(row, EventCol).getValue() === "")
          && (sheet.getRange(row, EventCol).getBackground() !== DefaultBGColor)) {
        
        sheet.getRange(row, EventCol, 1, 3).setBackground(null);
      }
          
      // 日付列の数式が不正(=#REF!)の場合
      if (sheet.getRange(row, DateCol).getValue() === ErrorFormula){
        sheet.getRange(row, DowCol).setFormula("=A" + String(row-1));
        sheet.getRange(row, DateCol).setFormula("=B" + String(row-1));        
      }
          
      // 最終変更行の次の行の数式補正
      if (row === rowEnd) {
        
        // 最終変更行が最終行でない場合
        if (row != lastRow) {
          var nextDateFormula = sheet.getRange(row+1, DateCol).getFormula();
          
          // 正しい数式
          var correctDateFormula = "=B" + String(row);
          
          // 日付が同日の場合のみ補正
          if ((!nextDateFormula.match(/\+/))
              && (nextDateFormula !== correctDateFormula)) {
            
            sheet.getRange(row+1, DowCol).setFormula("=A" + String(row));
            sheet.getRange(row+1, DateCol).setFormula("=B" + String(row));        
          }          
        }
      }
      
      // 罫線補正
      var dateFormula = sheet.getRange(row, DateCol).getFormula();
      
      // 日付が同月
      if (!dateFormula.match(/\+/)) {        
        // 上の罫線をクリア
        // (top, left, bottom, right, vertical, horizontal)
        sheet.getRange(row, DowCol, 1, 2).setBorder(false, null, null, null, null, null);
      } else {
        // 上の罫線をセット
        // (top, left, bottom, right, vertical, horizontal, color, style)
        sheet.getRange(row, DowCol, 1, 2).setBorder(true, null, null, null, null, null, "black", null);        
      }
      
    }
  }  
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 現在時刻をセット
function currentTime(){

  // 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 現在のシートを取得
  var sheet = spreadsheet.getActiveSheet();

  // 選択中のセルを取得
  var cell = sheet.getActiveCell();

  // 現在日時
  var now = new Date();
 
  /* 
   (未実装)
     ５分間隔でセット時刻を調節

  // 分
  var min = now.getMinutes();
  
  // 残余(5分間隔)
  var residue = min % 5;
  */
 
  // var time = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm');
  //var time = Utilities.formatDate(now, 'GMT', 'HH:mm');
  var time = Utilities.formatDate(now, 'GMT+1', 'HH:mm');

  // セルに値を設定
  cell.setValue(time);  
}

// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
/*
 [Memo]
 Accumシート変更時の修正箇所
 
 1.列を変更した場合
   (1.1) [btn_clearCurrentMonth] outCurrentMonthCol (今月の日付列)
   (1.2) [accumEachMonth] outPastMonthCol (過去月の日付列)
   (1.3) [accumEachMonth] outCurrentMonthCol (今月の日付列)
  
 2.行を変更した場合
   (2.1) [btn_clearCurrentMonth] outStartRow (今月の開始行)
   (2.2) [accumEachMonth] outStartRow (出力シートの開始行) 
   
*/

// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
// 今月のタスク積算時間をクリア
function btn_clearCurrentMonth(){
  
  // ---------- ---------- ----------  
  // 定数  
  const numCol = 2;

  const outStartRow = 3;
  const outCurrentMonthCol = 10;
  const outSheetName = "A";
  // ---------- ---------- ----------  

  // 出力シート
  var outSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(outSheetName);  
  var lastRow = outSheet.getLastRow();

  // 出力シート(今月分)をクリア
  outSheet.getRange(outStartRow, outCurrentMonthCol, lastRow, numCol).clearContent();
  outSheet.getRange(outStartRow, outCurrentMonthCol, lastRow, numCol).setBackground(null);  
}

// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
// 今月のタスク積算時間を計算
function btn_accumCurrentMonth(){
  accumEachMonth(null);
}

// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
// 過去月のタスク積算時間を計算
function btn_accumPastMonth(){
  accumEachMonth("ON");
}

// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
// タスクの積算時間を計算
function accumEachMonth(flag_Month){

  // ---------- ---------- ----------  
  // 定数  
  const numCol = 2;

  // Life Log file ID
  const llFileID = "1sOUFKC8k2INHb6ag71uxm4FBFmUSL5-hN1bM59UldyQ";

  const llStartRow = 2;
  const llDateCol = 2;
  const llDurationCol = 5;
  const llTaskCol = 6;
  
  const targetCol = 2;
  const targetMonthRow = 1;
  const targetTaskRow = 2;

  const outStartRow = 3;
  const outPastMonthCol = 5;
  const outCurrentMonthCol = 10;
  const outSheetName = "A";
  // ---------- ---------- ----------  

  // 出力シート
  var outSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(outSheetName);  
  var lastRow = outSheet.getLastRow();
  
  // タスク名
  var targetTask = outSheet.getRange(targetTaskRow, targetCol).getValue();
  
  // 過去月
  var pastMonth = outSheet.getRange(targetMonthRow, targetCol).getValue();
  
  // ---------- ---------- ----------  
  // 今月のシートを取得
  // 現在月
  var date = new Date();

  // 先頭の0を取り除く
  var monthStr = new String(date.getMonth()+1);
  month = monthStr.replace("/^0/","");
  
  // 今月のシート名
  var currentMonthSheetName = date.getFullYear() + '_' + month;

  var currentMonthSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(currentMonthSheetName);
  // ---------- ---------- ----------  
  
  if (flag_Month) {
    // 過去月
    
    var llSpreadSheet = SpreadsheetApp.openById(llFileID);
    var pastMonthSheet = llSpreadSheet.getSheetByName(pastMonth);
    
    if(pastMonthSheet){
    
      // 出力シート(過去月分)をクリア
      outSheet.getRange(outStartRow, outPastMonthCol, lastRow, numCol).clearContent();
    
      // 過去月のタスク積算時間を計算
      sub_accumTaskTime(pastMonthSheet, outSheet, llStartRow, outStartRow, llDateCol, outPastMonthCol, llTaskCol, llDurationCol, targetTask, null);
      
    } else {
      Browser.msgBox("過去月の指定が正しくありません。\\n(" + pastMonth + ")");
    }
      
  } else {
    // 今月

    // 計算済みの最終日を取得
    var lastCalcDate = sub_searchLastCalcDate(outSheet, outStartRow, outCurrentMonthCol);

    // 計算済み行が無い場合
    if (!lastCalcDate.row) {
      // 現在月のタスク積算時間を計算 (全て)
      sub_accumTaskTime(currentMonthSheet, outSheet, llStartRow, outStartRow, llDateCol, outCurrentMonthCol, llTaskCol, llDurationCol, targetTask, null);
    } else{
      // 計算済みの最終日に該当する先頭行を取得
      var iStartRow = sub_searchLastCalcRow(currentMonthSheet, llStartRow, llDateCol, lastCalcDate.date);
      
      // 現在月のタスク積算時間を計算 (計算済みの最終日以降)
      sub_accumTaskTime(currentMonthSheet, outSheet, iStartRow, lastCalcDate.row, llDateCol, outCurrentMonthCol, llTaskCol, llDurationCol, targetTask, lastCalcDate.time);
    }
  }
}

// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
// 一ヶ月分の積算時間を出力
function sub_accumTaskTime(iSheet, oSheet, iStartRow, oStartRow, iDateCol, oDateCol, iTaskCol, iDurationCol, targetTask, accumulatedTime){

  // ---------- ---------- ----------  
  // 定数  
  // 当日の背景色  
  const todayBGColor = "#d9ead3";
  // ---------- ---------- ----------  
  
  var lastRow = iSheet.getLastRow();
  var formerDate = iSheet.getRange(iStartRow, iDateCol).getValue();

  var outRow = oStartRow;

  // 積算時間をチェック
  if (!accumulatedTime) {
    accumulatedTime = new MyTime("");
  }
  
  // 現在日
  var now = new Date();
    
  // 最終行まで
  for (row=iStartRow; row<=lastRow; row++){

    // 現在行の日付
    var currentDate = iSheet.getRange(row, iDateCol).getValue();
        
    // 明日以降はスキップ
    if (currentDate.getTime() > now.getTime()){
      break;
    }
       
    // 日付が同じかをチェック
    if (currentDate.getTime() == formerDate.getTime()){

      // 現在行のタスク
      var currentTask = iSheet.getRange(row, iTaskCol).getValue();

      // 対象タスクかをチェック
      if (currentTask == targetTask){
        
        // 対象タスクの活動時間
        var currentTaskDate = iSheet.getRange(row, iDurationCol).getValue();

        if (currentTaskDate) {
          var currentTaskTimeString = currentTaskDate.toTimeString().split(" ");
          var currentTaskTime = new MyTime(currentTaskTimeString[0]);
          
          // 活動時間を加算
          accumulatedTime.add(currentTaskTime);
        }
      }
    } else {
      // 日付とタスクの合算時間を出力
      oSheet.getRange(outRow, oDateCol).setValue(formerDate);
      oSheet.getRange(outRow, oDateCol+1).setValue(accumulatedTime.disp());
     
      outRow++;
      formerDate = currentDate;
    }
    // 日付とタスクの合算時間を出力 (最終日)
    oSheet.getRange(outRow, oDateCol).setValue(formerDate);
    oSheet.getRange(outRow, oDateCol+1).setValue(accumulatedTime.disp());
  }

  // 当日チェック
  if ((formerDate.getFullYear() == now.getFullYear()) && 
    (formerDate.getMonth() == now.getMonth()) &&
      (formerDate.getDate() == now.getDate()) ) {
        // 背景色をセット
        oSheet.getRange(outRow, oDateCol, 1, 2).setBackground(todayBGColor);
      }
}

// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
// 計算済みの最終日を検索
function sub_searchLastCalcDate(oSheet, oStartRow, oDateCol){
  
  // ---------- ---------- ----------  
  // 定数  
  // 当日の背景色  
  const todayBGColor = "#d9ead3";
  // ---------- ---------- ----------  

  var lastRow = oSheet.getLastRow();

  var obj = new Object();
  var lastDate;  

  // 前日までの積算時間
  var accumulatedTime = null;

  for (row=lastRow; row>=oStartRow; row--){
    lastDate = oSheet.getRange(row, oDateCol).getValue();

    // 計算済みかをチェック
    if(lastDate){

      // 背景色をチェック
      var bgColor = oSheet.getRange(row, oDateCol).getBackground();
      if (bgColor = todayBGColor){
        // 背景色をクリア
        oSheet.getRange(row, oDateCol, 1, 2).setBackground(null);
      }
      
      // 2日目以降
      if (row > oStartRow){
        
        // 前日までの積算時間
        var taskTime = oSheet.getRange(row-1, oDateCol+1).getDisplayValue();

        // 秒は非表示のため、個別に取得
        var taskDate = oSheet.getRange(row-1, oDateCol+1).getValue();
        var taskDateSeconds = taskDate.toTimeString().split(" ")[0].split(":")[2];
        
        if (taskTime) {
          accumulatedTime = new MyTime(taskTime + ":" + taskDateSeconds);
          obj.time = accumulatedTime;
        }        
      }
      
      obj.row = row;
      obj.date = lastDate.getDate();
      break;
    }
  }

  return obj;
}

// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
// 指定日に該当する先頭行を検索
function sub_searchLastCalcRow(iSheet, iStartRow, iDateCol, targetDate){

  var lastRow = iSheet.getLastRow();

  var row;
  for (row=iStartRow; row<=lastRow; row++){
    currentDate = iSheet.getRange(row, iDateCol).getValue();

    // 日付が一致    
    if(currentDate.getDate() == targetDate){
      break;
    }
  }
  return row;
}

// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
// 時間計算用クラス
var MyTime = function(str) {
  
  var hms = str.split(":");

  this.h = parseInt(hms[0]|"0");
  this.m = parseInt(hms[1]|"0");
  this.s = parseInt(hms[2]|"0");
  
  this.time = this.h * 3600 + this.m*60 + this.s;
  
  this.add = function(n) {
    this.addHours(n.h);
    this.addMinutes(n.m);
    this.addSeconds(n.s);
  }

  this.addHours = function(n) {
    this.time += n*3600;
  };

  this.addMinutes = function(n) {
    this.time += n*60;
  };
  
  this.addSeconds = function(n) {
    this.time += n;
  };

  this.disp = function() {
    var minus = this.time<0?"-":"";
    
    var t = Math.abs(this.time);
    var hh = Math.floor(t/3600).toString();
    var mm = (100+Math.floor((t%3600)/60)).toString().substr(-2);
    var ss = (100+Math.round(t%60)).toString().substr(-2);

    return minus + hh + ":" + mm + ":" + ss;
    
  };
};

// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
// セルが結合されているかをチェック
function isMerge(targetRange) {
  var range = SpreadsheetApp.getActive().getActiveSheet().getRange(targetRange);

  var merges = [];
  for (var i = 0; i < range.getHeight(); i++)
  {
    var merge = range.offset(i, 0, 1, 1).isPartOfMerge();
    merges.push(merge);    
  }
  return merges;
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 指定文字列に該当する日付を取得
// args
//
// return:
function findDate(sheetName, range, val){  
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 指定行の直前の日付を取得
//
// args
//  row:日付を取得する対象行
//  col:日付列
//
// return:日付
function rowDate(row, col) {
  //Logger.log("row:%s, col:%s", row, col);

  // 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 現在のシートを取得
  var sheet = spreadsheet.getActiveSheet();

  // 対象のセル範囲を取得
  var range = sheet.getRange(row, col);

  var i = 0;
  var date = range.getValue();
  while(!date) {
    i--;
    date = range.offset(i, 0).getValue();
    // Logger.log("i:%s",i);   
  }

  // Logger.log("date:%s",date);
  return date;
}

// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
// 先月の最終行の値を取得
function getLastRowData(col) {

  // 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var currentSheetName = spreadsheet.getActiveSheet().getName();
  
  // '_'で分割
  var currentSheetNameList = currentSheetName.split('_');
  
  var year = currentSheetNameList[0];
  var month = currentSheetNameList[1];
  
  // 先月を計算
  var date = new Date(year, month);
  date.setMonth(date.getMonth()-2);

  // 先頭の0を取り除く
  lastMonth = new String(date.getMonth()+1);
  lastMonth = lastMonth.replace("/^0/","");

  // 先月のシート名
  var lastMonthSheetName = date.getFullYear() + '_' + lastMonth;

  var lastMonthSheet = spreadsheet.getSheetByName(lastMonthSheetName);

  // 最終行
  var lastRow = lastMonthSheet.getLastRow();
  
  var val = lastMonthSheet.getRange(lastRow, col).getValue();
  
  return val;

}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// テスト関数
// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 対象:rowDate()
function test_rowDate() {
  var date = rowDate("175", "2");
  Logger.log("<row:175>%s", date)
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// サンプル
// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// クラス
function sample_Class(name){
  this.name = name;
  this.msg = function(text){
    Logger.log("myClass(%s):%s", this.name, text);
  }
}

function sample_Function(){
  var test = new sample_Class("foo");
  test.msg("hello class");
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// snippet
// ---------- ---------- ---------- ---------- ---------- ---------- ----------

/*
 ログ出力
 
  // LogsとStackdriver logs(Execusions)の両方に出力
  Logger.log("%s", );

  console.log("%s", );
  
  // msgBoxで改行する場合: \\n
  Browser.msgBox("msg" + arg);

  // キャンセルボタンを配置
  var result = Browser.msgBox("msg", Browser.Buttons.OK_CANCEL);
  if (result === "cancel"){
  }
    
  // Google Document GAS用
  // (Browser.msgBoxが使用できないため)
  DocumentApp.getUi().alert("msg");

  SpreadsheetApp.getUi().alert("msg");
*/

/*
 オブジェクトのクラス名を取得
 "obj"に対象オブジェクトを指定
 
 var toString = Object.prototype.toString;
 Browser.msgBox(toString.call(obj));
*/

