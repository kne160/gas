// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
// 日付を全行に入力
function setDuplicateDate() {
  // ログフラグ
  var logFlag = Boolean("true");
  //var logFlag = Boolean("");
  
  // ---------- ---------- ----------  
  // 定数
  var sheetNameArray = ['2019_6', '2019_7', '2019_8', '2019_9', '2019_10', '2019_11', '2019_12', '2020_1'];
  
  var startRow = 3;
  var dateCol = 2;
  var dowCol = 1;
  // ---------- ---------- ----------  

  // 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  for (var i=0; i<sheetNameArray.length; i++) {
    // シートを取得
    var sheet = spreadsheet.getSheetByName(sheetNameArray[i]);

    // End Cell
    var lastRow = sheet.getLastRow();
    
    for (currentRow=startRow; currentRow<=lastRow; currentRow++){
      var dateCell = sheet.getRange(currentRow, dateCol).getValue();
      
      if (!dateCell) {
        var dateFormula = "=B";
        dateFormula += currentRow-1;
        
        var dowFormula = "=A";
        dowFormula += currentRow-1;
        
        if(logFlag){
          Logger.log("Row:%s\n dateFormula:%s\n dowFormula:%s",
                     currentRow, dateFormula, dowFormula);
        }
                
        sheet.getRange(currentRow, dateCol).setFormula(dateFormula);
        sheet.getRange(currentRow, dowCol).setFormula(dowFormula);
      }
    }    
  } 
}

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

