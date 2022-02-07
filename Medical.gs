// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
// 前回からの経過日数を計算
function calcDateInterval() {

  // ログフラグ
  var logFlag = Boolean("true");
  //var logFlag = Boolean("");

  // ---------- ---------- ----------  
  // 定数
  var intervalCol = 5;
  var clinicCol = 7;
  var firstRow = 2;
  var HsheetName = 'MedicalHistory';
  // ---------- ---------- ----------  

  // 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Gantt chart形式のシート
  var sheet = spreadsheet.getSheetByName(HsheetName);

  // Last Row
  var lastRow = sheet.getLastRow();
  
  for (currentRow=lastRow; currentRow>firstRow; currentRow--){

    // 経過日数を取得
    var intervalDate = sheet.getRange(currentRow, intervalCol).getValue();
    //var intervalDateF = sheet.getRange(currentRow, intervalCol).getFormula();
    
    // Nullチェック    
    if(!intervalDate){
      var clinicName = sheet.getRange(currentRow, clinicCol).getValue();
      var formerRow = sub_searchFormerHistory(sheet, clinicName, currentRow);
      
      if (formerRow >= firstRow) {
        var intervalFormula = "=E" + currentRow + "-E" + formerRow;

        if(logFlag){
          Logger.log("Row:%s\n Clinic Name:%s\n Former Row:%s\n intervalFormula)",
                     currentRow, clinicName, formerRow, intervalFormula);
        }

        sheet.getRange(currentRow, intervalCol).setFormula(intervalFormula);

      }
    }
  }    
}

// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
// 一回前の受診履歴を検索
function sub_searchFormerHistory(sheet, clinicName, currentRow){
  
  // ---------- ---------- ----------  
  // 定数
  var clinicCol = 8;
  var firstRow = 2;
  // ---------- ---------- ----------  
  
  for (fRow=currentRow-1; fRow>=firstRow; fRow--){
    var fclinicName = sheet.getRange(fRow, clinicCol).getValue();
    if (fclinicName === clinicName){
      return fRow;
    }
  }
  
  // 見つからなかった場合
  return 0;
}

// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
