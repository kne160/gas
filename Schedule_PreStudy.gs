// ---------- ---------- ---------- ---------- ---------- ---------- ---------- 
// オートフィルタで全て表示
function afAll(){
  var spreadsheet = SpreadsheetApp.getActive();

  var criteria = SpreadsheetApp.newFilterCriteria()
  .build();

  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(6, criteria);
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// オートフィルタでアクティブなタスクのみ表示
function afActive(){
  var spreadsheet = SpreadsheetApp.getActive();

  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['_C', 'F', 'K'])
  .build();

  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(6, criteria);
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// オートフィルタでアクティブ及びKeepタスクを表示
function afActiveKeep(){
  var spreadsheet = SpreadsheetApp.getActive();

  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['_C', 'F'])
  .build();

  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(6, criteria);
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// トータル時間の一致チェック
function checkTotalTime() {
  
  // ---------- ---------- ----------  
  // 定数
  tTotalTimeRow = 2;  
  tTotalTimeCol = 10;
  sTotalTimeRow = 1;  
  sTotalTimeCol = 8;
  
  sStartRow = 2;
  tStartRow = 3;
  
  sTaskCol = 6;
  tTaskCol = 3;
  
  sDateCol = 2;

  const scheduleSheetName = "Schedule_L";
  // ---------- ---------- ----------  

  // 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 現在のシートを取得
  var taskSheet = spreadsheet.getActiveSheet();
  
  // Scheduleシート
  var scheduleSheet = spreadsheet.getSheetByName(scheduleSheetName);  
  
  // トータル時間を取得
  tTotalTime = taskSheet.getRange(tTotalTimeRow, tTotalTimeCol).getDisplayValue();
  sTotalTime = scheduleSheet.getRange(sTotalTimeRow, sTotalTimeCol).getDisplayValue();

  // トータル時間を比較
  if (tTotalTime != sTotalTime) {
    Browser.msgBox("Total Times are unmatched\\n\\n[Task]: " + tTotalTime + "\\n[Schedule]: " + sTotalTime);    

    // 最終行  
    var sLastRow = scheduleSheet.getLastRow();
    var tLastRow = taskSheet.getLastRow();
    
    // TaskシートのTaskリストを作成
    var taskList = new Set();

    for(tRow=tStartRow; tRow<=tLastRow; tRow++) {
      tTask = taskSheet.getRange(tRow, tTaskCol).getValue();
      if(tTask) {
        taskList.add(tTask);
      }
    }
        
    // ScheduleシートのTaskがTaskシートにあるかチェック
    for(sRow=sStartRow; sRow<=sLastRow; sRow++){
      sTask = scheduleSheet.getRange(sRow, sTaskCol).getValue();

      // Taskシートに無い
      if(sTask && !taskList.has(sTask)) {
        var date = scheduleSheet.getRange(sRow, sDateCol).getValue();
        Browser.msgBox("Unlisted task is found:\\n\\n[" + sRow + "] (" + Utilities.formatDate(date, 'Asia/Tokyo', 'dd/M/YYYY') + ") " + sTask);
      }
    }

  } else {
    Browser.msgBox("Match both Total Times\\n\\n[Task]: " + tTotalTime + "\\n[Schedule]: " + sTotalTime);    
  }    
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 変更(行挿入 / 削除)時の補正
function onEdit(e){

  // ---------- ---------- ----------  
  // 定数
  const DowCol = 1;
  const DateCol = 2;
  const PCol = 5;
  
  const TaskCountCol = 9;
  const TaskTimeCol = 10;

  const ErrorFormula = "#REF!";
  const snSchedule_L = "Schedule_L";
  const snTask = "Task";  
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
  
  // シートが補正対象をチェック
  switch (sheetName) {
      
      // [Schedule_L]シート
    case snSchedule_L:
      
      // 変更範囲
      for (var row = rowStart; row <= rowEnd; row++) {
        
        // 1行目はスキップ
        if (row === 1){
          continue;
        }
        
        // 日付列に数式が設定されてない場合
        if (!sheet.getRange(row, DateCol).getFormula()){
          sheet.getRange(row, DowCol).setFormula("=A" + String(row-1));
          sheet.getRange(row, DateCol).setFormula("=B" + String(row-1));
        }
        
        // タスク作業時間列に数式が設定されていない場合
        if (!sheet.getRange(row, PCol).getFormula()){
          // 前日の数式を取得
          var formarFormula = sheet.getRange(row-1, PCol).getFormulaR1C1();
          sheet.getRange(row, PCol).setFormulaR1C1(formarFormula);        
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
      break;
      
      // [Task]シート
    case snTask:
      // 変更範囲
      for (var row = rowStart; row <= rowEnd; row++) {
        
        // 3行目まではスキップ
        if (row <= 3) {
          continue;
        }
        
        // タスク実施回数 / 作業時間列に数式が設定されていない場合
        if ((!sheet.getRange(row, TaskCountCol).getFormula())
            && (!sheet.getRange(row, TaskTimeCol).getFormula()) ) {
          sheet.getRange(row, TaskCountCol).setFormulaR1C1(sheet.getRange(row-1, TaskCountCol).getFormulaR1C1());
          sheet.getRange(row, TaskTimeCol).setFormulaR1C1(sheet.getRange(row-1, TaskTimeCol).getFormulaR1C1());
        }       
      }
      break;     
  }  
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 変更(行挿入 / 削除)時の補正
function onEdit(e){

  // ---------- ---------- ----------  
  // 定数
  const DowCol = 1;
  const DateCol = 2;
  const PCol = 5;
  
  const TaskCountCol = 9;
  const TaskTimeCol = 10;

  const ErrorFormula = "#REF!";
  const snSchedule_L = "Schedule_L";
  const snTask = "Task";  
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
  
  // シートが補正対象をチェック
  switch (sheetName) {
      
      // [Schedule_L]シート
    case snSchedule_L:
      
      // 変更範囲
      for (var row = rowStart; row <= rowEnd; row++) {
        
        // 1行目はスキップ
        if (row === 1){
          continue;
        }
        
        // 日付列に数式が設定されてない場合
        if (!sheet.getRange(row, DateCol).getFormula()){
          sheet.getRange(row, DowCol).setFormula("=A" + String(row-1));
          sheet.getRange(row, DateCol).setFormula("=B" + String(row-1));
        }
        
        // タスク作業時間列に数式が設定されていない場合
        if (!sheet.getRange(row, PCol).getFormula()){
          // 前日の数式を取得
          var formarFormula = sheet.getRange(row-1, PCol).getFormulaR1C1();
          sheet.getRange(row, PCol).setFormulaR1C1(formarFormula);        
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
      break;
      
      // [Task]シート
    case snTask:
      // 変更範囲
      for (var row = rowStart; row <= rowEnd; row++) {
        
        // 3行目まではスキップ
        if (row <= 3) {
          continue;
        }
        
        // タスク実施回数 / 作業時間列に数式が設定されていない場合
        if ((!sheet.getRange(row, TaskCountCol).getFormula())
            && (!sheet.getRange(row, TaskTimeCol).getFormula()) ) {
          sheet.getRange(row, TaskCountCol).setFormulaR1C1(sheet.getRange(row-1, TaskCountCol).getFormulaR1C1());
          sheet.getRange(row, TaskTimeCol).setFormulaR1C1(sheet.getRange(row-1, TaskTimeCol).getFormulaR1C1());
        }       
      }
      break;     
  }  
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 変更(行挿入 / 削除)時の補正
function onEdit(e){

  // ---------- ---------- ----------  
  // 定数
  const DowCol = 1;
  const DateCol = 2;
  const PCol = 5;
  
  const TaskCountCol = 9;
  const TaskTimeCol = 10;

  const ErrorFormula = "#REF!";
  const snSchedule_L = "Schedule_L";
  const snTask = "Task";  
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
  
  // シートが補正対象をチェック
  switch (sheetName) {
      
      // [Schedule_L]シート
    case snSchedule_L:
      
      // 変更範囲
      for (var row = rowStart; row <= rowEnd; row++) {
        
        // 1行目はスキップ
        if (row === 1){
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
      break;
      
      // [Task]シート
    case snTask:
      // 変更範囲
      for (var row = rowStart; row <= rowEnd; row++) {
        
        // 3行目まではスキップ
        if (row <= 3) {
          continue;
        }
        
        // タスク実施回数 / 作業時間列に数式が設定されていない場合
        if ((!sheet.getRange(row, TaskCountCol).getFormula())
            && (!sheet.getRange(row, TaskTimeCol).getFormula()) ) {
          sheet.getRange(row, TaskCountCol).setFormulaR1C1(sheet.getRange(row-1, TaskCountCol).getFormulaR1C1());
          sheet.getRange(row, TaskTimeCol).setFormulaR1C1(sheet.getRange(row-1, TaskTimeCol).getFormulaR1C1());
        }       
      }
      break;     
  }  
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
