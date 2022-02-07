// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 単語をランダムに選択
function randomSelectItem() {
  
  // ---------- ---------- ----------  
  // 定数
  wordCol = 1;
  posCol = 2;
  meaningCol = 3;
  // ---------- ---------- ----------  

  // 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 現在のシートを取得
  var sheet = spreadsheet.getActiveSheet();

  // 最終行
  var lastRow = sheet.getLastRow();

  flag = true
  while (flag) {
    row = Math.floor(Math.random()*lastRow)
    if (row != 1) {
      if ((!sheet.getRange(row, meaningCol).isBlank())
      && (!sheet.isRowHiddenByFilter(row))) {
        flag = false
        word = sheet.getRange(row, wordCol).getValue()
        pos = sheet.getRange(row, posCol).getValue()
        sheet.setActiveRange(sheet.getRange(row, 1))
        Browser.msgBox(word + " (" + pos + ")");
      }
    }
  }
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 単語をアルファベット順にソート
function sortVocabulary() {

  // Named Range(名前付きセル範囲)を修正
  fixNamedRanges();
    
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort({column: 1, ascending: true});

  // 重複項目を抽出
  checkDuplicateItems();
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// Named Range(名前付きセル範囲)を修正
function fixNamedRanges() {

  // ---------- ---------- ----------  
  // 定数
  vocabCol = 1;
  maxCount = 100;
  // ---------- ---------- ----------  
  
  // 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 現在のシートを取得
  var sheet = spreadsheet.getActiveSheet();
  
  // シート内の全ての名前付きセル範囲を取得
  var namedRanges = sheet.getNamedRanges();

  // 最終行
  var lastRow = sheet.getLastRow();

  if (namedRanges.length > 1) {
  
    for (i=0; i<namedRanges.length; i++) {
      var name = namedRanges[i].getName();
      var range = namedRanges[i].getRange();
      
      // 名前に"_"が付いている場合、取り除く
      if (name.match(/_/)) {
        name = name.replace(/_/,'');
      }
      
      // 名前付きセル範囲がズレている場合
      if (range.getValue() != name) {
        
        // 正しい名前付きセル範囲の位置を検索
        var count = 0;
        var row = range.getRow();
        var val = range.getValue();
        
        while (count <= maxCount) {
          count++;

          // 前にズレている場合(ex. text < U)
          if (val.charAt(0).toLowerCase() < name.charAt(0).toLowerCase()) {
            if (row === lastRow){
              Browser.msgBox("[Error]Not found until the end of file:" + name);
              return 1;
            }
            row++;
          } else {           
            // 後ろにズレている場合(ex. text > T)
            row--;
          }
          
          val = sheet.getRange(row, vocabCol).getValue();
          
          // 見つかった場合
          if (val === name) {
            var trueRange = sheet.getRange(row, vocabCol);
            
            // 名前付きセル範囲を補正
            namedRanges[i].setRange(trueRange);
            break;
          }            
        }
        
        // 見つからなかった場合
        if (val != name) {
          Browser.msgBox("[Error]Not found:" + name);
        }
      }
    }
  }
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 重複項目を抽出
function checkDuplicateItems() {

  // ---------- ---------- ----------  
  // 定数
  startRow = 3;
  vocabCol = 1;
  // ---------- ---------- ----------  
  
  // 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 現在のシートを取得
  var sheet = spreadsheet.getActiveSheet();
  
  // 最終行
  var lastRow = sheet.getLastRow();
  
  var dupVocabList = [];
  for (row=startRow; row<=lastRow-1; row++) {
    var vocab1 = sheet.getRange(row, vocabCol).getValue();
    var vocab2 = sheet.getRange(row+1, vocabCol).getValue();
    
    // 重複項目チェック
    if (vocab1 === vocab2) {
      dupVocabList.push("[" + row + "] " + vocab1);
      
      // 重複行を飛ばす
      row++;      
    }    
  }

  // 重複項目が見つかった場合
  if (dupVocabList.length > 0) {
    var msg = "";
    for(i=0; i<dupVocabList.length; i++) {
      // Browser.msgBoxの場合
      //msg += dupVocabList[i] + "\\n";
      msg += dupVocabList[i] + "\n";     
    }
    //Browser.msgBox("Same Vocaburay found\\n\\n" + msg);
    Logger.log("Duplicate vocaburay found (%s)\n\n%s", String(dupVocabList.length), msg);
  } else {
    Logger.log("No duplicate vocaburay");    
  }
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
