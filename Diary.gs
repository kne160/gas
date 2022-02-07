// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 新規日付を追加
function newDate() {

  // 文字数が未カウントの日記をチェックしてカウント
  sub_countWordsEachDay();
  
  // 今日の項目を追加
  sub_insertNewDate();
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 今日の項目を追加
function sub_insertNewDate() {
  // ---------- ---------- ----------  
  // 定数
  const DOW_LIST = new Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat");
  const NEW_INDEX = 5;
  // ---------- ---------- ----------  

  /*
   * 横線を引く (未使用)
   *
   var cursor = DocumentApp.getActiveDocument().getCursor();    
   var element = cursor.getElement();
   var parent = element.getParent();
   body.insertHorizontalRule(parent.getChildIndex(element));
  */

  // ドキュメントを取得
  var doc = DocumentApp.getActiveDocument();
  
  // Bodyを取得
  var body = doc.getBody();

  // 今日の文字列(曜日 日付)を生成
  var today = new Date();
  var dow_str = DOW_LIST[today.getDay()];
  var day = today.getDate();
  
  // 今日の項目を追加
  var newday_h = body.insertParagraph(NEW_INDEX, dow_str + " " + day + " ()");
  newday_h.setHeading(DocumentApp.ParagraphHeading.HEADING2);

  var index = NEW_INDEX;

  index++;
  body.insertParagraph(index, "");
  
  index++;
  body.insertParagraph(index, "Activities\n\n");

  index++;
  var newday_b = body.insertParagraph(index, "Interview questions\n\n\n");

  // 横線
  newday_b.appendHorizontalRule();

  // カーソル位置をセット
  var position = doc.newPosition(newday_b.getChild(0), 1);
  doc.setCursor(position);
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 一日分の文字数をカウント
function sub_countWordsEachDay() {
  
  // Private helper function
  // find text length of paragraph
  function paragraphLen( par ) {
    return par.asText().getText().length;
  }
  
  var doc = DocumentApp.getActiveDocument();
  var paragraphs = doc.getBody().getParagraphs();
  
  // Scan document
  for (i=0; i<paragraphs.length; i++) {    
    if (paragraphLen(paragraphs[i]) > 0) {
      // This paragraph has text
      var paragraphText = paragraphs[i].asText().getText();

      // 各日付タイトル and 文字数が空欄
      if ((paragraphs[i].getHeading() === DocumentApp.ParagraphHeading.HEADING2)
          && (paragraphText.match(/\(\)/))) {

        // HEADING2パラグラフをスキップ
        var j = i+1;
        var total = 0;
        
        // 次の日付(HEADING2)パラグラフまで
        while (!(
          (!paragraphs[j].findElement(DocumentApp.ElementType.HORIZONTAL_RULE)) &&
          (paragraphs[j].getHeading() === DocumentApp.ParagraphHeading.HEADING2))) {
          if (paragraphLen(paragraphs[j]) > 0) {
            // This paragraph has text
            var num = sub_countWordsText(paragraphs[j].asText().getText());
            total += num;
          }
          j++;
        }
        // 文字数を出力
        //DocumentApp.getUi().alert(paragraphText + ": " + total);
        paragraphs[i].replaceText("\\(\\)", "(" + total + ")");        
                
        // 最後の空行を除去
        
        // 処理フラグ
        var Flag_removeBlankline = true;
        
        // 本文がある場合(文字数が1以上)
        if (total > 0) {

          // 最後のパラグラフから
          var k = j-1;
                 
          // 区切り線パラグラフをチェック
          var rangeElement = paragraphs[k].findElement(DocumentApp.ElementType.HORIZONTAL_RULE);
          if (rangeElement) {          

            // 区切り線パラグラフに複数エレメントがある場合
            // (実際には最大2エレメントまでのはず)
            if (paragraphs[k].getNumChildren() > 1) {
              var sibling = rangeElement.getElement().getPreviousSibling();
              
              if (sibling) {
                if (sibling.asText().getText().trim().length === 0) {
                  // 空行のみのパラグラフは除去
                  sibling.removeFromParent();
                } else {
                  // 末尾の空行を取り除く
                  // trim()の場合、前方の改行も取り除かれる
                  //sibling.asText().setText(sibling.asText().getText().trim());
                  sibling.asText().setText(sibling.asText().getText().replace(/\r$/, ""));
                  
                  // 空行以外のパラグラフがすでに存在するため、空行除去は終了
                  Flag_removeBlankline = false;
                }
              }
            }
          }

          // 空行以外のパラグラフが見つかるまで、空行のみのパラグラフを除去
          if (Flag_removeBlankline) {
            // 区切り線パラグラフの一つ前から
            k--;
            
            while (paragraphs[k].asText().getText().trim().length === 0) {
              // 空行のみのパラグラフを除去
              paragraphs[k].removeFromParent();
              k--;
            }
          }
          
        }
      }
    }
  }
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// 文字数をカウント
function sub_countWordsText(text) {

  //A simple \n replacement didn't work, neither did \s not sure why
  s = text.replace(/\r\n|\r|\n/g, " ");

  //In cases where you have "...last word.First word..." 
  //it doesn't count the two words around the period.
  //so I replace all punctuation with a space
  var punctuationless = s.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()"?“”]/g, " ");

  //Finally, trim it down to single spaces (not sure this even matters)
  var finalString = punctuationless.replace(/\s{2,}/g, " ");

  //Actually count it
  var count = finalString.trim().split(/\s+/).length; 

  return count;
}  
  
// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// GASメニューを追加
function onOpen() {
  // ---------- ---------- ----------  
  // 定数
  const FILENAME_DIARY = "Diary";
  // ---------- ---------- ----------  

  // ファイル名
  var name = DocumentApp.getActiveDocument().getName();

  var reg = new RegExp(FILENAME_DIARY);
  
  // ファイル名が日記の場合
  if (name.match(reg)) {
    var ui = DocumentApp.getUi();
    ui.createMenu('GAS')
    .addItem('Run newDate()', 'newDate')
    .addToUi();
  }
}

// ---------- ---------- ---------- ---------- ---------- ---------- ----------
// snippet
// ---------- ---------- ---------- ---------- ---------- ---------- ----------

/*
 ログ出力

 DocumentApp.getUi().alert();
 
*/
