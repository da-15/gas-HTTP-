var ROW_START = 2; //データの開始行を指定。
var COL_URL = 2; // HTTPステータスチェックを行いたいURL列
var COL_STATUS = 1; // ステータス結果を出力したい列

/*
 * メニューを追加
 */
function onOpen(){
  //メニュー配列
  var myMenu=[
    {name: "実行", functionName: "main"},
    {name: "ステータスクリア", functionName: "fncClear"}
  ];
  //メニューを追加
  SpreadsheetApp.getActiveSpreadsheet().addMenu("ステータスチェック",myMenu);

}

/*
 * メイン実行
 */
function main(){
  var i; 
  var strURL;
  var resCode
  var sheet = SpreadsheetApp.getActiveSheet();

  
  //ステータスカラムをクリア
  fncClear();
  
  for(i=ROW_START; i<=sheet.getLastRow(); i++){
    // HTTPステータスを取得
    strURL = sheet.getRange(i, COL_URL).getValue();
    resCode = getHTTPStatusCode(strURL);
    
    //結果を書き込み
    sheet.getRange(i, COL_STATUS).setValue(resCode);
    // エラー時の色変更（ステータス２００以外はエラー判定）
    if(resCode != 200){
      sheet.getRange(i, COL_STATUS).setBackground('#FF0000');
      sheet.getRange(i, COL_STATUS).setFontColor('#FFFFFF'); 
    }
  }
}

/*
 * HTTPステータスチェック
 */
function getHTTPStatusCode(strURL){
  var options = {
    "muteHttpExceptions": true,　    // 404エラーでも処理を継続する
  };
  try{
    return resCode = UrlFetchApp.fetch(strURL, options).getResponseCode();
  }
  catch(ex){
    return 999; // エラー時は９９９を返す
  }
}


/*
 * ステータスカラムのクリア
 */
function fncClear(){
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(ROW_START, COL_STATUS, sheet.getLastRow()).clear();

}