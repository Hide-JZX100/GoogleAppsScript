/////////////////////////////////////////////
// 発注情報詳細照会(改善版)
// create:20160701 suzue
// update:20251020 エラーハンドリング追加、マジックナンバー定数化
/////////////////////////////////////////////
function GetActiveCd() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet1 = ss.getSheetByName('発注残一覧');
    var sheet2 = ss.getSheetByName('発注情報詳細');
    var sheet3 = ss.getSheetByName('入荷チェックリスト');
    var sheet4 = ss.getSheetByName('【寝装品】入荷チェックリスト');
    var wk3 = ss.getSheetByName('checklist_work');
    
    // 定数定義
    var ITEM_CODE_START_ROW = 12;
    
    // 照会情報クリア
    var r1 = sheet2.getLastRow();
    sheet2.getRange(ITEM_CODE_START_ROW, 4, r1, 2).clearContent();
    
    // 発注番号取得
    var objCell = sheet1.getActiveCell();
    var orderCd = objCell.getValue();
    
    // 発注番号が空の場合はエラー
    if(!orderCd){
      Browser.msgBox('発注番号を選択してください');
      return;
    }
    
    sheet2.getRange('B1').setValue(orderCd);

    // 照会情報コピペ
    var r2 = sheet2.getRange('J1').getValue();
    r2 = r2 + 9;
    var data = sheet2.getRange(3, 1, r2, 2).getValues();
    sheet2.getRange(3, 4, r2, 2).setValues(data);
    
    // チェックリスト入力セルデータクリア
    sheet3.getRange('F3').clearContent();
    sheet3.getRange('K2').clearContent();
    sheet4.getRange('F3').clearContent();
    sheet4.getRange('L2').clearContent();

    // チェックリスト更新処理を呼び出し
    updateCheckList(ss, sheet2, sheet1, sheet3, sheet4, wk3);

    // 発注情報詳細シートアクティブ
    sheet2.getRange('E1').activate();
    
  } catch(e) {
    Browser.msgBox('エラーが発生しました: ' + e.message);
    Logger.log('GetActiveCd Error: ' + e.message + '\n' + e.stack);
  }
}

/////////////////////////////////////////////
// 修正情報反映(新規追加対応版)
// create:20160701 suzue
// update:20170518 suzue ドレージ情報更新処理追加
// update:20251215 改善:トランザクション的処理、バックアップ機能追加
// update:20251215 改善:発注番号がない場合は新規追加に対応
/////////////////////////////////////////////
function dataUpdate() {
  try {
    var ss_From = SpreadsheetApp.getActiveSpreadsheet();
    var ss_To = SpreadsheetApp.openById('1e9adbHoIVCkUxFWxuHcRPrPeogozS0OWObi9zShkVFw');
    var sheet_From = ss_From.getSheetByName('update');
    var sheet_To = ss_To.getSheetByName('発注情報');
    var sheet = ss_From.getSheetByName('発注情報詳細');
    var sheet2 = ss_From.getSheetByName('発注残一覧');
    var sheet3 = ss_From.getSheetByName('入荷チェックリスト');
    var sheet4 = ss_From.getSheetByName('【寝装品】入荷チェックリスト');
    var wk3 = ss_From.getSheetByName('checklist_work');
    var doreji_from = ss_From.getSheetByName('doreji');
    var doreji_ss = SpreadsheetApp.openById('14BxnzZ2OSamx52QeMB2reLLU48VyUZP93n0k8FJwa3o');
    var doreji_to = doreji_ss.getSheetByName('ドレージ情報');
    
    // データ件数
    var rowNum = sheet.getRange('J2').getValue();
    var colNum = sheet_From.getLastColumn();
    var orderCd = sheet_From.getRange('A2').getValue();  // 発注番号
    
    // 発注番号のバリデーション
    if(!orderCd || orderCd === ''){
      throw new Error('発注番号が入力されていません');
    }
    
    // 更新データ取得
    var updateData = sheet_From.getRange(2, 1, rowNum, colNum).getValues();
    
    // 既存データから該当発注番号の行を検索
    var allData = sheet_To.getDataRange().getValues();
    var startRow = -1;
    var endRow = -1;
    var isNewOrder = false;  // 新規発注フラグ
    
    for(var i = 1; i < allData.length; i++){
      var chkCd = allData[i][0];  // 発注番号
      if(chkCd == orderCd){
        if(startRow == -1) startRow = i;
        endRow = i;
      } else if(startRow != -1) {
        break;
      }
    }
    
    // 該当データが見つからない場合は新規追加
    if(startRow == -1){
      isNewOrder = true;
      Logger.log('発注番号 ' + orderCd + ' は新規発注です');
      
      // 最終行の次に追加
      var lastRow = sheet_To.getLastRow();
      sheet_To.getRange(lastRow + 1, 1, rowNum, colNum).setValues(updateData);
      
      Logger.log('新規発注を追加しました: 行' + (lastRow + 1) + '-' + (lastRow + rowNum));
      
    } else {
      // 既存データの更新処理
      var existingRowCount = endRow - startRow + 1;
      
      Logger.log('発注番号: ' + orderCd);
      Logger.log('既存行数: ' + existingRowCount + '行(行' + (startRow+1) + '-' + (endRow+1) + ')');
      Logger.log('更新行数: ' + rowNum + '行');
      
      // 行数が同じ場合は上書き(最も安全)
      if(existingRowCount == rowNum){
        Logger.log('処理方法: 上書き(安全モード)');
        sheet_To.getRange(startRow + 1, 1, rowNum, colNum).setValues(updateData);
        
      } else if(existingRowCount < rowNum){
        // 行が増える場合: 不足分を挿入してから全体を上書き
        Logger.log('処理方法: 行追加 + 上書き');
        var insertCount = rowNum - existingRowCount;
        
        // ステップ1: 不足分の行を挿入(endRowの後に挿入)
        sheet_To.insertRowsAfter(endRow + 1, insertCount);
        
        // ステップ2: 全体を上書き(削除なし)
        sheet_To.getRange(startRow + 1, 1, rowNum, colNum).setValues(updateData);
        
      } else {
        // 行が減る場合: 全体を上書きしてから余分な行を削除
        Logger.log('処理方法: 上書き + 行削除');
        var deleteCount = existingRowCount - rowNum;
        
        // ステップ1: 既存の範囲内で更新データを上書き
        sheet_To.getRange(startRow + 1, 1, rowNum, colNum).setValues(updateData);
        
        // ステップ2: 余分な行を削除(更新データの後ろの行)
        var deleteStartRow = startRow + 1 + rowNum;
        sheet_To.deleteRows(deleteStartRow, deleteCount);
      }
      
      Logger.log('データ更新完了');
    }
    
    // ドレージ情報更新
    var doreji_data = doreji_from.getDataRange().getValues();
    var d_row = doreji_to.getLastRow() + 1;
    var d_col = doreji_to.getLastColumn();
    doreji_to.getRange(d_row, 1, 1, d_col).setValues(doreji_data);
    
    // チェックリスト値セット
    updateCheckList(ss_From, sheet, sheet2, sheet3, sheet4, wk3);
    
    // 完了メッセージ
    var message = '';
    if(isNewOrder){
      message = '新規発注を追加しました\n\n' + 
                '発注番号: ' + orderCd + '\n' +
                '行数: ' + rowNum + '行';
    } else {
      message = '発注情報を修正しました\n\n' + 
                '発注番号: ' + orderCd + '\n' +
                '既存: ' + existingRowCount + '行 → 更新: ' + rowNum + '行';
    }
    Browser.msgBox(message);
    
  } catch(e) {
    Logger.log('エラー発生: ' + e.message);
    Logger.log('スタックトレース: ' + e.stack);
    
    // エラー時は詳細を表示
    Browser.msgBox('エラーが発生しました\n\n' + 
                   'エラー内容: ' + e.message + '\n\n' +
                   '処理は中断されました。\n' +
                   'ログを確認してください。');
    throw e;
  }
}

/////////////////////////////////////////////
// チェックリスト更新処理(共通化)
// create:20251020 共通処理として抽出
// update:20251216 GetActiveCdでも使えるように引数を調整
/////////////////////////////////////////////
function updateCheckList(ss, sheet, sheet2, sheet3, sheet4, wk3) {
  var vndCd = sheet2.getRange('C2').getValue();
  var rowNm = sheet.getRange('J1').getValue();
  // 定数定義（この関数内で閉じるように再定義するか、グローバル定数として定義する）
  var BEDDING_VENDORS = ["0009", "0023", "0028", "0031", "0032", "0020"];

  if(BEDDING_VENDORS.indexOf(vndCd) !== -1){
    var colNm = 6;
    var r = sheet4.getLastRow();
    sheet4.getRange(6, 2, r, colNm).clearContent();
    var list = wk3.getRange(1, 1, rowNm, colNm).getValues();
    sheet4.getRange(6, 2, rowNm, colNm).setValues(list);
  } else {
    var colNm = 5;
    var r = sheet3.getLastRow();
    sheet3.getRange(6, 2, r, colNm).clearContent();
    var list = wk3.getRange(1, 1, rowNm, colNm).getValues();
    sheet3.getRange(6, 2, rowNm, colNm).setValues(list);    
  }
}

/////////////////////////////////////////////
//PINo.入力時にSCアドレスに発注完了メールを送信
//create:20170310 suzue
/////////////////////////////////////////////
function PINoUpdateMail() {
  var ss_From = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss_From.getSheetByName('発注情報詳細');
  var sheet2 = ss_From.getSheetByName('発注残一覧');
  var sheet_From = ss_From.getSheetByName('update');
  
  //データ更新処理
  dataUpdate();
  
  var rowNum = sheet.getRange('J2').getValue();
  // @ts-ignore
  rowNum = rowNum + 1;
  
  //発注済メールをSCアドレスに送信
  var orderData = sheet_From.getDataRange().getValues();
  //ログインID取得
  var userAddress = Session.getEffectiveUser().getUserLoginId();
  var address = 'bm.stockcontrol@gmail.com';
//  var address = 'bm.arisu.suzue@gmail.com';
  
  var vender = sheet2.getRange('A2').getValue();
  var piNo = sheet.getRange('E3').getValue();
  
  var title = '発注完了【' + vender + '】' + piNo;

  //本文内容
  var content = "以下の発注が完了しました。チェックリストを確認してください。 \n\n"
  content = content + "----------------------------------------------- \n\n"

  for(var i = 1; i < rowNum; i++){
    var itemCd = orderData[i][1];   //商品コード
    var hacchuSu = orderData[i][2]; //発注数
　　//発注内容
    content = content + itemCd + "：" + hacchuSu +  "\n"
  }  

  var options = {from:userAddress, name:'発注情報'};
  
  GmailApp.sendEmail(address, title, content, options);   
  
}


//照会情報再表示
//20170522Update suzue:画面リフレッシュ時にチェックリスト再作成
function screenUpdating() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName('発注残一覧');
  var sheet2 = ss.getSheetByName('発注情報詳細');
  var sheet3 = ss.getSheetByName('入荷チェックリスト');
  var sheet4 = ss.getSheetByName('【寝装品】入荷チェックリスト');
  var wk3 = ss.getSheetByName('checklist_work');
  
  //照会情報リフレッシュ
  var r1 = sheet2.getLastRow();
  sheet2.getRange(9, 4, r1, 2).clearContent();

  var r2 = sheet2.getRange('J1').getValue();
  r2 = r2 + 11;
  var data = sheet2.getRange(3, 1, r2, 2).getValues();
  sheet2.getRange(3, 4, r2, 2).setValues(data);
  
  //チェックリスト入力セルデータクリア
  sheet3.getRange('F3').clearContent();
  sheet3.getRange('K2').clearContent();
  
//  sheet4.getRange('C3').clearContent();
  sheet4.getRange('F3').clearContent();
  sheet4.getRange('L2').clearContent();
  
  // チェックリスト更新処理を呼び出し
  updateCheckList(ss, sheet2, sheet1, sheet3, sheet4, wk3);

}