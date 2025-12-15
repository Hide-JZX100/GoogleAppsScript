/////////////////////////////////////////////
//発注情報詳細照会
//create:20160701 suzue
/////////////////////////////////////////////
function GetActiveCd() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1 = ss.getSheetByName('発注残一覧');
  var sheet2 = ss.getSheetByName('発注情報詳細');
  var sheet3 = ss.getSheetByName('入荷チェックリスト');
  var sheet4 = ss.getSheetByName('【寝装品】入荷チェックリスト');
  var wk3 = ss.getSheetByName('checklist_work');
  
  //照会情報クリア
  var r1 = sheet2.getLastRow();
  var row = 12;     //商品コード開始行
  sheet2.getRange(row, 4, r1, 2).clearContent();
  
  //発注番号取得
  var objCell = sheet1.getActiveCell();
  var orderCd = objCell.getValue();
  sheet2.getRange('B1').setValue(orderCd);

  //照会情報コピペ
  var r2 = sheet2.getRange('J1').getValue();
  r2 = r2 + 9;
  var data = sheet2.getRange(3, 1, r2, 2).getValues();
  sheet2.getRange(3, 4, r2, 2).setValues(data);
  
  //チェックリスト入力セルデータクリア
  sheet3.getRange('F3').clearContent();
  sheet3.getRange('K2').clearContent();
  
//  sheet4.getRange('C3').clearContent();
  sheet4.getRange('F3').clearContent();
  sheet4.getRange('L2').clearContent();

  //チェックリスト値セット
  var vndCd = sheet1.getRange('C2').getValue();
  var rowNm = sheet2.getRange('J1').getValue();
  if(vndCd == "0009" || vndCd == "0023" || vndCd == "0028" || vndCd == "0031"|| vndCd == "0032"){
    var colNm = 6;
    var r = sheet4.getLastRow();
    sheet4.getRange(6, 2, r, colNm).clearContent();
    var list = wk3.getRange(1, 1, rowNm, colNm).getValues();
    sheet4.getRange(6, 2, rowNm, colNm).setValues(list);
  }else{
    var colNm = 5;
    var r = sheet3.getLastRow();
    sheet3.getRange(6, 2, r, colNm).clearContent();
    var list = wk3.getRange(1, 1, rowNm, colNm).getValues();
    sheet3.getRange(6, 2, rowNm, colNm).setValues(list);    
  }

  //発注情報詳細シートアクティブ
  sheet2.getRange('E1').activate();

}

/////////////////////////////////////////////
//修正情報反映
//create:20160701 suzue
//update:20170518 suzue ドレージ情報更新処理追加
/////////////////////////////////////////////
function dataUpdate() {

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
  
  //データ件数
  var r1 = sheet.getRange('J1').getValue();

  var chk_list = sheet_To.getDataRange().getValues();  //コピー先発注番号
  var orderCd = sheet_From.getRange('A2').getValue();  //発注番号
  
  //delete開始行
  for(var i = 1; i < chk_list.length; i++){
      var chkCd = chk_list[i][0];  //発注番号
     if(chkCd == orderCd){
      var r2 = i + 1;
      break;
     }
  }

　//行削除
  sheet_To.deleteRows(r2, r1);
  
  //データコピー
  var last_row = sheet_To.getLastRow();
  var r3 = last_row + 1; //貼付け開始行
  
  var rowNum = sheet.getRange('J2').getValue();
  var colNum = sheet_From.getLastColumn();
  var data = sheet_From.getRange(2, 1, rowNum, colNum).getValues();
  sheet_To.getRange(r3, 1, rowNum, colNum).setValues(data);
  
  //ドレージ情報更新
  var doreji_data = doreji_from.getDataRange().getValues();
  
  var d_row = doreji_to.getLastRow();
  d_row = d_row + 1;   //貼付開始行
  var d_col = doreji_to.getLastColumn();
  
  doreji_to.getRange(d_row, 1, 1, d_col).setValues(doreji_data);
  
  //チェックリスト値セット
  var vndCd = sheet2.getRange('C2').getValue();
  var rowNm = sheet.getRange('J1').getValue();

  if(vndCd == "0009" || vndCd == "0023" || vndCd == "0028" || vndCd == "0031"){
    var colNm = 6;
    var r = sheet4.getLastRow();
    sheet4.getRange(6, 2, r, colNm).clearContent();
    var list = wk3.getRange(1, 1, rowNm, colNm).getValues();
    sheet4.getRange(6, 2, rowNm, colNm).setValues(list);
  }else{
    var colNm = 5;
    var r = sheet3.getLastRow();
    sheet3.getRange(6, 2, r, colNm).clearContent();
    var list = wk3.getRange(1, 1, rowNm, colNm).getValues();
    sheet3.getRange(6, 2, rowNm, colNm).setValues(list);    
  }
  
  Browser.msgBox('発注情報を修正しました');
 
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
  
  //チェックリスト値セット
  var vndCd = sheet1.getRange('C2').getValue();
  var rowNm = sheet2.getRange('J1').getValue();
  if(vndCd == "0009" || vndCd == "0020"){
    var colNm = 6;
    var r = sheet4.getLastRow();
    sheet4.getRange(6, 2, r, colNm).clearContent();
    var list = wk3.getRange(1, 1, rowNm, colNm).getValues();
    sheet4.getRange(6, 2, rowNm, colNm).setValues(list);
  }else{
    var colNm = 5;
    var r = sheet3.getLastRow();
    sheet3.getRange(6, 2, r, colNm).clearContent();
    var list = wk3.getRange(1, 1, rowNm, colNm).getValues();
    sheet3.getRange(6, 2, rowNm, colNm).setValues(list);    
  }

}