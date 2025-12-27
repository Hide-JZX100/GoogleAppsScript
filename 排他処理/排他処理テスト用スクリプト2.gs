/**
 * プロジェクトBからの書き込みテスト
 * testFromProjectA実行中に実行してください
 */
const TARGET_SPREADSHEET_ID = "貴方のスプレッドシートIDをここに入れてください";

function testFromProjectB() {
  Logger.log("=== プロジェクトB: 書き込み開始 ===");
  
  // ★ 重要: lockIdを受け取る
  var lockId = LockLib.acquireSpreadsheetLockById(TARGET_SPREADSHEET_ID);
  
  if (lockId) {  // ★ nullチェック
    Logger.log("プロジェクトB: ロック取得成功 (ID: " + lockId + ")");
    
    try {
      var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
      var sheet = ss.getSheetByName("test");
      
      // 書き込み
      for (var i = 1; i <= 5; i++) {
        sheet.getRange("B" + i).setValue("ProjectB-" + new Date().toLocaleTimeString());
        Logger.log("プロジェクトB: " + i + "行目書き込み");
        Utilities.sleep(1000);
      }
      
    } finally {
      LockLib.releaseSpreadsheetLockById(TARGET_SPREADSHEET_ID, lockId);
      Logger.log("プロジェクトB: ロック解放");
    }
  } else {
    Logger.log("プロジェクトB: ロック取得失敗(タイムアウト)");
  }
}

function testFromProjectBWithDetailedLog() {
  Logger.log("=== プロジェクトB: 書き込み開始 ===");
  Logger.log("開始時刻: " + new Date().toLocaleTimeString());
  
  var startTime = new Date().getTime();
  
  // タイムアウトを60秒に設定(十分な余裕を持たせる)
  if (LockLib.acquireSpreadsheetLockById(TARGET_SPREADSHEET_ID, "_LOCK_", 45)) {
    var waitTime = (new Date().getTime() - startTime) / 1000;
    Logger.log("プロジェクトB: ロック取得成功");
    Logger.log("ロック取得までの待機時間: " + waitTime.toFixed(2) + "秒");
    Logger.log("取得時刻: " + new Date().toLocaleTimeString());
    
    try {
      var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
      var sheet = ss.getSheetByName("test");
      
      // 書き込み
      for (var i = 1; i <= 5; i++) {
        var timestamp = new Date().toLocaleTimeString();
        sheet.getRange("B" + i).setValue("ProjectB-" + timestamp);
        Logger.log("プロジェクトB: " + i + "行目書き込み (" + timestamp + ")");
        Utilities.sleep(1000);
      }
      
      Logger.log("プロジェクトB: 書き込み完了");
      
    } finally {
      LockLib.releaseSpreadsheetLockById(TARGET_SPREADSHEET_ID, lockId);
      Logger.log("プロジェクトB: ロック解放");
      Logger.log("終了時刻: " + new Date().toLocaleTimeString());
      
      var totalTime = (new Date().getTime() - startTime) / 1000;
      Logger.log("総実行時間: " + totalTime.toFixed(2) + "秒");
    }
  } else {
    var waitTime = (new Date().getTime() - startTime) / 1000;
    Logger.log("プロジェクトB: ロック取得失敗(タイムアウト)");
    Logger.log("待機時間: " + waitTime.toFixed(2) + "秒");
  }
}