/**
 * プロジェクトAからの書き込みテスト
 * このスクリプトを別のGASプロジェクトに配置
 */
const TARGET_SPREADSHEET_ID = "貴方のスプレッドシートIDをここに入れてください";

function testFromProjectA() {
  Logger.log("=== プロジェクトA: 書き込み開始 ===");
  
  // ライブラリを使用(識別子をLockLibと仮定)
  if (LockLib.acquireSpreadsheetLockById(TARGET_SPREADSHEET_ID)) {
    Logger.log("プロジェクトA: ロック取得成功");
    
    try {
      var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
      var sheet = ss.getSheetByName("test");
      
      // 10秒間ロックを保持(プロジェクトBがブロックされることを確認)
      for (var i = 1; i <= 10; i++) {
        sheet.getRange("A" + i).setValue("ProjectA-" + new Date().toLocaleTimeString());
        Logger.log("プロジェクトA: " + i + "行目書き込み");
        Utilities.sleep(1000);
      }
      
    } finally {
      LockLib.releaseSpreadsheetLockById(TARGET_SPREADSHEET_ID, lockId);
      Logger.log("プロジェクトA: ロック解放");
    }
  } else {
    Logger.log("プロジェクトA: ロック取得失敗");
  }
}