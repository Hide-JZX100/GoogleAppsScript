/**
 * @fileoverview Google Apps Script (GAS) を用いたスプレッドシートの排他処理ライブラリ
 *
 * @description
 * 複数のスクリプトやユーザーが同時にスプレッドシートを編集しようとすると発生する
 * 競合状態（レースコンディション）を防ぐための排他制御（ロック）機構を提供します。
 * スプレッドシートに特定の管理シート（デフォルト: `_LOCK_`）を作成し、
 * そのシートの状態を読み書きすることでロックを実現します。
 * GASの `LockService` と組み合わせることで、管理シートへのアクセス自体も保護し、信頼性を高めています。
 *
 * @usage
 *
 * ### ライブラリとして利用する場合 (推奨)
 * 1. このスクリプトをGASプロジェクトとして保存し、[デプロイ] > [新しいデプロイ] からライブラリとしてデプロイします。
 * 2. 別のGASプロジェクトで、[ライブラリ] > [+] からデプロイしたスクリプトのIDを追加します。(例: 識別子を `LockLib` とする)
 * 3. `safeWriteToSheet` 関数や、`acquireSpreadsheetLockById` / `releaseSpreadsheetLockById` を使用して排他処理を実装します。
 *
 *    // 例:
 *    const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID";
 *    if (LockLib.acquireSpreadsheetLockById(SPREADSHEET_ID)) {
 *      try {
 *        // ... ここに排他制御したい処理を記述 ...
 *      } finally {
 *        LockLib.releaseSpreadsheetLockById(SPREADSHEET_ID);
 *      }
 *    }
 *
 * ### スクリプトを直接コピーして利用する場合
 * 1. このスクリプトの内容を、利用したいスプレッドシートのコンテナバインドスクリプトにコピーします。
 * 2. `acquireSpreadsheetLock` / `releaseSpreadsheetLock` を使用して処理を囲みます。
 *
 * @functions
 * - `acquireSpreadsheetLock`: コンテナバインドスクリプトで、現在開いているスプレッドシートのロックを取得します。
 * - `acquireSpreadsheetLockById`: スプレッドシートIDを指定してロックを取得します。スタンドアロン環境や別ファイルからの操作に利用します。
 * - `releaseSpreadsheetLock`: コンテナバインドスクリプトでロックを解放します。
 * - `releaseSpreadsheetLockById`: スプレッドシートIDを指定してロックを解放します。
 * - `safeWriteToSheet`: ロックの取得、処理の実行、ロックの解放をまとめて行う安全な書き込み関数です。基本的にはこの関数の利用を推奨します。
 * - `forceUnlock`: 何らかのエラーでロックが残ってしまった場合に、強制的に解除するための関数です。
 * - `testCompleteLockMechanism`: ロック機構全体の動作をテストするための関数です。
 *
 * @see
 * このスクリプトは `TARGET_SPREADSHEET_ID` をテストおよび強制解除の際に使用します。
 * 利用する際は、対象のスプレッドシートIDに書き換えてください。
 */

const TARGET_SPREADSHEET_ID = "貴方のスプレッドシートIDをここに入れてください";


/**
 * スプレッドシート上でセマフォ(ロック)を取得する
 * コンテナバインドスクリプト用(getActiveSpreadsheet使用)
 * @param {string} lockSheetName - ロック管理用シート名(デフォルト: "_LOCK_")
 * @param {number} timeoutSeconds - タイムアウト秒数(デフォルト: 30秒)
 * @returns {boolean} - ロック取得成功時true、失敗時false
 */
function acquireSpreadsheetLock(lockSheetName = "_LOCK_", timeoutSeconds = 30) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lockSheet = ss.getSheetByName(lockSheetName);
  
  // ロック管理シートが存在しない場合は作成
  if (!lockSheet) {
    lockSheet = ss.insertSheet(lockSheetName);
    lockSheet.getRange("A1").setValue("ロック状態");
    lockSheet.getRange("B1").setValue("取得時刻");
    lockSheet.getRange("C1").setValue("プロジェクト名");
    lockSheet.getRange("D1").setValue("実行ユーザー");
    lockSheet.hideSheet();
  }
  
  var startTime = new Date().getTime();
  var projectName = ScriptApp.getScriptId();
  var userName = Session.getActiveUser().getEmail();
  
  while (true) {
    var scriptLock = LockService.getScriptLock();
    try {
      scriptLock.waitLock(5000);
      
      var lockStatus = lockSheet.getRange("A2").getValue();
      var lockTime = lockSheet.getRange("B2").getValue();
      
      if (lockStatus !== "LOCKED" || isLockExpired(lockTime, 300)) {
        lockSheet.getRange("A2").setValue("LOCKED");
        lockSheet.getRange("B2").setValue(new Date());
        lockSheet.getRange("C2").setValue(projectName);
        lockSheet.getRange("D2").setValue(userName);
        SpreadsheetApp.flush();
        
        Logger.log("ロック取得成功: " + projectName + " by " + userName);
        return true;
      }
      
    } catch (e) {
      Logger.log("シート操作エラー: " + e);
    } finally {
      scriptLock.releaseLock();
    }
    
    var elapsedTime = (new Date().getTime() - startTime) / 1000;
    if (elapsedTime > timeoutSeconds) {
      Logger.log("ロック取得タイムアウト: " + projectName);
      return false;
    }
    
    Utilities.sleep(500);
  }
}

/**
 * スプレッドシートIDを指定してロックを取得
 * 別プロジェクトから特定のスプレッドシートにアクセスする場合に使用
 * @param {string} spreadsheetId - 対象スプレッドシートのID
 * @param {string} lockSheetName - ロック管理用シート名(デフォルト: "_LOCK_")
 * @param {number} timeoutSeconds - タイムアウト秒数(デフォルト: 30秒)
 * @returns {boolean} - ロック取得成功時true、失敗時false
 */
function acquireSpreadsheetLockById(spreadsheetId, lockSheetName = "_LOCK_", timeoutSeconds = 30) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var lockSheet = ss.getSheetByName(lockSheetName);
  
  // ロック管理シートが存在しない場合は作成
  if (!lockSheet) {
    lockSheet = ss.insertSheet(lockSheetName);
    lockSheet.getRange("A1").setValue("ロック状態");
    lockSheet.getRange("B1").setValue("取得時刻");
    lockSheet.getRange("C1").setValue("プロジェクト名");
    lockSheet.getRange("D1").setValue("実行ユーザー");
    lockSheet.hideSheet();
  }
  
  var startTime = new Date().getTime();
  var projectName = ScriptApp.getScriptId();
  var userName = Session.getActiveUser().getEmail();
  
  while (true) {
    var scriptLock = LockService.getScriptLock();
    try {
      scriptLock.waitLock(5000);
      
      var lockStatus = lockSheet.getRange("A2").getValue();
      var lockTime = lockSheet.getRange("B2").getValue();
      
      if (lockStatus !== "LOCKED" || isLockExpired(lockTime, 300)) {
        lockSheet.getRange("A2").setValue("LOCKED");
        lockSheet.getRange("B2").setValue(new Date());
        lockSheet.getRange("C2").setValue(projectName);
        lockSheet.getRange("D2").setValue(userName);
        SpreadsheetApp.flush();
        
        Logger.log("ロック取得成功: " + projectName + " by " + userName);
        return true;
      }
      
    } catch (e) {
      Logger.log("シート操作エラー: " + e);
    } finally {
      scriptLock.releaseLock();
    }
    
    var elapsedTime = (new Date().getTime() - startTime) / 1000;
    if (elapsedTime > timeoutSeconds) {
      Logger.log("ロック取得タイムアウト: " + projectName);
      return false;
    }
    
    Utilities.sleep(500);
  }
}

/**
 * スプレッドシート上のロックを解放する
 * コンテナバインドスクリプト用
 * @param {string} lockSheetName - ロック管理用シート名(デフォルト: "_LOCK_")
 */
function releaseSpreadsheetLock(lockSheetName = "_LOCK_") {
  var scriptLock = LockService.getScriptLock();
  try {
    scriptLock.waitLock(5000);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var lockSheet = ss.getSheetByName(lockSheetName);
    
    if (lockSheet) {
      lockSheet.getRange("A2").setValue("UNLOCKED");
      lockSheet.getRange("B2").setValue("");
      lockSheet.getRange("C2").setValue("");
      lockSheet.getRange("D2").setValue("");
      SpreadsheetApp.flush();
      
      Logger.log("ロック解放成功");
    }
  } catch (e) {
    Logger.log("ロック解放エラー: " + e);
  } finally {
    scriptLock.releaseLock();
  }
}

/**
 * スプレッドシートIDを指定してロックを解放
 * 別プロジェクトから特定のスプレッドシートにアクセスする場合に使用
 * @param {string} spreadsheetId - 対象スプレッドシートのID
 * @param {string} lockSheetName - ロック管理用シート名(デフォルト: "_LOCK_")
 */
function releaseSpreadsheetLockById(spreadsheetId, lockSheetName = "_LOCK_") {
  var scriptLock = LockService.getScriptLock();
  try {
    scriptLock.waitLock(5000);
    
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var lockSheet = ss.getSheetByName(lockSheetName);
    
    if (lockSheet) {
      lockSheet.getRange("A2").setValue("UNLOCKED");
      lockSheet.getRange("B2").setValue("");
      lockSheet.getRange("C2").setValue("");
      lockSheet.getRange("D2").setValue("");
      SpreadsheetApp.flush();
      
      Logger.log("ロック解放成功");
    }
  } catch (e) {
    Logger.log("ロック解放エラー: " + e);
  } finally {
    scriptLock.releaseLock();
  }
}

/**
 * ロックがタイムアウトしているか確認
 * @param {Date} lockTime - ロック取得時刻
 * @param {number} timeoutSeconds - タイムアウト秒数(デフォルト: 300秒=5分)
 * @returns {boolean} - タイムアウト時true
 */
function isLockExpired(lockTime, timeoutSeconds = 300) {
  if (!lockTime) return true;
  
  var now = new Date().getTime();
  var lockTimestamp = new Date(lockTime).getTime();
  var elapsedSeconds = (now - lockTimestamp) / 1000;
  
  return elapsedSeconds > timeoutSeconds;
}

/**
 * ロック機構を使った安全な書き込み
 * acquireとreleaseを必ずペアで実行する
 * @param {string} spreadsheetId - 対象スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {string} range - 書き込み範囲(A1表記)
 * @param {Array} data - 書き込むデータの2次元配列
 * @returns {boolean} - 成功時true
 */
function safeWriteToSheet(spreadsheetId, sheetName, range, data) {
  // スプレッドシートレベルのロック取得
  if (!acquireSpreadsheetLockById(spreadsheetId)) {
    Logger.log("✗ ロック取得失敗 - 他の処理が実行中です");
    return false;
  }
  
  try {
    // 実際の書き込み処理
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(sheetName);
    sheet.getRange(range).setValues(data);
    SpreadsheetApp.flush();
    
    Logger.log("✓ 書き込み成功");
    return true;
    
  } catch (e) {
    Logger.log("✗ 書き込みエラー: " + e);
    return false;
    
  } finally {
    // 必ずスプレッドシートレベルのロックを解放
    releaseSpreadsheetLockById(spreadsheetId);
  }
}

/**
 * 緊急時: ロックが残ってしまった場合の強制解除
 * エラーで処理が中断し、_LOCK_シートが"LOCKED"のまま残った時に使用
 */
function forceUnlock() {
  Logger.log("⚠️ 強制的にロックを解除します");
  releaseSpreadsheetLockById(TARGET_SPREADSHEET_ID);
  Logger.log("✓ ロック解放完了");
}

/**
 * ロック機構の完全テスト
 * acquireとreleaseが正しく動作することを確認
 */
function testCompleteLockMechanism() {
  const ssId = TARGET_SPREADSHEET_ID;
  
  Logger.log("=== テスト1: ロック取得と解放 ===");
  
  // ロック取得
  if (acquireSpreadsheetLockById(ssId)) {
    Logger.log("✓ ロック取得成功");
    
    // _LOCK_シートの状態を確認
    var ss = SpreadsheetApp.openById(ssId);
    var lockSheet = ss.getSheetByName("_LOCK_");
    Logger.log("現在の状態: " + lockSheet.getRange("A2").getValue());
    Logger.log("取得時刻: " + lockSheet.getRange("B2").getValue());
    
    // 3秒待機
    Utilities.sleep(3000);
    
    // ロック解放
    releaseSpreadsheetLockById(ssId);
    Logger.log("✓ ロック解放実行");
    
    // 解放後の状態を確認
    Logger.log("解放後の状態: " + lockSheet.getRange("A2").getValue());
    
  } else {
    Logger.log("✗ ロック取得失敗");
  }
  
  Logger.log("");
  Logger.log("=== テスト2: 2重ロックの防止 ===");
  
  // 1回目のロック取得
  if (acquireSpreadsheetLockById(ssId, "_LOCK_", 5)) {
    Logger.log("✓ 1回目のロック取得成功");
    
    // ロックを解放せずに2回目を試す(別のトリガーや関数からの実行をシミュレート)
    Logger.log("解放せずに2回目を試みます...");
    
    if (acquireSpreadsheetLockById(ssId, "_LOCK_", 5)) {
      Logger.log("✗ 2回目も取得できてしまった(エラー!)");
    } else {
      Logger.log("✓ 2回目は正しく拒否されました");
    }
    
    // 1回目のロックを解放
    releaseSpreadsheetLockById(ssId);
    Logger.log("✓ 1回目のロック解放");
    
  }
  
  Logger.log("");
  Logger.log("=== テスト3: 実際の書き込み ===");
  
  var testData = [
    ["テスト1", "データ1", new Date()],
    ["テスト2", "データ2", new Date()]
  ];
  
  if (safeWriteToSheet(ssId, "test", "A1:C2", testData)) {
    Logger.log("✓ 安全な書き込み成功");
  } else {
    Logger.log("✗ 書き込み失敗");
  }
  
  Logger.log("=== 全テスト完了 ===");
}