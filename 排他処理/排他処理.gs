/**
 * @fileoverview Google Apps Script (GAS) を用いたスプレッドシートの排他処理ライブラリ
 *
 * @description
 * 複数のスクリプトやユーザーが同時にスプレッドシートを編集しようとすると発生する
 * 競合状態（レースコンディション）を防ぐための排他制御（ロック）機構を提供します。
 * スプレッドシートに特定の管理シート（デフォルト: `_LOCK_`）を作成し、
 * そのシートの状態を読み書きすることでロックを実現します。
 * GASの `LockService` と組み合わせることで、管理シートへのアクセス自体も保護し、信頼性を高めています。
 * 複数のGASプロジェクトから同一スプレッドシートへの同時書き込みによる
 * 競合状態を防ぐための排他制御ライブラリです。
 * 
 * 改善点:
 * - ユニークID(UUID)による所有者識別
 * - ダブルチェック方式による競合検出
 * - 安全なロック解放(自分のIDの場合のみ解放)
 * - キャッシュ対策の強化 
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
 * @version 2.0.0
 * @license MIT
 */

const TARGET_SPREADSHEET_ID = "貴方のスプレッドシートIDをここに入れてください";

/**
 * ユニークなロックIDを生成
 * プロジェクトID + タイムスタンプ + ランダム値で一意性を保証
 * @returns {string} - ユニークなロックID
 */
function generateLockId() {
  var projectId = ScriptApp.getScriptId();
  var timestamp = new Date().getTime();
  var random = Math.random().toString(36).substring(2, 15);
  return projectId + "-" + timestamp + "-" + random;
}

/**
 * スプレッドシート上でセマフォ(ロック)を取得する(改善版)
 * コンテナバインドスクリプト用
 * @param {string} lockSheetName - ロック管理用シート名(デフォルト: "_LOCK_")
 * @param {number} timeoutSeconds - タイムアウト秒数(デフォルト: 30秒)
 * @returns {string|null} - ロック取得成功時はロックID、失敗時はnull
 */
function acquireSpreadsheetLock(lockSheetName = "_LOCK_", timeoutSeconds = 30) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lockSheet = ss.getSheetByName(lockSheetName);
  
  // ロック管理シートが存在しない場合は作成
  if (!lockSheet) {
    lockSheet = ss.insertSheet(lockSheetName);
    lockSheet.getRange("A1").setValue("ロック状態");
    lockSheet.getRange("B1").setValue("ロックID");
    lockSheet.getRange("C1").setValue("取得時刻");
    lockSheet.getRange("D1").setValue("プロジェクト名");
    lockSheet.getRange("E1").setValue("実行ユーザー");
    lockSheet.hideSheet();
  }
  
  var startTime = new Date().getTime();
  var projectName = ScriptApp.getScriptId();
  var userName = Session.getActiveUser().getEmail();
  var myLockId = generateLockId();
  
  Logger.log("ロック取得開始: " + myLockId);
  
  while (true) {
    var scriptLock = LockService.getScriptLock();
    var hasScriptLock = false;
    
    try {
      hasScriptLock = scriptLock.tryLock(100);
      
      if (!hasScriptLock) {
        Logger.log("【DEBUG】リトライ" + retryCount + ": ScriptLock取得待機中...");
        Utilities.sleep(100);
        continue;
      }
      
      Logger.log("【DEBUG】リトライ" + retryCount + ": ScriptLock取得成功");
      
      // ★★★ 重要: ループ内でスプレッドシートを再取得してキャッシュをクリア ★★★: 他プロジェクトの書き込みを確実に読む
      SpreadsheetApp.flush();  // 保留中の変更を強制的に反映
      var freshSs = SpreadsheetApp.getActiveSpreadsheet();
      var freshLockSheet = freshSs.getSheetByName(lockSheetName);
      
      // 最新の値を取得
      var lockStatus = freshLockSheet.getRange("A2").getValue();
      var existingLockId = freshLockSheet.getRange("B2").getValue();
      var lockTime = freshLockSheet.getRange("C2").getValue();
      
      Logger.log("【DEBUG】_LOCKシート状態: " + lockStatus + " (時刻: " + lockTime + ")");
  
      // ロックが解放されているか、タイムアウトしているかチェック
      if (lockStatus !== "LOCKED" || isLockExpired(lockTime, 300)) {

        // ステップ1: 自分のIDを書き込む
        freshLockSheet.getRange("A2").setValue("LOCKED");
        freshLockSheet.getRange("B2").setValue(myLockId);
        freshLockSheet.getRange("C2").setValue(new Date());
        freshLockSheet.getRange("D2").setValue(projectName);
        freshLockSheet.getRange("E2").setValue(userName);
        SpreadsheetApp.flush();
        
        // ステップ2: 少し待機(他のプロジェクトの書き込みとの競合を検出)
        Utilities.sleep(200);
        
        // ステップ3: ダブルチェック - 本当に自分のIDが書かれているか確認
        SpreadsheetApp.flush();
        var verifyLockId = freshLockSheet.getRange("B2").getValue();
        
        if (verifyLockId === myLockId) {
          // 成功: 自分のIDが確実に書かれている
          Logger.log("ロック取得成功: " + myLockId);
          return myLockId;
        } else {
          // 失敗: 他のプロジェクトに上書きされた
          Logger.log("競合検出: 他のプロジェクトに上書きされました (期待: " + myLockId + ", 実際: " + verifyLockId + ")");
          // リトライへ
        }
      }
      
    } catch (e) {
      Logger.log("【DEBUG】シート操作エラー: " + e);
    } finally {
      if (hasScriptLock) {
        scriptLock.releaseLock();
      }
    }
    
    var elapsedTime = (new Date().getTime() - startTime) / 1000;
    if (elapsedTime > timeoutSeconds) {
      Logger.log("ロック取得タイムアウト: " + myLockId);
      return null;
    }
    
    Utilities.sleep(500);
  }
}

/**
 * スプレッドシートIDを指定してロックを取得(改善版)
 * 別プロジェクトから特定のスプレッドシートにアクセスする場合に使用
 * @param {string} spreadsheetId - 対象スプレッドシートのID
 * @param {string} lockSheetName - ロック管理用シート名(デフォルト: "_LOCK_")
 * @param {number} timeoutSeconds - タイムアウト秒数(デフォルト: 30秒)
 * @returns {string|null} - ロック取得成功時はロックID、失敗時はnull
 */
function acquireSpreadsheetLockById(spreadsheetId, lockSheetName = "_LOCK_", timeoutSeconds = 30) {
  var startTime = new Date().getTime();
  var projectName = ScriptApp.getScriptId();
  var userName = Session.getActiveUser().getEmail();
  var myLockId = generateLockId();
  
  Logger.log("ロック取得開始: " + myLockId);
  
  // 初回のシート確認と作成
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var lockSheet = ss.getSheetByName(lockSheetName);
  
  if (!lockSheet) {
    lockSheet = ss.insertSheet(lockSheetName);
    lockSheet.getRange("A1").setValue("ロック状態");
    lockSheet.getRange("B1").setValue("ロックID");
    lockSheet.getRange("C1").setValue("取得時刻");
    lockSheet.getRange("D1").setValue("プロジェクト名");
    lockSheet.getRange("E1").setValue("実行ユーザー");
    lockSheet.hideSheet();
    SpreadsheetApp.flush();
  }
  
  while (true) {
    var scriptLock = LockService.getScriptLock();
    var hasScriptLock = false;
    
    try {
      hasScriptLock = scriptLock.tryLock(100);
      
      if (!hasScriptLock) {
        Logger.log("【DEBUG】リトライ" + retryCount + ": ScriptLock取得待機中...");
        Utilities.sleep(100);
        continue;
      }
      
      Logger.log("【DEBUG】リトライ" + retryCount + ": ScriptLock取得成功");
      
      // ★★★ 重要: ループ内でスプレッドシートを再取得 ★★★
      SpreadsheetApp.flush();
      var freshSs = SpreadsheetApp.openById(spreadsheetId);
      var freshLockSheet = freshSs.getSheetByName(lockSheetName);
      
      var lockStatus = freshLockSheet.getRange("A2").getValue();
      var existingLockId = freshLockSheet.getRange("B2").getValue();
      var lockTime = freshLockSheet.getRange("C2").getValue();
      
      Logger.log("【DEBUG】_LOCKシート状態: " + lockStatus + " (時刻: " + lockTime + ")");
      
      if (lockStatus !== "LOCKED" || isLockExpired(lockTime, 300)) {
        
        // ステップ1: 書き込み
        freshLockSheet.getRange("A2").setValue("LOCKED");
        freshLockSheet.getRange("B2").setValue(myLockId);
        freshLockSheet.getRange("C2").setValue(new Date());
        freshLockSheet.getRange("D2").setValue(projectName);
        freshLockSheet.getRange("E2").setValue(userName);
        SpreadsheetApp.flush();
        
        Logger.log("【DEBUG】ロック取得成功: " + projectName + " (リトライ回数: " + retryCount + ")");
        
        // ステップ2: 待機
        Utilities.sleep(200);
        
        // ステップ3: ダブルチェック
        SpreadsheetApp.flush();
        var verifyLockId = freshLockSheet.getRange("B2").getValue();
        
        if (verifyLockId === myLockId) {
          Logger.log("ロック取得成功: " + myLockId);
          return myLockId;
        } else {
          Logger.log("競合検出: 上書きされました (期待: " + myLockId + ", 実際: " + verifyLockId + ")");
        }
      }
      
    } catch (e) {
      Logger.log("【DEBUG】シート操作エラー: " + e);
    } finally {
      if (hasScriptLock) {
        scriptLock.releaseLock();
      }
    }
    
    var elapsedTime = (new Date().getTime() - startTime) / 1000;
    if (elapsedTime > timeoutSeconds) {
      Logger.log("ロック取得タイムアウト: " + myLockId);
      return null;
    }
    
    Utilities.sleep(500);
  }
}

/**
 * スプレッドシート上のロックを解放する(改善版)
 * 自分のロックIDの場合のみ解放
 * @param {string} lockId - ロック取得時に返されたロックID
 * @param {string} lockSheetName - ロック管理用シート名(デフォルト: "_LOCK_")
 * @returns {boolean} - 解放成功時true、失敗時false
 */
function releaseSpreadsheetLock(lockId, lockSheetName = "_LOCK_") {
  if (!lockId) {
    Logger.log("ロック解放エラー: ロックIDが指定されていません");
    return false;
  }
  
  var scriptLock = LockService.getScriptLock();
  try {
    scriptLock.waitLock(5000);
    
    SpreadsheetApp.flush();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var lockSheet = ss.getSheetByName(lockSheetName);
    
    if (lockSheet) {
      var currentLockId = lockSheet.getRange("B2").getValue();
      
      // 安全性チェック: 自分のロックIDの場合のみ解放
      if (currentLockId === lockId) {
        lockSheet.getRange("A2").setValue("UNLOCKED");
        lockSheet.getRange("B2").setValue("");
        lockSheet.getRange("C2").setValue("");
        lockSheet.getRange("D2").setValue("");
        lockSheet.getRange("E2").setValue("");
        SpreadsheetApp.flush();
        
        Logger.log("ロック解放成功: " + lockId);
        return true;
      } else {
        Logger.log("ロック解放スキップ: 現在のロックIDが異なります (自分: " + lockId + ", 現在: " + currentLockId + ")");
        return false;
      }
    }
  } catch (e) {
    Logger.log("ロック解放エラー: " + e);
    return false;
  } finally {
    scriptLock.releaseLock();
  }
}

/**
 * スプレッドシートIDを指定してロックを解放(改善版)
 * @param {string} spreadsheetId - 対象スプレッドシートのID
 * @param {string} lockId - ロック取得時に返されたロックID
 * @param {string} lockSheetName - ロック管理用シート名(デフォルト: "_LOCK_")
 * @returns {boolean} - 解放成功時true、失敗時false
 */
function releaseSpreadsheetLockById(spreadsheetId, lockId, lockSheetName = "_LOCK_") {
  if (!lockId) {
    Logger.log("ロック解放エラー: ロックIDが指定されていません");
    return false;
  }
  
  var scriptLock = LockService.getScriptLock();
  try {
    scriptLock.waitLock(5000);
    
    SpreadsheetApp.flush();
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var lockSheet = ss.getSheetByName(lockSheetName);
    
    if (lockSheet) {
      var currentLockId = lockSheet.getRange("B2").getValue();
      
      if (currentLockId === lockId) {
        lockSheet.getRange("A2").setValue("UNLOCKED");
        lockSheet.getRange("B2").setValue("");
        lockSheet.getRange("C2").setValue("");
        lockSheet.getRange("D2").setValue("");
        lockSheet.getRange("E2").setValue("");
        SpreadsheetApp.flush();
        
        Logger.log("ロック解放成功: " + lockId);
        return true;
      } else {
        Logger.log("ロック解放スキップ: 現在のロックIDが異なります (自分: " + lockId + ", 現在: " + currentLockId + ")");
        return false;
      }
    }
  } catch (e) {
    Logger.log("ロック解放エラー: " + e);
    return false;
  } finally {
    scriptLock.releaseLock();
  }
  
  return false;
}


/**
 * ロックがタイムアウトしているか確認(デバッグ版)
 * @param {Date} lockTime - ロック取得時刻
 * @param {number} timeoutSeconds - タイムアウト秒数(デフォルト: 300秒=5分)
 * @returns {boolean} - タイムアウト時true
 */
function isLockExpired(lockTime, timeoutSeconds = 300) {
  if (!lockTime) {
    Logger.log("【DEBUG】isLockExpired: lockTimeが空 → true");
    return true;
  }
  
  var now = new Date().getTime();
  var lockTimestamp = new Date(lockTime).getTime();
  var elapsedSeconds = (now - lockTimestamp) / 1000;
  
  Logger.log("【DEBUG】isLockExpired: 経過時間=" + elapsedSeconds.toFixed(2) + "秒, タイムアウト=" + timeoutSeconds + "秒");
  
  return elapsedSeconds > timeoutSeconds;
}

/**
 * ロック機構を使った安全な書き込み(改善版)
 * @param {string} spreadsheetId - 対象スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {string} range - 書き込み範囲(A1表記)
 * @param {Array} data - 書き込むデータの2次元配列
 * @returns {boolean} - 成功時true
 */
function safeWriteToSheet(spreadsheetId, sheetName, range, data) {
  var lockId = acquireSpreadsheetLockById(spreadsheetId);
  
  if (!lockId) {
    Logger.log("ロック取得失敗 - 他の処理が実行中です");
    return false;
  }
  
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(sheetName);
    sheet.getRange(range).setValues(data);
    SpreadsheetApp.flush();
    
    Logger.log("書き込み成功");
    return true;
    
  } catch (e) {
    Logger.log("書き込みエラー: " + e);
    return false;
    
  } finally {
    releaseSpreadsheetLockById(spreadsheetId, lockId);
  }
}

/**
 * 緊急時: ロックが残ってしまった場合の強制解除(改善版)
 * タイムアウトしたロックのみ強制解除
 */
function forceUnlockExpired() {
  var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  var lockSheet = ss.getSheetByName("_LOCK_");
  
  if (!lockSheet) {
    Logger.log("_LOCKシートが存在しません");
    return;
  }
  
  var lockStatus = lockSheet.getRange("A2").getValue();
  var lockTime = lockSheet.getRange("C2").getValue();
  var lockId = lockSheet.getRange("B2").getValue();
  
  if (lockStatus === "LOCKED" && isLockExpired(lockTime, 300)) {
    Logger.log("⚠️ タイムアウトしたロックを強制解除します: " + lockId);
    
    lockSheet.getRange("A2").setValue("UNLOCKED");
    lockSheet.getRange("B2").setValue("");
    lockSheet.getRange("C2").setValue("");
    lockSheet.getRange("D2").setValue("");
    lockSheet.getRange("E2").setValue("");
    SpreadsheetApp.flush();
    
    Logger.log("✓ ロック強制解放完了");
  } else {
    Logger.log("有効なロックが存在するため、強制解除は行いませんでした");
  }
}

/**
 * テスト関数: プロジェクトA用(改善版)
 */
function testFromProjectA_Improved() {
  Logger.log("=== プロジェクトA: 書き込み開始(改善版) ===");
  
  var lockId = acquireSpreadsheetLockById(TARGET_SPREADSHEET_ID);
  
  if (lockId) {
    Logger.log("プロジェクトA: ロック取得成功 (ID: " + lockId + ")");
    
    try {
      var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
      var sheet = ss.getSheetByName("test");
      
      for (var i = 1; i <= 5; i++) {
        var timestamp = new Date().toLocaleTimeString();
        sheet.getRange("A" + i).setValue("ProjectA-" + timestamp);
        Logger.log("プロジェクトA: " + i + "行目書き込み");
        Utilities.sleep(1000);
      }
      
    } finally {
      releaseSpreadsheetLockById(TARGET_SPREADSHEET_ID, lockId);
      Logger.log("プロジェクトA: ロック解放");
    }
  } else {
    Logger.log("プロジェクトA: ロック取得失敗");
  }
}

/**
 * テスト関数: プロジェクトB用(改善版)
 */
function testFromProjectB_Improved() {
  Logger.log("=== プロジェクトB: 書き込み開始(改善版) ===");
  
  var lockId = acquireSpreadsheetLockById(TARGET_SPREADSHEET_ID, "_LOCK_", 60);
  
  if (lockId) {
    Logger.log("プロジェクトB: ロック取得成功 (ID: " + lockId + ")");
    
    try {
      var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
      var sheet = ss.getSheetByName("test");
      
      for (var i = 1; i <= 5; i++) {
        var timestamp = new Date().toLocaleTimeString();
        sheet.getRange("B" + i).setValue("ProjectB-" + timestamp);
        Logger.log("プロジェクトB: " + i + "行目書き込み");
        Utilities.sleep(1000);
      }
      
    } finally {
      releaseSpreadsheetLockById(TARGET_SPREADSHEET_ID, lockId);
      Logger.log("プロジェクトB: ロック解放");
    }
  } else {
    Logger.log("プロジェクトB: ロック取得失敗");
  }
}