/**
 * ========================================
 * 🧪 現場PoC専用：7月限定 自動監視デーモン（日付バグ修正版）
 * ========================================
 */
function daemon_PoC_JulyOnly() {
  // ----------------------------------------
  // ▼ PoC設定
  // ----------------------------------------
  var TARGET_YEAR    = 2026;
  var TARGET_MONTH   = 7;
  var TARGET_LEADERS = ['加藤', '望月', '関田']; 
  
  var TEST_ROOM_ID   = ''; // 空欄で本番ルームへ
  // ----------------------------------------

  var today = new Date();
  Logger.log('====== 🚀 PoCデーモン (7月限定) 起動 ======');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getDocumentProperties();
  var needIndexUpdate = false;
  
  var context;
  try { 
    context = buildContext(TARGET_YEAR, TARGET_MONTH); 
  } catch (e) { 
    Logger.log('❌ データ読み込み失敗: ' + e.message);
    return; 
  }
  
  context.clinicMaster.list.forEach(function(clinic) {
    var leaderText = (clinic.leaderId || '') + (clinic.director || '');
    var isTargetLeader = false;
    for (var i = 0; i < TARGET_LEADERS.length; i++) {
      if (leaderText.indexOf(TARGET_LEADERS[i]) !== -1) {
        isTargetLeader = true;
        break;
      }
    }
    if (!isTargetLeader) return; 
    
    var targetVal = TARGET_YEAR * 12 + TARGET_MONTH;
    if (clinic.openDate && Object.prototype.toString.call(clinic.openDate) === '[object Date]') {
      var openVal = clinic.openDate.getFullYear() * 12 + (clinic.openDate.getMonth() + 1);
      if (targetVal < openVal) return; 
    }
    
    var clinicNo = clinic.clinicNo;
    var lastDay = new Date(TARGET_YEAR, TARGET_MONTH, 0).getDate();
    var startKey = dateKey(new Date(TARGET_YEAR, TARGET_MONTH - 1, 1));
    var endKey   = dateKey(new Date(TARGET_YEAR, TARGET_MONTH - 1, lastDay));
    
    var currentSchedule = collectScheduleFromContext(context, clinicNo, startKey, endKey);
    var scheduleStr = JSON.stringify(currentSchedule);
    var currentHash = computeMD5_(scheduleStr);
    
    var propKey = 'POC_HASH_' + TARGET_YEAR + '_' + TARGET_MONTH + '_' + clinicNo;
    var lastHash = props.getProperty(propKey);
    
    var expectedSheetName = ('0' + TARGET_MONTH).slice(-2) + clinic.name;
    var targetSheet = ss.getSheetByName(expectedSheetName);
    var nowStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    
    var oldSchedule = getSavedScheduleState_(clinicNo, TARGET_YEAR, TARGET_MONTH);
    var isChanged = (currentHash !== lastHash || !targetSheet);
    
    if (isChanged) {
      var isFirstTime = (!lastHash); 
      var diffs = isFirstTime ? { filled: [], vacated: [] } : checkScheduleDifferences_(oldSchedule, currentSchedule);
      
      var shouldSendChat = true;
      if (!isFirstTime && diffs.filled.length === 0 && diffs.vacated.length === 0) {
        shouldSendChat = false; 
      }

      if (!shouldSendChat) {
        Logger.log('⏩ ' + clinic.name + ' : 細かな変更を検知しましたが、医師の追加・欠員ではないため、通知をスキップします。');
      } else {
        Logger.log('🔥 【処理開始】' + clinic.name + ' (' + TARGET_MONTH + '月分) - チャット送信あり');
      }
      
      var maxRetries = 2;
      for (var attempt = 1; attempt <= maxRetries; attempt++) {
        try {
          var sheet = generateScheduleWithContext(context, clinicNo, TARGET_YEAR, TARGET_MONTH);
          SpreadsheetApp.flush();
          Utilities.sleep(2000); 
          
          sheet.getRange('A1').clearNote();
          SpreadsheetApp.flush();
          
          var pdfFile = exportSheetToPDF(sheet, TARGET_YEAR, TARGET_MONTH, clinic.name);
          if (!pdfFile) throw new Error('PDF生成失敗');
          
          if (shouldSendChat) {
            var pdfBlob = pdfFile.getBlob();
            var sheetUrl = ss.getUrl() + '#gid=' + sheet.getSheetId();
            var type = isFirstTime ? 'monthly' : 'filled';
            
            var finalRoomId = TEST_ROOM_ID !== '' ? TEST_ROOM_ID : clinic.chatId;
            if (finalRoomId) {
              Logger.log('📥 Chatwork送信中... (宛先: ' + finalRoomId + ')');
              poc_sendChatworkDirectly_(clinic, TARGET_MONTH, type, diffs, pdfBlob, sheetUrl, finalRoomId);
            }
          }
          
          props.setProperty('LAST_UPDATE_' + sheet.getName(), nowStr);
          props.setProperty(propKey, currentHash);
          saveScheduleState_(clinicNo, TARGET_YEAR, TARGET_MONTH, currentSchedule);
          sheet.getRange('A1').setNote('🔄 最新データ更新: ' + nowStr + '\n(ハッシュ: ' + currentHash + ')');
          needIndexUpdate = true;
          
          break; 
        } catch (err) {
          if (attempt < maxRetries) {
            Logger.log('⚠️ 【リトライ】エラー再試行... (' + err.message + ')');
            Utilities.sleep(2000);
          } else {
            Logger.log('❌ 【エラー】' + clinic.name + ': ' + err.message);
          }
        }
      }
      Utilities.sleep(1500); 
    } else {
      Logger.log('✅ ' + clinic.name + ' は変更なし (スキップ)');
    }
  });
  
  if (needIndexUpdate) rebuildIndexSheet_standalone(ss);
  Logger.log('====== 🚀 PoCデーモン 終了 ======');
}

/**
 * 過去と現在のシフトを比較し、「医師が追加された枠」と「不在になった枠」を両方抽出する関数
 * ★日付の読み込みバグ（0月問題）を修正しました
 */
function checkScheduleDifferences_(oldSched, newSched) {
  var diff = { filled: [], vacated: [] };
  if (!oldSched) return diff;
  
  Object.keys(newSched).forEach(function(dKey) {
    // ★修正箇所: ハイフン( - )が含まれていてもいなくても、一律でハイフンを除去して「20260707」の形に統一します
    var cleanKey = dKey.replace(/-/g, ''); 
    
    // 統一された形から正確に月と日を切り出します
    var dateStr = parseInt(cleanKey.substring(4, 6), 10) + "月" + parseInt(cleanKey.substring(6, 8), 10) + "日";
    
    ['am', 'pm'].forEach(function(slot) {
      var oldDocs = (oldSched[dKey] && oldSched[dKey][slot]) ? oldSched[dKey][slot] : [];
      var newDocs = newSched[dKey][slot] || [];

      var oldHasDoc = oldDocs.some(function(d) { return d.name && d.name.replace(/\s/g, '') !== ""; });
      var newHasDoc = newDocs.some(function(d) { return d.name && d.name.replace(/\s/g, '') !== ""; });

      var slotName = (slot === 'am') ? '午前' : '午後(夜間含む)';

      if (!oldHasDoc && newHasDoc) {
        diff.filled.push(dateStr + "の" + slotName);
      } else if (oldHasDoc && !newHasDoc) {
        diff.vacated.push(dateStr + "の" + slotName);
      }
    });
  });
  return diff;
}

/**
 * PoC専用：Chatwork直接送信関数
 */
function poc_sendChatworkDirectly_(clinic, month, type, diffs, pdfBlob, sheetUrl, roomId) {
  var token = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN');
  if (!token) return;
  
  var toText = (clinic.leaderId || '') + '\n' + (clinic.sharedAccount || '') + '\n\n';
  var message = '';
  
  if (type === 'monthly') {
    message = toText + '[info][title]来月のシフト表送付[/title]\nお疲れ様です。来月' + month + '月の【' + clinic.name + '】のシフト表を共有させていただきます。\n医師不在がある箇所は、医師が調整され次第更新版をお送りいたします。\n\n保存先URL：\n' + sheetUrl + '\n[/info]';
  } else {
    var diffMessage = '';
    if (diffs.filled.length > 0) {
      diffMessage += '【医師が確定した枠】\n・' + diffs.filled.join('\n・') + '\n\n';
    }
    if (diffs.vacated.length > 0) {
      diffMessage += '【急遽不在(欠員)となった枠】\n・' + diffs.vacated.join('\n・') + '\n\n';
    }

    message = toText + '[info][title]差替シフト表送付[/title]\nお疲れ様です。\nシフトに変更がございましたので、差し替え版をお送りいたします。\n\n' + diffMessage + '適宜ご確認をお願いいたします。\n\n保存先URL：\n' + sheetUrl + '\n[/info]';
  }

  try {
    var boundary = "----WebKitFormBoundary" + Utilities.getUuid().replace(/-/g, '');
    var payload = [];
    function appendString(str) {
      var bytes = Utilities.newBlob(str).getBytes();
      for (var i = 0; i < bytes.length; i++) { payload.push(bytes[i]); }
    }
    appendString("--" + boundary + "\r\n");
    appendString("Content-Disposition: form-data; name=\"message\"\r\n\r\n");
    appendString(message + "\r\n");
    appendString("--" + boundary + "\r\n");
    appendString("Content-Disposition: form-data; name=\"file\"; filename=\"" + pdfBlob.getName() + "\"\r\n");
    appendString("Content-Type: " + pdfBlob.getContentType() + "\r\n\r\n");
    var fileBytes = pdfBlob.getBytes();
    for (var j = 0; j < fileBytes.length; j++) { payload.push(fileBytes[j]); }
    appendString("\r\n--" + boundary + "--\r\n");

    var url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/files';
    var options = {
      "method": "post",
      "headers": { "X-ChatWorkToken": token, "Content-Type": "multipart/form-data; boundary=" + boundary },
      "payload": payload,
      "muteHttpExceptions": true
    };
    UrlFetchApp.fetch(url, options);
  } catch(e) {
    Logger.log('❌ Chatwork送信エラー: ' + e.message);
  }
}

function _ss_getClinicObjectForPoC(clinic) {
  return {
    name: clinic.name,
    leaderId: clinic.leaderId,
    sharedAccount: clinic.sharedAccount,
    chatId: clinic.chatId
  };
}