/**
 * ========================================
 * Chatwork連携モジュール（夜間キュー保存 ＆ 朝の一斉送信）
 * ========================================
 */

/**
 * 夜間用：すぐに送信せず「送信待合室（キュー）」にメッセージとファイルを予約する
 * ★PoC知見 반영：欠員(vacated)と補充(filled)の両方を処理できる動的メッセージに対応
 */
function enqueueChatworkNotification_(clinic, month, type, diffs, pdfFile, sheetUrl) {
  var roomId = clinic.chatId; 
  if (!roomId) {
    Logger.log('【警告】' + clinic.name + ' の送信先ルームIDがないため、送信予約をスキップします。');
    return;
  }
  
  var toText = (clinic.leaderId || '') + '\n' + (clinic.sharedAccount || '') + '\n\n';
  var message = '';
  
  if (type === 'monthly') {
    message = toText + '[info][title]来月のシフト表送付[/title]\nお疲れ様です。来月' + month + '月の【' + clinic.name + '】のシフト表を共有させていただきます。\n医師不在がある箇所は、医師が調整され次第更新版をお送りいたします。\n\n保存先URL：\n' + sheetUrl + '\n[/info]';
  } else {
    // 差分の中身に応じてメッセージを動的に組み立てる
    var diffMessage = '';
    if (diffs.filled && diffs.filled.length > 0) {
      diffMessage += '【医師が確定した枠】\n・' + diffs.filled.join('\n・') + '\n\n';
    }
    if (diffs.vacated && diffs.vacated.length > 0) {
      diffMessage += '【急遽不在(欠員)となった枠】\n・' + diffs.vacated.join('\n・') + '\n\n';
    }

    message = toText + '[info][title]差替シフト表送付[/title]\nお疲れ様です。\nシフトに変更がございましたので、差し替え版をお送りいたします。\n\n' + diffMessage + '適宜ご確認をお願いいたします。\n\n保存先URL：\n' + sheetUrl + '\n[/info]';
  }

  // 送信待合室（非表示シート）に予約データを書き込む
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qSheet = ss.getSheetByName('⚙️_MessageQueue');
  if (!qSheet) {
    qSheet = ss.insertSheet('⚙️_MessageQueue');
    qSheet.hideSheet();
    qSheet.appendRow(['Timestamp', 'ClinicName', 'RoomId', 'Message', 'FileId']);
  }

  var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  qSheet.appendRow([nowStr, clinic.name, roomId, message, pdfFile.getId()]);
  Logger.log('【送信予約 📥】' + clinic.name + ' のメッセージを朝9時配信用に予約しました。');
}

/**
 * 朝9時用：待合室（キュー）に溜まったメッセージを一斉送信する
 */
function flushChatworkQueue_Morning() {
  // =========================================================
  // ★ 本番稼働ON ★
  // =========================================================
  var ENABLE_CHATWORK = true;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qSheet = ss.getSheetByName('⚙️_MessageQueue');
  if (!qSheet) return;

  var data = qSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log('【朝の定期配信】送信待ちの予約メッセージはありませんでした。');
    return;
  }

  var scriptProps = PropertiesService.getScriptProperties();
  var token = scriptProps.getProperty('CHATWORK_API_TOKEN');
  if (!token) {
    Logger.log('【エラー】CHATWORK_API_TOKENが設定されていません。');
    return;
  }

  // 予約されたメッセージを上から順に処理
  for (var i = 1; i < data.length; i++) {
    var clinicName = data[i][1];
    var roomId = data[i][2];
    var message = data[i][3];
    var fileId = data[i][4];

    try {
      // ドライブからPDFを取得（日本語ファイル名もここで復元されます）
      var file = DriveApp.getFileById(fileId);
      var pdfBlob = file.getBlob();

      var boundary = "----WebKitFormBoundary" + Utilities.getUuid().replace(/-/g, '');
      var payload = [];
      function appendString(str) {
        var bytes = Utilities.newBlob(str).getBytes();
        for (var j = 0; j < bytes.length; j++) { payload.push(bytes[j]); }
      }

      appendString("--" + boundary + "\r\n");
      appendString("Content-Disposition: form-data; name=\"message\"\r\n\r\n");
      appendString(message + "\r\n");

      appendString("--" + boundary + "\r\n");
      appendString("Content-Disposition: form-data; name=\"file\"; filename=\"" + pdfBlob.getName() + "\"\r\n");
      appendString("Content-Type: " + pdfBlob.getContentType() + "\r\n\r\n");

      var fileBytes = pdfBlob.getBytes();
      for (var k = 0; k < fileBytes.length; k++) { payload.push(fileBytes[k]); }
      appendString("\r\n--" + boundary + "--\r\n");

      var url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/files';
      var options = {
        "method": "post",
        "headers": { "X-ChatWorkToken": token, "Content-Type": "multipart/form-data; boundary=" + boundary },
        "payload": payload,
        "muteHttpExceptions": true
      };

      if (ENABLE_CHATWORK) {
        var response = UrlFetchApp.fetch(url, options);
        var responseCode = response.getResponseCode();
        if (responseCode >= 200 && responseCode < 300) {
          Logger.log('【送信成功 📤】' + clinicName + ' (Room: ' + roomId + ')');
        } else {
          Logger.log('【送信エラー ❌】' + clinicName + ' (HTTP: ' + responseCode + '): ' + response.getContentText());
        }
      } else {
        Logger.log('【送信ミュート中🔇】' + clinicName + ' 宛のチャット送信を消化（保留）しました。');
      }
    } catch(e) {
      Logger.log('【送信エラー ❌】' + clinicName + ': ' + e.message);
    }
    Utilities.sleep(1500); // 連続送信によるAPI制限を回避
  }

  // 全件の送信（または消化）が終わったら、待合室を空っぽにする
  qSheet.clear();
  qSheet.appendRow(['Timestamp', 'ClinicName', 'RoomId', 'Message', 'FileId']);
  Logger.log('【朝の定期配信完了】送信待合室（キュー）をクリアしました。');
}