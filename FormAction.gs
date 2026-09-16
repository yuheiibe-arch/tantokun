/**
 * ========================================
 * 所属先変更フォーム 連携モジュール（「別途記載」対応版）
 * ========================================
 */

/**
 * 1. フォーム送信時トリガー（チェックボックス準備・マスタ更新・シフト再生成・Chatwork送信）
 * ※インストーラブルトリガー（フォーム送信時）で設定
 */
function onFormSubmit_ShiftChange(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== '変更依頼') return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headerMap = buildHeaderMap(headers);

  // 列の特定
  var idxId = headerMap['医籍登録番号（数字のみ）'];
  var idxName = headerMap['医師名'];
  var idxUrl = headerMap['該当行'];
  var idxDone = headerMap['完了']; 
  var idxBaseClinic = headerMap['１）申告拠点を選択してください。']; 
  var idxDesiredName = headerMap['希望の記載名']; // ★追加：希望の記載名列の特定
  
  var idxNewClinic = -1;
  for (var key in headerMap) {
    if (key.indexOf('変更の所属先を記入') !== -1) {
      idxNewClinic = headerMap[key]; break;
    }
  }

  // 送信された行の取得
  var row = e.range.getRow();
  var dataRow = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  var targetId = String(dataRow[idxId]).trim();
  var targetName = String(dataRow[idxName]).trim();
  var baseClinic = (idxBaseClinic !== undefined) ? String(dataRow[idxBaseClinic]).trim() : '';
  var newClinic = String(dataRow[idxNewClinic]).trim();
  var desiredName = (idxDesiredName !== undefined) ? String(dataRow[idxDesiredName]).trim() : '';

  // ==========================================
  // ★追加：「別途記載」が選ばれた場合は「希望の記載名」のテキストを優先する
  // ==========================================
  if (newClinic === '別途記載' && desiredName !== '') {
    newClinic = desiredName;
  }

  // ==========================================
  // 新規行の「完了」列にチェックボックスを用意する
  // ==========================================
  if (idxDone !== undefined) {
    sheet.getRange(row, idxDone + 1).insertCheckboxes();
  }

  if (!targetId) return;

  // ① 医師マスタの更新（更新前の所属先を取得するため、更新処理内で現在の所属を取得）
  var oldClinic = getDoctorCurrentWork_(targetId);

  var masterUrl = updateDoctorMaster_(targetId, newClinic);
  if (masterUrl && idxUrl !== undefined) {
    // 該当行URLを書き込み
    sheet.getRange(row, idxUrl + 1).setFormula('=HYPERLINK("' + masterUrl + '", "該当行リンク")');
  }

  // ★指定のChatworkルーム（444531491）へ通知を送信
  sendAdminMasterChatworkAlert_(baseClinic, targetName, oldClinic, newClinic);

  // ② 影響範囲（該当医師が勤務している拠点と日付）の洗い出し
  var today = new Date();
  var currentYear = today.getFullYear();
  var currentMonth = today.getMonth() + 1;
  
  var context = buildContext(currentYear, currentMonth);
  var affectedClinicsMap = {}; // clinicNo -> 勤務日の配列

  // 当月（例: "2026-08"）の文字列を作って絞り込む
  var targetMonthPrefix = currentYear + '-' + ('0' + currentMonth).slice(-2);

  var shiftRows = context.shiftRows;
  for (var i = 0; i < shiftRows.length; i++) {
    if (shiftRows[i].ikiNo === targetId && shiftRows[i].dateKey.indexOf(targetMonthPrefix) === 0) {
      var cNo = shiftRows[i].clinicNo;
      var dKey = shiftRows[i].dateKey;
      if (!affectedClinicsMap[cNo]) affectedClinicsMap[cNo] = [];
      if (affectedClinicsMap[cNo].indexOf(dKey) === -1) {
        affectedClinicsMap[cNo].push(dKey);
      }
    }
  }

  // ③ シフト再生成とChatwork直接送信（該当拠点へ）
  var affectedClinicNos = Object.keys(affectedClinicsMap);
  affectedClinicNos.forEach(function(cNoStr) {
    var clinicNo = parseInt(cNoStr, 10);
    var clinic = context.clinicMaster.byClinicNo[clinicNo];
    if (!clinic) return;

    // シフト表の再生成
    var newSheet = generateScheduleWithContext(context, clinicNo, currentYear, currentMonth);
    SpreadsheetApp.flush();

    // PDF化
    var pdfFile = exportSheetToPDF(newSheet, currentYear, currentMonth, clinic.name);
    var sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl() + '#gid=' + newSheet.getSheetId();

    // 日付のフォーマット整形
    var dateStrs = affectedClinicsMap[cNoStr].map(function(dk) {
      var parts = dk.split('-');
      return parseInt(parts[1], 10) + '月' + parseInt(parts[2], 10) + '日';
    });
    var dateText = dateStrs.join('、');

    // Chatwork送信
    sendDirectChatworkAlert_(clinic, targetName, dateText, pdfFile, sheetUrl);
  });
}

/**
 * 2. 編集時トリガー（チェックボックスのON/OFFによるグレーアウトとタイムスタンプ）
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== '変更依頼') return;

  var row = e.range.getRow();
  if (row === 1) return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headerMap = buildHeaderMap(headers);
  var idxDone = headerMap['完了'];
  var idxDate = headerMap['対応日'];

  if (idxDone === undefined || idxDate === undefined) return;

  var col = e.range.getColumn();
  if (col === (idxDone + 1)) {
    var lastCol = sheet.getLastColumn();
    var rangeToStyle = sheet.getRange(row, 1, 1, lastCol);
    
    if (e.value === 'TRUE') {
      var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
      sheet.getRange(row, idxDate + 1).setValue(now);
      rangeToStyle.setBackground('#e0e0e0'); // グレーアウト
    } else {
      sheet.getRange(row, idxDate + 1).clearContent();
      rangeToStyle.setBackground(null); // 背景色リセット
    }
  }
}

/**
 * 医師の現在の所属先を取得するヘルパー関数
 */
function getDoctorCurrentWork_(targetId) {
  var masterId = '1-Ss0Zo1TZkl01d208_zVpG9DSLWOwLh7DntsvvK81Fc';
  try {
    var ss = SpreadsheetApp.openById(masterId);
    var sheet = ss.getSheetByName('医師マスタ');
    var data = sheet.getDataRange().getValues();
    var headerMap = buildHeaderMap(data[0]);
    var colIki = headerMap['医籍登録番号'];
    var colWork = headerMap['現在の主な勤務先名'];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colIki]).trim() === targetId) {
        return String(data[i][colWork]).trim();
      }
    }
  } catch (e) {
    Logger.log('医師マスタからの所属取得に失敗: ' + e.message);
  }
  return '不明';
}

/**
 * 医師マスタ更新用ヘルパー関数
 */
function updateDoctorMaster_(targetId, newClinicName) {
  var masterId = '1-Ss0Zo1TZkl01d208_zVpG9DSLWOwLh7DntsvvK81Fc';
  try {
    var ss = SpreadsheetApp.openById(masterId);
    var sheet = ss.getSheetByName('医師マスタ');
    var data = sheet.getDataRange().getValues();
    
    var headerMap = buildHeaderMap(data[0]);
    var colIki = headerMap['医籍登録番号'];
    var colWork = headerMap['現在の主な勤務先名'];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][colIki]).trim() === targetId) {
        sheet.getRange(i + 1, colWork + 1).setValue(newClinicName);
        SpreadsheetApp.flush();
        return ss.getUrl() + '#gid=' + sheet.getSheetId() + '&range=' + sheet.getRange(i + 1, colWork + 1).getA1Notation();
      }
    }
  } catch (e) {
    Logger.log('医師マスタの更新に失敗: ' + e.message);
  }
  return null;
}

/**
 * 管理用Chatworkルーム（444531491）へ通知を送る関数
 */
function sendAdminMasterChatworkAlert_(baseClinic, doctorName, oldClinic, newClinic) {
  var roomId = '444531491';
  var token = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN');
  if (!token) return;

  var message = '[info][title]勤務先差し替え連絡[/title]' +
    '勤務先変更要望が拠点よりございました。\n\n' +
    '【申告拠点】' + baseClinic + '\n' +
    '対象医師名：' + doctorName + '\n' +
    '現在の所属：' + oldClinic + '\n' +
    '差し替え後の所属：' + newClinic + '\n' +
    '[/info]';

  var url = 'https://api.chatwork.com/v2/rooms/' + roomId + '/messages';
  var options = {
    "method": "post",
    "headers": { "X-ChatWorkToken": token },
    "payload": { "body": message },
    "muteHttpExceptions": true
  };

  try {
    UrlFetchApp.fetch(url, options);
  } catch (e) {
    Logger.log('管理ルームへの通知送信エラー: ' + e.message);
  }
}

/**
 * 変更差し替え用の即時Chatwork送信関数（各拠点向け）
 */
function sendDirectChatworkAlert_(clinic, doctorName, dateText, pdfFile, sheetUrl) {
  var roomId = clinic.chatId; 
  if (!roomId) return;
  
  var token = PropertiesService.getScriptProperties().getProperty('CHATWORK_API_TOKEN');
  if (!token) return;

  var toText = (clinic.leaderId || '') + '\n' + (clinic.sharedAccount || '') + '\n\n';
  var message = toText + 
    '[info][title]勤務先差し替え連絡[/title]' +
    '勤務先変更要望が医師よりございました。\n' +
    '差し替えを送付します。\n\n' +
    '医師名：' + doctorName + '\n' +
    '該当日：' + dateText + '\n\n' +
    '保存先URL：\n' + sheetUrl + '\n[/info]';

  try {
    var pdfBlob = pdfFile.getBlob();
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
    UrlFetchApp.fetch(url, {
      "method": "post",
      "headers": { "X-ChatWorkToken": token, "Content-Type": "multipart/form-data; boundary=" + boundary },
      "payload": payload,
      "muteHttpExceptions": true
    });
  } catch(e) {
    Logger.log('通知送信エラー: ' + e.message);
  }
}