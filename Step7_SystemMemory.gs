/**
 * ========================================
 * 差分比較 ＆ データ記憶モジュール
 * ========================================
 */

/**
 * 過去のシフトと今のシフトを比較し、「空欄から名前が入った」日付・枠を抽出する
 */
function getFilledVacancies_(oldSched, newSched) {
  var filledMsgs = [];
  if (!oldSched) return filledMsgs;
  
  Object.keys(newSched).forEach(function(dKey) {
    var dateStr = parseInt(dKey.substring(4, 6), 10) + "月" + parseInt(dKey.substring(6, 8), 10) + "日";
    ['am', 'pm'].forEach(function(slot) {
      var oldDocs = (oldSched[dKey] && oldSched[dKey][slot]) ? oldSched[dKey][slot] : [];
      var newDocs = newSched[dKey][slot] || [];

      var oldHasDoc = oldDocs.some(function(d) { return d.name && d.name.replace(/\s/g, '') !== ""; });
      var newHasDoc = newDocs.some(function(d) { return d.name && d.name.replace(/\s/g, '') !== ""; });

      if (!oldHasDoc && newHasDoc) {
        var slotName = (slot === 'am') ? '午前' : '午後(夜間含む)';
        filledMsgs.push(dateStr + "の" + slotName);
      }
    });
  });
  return filledMsgs;
}

/**
 * システム記憶用：状態の保存
 */
function saveScheduleState_(clinicNo, year, month, scheduleObj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = '⚙️_SystemData';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.hideSheet();
  }
  var key = year + '_' + month + '_' + clinicNo;
  var data = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == key) {
      sheet.getRange(i + 1, 2).setValue(JSON.stringify(scheduleObj));
      found = true;
      break;
    }
  }
  if (!found) sheet.appendRow([key, JSON.stringify(scheduleObj)]);
}

/**
 * システム記憶用：状態の読み込み
 */
function getSavedScheduleState_(clinicNo, year, month) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('⚙️_SystemData');
  if (!sheet) return null;
  var key = year + '_' + month + '_' + clinicNo;
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == key) return JSON.parse(data[i][1]);
  }
  return null;
}

/**
 * ハッシュ計算用関数
 */
function computeMD5_(input) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input, Utilities.Charset.UTF_8);
  var hashStr = '';
  for (var i = 0; i < rawHash.length; i++) {
    var byteVal = rawHash[i];
    if (byteVal < 0) byteVal += 256;
    var byteString = byteVal.toString(16);
    if (byteString.length == 1) byteString = "0" + byteString;
    hashStr += byteString;
  }
  return hashStr;
}