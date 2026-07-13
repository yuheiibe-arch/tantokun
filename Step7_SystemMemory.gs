/**
 * ========================================
 * ⚙️ データ記憶 ＆ ハッシュ計算モジュール
 * ========================================
 * システムが前回の状態を記憶し、変更を検知するための裏側の仕組みです。
 * ※差分比較ロジックは本番デーモン内に統合されたため、ここには記憶システムのみ残します。
 */

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
 * （データの文字列を暗号化して短い文字列にし、変更があったかを高速判定する）
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