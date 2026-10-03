/**
 * ========================================
 * 第7段階：舌下・オンライン別表の生成（外部スプシ保存版）
 * ========================================
 */

function generateQualTable(context, clinicNo, year, month, targetFolder) {
  var clinic = context.clinicMaster.byClinicNo[clinicNo];
  if (!clinic) return null;

  var clinicName = clinic.name;
  var sheetName = ('0' + month).slice(-2) + clinicName + '_別表';
  var fileName = year + '年' + ('0' + month).slice(-2) + '月_' + clinicName + '_資格表.pdf';

  var qualSs = SpreadsheetApp.openById(QUAL_SPREADSHEET_ID); // 外部スプシ（保存先）
  var ss = SpreadsheetApp.getActiveSpreadsheet();            // 担当くん本体
  
  // ★ 本体の「舌下登録・オンライン原本」をコピー元にする
  var templateSheet = ss.getSheetByName('舌下登録・オンライン原本');
  if (!templateSheet) throw new Error('担当くん本体に「舌下登録・オンライン原本」シートが見つかりません');

  // 1. 外部スプシに同名の古いシートがあれば削除して作り直す
  var existingSheet = qualSs.getSheetByName(sheetName);
  if (existingSheet) qualSs.deleteSheet(existingSheet);

  var sheet = templateSheet.copyTo(qualSs);
  sheet.setName(sheetName);

  // 2. ヘッダー情報の書き換え
  var isInternalMed = (clinicName === '北葛西' || clinicName === '亀有');
  var deptName = isInternalMed ? '内科' : '小児科';

  replacePlaceholdersOrdered(sheet, {
    period: [],
    simple: { '{{拠点名}}': clinicName, '{{診療科}}': deptName }
  });

  if (clinicName === '北葛西') {
    sheet.getRange('D3').setValue('18:00-20:00');
  }

  // 3. シフトデータの振り分け
  var startKey = dateKey(new Date(year, month - 1, 1));
  var endKey = dateKey(new Date(year, month, 0));
  
  var qualSchedule = {};
  var rows = context.shiftRows;
  var slot1 = { start: 9 * 60, end: 13 * 60 };
  var slot2 = { start: 15 * 60, end: 18 * 60 };
  var slot3 = { start: 18 * 60, end: 21 * 60 };

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r.clinicNo !== clinicNo) continue;
    if (r.dateKey < startKey || r.dateKey > endKey) continue;

    if (!qualSchedule[r.dateKey]) qualSchedule[r.dateKey] = { s1: [], s2: [], s3: [] };

    var inS1 = isQualOverlap_(r.start, r.end, slot1.start, slot1.end);
    var inS2 = isQualOverlap_(r.start, r.end, slot2.start, slot2.end);
    var inS3 = isQualOverlap_(r.start, r.end, slot3.start, slot3.end);

    if (inS1) qualSchedule[r.dateKey].s1.push(r);
    if (inS2) qualSchedule[r.dateKey].s2.push(r);
    if (inS3) qualSchedule[r.dateKey].s3.push(r);
  }

  // 4. 日付ごとの書き込み処理
  var w = ['日','月','火','水','木','金','土'];
  var START_ROW = 3; 

  for (var day = 1; day <= 31; day++) {
    var rowIdx = START_ROW + day - 1; 
    var rowRange = sheet.getRange(rowIdx, 1, 1, 4);

    var d = new Date(year, month - 1, day);
    if (d.getMonth() + 1 === month) {
      var dKey = dateKey(d);
      var isHol = context.holidays[dKey] === true;
      var dayStr = w[d.getDay()] + (isHol ? '祝' : '');

      var dateCell = sheet.getRange(rowIdx, 1);
      dateCell.setValue(month + '/' + day + '(' + dayStr + ')');
      
      if (d.getDay() === 0 || isHol) {
        dateCell.setBackground('#f4cccc');
      } else if (d.getDay() === 6) {
        dateCell.setBackground('#cfe2f3');
      } else {
        dateCell.setBackground(null);
      }
      
      var dailyData = qualSchedule[dKey] || { s1: [], s2: [], s3: [] };
      writeQualSlot_(sheet, rowIdx, 2, dailyData.s1); 
      writeQualSlot_(sheet, rowIdx, 3, dailyData.s2); 
      writeQualSlot_(sheet, rowIdx, 4, dailyData.s3); 
    } else {
      rowRange.setValue('');
      rowRange.setBackground('#f3f3f3');
    }
  }

  SpreadsheetApp.flush();

  // 5. Chatwork配信用に、このシート単体のPDFを作成
  var url = qualSs.getUrl().replace(/edit$/, '') + 'export?'
    + 'exportFormat=pdf&format=pdf'
    + '&gid=' + sheet.getSheetId()
    + '&size=A4&portrait=true&scale=4'
    + '&top_margin=0.25&bottom_margin=0.25&left_margin=0.25&right_margin=0.25'
    + '&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false';

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  var pdfFile = null;
  if (response.getResponseCode() === 200) {
    if (targetFolder) {
      var existingFiles = targetFolder.getFilesByName(fileName);
      while (existingFiles.hasNext()) existingFiles.next().setTrashed(true);
      var blob = response.getBlob().setName(fileName);
      pdfFile = targetFolder.createFile(blob);
    }
  } else {
    Logger.log('PDF生成エラー: ' + response.getContentText());
  }

  // 6. 戻り値
  var sheetUrl = qualSs.getUrl() + '#gid=' + sheet.getSheetId();
  return {
    sheetUrl: sheetUrl,
    pdfFile: pdfFile
  };
}

function isQualOverlap_(start, end, slotStart, slotEnd) {
  if (start === null || end === null) return false;
  return start < slotEnd && end > slotStart;
}

function writeQualSlot_(sheet, row, col, doctors) {
  var cell = sheet.getRange(row, col);
  
  if (!doctors || doctors.length === 0) {
    cell.setValue('');
    cell.setBackground('#f3f3f3');
    return;
  }
  cell.setBackground(null);

  var uniqueDocs = [];
  var seen = {};
  doctors.forEach(function(doc) {
    if (!seen[doc.ikiNo]) {
      seen[doc.ikiNo] = true;
      uniqueDocs.push(doc);
    }
  });

  var fullText = '';
  var redSegments = [];

  uniqueDocs.forEach(function(doc, idx) {
    if (idx > 0) fullText += '\n';

    var zStatus = doc.zekka || 'NG';
    var oStatus = doc.online || 'NG';
    
    var lineText = doc.name + '(舌下:' + zStatus + ') (オンライン:' + oStatus + ')';
    var lineStart = fullText.length;
    fullText += lineText;

    var regex = /NG/g;
    var match;
    while ((match = regex.exec(lineText)) !== null) {
      var globalStart = lineStart + match.index;
      redSegments.push({ start: globalStart, end: globalStart + 2 });
    }
  });

  var builder = SpreadsheetApp.newRichTextValue().setText(fullText);
  var redStyle = SpreadsheetApp.newTextStyle().setForegroundColor('#E60012').setBold(true).build();

  redSegments.forEach(function(seg) {
    builder.setTextStyle(seg.start, seg.end, redStyle);
  });

  cell.setRichTextValue(builder.build());
  cell.setWrap(true).setVerticalAlignment('middle');
}