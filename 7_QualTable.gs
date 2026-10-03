/**
 * ========================================
 * 第7段階：舌下・オンライン別表の生成（外部スプシ保存版・PDF出力なし）
 * ========================================
 */

function generateQualTable(context, clinicNo, year, month, targetFolder) {
  var clinic = context.clinicMaster.byClinicNo[clinicNo];
  if (!clinic) return null;

  var clinicName = clinic.name;
  var isInternalMed = (clinicName === '北葛西' || clinicName === '亀有');
  var targetDepts = isInternalMed ? ['小児科', '内科'] : ['小児科'];

  var qualSs = SpreadsheetApp.openById(QUAL_SPREADSHEET_ID); // 外部スプシ（保存先）
  var ss = SpreadsheetApp.getActiveSpreadsheet();            // 担当くん本体
  
  var templateSheet = ss.getSheetByName('舌下登録・オンライン原本');
  if (!templateSheet) throw new Error('担当くん本体に「舌下登録・オンライン原本」シートが見つかりません');

  var results = []; 

  for (var d = 0; d < targetDepts.length; d++) {
    var currentDept = targetDepts[d];
    
    var sheetNameSuffix = isInternalMed ? '（' + currentDept + '）別表' : '_別表';
    var sheetName = ('0' + month).slice(-2) + clinicName + sheetNameSuffix;

    var existingSheet = qualSs.getSheetByName(sheetName);
    if (existingSheet) qualSs.deleteSheet(existingSheet);

    var sheet = templateSheet.copyTo(qualSs);
    sheet.setName(sheetName);

    replacePlaceholdersOrdered(sheet, {
      period: [],
      simple: { '{{拠点名}}': clinicName, '{{診療科}}': currentDept }
    });

    if (clinicName === '北葛西') {
      sheet.getRange('D3').setValue('18:00-20:00');
    }

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

      if (String(r.dept).indexOf('小児科ワクチン専任(対象：小児～成人)') !== -1) {
        continue;
      }

      var docIsInternal = (String(r.dept).indexOf('内科') !== -1);
      if (currentDept === '内科' && !docIsInternal) continue; 
      if (currentDept === '小児科' && docIsInternal) continue; 

      var docData = context.doctorMaster[r.ikiNo] || { zekka: 'NG', online: 'NG' };
      r.zekka = docData.zekka === 'OK' ? 'OK' : 'NG';
      r.online = docData.online === 'OK' ? 'OK' : 'NG';

      if (!qualSchedule[r.dateKey]) qualSchedule[r.dateKey] = { s1: [], s2: [], s3: [] };

      var inS1 = isQualOverlap_(r.start, r.end, slot1.start, slot1.end);
      var inS2 = isQualOverlap_(r.start, r.end, slot2.start, slot2.end);
      var inS3 = isQualOverlap_(r.start, r.end, slot3.start, slot3.end);

      if (inS1) qualSchedule[r.dateKey].s1.push(r);
      if (inS2) qualSchedule[r.dateKey].s2.push(r);
      if (inS3) qualSchedule[r.dateKey].s3.push(r);
    }

    var w = ['日','月','火','水','木','金','土'];
    var START_ROW = 3; 

    for (var day = 1; day <= 31; day++) {
      var rowIdx = START_ROW + day - 1; 
      var rowRange = sheet.getRange(rowIdx, 1, 1, 4);

      var dObj = new Date(year, month - 1, day);
      if (dObj.getMonth() + 1 === month) {
        var dKey = dateKey(dObj);
        var isHol = context.holidays[dKey] === true;
        var dayStr = w[dObj.getDay()] + (isHol ? '祝' : '');

        var dateCell = sheet.getRange(rowIdx, 1);
        dateCell.setValue(month + '/' + day + '(' + dayStr + ')');
        
        if (dObj.getDay() === 0 || isHol) {
          dateCell.setBackground('#f4cccc');
        } else if (dObj.getDay() === 6) {
          dateCell.setBackground('#cfe2f3');
        } else {
          dateCell.setBackground(null);
        }
        
        var dailyData = qualSchedule[dKey] || { s1: [], s2: [], s3: [] };
        
        writeQualSlotSingleDept_(sheet, rowIdx, 2, dailyData.s1, currentDept); 
        writeQualSlotSingleDept_(sheet, rowIdx, 3, dailyData.s2, currentDept); 
        writeQualSlotSingleDept_(sheet, rowIdx, 4, dailyData.s3, currentDept); 
      } else {
        rowRange.setValue('');
        rowRange.setBackground('#f3f3f3');
      }
    }

    SpreadsheetApp.flush();

    results.push({
      sheetUrl: qualSs.getUrl() + '#gid=' + sheet.getSheetId(),
      pdfFile: null // PDF作成は廃止
    });
  }

  return {
    sheetUrl: results[0].sheetUrl,
    pdfFile: null
  };
}

function isQualOverlap_(start, end, slotStart, slotEnd) {
  if (start === null || end === null) return false;
  return start < slotEnd && end > slotStart;
}

function writeQualSlotSingleDept_(sheet, row, col, doctors, targetDept) {
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

    var lineText = '';
    var oStatus = doc.online || 'NG';
    
    if (targetDept === '内科') {
      lineText = doc.name + '(オンライン:' + oStatus + ')';
    } else {
      var zStatus = doc.zekka || 'NG';
      lineText = doc.name + '(舌下:' + zStatus + ') (オンライン:' + oStatus + ')';
    }
    
    var lineStart = fullText.length;
    fullText += lineText;
    extractNgSegments_(lineText, lineStart, redSegments);
  });

  var builder = SpreadsheetApp.newRichTextValue().setText(fullText);
  var redStyle = SpreadsheetApp.newTextStyle().setForegroundColor('#E60012').setBold(true).build();

  redSegments.forEach(function(seg) {
    builder.setTextStyle(seg.start, seg.end, redStyle);
  });

  cell.setRichTextValue(builder.build());
  cell.setWrap(true).setVerticalAlignment('middle');
}

function extractNgSegments_(lineText, lineStart, redSegments) {
  var regex = /NG/g;
  var match;
  while ((match = regex.exec(lineText)) !== null) {
    var globalStart = lineStart + match.index;
    redSegments.push({ start: globalStart, end: globalStart + 2 });
  }
}