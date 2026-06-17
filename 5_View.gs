/**
 * ========================================
 * 第5段階：ビュー（シートへの書き込み・レイアウト処理）
 * ========================================
 */

function prepareOutputSheet(ss, clinicName, newSheetName) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var m = sheets[i].getName().match(/^(\d{2})(.+)$/);
    if (m && m[2] === clinicName) { ss.deleteSheet(sheets[i]); break; }
  }
  var genpon = ss.getSheetByName(GENPON_NAME);
  var copy = genpon.copyTo(ss);
  copy.setName(newSheetName);
  ss.setActiveSheet(copy);
  ss.moveActiveSheet(genpon.getIndex() + 1);
  return copy;
}

function formatPeriod(year, m1, d1, m2, d2) {
  var w = ['日','月','火','水','木','金','土'];
  var s = new Date(year, m1 - 1, d1), e = new Date(year, m2 - 1, d2);
  return m1 + '月' + d1 + '日(' + w[s.getDay()] + ')〜' + m2 + '月' + d2 + '日(' + w[e.getDay()] + ')';
}

function replacePlaceholdersOrdered(sheet, opts) {
  var range = sheet.getDataRange();
  var values = range.getValues();
  var pi = 0;
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var cell = values[r][c];
      if (typeof cell !== 'string' || cell.indexOf('{{') === -1) continue;
      var v = cell;
      if (v.indexOf('{{期間}}') !== -1) {
        var pv = (opts.period[pi] !== undefined) ? opts.period[pi] : opts.period[opts.period.length - 1];
        v = v.split('{{期間}}').join(pv); pi++;
      }
      Object.keys(opts.simple).forEach(function(k) {
        if (v.indexOf(k) !== -1) v = v.split(k).join(opts.simple[k]);
      });
      values[r][c] = v;
    }
  }
  range.setValues(values);
}

function fillBlock(sheet, schedule, context, clinic, year, month, startDay, endDay, blockStartRow, useLabel) {
  var w = ['日','月','火','水','木','金','土'];
  for (var day = startDay; day <= endDay; day++) {
    var topRow = blockStartRow + (day - startDay) * 2;
    var bottomRow = topRow + 1;
    var d = new Date(year, month - 1, day);
    var dKey = dateKey(d);
    
    var dateRange = sheet.getRange(topRow, 1, 2, 1);
    dateRange.getCell(1,1).setValue(month + '月' + day + '日(' + w[d.getDay()] + ')');

    // 土日祝の色分け処理
    var isSat = d.getDay() === 6;
    var isSun = d.getDay() === 0;
    var isHol = context.holidays[dKey] === true;
    
    if (isSun || isHol) {
      dateRange.setBackground('#f4cccc'); // 赤
    } else if (isSat) {
      dateRange.setBackground('#cfe2f3'); // 水色
    } else {
      dateRange.setBackground(null); // 平日はリセット
    }

    // 休館日の判定
    var cDays = context.closedDays[dKey] || [];
    var isAllDayClosed = false;
    var isPmClosed = false;
    for (var c = 0; c < cDays.length; c++) {
      var targetClinic = cDays[c].clinicName;
      if (targetClinic === '全拠点' || targetClinic.indexOf(clinic.name) !== -1) {
        if (cDays[c].time === '全日') isAllDayClosed = true;
        if (cDays[c].time === '午後夜間' || cDays[c].time === '午後') isPmClosed = true;
      }
    }

    var slot = schedule[dKey] || { am: [], pm: [] };

    if (isAllDayClosed) {
      writeClosedSlot(sheet, topRow, bottomRow, 2, '休館日');
      writeClosedSlot(sheet, topRow, bottomRow, 3, '休館日');
    } else {
      writeSlot(sheet, topRow, bottomRow, 2, slot.am, context.doctorMaster, useLabel);
      if (isPmClosed) {
        writeClosedSlot(sheet, topRow, bottomRow, 3, '休館日');
      } else {
        writeSlot(sheet, topRow, bottomRow, 3, slot.pm, context.doctorMaster, useLabel);
      }
    }
  }
}

function writeClosedSlot(sheet, topRow, bottomRow, col, text) {
  sheet.getRange(topRow, col).setValue(text).setFontSize(14).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sheet.getRange(bottomRow, col).setValue('');
}

function writeSlot(sheet, topRow, bottomRow, col, doctors, doctorMaster, useLabel) {
  var topCell = sheet.getRange(topRow, col);
  var bottomCell = sheet.getRange(bottomRow, col);

  if (!doctors || doctors.length === 0) {
    topCell.setValue('');
    bottomCell.setValue('');
    return;
  }

  // 空欄パディング処理
  var displayDoctors = doctors.slice();
  if (col === 3 && displayDoctors.length > 0) {
    var allNight = true;
    for (var k = 0; k < displayDoctors.length; k++) {
      if (displayDoctors[k].start !== null && displayDoctors[k].start < 18 * 60) { 
        allNight = false; break; 
      }
    }
    if (allNight) {
      displayDoctors.unshift({ name: ' ', dept: '', ikiNo: 'DUMMY' });
    }
  }

  topCell.setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  bottomCell.setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');

  // ★人数に応じてフォントサイズを動的に変更
  var numDocs = displayDoctors.length;
  var FONT_SIZE_NAME, FONT_SIZE_WORK, FONT_SIZE_LABEL;

  if (!useLabel) {
    // 小児科単独などラベルがない拠点
    if (numDocs === 1) {
      FONT_SIZE_NAME = 14; FONT_SIZE_WORK = 10;
    } else if (numDocs === 2) {
      FONT_SIZE_NAME = 12; FONT_SIZE_WORK = 9;
    } else {
      FONT_SIZE_NAME = 11; FONT_SIZE_WORK = 8; // 3名以上
    }
  } else {
    // 亀有・北葛西など診療科ラベルがある拠点
    if (numDocs === 1) {
      FONT_SIZE_NAME = 13; FONT_SIZE_LABEL = 9; FONT_SIZE_WORK = 9;
    } else if (numDocs === 2) {
      FONT_SIZE_NAME = 11; FONT_SIZE_LABEL = 8; FONT_SIZE_WORK = 8;
    } else {
      FONT_SIZE_NAME = 10; FONT_SIZE_LABEL = 8; FONT_SIZE_WORK = 8; // 3名以上
    }
  }

  if (!useLabel) {
    var names = displayDoctors.map(function(doc) { return doc.name; }).join(' / ');
    topCell.setValue(names).setFontSize(FONT_SIZE_NAME);

    var works = displayDoctors.map(function(doc) {
      if (doc.ikiNo === 'DUMMY') return ' ';
      // ★ フィルター関数を通す
      return getValidWorkName(doctorMaster[doc.ikiNo]);
    }).join(' / ');
    bottomCell.setValue(works).setFontSize(FONT_SIZE_WORK);
    
  } else {
    var fullText = '';
    var workText = '';
    var segments = [];
    var internDocs = [];
    var pediaDocs = [];
    
    displayDoctors.forEach(function(doc) {
      if (doc.ikiNo === 'DUMMY') {
        var firstReal = displayDoctors[1];
        if (firstReal && firstReal.dept.indexOf('内科') !== -1) internDocs.push(doc);
        else pediaDocs.push(doc);
      } else if (doc.dept.indexOf('内科') !== -1) {
        internDocs.push(doc);
      } else {
        pediaDocs.push(doc);
      }
    });
    
    if (displayDoctors.length <= 3) {
      var lines = [];
      if (pediaDocs.length > 0) lines = lines.concat(pediaDocs);
      if (internDocs.length > 0) lines = lines.concat(internDocs);

      lines.forEach(function(doc, idx) {
        if (idx > 0) fullText += ' / ';
        var nameStart = fullText.length;
        fullText += doc.name;
        segments.push({ start: nameStart, end: fullText.length, size: FONT_SIZE_NAME });
        
        var label = deptLabel(doc.dept);
        if (label && doc.ikiNo !== 'DUMMY') {
          var labelStart = fullText.length;
          fullText += label;
          segments.push({ start: labelStart, end: fullText.length, size: FONT_SIZE_LABEL });
        }
      });

      workText = lines.map(function(doc) {
        if (doc.ikiNo === 'DUMMY') return ' ';
        // ★ フィルター関数を通す
        return getValidWorkName(doctorMaster[doc.ikiNo]);
      }).join(' / ');
      
    } else {
      var catLines = [];
      if (pediaDocs.length > 0) catLines.push(pediaDocs);
      if (internDocs.length > 0) catLines.push(internDocs);

      var workLines = [];
      
      catLines.forEach(function(lineDocs, lineIdx) {
        if (lineIdx > 0) fullText += '\n';
        
        var currentWorkLine = [];
        lineDocs.forEach(function(doc, docIdx) {
          if (docIdx > 0) fullText += ' / ';
          
          var nameStart = fullText.length;
          fullText += doc.name;
          segments.push({ start: nameStart, end: fullText.length, size: FONT_SIZE_NAME });
          
          var label = deptLabel(doc.dept);
          if (label && doc.ikiNo !== 'DUMMY') {
            var labelStart = fullText.length;
            fullText += label;
            segments.push({ start: labelStart, end: fullText.length, size: FONT_SIZE_LABEL });
          }
          // ★ フィルター関数を通す
          currentWorkLine.push(doc.ikiNo === 'DUMMY' ? ' ' : getValidWorkName(doctorMaster[doc.ikiNo]));
        });
        workLines.push(currentWorkLine.join(' / '));
      });
      workText = workLines.join('\n');
    }

    var builder = SpreadsheetApp.newRichTextValue().setText(fullText);
    var baseStyle = SpreadsheetApp.newTextStyle().setFontSize(FONT_SIZE_LABEL).build();
    builder.setTextStyle(0, fullText.length, baseStyle);
    
    segments.forEach(function(seg) {
      var style = SpreadsheetApp.newTextStyle().setFontSize(seg.size).build();
      builder.setTextStyle(seg.start, seg.end, style);
    });
    
    topCell.setRichTextValue(builder.build());
    bottomCell.setValue(workText).setFontSize(FONT_SIZE_WORK);
  }
}

function deptLabel(dept) {
  if (dept.indexOf('小児') !== -1) return '（小児科）';
  if (dept.indexOf('内科') !== -1) return '（内科）';
  return '';
}

function trimUnusedRows(sheet, lastDay) {
  var usedSets = lastDay - 16 + 1;
  var unusedSets = SECOND_BLOCK_SETS - usedSets;
  if (unusedSets <= 0) return;
  var dataLastRow = SECOND_BLOCK_START_ROW + SECOND_BLOCK_SETS * 2 - 1;
  sheet.deleteRows(dataLastRow - unusedSets * 2 + 1, unusedSets * 2);
}

function redrawBorders(sheet, lastDay) {
  drawBlockBorders(sheet, FIRST_BLOCK_START_ROW, 15);
  drawBlockBorders(sheet, SECOND_BLOCK_START_ROW, lastDay - 16 + 1);
}

function drawBlockBorders(sheet, startRow, setCount) {
  for (var i = 0; i < setCount; i++) {
    var top = startRow + i * 2;
    sheet.getRange(top, 1, 2, 3).setBorder(true, true, true, true, true, false,
      'black', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(top, 2, 1, 2).setBorder(null, null, true, null, null, null,
      'black', SpreadsheetApp.BorderStyle.DOTTED);
  }
}

/**
 * ★所属先（勤務先）の表記をチェックし、ネガティブなものや空欄を統一する
 */
function getValidWorkName(workName) {
  if (!workName) return 'キャップスクリニック';
  
  var w = String(workName).trim();
  var negativeWords = ['フリーランス', '休職', '休職中', 'なし'];
  
  if (w === '' || negativeWords.indexOf(w) !== -1) {
    return 'キャップスクリニック';
  }
  
  return w;
}