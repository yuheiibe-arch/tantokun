/**
 * ========================================
 * 第3段階：メイン処理・データ準備
 * ========================================
 */

var GENPON_NAME = '原本';
var FIRST_BLOCK_START_ROW = 4;
var SECOND_BLOCK_START_ROW = 39;
var SECOND_BLOCK_SETS = 16;

// 診療科ラベルを付ける特殊拠点
var DEPT_LABEL_CLINICS = ['北葛西', '亀有'];

/**
 * 祝日データを取得する
 */
function getHolidayMap(year, month) {
  var ssData = SpreadsheetApp.openById(TOOL_SPREADSHEET_ID).getSheetByName('データセット');
  var data = ssData.getDataRange().getValues();
  var fileId = '';
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === '祝日') { fileId = data[i][1]; break; }
  }
  if (!fileId) return {};

  var fy = (month <= 3) ? year - 1 : year;
  var sheetName = fy + '年度';
  var ss;
  try {
    ss = SpreadsheetApp.openById(fileId);
  } catch(e) { return {}; }

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return {};

  var values = sheet.getDataRange().getValues();
  var hMap = {};
  for (var r = 1; r < values.length; r++) {
    var d = values[r][0];
    var isHoliday = values[r][3];
    if (Object.prototype.toString.call(d) === '[object Date]' && isHoliday) {
      hMap[dateKey(d)] = true;
    }
  }
  return hMap;
}

/**
 * 休館日データを取得する
 */
function getClosedDaysMap(year, month) {
  var opened = openSource('休館日');
  var values = opened.values;
  var h = opened.headerMap;
  var idxDate = requireColumn(h, '日付', '休館日');
  var idxClinic = requireColumn(h, '拠点名', '休館日');
  var idxTime = requireColumn(h, '時間', '休館日');

  var startKey = dateKey(new Date(year, month - 1, 1));
  var endKey = dateKey(new Date(year, month, 0));

  var closedMap = {};
  for (var i = 1; i < values.length; i++) {
    var d = values[i][idxDate];
    if (Object.prototype.toString.call(d) !== '[object Date]') continue;
    var dKey = dateKey(d);
    if (dKey < startKey || dKey > endKey) continue;

    var cName = String(values[i][idxClinic]).trim();
    var tRange = String(values[i][idxTime]).trim();

    if (!closedMap[dKey]) closedMap[dKey] = [];
    closedMap[dKey].push({ clinicName: cName, time: tRange });
  }
  return closedMap;
}

/**
 * 高速化のため、重いデータを1回だけ読み込んでcontextを作る。
 */
function buildContext(year, month) {
  var opened = openSource('確定シフト');
  var values = opened.values;
  var h = opened.headerMap;
  var idxClinicNo = requireColumn(h, 'クリニックNo', '確定シフト');
  var idxIki      = requireColumn(h, '医籍番号', '確定シフト');
  var idxName     = requireColumn(h, '名前', '確定シフト');
  var idxDate     = requireColumn(h, '勤務日', '確定シフト');
  var idxStart    = requireColumn(h, '勤務開始時間', '確定シフト');
  var idxEnd      = requireColumn(h, '勤務終了時間', '確定シフト');
  var idxDept     = optionalColumn(h, '診療科');

  var shiftRows = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var d = row[idxDate];
    if (Object.prototype.toString.call(d) !== '[object Date]') continue;
    shiftRows.push({
      clinicNo: row[idxClinicNo],
      ikiNo: normalizeId(row[idxIki]),
      name: String(row[idxName]).trim(),
      dateKey: dateKey(d),
      start: timeToMinutes(row[idxStart]),
      end: timeToMinutes(row[idxEnd]),
      dept: (idxDept !== null) ? String(row[idxDept]).trim() : ''
    });
  }
  return {
    shiftRows: shiftRows,
    doctorMaster: getDoctorMaster(),
    clinicMaster: getClinicMaster(),
    holidays: getHolidayMap(year, month),
    closedDays: getClosedDaysMap(year, month)
  };
}

/**
 * 1拠点を生成（context使用）。これが本番の生成単位。
 */
function generateScheduleWithContext(context, clinicNo, year, month) {
  var ss = SpreadsheetApp.openById(TOOL_SPREADSHEET_ID);
  var clinic = context.clinicMaster.byClinicNo[clinicNo];
  if (!clinic) throw new Error('クリニックNo ' + clinicNo + ' が拠点マスターにありません。');

  var lastDay = new Date(year, month, 0).getDate();
  var startKey = dateKey(new Date(year, month - 1, 1));
  var endKey   = dateKey(new Date(year, month - 1, lastDay));

  var schedule = collectScheduleFromContext(context, clinicNo, startKey, endKey);
  mergeSameDoctors(schedule);
  
  var useLabel = isDeptLabelClinic(clinic.name);
  sortScheduleByDeptAndTime(schedule, useLabel);

  var sheetName = ('0' + month).slice(-2) + clinic.name;
  var sheet = prepareOutputSheet(ss, clinic.name, sheetName);

  var firstPeriod  = formatPeriod(year, month, 1, month, 15);
  var secondPeriod = formatPeriod(year, month, 16, month, lastDay);
  
  var deptText = useLabel ? '小児科・内科' : '小児科';

  replacePlaceholdersOrdered(sheet, {
    period: [firstPeriod, secondPeriod],
    simple: {
      '{{月}}': month, 
      '{{管理医師}}': clinic.director || '',
      '{{拠点名}}': clinic.name, 
      '{{診療科}}': deptText
    }
  });

  fillBlock(sheet, schedule, context, clinic, year, month, 1, 15, FIRST_BLOCK_START_ROW, useLabel);
  fillBlock(sheet, schedule, context, clinic, year, month, 16, lastDay, SECOND_BLOCK_START_ROW, useLabel);

  trimUnusedRows(sheet, lastDay);
  redrawBorders(sheet, lastDay);

  return sheet;
}

/**
 * 1拠点を生成（context無し版。単発生成用）。
 */
function generateSchedule(clinicNo, year, month) {
  var context = buildContext(year, month);
  var sheet = generateScheduleWithContext(context, clinicNo, year, month);
  SpreadsheetApp.flush();
  return sheet;
}

/**
 * 全拠点一括生成。
 */
function generateAllClinics(year, month) {
  var t0 = new Date().getTime();
  var context = buildContext(year, month);
  var n = 0;
  context.clinicMaster.list.forEach(function(clinic) {
    generateScheduleWithContext(context, clinic.clinicNo, year, month);
    n++;
  });
  SpreadsheetApp.flush();
  var sec = Math.round((new Date().getTime() - t0) / 1000);
  Logger.log('一括生成完了: ' + n + '拠点 / ' + year + '年' + month + '月 / ' + sec + '秒');
}

function isDeptLabelClinic(clinicName) {
  return DEPT_LABEL_CLINICS.indexOf(clinicName) !== -1;
}