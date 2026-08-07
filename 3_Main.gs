/**
 * ========================================
 * 第3段階：メイン処理・データ準備（保険機能・ヘッダー揺れ対応版）
 * ========================================
 */

var GENPON_NAME = '原本';
var FIRST_BLOCK_START_ROW = 4;
var SECOND_BLOCK_START_ROW = 39;
var SECOND_BLOCK_SETS = 16;

// 診療科ラベルを付ける特殊拠点
var DEPT_LABEL_CLINICS = ['北葛西', '亀有'];

/**
 * 祝日データを取得する（共通の仕組みを利用するように修正済み）
 */
function getHolidayMap(year, month) {
  var sources = getDataSources();
  var src = sources['祝日'];
  
  if (!src || !src.id) {
    Logger.log('【警告】データセットに「祝日」の設定がないか、URLが不正です。');
    return {};
  }

  // 今年の年度を計算（1〜3月は前年扱い）
  var fy = (month <= 3) ? year - 1 : year;
  var sheetName = fy + '年度';
  
  var ss;
  try {
    ss = SpreadsheetApp.openById(src.id);
  } catch(e) { 
    Logger.log('【警告】祝日ファイルのオープンに失敗しました: ' + e.message);
    return {}; 
  }

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('【警告】祝日ファイルに「' + sheetName + '」シートが見つかりません。');
    return {};
  }

  var values = sheet.getDataRange().getValues();
  var hMap = {};
  for (var r = 1; r < values.length; r++) {
    var d = values[r][0];         // A列：日付
    var isHoliday = values[r][3]; // D列：祝日フラグ
    
    // 日付型であり、かつ祝日フラグ（TRUEなど）が入っている場合のみ登録
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
 * 複数の見出し名の候補から見つかったものを返すヘルパー（ヘッダー揺れ対応）
 */
function requireColumnAny(headerMap, headerNames, sheetLabel) {
  for (var i = 0; i < headerNames.length; i++) {
    if (headerNames[i] in headerMap) {
      return headerMap[headerNames[i]];
    }
  }
  var available = Object.keys(headerMap).join('", "');
  throw new Error(
    '「' + sheetLabel + '」に見出し「' + headerNames.join('」または「') + '」が見つかりません。\n' +
    '存在する見出し: "' + available + '"'
  );
}

/**
 * 高速化のため、重いデータを1回だけ読み込んでcontextを作る。
 * ★改修：権限エラーやデータ0件時にフォールバック（保険）する機能を搭載。
 *        ヘッダー名の微妙な違いも吸収します。
 */
function buildContext(year, month) {
  var today = new Date();
  var currentYear = today.getFullYear();
  var currentMonth = today.getMonth() + 1;
  
  // 今月なら「当月」、来月なら「確定」を第一候補とする
  var isCurrentMonth = (year === currentYear && month === currentMonth);
  var primarySource = isCurrentMonth ? '当月シフト' : '確定シフト';
  var fallbackSource = isCurrentMonth ? '確定シフト' : '当月シフト';
  
  // 第一候補 → 第二候補（保険） の順で試行
  var attemptSources = [primarySource, fallbackSource];
  var shiftRows = [];

  for (var s = 0; s < attemptSources.length; s++) {
    var sourceName = attemptSources[s];
    Logger.log('📂 試行: データソース 【 ' + sourceName + ' 】');
    
    try {
      var opened = openSource(sourceName);
      var values = opened.values;
      var h = opened.headerMap;
      
      // ヘッダー名が微妙に違っても吸収
      var idxClinicNo = requireColumnAny(h, ['クリニックNo', '拠点No', 'クリニック番号'], sourceName);
      var idxIki      = requireColumnAny(h, ['医籍番号', '医師番号', '医籍登録番号'], sourceName);
      var idxName     = requireColumnAny(h, ['名前', '氏名', '医師名'], sourceName);
      var idxDate     = requireColumnAny(h, ['勤務日', '日付', '出勤日'], sourceName);
      var idxStart    = requireColumnAny(h, ['勤務開始時間', '開始時間', '勤務開始'], sourceName);
      var idxEnd      = requireColumnAny(h, ['勤務終了時間', '終了時間', '勤務終了'], sourceName);
      var idxDept = optionalColumn(h, '診療科') || optionalColumn(h, '科目') || null;

      var tempRows = [];
      var hasTargetMonthData = false;

      for (var i = 1; i < values.length; i++) {
        var row = values[i];
        var d = row[idxDate];
        if (Object.prototype.toString.call(d) !== '[object Date]') continue;
        
        // 対象年月のデータが含まれているかチェック
        if (d.getFullYear() === year && (d.getMonth() + 1) === month) {
          hasTargetMonthData = true;
        }
        
        var parsedClinicNo = parseInt(row[idxClinicNo], 10);
        tempRows.push({
          clinicNo: parsedClinicNo,
          ikiNo: normalizeId(row[idxIki]),
          name: String(row[idxName]).trim(),
          dateKey: dateKey(d),
          start: timeToMinutes(row[idxStart]),
          end: timeToMinutes(row[idxEnd]),
          dept: (idxDept !== null) ? String(row[idxDept]).trim() : ''
        });
      }

      // 1件でもあれば採用してループを抜ける
      if (hasTargetMonthData) {
        shiftRows = tempRows;
        Logger.log('✅ 【 ' + sourceName + ' 】から対象月のデータを発見。このデータを使用します。');
        break; 
      } else {
        Logger.log('⚠️ 【 ' + sourceName + ' 】にはデータが0件でした。保険(第二候補)を探します。');
      }
    } catch (e) {
      Logger.log('❌ 【 ' + sourceName + ' 】エラー: ' + e.message + ' -> 保険を探します。');
    }
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