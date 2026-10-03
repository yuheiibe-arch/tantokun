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
 */
function buildContext(year, month) {
  var today = new Date();
  var currentYear = today.getFullYear();
  var currentMonth = today.getMonth() + 1;
  
  var isCurrentMonth = (year === currentYear && month === currentMonth);
  var primarySource = isCurrentMonth ? '当月シフト' : '確定シフト';
  var fallbackSource = isCurrentMonth ? '確定シフト' : '当月シフト';
  
  var attemptSources = [primarySource, fallbackSource];
  var shiftRows = [];

  for (var s = 0; s < attemptSources.length; s++) {
    var sourceName = attemptSources[s];
    Logger.log('📂 試行: データソース 【 ' + sourceName + ' 】');
    
    try {
      var opened = openSource(sourceName);
      var values = opened.values;
      var h = opened.headerMap;
      
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
 * ★改修：シフト表に加え、別表(資格表)の生成と、目次へのリンク追記機能を追加
 */
function generateScheduleWithContext(context, clinicNo, year, month, targetFolder) {
  var ss = SpreadsheetApp.openById(TOOL_SPREADSHEET_ID);
  var clinic = context.clinicMaster.byClinicNo[clinicNo];
  if (!clinic) throw new Error('クリニックNo ' + clinicNo + ' が拠点マスターにありません。');

  var lastDay = new Date(year, month, 0).getDate();
  var startKey = dateKey(new Date(year, month - 1, 1));
  var endKey   = dateKey(new Date(year, month - 1, lastDay));

  // --- 1. メインのシフト表を作成 ---
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
  SpreadsheetApp.flush();

  // シフト表のシートURL
  var shiftSheetUrl = ss.getUrl() + '#gid=' + sheet.getSheetId();
  
  // シフト表のPDFを作成 (第6段階の機能を利用)
  var shiftPdf = null;
  if (targetFolder) {
     shiftPdf = exportSheetToPDF(sheet, year, month, clinic.name, targetFolder);
  }

  // --- 2. 別表（資格表）を作成 ---
  var qualResult = null;
  try {
     qualResult = generateQualTable(context, clinicNo, year, month, targetFolder);
  } catch(e) {
     Logger.log('⚠️ ' + clinic.name + ' の別表生成に失敗しました: ' + e.message);
  }

  // --- 3. 目次（リンク集）へ書き込み ---
  updateIndexLinks(year, month, clinic.name, shiftSheetUrl, qualResult ? qualResult.sheetUrl : '');

  // ★修正：返り値に「別表のURL (qualSheetUrl)」を含めるように変更
  return {
     sheet: sheet,
     shiftPdf: shiftPdf,
     shiftSheetUrl: shiftSheetUrl,
     qualSheetUrl: qualResult ? qualResult.sheetUrl : null
  };
}

/**
 * 目次スプシにシフト表と別表のリンクを追記する関数
 */
function updateIndexLinks(year, month, clinicName, shiftUrl, qualUrl) {
  try {
    var qualSs = SpreadsheetApp.openById(QUAL_SPREADSHEET_ID);
    var indexSheet = qualSs.getSheetByName('✨ 目次');
    if (!indexSheet) return;

    var data = indexSheet.getDataRange().getValues();
    var header = data[0];
    
    // ヘッダーから列番号を探す
    var colClinic = -1, colShift = -1, colQual = -1;
    for (var c = 0; c < header.length; c++) {
      var h = String(header[c]).trim();
      if (h === '拠点名') colClinic = c;
      if (h.indexOf('シフト') !== -1) colShift = c;
      if (h.indexOf('別表') !== -1) colQual = c;
    }
    
    // 拠点名列が見つからなければ処理しない
    if (colClinic === -1) return;

    // 該当する拠点の行を探してリンクを書き込む
    var updated = false;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][colClinic]).trim() === clinicName) {
        var rowNum = r + 1;
        if (colShift !== -1 && shiftUrl) {
          indexSheet.getRange(rowNum, colShift + 1).setFormula('=HYPERLINK("' + shiftUrl + '", "シフト表")');
        }
        if (colQual !== -1 && qualUrl) {
          indexSheet.getRange(rowNum, colQual + 1).setFormula('=HYPERLINK("' + qualUrl + '", "別表")');
        }
        updated = true;
        break;
      }
    }
    
    // 新規拠点などで行が見つからなかった場合は最終行に追記
    if (!updated) {
      var newRow = [];
      for (var c = 0; c < header.length; c++) newRow.push('');
      newRow[colClinic] = clinicName;
      indexSheet.appendRow(newRow);
      var rowNum = indexSheet.getLastRow();
      
      if (colShift !== -1 && shiftUrl) indexSheet.getRange(rowNum, colShift + 1).setFormula('=HYPERLINK("' + shiftUrl + '", "シフト表")');
      if (colQual !== -1 && qualUrl) indexSheet.getRange(rowNum, colQual + 1).setFormula('=HYPERLINK("' + qualUrl + '", "別表")');
    }
    
  } catch(e) {
    Logger.log('目次の更新エラー: ' + e.message);
  }
}

/**
 * 1拠点を生成（context無し版。単発生成用）。
 */
function generateSchedule(clinicNo, year, month) {
  var context = buildContext(year, month);
  var result = generateScheduleWithContext(context, clinicNo, year, month, null);
  SpreadsheetApp.flush();
  return result.sheet;
}

/**
 * 全拠点一括生成。
 */
function generateAllClinics(year, month) {
  var t0 = new Date().getTime();
  var context = buildContext(year, month);
  
  // 保存先フォルダの準備 (PDF処理と同じロジック)
  var PDF_BASE_FOLDER_ID = '1O5ScGBUVKOvmhjpSrIH_9KbrtkYu_0DB';
  var baseFolder = DriveApp.getFolderById(PDF_BASE_FOLDER_ID);
  var folderName = year + '年' + ('0' + month).slice(-2) + '月';
  var targetFolder;
  var folders = baseFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    targetFolder = folders.next();
  } else {
    targetFolder = baseFolder.createFolder(folderName);
  }

  var n = 0;
  context.clinicMaster.list.forEach(function(clinic) {
    generateScheduleWithContext(context, clinic.clinicNo, year, month, targetFolder);
    n++;
  });
  SpreadsheetApp.flush();
  var sec = Math.round((new Date().getTime() - t0) / 1000);
  Logger.log('一括生成完了: ' + n + '拠点 / ' + year + '年' + month + '月 / ' + sec + '秒');
}

function isDeptLabelClinic(clinicName) {
  return DEPT_LABEL_CLINICS.indexOf(clinicName) !== -1;
}