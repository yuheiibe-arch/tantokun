/**
 * 自動監視デーモン（メイン処理：深夜帯に実行）
 */
function daemon_checkAndSyncSchedules() {
  var today = new Date();
  var currentYear = today.getFullYear();
  var currentMonth = today.getMonth() + 1;
  var lastDayOfThisMonth = new Date(currentYear, currentMonth, 0).getDate();
  var targets = [];

  if (today.getDate() >= lastDayOfThisMonth - 1) {
    var nextYear = currentMonth === 12 ? currentYear + 1 : currentYear;
    var nextMonth = currentMonth === 12 ? 1 : currentMonth + 1;
    targets.push({ year: nextYear, month: nextMonth, isNewMonth: true });
  } else {
    targets.push({ year: currentYear, month: currentMonth, isNewMonth: false });
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getDocumentProperties();
  var needIndexUpdate = false;
  
  targets.forEach(function(target) {
    var year = target.year;
    var month = target.month;
    var context;
    try { context = buildContext(year, month); } catch (e) { return; }
    
    context.clinicMaster.list.forEach(function(clinic) {
      var targetVal = year * 12 + month;
      if (clinic.openDate && Object.prototype.toString.call(clinic.openDate) === '[object Date]') {
        var openVal = clinic.openDate.getFullYear() * 12 + (clinic.openDate.getMonth() + 1);
        if (targetVal < openVal) return; // 未開院スキップ
      }
      
      var clinicNo = clinic.clinicNo;
      var lastDay = new Date(year, month, 0).getDate();
      var startKey = dateKey(new Date(year, month - 1, 1));
      var endKey   = dateKey(new Date(year, month - 1, lastDay));
      
      var currentSchedule = collectScheduleFromContext(context, clinicNo, startKey, endKey);
      var scheduleStr = JSON.stringify(currentSchedule);
      var currentHash = computeMD5_(scheduleStr);
      var propKey = 'DAEMON_HASH_' + year + '_' + month + '_' + clinicNo;
      var lastHash = props.getProperty(propKey);
      var expectedSheetName = ('0' + month).slice(-2) + clinic.name;
      var targetSheet = ss.getSheetByName(expectedSheetName);
      var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
      var oldSchedule = getSavedScheduleState_(clinicNo, year, month);
      
      if (currentHash !== lastHash || !targetSheet) {
        Logger.log('【変更検知 🔄】' + clinic.name);
        try {
          var sheet = generateScheduleWithContext(context, clinicNo, year, month);
          
          sheet.getRange('A1').clearNote();
          SpreadsheetApp.flush();
          
          var pdfFile = exportSheetToPDF(sheet, year, month, clinic.name);
          if (!pdfFile) throw new Error('exportSheetToPDF からファイルが返却されませんでした。');
          
          var sheetUrl = ss.getUrl() + '#gid=' + sheet.getSheetId();
          var isFirstTime = (!lastHash); 
          var filledVacancies = isFirstTime ? [] : getFilledVacancies_(oldSchedule, currentSchedule);
          
          // ★ 送信するのではなく、朝9時送信用に「予約」する
          if (isFirstTime && target.isNewMonth) {
            enqueueChatworkNotification_(clinic, month, 'monthly', [], pdfFile, sheetUrl);
          } else if (filledVacancies.length > 0) {
            enqueueChatworkNotification_(clinic, month, 'filled', filledVacancies, pdfFile, sheetUrl);
          }
          
          props.setProperty('LAST_UPDATE_' + sheet.getName(), nowStr);
          props.setProperty(propKey, currentHash);
          saveScheduleState_(clinicNo, year, month, currentSchedule);
          sheet.getRange('A1').setNote('🔄 最終変更検知・更新: ' + nowStr + '\n(データハッシュ: ' + currentHash + ')');
          needIndexUpdate = true;
          
        } catch (err) {
          Logger.log('【エラー ❌】' + clinic.name + ': ' + err.message);
        }
        Utilities.sleep(1500); 
      }
    });
  });
  if (needIndexUpdate) rebuildIndexSheet_standalone(ss);
}

/**
 * 自動監視用の目次再構築関数
 */
function rebuildIndexSheet_standalone(ss) {
  var sheetName = '✨ 目次';
  var indexSheet = ss.getSheetByName(sheetName);
  
  if (!indexSheet) {
    indexSheet = ss.insertSheet(sheetName, 0);
  } else {
    ss.setActiveSheet(indexSheet);
    ss.moveActiveSheet(1);
  }
  
  indexSheet.clear();

  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  indexSheet.getRange('A1:D1').merge();
  indexSheet.getRange('A1').setValue('✨ 担当くん 総合目次 (自動監視更新: ' + now + ')')
       .setBackground('#4a86e8').setFontColor('white').setFontWeight('bold').setFontSize(12);
  indexSheet.getRange('A2:D2').setValues([['エリア', '拠点・シート名', '最終更新', 'リンク']])
       .setBackground('#cccccc').setFontWeight('bold');

  var clinicMaster = getClinicMaster();
  var props = PropertiesService.getDocumentProperties();
  var sheets = ss.getSheets();
  var rows = [];

  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName();
    var m = sName.match(/^(\d{2})(.+)$/);
    if (!m) continue; 
    
    var clinicName = m[2];
    var area = 'その他エリア';
    
    for (var j = 0; j < clinicMaster.list.length; j++) {
      if (clinicMaster.list[j].name === clinicName) {
        area = clinicMaster.list[j].area || 'その他エリア';
        break;
      }
    }
    
    var lastUpdate = props.getProperty('LAST_UPDATE_' + sName) || '新規生成';
    var sheetUrl = ss.getUrl() + '#gid=' + sheets[i].getSheetId();
    var formula = '=HYPERLINK("' + sheetUrl + '", "開く")';
    
    rows.push({
      area: area,
      sheetName: sName,
      lastUpdate: lastUpdate,
      formula: formula
    });
  }

  rows.sort(function(a, b) {
    if (a.area !== b.area) return a.area.localeCompare(b.area);
    return a.sheetName.localeCompare(b.sheetName);
  });

  var valueRows = [];
  rows.forEach(function(row) {
    valueRows.push([row.area, row.sheetName, row.lastUpdate, row.formula]);
  });

  if (valueRows.length > 0) {
    indexSheet.getRange(3, 1, valueRows.length, 4).setValues(valueRows);
    indexSheet.getRange(2, 1, valueRows.length + 1, 4).setBorder(true, true, true, true, true, true);
  }

  indexSheet.setColumnWidth(1, 120);
  indexSheet.setColumnWidth(2, 200);
  indexSheet.setColumnWidth(3, 160);
  indexSheet.setColumnWidth(4, 80);
}