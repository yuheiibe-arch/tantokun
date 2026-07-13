/**
 * ========================================
 * 🌟 本番用：夜間自動監視 ＆ 送信予約デーモン（完全決定版）
 * ========================================
 * （日付バグ修正 ＆ スパム通知防止 ＆ 欠員・追加検知 搭載）
 * 夜中に変更をスキャンし、必要なものだけをキュー（待合室）に予約します。
 */
function daemon_checkAndSyncSchedules() {
  var today = new Date();
  var currentYear = today.getFullYear();
  var currentMonth = today.getMonth() + 1;
  var lastDayOfThisMonth = new Date(currentYear, currentMonth, 0).getDate();
  var targets = [];

  // ①【今月分】の監視（日常の欠員・補充の検知用）
  targets.push({ year: currentYear, month: currentMonth, isNewMonth: false });

  // ②【来月分】の監視（月末一斉送信用：今日が月末1日前以降なら作動）
  if (today.getDate() >= lastDayOfThisMonth - 1) {
    var nextYear = currentMonth === 12 ? currentYear + 1 : currentYear;
    var nextMonth = currentMonth === 12 ? 1 : currentMonth + 1;
    targets.push({ year: nextYear, month: nextMonth, isNewMonth: true });
  }

  Logger.log('====== 🌟 本番デーモン (夜間監視＆予約) 起動 ======');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getDocumentProperties();
  var needIndexUpdate = false;

  targets.forEach(function(target) {
    var TARGET_YEAR = target.year;
    var TARGET_MONTH = target.month;
    
    Logger.log('▶ ' + TARGET_YEAR + '年' + TARGET_MONTH + '月分の処理を開始...');
    
    var context;
    try { 
      context = buildContext(TARGET_YEAR, TARGET_MONTH); 
    } catch (e) { 
      Logger.log('❌ データ読み込み失敗: ' + e.message);
      return; 
    }
    
    context.clinicMaster.list.forEach(function(clinic) {
      // 開院判定
      var targetVal = TARGET_YEAR * 12 + TARGET_MONTH;
      if (clinic.openDate && Object.prototype.toString.call(clinic.openDate) === '[object Date]') {
        var openVal = clinic.openDate.getFullYear() * 12 + (clinic.openDate.getMonth() + 1);
        if (targetVal < openVal) return; 
      }
      
      var clinicNo = clinic.clinicNo;
      var lastDay = new Date(TARGET_YEAR, TARGET_MONTH, 0).getDate();
      var startKey = dateKey(new Date(TARGET_YEAR, TARGET_MONTH - 1, 1));
      var endKey   = dateKey(new Date(TARGET_YEAR, TARGET_MONTH - 1, lastDay));
      
      var currentSchedule = collectScheduleFromContext(context, clinicNo, startKey, endKey);
      var scheduleStr = JSON.stringify(currentSchedule);
      var currentHash = computeMD5_(scheduleStr);
      
      var propKey = 'DAEMON_HASH_' + TARGET_YEAR + '_' + TARGET_MONTH + '_' + clinicNo;
      var lastHash = props.getProperty(propKey);
      
      var expectedSheetName = ('0' + TARGET_MONTH).slice(-2) + clinic.name;
      var targetSheet = ss.getSheetByName(expectedSheetName);
      var nowStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
      
      var oldSchedule = getSavedScheduleState_(clinicNo, TARGET_YEAR, TARGET_MONTH);
      var isChanged = (currentHash !== lastHash || !targetSheet);
      
      if (isChanged) {
        var isFirstTime = (!lastHash); 
        var diffs = isFirstTime ? { filled: [], vacated: [] } : checkScheduleDifferences_Main_(oldSchedule, currentSchedule);
        
        // ★ 厳格なスパム通知防止ロジック
        var shouldSendChat = true;
        if (!isFirstTime && diffs.filled.length === 0 && diffs.vacated.length === 0) {
          shouldSendChat = false; 
        }

        if (!shouldSendChat) {
          Logger.log('⏩ ' + clinic.name + ' : 細かな変更のため通知スキップ（データは裏で最新化します）');
          props.setProperty(propKey, currentHash);
          saveScheduleState_(clinicNo, TARGET_YEAR, TARGET_MONTH, currentSchedule);
        } else {
          Logger.log('🔥 【処理開始】' + clinic.name + ' (' + TARGET_MONTH + '月分) - 送信予約あり');
          
          var maxRetries = 2;
          for (var attempt = 1; attempt <= maxRetries; attempt++) {
            try {
              var sheet = generateScheduleWithContext(context, clinicNo, TARGET_YEAR, TARGET_MONTH);
              SpreadsheetApp.flush();
              Utilities.sleep(2000); 
              
              sheet.getRange('A1').clearNote();
              SpreadsheetApp.flush();
              
              var pdfFile = exportSheetToPDF(sheet, TARGET_YEAR, TARGET_MONTH, clinic.name);
              if (!pdfFile) throw new Error('PDF生成失敗');
              
              var sheetUrl = ss.getUrl() + '#gid=' + sheet.getSheetId();
              
              // ★ 直に送るのではなく、朝9時用のキュー（待合室）に予約する
              var type = isFirstTime ? 'monthly' : 'filled';
              enqueueChatworkNotification_(clinic, TARGET_MONTH, type, diffs, pdfFile, sheetUrl);
              
              props.setProperty('LAST_UPDATE_' + sheet.getName(), nowStr);
              props.setProperty(propKey, currentHash);
              saveScheduleState_(clinicNo, TARGET_YEAR, TARGET_MONTH, currentSchedule);
              sheet.getRange('A1').setNote('🔄 最終変更検知・更新: ' + nowStr + '\n(データハッシュ: ' + currentHash + ')');
              needIndexUpdate = true;
              
              break; 
            } catch (err) {
              if (attempt < maxRetries) {
                Logger.log('⚠️ 【リトライ】エラー再試行... (' + err.message + ')');
                Utilities.sleep(2000);
              } else {
                Logger.log('❌ 【エラー】' + clinic.name + ': ' + err.message);
              }
            }
          }
        }
        Utilities.sleep(1500); 
      }
    });
  });
  
  if (needIndexUpdate) rebuildIndexSheet_standalone(ss);
  Logger.log('====== 🌟 本番デーモン 終了 ======');
}

/**
 * 【本番用】差分検知関数（「0月」日付バグ修正済み）
 */
function checkScheduleDifferences_Main_(oldSched, newSched) {
  var diff = { filled: [], vacated: [] };
  if (!oldSched) return diff;
  
  Object.keys(newSched).forEach(function(dKey) {
    var cleanKey = dKey.replace(/-/g, ''); 
    var dateStr = parseInt(cleanKey.substring(4, 6), 10) + "月" + parseInt(cleanKey.substring(6, 8), 10) + "日";
    
    ['am', 'pm'].forEach(function(slot) {
      var oldDocs = (oldSched[dKey] && oldSched[dKey][slot]) ? oldSched[dKey][slot] : [];
      var newDocs = newSched[dKey][slot] || [];

      var oldHasDoc = oldDocs.some(function(d) { return d.name && d.name.replace(/\s/g, '') !== ""; });
      var newHasDoc = newDocs.some(function(d) { return d.name && d.name.replace(/\s/g, '') !== ""; });

      var slotName = (slot === 'am') ? '午前' : '午後(夜間含む)';

      if (!oldHasDoc && newHasDoc) {
        diff.filled.push(dateStr + "の" + slotName);
      } else if (oldHasDoc && !newHasDoc) {
        diff.vacated.push(dateStr + "の" + slotName);
      }
    });
  });
  return diff;
}