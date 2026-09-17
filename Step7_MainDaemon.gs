/**
 * ========================================
 * 🌟 本番用：夜間自動監視 ＆ 送信予約デーモン（リレー方式・完全決定版）
 * ========================================
 */
function daemon_checkAndSyncSchedules() {
  Logger.log('====== 🌟 本番デーモン (夜間監視＆予約) 起動 ======');
  var today = new Date();
  var currentYear = today.getFullYear();
  var currentMonth = today.getMonth() + 1;
  var lastDayOfThisMonth = new Date(currentYear, currentMonth, 0).getDate();
  var targets = [];

  // ①【今月分】の監視
  targets.push({ year: currentYear, month: currentMonth, isNewMonth: false });

  // ②【来月分】の監視（今日が月末1日前以降なら作動）
  if (today.getDate() >= lastDayOfThisMonth - 1) {
    var nextYear = currentMonth === 12 ? currentYear + 1 : currentYear;
    var nextMonth = currentMonth === 12 ? 1 : currentMonth + 1;
    targets.push({ year: nextYear, month: nextMonth, isNewMonth: true });
  }

  // 拠点一覧を取得（マスター取得用に一度だけ呼び出す）
  var context;
  try { 
    context = buildContext(currentYear, currentMonth); 
  } catch (e) { 
    Logger.log('❌ マスター取得エラー: ' + e.message);
    return; 
  }
  
  // キュー（処理待ちリスト）を作成
  var queue = [];
  targets.forEach(function(target) {
    context.clinicMaster.list.forEach(function(clinic) {
      queue.push({
        clinicNo: clinic.clinicNo,
        clinicName: clinic.name,
        year: target.year,
        month: target.month
      });
    });
  });

  var props = PropertiesService.getDocumentProperties();
  props.setProperty('DAEMON_QUEUE', JSON.stringify(queue));
  props.setProperty('DAEMON_INDEX_UPDATE_NEEDED', 'false');

  clearDaemonTriggers_();
  
  // 直ちにキューの処理を開始
  daemon_processQueue();
}

/**
 * リレー方式でキューを処理する実行関数
 */
function daemon_processQueue() {
  var startTime = Date.now();
  var props = PropertiesService.getDocumentProperties();
  var queueStr = props.getProperty('DAEMON_QUEUE');
  if (!queueStr) return;
  var queue = JSON.parse(queueStr);
  
  if (queue.length === 0) {
    Logger.log("====== 🌟 本番デーモン 全件完了 ======");
    if (props.getProperty('DAEMON_INDEX_UPDATE_NEEDED') === 'true') {
      updateIndexSheet(null, SpreadsheetApp.getActiveSpreadsheet());
      props.setProperty('DAEMON_INDEX_UPDATE_NEEDED', 'false');
    }
    props.deleteProperty('DAEMON_QUEUE');
    clearDaemonTriggers_();
    return;
  }

  clearDaemonTriggers_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date();
  var nowStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  
  var currentContextYearMonth = "";
  var context = null;

  while(queue.length > 0) {
    // ★ 実行時間が約4.5分 (270000ミリ秒) を超えたらリレー予約して終了
    if (Date.now() - startTime > 270000) {
      Logger.log("⏳ タイムアウト回避: 残りのキュー(" + queue.length + "件)を次回の実行に持ち越します。");
      props.setProperty('DAEMON_QUEUE', JSON.stringify(queue));
      ScriptApp.newTrigger('daemon_processQueue').timeBased().after(1000).create();
      return;
    }

    var task = queue.shift();
    
    // 対象月のcontextがまだ無い場合のみ生成（高速化）
    var targetYM = task.year + "_" + task.month;
    if (currentContextYearMonth !== targetYM) {
      try {
        context = buildContext(task.year, task.month);
        currentContextYearMonth = targetYM;
      } catch (e) {
        Logger.log('❌ データ読み込み失敗: ' + e.message);
        continue;
      }
    }

    var clinic = context.clinicMaster.byClinicNo[task.clinicNo];
    if (!clinic) continue;
    
    var targetVal = task.year * 12 + task.month;
    if (clinic.openDate && Object.prototype.toString.call(clinic.openDate) === '[object Date]') {
      var openVal = clinic.openDate.getFullYear() * 12 + (clinic.openDate.getMonth() + 1);
      if (targetVal < openVal) continue; 
    }
    
    var clinicNo = task.clinicNo;
    var TARGET_YEAR = task.year;
    var TARGET_MONTH = task.month;
    
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
    var oldSchedule = getSavedScheduleState_(clinicNo, TARGET_YEAR, TARGET_MONTH);
    var isChanged = (currentHash !== lastHash || !targetSheet);
    
    if (isChanged) {
        var isFirstTime = (!lastHash); 
        var diffs = isFirstTime ? { filled: [], vacated: [] } : checkScheduleDifferences_Main_(oldSchedule, currentSchedule);
        
        var shouldSendChat = true;
        if (!isFirstTime && diffs.filled.length === 0 && diffs.vacated.length === 0) {
          shouldSendChat = false; 
        }

        if (!shouldSendChat) {
          Logger.log('⏩ ' + clinic.name + ' : 細かな変更のため通知スキップ（裏で最新化）');
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
              var type = isFirstTime ? 'monthly' : 'filled';
              enqueueChatworkNotification_(clinic, TARGET_MONTH, type, diffs, pdfFile, sheetUrl);
              
              props.setProperty('LAST_UPDATE_' + sheet.getName(), nowStr);
              props.setProperty(propKey, currentHash);
              saveScheduleState_(clinicNo, TARGET_YEAR, TARGET_MONTH, currentSchedule);
              sheet.getRange('A1').setNote('🔄 最終更新: ' + nowStr + '\n(Hash: ' + currentHash + ')');
              props.setProperty('DAEMON_INDEX_UPDATE_NEEDED', 'true');
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
  }

  // キューが空になったら再帰呼び出しして完了処理へ
  props.setProperty('DAEMON_QUEUE', JSON.stringify(queue));
  daemon_processQueue();
}

function clearDaemonTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'daemon_processQueue') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * 【本番用】差分検知関数（バイネーム＆重複排除対応版）
 */
function checkScheduleDifferences_Main_(oldSched, newSched) {
  var diff = { filled: [], vacated: [] };
  if (!oldSched) return diff;

  // 重複を排除してソートされた文字列を返すヘルパー
  function getUniqueNamesStr(docs) {
    if (!docs || docs.length === 0) return "";
    var names = docs.map(function(d){ return d.name ? d.name.replace(/\s/g, '') : ''; })
                    .filter(function(n){ return n !== ''; });
    var unique = [];
    for (var i = 0; i < names.length; i++) {
      if (unique.indexOf(names[i]) === -1) unique.push(names[i]);
    }
    return unique.sort().join(', ');
  }

  Object.keys(newSched).forEach(function(dKey) {
    var cleanKey = dKey.replace(/-/g, ''); 
    var dateStr = parseInt(cleanKey.substring(4, 6), 10) + "月" + parseInt(cleanKey.substring(6, 8), 10) + "日";

    ['am', 'pm'].forEach(function(slot) {
      var oldDocs = (oldSched[dKey] && oldSched[dKey][slot]) ? oldSched[dKey][slot] : [];
      var newDocs = newSched[dKey][slot] || [];

      var oldNames = getUniqueNamesStr(oldDocs);
      var newNames = getUniqueNamesStr(newDocs);
      var slotName = (slot === 'am') ? '午前' : '午後(夜間含む)';

      if (oldNames !== newNames) {
        if (oldNames === "" && newNames !== "") {
          diff.filled.push(dateStr + "の" + slotName + " (" + newNames + " が追加)");
        } else if (oldNames !== "" && newNames === "") {
          diff.vacated.push(dateStr + "の" + slotName + " (" + oldNames + " が削除)");
        } else {
          diff.filled.push(dateStr + "の" + slotName + " (変更: " + oldNames + " → " + newNames + ")");
        }
      }
    });
  });
  return diff;
}