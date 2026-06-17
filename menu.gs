/**
 * ========================================
 * 第5段階：巨大UI起動 ＆ バックグラウンド実行コントローラー（UIスキップ・履歴リンク・目次統一版）
 * ========================================
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('担当くん')
    .addItem('🎮 担当くんを起動', 'openAppUI')
    .addSeparator()
    .addItem('🛑 バックグラウンド処理を強制停止', 'stopBackgroundProcess')
    .addToUi();
}

function openAppUI() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(800)
    .setHeight(680)
    .setTitle('担当くん - シフト自動生成システム');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getUIData() {
  var master = getClinicMaster();
  var list = master.list;

  var hierarchy = {};
  list.forEach(function(c) {
    var area = c.area || 'その他エリア';
    var group = c.group || 'その他グループ';
    if (!hierarchy[area]) hierarchy[area] = {};
    if (!hierarchy[area][group]) hierarchy[area][group] = [];
    
    var openTime = null;
    if (c.openDate && Object.prototype.toString.call(c.openDate) === '[object Date]') {
      openTime = { y: c.openDate.getFullYear(), m: c.openDate.getMonth() + 1 };
    }

    hierarchy[area][group].push({ name: c.name, clinicNo: c.clinicNo, open: openTime });
  });

  var today = new Date();
  var currentMonth = today.getMonth() + 1;
  var targetMonth = currentMonth === 12 ? 1 : currentMonth + 1;
  var targetYear = today.getFullYear();
  if (currentMonth === 12) targetYear++;

  return { hierarchy: hierarchy, defaultMonth: targetMonth, defaultYear: targetYear };
}

function setupBackgroundProcess(payload) {
  var props = PropertiesService.getDocumentProperties();
  props.setProperty('TARGET_YEAR', payload.targetYear);
  props.setProperty('TARGET_MONTH', payload.targetMonth);
  
  var queue = payload.clinicNos;
  if (queue.length === 1 && queue[0] === 'ALL') {
    queue = getValidClinicNos(payload.targetYear, payload.targetMonth);
  }
  
  props.setProperty('CLINIC_QUEUE', JSON.stringify(queue));
  props.setProperty('PROCESS_RESULTS', JSON.stringify([]));
  props.setProperty('MAIN_SS_ID', SpreadsheetApp.getActiveSpreadsheet().getId());
  
  props.setProperty('TOTAL_COUNT', queue.length.toString());
  props.setProperty('PROCESSED_COUNT', '0');
  props.setProperty('CURRENT_CLINIC_NAME', '準備中...');
  props.setProperty('IS_COMPLETED', 'false');
  props.setProperty('COMPLETION_DATA', '');
  
  clearTriggers();
  ScriptApp.newTrigger('processBatch').timeBased().after(1000).create();
  return true;
}

function getProgress() {
  var props = PropertiesService.getDocumentProperties();
  return {
    isCompleted: props.getProperty('IS_COMPLETED') === 'true',
    total: parseInt(props.getProperty('TOTAL_COUNT') || '1', 10),
    processed: parseInt(props.getProperty('PROCESSED_COUNT') || '0', 10),
    currentName: props.getProperty('CURRENT_CLINIC_NAME') || '',
    completionData: props.getProperty('COMPLETION_DATA') || ''
  };
}

function processBatch() {
  var startTime = Date.now();
  var props = PropertiesService.getDocumentProperties();
  var queueStr = props.getProperty('CLINIC_QUEUE');
  
  if (!queueStr) return;
  
  var queue = JSON.parse(queueStr);
  var results = JSON.parse(props.getProperty('PROCESS_RESULTS') || '[]');
  var year = parseInt(props.getProperty('TARGET_YEAR'), 10);
  var month = parseInt(props.getProperty('TARGET_MONTH'), 10);
  var ssId = props.getProperty('MAIN_SS_ID');
  
  var ss = SpreadsheetApp.openById(ssId);
  SpreadsheetApp.setActiveSpreadsheet(ss);
  
  clearTriggers();
  var totalCount = parseInt(props.getProperty('TOTAL_COUNT'), 10);
  
  var context = null;
  try {
    context = buildContext(year, month);
  } catch (e) {
    ss.toast('マスターデータの読み込みに失敗しました: ' + e.message, 'エラー', 10);
    return;
  }
  
  while (queue.length > 0) {
    if (Date.now() - startTime > 270000) {
      props.setProperty('CLINIC_QUEUE', JSON.stringify(queue));
      props.setProperty('PROCESS_RESULTS', JSON.stringify(results));
      ScriptApp.newTrigger('processBatch').timeBased().after(1000).create();
      ss.toast('GASの実行制限(6分)が近づいたため、自動で再起動して続きから再開します...', '⏳ 一時中断・リレー待機', 10);
      return;
    }
    
    var rawClinicNo = queue.shift();
    var clinicNo = isNaN(Number(rawClinicNo)) ? rawClinicNo : Number(rawClinicNo);
    var currentIndex = results.length + 1;
    
    try {
      var clinic = context.clinicMaster.byClinicNo[clinicNo];
      if (!clinic) throw new Error('拠点マスターに見つかりません。');
      
      props.setProperty('CURRENT_CLINIC_NAME', clinic.name);

      // --- ★ UI実行時も変更チェック（ハッシュ比較）を行う ---
      var lastDay = new Date(year, month, 0).getDate();
      var startKey = dateKey(new Date(year, month - 1, 1));
      var endKey   = dateKey(new Date(year, month - 1, lastDay));
      var currentSchedule = collectScheduleFromContext(context, clinicNo, startKey, endKey);
      var currentHash = computeMD5_(JSON.stringify(currentSchedule));
      
      var propKey = 'DAEMON_HASH_' + year + '_' + month + '_' + clinicNo;
      var lastHash = props.getProperty(propKey);
      
      var expectedSheetName = ('0' + month).slice(-2) + clinic.name;
      var targetSheet = ss.getSheetByName(expectedSheetName);

      // 変更がなく、かつシートがすでに存在する場合はスキップ
      if (currentHash === lastHash && targetSheet) {
        ss.toast(currentIndex + '/' + totalCount + ' 件目: 変更なし・スキップ\n残り ' + queue.length + ' 件', '⏭️ スキップ: ' + clinic.name, 3);
        
        results.push({
          success: true, clinicNo: clinicNo, clinicName: clinic.name,
          area: clinic.area || 'その他', sheetName: expectedSheetName,
          sheetId: targetSheet.getSheetId(), folderUrl: props.getProperty('LAST_FOLDER_URL') || ''
        });
        
      } else {
        // --- 変更があった場合、またはシートが存在しない場合のみ生成 ---
        ss.toast(currentIndex + '/' + totalCount + ' 件目を生成中...\n残り ' + queue.length + ' 件', '⚙️ 実行中: ' + clinic.name, 10);

        var sheet = generateScheduleWithContext(context, clinicNo, year, month);
        sheet.showSheet(); // 原本が非表示の場合でも強制的に表示させてリンクを使えるようにする
        
        var pdfFile = exportSheetToPDF(sheet, year, month, clinic.name);
        
        var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
        props.setProperty('LAST_UPDATE_' + sheet.getName(), nowStr); // 更新日時の保存
        props.setProperty(propKey, currentHash); // 次回比較用のハッシュを保存
        
        var folders = pdfFile.getParents();
        var folderUrl = folders.hasNext() ? folders.next().getUrl() : '';
        props.setProperty('LAST_FOLDER_URL', folderUrl);

        results.push({
          success: true, clinicNo: clinicNo, clinicName: clinic.name,
          area: clinic.area || 'その他', sheetName: sheet.getName(),
          sheetId: sheet.getSheetId(), folderUrl: folderUrl
        });
      }
    } catch (e) {
      results.push({ success: false, clinicNo: rawClinicNo, error: e.message });
    }
    
    props.setProperty('CLINIC_QUEUE', JSON.stringify(queue));
    props.setProperty('PROCESS_RESULTS', JSON.stringify(results));
    props.setProperty('PROCESSED_COUNT', results.length.toString());
  }
  
  ss.toast('最終処理（目次の作成）を行っています...', '✨ 仕上げ中', 5);
  // 手動実行時も、全自動スキャン型の4列目次ジェネレーターを呼び出す
  var urls = updateIndexSheet(null, ss);
  
  props.setProperty('IS_COMPLETED', 'true');
  props.setProperty('COMPLETION_DATA', JSON.stringify(urls));
  
  props.deleteProperty('CLINIC_QUEUE');
  props.deleteProperty('PROCESS_RESULTS');
  props.deleteProperty('MAIN_SS_ID');
  
  ss.toast('すべての予定表とPDFの生成が完了しました！\n✨ 目次シートをご確認ください。', '✅ 完全完了', -1);
}

function getValidClinicNos(targetYear, targetMonth) {
  var master = getClinicMaster();
  var targetVal = targetYear * 12 + targetMonth;
  var validNos = [];
  
  master.list.forEach(function(c) {
    if (c.openDate && Object.prototype.toString.call(c.openDate) === '[object Date]') {
      var openVal = c.openDate.getFullYear() * 12 + (c.openDate.getMonth() + 1);
      if (targetVal < openVal) return; 
    }
    validNos.push(c.clinicNo);
  });
  return validNos;
}

function clearTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processBatch') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function stopBackgroundProcess() {
  clearTriggers();
  PropertiesService.getDocumentProperties().deleteProperty('CLINIC_QUEUE');
  PropertiesService.getDocumentProperties().setProperty('IS_COMPLETED', 'true');
  SpreadsheetApp.getActiveSpreadsheet().toast('バックグラウンド処理を強制停止し、キューをクリアしました。', '🛑 停止', 10);
}

/**
 * ★目次生成ロジックの統一化（自動監視側と完全に同じ4列スキャン方式）
 */
function updateIndexSheet(unused_results, ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
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
  indexSheet.getRange('A1').setValue('✨ 担当くん 総合目次 (UI手動生成: ' + now + ')')
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
    
    var lastUpdate = props.getProperty('LAST_UPDATE_' + sName) || '-';
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

  return {
    indexUrl: ss.getUrl() + '#gid=' + indexSheet.getSheetId(),
    folderUrl: props.getProperty('LAST_FOLDER_URL') || ''
  };
}

/**
 * 手動UIからハッシュ計算を行うための共通関数
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