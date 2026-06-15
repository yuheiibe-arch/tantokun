/**
 * ========================================
 * 第5段階：巨大UI起動 ＆ バックグラウンド実行コントローラー（爆速化版）
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

/**
 * ★爆速化：ループの中で毎回データを読み込むのをやめました
 */
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
  
  // ========================================================
  // ★ここで1回だけマスターデータを読み込んで使い回す（超高速化）
  // ========================================================
  var context = null;
  try {
    context = buildContext(year, month);
  } catch (e) {
    ss.toast('マスターデータの読み込みに失敗しました: ' + e.message, 'エラー', 10);
    return;
  }
  
  while (queue.length > 0) {
    // 6分制限が近づいたら次のトリガーへバトンタッチ
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
      ss.toast(currentIndex + '/' + totalCount + ' 件目を生成中...\n残り ' + queue.length + ' 件', '⚙️ 実行中: ' + clinic.name, 10);

      // 生成とPDF化
      var sheet = generateScheduleWithContext(context, clinicNo, year, month);
      var pdfFile = exportSheetToPDF(sheet, year, month, clinic.name);
      
      var folders = pdfFile.getParents();
      var folderUrl = folders.hasNext() ? folders.next().getUrl() : '';

      results.push({
        success: true, clinicNo: clinicNo, clinicName: clinic.name,
        area: clinic.area || 'その他', sheetName: sheet.getName(),
        sheetId: sheet.getSheetId(), folderUrl: folderUrl
      });
    } catch (e) {
      results.push({ success: false, clinicNo: rawClinicNo, error: e.message });
    }
    
    props.setProperty('CLINIC_QUEUE', JSON.stringify(queue));
    props.setProperty('PROCESS_RESULTS', JSON.stringify(results));
    props.setProperty('PROCESSED_COUNT', results.length.toString());
  }
  
  ss.toast('最終処理（目次の作成）を行っています...', '✨ 仕上げ中', 5);
  var urls = updateIndexSheet(results, ss);
  
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

function updateIndexSheet(results, ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = '✨ 目次';
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName, 0);
  } else {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1);
  }
  
  sheet.clear();

  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  sheet.getRange('A1:C1').merge();
  sheet.getRange('A1').setValue('✨ 目次 (最終更新: ' + now + ')')
       .setBackground('#4a86e8').setFontColor('white').setFontWeight('bold').setFontSize(12);

  sheet.getRange('A2:C2').setValues([['エリア', '拠点・シート名', 'リンク']])
       .setBackground('#cccccc').setFontWeight('bold');

  var rows = [];
  var folderUrlToReturn = '';

  for (var i = 0; i < results.length; i++) {
    var r = results[i];
    if (!r) continue;
    
    if (r.folderUrl) folderUrlToReturn = r.folderUrl;
    
    if (!r.success) {
      rows.push([r.area || '-', 'エラー (No.' + r.clinicNo + ')', r.error]);
      continue;
    }
    
    var sheetUrl = ss.getUrl() + '#gid=' + r.sheetId;
    var formula = '=HYPERLINK("' + sheetUrl + '", "開く")';
    rows.push([r.area, r.sheetName, formula]);
  }

  if (rows.length > 0) {
    sheet.getRange(3, 1, rows.length, 3).setValues(rows);
    sheet.getRange(2, 1, rows.length + 1, 3).setBorder(true, true, true, true, true, true);
  }

  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 80);

  return {
    indexUrl: ss.getUrl() + '#gid=' + sheet.getSheetId(),
    folderUrl: folderUrlToReturn
  };
}