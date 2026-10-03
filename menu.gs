/**
 * ========================================
 * 第5段階：巨大UI起動 ＆ バックグラウンド実行コントローラー（本体目次に別表リンク追加版）
 * ========================================
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('担当くん')
    .addItem('🎮 担当くんを起動', 'openAppUI')
    .addSeparator()
    .addItem('🛑 バックグラウンド処理を強制停止', 'stopBackgroundProcess')
    .addSeparator()
    .addItem('📝 所属先変更依頼', 'showFormDialog')
    .addItem('🎓 舌下・オンライン 受講依頼', 'openTrainingRequestDialog') // ★ 今回追加した受講依頼フォーム呼び出し
    .addToUi();
}

/**
 * フォームを開くためのダイアログを表示する関数（所属先変更依頼用）
 */
function showFormDialog() {
  var formUrl = 'https://docs.google.com/forms/d/e/1FAIpQLSe9yFWpRopjnyx46W4CwFGXe99K_KSH1CN0kPr4KtpvGtKILg/viewform?usp=dialog';
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family: sans-serif; padding: 10px;">' +
    '<p>以下のボタンをクリックしてフォームを開いてください。</p>' +
    '<a href="' + formUrl + '" target="_blank" style="display:inline-block; padding:10px 20px; background:#4a86e8; color:#fff; text-decoration:none; border-radius:5px;">所属先変更依頼フォームを開く</a>' +
    '</div>'
  )
  .setWidth(350)
  .setHeight(150);
  
  SpreadsheetApp.getUi().showModalDialog(html, '所属先変更依頼');
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

  // ★ PDF保存用フォルダの準備
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

      var lastDay = new Date(year, month, 0).getDate();
      var startKey = dateKey(new Date(year, month - 1, 1));
      var endKey   = dateKey(new Date(year, month - 1, lastDay));
      var currentSchedule = collectScheduleFromContext(context, clinicNo, startKey, endKey);
      var currentHash = computeMD5_(JSON.stringify(currentSchedule));
      
      var propKey = 'DAEMON_HASH_' + year + '_' + month + '_' + clinicNo;
      var lastHash = props.getProperty(propKey);
      
      var expectedSheetName = ('0' + month).slice(-2) + clinic.name;
      var targetSheet = ss.getSheetByName(expectedSheetName);

      if (currentHash === lastHash && targetSheet) {
        ss.toast(currentIndex + '/' + totalCount + ' 件目: 変更なし・スキップ\n残り ' + queue.length + ' 件', '⏭️ スキップ: ' + clinic.name, 3);
        
        results.push({
          success: true, clinicNo: clinicNo, clinicName: clinic.name,
          area: clinic.area || 'その他', sheetName: expectedSheetName,
          sheetId: targetSheet.getSheetId(), folderUrl: props.getProperty('LAST_FOLDER_URL') || ''
        });
        
      } else {
        ss.toast(currentIndex + '/' + totalCount + ' 件目を生成中...\n残り ' + queue.length + ' 件', '⚙️ 実行中: ' + clinic.name, 10);

        // ★ シフト表と別表の両方を生成・保存させる
        var resultObj = generateScheduleWithContext(context, clinicNo, year, month, targetFolder);
        var sheet = resultObj.sheet;
        sheet.showSheet(); 
        
        var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
        props.setProperty('LAST_UPDATE_' + sheet.getName(), nowStr); 
        props.setProperty(propKey, currentHash); 
        
        var folderUrl = targetFolder.getUrl();
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
 * ★ 目次生成ロジック（別表のリンクも引っ張ってくる完全版）
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
  indexSheet.getRange('A1:E1').merge();
  indexSheet.getRange('A1').setValue('✨ 担当くん 総合目次 (UI/デーモン自動生成: ' + now + ')')
        .setBackground('#4a86e8').setFontColor('white').setFontWeight('bold').setFontSize(12);
  
  // ★ E列に「別表(資格表)」を追加
  indexSheet.getRange('A2:E2').setValues([['エリア', '拠点・シート名', '最終更新', 'シフト表', '別表(資格表)']])
        .setBackground('#cccccc').setFontWeight('bold');

  var clinicMaster = getClinicMaster();
  var props = PropertiesService.getDocumentProperties();
  var sheets = ss.getSheets();
  var rows = [];

  // ★ 外部スプシ（別表）のシート一覧を事前取得
  var qualSs = null;
  var qualSheetMap = {};
  try {
    qualSs = SpreadsheetApp.openById(QUAL_SPREADSHEET_ID);
    var qSheets = qualSs.getSheets();
    for (var k = 0; k < qSheets.length; k++) {
      qualSheetMap[qSheets[k].getName()] = qualSs.getUrl() + '#gid=' + qSheets[k].getSheetId();
    }
  } catch (e) {
    Logger.log('別表スプシの読み込みエラー: ' + e.message);
  }

  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName();
    var m = sName.match(/^(\d{2})(.+)$/);
    if (!m) continue; 
    
    var monthStr = m[1];
    var clinicName = m[2];
    var area = 'その他エリア';
    
    for (var j = 0; j < clinicMaster.list.length; j++) {
      if (clinicMaster.list[j].name === clinicName) {
        area = clinicMaster.list[j].area || 'その他エリア';
        break;
      }
    }
    
    var lastUpdate = props.getProperty('LAST_UPDATE_' + sName) || '-';
    var shiftUrl = ss.getUrl() + '#gid=' + sheets[i].getSheetId();
    var shiftFormula = '=HYPERLINK("' + shiftUrl + '", "開く")';
    
    // ★ 別表のURLを動的に探し出して関数化する
    var qualLinks = [];
    var qName1 = monthStr + clinicName + '_別表';
    var qName2 = monthStr + clinicName + '（小児科）別表';
    var qName3 = monthStr + clinicName + '（内科）別表';

    if (qualSheetMap[qName1]) {
      qualLinks.push('HYPERLINK("' + qualSheetMap[qName1] + '", "別表")');
    } else {
      // 2つに分かれている場合（北葛西・亀有）
      if (qualSheetMap[qName2]) qualLinks.push('HYPERLINK("' + qualSheetMap[qName2] + '", "小児科")');
      if (qualSheetMap[qName3]) qualLinks.push('HYPERLINK("' + qualSheetMap[qName3] + '", "内科")');
    }

    var qualFormula = '';
    if (qualLinks.length === 1) {
      qualFormula = '=' + qualLinks[0];
    } else if (qualLinks.length === 2) {
      // & で繋いで「小児科 / 内科」の見た目にする
      qualFormula = '=' + qualLinks[0] + ' & " / " & ' + qualLinks[1];
    }
    
    rows.push({
      area: area,
      sheetName: sName,
      lastUpdate: lastUpdate,
      shiftFormula: shiftFormula,
      qualFormula: qualFormula
    });
  }

  rows.sort(function(a, b) {
    if (a.area !== b.area) return a.area.localeCompare(b.area);
    return a.sheetName.localeCompare(b.sheetName);
  });

  var valueRows = [];
  rows.forEach(function(row) {
    // 5列分の配列を作成
    valueRows.push([row.area, row.sheetName, row.lastUpdate, row.shiftFormula, row.qualFormula]);
  });

  if (valueRows.length > 0) {
    indexSheet.getRange(3, 1, valueRows.length, 5).setValues(valueRows);
    indexSheet.getRange(2, 1, valueRows.length + 1, 5).setBorder(true, true, true, true, true, true);
  }

  indexSheet.setColumnWidth(1, 120);
  indexSheet.setColumnWidth(2, 200);
  indexSheet.setColumnWidth(3, 160);
  indexSheet.setColumnWidth(4, 80);
  indexSheet.setColumnWidth(5, 120); // 別表列の幅

  return {
    indexUrl: ss.getUrl() + '#gid=' + indexSheet.getSheetId(),
    folderUrl: props.getProperty('LAST_FOLDER_URL') || ''
  };
}

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