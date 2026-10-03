/**
 * 臨時デバッグ用：2026年10月の「別表（資格表）」を全拠点分一括で作成する（超高速版）
 */
function debug_GenerateAllQualTables_202610() {
  var year = 2026;
  var month = 10;
  
  Logger.log('===== 全拠点 別表一括作成（' + year + '年' + month + '月） 開始 =====');
  
  try {
    var context = buildContext(year, month);
    var clinicList = context.clinicMaster.list;
    
    var successCount = 0;
    var errorCount = 0;

    for (var i = 0; i < clinicList.length; i++) {
      var clinicNo = clinicList[i].clinicNo;
      var clinicName = clinicList[i].name;
      
      try {
        Logger.log((i + 1) + '/' + clinicList.length + ' : ' + clinicName + ' の別表を作成中...');
        generateQualTable(context, clinicNo, year, month, null);
        successCount++;
      } catch (e) {
        Logger.log('⚠️ ' + clinicName + ' の別表生成エラー: ' + e.message);
        errorCount++;
      }
    }
    
    SpreadsheetApp.flush();
    updateIndexLinks_ForQualOnly_(year, month);
    
    Logger.log('===== 全拠点 別表一括作成 完了 =====');
    Logger.log('成功: ' + successCount + ' 拠点');
    Logger.log('失敗: ' + errorCount + ' 拠点');

  } catch (e) {
    Logger.log('❌ コンテキスト読み込みエラー: ' + e.message);
  }
}

/**
 * 今回の臨時スクリプト用の目次更新ヘルパー
 */
function updateIndexLinks_ForQualOnly_(year, month) {
  try {
    var qualSs = SpreadsheetApp.openById(QUAL_SPREADSHEET_ID); 
    var indexSheet = qualSs.getSheetByName('✨ 目次');
    if (!indexSheet) return;

    var sheets = qualSs.getSheets();
    var data = indexSheet.getDataRange().getValues();
    var header = data[0];
    
    var colClinic = -1, colQual = -1;
    for (var c = 0; c < header.length; c++) {
      var h = String(header[c]).trim();
      if (h === '拠点名') colClinic = c;
      if (h.indexOf('別表') !== -1) colQual = c;
    }
    if (colClinic === -1 || colQual === -1) return;

    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      var m = sheetName.match(/^(\d{2})(.+)$/);
      if (m && parseInt(m[1], 10) === month) {
        var clinicName = m[2].replace(/（.*）別表$/, '').replace(/_別表$/, '');
        var url = qualSs.getUrl() + '#gid=' + sheets[i].getSheetId();
        
        for (var r = 1; r < data.length; r++) {
          if (String(data[r][colClinic]).trim() === clinicName) {
            indexSheet.getRange(r + 1, colQual + 1).setFormula('=HYPERLINK("' + url + '", "別表")');
            break;
          }
        }
      }
    }
    Logger.log('目次シートのリンクを更新しました。');
  } catch(e) {
    Logger.log('目次更新エラー: ' + e.message);
  }
}