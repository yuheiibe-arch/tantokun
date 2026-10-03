/**
 * ========================================
 * 第6段階：PDF自動出力（物理分割・はみ出し解消版）
 * ========================================
 */

// ※ targetFolderは呼び出し元(3_Main.gs)から渡されます
function exportSheetToPDF(sheet, year, month, clinicName, targetFolder) {
  var fileName = year + '年' + ('0' + month).slice(-2) + '月_' + clinicName + '.pdf';

  // 既存の同名ファイルがあればゴミ箱へ
  if (targetFolder) {
    var existingFiles = targetFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }
  }

  SpreadsheetApp.flush(); 

  // ========================================================
  // ★裏側で一時的なスプレッドシートを作り、2枚のタブに物理切断
  // ========================================================
  var tempSs = SpreadsheetApp.create('Temp_PDF_' + clinicName);
  var tempSsId = tempSs.getId();

  try {
    var page1Sheet = sheet.copyTo(tempSs);
    page1Sheet.showSheet(); 
    page1Sheet.setName('Page1');
    
    var page2Sheet = sheet.copyTo(tempSs);
    page2Sheet.showSheet(); 
    page2Sheet.setName('Page2');

    var defaultSheet = tempSs.getSheets()[0];
    if (tempSs.getSheets().length > 1) {
      tempSs.deleteSheet(defaultSheet);
    }

    var splitRow = 36;
    
    var maxRows1 = page1Sheet.getMaxRows();
    if (maxRows1 >= splitRow) {
      page1Sheet.deleteRows(splitRow, maxRows1 - splitRow + 1);
    }

    page2Sheet.deleteRows(1, splitRow - 1);

    SpreadsheetApp.flush();

    var url = tempSs.getUrl().replace(/edit$/, '') + 'export?'
      + 'exportFormat=pdf&format=pdf'
      + '&size=A4'
      + '&portrait=true'
      + '&scale=4'
      + '&top_margin=0.25'
      + '&bottom_margin=0.25'
      + '&left_margin=0.25'
      + '&right_margin=0.25'
      + '&sheetnames=false'
      + '&printtitle=false'
      + '&pagenumbers=false'
      + '&gridlines=false'
      + '&fzr=false';

    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('PDF出力に失敗しました。');
    }

    if (targetFolder) {
      var blob = response.getBlob().setName(fileName);
      return targetFolder.createFile(blob);
    }
    return null;

  } finally {
    DriveApp.getFileById(tempSsId).setTrashed(true);
  }
}