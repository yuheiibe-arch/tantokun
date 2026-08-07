/**
 * ========================================
 * 第6段階：PDF自動出力（物理分割・はみ出し解消版）
 * ========================================
 */

// 保存先のベースとなるGoogleドライブフォルダID
var PDF_BASE_FOLDER_ID = '1O5ScGBUVKOvmhjpSrIH_9KbrtkYu_0DB';

function exportSheetToPDF(sheet, year, month, clinicName) {
  var ss = sheet.getParent();

  // 1. フォルダの準備
  var baseFolder = DriveApp.getFolderById(PDF_BASE_FOLDER_ID);
  var folderName = year + '年' + ('0' + month).slice(-2) + '月';
  var targetFolder;

  var folders = baseFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    targetFolder = folders.next();
  } else {
    targetFolder = baseFolder.createFolder(folderName);
  }

  var fileName = year + '年' + ('0' + month).slice(-2) + '月_' + clinicName + '.pdf';

  var existingFiles = targetFolder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  SpreadsheetApp.flush(); 

  // ========================================================
  // ★裏側で一時的なスプレッドシートを作り、2枚のタブに物理切断
  // ========================================================
  var tempSs = SpreadsheetApp.create('Temp_PDF_' + clinicName);
  var tempSsId = tempSs.getId();

  try {
    var page1Sheet = sheet.copyTo(tempSs);
    page1Sheet.showSheet(); // 非表示でコピーされても強制的に「表示」させる
    page1Sheet.setName('Page1');
    
    var page2Sheet = sheet.copyTo(tempSs);
    page2Sheet.showSheet(); // こちらも強制的に「表示」させる
    page2Sheet.setName('Page2');

    // デフォルトの空シートを削除（表示されているシートが確実にある状態で消す）
    var defaultSheet = tempSs.getSheets()[0];
    if (tempSs.getSheets().length > 1) {
      tempSs.deleteSheet(defaultSheet);
    }

    // 分割基準行（36行目以降を2ページ目とする）
    var splitRow = 36;
    
    // Page1（前半）：後半部分以降をすべて削除
    var maxRows1 = page1Sheet.getMaxRows();
    if (maxRows1 >= splitRow) {
      page1Sheet.deleteRows(splitRow, maxRows1 - splitRow + 1);
    }

    // Page2（後半）：前半部分をすべて削除
    page2Sheet.deleteRows(1, splitRow - 1);

    SpreadsheetApp.flush();

    // 一時ファイル全体をPDF化
    // ★修正点：scale=2(幅に合わせる)だと縦にはみ出るため、
    // scale=4(ページに合わせる)に変更して確実に1タブ1ページに収める。
    var url = tempSs.getUrl().replace(/edit$/, '') + 'export?'
      + 'exportFormat=pdf&format=pdf'
      + '&size=A4'
      + '&portrait=true'
      + '&scale=4'             // ★変更: 4 = ページに合わせる（縦にも横にも収める）
      + '&top_margin=0.25'     // 余白：狭い
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

    var blob = response.getBlob().setName(fileName);
    var newFile = targetFolder.createFile(blob);

    return newFile;

  } finally {
    // 処理が終わったら、一時的に作ったスプレッドシートを必ずゴミ箱へ捨てる
    DriveApp.getFileById(tempSsId).setTrashed(true);
  }
}