/**
 * ========================================
 * 第8段階：受講依頼フォーム（UI・別スプシ台帳・Slack連携）
 * ========================================
 */

/**
 * カスタムメニューから呼ばれる：フォーム（ダイアログ）を開く
 */
function openTrainingRequestDialog() {
  var html = HtmlService.createHtmlOutputFromFile('RequestForm')
      .setWidth(450)
      .setHeight(480)
      .setTitle('舌下・オンライン 受講依頼');
  SpreadsheetApp.getUi().showModalDialog(html, '受講依頼フォーム');
}

/**
 * HTML側から呼ばれる：送信データを受け取り処理する
 */
function submitTrainingRequest(formData) {
  try {
    // 1. 外部スプレッドシート（別表スプシ）を開く
    var ss = SpreadsheetApp.openById(QUAL_SPREADSHEET_ID);
    var sheetName = '受講依頼台帳';
    var sheet = ss.getSheetByName(sheetName);

    // シートがなければ自動で作成し、ヘッダーをセットする
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['タイムスタンプ', '依頼元拠点', '対象医師名', '依頼内容', 'ステータス', '備考']);
      sheet.getRange("A1:F1").setBackground('#4a86e8').setFontColor('white').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');

    // 2. 台帳にデータを書き込む
    sheet.appendRow([
      timestamp,
      formData.clinicName,
      formData.doctorName,
      formData.trainingType,
      '未対応',  // デフォルトのステータス
      formData.notes
    ]);

    SpreadsheetApp.flush();

    // 3. Slackへ通知を送る
    // ※ スクリプトプロパティに 'SLACK_WEBHOOK_URL' を設定しておく必要があります
    var slackWebhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
    if (!slackWebhookUrl) {
       return { success: true, message: '台帳に記録しました。（※Slack通知URLが未設定のため通知はスキップされました）' };
    }

    var slackMessage = "*【新規】受講依頼のお知らせ*\n" +
                       "担当くんメニューより、新しい受講依頼が届きました。\n\n" +
                       "*依頼元拠点:* " + formData.clinicName + "\n" +
                       "*対象医師名:* " + formData.doctorName + "\n" +
                       "*依頼内容:* " + formData.trainingType + "\n" +
                       "*備考:* " + (formData.notes || "なし") + "\n\n" +
                       "<" + ss.getUrl() + "#gid=" + sheet.getSheetId() + "|👉 台帳スプレッドシートを確認する>";

    var payload = {
      "text": slackMessage
    };

    var options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    UrlFetchApp.fetch(slackWebhookUrl, options);

    return { success: true };

  } catch (e) {
    Logger.log('受講依頼送信エラー: ' + e.message);
    return { success: false, message: e.message };
  }
}