/**
 * ========================================
 * 第8段階：受講依頼フォーム（UI・別スプシ台帳・Slack Bot連携）
 * ========================================
 */

function openTrainingRequestDialog() {
  var html = HtmlService.createHtmlOutputFromFile('RequestForm')
      .setWidth(620)
      .setHeight(430)
      .setTitle('舌下・オンライン 受講依頼');
  SpreadsheetApp.getUi().showModalDialog(html, '受講依頼フォーム');
}

function getNgDoctorsForForm(clinicNo, year, month) {
  try {
    var existingRequests = {}; 
    try {
      var ss = SpreadsheetApp.openById(QUAL_SPREADSHEET_ID);
      var sheet = ss.getSheetByName('受講依頼台帳');
      if (sheet) {
        var data = sheet.getDataRange().getValues();
        var targetY = parseInt(year, 10);
        var targetM = parseInt(month, 10);
        var targetMonthStr = targetY + "年" + targetM + "月"; 

        for (var i = 1; i < data.length; i++) {
          var cellVal = data[i][1]; 
          var cellStr = "";
          
          if (Object.prototype.toString.call(cellVal) === '[object Date]') {
            cellStr = cellVal.getFullYear() + "年" + (cellVal.getMonth() + 1) + "月";
          } else {
            var m = String(cellVal).match(/(\d{4})年(\d{1,2})月/);
            if (m) {
              cellStr = parseInt(m[1], 10) + "年" + parseInt(m[2], 10) + "月";
            } else {
              cellStr = String(cellVal).trim();
            }
          }

          if (cellStr === targetMonthStr) {
            var reqDate = data[i][0]; 
            if (Object.prototype.toString.call(reqDate) === '[object Date]') {
              reqDate = Utilities.formatDate(reqDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
            } else {
              reqDate = String(reqDate).split(' ')[0]; 
            }
            var reqClinic = String(data[i][2]).trim(); 
            var docName = String(data[i][3]).trim();   
            var reqStatus = String(data[i][5]).trim() || '未対応'; 
            
            var normDocName = docName.replace(/\s+/g, '');
            existingRequests[normDocName] = { clinic: reqClinic, date: reqDate, status: reqStatus };
          }
        }
      }
    } catch(e) {
      Logger.log('台帳読み込みエラー（無視して続行）: ' + e.message);
    }

    var context = buildContext(parseInt(year, 10), parseInt(month, 10));
    var ngDocs = {};

    context.shiftRows.forEach(function(r) {
      if (r.clinicNo === parseInt(clinicNo, 10)) {
        var docData = context.doctorMaster[r.ikiNo] || { zekka: 'NG', online: 'NG' };
        
        if (docData.zekka !== 'OK' || docData.online !== 'OK') {
          var normName = String(r.name).replace(/\s+/g, ''); 
          
          if (!ngDocs[r.ikiNo]) {
            ngDocs[r.ikiNo] = {
              ikiNo: r.ikiNo, // ★ここで医籍番号をセット
              name: r.name,
              zekka: docData.zekka === 'OK' ? 'OK' : 'NG',
              online: docData.online === 'OK' ? 'OK' : 'NG',
              requestedInfo: existingRequests[normName] || null,
              dates: []
            };
          }
          ngDocs[r.ikiNo].dates.push(r.dateKey);
        }
      }
    });

    var list = [];
    var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    for (var k in ngDocs) {
      var dates = ngDocs[k].dates.sort();
      var closest = dates[0]; 
      for (var j = 0; j < dates.length; j++) {
        if (dates[j] >= todayStr) {
          closest = dates[j]; 
          break;
        }
      }
      
      var parts = closest.split('-');
      ngDocs[k].closestDateStr = parseInt(parts[0], 10) + "年" + parseInt(parts[1], 10) + "月" + parseInt(parts[2], 10) + "日";
      list.push(ngDocs[k]);
    }
    return list;
  } catch(e) {
    Logger.log('NG医師取得エラー: ' + e.message);
    return [];
  }
}

function submitTrainingRequest(formData) {
  try {
    var ss = SpreadsheetApp.openById(QUAL_SPREADSHEET_ID);
    var sheetName = '受講依頼台帳';
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['タイムスタンプ', '対象月', '依頼元拠点', '対象医師名', '依頼内容', 'ステータス', '備考', '医籍番号']);
      sheet.getRange("A1:H1").setBackground('#4a86e8').setFontColor('white').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    var targetMonthStr = "'" + formData.year + "年" + formData.month + "月";

    // ★ここでH列にformData.ikiNoを書き込む
    sheet.appendRow([
      timestamp,
      targetMonthStr,
      formData.clinicName,
      formData.doctorName,
      formData.trainingTypes,
      '未対応',
      formData.notes,
      formData.ikiNo
    ]);

    var lastRow = sheet.getLastRow();
    var newRowRange = sheet.getRange(lastRow, 1, 1, 8);
    newRowRange.setHorizontalAlignment('left');

    var statusCell = sheet.getRange(lastRow, 6);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['未対応', '対応中', '完了'], true)
      .setAllowInvalid(true) 
      .build();
    statusCell.setDataValidation(rule);

    SpreadsheetApp.flush();

    var slackToken = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
    if (!slackToken) {
       return { success: true, message: '台帳に記録しました。（Slack通知スキップ）' };
    }

    var displayDateStr = formData.closestDate || (formData.year + "年" + formData.month + "月");

    var slackMessage = "*【新規】受講依頼のお知らせ*\n" +
                       "担当くんより、新しい受講依頼が届きました。\n\n" +
                       "*対象勤務日:* " + displayDateStr + "\n" +
                       "*依頼元拠点:* " + formData.clinicName + "\n" +
                       "*対象医師名:* " + formData.doctorName + "\n" +
                       "*依頼内容:* " + formData.trainingTypes + "\n" +
                       "*備考:* " + (formData.notes || "なし") + "\n\n" +
                       "<" + ss.getUrl() + "#gid=" + sheet.getSheetId() + "|👉 台帳スプレッドシートを確認する>";

    var payload = {
      "channel": "C09TRHGU64D",
      "text": slackMessage
    };

    var options = {
      "method": "post",
      "headers": {
        "Authorization": "Bearer " + slackToken,
        "Content-Type": "application/json"
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    var response = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", options);
    var resJson = JSON.parse(response.getContentText());
    
    if (!resJson.ok) {
       return { success: true, message: '台帳に記録しました。（Slack通知失敗: ' + resJson.error + '）' };
    }

    return { success: true };

  } catch (e) {
    Logger.log('受講依頼送信エラー: ' + e.message);
    return { success: false, message: e.message };
  }
}