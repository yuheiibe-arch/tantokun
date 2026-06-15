/**
 * ========================================
 * 担当医師予定表 自動生成ツール
 * 第1段階：共通の土台（ユーティリティ）
 * ========================================
 *
 * このファイルには、これ以降のすべての処理が使う共通部品を入れています。
 * - 設定（ツール本体のID）
 * - データセットから諸元を読む関数
 * - 動的にヘッダーを探す関数（★最重要原則）
 * - 土台が正しく動くか確認するテスト関数
 */

// ---- 設定 ----
// ツール本体（担当くん）のスプレッドシートID。ここだけは起点なので定数で持つ。
var TOOL_SPREADSHEET_ID = '1FPwc7s1Iw_QoNh0rQ-bnz-nuCpn-qiXUe_SKn0ozPMQ';
var DATASET_SHEET_NAME = 'データセット';


/**
 * データセットシートを読み、諸元（名称→{id, sheetName}）の対応表を返す。
 * 例: getDataSources()['確定シフト'] = { id: '1LFV...', sheetName: '確定シフト' }
 */
function getDataSources() {
  var ss = SpreadsheetApp.openById(TOOL_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(DATASET_SHEET_NAME);
  if (!sheet) {
    throw new Error('「' + DATASET_SHEET_NAME + '」シートが見つかりません。');
  }

  var values = sheet.getDataRange().getValues();
  // ヘッダーを動的に探す（名称・URL・シート名の列位置）
  var headerMap = buildHeaderMap(values[0]);
  var idxName = requireColumn(headerMap, '名称', DATASET_SHEET_NAME);
  var idxUrl = requireColumn(headerMap, 'URL', DATASET_SHEET_NAME);
  var idxSheet = requireColumn(headerMap, 'シート名', DATASET_SHEET_NAME);

  var sources = {};
  for (var i = 1; i < values.length; i++) {
    var name = values[i][idxName];
    if (!name) continue;
    var url = String(values[i][idxUrl]);
    var id = extractSpreadsheetId(url);
    sources[name] = {
      id: id,
      sheetName: values[i][idxSheet],
      url: url
    };
  }
  return sources;
}


/**
 * スプレッドシートのURLからIDを抽出する。
 * 例: https://docs.google.com/spreadsheets/d/XXXX/edit → XXXX
 */
function extractSpreadsheetId(url) {
  var m = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}


/**
 * ★最重要原則：ヘッダー行（見出しの配列）から「見出し名→列番号(0始まり)」の対応表を作る。
 * 列の位置を固定せず、毎回これで見出し名から列を特定する。
 * 前後の空白は無視し、全角・半角はそのまま扱う。
 */
function buildHeaderMap(headerRow) {
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    var key = String(headerRow[i]).trim();
    if (key !== '' && !(key in map)) {
      map[key] = i; // 同名が複数あれば最初の1つを採用
    }
  }
  return map;
}


/**
 * ヘッダー対応表から見出しの列番号を取得する。
 * 見つからなければ分かりやすいエラーを出す（どのシートの何列かを明示）。
 */
function requireColumn(headerMap, headerName, sheetLabel) {
  if (headerName in headerMap) {
    return headerMap[headerName];
  }
  var available = Object.keys(headerMap).join('", "');
  throw new Error(
    '「' + sheetLabel + '」に見出し「' + headerName + '」が見つかりません。\n' +
    '存在する見出し: "' + available + '"'
  );
}


/**
 * 見出しが「あれば取る、なければnull」という緩い取得（任意項目用）。
 */
function optionalColumn(headerMap, headerName) {
  return (headerName in headerMap) ? headerMap[headerName] : null;
}


/**
 * 諸元名（例:'確定シフト'）を指定して、その対象シートと
 * ヘッダー対応表をまとめて返すヘルパー。
 * 返り値: { sheet, headerMap, values }
 *   values は全データ（ヘッダー行含む）の二次元配列
 */
function openSource(sourceName) {
  var sources = getDataSources();
  var src = sources[sourceName];
  if (!src) {
    throw new Error('データセットに「' + sourceName + '」が登録されていません。');
  }
  if (!src.id) {
    throw new Error('「' + sourceName + '」のURLからIDを取得できませんでした。');
  }

  var ss = SpreadsheetApp.openById(src.id);
  // シート名が空欄の登録（祝日など）は、呼び出し側で別途シートを選ぶ想定。
  var sheet = src.sheetName ? ss.getSheetByName(src.sheetName) : null;

  var result = { spreadsheet: ss, sheet: sheet, source: src };
  if (sheet) {
    var values = sheet.getDataRange().getValues();
    result.values = values;
    result.headerMap = buildHeaderMap(values[0]);
  }
  return result;
}


/**
 * ========================================
 * テスト関数：第1段階の土台が正しく動くか確認する
 * ここを実行して、各データソースが開けてヘッダーが読めるか検証する。
 * ========================================
 */
function test_step1_foundation() {
  Logger.log('===== 第1段階 土台テスト 開始 =====');

  // (1) データセットの諸元を取得できるか
  var sources = getDataSources();
  Logger.log('▼ データセットから取得した諸元一覧:');
  Object.keys(sources).forEach(function(name) {
    var s = sources[name];
    Logger.log('  ・' + name + ' → id=' + s.id + ' / sheet=' + s.sheetName);
  });

  // (2) 主要データソースを開いてヘッダーが読めるか
  var checkList = ['正規表現', '確定シフト', '医師マスタ', '休館日'];
  checkList.forEach(function(name) {
    Logger.log('\n▼ 「' + name + '」を開いてヘッダー確認:');
    try {
      var opened = openSource(name);
      if (!opened.sheet) {
        Logger.log('  （シート名が空欄のため、ここではスキップ）');
        return;
      }
      var headers = Object.keys(opened.headerMap);
      Logger.log('  シート名: ' + opened.sheet.getName());
      Logger.log('  行数: ' + opened.values.length + '（ヘッダー含む）');
      Logger.log('  見出し: ' + headers.join(' | '));
    } catch (e) {
      Logger.log('  ★エラー: ' + e.message);
    }
  });

  // (3) 「医師マスタ」は実シート名が「シート1」。データセットの登録シート名と
  //     実際のシート名が食い違っていないかを確認する。
  Logger.log('\n▼ 医師マスタの実シート名チェック:');
  try {
    var docMaster = getDataSources()['医師マスタ'];
    var ms = SpreadsheetApp.openById(docMaster.id);
    Logger.log('  登録シート名: ' + docMaster.sheetName);
    Logger.log('  実在シート: ' + ms.getSheets().map(function(s){return s.getName();}).join(', '));
  } catch (e) {
    Logger.log('  ★エラー: ' + e.message);
  }

  Logger.log('\n===== 第1段階 土台テスト 完了 =====');
}