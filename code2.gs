/**
 * ========================================
 * 第2段階-A：マスター読み込み
 * 拠点名マスターと医師マスタを、照合しやすい対応表に変換する。
 * ========================================
 */

/**
 * 拠点名マスターを読み込み、クリニックNo→拠点情報 の対応表を返す。
 * 返り値の例:
 * getClinicMaster()[14] = {
 * name: '川口', group: '埼玉', area: '関東',
 * openDate: Date, clinicNo: 14, director: '青山舞'
 * }
 * ※ MQC など正規記載が無い行はスキップ。
 */
function getClinicMaster() {
  var opened = openSource('正規表現');
  var values = opened.values;
  var h = opened.headerMap;

  var idxSeiki   = requireColumn(h, '正規記載', '拠点名');
  var idxGroup   = requireColumn(h, '拠点グループ', '拠点名');
  var idxArea    = optionalColumn(h, 'エリア');
  var idxOpen    = requireColumn(h, '開院日', '拠点名');
  var idxClinicNo= requireColumn(h, 'クリニックNo', '拠点名');
  var idxDirector= optionalColumn(h, '院長氏名');
  
  // ★追加：Chatwork通知用のID列を取得
  var idxLeaderId      = optionalColumn(h, '拠点長ID');
  var idxSharedAccount = optionalColumn(h, '共用アカウント');
  var idxChatId        = optionalColumn(h, '全体チャットID');

  var master = {};       // clinicNo -> info
  var listForMenu = [];  // メニュー用の並び（正規記載・グループ）

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var seiki = String(row[idxSeiki]).trim();
    if (!seiki) continue;                 // 正規記載が空ならスキップ
    if (seiki === 'MQC') continue;        // MQCは対象外

    var clinicNo = row[idxClinicNo];
    if (clinicNo === '' || clinicNo === null) continue; // No未設定はスキップ

    var info = {
      clinicNo: clinicNo,
      name: seiki,
      group: String(row[idxGroup]).trim(),
      area: (idxArea !== null) ? String(row[idxArea]).trim() : '',
      openDate: row[idxOpen],
      director: (idxDirector !== null) ? String(row[idxDirector]).trim() : '',
      // ★追加：メンション・通知用のデータを格納
      leaderId: (idxLeaderId !== null) ? String(row[idxLeaderId]).trim() : '',
      sharedAccount: (idxSharedAccount !== null) ? String(row[idxSharedAccount]).trim() : '',
      chatId: (idxChatId !== null) ? String(row[idxChatId]).trim() : ''
    };
    master[clinicNo] = info;
    listForMenu.push(info);
  }

  return { byClinicNo: master, list: listForMenu };
}


/**
 * 医師マスタを読み込み、医籍番号→勤務先・資格情報の対応表を返す。
 */
function getDoctorMaster() {
  var opened = openSource('医師マスタ');
  var values = opened.values;
  var h = opened.headerMap;

  var idxId   = requireColumn(h, '医籍登録番号', '医師マスタ');
  var idxWork = requireColumn(h, '現在の主な勤務先名', '医師マスタ');
  
  // ★ 追加：資格情報の列（見つからなければnullになる安全設計）
  var idxZekka  = optionalColumn(h, '舌下免疫療法E-ラーニング受講');
  var idxOnline = optionalColumn(h, 'オンライン診療研修受講');

  var map = {};
  for (var i = 1; i < values.length; i++) {
    var id = values[i][idxId];
    if (id === '' || id === null) continue;
    
    var work = String(values[i][idxWork]).trim();
    
    // 「済」ならOK、それ以外（未・空欄）はNG
    var zekkaStatus = (idxZekka !== null && String(values[i][idxZekka]).trim() === '済') ? 'OK' : 'NG';
    var onlineStatus = (idxOnline !== null && String(values[i][idxOnline]).trim() === '済') ? 'OK' : 'NG';

    // これまでは work の文字列だけだったものを、オブジェクト（複数データ）に変更
    map[normalizeId(id)] = {
      work: work,
      zekka: zekkaStatus,
      online: onlineStatus
    };
  }
  return map;
}


/**
 * 医籍番号を照合用に正規化する。
 * 数値・文字列・前後空白の違いを吸収して、確実に突き合わせられるようにする。
 */
function normalizeId(id) {
  return String(id).trim();
}


/**
 * ========================================
 * テスト関数：第2段階-A マスター読み込みの確認
 * ========================================
 */
function test_step2a_masters() {
  Logger.log('===== 第2段階-A マスター読み込みテスト 開始 =====');

  // (1) 拠点名マスター
  var clinic = getClinicMaster();
  var clinicKeys = Object.keys(clinic.byClinicNo);
  Logger.log('▼ 拠点名マスター: ' + clinicKeys.length + '拠点（MQC除く）');
  // 川口（No=14想定）が引けるか確認
  if (clinic.byClinicNo[14]) {
    var k = clinic.byClinicNo[14];
    Logger.log('  例) No14 → ' + k.name + ' / グループ=' + k.group +
               ' / エリア=' + k.area + ' / 院長=' + k.director);
  } else {
    Logger.log('  ★No14（川口想定）が見つかりません。クリニックNoの値を確認してください。');
  }
  // 拠点グループの種類を一覧（メニュー分類用）
  var groups = {};
  clinic.list.forEach(function(c){ if (c.group) groups[c.group] = (groups[c.group]||0)+1; });
  Logger.log('  拠点グループ一覧: ');
  Object.keys(groups).forEach(function(g){ Logger.log('    ・' + g + ' … ' + groups[g] + '拠点'); });

  // (2) 医師マスタ
  var doc = getDoctorMaster();
  var docKeys = Object.keys(doc);
  Logger.log('\n▼ 医師マスタ: ' + docKeys.length + '件の医籍番号→勤務先');
  // 確定シフトに出てきた医籍番号でいくつか試し引き
  var sampleIds = ['532756', '458267', '441829'];
  sampleIds.forEach(function(id){
    var docInfo = doc[normalizeId(id)];
    if(docInfo){
       Logger.log('  医籍番号 ' + id + ' → 勤務先: ' + docInfo.work + ' / 舌下: ' + docInfo.zekka + ' / オンライン: ' + docInfo.online);
    } else {
       Logger.log('  医籍番号 ' + id + ' → （データなし）');
    }
  });

  Logger.log('\n===== 第2段階-A マスター読み込みテスト 完了 =====');
}