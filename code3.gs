/**
 * ========================================
 * 第2段階-B：コアロジック
 * 指定した「拠点(クリニックNo) × 期間」のシフトを集め、
 * 日付×午前/午後の枠に医師を配置したデータ構造を作る。
 * ========================================
 */

// 時間帯の定義（分単位。9:00 = 540分）
var SLOT_AM_START = 9 * 60;    // 09:00
var SLOT_AM_END   = 13 * 60;   // 13:00
var SLOT_PM_START = 15 * 60;   // 15:00
var SLOT_PM_END   = 21 * 60;   // 21:00（北葛西は20:00だが重なり判定には影響しない）

/**
 * 確定シフトの時刻セル（1899年の日付つきDate）から「その日の何分か」を取り出す。
 * 例: Sat Dec 30 1899 09:00:00 → 540
 */
function timeToMinutes(cellValue) {
  if (cellValue === '' || cellValue === null) return null;
  if (Object.prototype.toString.call(cellValue) === '[object Date]') {
    return cellValue.getHours() * 60 + cellValue.getMinutes();
  }
  return null;
}

/**
 * 勤務時間[start,end]が、ある枠[slotStart,slotEnd]に重なるか判定。
 */
function overlaps(start, end, slotStart, slotEnd) {
  if (start === null || end === null) return false;
  return start < slotEnd && end > slotStart;
}

/**
 * 日付を 'YYYY-MM-DD' の文字列キーにする（時刻を無視して日単位で扱う）。
 */
function dateKey(d) {
  if (Object.prototype.toString.call(d) !== '[object Date]') return null;
  var y = d.getFullYear();
  var m = ('0' + (d.getMonth() + 1)).slice(-2);
  var day = ('0' + d.getDate()).slice(-2);
  return y + '-' + m + '-' + day;
}

/**
 * 指定拠点・期間のシフトを集めて、日付ごと・午前/午後ごとの医師リストを作る。
 *
 * @param {number} clinicNo 対象拠点のクリニックNo
 * @param {Date} startDate 期間開始（その日を含む）
 * @param {Date} endDate 期間終了（その日を含む）
 * @return {Object} schedule
 *   schedule['2026-06-01'] = {
 *     am: [ {ikiNo:'...', name:'山口いぶき', start:540, end:780}, ... ],
 *     pm: [ ... ]
 *   }
 */
function collectSchedule(clinicNo, startDate, endDate) {
  var opened = openSource('確定シフト');
  var values = opened.values;
  var h = opened.headerMap;

  var idxClinicNo = requireColumn(h, 'クリニックNo', '確定シフト');
  var idxIki      = requireColumn(h, '医籍番号', '確定シフト');
  var idxName     = requireColumn(h, '名前', '確定シフト');
  var idxDate     = requireColumn(h, '勤務日', '確定シフト');
  var idxStart    = requireColumn(h, '勤務開始時間', '確定シフト');
  var idxEnd      = requireColumn(h, '勤務終了時間', '確定シフト');

  // 期間を日付キーの範囲に
  var startKey = dateKey(startDate);
  var endKey = dateKey(endDate);

  var schedule = {};

  for (var i = 1; i < values.length; i++) {
    var row = values[i];

    // (1) 拠点が一致するか（クリニックNoで照合）
    if (row[idxClinicNo] !== clinicNo) continue;

    // (2) 勤務日が期間内か
    var d = row[idxDate];
    var dKey = dateKey(d);
    if (!dKey || dKey < startKey || dKey > endKey) continue;

    // (3) 勤務時間を分に変換
    var startMin = timeToMinutes(row[idxStart]);
    var endMin = timeToMinutes(row[idxEnd]);

    // (4) 午前・午後どちらの枠に入るか（重なり判定）
    var inAM = overlaps(startMin, endMin, SLOT_AM_START, SLOT_AM_END);
    var inPM = overlaps(startMin, endMin, SLOT_PM_START, SLOT_PM_END);

    var doctor = {
      ikiNo: normalizeId(row[idxIki]),
      name: String(row[idxName]).trim(),
      start: startMin,
      end: endMin
    };

    if (!schedule[dKey]) schedule[dKey] = { am: [], pm: [] };
    if (inAM) schedule[dKey].am.push(doctor);
    if (inPM) schedule[dKey].pm.push(doctor);
  }

  return schedule;
}


/**
 * ========================================
 * テスト関数：第2段階-B コアロジックの確認
 * 川口（No14）の2026年6月前半（6/1〜6/15）を集めてログ表示。
 * 完成見本の川口シートと見比べて、正しく集計できているか確認する。
 * ========================================
 */
function test_step2b_core() {
  Logger.log('===== 第2段階-B コアロジックテスト 開始 =====');

  var clinicNo = 14; // 川口
  var start = new Date(2026, 5, 1);   // 2026-06-01（月は0始まりなので5=6月）
  var end   = new Date(2026, 5, 15);  // 2026-06-15

  var schedule = collectSchedule(clinicNo, start, end);
  var doctorMaster = getDoctorMaster();

  // 日付順に並べて表示
  var keys = Object.keys(schedule).sort();
  Logger.log('対象: 川口(No' + clinicNo + ') / ' + keys.length + '日分のデータ');
  Logger.log('----------------------------------------');

  keys.forEach(function(dKey) {
    var slot = schedule[dKey];
    Logger.log('■ ' + dKey);
    Logger.log('  午前: ' + formatSlotForLog(slot.am, doctorMaster));
    Logger.log('  午後: ' + formatSlotForLog(slot.pm, doctorMaster));
  });

  Logger.log('\n===== 第2段階-B コアロジックテスト 完了 =====');
}

// ログ表示用：枠内の医師を「名前(所属先)」で連結
function formatSlotForLog(doctors, doctorMaster) {
  if (!doctors || doctors.length === 0) return '（なし）';
  return doctors.map(function(doc) {
    var work = doctorMaster[doc.ikiNo] || '（所属先なし）';
    return doc.name + '(' + work + ')';
  }).join(' / ');
}