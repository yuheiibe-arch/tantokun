/**
 * ========================================
 * 第4段階：ロジック（スケジュールの集計・並べ替え処理）
 * ========================================
 */

function collectScheduleFromContext(context, clinicNo, startKey, endKey) {
  var schedule = {};
  var rows = context.shiftRows;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r.clinicNo !== clinicNo) continue;
    if (r.dateKey < startKey || r.dateKey > endKey) continue;
    var inAM = overlaps(r.start, r.end, SLOT_AM_START, SLOT_AM_END);
    var inPM = overlaps(r.start, r.end, SLOT_PM_START, SLOT_PM_END);
    var doctor = { ikiNo: r.ikiNo, name: r.name, start: r.start, end: r.end, dept: r.dept };
    if (!schedule[r.dateKey]) schedule[r.dateKey] = { am: [], pm: [] };
    if (inAM) schedule[r.dateKey].am.push(doctor);
    if (inPM) schedule[r.dateKey].pm.push(doctor);
  }
  return schedule;
}

function mergeSameDoctors(schedule) {
  Object.keys(schedule).forEach(function(dKey) {
    ['am', 'pm'].forEach(function(slot) {
      var seen = {}, merged = [];
      schedule[dKey][slot].forEach(function(doc) {
        if (!seen[doc.ikiNo]) { seen[doc.ikiNo] = true; merged.push(doc); }
      });
      schedule[dKey][slot] = merged;
    });
  });
}

function sortScheduleByDeptAndTime(schedule, useLabel) {
  Object.keys(schedule).forEach(function(dKey) {
    ['am', 'pm'].forEach(function(slot) {
      schedule[dKey][slot].sort(function(a, b) {
        if (useLabel) {
          var da = deptRank(a.dept), db = deptRank(b.dept);
          if (da !== db) return da - db;
        }
        var sa = (a.start === null) ? 99999 : a.start;
        var sb = (b.start === null) ? 99999 : b.start;
        return sa - sb;
      });
    });
  });
}

function deptRank(dept) {
  if (dept.indexOf('小児') !== -1) return 0;
  if (dept.indexOf('内科') !== -1) return 1;
  return 2;
}