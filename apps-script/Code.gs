/**
 * 노무 리서치 포털 — Google Apps Script 중계 서버
 * ------------------------------------------------------------
 * 역할 3가지:
 *   1) 홈페이지 폼 제출 접수  (doPost, action 없음)
 *   2) 실행 머신이 새 질의를 가져감  (doGet ?action=pending&key=SECRET)
 *   3) 실행 머신이 상태/답변URL 갱신  (doPost action=update)
 *
 * [설치]
 *   1. 구글 시트를 새로 만들고, 확장 프로그램 > Apps Script 로 이 코드를 붙여넣기
 *   2. 아래 SECRET 을 아무도 모를 값으로 바꾸기 (실행 머신과 공유)
 *   3. 배포 > 새 배포 > 유형:웹앱 > 실행:나 / 액세스:모든 사용자
 *   4. 나오는 /exec URL 을 config.js 의 APPS_SCRIPT_URL 에 붙여넣기
 *   5. 같은 URL + &key=SECRET 을 실행 머신 런북에 등록
 */

var SECRET = "여기를_긴_임의문자열로_바꾸세요";   // 실행 머신과 공유하는 비밀키
var SHEET_NAME = "queue";

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["id","ts","date","name","dept","depth","question","status","answer_url","processed_at","worker"]);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 홈페이지 폼 제출 & 실행 머신 상태 갱신 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || "{}");

    // (A) 실행 머신이 상태 갱신
    if (body.action === "update") {
      if (body.key !== SECRET) return json_({ ok:false, error:"unauthorized" });
      return updateRow_(body);
    }

    // (B) 홈페이지 신규 접수
    var q = (body.question || "").toString().trim();
    var name = (body.name || "").toString().trim();
    if (!q || !name) return json_({ ok:false, error:"이름/질문 누락" });

    var sh = sheet_();
    var now = new Date();
    var id = "NR-" + Utilities.formatDate(now, "Asia/Seoul", "yyMMdd") + "-" +
             ("000" + (sh.getLastRow())).slice(-3);
    var depth = ["full","standard","quick"].indexOf(body.depth) >= 0 ? body.depth : "standard";

    sh.appendRow([
      id, now.getTime(), Utilities.formatDate(now,"Asia/Seoul","yyyy-MM-dd HH:mm"),
      name, (body.dept||"").toString().trim(), depth, q,
      "recv", "", "", ""
    ]);
    return json_({ ok:true, id:id });

  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

/** 실행 머신이 대기열/전체를 조회 */
function doGet(e) {
  var p = e.parameter || {};
  if (p.key !== SECRET) return json_({ ok:false, error:"unauthorized" });

  var sh = sheet_();
  var vals = sh.getDataRange().getValues();
  var head = vals.shift();
  var items = vals.map(function(r){
    var o = {}; head.forEach(function(h,i){ o[h]=r[i]; }); return o;
  });

  if (p.action === "pending") {
    // 처리할 항목: 아직 recv 상태인 것 (오래된 순)
    items = items.filter(function(o){ return o.status === "recv"; });
  }
  return json_({ ok:true, items:items });
}

function updateRow_(body) {
  var sh = sheet_();
  var vals = sh.getDataRange().getValues();
  for (var i=1; i<vals.length; i++) {
    if (vals[i][0] === body.id) {
      if (body.status)      sh.getRange(i+1, 8).setValue(body.status);       // status
      if (body.answer_url)  sh.getRange(i+1, 9).setValue(body.answer_url);   // answer_url
      sh.getRange(i+1, 10).setValue(Utilities.formatDate(new Date(),"Asia/Seoul","yyyy-MM-dd HH:mm")); // processed_at
      if (body.worker)      sh.getRange(i+1, 11).setValue(body.worker);      // worker
      return json_({ ok:true, id:body.id });
    }
  }
  return json_({ ok:false, error:"id not found: " + body.id });
}
