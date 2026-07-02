/**
 * 노무 리서치 포털 — Google Apps Script 중계 서버 (자동화판)
 * ------------------------------------------------------------
 * 게시판이 GitHub 대신 이 구글시트를 직접 읽습니다.
 * → 답변이 생기면 GitHub 업로드 없이 자동으로 게시판에 뜹니다.
 *
 * 역할:
 *   1) 홈페이지 폼 접수            doPost (accessCode 확인)
 *   2) 게시판 목록 제공(공개)       doGet ?action=board&code=ACCESS_CODE
 *   3) 답변 1건 열람(공개)         doGet ?action=answer&id=..&code=ACCESS_CODE
 *   4) 실행 루프가 새 질문 가져감    doGet ?action=pending&key=SECRET
 *   5) 실행 루프가 답변/상태 기록    doPost {action:update, key:SECRET, ...}
 *
 * [수정 후 반드시] 배포 > 배포 관리 > 연필 > 버전:새 버전 > 배포
 */

var SECRET = "nomu2026";        // 실행 루프와 공유하는 비밀키
var ACCESS_CODE = "beyond";     // 직원이 질문 낼 때 + 게시판 볼 때 입력하는 공통 코드
var SHEET_NAME = "queue";

// 열 순서: 1id 2ts 3date 4name 5dept 6depth 7question 8status 9answer_url 10processed_at 11worker 12answer_html
function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["id","ts","date","name","dept","depth","question","status","answer_url","processed_at","worker","answer_html"]);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function rows_() {
  var sh = sheet_();
  var vals = sh.getDataRange().getValues();
  var head = vals.shift();
  return vals.map(function(r){ var o={}; head.forEach(function(h,i){ o[h]=r[i]; }); return o; });
}

/** 폼 접수 & 루프의 상태/답변 기록 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || "{}");

    // (A) 실행 루프가 답변/상태 기록
    if (body.action === "update") {
      if (body.key !== SECRET) return json_({ ok:false, error:"unauthorized" });
      return updateRow_(body);
    }

    // (B) 홈페이지 신규 접수 — 팀 접속코드 확인
    if ((body.accessCode || "") !== ACCESS_CODE) {
      return json_({ ok:false, error:"접속코드가 올바르지 않습니다. 회사에서 받은 코드를 확인하세요." });
    }
    var q = (body.question || "").toString().trim();
    var name = (body.name || "").toString().trim();
    if (!q || !name) return json_({ ok:false, error:"이름/질문 누락" });

    var sh = sheet_();
    var now = new Date();
    var id = "NR-" + Utilities.formatDate(now, "Asia/Seoul", "yyMMdd") + "-" + ("000" + sh.getLastRow()).slice(-3);
    var depth = ["full","standard","quick"].indexOf(body.depth) >= 0 ? body.depth : "standard";
    sh.appendRow([ id, now.getTime(), Utilities.formatDate(now,"Asia/Seoul","yyyy-MM-dd HH:mm"),
      name, (body.dept||"").toString().trim(), depth, q, "recv", "", "", "", "" ]);
    return json_({ ok:true, id:id });

  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

function doGet(e) {
  var p = e.parameter || {};

  // (1) 게시판 목록 — 접속코드로 보호(답변 본문은 뺌)
  if (p.action === "board") {
    if (p.code !== ACCESS_CODE) return json_({ ok:false, error:"unauthorized" });
    var items = rows_().map(function(o){
      return { id:o.id, ts:o.ts, date:o.date, name:o.name, dept:o.dept, depth:o.depth,
               question:o.question, status:o.status, has_answer: !!(o.answer_html && String(o.answer_html).length>10) };
    });
    return json_({ ok:true, items:items });
  }

  // (2) 답변 1건 열람 — HTML 로 렌더
  if (p.action === "answer") {
    if (p.code !== ACCESS_CODE) return HtmlService.createHtmlOutput("<h3>접근 코드가 필요합니다.</h3>");
    var found = null;
    rows_().forEach(function(o){ if (o.id === p.id) found = o; });
    if (!found) return HtmlService.createHtmlOutput("<h3>해당 접수번호를 찾을 수 없습니다: " + (p.id||"") + "</h3>");
    if (found.status !== "done") return HtmlService.createHtmlOutput("<h3>아직 답변이 게시되지 않았습니다. (상태: " + found.status + ")</h3>");
    return HtmlService.createHtmlOutput(String(found.answer_html || "<p>답변 내용이 비어 있습니다.</p>"))
      .setTitle("노무 리서치 답변 " + found.id);
  }

  // (3) 실행 루프 전용 — 새 질문(recv) 목록
  if (p.key !== SECRET) return json_({ ok:false, error:"unauthorized" });
  var all = rows_();
  if (p.action === "pending") all = all.filter(function(o){ return o.status === "recv"; });
  return json_({ ok:true, items:all });
}

function updateRow_(body) {
  var sh = sheet_();
  var vals = sh.getDataRange().getValues();
  for (var i=1; i<vals.length; i++) {
    if (vals[i][0] === body.id) {
      if (body.status)       sh.getRange(i+1, 8).setValue(body.status);
      if (body.answer_url)   sh.getRange(i+1, 9).setValue(body.answer_url);
      sh.getRange(i+1, 10).setValue(Utilities.formatDate(new Date(),"Asia/Seoul","yyyy-MM-dd HH:mm"));
      if (body.worker)       sh.getRange(i+1, 11).setValue(body.worker);
      if (body.answer_html)  sh.getRange(i+1, 12).setValue(body.answer_html);
      return json_({ ok:true, id:body.id });
    }
  }
  return json_({ ok:false, error:"id not found: " + body.id });
}
