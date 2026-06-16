/**
 * 고객 DB 관리 시스템 — 구글시트 백엔드 (Google Apps Script 웹앱)
 *
 * [설치]
 * 1) 구글시트 새로 만들기 → 메뉴 [확장 프로그램] → [Apps Script]
 * 2) 기본 코드를 지우고 이 파일 내용을 통째로 붙여넣기 → 저장(💾)
 * 3) 오른쪽 위 [배포] → [새 배포] → 유형 "웹 앱"
 *      - 실행 주체: 나
 *      - 액세스 권한: "모든 사용자"
 *    → 배포 → 권한 승인 → 생성된 웹앱 URL(.../exec)을 복사
 * 4) 앱(고객 DB)의 "☁ 동기화"에 그 URL을 붙여넣고 [지금 동기화]
 *
 * 코드를 수정하면 [배포] → [배포 관리] → 기존 배포 편집 → 버전 "새 버전"으로 다시 배포하세요.
 */

var SHEET_NAME = 'DB';

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  return sh;
}

// 읽기: 시트 전체를 2차원 배열(첫 줄=헤더)로 반환. callback 파라미터가 있으면 JSONP.
function doGet(e) {
  var sh = getSheet_();
  var values = sh.getLastRow() > 0 ? sh.getDataRange().getValues() : [];
  var payload = JSON.stringify({ ok: true, values: values });
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + payload + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

// 쓰기: { action:'save', values:[[헤더...],[행...]] } 를 받아 시트를 통째로 교체.
function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { /* 잠금 실패 시에도 진행 */ }
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'save' && body.values && body.values.length) {
      var sh = getSheet_();
      sh.clearContents();
      var rows = body.values;
      var cols = rows[0].length;
      // 모든 행을 동일한 열 수로 정규화
      var norm = rows.map(function (r) {
        r = r || [];
        if (r.length < cols) r = r.concat(new Array(cols - r.length).fill(''));
        return r.slice(0, cols);
      });
      sh.getRange(1, 1, norm.length, cols).setValues(norm);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
