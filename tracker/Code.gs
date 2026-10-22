/**
 * 포트폴리오 방문 추적 — Google Sheets 연동 (Apps Script)
 * 1) 추적용 스프레드시트에서 [확장 프로그램] → [Apps Script] 열기
 * 2) 이 코드를 전체 붙여넣고 저장
 * 3) setup() 한 번 실행 (권한 허용) → 'log' / '요약' 시트 생성
 * 4) [배포] → [새 배포] → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → URL 복사
 * 5) index.html 의 CONFIG.TRACK_URL 에 붙여넣기
 */

const LOG = 'log';
const SUMMARY = '요약';
const HEADERS = ['시각', '이벤트', '유입(ref)', '섹션', '라벨', '체류(초)', '스크롤(%)',
                 '세션ID', '방문자ID', '이전 페이지', '기기', '언어', '경로'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const sh = sheet_(LOG);
    sh.appendRow([
      new Date(), clip_(d.event), clip_(d.ref), clip_(d.section), clip_(d.label),
      num_(d.duration), num_(d.scroll), clip_(d.sid), clip_(d.vid),
      clip_(d.referrer), clip_(d.device), clip_(d.lang), clip_(d.path)
    ]);
  } catch (err) {
    console.error(err);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}

// 브라우저에서 URL을 열었을 때 동작 확인용
function doGet() {
  return ContentService.createTextOutput('portfolio tracker is running');
}

function setup() {
  const sh = sheet_(LOG);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#13213c').setFontColor('#ffffff');
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }

  const s = sheet_(SUMMARY);
  s.clear();
  const navy = '#13213c';
  s.getRange('A1').setValue('유입별 방문 (ref)').setFontWeight('bold').setFontColor(navy);
  s.getRange('A2').setFormula(
    "=IFERROR(QUERY(log!A:M,\"select C, count(H), max(A) where B='visit' group by C order by count(H) desc label C '유입', count(H) '방문수', max(A) '최근 방문'\",1),\"데이터 없음\")");

  s.getRange('E1').setValue('섹션 도달 수').setFontWeight('bold').setFontColor(navy);
  s.getRange('E2').setFormula(
    "=IFERROR(QUERY(log!A:M,\"select D, count(H) where B='section' group by D order by count(H) desc label D '섹션', count(H) '도달 수'\",1),\"데이터 없음\")");

  s.getRange('H1').setValue('유입별 평균 체류·스크롤').setFontWeight('bold').setFontColor(navy);
  s.getRange('H2').setFormula(
    "=IFERROR(QUERY(log!A:M,\"select C, avg(F), avg(G), count(H) where B='leave' group by C label C '유입', avg(F) '평균 체류(초)', avg(G) '평균 스크롤(%)', count(H) '이탈 기록'\",1),\"데이터 없음\")");

  s.getRange('M1').setValue('클릭').setFontWeight('bold').setFontColor(navy);
  s.getRange('M2').setFormula(
    "=IFERROR(QUERY(log!A:M,\"select E, count(H) where B='click' group by E order by count(H) desc label E '항목', count(H) '클릭 수'\",1),\"데이터 없음\")");

  s.getRange('A20').setValue('총 방문 / 고유 방문자').setFontWeight('bold').setFontColor(navy);
  s.getRange('A21').setFormula("=COUNTIF(log!B:B,\"visit\")");
  s.getRange('B21').setFormula("=IFERROR(COUNTUNIQUE(FILTER(log!I:I,log!B:B=\"visit\")),0)");
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function clip_(v) { return v == null ? '' : String(v).slice(0, 300); }
function num_(v) { return v == null || v === '' || isNaN(v) ? '' : Number(v); }
