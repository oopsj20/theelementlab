/**
 * 더엘리먼트랩 홈페이지 문의 → 구글 시트 저장 + 알림 메일
 *
 * 사용법
 * 1. 구글 시트 새로 만들기 → 확장 프로그램 → Apps Script
 * 2. 이 코드 전체를 붙여넣고 저장
 * 3. 배포 → 새 배포 → 유형: 웹 앱
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 * 4. 권한 승인 후 나오는 "웹 앱 URL"을 홈페이지 index.html의 FORM_ENDPOINT에 넣기
 */

// 알림 메일 받을 주소 (비워두면 배포한 계정의 Gmail로 갑니다)
const NOTIFY_EMAIL = '';

function doPost(e) {
  const p = e.parameter;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['접수일시', '이름', '연락처', '관심 프로그램', '문의 내용', '처리 상태']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f5f0e6');
    sheet.setFrozenRows(1);
  }

  // 수식 실행 방지: = + - @ 로 시작하는 값 앞에 ' 붙이기
  const clean = v => {
    const s = String(v || '').slice(0, 5000);
    return /^[=+\-@]/.test(s) ? "'" + s : s;
  };

  sheet.appendRow([
    new Date(),
    clean(p.name),
    clean(p.contact),
    clean(p.program),
    clean(p.msg),
    '신규',
  ]);

  const to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({
    to: to,
    subject: `[더엘리먼트랩 문의] ${p.program || ''} - ${p.name || ''}`,
    body:
      `이름: ${p.name}\n연락처: ${p.contact}\n관심 프로그램: ${p.program}\n\n${p.msg || ''}\n\n` +
      `전체 문의 보기: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`,
  });

  return ContentService.createTextOutput('ok');
}
