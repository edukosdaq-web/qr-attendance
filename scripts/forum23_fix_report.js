function fixReportFormula() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('보고용_23기(후보자)');
  if (!sheet) { SpreadsheetApp.getUi().alert('보고용_23기(후보자) 시트가 없습니다'); return; }

  // Row 4에 있던 잘못된 QUERY 삭제
  sheet.getRange('A4').clearContent();

  // 헤더 파라미터 1로 수정한 QUERY 재삽입
  // 상세관리 시트의 row 1이 헤더임을 QUERY에게 알림
  const formula = "=QUERY('상세관리_23기(후보자)'!A:Q, \"SELECT A,B,C,D,E,L,P,Q WHERE C IS NOT NULL\", 1)";
  sheet.getRange('A4').setFormula(formula);

  SpreadsheetApp.getUi().alert('✅ 보고용 QUERY 수정 완료. 헤더 중복 제거됨.');
}
