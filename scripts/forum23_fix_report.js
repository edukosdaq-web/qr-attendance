function fixReportFormula() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('보고용_23기(후보자)');
  if (!sheet) { SpreadsheetApp.getUi().alert('보고용_23기(후보자) 시트가 없습니다'); return; }

  // A4부터 아래로 전부 삭제
  const lastRow = sheet.getLastRow();
  if (lastRow >= 4) {
    sheet.getRange(4, 1, lastRow - 3, sheet.getLastColumn()).clearContent();
  }

  // 상세관리 시트의 A2부터 pull (헤더 행 제외) → Col1/Col2 등 위치 기반 참조
  const formula = "=QUERY('상세관리_23기(후보자)'!A2:Q, \"SELECT Col1,Col2,Col3,Col4,Col5,Col12,Col16,Col17 WHERE Col3 IS NOT NULL\", 0)";
  sheet.getRange('A4').setFormula(formula);

  SpreadsheetApp.getUi().alert('✅ 보고용 QUERY 수정 완료. 헤더 중복 완전 제거.');
}

// 신청자 보고용도 같은 방식으로 수정
function fixReportApplicant() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('보고용_23기(신청자)');
  if (!sheet) { SpreadsheetApp.getUi().alert('보고용_23기(신청자) 시트가 없습니다'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow >= 4) {
    sheet.getRange(4, 1, lastRow - 3, sheet.getLastColumn()).clearContent();
  }

  // A=NO, B=회사명, C=성명, D=직위, E=주요제품, T=시장구분, S=상장일, W=시가총액, X=매출액, AB=최대주주, AG=비고
  // A2:AG 범위에서 Col1,Col2,Col3,Col4,Col5,Col19(시장구분T),Col18(상장일S),Col22(시가W),Col23(매출X),Col27(최대주주AB),Col32(비고AG)
  const formula = "=QUERY('상세관리_23기(신청자)'!A2:AG, \"SELECT Col1,Col2,Col3,Col4,Col5,Col20,Col19,Col23,Col24,Col28,Col33 WHERE Col2 IS NOT NULL\", 0)";
  sheet.getRange('A4').setFormula(formula);

  SpreadsheetApp.getUi().alert('✅ 신청자 보고용 QUERY도 수정 완료.');
}
