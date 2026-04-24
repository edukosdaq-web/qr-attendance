/**
 * 제23기 코스닥CEO포럼 · 신청자 관리 시트 + DART/Yahoo 자동 조회
 * 버전 3.0 (2026-04-24) — 신청자↔후보자 이동 버튼 추가
 *
 * 기능:
 * - 신청자 상세/보고용 시트 + 후보자 상세/보고용 시트 + 대시보드 자동 생성
 * - DART API 연결 + 별도 재무제표 모두 조회
 * - Yahoo Finance API로 시가총액 자동 조회
 * - 폼 제출 시 자동 조회·채움 (트리거 필요)
 * - 신청자↔후보자 이동 메뉴 (상단 메뉴 "CEO포럼 관리")
 *
 * 사용법:
 * 1. 확장 프로그램 → Apps Script → 코드 전체 교체
 * 2. 함수 "oneTimeSetup" 실행 → 권한 승인
 * 3. 트리거 추가 → "onFormSubmitEnrich" → 양식 제출 시
 * 4. 스프레드시트 새로고침 → 상단 "CEO포럼 관리" 메뉴 나타남
 */

// ========== 설정 ==========
const DART_API_KEY = 'fc69a2acc94e67abc8fd1d5ffa86e29009f171f5';
const RESPONSE_SHEET_NAME = '설문지 응답 시트1';

// ========== 커스텀 메뉴 (스프레드시트 열 때 자동 실행) ==========
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CEO포럼 관리')
    .addItem('✔ 선택한 신청자를 후보자로 복사', 'menuApplicantToCandidate')
    .addItem('✔ 선택한 후보자를 신청자로 복사', 'menuCandidateToApplicant')
    .addSeparator()
    .addItem('📥 신청자 전체 DART 재조회 (공백 채움)', 'menuBackfillApplicants')
    .addItem('🔄 DART 회사 맵 갱신 (월 1회)', 'menuRefreshCorpMap')
    .addToUi();
}

// ========== 1회용 초기 세팅 ==========
function oneTimeSetup() {
  PropertiesService.getScriptProperties().setProperty('DART_API_KEY', DART_API_KEY);

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 기존 쓰레기 시트 정리
  ['시트1', '보고용(신청자)_23기', '상세관리_23기(신청자)', '보고용_23기(신청자)',
   '상세관리_23기(후보자)', '보고용_23기(후보자)', '대시보드_23기', '수기_신청자_23기'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) ss.deleteSheet(s);
  });

  createManualApplicantSheet(ss);   // 수기 신청자 (후보자→신청자 이동용)
  createDetailApplicantSheet(ss);
  createReportApplicantSheet(ss);
  createDetailCandidateSheet(ss);
  createReportCandidateSheet(ss);
  createDashboard(ss);

  Logger.log('✅ 초기 세팅 완료 - 6개 시트 생성됨');
}

function createManualApplicantSheet(ss) {
  const sheet = ss.insertSheet('수기_신청자_23기');
  // 컬럼 확장 (21)
  if (sheet.getMaxColumns() < 21) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 21 - sheet.getMaxColumns());
  }
  // 폼 응답 시트와 동일한 컬럼 순서 (A=타임스탬프, B=회사명...)
  const headers = ['타임스탬프', '회사명', '성명', '직위', '주요제품',
    '회사연락처', '휴대폰', '이메일', '생년월일', '생일양력음력', '골프핸디',
    '자택주소', '연락담당자', '담당자직위', '담당자연락처', '담당자이메일',
    '학력', '경력', '특이사항', '추천인', '개인정보동의'];
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#4a7d5a')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.hideSheet(); // 뒷편 관리용이라 기본 숨김
}

// ========== 시트 생성 ==========
function createDetailApplicantSheet(ss) {
  const sheet = ss.insertSheet('상세관리_23기(신청자)');

  // 컬럼 확장 (기본 26 → 33)
  const needCols = 33;
  const curCols = sheet.getMaxColumns();
  if (curCols < needCols) {
    sheet.insertColumnsAfter(curCols, needCols - curCols);
  }

  const headers = [
    'NO',                                          // A
    '회사명', '성명', '직위', '주요제품',          // B-E
    '휴대폰', '이메일', '생년월일', '자택주소', '추천인',       // F-J
    '연락담당자', '담당자직위', '담당자연락처', '담당자이메일', // K-N
    '학력', '경력', '특이사항', '골프핸디',        // O-R
    // DART 자동 조회
    '상장일', '시장구분', '업종',                  // S-U
    '본사주소',                                    // V
    '시가총액(억)',                                // W
    '매출액_연결(억)', '영업이익_연결(억)',        // X-Y
    '매출액_별도(억)', '영업이익_별도(억)',        // Z-AA
    '최대주주', '오너',                            // AB-AC
    // 운영 관리 (수기)
    '계산서', '입금일자', '금액', '비고'           // AD-AG
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1a4d6e')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);

  // QUERY: 폼 응답 + 수기 신청자 UNION
  // 두 소스에서 동일한 컬럼 순서로 pull 후 세로 결합
  const formula =
    "={" +
      "QUERY('" + RESPONSE_SHEET_NAME + "'!A2:U, " +
        '"SELECT B,C,D,E,G,H,I,L,T,M,N,O,P,Q,R,S,K ' +
        'WHERE B IS NOT NULL ORDER BY A ASC", 0);' +
      "IFERROR(QUERY('수기_신청자_23기'!A2:U, " +
        '"SELECT B,C,D,E,G,H,I,L,T,M,N,O,P,Q,R,S,K ' +
        'WHERE B IS NOT NULL", 0), {\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\",\"\"})' +
    "}";
  sheet.getRange('B2').setFormula(formula);

  // NO 자동 번호
  sheet.getRange('A2').setFormula(
    '=ARRAYFORMULA(IF(LEN(B2:B), ROW(B2:B)-1, ""))'
  );

  // 컬럼 너비
  const widths = [45,160,80,80,180,130,180,100,250,120,90,80,130,180,200,200,150,70,90,90,120,200,110,110,110,110,110,150,100,90,100,100,150];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}

function createReportApplicantSheet(ss) {
  const sheet = ss.insertSheet('보고용_23기(신청자)');

  // 컬럼 확장 (11)
  if (sheet.getMaxColumns() < 11) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 11 - sheet.getMaxColumns());
  }

  sheet.getRange('A1').setValue('「제23기 코스닥CEO포럼」 참가자 명단')
    .setFontWeight('bold').setFontSize(14);
  sheet.getRange('A1:K1').merge().setHorizontalAlignment('center');

  sheet.getRange('A2').setValue('갱신일:').setFontWeight('bold');
  sheet.getRange('B2').setFormula('=TEXT(NOW(),"yyyy-mm-dd hh:mm")');

  const headers = ['NO', '회사명', '성명', '직위', '주요제품',
    '시장구분', '상장일', '시가총액(억)', '매출액(억)', '최대주주', '비고'];
  sheet.getRange(3, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#8b3a1c')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // 상세관리에서 공유용 컬럼만 pull (민감정보 제외)
  // A=NO, B=회사명, C=성명, D=직위, E=주요제품, T=시장구분, S=상장일, W=시가총액, X=매출액_연결, AB=최대주주, AG=비고
  const formula = "=QUERY('상세관리_23기(신청자)'!A:AG, \"SELECT A,B,C,D,E,T,S,W,X,AB,AG WHERE B IS NOT NULL\", 0)";
  sheet.getRange('A4').setFormula(formula);

  sheet.setFrozenRows(3);
}

function createDetailCandidateSheet(ss) {
  const sheet = ss.insertSheet('상세관리_23기(후보자)');

  // 컬럼 확장 (23)
  if (sheet.getMaxColumns() < 23) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 23 - sheet.getMaxColumns());
  }

  const headers = [
    'NO', '등록일자', '회사명', '성명', '직위',
    '연락처', '휴대폰', '이메일',
    '오너', '생년', '주요제품', '법인분류', '본사/사무소', '상장일', '추천인',
    '연락상태', '비고',
    '담당자', '담당자직위', '담당자연락처', '담당자휴대폰', '담당자이메일',
    '통화내용'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1a4d6e')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.getRange('A2').setFormula(
    '=ARRAYFORMULA(IF(LEN(C2:C), ROW(C2:C)-1, ""))'
  );
}

function createReportCandidateSheet(ss) {
  const sheet = ss.insertSheet('보고용_23기(후보자)');

  sheet.getRange('A1').setValue('제23기 코스닥CEO포럼 후보자 명단')
    .setFontWeight('bold').setFontSize(14);
  sheet.getRange('A1:H1').merge().setHorizontalAlignment('center');

  const headers = ['NO', '등록일자', '회사명', '성명', '직위', '법인분류', '연락상태', '비고'];
  sheet.getRange(3, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#8b3a1c')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  const formula = "=QUERY('상세관리_23기(후보자)'!A:Q, \"SELECT A,B,C,D,E,L,P,Q WHERE C IS NOT NULL\", 0)";
  sheet.getRange('A4').setFormula(formula);

  sheet.setFrozenRows(3);
}

function createDashboard(ss) {
  const sheet = ss.insertSheet('대시보드_23기');

  sheet.getRange('A1').setValue('제23기 CEO포럼 · 실시간 대시보드')
    .setFontWeight('bold').setFontSize(16);
  sheet.getRange('A1:D1').merge();

  sheet.getRange('A3').setValue('총 신청자').setFontWeight('bold');
  sheet.getRange('B3').setFormula(`=COUNTA('${RESPONSE_SHEET_NAME}'!B2:B)`);

  sheet.getRange('A4').setValue('총 후보자').setFontWeight('bold');
  sheet.getRange('B4').setFormula("=COUNTA('상세관리_23기(후보자)'!C2:C)");

  sheet.getRange('A5').setValue('오늘 신청').setFontWeight('bold');
  sheet.getRange('B5').setFormula(
    `=COUNTIFS('${RESPONSE_SHEET_NAME}'!A:A,">="&TODAY(),'${RESPONSE_SHEET_NAME}'!A:A,"<"&TODAY()+1)`
  );

  sheet.getRange('A7').setValue('최근 신청 5명').setFontWeight('bold').setFontSize(12);
  sheet.getRange('A8').setFormula(
    `=QUERY('${RESPONSE_SHEET_NAME}'!A:D, "SELECT A,B,C,D WHERE B IS NOT NULL ORDER BY A DESC LIMIT 5", 0)`
  );
}

// ========== DART 회사 매핑 ==========
function refreshCorpCodeMap() {
  const key = PropertiesService.getScriptProperties().getProperty('DART_API_KEY');
  const url = `https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key=${key}`;

  const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const blob = response.getBlob();
  blob.setContentType('application/zip');

  const unzipped = Utilities.unzip(blob)[0];
  const xml = unzipped.getDataAsString('UTF-8');

  const nameMap = {}; // 회사명 → {corpCode, stockCode}
  const regex = /<list>[\s\S]*?<corp_code>(\d+)<\/corp_code>[\s\S]*?<corp_name>([^<]+)<\/corp_name>[\s\S]*?<stock_code>([^<]*)<\/stock_code>[\s\S]*?<\/list>/g;
  let match;
  while ((match = regex.exec(xml)) !== null) {
    const corpCode = match[1];
    const corpName = match[2].trim();
    const stockCode = match[3].trim();
    if (stockCode) {
      nameMap[corpName] = { corpCode, stockCode };
    }
  }

  // 용량 제한 때문에 한 개 property에 30만자 남짓밖에 못 넣어 — 청킹 필요
  const mapStr = JSON.stringify(nameMap);
  const chunkSize = 400000;
  const chunks = Math.ceil(mapStr.length / chunkSize);
  const props = PropertiesService.getScriptProperties();

  // 기존 맵 청크 제거
  for (let i = 0; i < 20; i++) props.deleteProperty('CORP_MAP_' + i);

  for (let i = 0; i < chunks; i++) {
    props.setProperty('CORP_MAP_' + i, mapStr.slice(i * chunkSize, (i + 1) * chunkSize));
  }
  props.setProperty('CORP_MAP_CHUNKS', String(chunks));
  Logger.log(`회사 맵 갱신: ${Object.keys(nameMap).length}개, ${chunks}청크`);
}

function getCorpMap() {
  const props = PropertiesService.getScriptProperties();
  const chunks = parseInt(props.getProperty('CORP_MAP_CHUNKS') || '0');
  if (chunks === 0) {
    refreshCorpCodeMap();
    return getCorpMap();
  }
  let str = '';
  for (let i = 0; i < chunks; i++) {
    str += props.getProperty('CORP_MAP_' + i) || '';
  }
  return JSON.parse(str);
}

function lookupCompany(companyName) {
  const map = getCorpMap();
  const clean = companyName.trim().replace(/\(주\)|㈜/g, '').trim();
  // 정확 매칭 우선
  if (map[clean]) return map[clean];
  if (map[companyName]) return map[companyName];
  // Fuzzy: 공백 제거 후 매칭
  const noSpace = clean.replace(/\s/g, '');
  for (const [key, val] of Object.entries(map)) {
    if (key.replace(/\s/g, '') === noSpace) return val;
  }
  return null;
}

// ========== DART 데이터 조회 ==========
function fetchCompanyInfo(corpCode) {
  const key = PropertiesService.getScriptProperties().getProperty('DART_API_KEY');
  const url = `https://opendart.fss.or.kr/api/company.json?crtfc_key=${key}&corp_code=${corpCode}`;
  const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  return JSON.parse(response.getContentText());
}

/**
 * 재무제표 조회 - 연결(CFS) + 별도(OFS) 모두
 * @return {cfs: {revenue, operatingProfit}, ofs: {revenue, operatingProfit}}
 */
function fetchFinancials(corpCode, year) {
  const key = PropertiesService.getScriptProperties().getProperty('DART_API_KEY');
  const bsnsYear = year || (new Date().getFullYear() - 1);
  const url = `https://opendart.fss.or.kr/api/fnlttSinglAcnt.json?crtfc_key=${key}&corp_code=${corpCode}&bsns_year=${bsnsYear}&reprt_code=11011`;
  const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const data = JSON.parse(response.getContentText());

  if (data.status !== '000') {
    // 작년 데이터 없으면 2년 전으로 재시도
    if (!year && bsnsYear === new Date().getFullYear() - 1) {
      return fetchFinancials(corpCode, bsnsYear - 1);
    }
    return null;
  }

  const result = { cfs: {}, ofs: {} };
  (data.list || []).forEach(item => {
    const bucket = item.fs_div === 'CFS' ? result.cfs : (item.fs_div === 'OFS' ? result.ofs : null);
    if (!bucket) return;
    if (item.sj_div === 'IS') {
      if (item.account_nm === '매출액') bucket.revenue = item.thstrm_amount;
      if (item.account_nm === '영업이익') bucket.operatingProfit = item.thstrm_amount;
    }
  });
  return result;
}

function fetchMajorShareholder(corpCode, year) {
  const key = PropertiesService.getScriptProperties().getProperty('DART_API_KEY');
  const bsnsYear = year || (new Date().getFullYear() - 1);
  const url = `https://opendart.fss.or.kr/api/hyslrSttus.json?crtfc_key=${key}&corp_code=${corpCode}&bsns_year=${bsnsYear}&reprt_code=11011`;
  const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const data = JSON.parse(response.getContentText());

  if (data.status !== '000' || !data.list || data.list.length === 0) return null;
  const top = data.list.reduce((a, b) =>
    (parseFloat(a.trmend_posesn_stock_qota_rt) > parseFloat(b.trmend_posesn_stock_qota_rt)) ? a : b
  );
  return {
    name: top.nm,
    ratio: top.trmend_posesn_stock_qota_rt
  };
}

// ========== 시가총액 조회 (Yahoo Finance) ==========
/**
 * 한국 종목은 '005930.KS' 형식. 코스닥은 '.KQ' 종목도 있음.
 * Yahoo에서 둘 다 시도.
 */
function fetchMarketCap(stockCode) {
  const suffixes = ['.KS', '.KQ'];
  for (const suffix of suffixes) {
    const symbol = stockCode + suffix;
    const url = `https://query1.finance.yahoo.com/v7/finance/quote?symbols=${symbol}`;
    try {
      const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true, headers: {'User-Agent': 'Mozilla/5.0'}});
      const data = JSON.parse(response.getContentText());
      if (data.quoteResponse && data.quoteResponse.result && data.quoteResponse.result.length > 0) {
        const quote = data.quoteResponse.result[0];
        if (quote.marketCap) {
          return quote.marketCap; // KRW 단위
        }
      }
    } catch (e) {
      Logger.log('Yahoo 조회 실패: ' + symbol + ' / ' + e);
    }
  }
  return null;
}

// ========== 자동 채움 공통 로직 ==========
function enrichRow(sheet, rowNum, companyName) {
  const comp = lookupCompany(companyName);
  if (!comp) {
    Logger.log(`매칭 실패: ${companyName}`);
    return false;
  }

  const info = fetchCompanyInfo(comp.corpCode);
  const financials = fetchFinancials(comp.corpCode);
  const shareholder = fetchMajorShareholder(comp.corpCode);
  const marketCap = fetchMarketCap(comp.stockCode);

  // 컬럼 인덱스 (1-based)
  // S=19 상장일, T=20 시장구분, U=21 업종, V=22 본사주소
  // W=23 시가총액(억)
  // X=24 매출액_연결, Y=25 영업이익_연결
  // Z=26 매출액_별도, AA=27 영업이익_별도
  // AB=28 최대주주

  if (info && info.status === '000') {
    sheet.getRange(rowNum, 19).setValue(info.list_dt || '');
    sheet.getRange(rowNum, 20).setValue(info.corp_cls_nm || corpClsToKorean(info.corp_cls));
    sheet.getRange(rowNum, 21).setValue(info.induty_code || '');
    sheet.getRange(rowNum, 22).setValue(info.adres || '');
  }

  if (marketCap) {
    sheet.getRange(rowNum, 23).setValue(Math.round(marketCap / 100000000));
  }

  if (financials) {
    if (financials.cfs.revenue) sheet.getRange(rowNum, 24).setValue(toBillion(financials.cfs.revenue));
    if (financials.cfs.operatingProfit) sheet.getRange(rowNum, 25).setValue(toBillion(financials.cfs.operatingProfit));
    if (financials.ofs.revenue) sheet.getRange(rowNum, 26).setValue(toBillion(financials.ofs.revenue));
    if (financials.ofs.operatingProfit) sheet.getRange(rowNum, 27).setValue(toBillion(financials.ofs.operatingProfit));
  }

  if (shareholder) {
    sheet.getRange(rowNum, 28).setValue(`${shareholder.name} (${shareholder.ratio}%)`);
  }

  return true;
}

function toBillion(str) {
  if (!str) return '';
  const num = Number(String(str).replace(/,/g, ''));
  if (isNaN(num)) return '';
  return Math.round(num / 100000000);
}

function corpClsToKorean(cls) {
  const map = { Y: '유가증권', K: '코스닥', N: '코넥스', E: '기타' };
  return map[cls] || cls;
}

// ========== 폼 제출 트리거 ==========
function onFormSubmitEnrich(e) {
  try {
    const values = e.values || [];
    const companyName = (values[1] || '').trim();
    if (!companyName) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('상세관리_23기(신청자)');
    if (!sheet) return;

    // 최근 추가 행 찾기
    const lastRow = sheet.getLastRow();
    let targetRow = -1;
    for (let r = lastRow; r >= 2; r--) {
      if (String(sheet.getRange(r, 2).getValue()).trim() === companyName) {
        targetRow = r;
        break;
      }
    }
    if (targetRow === -1) {
      Utilities.sleep(2000); // QUERY 반영 대기
      for (let r = sheet.getLastRow(); r >= 2; r--) {
        if (String(sheet.getRange(r, 2).getValue()).trim() === companyName) {
          targetRow = r;
          break;
        }
      }
    }
    if (targetRow === -1) return;

    enrichRow(sheet, targetRow, companyName);
    Logger.log(`자동 채움: ${companyName} → 행 ${targetRow}`);
  } catch (err) {
    Logger.log('onFormSubmitEnrich 오류: ' + err.toString());
  }
}

// ========== 메뉴 액션: 신청자 → 후보자 ==========
function menuApplicantToCandidate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== '상세관리_23기(신청자)') {
    ui.alert('먼저 "상세관리_23기(신청자)" 시트에서 이동할 행을 선택하세요.');
    return;
  }

  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow < 2) {
    ui.alert('헤더 행은 이동할 수 없습니다.');
    return;
  }

  const candidateSheet = ss.getSheetByName('상세관리_23기(후보자)');
  if (!candidateSheet) {
    ui.alert('후보자 시트가 없습니다. oneTimeSetup을 먼저 실행하세요.');
    return;
  }

  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  let moved = 0;

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const data = sheet.getRange(row, 1, 1, 33).getValues()[0];
    // data 인덱스:
    // 0 NO, 1 회사명, 2 성명, 3 직위, 4 주요제품,
    // 5 휴대폰, 6 이메일, 7 생년월일, 8 자택주소, 9 추천인,
    // 10 연락담당자, 11 담당자직위, 12 담당자연락처, 13 담당자이메일,
    // 14 학력, 15 경력, 16 특이사항, 17 골프핸디,
    // 18 상장일, 19 시장구분, 20 업종, 21 본사주소, 22 시가총액,
    // 23 매출액_연결, 24 영업이익_연결, 25 매출액_별도, 26 영업이익_별도,
    // 27 최대주주, 28 오너, 29 계산서, 30 입금일자, 31 금액, 32 비고

    if (!data[1]) continue; // 회사명 없으면 스킵

    // 후보자 시트 컬럼 순서에 맞게 매핑
    const candidateRow = [
      '',              // A NO (수식이 자동 채움)
      today,           // B 등록일자
      data[1],         // C 회사명
      data[2],         // D 성명
      data[3],         // E 직위
      data[5],         // F 연락처 (휴대폰)
      data[5],         // G 휴대폰
      data[6],         // H 이메일
      data[28],        // I 오너
      extractYear(data[7]), // J 생년
      data[4],         // K 주요제품
      data[19],        // L 법인분류
      data[21],        // M 본사/사무소
      data[18],        // N 상장일
      data[9],         // O 추천인
      '신청자에서 이관(' + today + ')', // P 연락상태
      data[32],        // Q 비고
      data[10],        // R 담당자
      data[11],        // S 담당자직위
      data[12],        // T 담당자연락처
      data[12],        // U 담당자휴대폰
      data[13],        // V 담당자이메일
      ''               // W 통화내용
    ];

    const appendRow = candidateSheet.getLastRow() + 1;
    candidateSheet.getRange(appendRow, 1, 1, candidateRow.length).setValues([candidateRow]);
    moved++;
  }

  ui.alert(`✅ ${moved}명이 후보자로 복사되었습니다.\n\n※ 신청자 시트는 폼 원본이라 자동 삭제되지 않습니다. 필요 시 비고 컬럼에 "후보이관" 등 표시해주세요.`);
}

// ========== 메뉴 액션: 후보자 → 신청자 ==========
function menuCandidateToApplicant() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  if (sheet.getName() !== '상세관리_23기(후보자)') {
    ui.alert('먼저 "상세관리_23기(후보자)" 시트에서 이동할 행을 선택하세요.');
    return;
  }

  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  if (startRow < 2) {
    ui.alert('헤더 행은 이동할 수 없습니다.');
    return;
  }

  const manualSheet = ss.getSheetByName('수기_신청자_23기');
  if (!manualSheet) {
    ui.alert('수기_신청자_23기 시트가 없습니다. oneTimeSetup을 먼저 실행하세요.');
    return;
  }

  // 숨긴 시트면 잠시 보이게
  if (manualSheet.isSheetHidden()) {
    manualSheet.showSheet();
  }

  const now = new Date();
  let moved = 0;

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const data = sheet.getRange(row, 1, 1, 23).getValues()[0];
    // 후보자 시트 컬럼 인덱스:
    // 0 NO, 1 등록일자, 2 회사명, 3 성명, 4 직위,
    // 5 연락처, 6 휴대폰, 7 이메일,
    // 8 오너, 9 생년, 10 주요제품, 11 법인분류, 12 본사/사무소, 13 상장일, 14 추천인,
    // 15 연락상태, 16 비고,
    // 17 담당자, 18 담당자직위, 19 담당자연락처, 20 담당자휴대폰, 21 담당자이메일, 22 통화내용

    if (!data[2]) continue; // 회사명 없으면 스킵

    // 수기_신청자 시트 컬럼 순서 (폼 응답 시트와 동일: A=타임스탬프, B=회사명...)
    const birthStr = data[9] ? (data[9] + '년생') : '';
    const manualRow = [
      now,              // A 타임스탬프
      data[2],          // B 회사명
      data[3],          // C 성명
      data[4],          // D 직위
      data[10],         // E 주요제품
      data[5],          // F 회사연락처
      data[6],          // G 휴대폰
      data[7],          // H 이메일
      birthStr,         // I 생년월일
      '',               // J 생일양력음력
      '',               // K 골프핸디
      data[12],         // L 자택주소 (본사로 대체)
      data[17],         // M 연락담당자
      data[18],         // N 담당자직위
      data[19],         // O 담당자연락처
      data[21],         // P 담당자이메일
      '',               // Q 학력
      '',               // R 경력
      data[16],         // S 특이사항 (비고)
      data[14],         // T 추천인
      '동의'             // U 개인정보동의
    ];

    const appendRow = manualSheet.getLastRow() + 1;
    manualSheet.getRange(appendRow, 1, 1, manualRow.length).setValues([manualRow]);
    moved++;
  }

  ui.alert(`✅ ${moved}명이 신청자로 복사되었습니다.\n\n※ 수기_신청자_23기 시트에 추가되어 상세관리_23기(신청자)에 자동 반영됩니다.\n※ 후보자 시트는 그대로 유지됩니다. 필요 시 연락상태를 "신청완료" 등으로 변경해주세요.`);
}

// ========== 메뉴 액션: 일괄 DART 조회 ==========
function menuBackfillApplicants() {
  backfillAllCompanies();
  SpreadsheetApp.getUi().alert('✅ 신청자 전체 DART 재조회 완료 (실행 로그 확인)');
}

// ========== 메뉴 액션: 회사 맵 갱신 ==========
function menuRefreshCorpMap() {
  refreshCorpCodeMap();
  SpreadsheetApp.getUi().alert('✅ DART 회사 맵 갱신 완료');
}

// ========== 유틸: 생년월일에서 년도 추출 ==========
function extractYear(dateStr) {
  if (!dateStr) return '';
  const match = String(dateStr).match(/(\d{4})/);
  return match ? match[1] : '';
}

// ========== 수동 일괄 채움 ==========
function backfillAllCompanies() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('상세관리_23기(신청자)');
  const lastRow = sheet.getLastRow();
  let success = 0, fail = 0;

  for (let r = 2; r <= lastRow; r++) {
    const companyName = String(sheet.getRange(r, 2).getValue()).trim();
    if (!companyName) continue;

    // 이미 채워진 행은 스킵
    const existing = sheet.getRange(r, 23).getValue();
    if (existing) continue;

    if (enrichRow(sheet, r, companyName)) success++;
    else fail++;

    Utilities.sleep(500);
  }
  Logger.log(`일괄 채움 완료: 성공 ${success} / 실패 ${fail}`);
}
