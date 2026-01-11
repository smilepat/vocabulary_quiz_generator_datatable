/**
 * SetupSheet.gs
 * 시트 초기 설정 함수
 *
 * 사용법: Google Apps Script에서 새 파일로 추가
 * (+ 버튼 > 스크립트 > 파일명: SetupSheet)
 */

/**
 * 시트 초기 설정
 * H~Q열에 문항 유형 헤더 자동 설정
 */
function setupSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();

  // H~Q열 헤더 설정 (10개 문항 유형) - 2D 배열로 한 번에 쓰기
  const itemTypes = [[
    '영한번역 빈칸채우기',  // H열
    '객관식',              // I열
    '유의어찾기',          // J열
    '반의어찾기',          // K열
    '품사맞추기',          // L열
    '영영풀이',            // M열
    '문장완성',            // N열
    '철자맞추기',          // O열
    '문맥추론',            // P열
    '단어배열'             // Q열
  ]];

  // CONFIG는 메인 코드에서 정의됨 (ITEM_START_COL = 8)
  const ITEM_START_COL = 8;

  // 한 번에 쓰기 (H1:Q1)
  sheet.getRange(1, ITEM_START_COL, 1, itemTypes[0].length).setValues(itemTypes);

  SpreadsheetApp.getUi().alert('H~Q열에 10개 문항 유형이 설정되었습니다.');
}
