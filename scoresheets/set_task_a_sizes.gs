/**
 * Google Sheets Apps Script: set the TASK A sheet to the TASK B sizes (px).
 * 확장 프로그램 > Apps Script 에 붙여넣고, TASK A 시트를 연 상태에서 setTaskASizes 실행.
 */
function setTaskASizes() {
  const sh = SpreadsheetApp.getActiveSheet();
  // columns: A=116, B=481, C..N=100
  sh.setColumnWidth(1, 116);
  sh.setColumnWidth(2, 481);
  sh.setColumnWidths(3, 12, 100);
  // rows 1-4: 89, 32, 37, 21
  sh.setRowHeightsForced(1, 1, 89);
  sh.setRowHeightsForced(2, 1, 32);
  sh.setRowHeightsForced(3, 1, 37);
  sh.setRowHeightsForced(4, 1, 21);
  // 항목 9행 (5-13) = 40, 총합/유의 사항 (14-17) = 21, 노트 (18) = 110
  sh.showRows(5, 14);
  sh.setRowHeightsForced(5, 9, 40);
  sh.setRowHeightsForced(14, 4, 21);
  sh.setRowHeightsForced(18, 1, 110);
}
