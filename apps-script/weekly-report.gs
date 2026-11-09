/**
 * 업무 대시보드 ↔ 「2026 드림메이커스 주간업무보고 ★」 연동
 *
 * 지정한 주의 시트에서 NAME(박예영) 행의 '업무 보고'(E열)·'업무 계획'(F열) 칸만 돌려줍니다.
 * 다른 사람의 칸은 읽지도 내보내지도 않습니다.
 *
 * 설치: 시트 → 확장 프로그램 → Apps Script → 이 코드 붙여넣기 → 배포 → 새 배포 → 웹 앱
 *       (실행 사용자: 나 / 액세스 권한: 모든 사용자) → 나온 주소를 대시보드
 *       운영자 모드 > 보고 양식 > '시트 연동 주소'에 붙여넣기
 *
 * 호출: GET <웹 앱 주소>?week=2026-11-02   (보고 주의 월요일)
 */
const SHEET_ID = '18J11mHW90HkDVFw6xv6npeoLgYf31FVa5FNBzvkPrgU';
const NAME = '박예영';
const COL_REPORT = 5; // E열
const COL_PLAN = 6;   // F열

function doGet(e) {
  try {
    const week = String((e && e.parameter && e.parameter.week) || '');
    const m = week.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return out({ error: 'week 값이 필요합니다 (예: 2026-11-02)' });
    const key = Number(m[2]) + '.' + Number(m[3]) + '.('; // "11.2.(" 형태로 비교
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const flat = s => String(s || '').replace(/\s+/g, '');

    for (const sh of ss.getSheets()) {
      const last = sh.getLastRow();
      if (last < 4) continue;
      const rows = sh.getRange(1, 1, last, Math.max(COL_PLAN, 7)).getDisplayValues();
      // 머리글 중 '보고 기간' 칸(E열)이 해당 주 월요일로 시작하는 시트 찾기
      const hit = rows.slice(0, 10).some(r => flat(r[COL_REPORT - 1]).startsWith(key));
      if (!hit) continue;
      // A~D열에 이름이 있는 행
      const row = rows.find(r => r.slice(0, 4).some(c => flat(c).includes(NAME)));
      if (!row) return out({ sheet: sh.getName(), error: NAME + ' 칸을 찾지 못했습니다' });
      const clean = s => String(s || '').replace(/\r\n?/g, '\n').replace(/\s+$/, '');
      return out({ sheet: sh.getName(), report: clean(row[COL_REPORT - 1]), plan: clean(row[COL_PLAN - 1]) });
    }
    return out({ error: Number(m[2]) + '. ' + Number(m[3]) + '.(월)부터 시작하는 주의 시트가 아직 없습니다' });
  } catch (err) {
    return out({ error: String(err && err.message || err) });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
