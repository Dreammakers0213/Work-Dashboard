/**
 * 업무 대시보드 ↔ 「2026 드림메이커스 주간업무보고 ★」 연동 (시트 → 대시보드로 보내기)
 *
 * 1시간마다 주간업무보고 시트의 주별 탭에서 내 이름 행의 '업무 보고'(E열)·'업무 계획'(F열) 칸만 읽어
 * 대시보드 계정의 sheetweeks/<보고 주 월요일> 에 넣습니다. 다른 사람의 칸은 읽지도 보내지도 않습니다.
 * 대시보드에서는 주간업무보고 화면의 '시트에서 불러오기'를 누르면 이 내용이 채워집니다.
 *
 * 설치 (한 번만)
 *  1) script.google.com → 새 프로젝트 → 이 코드 전체를 붙여넣고 저장
 *  2) 왼쪽 톱니바퀴(프로젝트 설정) → 맨 아래 '스크립트 속성' → 속성 추가
 *       SHEET_ID      : 주간업무보고 시트 ID (시트 주소의 /d/ 와 /edit 사이 글자)
 *       SHEET_NAME    : 시트에서 찾을 내 이름 (A~D열에 적힌 그대로)
 *       DASH_NAME     : 대시보드 로그인 이름
 *       DASH_PASSWORD : 대시보드 로그인 비밀번호 (다른 곳에서 쓰지 않는 비밀번호)
 *     ※ 시트 ID·이름·비밀번호는 이 코드에 직접 적지 마세요. 코드는 공개 저장소에 올라갑니다.
 *  3) 위쪽 함수 목록에서 setup 을 고르고 ▶실행 → 권한 허용
 *     (지금 한 번 보내고, 이후 1시간마다 자동으로 보내도록 예약합니다)
 *  시트를 고친 뒤 바로 반영하고 싶으면 sync 를 ▶실행하면 됩니다.
 */
const COL_REPORT = 5;      // E열
const COL_PLAN = 6;        // F열
const DAYS_BACK = 70;      // 최근 몇 일치 주만 보낼지

// 대시보드(Firebase) 정보 — 바꾸지 마세요
const FB_API_KEY = 'AIzaSyC60VvBo_njL-4Ak0GI4cg9qVaegUYealE';
const FB_PROJECT = 'work-dashboard-48e63';
const FB_DOMAIN = '@work-dashboard.app';
const DASH_ORIGIN = 'https://dreammakers0213.github.io/';

/** 처음 한 번: 지금 보내고, 1시간마다 자동 실행 예약 */
function setup() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sync').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sync').timeBased().everyHours(1).create();
  sync();
}

/** 시트의 최근 주들을 대시보드로 보냅니다 */
function sync() {
  const props = PropertiesService.getScriptProperties();
  const name = props.getProperty('DASH_NAME'), pw = props.getProperty('DASH_PASSWORD');
  if (!name || !pw) throw new Error("프로젝트 설정 > 스크립트 속성에 DASH_NAME, DASH_PASSWORD 를 넣어 주세요");
  const sheetId = props.getProperty('SHEET_ID'), me = props.getProperty('SHEET_NAME') || name;
  if (!sheetId) throw new Error("프로젝트 설정 > 스크립트 속성에 SHEET_ID 를 넣어 주세요");
  const auth = signIn(name, pw);

  const weeks = readWeeks(sheetId, me);
  const since = Utilities.formatDate(new Date(Date.now() - DAYS_BACK * 864e5), 'Asia/Seoul', 'yyyy-MM-dd');
  let sent = 0, same = 0;
  weeks.filter(w => w.week >= since).forEach(w => {
    const sig = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, w.sheet + '\n' + w.report + '\n' + w.plan, Utilities.Charset.UTF_8));
    const key = 'sig_' + w.week;
    if (props.getProperty(key) === sig) { same++; return; }
    writeDoc(auth, 'sheetweeks/' + w.week, { week: w.week, sheet: w.sheet, report: w.report, plan: w.plan, syncedAt: new Date().toISOString() });
    props.setProperty(key, sig);
    sent++;
  });
  console.log('보냄 ' + sent + '주, 변경 없음 ' + same + '주 (시트에서 찾은 주 ' + weeks.length + '개)');
}

/** 각 탭에서 보고 주(E열 머리글의 시작일)와 내 칸을 찾습니다 */
function readWeeks(sheetId, me) {
  const ss = SpreadsheetApp.openById(sheetId);
  const flat = s => String(s || '').replace(/\s+/g, '');
  const clean = s => String(s || '').replace(/\r\n?/g, '\n').replace(/\s+$/, '');
  const fileYear = (ss.getName().match(/20\d\d/) || [String(new Date().getFullYear())])[0];
  const pad = n => ('0' + n).slice(-2);
  const out = [];
  for (const sh of ss.getSheets()) {
    const last = sh.getLastRow();
    if (last < 4) continue;
    const rows = sh.getRange(1, 1, last, Math.max(COL_PLAN, 7)).getDisplayValues();
    const head = rows.slice(0, 10);
    let md = null, year = null;
    for (const r of head) {
      const e = flat(r[COL_REPORT - 1]);
      const m = e.match(/^(\d{1,2})\.(\d{1,2})\.\(/);
      if (m && !md) md = m;
      const y = e.match(/(20\d\d)년/);
      if (y && !year) year = y[1];
    }
    if (!md) continue;
    const row = rows.find(r => r.slice(0, 4).some(c => flat(c).includes(flat(me))));
    if (!row) continue;
    out.push({
      week: (year || fileYear) + '-' + pad(md[1]) + '-' + pad(md[2]),
      sheet: sh.getName(),
      report: clean(row[COL_REPORT - 1]),
      plan: clean(row[COL_PLAN - 1])
    });
  }
  return out;
}

/** 대시보드 로그인 이름 → 계정 주소 (대시보드와 같은 규칙) */
function nameToEmail(n) {
  const norm = String(n || '').normalize('NFC').trim().replace(/\s+/g, ' ').toLowerCase();
  const hex = Utilities.newBlob('').setDataFromString(norm, 'UTF-8').getBytes().map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
  return 'u' + hex + FB_DOMAIN;
}

function signIn(name, pw) {
  const res = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=' + FB_API_KEY, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Referer: DASH_ORIGIN },
    payload: JSON.stringify({ email: nameToEmail(name), password: pw, returnSecureToken: true })
  });
  const d = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200) throw new Error('대시보드 로그인 실패: ' + ((d.error && d.error.message) || res.getResponseCode()) + ' / 스크립트 속성의 DASH_NAME, DASH_PASSWORD 를 확인해 주세요');
  return { uid: d.localId, token: d.idToken };
}

function writeDoc(auth, path, obj) {
  const fields = {};
  Object.keys(obj).forEach(k => fields[k] = { stringValue: String(obj[k]) });
  const url = 'https://firestore.googleapis.com/v1/projects/' + FB_PROJECT + '/databases/(default)/documents/users/' + auth.uid + '/' + path;
  const res = UrlFetchApp.fetch(url, {
    method: 'patch', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + auth.token, Referer: DASH_ORIGIN },
    payload: JSON.stringify({ fields })
  });
  if (res.getResponseCode() !== 200) throw new Error('대시보드에 쓰지 못했습니다 (' + res.getResponseCode() + '): ' + res.getContentText().slice(0, 300));
}
