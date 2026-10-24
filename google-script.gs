/**
 * 푸드코트 투표 — 구글폼 자동 생성 + 집계 + 실시간 순위
 *
 * 1) 새 구글 시트 → 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기 → 저장
 * 2) 위쪽 함수 목록에서 createForm 선택 → 실행 (처음 한 번만, 권한 승인 필요)
 * 3) '설정' 시트의 '웹페이지 연결 링크'를 index.html 의 formPrefillLink 에 붙여넣기
 * 4) 실시간 순위: 배포 → 새 배포 → 웹 앱 (액세스: 모든 사용자) → 웹 앱 URL 을 index.html 의 liveUrl 에 붙여넣기
 * 5) 운영본부: 시트 새로고침 → 메뉴 '🏆 투표 집계 → 운영본부 코드·PIN 설정' (코드를 붙여넣을 때마다 값은 유지됨)
 * 6) 최종 집계: 메뉴 '🏆 투표 집계 → 집계 실행'
 *
 * 코드를 고친 뒤에는 '배포 → 배포 관리 → ✏️ → 버전: 새 버전 → 배포' (새 배포를 하면 웹 앱 주소가 바뀜)
 *
 * 집계 규칙
 * - 같은 사람(기수+중대+이름)이 여러 번 투표하면 '마지막 응답'만 인정 (투표 변경)
 * - 같은 기기(폰)에서 온 표도 '마지막 응답'만 인정. 단, PIN으로 인증된 운영본부 폰은 예외
 * - 한 사람이 같은 부스를 두 번 고른 경우 1표로 처리
 * - 순위: 득표수 → 동점이면 '단독 지지'(그 부스 한 곳만 고른 표) 많은 순 → 그래도 같으면 같은 순위(동점)
 * - 운영본부가 투표를 마감한 동안 들어온 응답은 무효 (구글폼이 받아들여도 집계에서 뺌)
 * - 투표를 마감하면 실시간 순위도 자동으로 숨김 (시상식 때 '순위 공개'로 다시 보여줄 수 있음)
 * - 결과는 '집계' 시트에, 무효가 된 이전 투표는 '이전투표', 마감 중 응답은 '마감후응답' 시트에 기록
 */
const FORM_TITLE = '공군사관학교 축제 푸드코트 맛 대결';
const FIELDS = ['기수', '중대', '이름', '부스1', '부스2'];
const DEVICE_FIELD = '기기';
const AWARDS = ['교장상', '생도대장상', '생도전대장상'];
const SETUP_SHEET = '설정';
const RESULT_SHEET = '집계';
const OLD_SHEET = '이전투표';
const LATE_SHEET = '마감후응답';
const CACHE_SECONDS = 15; // 실시간 순위를 다시 계산하는 최소 간격

// 운영본부 인증: 운영본부 폰에서 '투표주소/?staff=코드' 로 들어온 뒤 PIN 을 입력해야 함.
// 코드·PIN은 이 파일에 적지 않고 Apps Script '스크립트 속성'에 저장함 (이 파일은 공개 저장소에 올라감).
// 시트 메뉴 '운영본부 코드·PIN 설정'에서 입력·변경. 바꾸면 기존에 로그인한 폰은 모두 로그아웃됨 (재배포 불필요).
const MAX_PIN_TRIES = 5; // 링크를 가진 사람이 PIN을 이만큼 틀리면 10분간 잠김

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏆 투표 집계')
    .addItem('집계 실행', 'tally')
    .addSeparator()
    .addItem('실시간 순위 숨기기', 'hideLive')
    .addItem('실시간 순위 공개', 'showLive')
    .addSeparator()
    .addItem('투표 마감', 'closeVoting')
    .addItem('투표 다시 열기', 'openVoting')
    .addSeparator()
    .addItem('운영본부 코드·PIN 설정', 'setStaffSecrets')
    .addSeparator()
    .addItem('구글폼 만들기 (처음 한 번만)', 'createForm')
    .addItem('기기 질문 추가 (한 번만)', 'addDeviceField')
    .addToUi();
}

function props_() {
  return PropertiesService.getScriptProperties();
}

function spreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(props_().getProperty('SHEET_ID'));
}

function form_() {
  const id = props_().getProperty('FORM_ID');
  if (id) return FormApp.openById(id);
  const linked = spreadsheet_().getSheets().find((s) => s.getFormUrl());
  return FormApp.openByUrl(linked.getFormUrl());
}

function tryRun_(fn) {
  try { fn(); } catch (e) { Logger.log('건너뜀: ' + e.message); }
}

/* ───────────── 운영본부 인증 ───────────── */

// 설정되지 않았으면 null → 운영본부 기능(로그인·관리·기기 제한 예외)이 모두 꺼짐
function staffSecrets_() {
  const p = props_().getProperties();
  const code = String(p.STAFF_CODE || '').trim();
  const pin = String(p.STAFF_PIN || '').trim();
  return code && pin ? { code, pin } : null;
}

// 운영본부 인증 토큰: 코드와 PIN으로 만들어, 둘 중 하나를 바꾸면 기존 토큰은 모두 무효가 됨
function staffToken_() {
  const s = staffSecrets_();
  if (!s) return null;
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, `${s.code}:${s.pin}`, Utilities.Charset.UTF_8);
  return bytes.map((b) => ((b + 256) % 256).toString(16).padStart(2, '0')).join('').slice(0, 32);
}

// 시트 메뉴에서 실행: 운영본부 링크 코드와 PIN 입력
function setStaffSecrets() {
  const ui = SpreadsheetApp.getUi();
  const now = staffSecrets_();
  const current = now ? `\n\n현재 코드: ${now.code} / 현재 PIN: ${now.pin}` : '\n\n(아직 설정되지 않음)';
  const code = ui.prompt('운영본부 코드', `운영본부 링크의 ?staff= 뒤에 붙일 코드를 입력하세요. (예: hq-ab12cd34)${current}`, ui.ButtonSet.OK_CANCEL);
  if (code.getSelectedButton() !== ui.Button.OK) return;
  const pin = ui.prompt('운영본부 PIN', '운영본부 폰에서 입력할 PIN을 입력하세요. (숫자 6자리 권장)', ui.ButtonSet.OK_CANCEL);
  if (pin.getSelectedButton() !== ui.Button.OK) return;

  const c = code.getResponseText().trim();
  const p = pin.getResponseText().trim();
  if (!c || !p) return ui.alert('코드와 PIN을 모두 입력해야 합니다. 저장하지 않았습니다.');
  props_().setProperties({ STAFF_CODE: c, STAFF_PIN: p });
  ui.alert(`저장했습니다.\n\n운영본부 링크: https://leeejunseo.github.io/foodcourt-vote/?staff=${c}\nPIN: ${p}\n\n값을 바꿨다면 운영본부 폰은 다시 로그인해야 합니다.`);
}

/* ───────────── 투표 마감 ─────────────
 * 구글폼의 '응답 받기' 설정만으로는 실제로 막히지 않는 경우가 있어,
 * 마감 기간을 직접 기록하고 그 사이에 들어온 응답은 집계에서 뺌.
 * CLOSED_PERIODS = [[시작ms, 끝ms 또는 null(아직 마감 중)], ...]
 */
function closedPeriods_() {
  try { return JSON.parse(props_().getProperty('CLOSED_PERIODS') || '[]'); } catch (e) { return []; }
}

function isClosed_() {
  const periods = closedPeriods_();
  return periods.length > 0 && periods[periods.length - 1][1] == null;
}

function setClosed_(closed) {
  const periods = closedPeriods_();
  const open = !(periods.length > 0 && periods[periods.length - 1][1] == null);
  if (closed && open) periods.push([Date.now(), null]);
  if (!closed && !open) periods[periods.length - 1][1] = Date.now();
  props_().setProperty('CLOSED_PERIODS', JSON.stringify(periods));
  if (closed) props_().setProperty('LIVE_HIDDEN', '1'); // 마감하면 결과 발표 전까지 순위 비공개
  tryRun_(() => form_().setAcceptingResponses(!closed));
  CacheService.getScriptCache().remove('live');
}

function closeVoting() {
  setClosed_(true);
  SpreadsheetApp.getActive().toast('지금부터 들어오는 응답은 집계되지 않고, 실시간 순위도 숨겨집니다.', '투표 마감');
}

function openVoting() {
  setClosed_(false);
  SpreadsheetApp.getActive().toast('다시 투표를 받습니다.', '투표 다시 열기');
}

/* ───────────── 구글폼 만들기 ───────────── */

function createForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (props_().getProperty('FORM_ID') || ss.getSheets().some((s) => s.getFormUrl())) {
    throw new Error("이미 구글폼이 연결되어 있습니다. '설정' 시트를 확인하세요.");
  }

  // 새 시트의 빈 '시트1'은 '설정' 시트로 재활용
  const sheets = ss.getSheets();
  if (sheets.length === 1 && sheets[0].getLastRow() === 0) sheets[0].setName(SETUP_SHEET);

  const form = FormApp.create(FORM_TITLE)
    .setDescription('투표 웹페이지에서 자동으로 응답이 들어오는 폼입니다.')
    .setConfirmationMessage('투표해 주셔서 감사합니다.');
  FIELDS.forEach((title) => form.addTextItem().setTitle(title).setRequired(title !== '부스2'));
  form.addTextItem().setTitle(DEVICE_FIELD).setHelpText('자동 입력 — 비워두세요');

  // 웹페이지에서 응답을 받으려면 로그인 요구·이메일 수집·1회 제한이 모두 꺼져 있어야 함
  form.setLimitOneResponsePerUser(false);
  tryRun_(() => form.setCollectEmail(false));
  tryRun_(() => form.setRequireLogin(false)); // 학교(Workspace) 계정에만 있는 설정
  tryRun_(() => form.setPublished(true));
  form.setAcceptingResponses(true);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  const prefillUrl = buildPrefill_(form);
  props_().setProperties({ FORM_ID: form.getId(), SHEET_ID: ss.getId() });

  const setup = ss.getSheetByName(SETUP_SHEET) || ss.insertSheet(SETUP_SHEET);
  setup.clear();
  setup.getRange(1, 1, 6, 2).setValues([
    ['푸드코트 투표 설정', ''],
    ['', ''],
    ['웹페이지 연결 링크', prefillUrl],
    ['', '↑ 이 링크를 index.html 의 formPrefillLink 에 붙여넣기'],
    ['구글폼 편집', form.getEditUrl()],
    ['응답 마감 방법', '구글폼 편집 → 응답 탭 → 응답 받기 끄기'],
  ]);
  setup.getRange('A1').setFontSize(14).setFontWeight('bold');
  setup.getRange('A3:A6').setFontWeight('bold');
  setup.getRange('B4').setFontColor('#888888');
  setup.setColumnWidth(1, 140);
  setup.setColumnWidth(2, 700);
  setup.activate();

  Logger.log('웹페이지 연결 링크 (index.html 의 formPrefillLink 에 붙여넣기):\n' + prefillUrl);
  // alert 는 시트 탭에서 확인을 누를 때까지 실행이 멈추므로 사라지는 알림(toast)을 씀
  tryRun_(() => ss.toast("'설정' 시트의 웹페이지 연결 링크를 index.html 에 붙여넣으세요.", '구글폼을 만들었습니다'));
}

// 이미 만든 폼에 '기기' 질문을 추가하고 새 연결 링크를 만듦 (기존 응답·질문은 그대로)
function addDeviceField() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const linked = ss.getSheets().find((s) => s.getFormUrl());
  if (!linked) throw new Error('연결된 구글폼이 없습니다. 먼저 createForm 을 실행하세요.');
  const formId = props_().getProperty('FORM_ID');
  const form = formId ? FormApp.openById(formId) : FormApp.openByUrl(linked.getFormUrl());
  if (form.getItems().some((item) => item.getTitle() === DEVICE_FIELD)) {
    throw new Error("이미 '기기' 질문이 있습니다. '설정' 시트의 링크를 그대로 쓰면 됩니다.");
  }
  form.addTextItem().setTitle(DEVICE_FIELD).setHelpText('자동 입력 — 비워두세요');
  tryRun_(() => form.setPublished(true));

  const prefillUrl = buildPrefill_(form);
  const setup = ss.getSheetByName(SETUP_SHEET);
  if (setup) setup.getRange('B3').setValue(prefillUrl);

  Logger.log('새 웹페이지 연결 링크 (index.html 의 formPrefillLink 에 붙여넣기):\n' + prefillUrl);
  tryRun_(() => ss.toast("'설정' 시트의 새 웹페이지 연결 링크를 index.html 에 붙여넣으세요.", "'기기' 질문을 추가했습니다"));
}

// 각 칸에 질문 제목을 미리 채운 링크 → 웹페이지가 이걸 보고 entry ID를 찾음
function buildPrefill_(form) {
  const prefill = form.createResponse();
  form.getItems().forEach((item) => prefill.withItemResponse(item.asTextItem().createResponse(item.getTitle())));
  return prefill.toPrefilledUrl();
}

/* ───────────── 집계 계산 ───────────── */

function computeTally_(ss) {
  const src = ss.getSheets().find((s) => s.getFormUrl());
  if (!src) throw new Error('응답 시트를 찾을 수 없습니다. 먼저 createForm 을 실행하세요.');
  const values = src.getDataRange().getValues();
  const header = values[0].map((h) => String(h).trim());

  const col = (name) => {
    const i = header.findIndex((h) => h === name || h.includes(name));
    if (i < 0) throw new Error(`응답 시트에서 '${name}' 열을 찾을 수 없습니다. 구글폼 질문 제목을 확인하세요.`);
    return i;
  };
  const c = { gisu: col('기수'), company: col('중대'), name: col('이름'), b1: col('부스1'), b2: col('부스2') };
  const dev = header.indexOf(DEVICE_FIELD); // 없으면 기기 기준 중복 처리는 건너뜀

  const norm = (v) => String(v).replace(/\s+/g, '').toLowerCase();
  const rows = values.slice(1).filter((r) => r.some((v) => v !== ''));
  rows.sort((a, b) => new Date(a[0]) - new Date(b[0])); // 타임스탬프 순

  // 마감 중에 들어온 응답은 무효
  const periods = closedPeriods_();
  const late = [];
  const inTime = rows.filter((r) => {
    const t = new Date(r[0]).getTime();
    const closed = periods.some(([start, end]) => t >= start && (end == null || t < end));
    if (closed) late.push(r);
    return !closed;
  });

  // 최신 응답부터 보면서, 같은 사람이나 같은 기기에서 이미 더 늦은 표가 있으면 무효
  const token = staffToken_();
  const staffDevice = token ? `staff:${token}` : null;
  const seenPeople = new Set();
  const seenDevices = new Set();
  const valid = [];
  const replaced = [];
  for (const r of [...inTime].reverse()) {
    if (!norm(r[c.name])) continue;
    const person = [r[c.gisu], r[c.company], r[c.name]].map(norm).join('|');
    const device = dev < 0 ? '' : String(r[dev]).trim();
    const deviceKey = staffDevice && device === staffDevice ? '' : device; // 인증된 운영본부 폰은 기기 제한 없음
    if (seenPeople.has(person) || (deviceKey && seenDevices.has(deviceKey))) {
      replaced.push(r);
      continue;
    }
    seenPeople.add(person);
    if (deviceKey) seenDevices.add(deviceKey);
    valid.push(r);
  }
  replaced.reverse();

  const counts = new Map();
  const solo = new Map(); // 단독 지지: 그 부스 한 곳만 고른 투표 수 (동점일 때 기준)
  let voters = 0;
  for (const r of valid) {
    const picks = [...new Set([r[c.b1], r[c.b2]].map((v) => String(v).trim()).filter(Boolean))];
    if (!picks.length) continue;
    voters++;
    picks.forEach((b) => counts.set(b, (counts.get(b) || 0) + 1));
    if (picks.length === 1) solo.set(picks[0], (solo.get(picks[0]) || 0) + 1);
  }

  return { header, rows, replaced, late, voters, counts, solo };
}

/* ───────────── 실시간 순위 (웹 앱) ───────────── */

function doGet() {
  const cache = CacheService.getScriptCache();
  let json = cache.get('live');
  if (!json) {
    json = JSON.stringify(liveData_());
    cache.put('live', json, CACHE_SECONDS);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function liveData_() {
  const updatedAt = new Date().toISOString();
  const t = computeTally_(spreadsheet_());
  const closed = isClosed_(); // 마감되면 웹페이지가 투표 화면을 닫음
  // 숨김 중에도 참여자 수는 보여줌 (순위·득표는 비공개)
  if (props_().getProperty('LIVE_HIDDEN') === '1') return { hidden: true, closed, updatedAt, voters: t.voters };
  return { hidden: false, closed, updatedAt, voters: t.voters, counts: toObject_(t.counts), solo: toObject_(t.solo) };
}

function toObject_(map) {
  const obj = {};
  map.forEach((value, key) => { obj[key] = value; });
  return obj;
}

function hideLive() { setLiveHidden_(true); }
function showLive() { setLiveHidden_(false); }

function setLiveHidden_(hidden) {
  props_().setProperties({ LIVE_HIDDEN: hidden ? '1' : '0', SHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId() });
  CacheService.getScriptCache().remove('live');
  SpreadsheetApp.getActive().toast(
    hidden ? '웹페이지에서 순위가 가려집니다.' : '웹페이지에 순위가 다시 표시됩니다.',
    hidden ? '실시간 순위 숨김' : '실시간 순위 공개'
  );
}

/* ───────────── 운영본부 관리 (운영본부 폰의 웹페이지에서 호출) ───────────── */

// 로그인: { action: 'login', code, pin } → { token }
// 그 외:  { action, token }  action = status | hide | show | close | open
function doPost(e) {
  const out = (obj) => ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  let req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) { /* 잘못된 요청은 아래에서 거절됨 */ }

  const secrets = staffSecrets_();
  if (!secrets) return out({ ok: false, error: '운영본부 코드·PIN이 아직 설정되지 않았어요. 시트 메뉴에서 설정해 주세요.' });

  if (req.action === 'login') {
    // 링크(코드)가 틀린 요청은 잠금 횟수에 넣지 않음 → 외부인이 운영본부를 잠글 수 없음
    if (req.code !== secrets.code) return out({ ok: false, error: '운영본부 링크가 올바르지 않아요.' });
    const cache = CacheService.getScriptCache();
    const fails = Number(cache.get('staff-fails') || 0);
    if (fails >= MAX_PIN_TRIES) return out({ ok: false, error: 'PIN을 여러 번 틀려서 10분 동안 잠겼어요.' });
    if (String(req.pin) !== secrets.pin) {
      cache.put('staff-fails', String(fails + 1), 600);
      const left = MAX_PIN_TRIES - fails - 1;
      return out({ ok: false, error: left > 0 ? `PIN이 올바르지 않아요. (남은 시도 ${left}회)` : 'PIN을 여러 번 틀려서 10분 동안 잠겼어요.' });
    }
    cache.remove('staff-fails');
    return out({ ok: true, token: staffToken_() });
  }

  if (req.token !== staffToken_()) {
    return out({ ok: false, auth: true, error: '운영본부 인증이 필요해요. 운영본부 링크로 다시 들어와 PIN을 입력해 주세요.' });
  }

  if (req.action === 'hide' || req.action === 'show') props_().setProperty('LIVE_HIDDEN', req.action === 'hide' ? '1' : '0');
  if (req.action === 'close' || req.action === 'open') setClosed_(req.action === 'close');
  if (req.action !== 'status') CacheService.getScriptCache().remove('live');
  // 구글폼의 '응답 받기' 설정을 우리 마감 상태와 맞춰 둠
  if (req.action === 'status') tryRun_(() => {
    const form = form_();
    if (form.isAcceptingResponses() === isClosed_()) form.setAcceptingResponses(!isClosed_());
  });

  // 운영본부에는 순위를 숨긴 상태에서도 전체 득표를 보여줌
  const t = computeTally_(spreadsheet_());
  return out({
    ok: true,
    hidden: props_().getProperty('LIVE_HIDDEN') === '1',
    accepting: !isClosed_(),
    late: t.late.length,
    voters: t.voters,
    counts: toObject_(t.counts),
    solo: toObject_(t.solo),
    updatedAt: new Date().toISOString(),
  });
}

/* ───────────── 최종 집계 (시트에 기록) ───────────── */

function tally() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  props_().setProperty('SHEET_ID', ss.getId());
  const t = computeTally_(ss);

  // 순위: 득표수 → 단독 지지 순. 둘 다 같으면 같은 순위(1, 1, 3 …)로 두고 동점 표시
  const ranked = [...t.counts.entries()]
    .map(([booth, votes]) => ({ booth, votes, solo: t.solo.get(booth) || 0 }))
    .sort((a, b) => b.votes - a.votes || b.solo - a.solo || a.booth.localeCompare(b.booth, 'ko'));
  ranked.forEach((r) => {
    r.rank = 1 + ranked.filter((o) => o.votes > r.votes || (o.votes === r.votes && o.solo > r.solo)).length;
    const tied = ranked.filter((o) => o.votes === r.votes && o.solo === r.solo).length > 1;
    const award = AWARDS[r.rank - 1] || '';
    r.award = award && tied ? `${award} (동점 — 운영본부 결정 필요)` : award;
  });

  // 집계 시트 작성
  const out = ss.getSheetByName(RESULT_SHEET) || ss.insertSheet(RESULT_SHEET);
  out.clear();
  const summary = [
    ['푸드코트 투표 집계', ''],
    ['집계 시각', new Date()],
    ['전체 응답 수', t.rows.length],
    ['유효 투표자 수', t.voters],
    ['무효가 된 이전 투표', t.replaced.length],
    ['마감 중 들어온 응답 (무효)', t.late.length],
  ];
  out.getRange(1, 1, summary.length, 2).setValues(summary);
  out.getRange('A1').setFontSize(14).setFontWeight('bold');
  out.getRange('B2').setNumberFormat('yyyy-mm-dd hh:mm:ss');

  const tableTop = summary.length + 2;
  const table = [['순위', '부스', '득표수', '단독 지지', '수상']].concat(ranked.map((r) => [r.rank, r.booth, r.votes, r.solo, r.award]));
  out.getRange(tableTop, 1, table.length, 5).setValues(table);
  out.getRange(tableTop, 1, 1, 5).setFontWeight('bold').setBackground('#0b2a5b').setFontColor('#ffffff');
  ranked.forEach((r, i) => {
    const bg = { 1: '#fdf3d0', 2: '#eef1f4', 3: '#f6e6d8' }[r.rank];
    if (bg) out.getRange(tableTop + 1 + i, 1, 1, 5).setBackground(bg).setFontWeight('bold');
  });
  out.autoResizeColumns(1, 5);

  // 이전투표 시트 작성 (같은 사람이 바꾼 표 + 같은 폰에서 다른 사람이 덮어쓴 표)
  const oldOut = ss.getSheetByName(OLD_SHEET) || ss.insertSheet(OLD_SHEET);
  oldOut.clear();
  oldOut.getRange(1, 1, 1, t.header.length).setValues([t.header]).setFontWeight('bold');
  if (t.replaced.length) oldOut.getRange(2, 1, t.replaced.length, t.header.length).setValues(t.replaced);

  // 마감후응답 시트 작성 (마감 중에 들어와 집계에서 뺀 응답)
  const lateOut = ss.getSheetByName(LATE_SHEET) || ss.insertSheet(LATE_SHEET);
  lateOut.clear();
  lateOut.getRange(1, 1, 1, t.header.length).setValues([t.header]).setFontWeight('bold');
  if (t.late.length) lateOut.getRange(2, 1, t.late.length, t.header.length).setValues(t.late);

  out.activate();
  ss.toast(`유효 투표자 ${t.voters}명, 무효 ${t.replaced.length + t.late.length}건`, '집계 완료');
}
