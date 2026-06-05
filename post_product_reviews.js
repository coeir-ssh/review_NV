/**
 * 코에르 리뷰 미답변 자동 답변 등록
 *
 * [흐름]
 *  1. 셀러센터 리뷰 페이지 접속
 *  2. 필터: 오늘 작성 + 답글미등록
 *  3. 환불검토 대상 판단 (1~2점 자동, 3점은 Claude 판단)
 *  4. 환불검토 대상이 아닌 경우 Claude API로 개인화 답변 생성 및 등록
 *  5. 결과 Excel 저장
 *  6. 요약 JSON 저장 (Slack DM용)
 *
 * 실행: node post_product_reviews.js
 */

const puppeteer  = require('puppeteer');
const Anthropic  = require('@anthropic-ai/sdk');
const readline   = require('readline');
const path       = require('path');
const fs         = require('fs');
const https      = require('https');
const cfg        = require('./config');
const { getMasterProductName, inferOption } = require('./master_product_names');

// 리뷰순위 표시 헬퍼: 양수=N위 / -2=100위 밖 / 그 외=미확인
function formatReviewPosition(p) {
  if (p > 0) return `${p}위`;
  if (p === -2) return '100위 밖';
  return '미확인';
}

// 표시용 옵션: AG Grid 에서 잡힌 buyer 옵션 우선, 없으면 풀네임에서 추출한 변형명
function resolveDisplayOption(r) {
  return (r.optionName && r.optionName.trim()) || inferOption(r.productName) || '';
}

// 리뷰 순위 정렬 키: 1~N → N, 100위 밖(-2) → 9999, 미확인(0/-1) → 99999
function rankSortKey(p) {
  if (typeof p === 'number' && p > 0) return p;
  if (p === -2) return 9999;
  return 99999;
}
function sortByRank(list) {
  return [...list].sort((a, b) => rankSortKey(a.reviewPosition) - rankSortKey(b.reviewPosition));
}

// 판단 근거 표시 여부:
// - 환불검토는 항상 표시
// - 답변이라도 confidence 가 낮으면(< 90) 애매한 케이스로 보고 표시
const REASON_CONFIDENCE_THRESHOLD = 90;
function shouldShowJudgeReason(r) {
  if (!r.judgeReason) return false;
  if (r.judgeLabel === '환불검토') return true;
  if (r.judgeConfidence != null && r.judgeConfidence < REASON_CONFIDENCE_THRESHOLD) return true;
  return false;
}
const {
  Document, Packer, Paragraph, TextRun,
  AlignmentType, BorderStyle, HeadingLevel,
  ShadingType,
} = require('docx');

// ─────────────────────────────────────────────────────────
// 설정
// ─────────────────────────────────────────────────────────
const CONFIG = {
  MAX_REVIEWS:  500,                 // 처리할 최대 리뷰 수 (사실상 무제한)
  MODEL:        'claude-sonnet-4-5', // Sonnet 모델
  sessionFile:  path.join(__dirname, '.seller_session.json'),
  reviewUrl:    'https://sell.smartstore.naver.com/#/review/search',
  headless:     true,
  replyDelay:   3000,
};

const log   = msg => console.log(msg);
const sleep = ms  => new Promise(r => setTimeout(r, ms));

// ─────────────────────────────────────────────────────────
// 진단 덤프: 실패 시점의 스크린샷 + HTML + URL을 logs/에 저장
// ─────────────────────────────────────────────────────────
async function dumpDiagnostics(page, label = 'error') {
  try {
    const logsDir = path.join(__dirname, 'logs');
    if (!fs.existsSync(logsDir)) fs.mkdirSync(logsDir, { recursive: true });
    const now = new Date();
    const ts  = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}_${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}${String(now.getSeconds()).padStart(2,'0')}`;
    const base = path.join(logsDir, `diag_${ts}_${label}`);
    try {
      await page.screenshot({ path: base + '.png', fullPage: true });
      log(`  [진단] 스크린샷: ${base}.png`);
    } catch (e) { log(`  [진단] 스크린샷 실패: ${e.message}`); }
    try {
      const html = await page.content();
      fs.writeFileSync(base + '.html', html);
      log(`  [진단] HTML 덤프: ${base}.html (${html.length}자)`);
    } catch (e) { log(`  [진단] HTML 덤프 실패: ${e.message}`); }
    try {
      const url = page.url();
      const title = await page.title().catch(() => '');
      const bodyText = await page.evaluate(() => document.body ? document.body.innerText.substring(0, 500) : '').catch(() => '');
      log(`  [진단] URL: ${url}`);
      log(`  [진단] 제목: ${title}`);
      log(`  [진단] 본문 첫 500자: ${bodyText.replace(/\n/g, ' ')}`);
    } catch (e) {}
  } catch (e) { log(`  [진단] 전체 실패: ${e.message}`); }
}

const ai = new Anthropic({ apiKey: cfg.ANTHROPIC_API_KEY });

// ─────────────────────────────────────────────────────────
// 영업일 가드 (공공데이터포털 한국천문연구원 특일정보 API + 로컬 캐시)
// ─────────────────────────────────────────────────────────
const HOLIDAY_CACHE_FILE = path.join(__dirname, 'holidays_cache.json');
const HOLIDAY_CACHE_TTL_DAYS = 30; // 캐시 만료 주기 (임시공휴일 신규 발표 반영용)

// 공공데이터포털에서 특정 연도의 공휴일 가져오기
function fetchHolidaysFromAPI(year) {
  return new Promise((resolve, reject) => {
    const key = encodeURIComponent(cfg.DATA_GO_KR_KEY);
    const url = `http://apis.data.go.kr/B090041/openapi/service/SpcdeInfoService/getRestDeInfo?serviceKey=${key}&solYear=${year}&numOfRows=100&_type=json`;
    const http = require('http');
    http.get(url, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const j = JSON.parse(data);
          const code = j.response?.header?.resultCode;
          if (code !== '00') return reject(new Error(`API 오류: ${code} ${j.response?.header?.resultMsg}`));
          let items = j.response?.body?.items?.item || [];
          if (!Array.isArray(items)) items = [items]; // 1건일 때 객체로 옴
          const dates = items
            .filter(it => it.isHoliday === 'Y')
            .map(it => {
              const s = String(it.locdate);
              return `${s.slice(0,4)}-${s.slice(4,6)}-${s.slice(6,8)}`;
            });
          resolve(dates);
        } catch (e) { reject(e); }
      });
    }).on('error', reject);
  });
}

// 캐시 우선 + 만료/누락 시 API 호출 → 통합 공휴일 Set 반환
async function getKoreanHolidays() {
  const now = new Date();
  const thisYear = now.getFullYear();
  const nextYear = thisYear + 1;

  let cache = null;
  try {
    if (fs.existsSync(HOLIDAY_CACHE_FILE)) cache = JSON.parse(fs.readFileSync(HOLIDAY_CACHE_FILE, 'utf8'));
  } catch (e) { cache = null; }

  const ageDays = cache ? (Date.now() - new Date(cache.fetchedAt).getTime()) / 86400000 : Infinity;
  const hasCurrentYear = cache && Array.isArray(cache.years?.[thisYear]);
  const hasNextYear    = cache && Array.isArray(cache.years?.[nextYear]);
  const cacheValid     = hasCurrentYear && hasNextYear && ageDays < HOLIDAY_CACHE_TTL_DAYS;

  if (cacheValid) {
    return new Set([...cache.years[thisYear], ...cache.years[nextYear]]);
  }

  // 캐시 미존재/만료 → API 갱신
  try {
    const [thisYearList, nextYearList] = await Promise.all([
      fetchHolidaysFromAPI(thisYear),
      fetchHolidaysFromAPI(nextYear),
    ]);
    const fresh = {
      fetchedAt: new Date().toISOString(),
      years: { [thisYear]: thisYearList, [nextYear]: nextYearList },
    };
    fs.writeFileSync(HOLIDAY_CACHE_FILE, JSON.stringify(fresh, null, 2));
    log(`[공휴일] API 갱신 완료: ${thisYear}년 ${thisYearList.length}건, ${nextYear}년 ${nextYearList.length}건`);
    return new Set([...thisYearList, ...nextYearList]);
  } catch (e) {
    log(`[공휴일] API 호출 실패: ${e.message}`);
    if (cache) {
      log(`[공휴일] 만료된 캐시(${Math.floor(ageDays)}일 경과) 폴백 사용`);
      const all = [];
      for (const y of Object.keys(cache.years || {})) all.push(...cache.years[y]);
      return new Set(all);
    }
    log(`[공휴일] 캐시도 없음 → 주말만 가드`);
    return new Set();
  }
}

async function shouldSkipToday() {
  const now = new Date();
  const dow = now.getDay(); // 0=일, 6=토
  if (dow === 0 || dow === 6) return { skip: true, reason: '주말' };

  const ymd = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
  const holidays = await getKoreanHolidays();
  if (holidays.has(ymd)) return { skip: true, reason: `공휴일 (${ymd})` };
  return { skip: false };
}

// 스케줄러에서 호출할 때만 가드 동작 (수동 실행은 영향 없음)
async function runScheduleGuard() {
  if (!process.argv.includes('--scheduled')) return;
  const { skip, reason } = await shouldSkipToday();
  if (skip) {
    console.log(`[스케줄러] ${reason} → 자동 답변 작업 건너뜀. 정상 종료.`);
    process.exit(0);
  }
}

// ─────────────────────────────────────────────────────────
// API 토큰 사용량 / 비용 추적
// ─────────────────────────────────────────────────────────
// Claude Sonnet 4.5 단가 (USD / 1M tokens)
const PRICING = {
  input:        3.00,
  output:      15.00,
  cacheWrite:   3.75,
  cacheRead:    0.30,
};
const USD_TO_KRW = 1380; // 대략적인 환율 (표기용)

const tokenUsage = {
  input:       0,
  output:      0,
  cacheWrite:  0,
  cacheRead:   0,
  calls:       0,
};

function accumulateUsage(u) {
  if (!u) return;
  tokenUsage.calls      += 1;
  tokenUsage.input      += u.input_tokens || 0;
  tokenUsage.output     += u.output_tokens || 0;
  tokenUsage.cacheWrite += u.cache_creation_input_tokens || 0;
  tokenUsage.cacheRead  += u.cache_read_input_tokens || 0;
}

function getUsageSummary() {
  const costUSD =
    (tokenUsage.input      / 1_000_000) * PRICING.input +
    (tokenUsage.output     / 1_000_000) * PRICING.output +
    (tokenUsage.cacheWrite / 1_000_000) * PRICING.cacheWrite +
    (tokenUsage.cacheRead  / 1_000_000) * PRICING.cacheRead;
  const costKRW = costUSD * USD_TO_KRW;
  const totalTokens = tokenUsage.input + tokenUsage.output + tokenUsage.cacheWrite + tokenUsage.cacheRead;
  return {
    ...tokenUsage,
    totalTokens,
    costUSD: Math.round(costUSD * 10000) / 10000,   // 0.0001 USD
    costKRW: Math.round(costKRW),
  };
}

// ─────────────────────────────────────────────────────────
// 세션
// ─────────────────────────────────────────────────────────
function saveSession(c) { fs.writeFileSync(CONFIG.sessionFile, JSON.stringify(c)); }
function loadSession() {
  try { if (fs.existsSync(CONFIG.sessionFile)) return JSON.parse(fs.readFileSync(CONFIG.sessionFile)); }
  catch(e) {}
  return null;
}

function askQuestion(prompt) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise(resolve => rl.question(prompt, ans => { rl.close(); resolve(ans.trim()); }));
}

// ─────────────────────────────────────────────────────────
// 로그인 확인 / 처리
// ─────────────────────────────────────────────────────────
async function isLoggedIn(page) {
  try {
    await page.waitForFunction(() => {
      const t = document.body.innerText;
      return t.includes('상품관리') || t.includes('리뷰관리') || t.includes('주문관리') ||
             !!document.querySelector('input[placeholder="아이디 또는 이메일 주소"]');
    }, { timeout: 12000 });
  } catch(e) {}
  return page.evaluate(() => {
    const t = document.body.innerText;
    return t.includes('상품관리') || t.includes('리뷰관리') || t.includes('주문관리');
  });
}

async function loginToSellerCenter(page) {
  const saved = loadSession();
  if (saved) {
    log('  저장된 세션 확인 중...');
    await page.setCookie(...saved);
    await page.goto('https://sell.smartstore.naver.com/', { waitUntil: 'networkidle2', timeout: 30000 });
    if (await isLoggedIn(page)) { log('  ✓ 세션 로그인 성공'); return; }
    log('  세션 만료 → 재로그인');
    try { fs.unlinkSync(CONFIG.sessionFile); } catch(e) {}
  }

  await page.goto('https://accounts.commerce.naver.com/login', { waitUntil: 'networkidle2', timeout: 20000 });
  await sleep(1500);

  const idF = await page.$('input[placeholder="아이디 또는 이메일 주소"]');
  const pwF = await page.$('input[type="password"]');
  if (idF && pwF) {
    log('  ID/PW 입력 중...');
    await idF.click({ clickCount: 3 }); await idF.type(cfg.SELLER_ID, { delay: 80 });
    await sleep(300);
    await pwF.click({ clickCount: 3 }); await pwF.type(cfg.SELLER_PW, { delay: 80 });
    await sleep(300);
    await page.evaluate(() => {
      const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '로그인');
      if (btn) btn.click();
    });
    await page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 20000 }).catch(() => {});
    await sleep(2000);
  }

  if (!(await isLoggedIn(page))) {
    // 무인 환경(오전 8시 자동 실행)에서 askQuestion()으로 블로킹하면 영구 멈춤
    // → 에러 throw로 즉시 종료, main().catch()가 Slack 알림 + 락파일 정리 담당
    if (process.argv.includes('--scheduled')) {
      throw new Error('세션 만료: 수동 재로그인 필요. 터미널에서 한 번 직접 실행해서 세션을 갱신해 주세요.');
    }
    // 수동 실행 시에는 기존처럼 대기
    log('  추가 인증이 필요합니다. 브라우저에서 인증 완료 후 Enter를 누르세요.');
    await askQuestion('  인증 완료 후 Enter ▶ ');
    await sleep(2000);
  }
  saveSession(await page.cookies());
  log('  ✓ 로그인 완료');
}

// ─────────────────────────────────────────────────────────
// 스토어 확인 / 전환 (계정에 코에르 + 자유생활 여러 스토어가 있음)
// 다른 스토어에 들어가면 그 스토어의 인증 모달 등이 코에르 작업을 막음
// ─────────────────────────────────────────────────────────
const TARGET_STORE_NAME = '코에르';

// 좌측 상단 / 헤더에서 현재 접속 중인 스토어명 읽기
async function readCurrentStoreName(page) {
  return page.evaluate(() => {
    // 스토어 이름이 들어갈 만한 좁은 영역 후보들
    const sels = [
      'aside [class*="store"]', 'aside [class*="profile"]', 'aside [class*="name"]',
      'header [class*="store"]', 'header [class*="profile"]',
      '[class*="GnbStore"]', '[class*="storeName"]', '[class*="store_name"]',
    ];
    const elems = new Set();
    for (const s of sels) document.querySelectorAll(s).forEach(e => elems.add(e));
    for (const el of elems) {
      const t = (el.textContent || '').trim().replace(/\s+/g, ' ');
      if (!t || t.length > 30) continue;
      if (/코에르|자\s*유\s*생\s*활/.test(t)) return t;
    }
    // 폴백: 페이지 상단(첫 1000자) 텍스트에서 후보 검출
    const head = (document.body?.innerText || '').slice(0, 1500);
    if (/자\s*유\s*생\s*활/.test(head)) return '자 유 생 활';
    if (/코에르/.test(head))           return '코에르';
    return null;
  });
}

// 화면을 덮는 알림/안내 모달들 자동 닫기
async function dismissBlockingModals(page) {
  const closed = await page.evaluate(() => {
    let n = 0;
    // 1. 명시적인 "닫기/취소/나중에/다음에" 버튼
    const closeBtns = Array.from(document.querySelectorAll('button'))
      .filter(b => /^(닫기|취소|나중에|다음에|건너뛰기)$/.test((b.textContent || '').trim()));
    for (const b of closeBtns.slice(0, 5)) {
      try { b.click(); n++; } catch(e) {}
    }
    // 2. × 닫기 아이콘 (모달 헤더 내부)
    const modals = document.querySelectorAll('[class*="modal"], [class*="popup"], [class*="dialog"], [class*="layer"]');
    for (const m of modals) {
      const x = m.querySelector('button[aria-label*="닫기"], button[class*="close"], [class*="closeBtn"], button[title="닫기"]');
      if (x) { try { x.click(); n++; } catch(e) {} }
    }
    return n;
  });
  if (closed > 0) { log(`  [모달] ${closed}개 닫음`); await sleep(700); }
}

// "스토어 이동" 드롭다운 → 코에르 라디오 → 적용
async function switchToCoeirStore(page) {
  // 1) "스토어 이동" 드롭다운 / 버튼 열기
  const opened = await page.evaluate(() => {
    const cands = Array.from(document.querySelectorAll('button, a, span, div'))
      .filter(el => /스토어\s*이동/.test((el.textContent || '').trim()));
    // 너무 큰 컨테이너 배제 — 작은 헤더 버튼 우선
    cands.sort((a, b) => {
      const ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect();
      return (ra.width * ra.height) - (rb.width * rb.height);
    });
    for (const el of cands) {
      const r = el.getBoundingClientRect();
      if (r.width > 5 && r.width < 250 && r.height < 80) {
        try { el.click(); return true; } catch(e) {}
      }
    }
    if (cands[0]) { try { cands[0].click(); return true; } catch(e) {} }
    return false;
  });
  if (!opened) { log('  [경고] "스토어 이동" 버튼 못 찾음'); return false; }
  await sleep(1500);

  // 2) 모달에서 코에르 항목 / 라디오 클릭
  const picked = await page.evaluate(() => {
    // 모든 라디오를 훑어 그 라벨에 코에르가 포함된 것 우선
    const radios = Array.from(document.querySelectorAll('input[type="radio"]'));
    for (const r of radios) {
      const lbl = r.closest('label') || (r.id && document.querySelector(`label[for="${r.id}"]`));
      const txt = ((lbl?.textContent) || r.parentElement?.textContent || '').trim();
      if (txt.includes('코에르') && !txt.includes('자 유 생 활') && !txt.includes('자유생활')) {
        try { r.click(); return 'radio'; } catch(e) {}
      }
    }
    // 라디오 못 찾으면 텍스트가 코에르인 클릭 가능 요소
    const all = Array.from(document.querySelectorAll('label, li, button, [role="button"], div'));
    for (const el of all) {
      const t = (el.textContent || '').trim();
      if (!t.includes('코에르')) continue;
      if (t.length > 60) continue;
      try { el.click(); return 'text'; } catch(e) {}
    }
    return false;
  });
  if (!picked) { log('  [경고] 코에르 항목 못 찾음'); return false; }
  log(`  [스토어 전환] 코에르 항목 선택 (${picked})`);
  await sleep(700);

  // 3) "확인/적용/이동/접속" 버튼
  await page.evaluate(() => {
    const btn = Array.from(document.querySelectorAll('button'))
      .find(b => /^(확인|적용|이동|선택|접속|접속하기|이동하기)$/.test((b.textContent || '').trim()));
    if (btn) btn.click();
  });

  // 4) 페이지가 코에르로 갱신될 때까지 대기 (헤더 텍스트 포함 여부로 확인)
  await page.waitForFunction(() => {
    const head = (document.body?.innerText || '').slice(0, 2000);
    return head.includes('코에르') && !head.includes('자 유 생 활') && !head.includes('자유생활');
  }, { timeout: 20000 }).catch(() => {});
  await sleep(1500);
  return true;
}

// 메인 진입점: 로그인 후 코에르 스토어인지 확인하고 아니면 전환
async function ensureCoeirStore(page) {
  log('[스토어 확인] 현재 스토어 확인 중...');
  await dismissBlockingModals(page);
  await sleep(500);

  let cur = await readCurrentStoreName(page);
  log(`  현재 스토어: ${cur || '(미확인)'}`);
  if (cur && cur.includes(TARGET_STORE_NAME)) {
    log(`  ✓ 이미 ${TARGET_STORE_NAME} 스토어 접속 중`);
    return;
  }

  log(`  → ${TARGET_STORE_NAME}로 전환 시도...`);
  const ok = await switchToCoeirStore(page);
  if (!ok) {
    await dumpDiagnostics(page, 'store_switch_failed');
    throw new Error(`스토어 전환 실패: 현재 "${cur}" → 목표 "${TARGET_STORE_NAME}"`);
  }

  // 전환 후 잔여 모달 다시 닫기
  await dismissBlockingModals(page);
  await sleep(400);

  cur = await readCurrentStoreName(page);
  log(`  전환 후 스토어: ${cur || '(미확인)'}`);
  if (!cur || !cur.includes(TARGET_STORE_NAME)) {
    await dumpDiagnostics(page, 'store_switch_verify_failed');
    throw new Error(`스토어 전환 검증 실패: 전환 후에도 "${cur}" — 수동 확인 필요`);
  }
  log(`  ✓ ${TARGET_STORE_NAME} 스토어로 전환 완료`);

  // 전환된 세션을 캐시에 저장 (다음 실행 때 바로 코에르로 들어가도록)
  try { saveSession(await page.cookies()); } catch(e) {}
}

// ─────────────────────────────────────────────────────────
// AG Grid 전체 선택 해제
// ─────────────────────────────────────────────────────────
async function deselectAllRows(page) {
  // ElementHandle.click() — 요소 중심으로 실제 마우스 이벤트 발생 (검증된 방식)
  const rowElements = await page.$$('.ag-pinned-left-cols-container .ag-row-selected');
  for (const rowEl of rowElements) {
    try {
      const cbHandle = await rowEl.$('input[type="checkbox"]');
      if (cbHandle) {
        await cbHandle.evaluate(el => el.scrollIntoView({ behavior: 'instant', block: 'center' }));
        await sleep(100);
        await cbHandle.click();
      } else {
        await rowEl.click();
      }
      await sleep(250);
    } catch(e) {}
  }
  await sleep(300);
}

// 현재 선택된 행 수 확인 (ag-row-selected 기준 — AG Grid 실제 선택 상태)
async function getCheckedCount(page) {
  return page.evaluate(() => {
    const bySelected = document.querySelectorAll(
      '.ag-pinned-left-cols-container .ag-row-selected'
    ).length;
    const byChecked = document.querySelectorAll(
      '.ag-pinned-left-cols-container input[type="checkbox"]:checked'
    ).length;
    return Math.max(bySelected, byChecked);
  });
}

// ─────────────────────────────────────────────────────────
// 체크박스 클릭 (row-index 기반)
// ─────────────────────────────────────────────────────────
async function clickRowCheckbox(page, rowIndex) {
  const handle = await page.evaluateHandle((idx) => {
    const pinnedRows = Array.from(document.querySelectorAll('.ag-pinned-left-cols-container .ag-row'));
    const target = pinnedRows.find(r => r.getAttribute('row-index') === idx);
    if (target) {
      return target.querySelector('input[type="checkbox"]') ||
             target.querySelector('.ag-cell') || null;
    }
    const centerRows = Array.from(document.querySelectorAll('.ag-center-cols-container .ag-row'));
    return centerRows.find(r => r.getAttribute('row-index') === idx) || null;
  }, rowIndex);

  const el = handle.asElement();
  if (!el) return false;
  try {
    await page.evaluate(e => e.scrollIntoView({ behavior: 'instant', block: 'center' }), el);
    await sleep(200);
    await el.click();
    return true;
  } catch(e) {
    return page.evaluate((idx) => {
      const rows = Array.from(document.querySelectorAll('.ag-center-cols-container .ag-row'));
      const r = rows.find(r => r.getAttribute('row-index') === idx);
      if (r) { r.click(); return true; }
      return false;
    }, rowIndex);
  }
}

// ─────────────────────────────────────────────────────────
// 답글작성 버튼 클릭
// ─────────────────────────────────────────────────────────
async function clickReplyButton(page) {
  const handle = await page.evaluateHandle(() => {
    const btns = Array.from(document.querySelectorAll('button'));
    return btns.find(b =>
      (b.textContent.trim() === '답글작성' || b.textContent.trim() === '답글 작성') &&
      !b.disabled && b.getAttribute('aria-disabled') !== 'true'
    ) || null;
  });
  const el = handle.asElement();
  if (!el) return false;
  try {
    await page.evaluate(e => e.scrollIntoView({ behavior: 'instant', block: 'center' }), el);
    await sleep(200);
    await el.click();
    return true;
  } catch(e) {
    return page.evaluate(() => {
      const btn = Array.from(document.querySelectorAll('button'))
        .find(b => (b.textContent.trim() === '답글작성' || b.textContent.trim() === '답글 작성') && !b.disabled);
      if (btn) { btn.click(); return true; }
      return false;
    });
  }
}

// ─────────────────────────────────────────────────────────
// 모달 textarea 입력
// ─────────────────────────────────────────────────────────
async function typeReplyText(page, text) {
  let el = null;
  for (let i = 0; i < 15; i++) {
    const handle = await page.evaluateHandle(() => {
      const MODAL_SELS = [
        '.seller-layer-modal', '.modal.show', '.modal.fade',
        '[role="dialog"]', 'mat-dialog-container', '.layer-popup',
      ];
      for (const sel of MODAL_SELS) {
        const modal = document.querySelector(sel);
        if (modal) {
          const ta = modal.querySelector('textarea:not([readonly]):not([disabled])');
          if (ta) return ta;
        }
      }
      const all = document.querySelectorAll('textarea:not([readonly]):not([disabled])');
      for (const ta of all) {
        const r = ta.getBoundingClientRect();
        if (r.width > 50 && r.height > 20) return ta;
      }
      return null;
    });
    el = handle.asElement();
    if (el) break;
    await sleep(400);
  }
  if (!el) return false;

  await page.evaluate(ta => {
    ta.scrollIntoView({ behavior: 'instant', block: 'center' });
    ta.focus(); ta.select(); ta.value = '';
  }, el);
  await sleep(300);

  await page.keyboard.down('Control');
  await page.keyboard.press('a');
  await page.keyboard.up('Control');
  await page.keyboard.press('Delete');
  await sleep(150);

  await page.keyboard.type(text, { delay: 10 });
  await page.evaluate(e => {
    e.dispatchEvent(new Event('input',  { bubbles: true }));
    e.dispatchEvent(new Event('change', { bubbles: true }));
  }, el);
  return true;
}

// ─────────────────────────────────────────────────────────
// 등록 버튼 클릭
// ─────────────────────────────────────────────────────────
async function clickSubmit(page) {
  return page.evaluate(() => {
    const LABELS = ['등록', '저장', '확인', '작성완료', '등록하기'];
    const containers = [
      ...document.querySelectorAll('.seller-layer-modal'),
      ...document.querySelectorAll('[role="dialog"]'),
      ...document.querySelectorAll('.modal.show'),
      ...document.querySelectorAll('mat-dialog-container'),
    ];
    for (const c of containers) {
      const btn = Array.from(c.querySelectorAll('button'))
        .find(b => LABELS.includes(b.textContent.trim()) && !b.disabled);
      if (btn) { btn.click(); return btn.textContent.trim(); }
    }
    const btn = Array.from(document.querySelectorAll('button'))
      .find(b => LABELS.includes(b.textContent.trim()) && !b.disabled);
    if (btn) { btn.click(); return btn.textContent.trim(); }
    return false;
  });
}

// ─────────────────────────────────────────────────────────
// 팝업 감지 및 닫기
// ─────────────────────────────────────────────────────────
async function checkAndClosePopup(page) {
  return page.evaluate(() => {
    const modal = document.querySelector('.modal.show, [role="alertdialog"], .layer-popup');
    if (!modal) return null;
    const txt = modal.textContent.trim().substring(0, 100);
    const btn = Array.from(modal.querySelectorAll('button'))
      .find(b => ['확인', '닫기'].includes(b.textContent.trim()));
    if (btn) btn.click();
    return txt;
  });
}

// ─────────────────────────────────────────────────────────
// 모달 닫기 (Escape)
// ─────────────────────────────────────────────────────────
async function closeModal(page) {
  try {
    await page.keyboard.press('Escape');
    await sleep(800);
    await page.evaluate(() => {
      const btn = Array.from(document.querySelectorAll('button'))
        .find(b => ['닫기','취소','×','X'].includes(b.textContent.trim()));
      if (btn) btn.click();
    });
    await sleep(400);
  } catch(e) {}
}

// ─────────────────────────────────────────────────────────
// 리뷰 종합 판단 (환불검토 vs 답변으로 대처)
// - 모든 별점에 대해 Claude가 별점·내용·순위를 종합 판단
// - 환불은 당분간 사람이 수동 처리 (이 함수는 판단 + 근거 제공만)
// ─────────────────────────────────────────────────────────
async function judgeReview({ productName, rating, reviewText, reviewPosition }) {
  const posStr = formatReviewPosition(reviewPosition);

  const systemPrompt = `당신은 코에르(COEIR) 브랜드 운영 담당자입니다.
네이버 스마트스토어 리뷰를 보고 "환불검토가 필요한 건인지" vs "답변으로 대처하고 넘어갈 수 있는 건인지" 종합 판단합니다.

[판단 원칙 — 별점만 믿지 말고 내용·순위까지 종합]
1. 별점 1~2 + 제품 실제 불량/품질 문제 → 환불검토 (true)
2. 별점 1~2 + 오배송만 문제 (제품은 OK) → 답변으로 사과 (false)
   단, 리뷰 순위 10위 이내면 노출 영향 커서 환불검토 (true)
3. 별점 1~2 + 단순 변심·가벼운 불만 + 순위 50위 이후 → 답변으로 대처 (false)
4. 별점 4~5 + 내용이 악평·실제 불만 (별점 디코이) → 환불검토 (true)
   단, 별점 4~5이라도 리뷰 순위 1~10위 + 제품 사용 시 불편함·통증·기능 문제 언급 → 환불검토 (true)
   (별점이 높아도 상위 노출 리뷰에 부정 내용이 있으면 브랜드 이미지 타격이 크므로 적극 대응)
   ★ 특히 리뷰 순위 1~5위는 구매 전환에 직접 영향 → "귀찮다", "번거롭다", "두 개 샀는데 하나만 쓴다", "아쉽다" 같은 미묘한 부정 뉘앙스라도 환불검토 (true)로 보수적 판단
   ★★ 리뷰 순위 1~20위 + **부정 내용이 전체 리뷰의 절반 이상**(부정으로 리뷰가 시작되거나, 부정 표현이 더 길게/구체적으로 묘사되거나, 별점 차감 사유로 명시된 경우) → 환불검토 (true)
       (예: 욕실매트 11위 "비싼편이에요~ 길이도 43cm밖에 안되서 좁은 느낌" → 부정으로 시작 + 사이즈 불만 + 긍정 내용 없음 → 환불검토)
   🚨 **가격 불만 단독은 환불검토 사유 아님** — 본문에 "잘 쓰고 있다 / 디자인 예쁘다 / 만족 / 좋아요" 같은 긍정 표현이 함께 있으면, 가격 비싸다·가격대 쎄다·비싼감 있다 표현이 있어도 → 답변 (false)
       (예: "디자인 너무예쁘구요 사이즈 선택도 잘했어요. 가격대는 좀 쎄네요" → 디자인·사이즈 만족이 핵심, 가격은 부수적 의견 → 답변)
       (예: "잘쓰고있어요. 비싼감이있지만" → 잘 쓰고 있다가 핵심 → 답변)
       (예: "좀 비싼거 빼고 안미끌려 좋아요 디쟌 이쁘고" → 좋다·예쁘다가 핵심 → 답변)
   🚨🚨 **단, 1~5위 최상위 + 본문 전체가 부정 내용만 (긍정 표현 0)** → 가격 불만이라도 환불검토 (true) 적극 분류
       이유: 최상위 노출 리뷰에 긍정 한 마디도 없이 부정만 있으면 구매 전환에 치명적
       (예: 1위 "용도에 비해 비싼거같습니다" → 본문 100% 가격 부정, 긍정 표현 없음 → 환불검토)
       (예: 3위 "사이즈가 작네요" → 본문 100% 사이즈 부정 → 환불검토)
       → 같은 가격 불만이라도 본문에 "좋아요/만족" 등 긍정 표현 한 마디라도 있으면 답변, 부정만 있으면 환불검토
5. 별점 4~5 + 전반적 긍정·중립 → 답변으로 대처 (false)
6. 별점 3 + 내용이 부정적·노출 영향 큼 → 환불검토 (true)
   단, 별점 3이라도 **본문이 명백히 긍정**("배송 빠르다", "제품 좋다", "만족" 등 칭찬 위주)이고 **부정 내용 없음** → 답변으로 충분 (false)
   (별점은 실수 가능성이 있으니 본문이 진짜 의도. 답변에서 가벼운 위트로 "5점 아니라 3점은 실수로 잘못 누르신거죠?^^;" 류 톤 사용 가능)
7. 애매한 경우 confidence를 낮게 설정 (사람 검토 유도)
8. [예외 - 답변으로 처리하면 충분한 케이스]
   - 스테인리스 욕실선반 "동봉된 스패너로 설치 시 긁힘 발생" → 환불 사유 아님. 스패너 사용법(제품과 닿지 않게 돌리기 / 마지막에 꽉 조이기 / 약간 긁혀도 가려져서 안 보임) 안내로 충분 → false
   - 단순 설치 미숙으로 인한 표면 흠집은 환불 사유가 아님 (제품 결함 아님) → false
   - 별점 3점이라도 본문이 명백히 긍정·중립이면 환불 사유 없음 → false (위 원칙 6 단서)
   - **디스펜서 "입구가 작다 / 좁다 / 세제 리필이 번거롭다"** → 환불 사유 아님. 디스펜서 입구는 표준 크기이며 깔때기 사용을 안내하면 충분 → false
   - **테라조 비누받침대 "비누가 붙어서 떼야 한다 / 번거롭다"** → 제품 문제가 아니라 비누 특성(무른 비누 vs 솔리드한 비누) 차이. 솔리드한 비누 추천 안내로 충분 → false
   - 위 두 케이스는 리뷰가 "너무 귀엽다 / 디자인 좋다" 같은 긍정 표현으로 시작하면 더더욱 환불 대상 아님 (긍정 톤이 우선)
   - **배송 파손 / 박스 손상이 있었지만 "잘 쓰고 있다 / 사용에 문제 없다"** → 환불 사유 아님 (이미 해결된 케이스). 사과 + 포장·배송 시스템 개편 안내로 충분 → false
       (예: "배송중 파손되서 몇번씩 반품했네요..잘 쓰고 있어요.." → 결국 잘 쓰고 있다 → false)
   - **발리콘 욕실화 "딱딱하다 / 무겁다" + "미끄럼 방지는 좋다"** → 환불 사유 아님. 묵직한 무게감 = 미끄럼방지를 위한 의도된 설계 (제품 특징). 단단함은 EVA와 달리 안전·내구성을 위한 경도. 정정 안내로 충분 → false
   - **본문 전체가 긍정/중립이며 가격 불만만 부수적**으로 언급된 모든 케이스 → 환불 사유 아님 (위 원칙 4 단서)
   - **"사이즈가 더 컸으면 / 작았으면 좋겠다" 같은 단순 사이즈 아쉬움 단독** + 본문 긍정("깔끔합니다", "좋아요" 등) → 환불 사유 아님. "최적의 사이즈로 개발했다" 안내로 충분 → false
   - **"수압이 한 가지" / "기능 다양성이 부족하다" 같은 비핵심 기능 아쉬움 단독** + 본문 긍정 → 환불 사유 아님. "제품팀에 전달하여 개발 검토하겠다" 정도의 약속 톤으로 대응 → false
       (예: 샤워기 "수압이 쎄서 좋고 감 부드럽고 색상 예뻐요. 아쉬운점은 수압이 한가지" → 전체 긍정, 수압 다양성은 핵심 USP 아님 → false)

[리뷰 순위의 의미]
- 네이버 브랜드스토어 랭킹순 기준 순위
- 상위(1~10위)는 구매자가 가장 먼저 보는 리뷰 → 노출 영향 큼
- 하위(50위 이후)는 노출 영향 작음 → 답변으로 충분한 경우 많음

[출력 — 반드시 순수 JSON만, 코드블록·설명 금지]
{
  "needsRefund": true|false,
  "reason": "판단 근거를 한 문장으로 (왜 그렇게 판단했는지 — 별점·내용·순위 중 핵심 요소 언급)",
  "confidence": 0-100
}`;

  const userPrompt = `제품명: ${productName || '-'}
별점: ${rating}점
리뷰 순위: ${posStr}
리뷰 내용: "${reviewText}"

위 리뷰를 판단하고 JSON으로만 출력하세요.`;

  try {
    const msg = await ai.messages.create({
      model:       CONFIG.MODEL,
      max_tokens:  300,
      temperature: 0,
      system:      systemPrompt,
      messages:    [{ role: 'user', content: userPrompt }],
    });
    accumulateUsage(msg.usage);
    let raw = msg.content[0].text.trim();
    raw = raw.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
    const s = raw.indexOf('{'), e = raw.lastIndexOf('}');
    if (s >= 0 && e > s) raw = raw.slice(s, e + 1);

    const parsed = JSON.parse(raw);
    return {
      needsRefund: parsed.needsRefund === true,
      reason:      (typeof parsed.reason === 'string' ? parsed.reason : '').trim() || '근거 없음',
      confidence:  typeof parsed.confidence === 'number' ? parsed.confidence : null,
    };
  } catch (e) {
    // 파싱/호출 실패는 보수적으로 환불검토 처리 (사람이 확인)
    return {
      needsRefund: true,
      reason:      `API/파싱 오류 — 수동 확인 필요 (${e.message})`,
      confidence:  0,
    };
  }
}

// ─────────────────────────────────────────────────────────
// Claude API로 리뷰 답변 생성
// ─────────────────────────────────────────────────────────
// 제품명에서 카테고리 감지 후 공식 USP / 카피 / 주의사항 지식 반환
// 출처: 코에르 공식 시트 (주요 USP · 카피 · 주의사항/제약문구)
function getProductKnowledge(productName) {
  const name = (productName || '').toLowerCase();
  const k = [];

  // ── 욕실 미끄럼방지 매트 ────────────────────────
  if (name.includes('미끄럼') || (name.includes('매트') && !name.includes('규조토') && !name.includes('발매트'))) {
    k.push(`[욕실 미끄럼방지 매트] (22,000원)
- 카피: "물 빠짐은 빠르게, 바닥엔 빈틈없이 밀착 - 관리까지 쉬운 안전한 욕실 미끄럼방지 매트"
- USP: 뛰어난 미끄럼 방지 / 고급 TPE 소재 / KC마크 / **항곰팡이 테스트 통과** / 간편한 세척 관리
- 누적 판매 15만개+, 네이버 쇼핑 1위, 구매 후기 9,000개+, 만족도 4.8/5
- 사이즈: S 55×55 / M 45×75 / L 60×90, 컬러: 그레이/베이지/화이트/블루
- 설치 팁: 빨판을 하나하나 꾹 눌러 밀착
- 주의: 'KC인증' 표현 X → 'KC마크가 있는 제품' 으로 표기
- 주의: TPE를 "젖병 젖꼭지 소재"로 언급 X → "의료용품에 사용되는 소재" 로만 언급
- [촉감/경도 관련 답변] "발에 닿는 느낌이 좋다 / 부드럽다 / 푹신하다" 같은 촉감 리뷰에는:
  → "너무 부드러워 미끄러지지도, 너무 단단하여 발바닥이 아프지도 않은 **최적의 경도를 여러 번 테스트하여 개발**" 했다는 점을 강조 (기능성+촉감 양쪽을 잡은 설계)
- [가격 방어 답변] "비싸다 / 가격이 부담된다 / 가격대 쎄다 / 비싼감" 리뷰에는:
  → 환불 사유 아님. 만족 표현이 함께 있으면 감사 인사 먼저, 그 다음 부드러운 가격 방어
  → 가격 방어 키워드: **"의료 용품에 사용되는 소재"** + KC마크 + 다양한 유해물질 테스트 통과 + **"원재료 값이 월등히 비싼 게 사실"** + "최근 전쟁으로 모든 비용이 증가했지만 **고객님 부담 가중을 막기 위해 기존 가격을 유지**" + "회사에 건의해보겠다" 톤
  → 사이즈 다양성 불만이 함께 있으면 → "S(55×55) / M(45×75) / L(60×90) **3가지 사이즈를 공간에 맞게 조합 사용**할 수 있도록 제공" 안내
- ★ [사이즈 안 맞아서 여러 개 붙여 쓴다 / 두 개 사서 붙여서 쓴다 리뷰]
  → 환불 사유 아님. **원래 이 제품은 욕실 크기와 사용 공간이 다양해서, 용도에 맞게 여러 개 조합해 붙여 쓰도록 설계된 제품**. 이를 자연스럽게 안내
  → 🚨 금지 표현: "공간에 더 잘 맞는 사이즈를 안내해드리지 못한 점 아쉽습니다" — 절대 사용 금지 (제품 의도와 어긋남)
  → 권장 톤: "**S/M/L 3가지 사이즈를 공간 형태에 맞게 자유롭게 조합·연결하여 사용**하실 수 있도록 설계되어 있어요. 두 개를 붙여 쓰시면 욕실 모양에 딱 맞춰지죠!"
- ★ [사이즈가 작다 + 가격 비싸다 동반 리뷰 — 최대 사이즈 디펜스]
  → 핵심 디펜스: "저희 제품의 가장 큰 장점은 **한 손으로 들고 세척이 가능할 만큼 가볍다**는 점이에요. 사이즈가 더 크면 한 손 세척이 어려워, 세척 편의를 위해 **가장 큰 사이즈를 60×90cm 로 설계**했습니다. 더 넓은 공간이 필요하시면 **여러 장을 붙여 사용**하실 수 있도록 작은 사이즈도 함께 제공하고 있어요"
  → 가벼움(한 손 세척) ↔ 사이즈 한계가 트레이드오프임을 설명하면 가격·사이즈 불만이 동시에 풀림
- ★ [사이즈 실제와 차이가 난다 리뷰]
  → 디펜스: "**제조 공정 특성상 ±1cm 정도 미세한 오차는 발생**할 수 있으나 거의 정확합니다. 다만 **받자마자 사용하시면 끝부분이 약간 말려 있어 실제 사이즈보다 작게 느껴질 수 있는데**, 평평하게 펴서 잠시 두시면 자연스럽게 원상복귀 됩니다" 톤
- ★ [물이 고인다 / 매트 밑에 물 고임 리뷰] — 사용 가이드 안내
  → "물이 고여도 고정 잘 된다" 류 긍정 리뷰에는 → 위생적인 사용법으로 자연스럽게 가이드:
  → **"계속 깔아두지 마시고, 다 쓰신 뒤에는 떼서 간단히 씻고 고리나 선반에 걸어 말린 다음, 다시 사용하실 때 깔아두고 사용하시면 더 오래 위생적으로 사용**하실 수 있어요" 톤
  → 양면 공기 순환으로 빠른 건조 + 곰팡이·세균 억제 강조
- [건조 / 말리기 관련 답변] "고리에 걸어 말린다 / 벽에 붙여 말린다" 류 리뷰에는:
  → "**고리나 선반에 걸쳐 말리는 게 벽에 붙여 말리는 것보다 훨씬 건조가 잘 됨**" 안내. 매트 양면이 공기 순환되어 빠른 건조 + 곰팡이·세균 억제에 유리
  → 벽에 붙여 말린다는 리뷰에도 동일하게 "고리/선반이 더 잘 마른다"고 부드럽게 안내
- [가벼움 답변] 욕실매트는 욕실화와 달리 '가벼움'이 장점 → "가볍다"는 리뷰엔 "가벼워서 설치와 세척이 훨씬 편리하다"고 긍정 답변 (욕실화는 절대 X)
- [흡착판 관련 안내] 흡착판이 바닥에 붙는지 여부는 바닥 마감 재질에 따라 다름. 안전 기준:
  1) 건식 사용: 흡착 안 돼도 매트가 밀리지 않아 안전
  2) 잠깐 깔고 사용(물기 있는 상태): 물기로 인해 미끄러지지 않아 안전
  3) 물기 많은 바닥에 계속 깔아두는 경우: 바닥 재질/마감에 따라 매트가 밀릴 수 있음
     → 사용 전 안전 테스트 필수. 안전하지 않다고 판단되면 환불 안내
- "바닥에 안 붙는다"는 리뷰에는 위 기준으로 맥락을 설명하고, 안전 우려면 환불 안내`);
  }

  // ── 워셔블 레더 규조토 발매트 ───────────────────
  if (name.includes('규조토') || name.includes('발매트') || name.includes('워셔블')) {
    k.push(`[워셔블 레더 규조토 발매트] (25,000원)
- 정식명: '워셔블 레더 규조토 발매트' (기존 딱딱한 규조토 발매트와 완전히 다른 카테고리)
- 카피: "규조토의 흡수력은 그대로, 세척과 관리는 훨씬 간편하게 - 워셔블 레더로 완성한 새로운 기준의 규조토 발매트"
- USP: KC마크 / 워셔블 레더 / 항균 / 리얼 규조토 함유(빠른 흡수·건조) / 논슬립 바닥면 / 세련된 디자인
- 사이즈: S 40×60 / M 50×70cm, 디자인: 코에르 / 문페이즈 / 프렌치무드 / 파스텔 테라조
- 세탁기 세탁 가능 (기존 딱딱한 규조토와 차별)
- 기존 섬유 소재 빨아쓰는 규조토 대비 워셔블 레더는 때·오염이 덜 타고 제거도 쉬움
- 주의: 'KC인증' 표현 X → 'KC마크가 있는 제품'
- 주의: 고객이 빨아쓰는 규조토 발매트 개념을 모를 수 있으므로 설명이 필요할 수 있음
- [말림 / 끝부분이 바닥에 안 붙는다는 리뷰]
  → "워셔블 레더 소재 특성상 처음 사용 시 끝부분이 살짝 말리는 경우가 있으나, **시간이 지나면 자연스럽게 원상복귀**됩니다" 라고 안심시키는 안내 추가
- ★ [흡수력·소재 설명 시 핵심 맥락]
  → "워셔블 레더 재질은 **오염에 강한 가죽 재질이면서도 물기는 흡수가 되도록 개발된 제품**" 이라는 두 가지 특성(오염 강함 + 흡수 가능)을 함께 강조. 단순히 "흡수 잘 된다"만 말하지 말고 "가죽 재질이라 오염에는 강한데 흡수까지 되도록 개발했다"는 차별점 어필
- [가격 비싸다 리뷰] → 가격 방어: "안미끌리고 디자인 예쁘다"는 만족 포인트 인정 + 워셔블 레더의 차별성·세탁 편의성으로 자연스럽게 가치 어필
- [사이즈 아쉬움 리뷰 — "더 컸으면 좋겠다"] → 환불 사유 아님
  → "더 크면 사용성이 생각보다 좋지 않아 **여러 차례 테스트를 거쳐 최적의 사이즈로 개발**" 했다는 톤으로 자연스럽게 안내. S(40×60) / M(50×70) 두 가지 옵션 제공 사실도 함께 언급 가능
- ★ [두께·도톰함 리뷰] "도톰하다", "두께가 있다" 류 칭찬엔 비교 강조:
  → "**시중의 대부분 발매트는 2~3mm 내외로 얇아서 푹신하지도 않고 흡수력도 떨어지는데, 저희 제품은 약 5mm 두께**로 푹신함과 큰 흡수 용량을 모두 잡았어요" 톤
  → 단순히 "도톰하다"는 사실 인정에서 그치지 말고 차별점 어필`);
  }

  // ── 발리콘 욕실화 ───────────────────────────────
  if (name.includes('욕실화') || name.includes('슬리퍼') || name.includes('발리콘')) {
    k.push(`[발리콘 욕실화] (22,000원)
- 포지션: '안전'과 '착화감' (경쟁사 '디어커스'의 "깃털같이 가벼운" 과 명확히 차별화)
- 카피: "가벼움의 한계를 넘은 안전한 무게와 편안한 착화감의 욕실화"
- USP: 안전·착화감 고려한 최적의 무게 / 안전한 PVC 소재·항균효과 / 원활한 물 빠짐 14개 배수 홀 / 고급스러운 디자인 / 국내 제작
- 사이즈: M(250) / L(270), 컬러: 베이지 등
- '가볍다'는 리뷰에 가벼움 자체를 칭찬 X → "묵직한 무게감으로 미끄럼을 확실히 잡아주는 설계"로 리프레이밍
- 주의: 굴패각 함유 PVC를 '친환경 소재' 로 표현 X
- 주의: 굴패각과 항균을 인과관계로 연결 X (굴껍질 덕분에 항균효과가 있는 것이 아님)
- 주의: "욕실문에 걸리지 않는 욕실화"라는 표현 X (문 사이즈에 따라 걸릴 수 있음)
- 금지 표현: "욕실용으로 딱 맞게"
- ★ ["딱딱하다 / 무겁다" + "미끄럼 방지 좋다" 리뷰] — 환불 사유 아님 (제품 특징)
  → 무게감은 **가벼워 미끄러운 욕실화와 달리 적당한 무게감으로 미끄럼방지에 최적화** 된 설계임을 강조
  → 단단함은 **최적의 착화감을 위한 모양과 경도**로 제작되었음을 부드럽게 정정. "소재 자체가 딱딱할 수는 없는데, 고객님께는 조금 단단하게 느껴지셨나 보네요^^;" 류 톤
  → 마무리: "단단하고 무게감이 있어 안전할 수 있다는 점 기억해주시면 좋을 것 같습니다!"
  → 예시: "안녕하세요, 고객님! 코에르 발리콘 욕실화는 가벼워 미끄러운 욕실화와는 달리 적당한 무게감으로 미끄럼방지에 최적화되어 설계되었습니다! 동시에 최적의 착화감을 제공하기 위한 모양과 경도로 제작하였습니다. 소재 자체가 딱딱할 수는 없는데, 고객님께는 조금 단단하게 느껴지셨나 보네요^^; 단단하고 무게감이 있어 안전할 수 있다는 점 기억해주시면 좋을 것 같습니다! 리뷰 감사드리며 오늘도 좋은 하루 보내세요:)"
- ★ [세일 가격 / 저렴하게 잘 샀다 리뷰]
  → "**코에르는 자주 세일을 진행하지 않는데, 마침 기간을 잘 맞춰 합리적으로 구매**하셨네요!" 톤
  → 의도: '코에르가 자주 세일하지 않는다' 는 인상을 심어 다음 구매 시점에도 가격 협상력 유지
- ★ [욕실 외 공간 사용 리뷰] — 기숙사·거실·베란다·야외 등에서 신는다는 후기
  → "욕실 외에도 **거실·베란다·기숙사·야외용으로 활용**하시는 고객님들이 많으시며, 용도에도 적합하게 설계되어 있습니다" 톤
  → 사용 다양성을 인정·강조해 제품 가치 확장
- ★ [사이즈 큰 케이스 — "작은 거 살걸", "크게 나오네요" 류]
  → "여유 있는 사이즈" 같은 표현 X (사이즈가 크다는 불만에 대한 핀트 어긋남)
  → "**다음엔 더 잘 맞는 사이즈로 안내**해 드릴 수 있도록 제품팀에 전달하겠습니다" 또는 "한 치수 작은 사이즈가 더 잘 맞으셨을 것 같아 아쉽습니다" 류 톤
- ★ [발등이 낮다는 리뷰 — 의도된 설계, 장점으로 리프레이밍]
  → "발등이 낮은" 디자인은 의도된 설계 포인트. 욕실 문턱·문 사이에 **"걸릴 일이 적다 / 잘 걸리지 않는다"** 부드러운 톤으로 안내
  → 🚨 "절대 걸리지 않는다 / 욕실문에 걸리지 않는 욕실화" 같은 단정 금지 (문 사이즈에 따라 다름)
  → 권장 톤: "발등이 낮은 디자인 덕분에 욕실 문턱이나 문 사이에 **잘 걸리지 않아** 편하게 신고 벗으실 수 있어요"
- ★ [발등이 낮아 불편하다 / 발등 낮은 게 아쉽다는 부정 뉘앙스 리뷰]
  → 발등이 높지 않아도 **착화감이 불편하지 않도록 설계**되어 **많은 고객님들께서 편안하게 신고 계신다**는 톤으로 안내
  → 권장 톤: "발등 라인을 낮게 디자인했지만 착화감이 불편하지 않도록 신경 써 설계해, 많은 고객님들께서 편하게 신고 계세요. 신고 벗기 편하고 문턱에도 잘 걸리지 않는 장점도 함께 느껴보세요"`);
  }

  // ── 스테인리스 제품 공통 ────────────────────────
  const isStainless = name.includes('스테인리스') || name.includes('스텐') ||
    (name.includes('비누받침') && !name.includes('테라조')) ||
    (name.includes('칫솔') && !name.includes('테라조')) ||
    name.includes('멀티홀더') || name.includes('디스펜서') || name.includes('트레이') || name.includes('선반');

  if (isStainless && !name.includes('테라조')) {
    k.push(`[스테인리스 라인 공통]
- SUS304 (녹·부식에 강함) / 브러쉬드 피니쉬 기법 / 샌드블라스팅·전해 연마
- 물빠짐 구조로 물때·곰팡이 방지, 위생적이고 내구성 우수
- 🚨 **안티 핑거프린트 코팅(AFC) 은 일부 제품에만 적용됨**:
  - AFC 적용 ✅: 디스펜서 / 트레이 / 멀티홀더 / 욕실선반
  - AFC 미적용 ❌: **비누받침대, 칫솔꽂이** — 답변에 AFC 언급 금지
- ★ [AFC 언급 시 표현법] 고객은 "AFC" 라는 용어를 모르므로, 약어 단독으로 쓰지 말고 **풀어서 안내**:
  → "안티 핑거프린트 코팅(AFC)으로 **지문과 물자국이 잘 남지 않아** 깔끔한 모습을 오래 유지" 식으로
  → AFC 단독 명시(약어만) 답변 금지`);
  }
  if (name.includes('비누받침') && !name.includes('테라조')) {
    k.push(`[스테인리스 비누받침대] (22,000원 / A·B·C 타입 3종)
- 카피: "SUS304, 마감·구조까지 섬세하게 - 코에르가 만든 스테인리스 기준"
- USP: 원활한 배수 (비누 무름 방지) / SUS304 / 뛰어난 내구성 / 세심히 선정한 사이즈 / 고급스러운 디자인
- 무게감 있는 고급스러움, 타입별 디자인 차이 있음
- 🚨 **AFC(안티 핑거프린트 코팅) 미적용 제품** — 답변에 AFC 언급 금지`);
  }
  if (name.includes('칫솔') && !name.includes('테라조')) {
    k.push(`[스테인리스 칫솔꽂이] (27,000원)
- USP: SUS304 / 물고임 방지 6개 배수 라인 / 활용도 높은 다기능 파티션 / 칫솔 이탈 방지 거치 레이어 / **샌드블라스팅·전해 연마 공법** / 유니크한 스퀘어 오벌 디자인
- ★ [샌드블라스팅 마감 강조] 답변 시 적극 언급할 만한 포인트:
  → "샌드블라스팅 마감으로 표면의 미세한 분쇄물(분술물)을 제거하고 고급스러운 매트한 질감"
  → "전해 연마까지 더해져 SUS304 본연의 깔끔함 + 위생적 마감"
  → 단단함·고급스러움 칭찬 리뷰엔 이 표현 활용 가능
- 🚨 **AFC(안티 핑거프린트 코팅) 미적용 제품** — 답변에 AFC 언급 금지`);
  }
  if (name.includes('멀티홀더') && !name.includes('테라조')) {
    k.push(`[스테인리스 멀티홀더] (25,000원)
- USP: SUS304 / 안티 핑거프린트 코팅 (AFC) / 브러쉬드 피니쉬 기법 / 인체 공학적 설계 / 세련된 미니멀 디자인
- 무광 마감으로 물자국 덜 남음, 치약·칫솔·펜·소품 보관 등 다용도`);
  }
  if (name.includes('디스펜서') && !name.includes('테라조')) {
    k.push(`[스테인리스 디스펜서] (27,000원)
- 소재: SUS304 / 무광(브러쉬드 피니쉬) 마감
- USP: SUS304 / 안티 핑거프린트 코팅 (AFC) / 브러쉬드 피니쉬 / 손쉬운 펌핑 / 세련된 미니멀 디자인
- 핸드워시·주방세제 등 액상 제품용. 거품펌프로의 교환은 불가 (안내 시 주의)
- [입구가 작다 / 좁다 / 세제 리필 번거롭다 리뷰 — 환불 사유 아님]
  → 디스펜서 입구는 **표준 디스펜서 크기**이며, 디스펜서 카테고리 자체가 입구를 작게 설계함 (재현성·내용물 보존을 위해)
  → **깔때기 사용**을 자연스럽게 안내하면 충분. "저희도 입구를 조금 더 크게 만들 수 있는 방안이 있는지 고민해보겠다"는 메시지로 마무리
  → 예시: "안녕하세요, 고객님. 😊 디스펜서 입구가 좁아 불편하셨군요! 그런데 저희 제품의 입구 크기는 표준 디스펜서 입구 크기 정도라 저희 제품이 딱히 작다기 보다 디스펜서라는 제품의 특성상 전반적으로 조금 작게 만들어지고 있는 것 같습니다^^; 그래서 보통 디스펜서에 세제 등을 넣으실 때는 깔때기를 많이 사용들 하십니다. 저희도 입구를 조금 더 크게 만들 수 있는 방안이 있는지 고민해보도록 하겠습니다! 리뷰 감사드리며 오늘도 좋은 하루 보내세요!"`);
  }
  if (name.includes('트레이') && !name.includes('테라조') && !name.includes('드라이')) {
    k.push(`[스테인리스 트레이] (22,000원)
- 사이즈: S(18cm) / M(23cm) / L(30cm)
- USP: SUS304 / 안티 핑거프린트 코팅 (AFC) / 브러쉬드 피니쉬 / 편의성까지 고려한 세련된 디자인 / 다양하게 활용 가능한 멀티 아이템
- 비누받침·가글컵·칫솔 등 욕실 소품 정리용, 물자국 잘 지워짐`);
  }
  if (name.includes('선반')) {
    k.push(`[스테인리스 욕실선반]
- 베이직 (S 20cm 33,000원 / M 30cm 43,000원)
- 플랫 (M 30cm 43,000원 / L 40cm 53,000원) — 드라이기 거치대 포함, 지문 방지 코팅, 녹슬지 않는 나사까지 전체 스텐
- 히든 (43,000원) — 무타공 접착식, 코너 공간 활용 특화
- 공통 USP: 안티 핑거프린트 코팅 (AFC) / SUS304 / 배수라인 / 안전·깔끔 마감 / 실용적 사이즈 / 간편하지만 강력한 무타공
- 무타공 접착식이지만 튼튼하게 고정 (셀프 설치 가능), 타공 옵션도 제공
- 부속품: '무타공 스티커 세트' / '타공 나사 세트' (각 1,000원) 별도 판매
- ★ [무타공 견고함 — 매우 중요] 베이직/히든/플랫 **모두** 무타공 스티커로 설치해도 굉장히 견고함. 물건 넣고 펌핑까지 강하게 해도 흔들림 없는 안정성. "물건 넣으려면 타공 필수"·"무타공이라 약하다" 류 리뷰는 사실과 다르므로 정정하며 적극 강조해야 함. 다만 벽면 상태에 따라 무타공 설치 불가한 벽이 있어, 그 경우만 상세페이지 하단의 안내를 참조하여 타공으로 설치하도록 안내
- [예시 답변 톤 — 무타공 오해 정정]
  "어머, 아니에요! 저희 욕실선반은 모두 무타공으로 설치해도 굉장히 견고하답니다! 물건들 넣고, 심지어 펌핑까지 강하게 해도 전혀 흔들림없는 안정성을 자랑해요^^ 다만 벽면의 상태에 따라 무타공으로 설치가 불가할 수 있는데(상세페이지 하단 안내 참조), 그럴 경우에만 타공으로 설치하셔서 사용하시면 됩니다!"
- [설치 시 스패너 사용 안내] "동봉된 스패너 사용 시 긁힘이 생긴다"는 리뷰에는:
  → 환불 사유가 아니며, 다음 가이드를 부드럽게 안내하면 충분 (환불검토 불필요):
    1) 스패너를 제품과 직접 닿지 않게 돌리면 긁힘 없이 설치 가능
    2) 손으로만 조이면 충분히 단단해지지 않으므로 마지막엔 반드시 스패너로 마감 너트를 꽉 조여줘야 안전
    3) 약간의 긁힘이 생기더라도 깊지 않고, 실제 사용 시 거치 제품 뒤에 가려져 보이지 않음
  → 위 3가지를 안내하며 "안전하게 사용해 주세요" 톤으로 마무리. 답변 길어져도 OK (글자수 예외)`);
  }
  // ── 스텐 세트 가격 정보 ────────────────────────
  if (name.includes('스테인리스') && name.includes('세트')) {
    k.push(`[스테인리스 세트 (약 10% 할인)]
- 2종 세트 (칫솔꽂이+비누받침대): 47,000원
- 3종 세트 (멀티홀더+디스펜서+트레이): 69,500원
- 4종 세트 (3종+욕실선반): 93,000원
- 5종 세트 (칫솔꽂이+비누받침대+멀티홀더+디스펜서+트레이): 112,000원
- 비누받침대 A/B/C 타입, 트레이 S/M/L, 선반 S/M 조합 선택 가능`);
  }

  // ── 테라조 라인 ─────────────────────────────────
  if (name.includes('테라조')) {
    k.push(`[테라조 라인 공통]
- 소재: 레진 기반, 천연스톤을 하나하나 담아낸 100% 핸드메이드
- 공통 USP: 내구성 좋고 관리 편한 레진 소재 / 물때·미끄럼 방지 논슬립 패드 / 100% 핸드메이드 프리미엄 퀄리티
- 세트(비누받침+칫솔꽂이 / 디스펜서+멀티홀더+트레이 등) 사용 시 공간 통일감·인테리어 효과 극대화
- 카피: "천연스톤을 하나하나 담아낸 100% 핸드메이드, 내구성과 감도를 모두 갖춘 코에르 테라조 라인"
- 설명이 필요할 때만 "천연스톤과 레진을 조합하여 만든 저희 테라조 제품은" 표현 사용 (매 답변마다 X)
- ★ [배송 중 파손 / 박스 손상 리뷰 — 환불 사유 아님 (이미 해결된 케이스)]
  → 본문에 "잘 쓰고 있다 / 사용에 문제없다" 가 있으면 → 사과 + **포장·배송 시스템 전면 개편** 안내로 충분
  → 키 메시지: "본 제품이 무게가 무겁고 길다보니 파손 사례가 가끔 발생함" + "**내부 포장 방법을 전면 개편**" + "**배송 시스템을 전면 개편**" + "솔직한 피드백 감사" 톤
  → 예시: "안녕하세요, 고객님. 파손으로 불편함을 끼쳐드려 정말 죄송합니다. 본 제품이 무게가 무겁고 길다보니 파손이 되는 사례가 가끔 발생합니다. 해당 부분을 개선하기 위해 내부 포장 방법을 전면 개편하여 지금은 파손이 되지 않도록 배송 시스템을 전면 개편하였습니다! 솔직한 피드백 감사드리며, 작은 부분까지 발전하는 코에르가 되도록 노력하겠습니다^^ 좋은 하루 보내세요:)"`);
    if (name.includes('비누받침')) {
      k.push(`- 테라조 비누받침대 베이직 (24,000원) / 웨이브 (27,000원) — 웨이브가 더 고급 디자인
- 배수 홀로 물빠짐 원활, 비누 무름 방지. '묵직해서 움직이지 않는' 안정감
- 베이직에서 "비누가 미끄러진다"는 리뷰에는 웨이브 모델을 언급할 수 있음
- [비누가 받침대에 붙는다 / 떼어내야 한다 리뷰 — 환불 사유 아님]
  → 제품 문제가 아니라 **비누의 특성** 차이. 무른 비누는 더 잘 달라붙고, 솔리드(단단)한 비누는 잘 안 붙음
  → "다양한 비누로 테스트해본 결과 비누 성질에 따라 다르더라" + "솔리드한 비누 추천" 톤으로 안내
  → 예시: "안녕하세요, 고객님. 😊 비누가 받침대에 붙어서 불편하셨군요! 제품을 개발하며 저희가 다양한 비누로 테스트를 진행했었는데, 비누가 받침대에 붙는 문제는 제품보다는 비누의 특성에 따라 달랐어요! 많이 무른 비누의 경우엔 제품에 상대적으로 더 잘 달라붙지만, 좀 더 솔리드한 비누의 경우엔 붙지 않고 편안하게 사용할 수 있었습니다. 현재 비누가 불편하시다면 조금 더 단단하고 솔리드한 비누로 바꾸어 사용해보시길 추천드립니다! 리뷰 감사드리며 오늘도 좋은 하루 보내세요!"`);
    }
    if (name.includes('칫솔')) {
      k.push(`- 테라조 칫솔꽂이 라운드 (31,900원, 원가 36,000원) / 스퀘어 (38,000원) — 스퀘어가 상위 라인
- USP (마스터 테이블 기준):
  1) 내구성 좋고 관리가 편한 **레진 소재**
  2) 세심하게 설계된 **배수 홀과 배수 라인** (구체적 개수 명시 X — "6개" 같은 숫자는 SS 칫솔꽂이 USP 이므로 혼동 금지)
  3) 물때 및 미끄럼 방지를 위한 **논슬립 패드** (바닥 고무패킹)
  4) 안정감 있고 위생을 고려한 **거치 홀**
  5) **100% 핸드메이드**, 프리미엄 퀄리티 (천연스톤을 하나하나 담아냄)
- 🚨 **TZ 칫솔꽂이에 "6개 배수 라인" 같은 숫자 명시 금지** — 그건 SS(스테인리스) 칫솔꽂이 USP. TZ는 "배수 홀·배수 라인이 있다" 정도로만 표현
- 가격 비싸다는 리뷰엔 천연스톤 100% 핸드메이드 + 레진 내구성 + 자연스러운 무늬(각 제품 고유) 강조`);
    }
    if (name.includes('멀티홀더')) k.push(`- 테라조 멀티홀더 (29,000원): 칫솔·치약·양치컵 등 여러 아이템 통합 수납`);
    if (name.includes('디스펜서')) k.push(`- 테라조 디스펜서 (36,000원): 실버 톤 뚜껑, 공간 정돈·인테리어 효과. 거품펌프 교환 불가 (안내 시 주의)`);
    if (name.includes('트레이') && !name.includes('드라이')) k.push(`- 테라조 트레이 (29,000원~): S(20cm) / M(27cm, +5,000원), 비누·칫솔·머리끈 등 다용도 정리`);
    if (name.includes('드라이') || name.includes('드라이트레이')) k.push(`- 테라조 드라이트레이 (40,000원): M(30cm) / L(40cm), 물이 아래로 흐르는 배수 강화 구조 → 곰팡이 억제, 비누 수명 연장`);
    if (name.includes('캐니스터')) k.push(`- 테라조 캐니스터 (32,000원): 뚜껑 있음, 화장솜·면봉 등 위생적 수납. 호텔 같은 고급스러움`);
  }
  // ── 테라조 세트 가격 정보 ──────────────────────
  if (name.includes('테라조') && name.includes('세트')) {
    k.push(`[테라조 세트 (약 10% 할인)]
- 2종 세트 (비누받침대+칫솔꽂이): 50,000원 — 베이직/웨이브 × 라운드/스퀘어 4가지 조합
- 3종 세트 (디스펜서+멀티홀더+트레이): 90,000원 — 트레이 S/M 선택
- 5종 세트 (디스펜서+멀티홀더+트레이+비누받침대+칫솔꽂이): 134,000원`);
  }

  // ── 오가닉 에버닌 그린 타월 ─────────────────────
  if (name.includes('타월') || name.includes('수건') || name.includes('에버닌')) {
    k.push(`[오가닉 에버닌 그린 타월]
- 카피: "3년 이상 화학비료 없이 재배된 토양에서 자란, GOTS 최고 등급 오가닉 코튼 100%"
- USP: GOTS 인증 최고급 오가닉 코튼 / OEKO-TEX 1등급 / 친환경 EFD 염색(무형광·무표백) / 호텔 스펙 능가 40수·220g·543gsm / 먼지 걱정 없는 '슈퍼 코마사' / 리버시블 디자인 (앞뒤 컬러 차이)
- 소재 표기는 반드시 '슈퍼 코마사' (그냥 '코마사' 아님)
- 앞뒤 컬러가 다른 디자인은 반드시 "리버시블 디자인"으로 표현 ('양면 컬러' 표현 금지)
- 디자인: 스트라이프 / 체커드 / 플로라 / 블룸 (모두 에버닌 그린 파스텔 톤)
- 구성/가격: 단품 1P (19,000원) / 벌크 4P·8P (69,000원~, 1P씩 깔끔 포장) / 패키지 4P·8P (72,000원~, 선물용 박스 포함)
- 선물·답례품 리뷰 대응 시 활용 가능 포인트: **타월은 전용 패키지 선물 박스**(그리고 원하시면 쇼핑백)에 담아 드릴 수 있어 답례품·집들이 선물로 바로 활용 가능
  ※ 현재 선물용 전용 패키지 박스가 준비된 품목은 '타월'만이므로 다른 제품에서는 이 포인트 언급 금지
- 일반 수건보다 도톰하고 크며, 먼지 적고 여러 번 세탁에도 변형 없음
- 주의: 'KC인증' 표현보다 'KC 마크를 획득한 제품' 이 더 정확`);
  }

  // ── 샤워기·필터 라인 ────────────────────────────
  if (name.includes('샤워기')) {
    k.push(`[PLA 필터 온오프 샤워기 클라우드 화이트] (39,000원)
- ★★ 핵심 차별점 (최우선 강조): 샤워기 본체와 필터가 100% 식물 유래 PLA 소재 → **"미세플라스틱이 검출되지 않는"** 샤워기
  (일반 플라스틱 샤워기는 물이 흐를 때 미세플라스틱이 배출되지만, 코에르 PLA 샤워기는 소재 자체가 미세플라스틱을 배출하지 않음)
- 카피: "식물 유래 PLA 소재로 미세플라스틱부터 잔류 염소·유해 물질까지 걸러내는 듀얼 필터 온오프 샤워기"
- USP: Non-GMO 식물 유래 PLA(생분해성) / 바디·헤드 듀얼 필터 케어 / 원터치 STOP 온오프 버튼
- 듀얼 필터: ① PLA 바디 필터 (미세플라스틱 미검출·녹물·불순물 여과·세균 99.9% 항균) ② ACF 헤드필터 (잔류 염소·냄새 흡착)
- 편의: 수도꼭지까지 안 가고 손잡이 STOP 버튼으로 물 조절 → 물 절약, 고압 수압
- 무광 화이트 컬러, 욕실 인테리어 매치 우수
- 구성품: 샤워기 본체 + PLA 필터 기본 포함 (샤워호스·추가 PLA·ACF 헤드필터는 별도 구매)
- ※ 답변 시 "미세플라스틱 미검출/검출되지 않음" 포인트를 반드시 우선 언급 (단순 필터 기능보다 PLA 소재 자체의 안전성이 진짜 차별점)
- ★ ["수압이 한 가지" / "수압 모드 다양했으면" 류 비핵심 기능 아쉬움 — 환불 사유 아님]
  → 수압 다양성은 코에르 샤워기의 핵심 USP가 아님 (핵심은 미세플라스틱 미검출 + 듀얼 필터 + 온오프). 본문이 전체 긍정이면 답변으로 충분
  → 톤: 만족 포인트(수압·감·색상 등) 공감 + "**수압을 다양하게 개발해볼 수 있도록 제품팀에 전달하겠다**" 약속 톤. 환불·교환 안내 X`);
  }
  if (name.includes('pla') && name.includes('필터') && !name.includes('샤워기')) {
    k.push(`[PLA 필터] (16,000원)
- ★★ 핵심 차별점: **100% 식물 유래 PLA 소재라서 필터에서 미세플라스틱이 검출되지 않음**
  (일반 플라스틱 필터는 오히려 미세플라스틱을 방출할 수 있음)
- 카피: "100% 식물 유래 소재로 미세플라스틱이 검출되지 않는 항균 필터"
- USP: 식물 유래 PLA / 녹물·불순물 여과 / 필터 속 세균까지 99.9% 항균
- PLA 필터 온오프 샤워기 전용 교체 필터
- ※ 답변 시 "미세플라스틱 미검출" 포인트 최우선 강조`);
  }
  if (name.includes('acf') || name.includes('헤드필터')) {
    k.push(`[ACF 헤드필터] (10,000원)
- 카피: "잔류 염소와 냄새를 흡착해 피부에 닿기 전 한 번 더 걸러주는 ACF 헤드필터"
- USP: 잔류 염소 제거·냄새 흡착 / 케이스·부직포까지 식물 유래 소재
- PLA 샤워기 헤드에 장착하는 추가 필터 (듀얼 케어 강화용)`);
  }
  if (name.includes('샤워호스') || (name.includes('호스') && !name.includes('홀'))) {
    k.push(`[샤워호스] (17,000원)
- 카피: "물때와 세균 번식을 줄이고 360º 회전으로 꼬임 없는 샤워호스"
- USP: 물때·세균 번식 억제 / 호스 꼬임 없는 360º 회전 구조 / 내구성 높인 견고한 설계`);
  }

  // ── 액세서리 ────────────────────────────────────
  if (name.includes('기프트 박스') || name.includes('기프트박스')) {
    k.push(`[기프트 박스] (3,000원) — 선물용 포장 박스, 타월·테라조·스텐 제품 선물 세트에 추가 구매`);
  }
  if (name.includes('쇼핑백')) {
    k.push(`[쇼핑백] (2,500원, L은 +500원) — 선물·지참용 쇼핑백`);
  }
  if (name.includes('무타공 스티커')) {
    k.push(`[무타공 스티커 세트] (1,000원) — 스테인리스 선반 무타공 설치/재부착용 스티커`);
  }
  if (name.includes('타공 나사')) {
    k.push(`[타공 나사 세트] (1,000원) — 스테인리스 선반 타공 설치용 스테인리스 나사 부속`);
  }
  if (name.includes('논슬립 패드')) {
    k.push(`[논슬립 패드] (1,000원) — 테라조 라인(멀티홀더·디스펜서·트레이 등) 하단 부착용 미끄럼방지 패드`);
  }

  return k.join('\n\n');
}

async function generateReply(review) {
  const { productName, rating, reviewText, writer } = review;
  const productKnowledge = getProductKnowledge(productName);

  const systemPrompt = `당신은 코에르(COEIR) 브랜드의 고객 응대 담당자입니다.
코에르는 욕실용품 전문 브랜드로, 스테인리스·테라조·규조토·슈퍼코마사 등 프리미엄 소재의 세련된 욕실 소품을 판매합니다.

[브랜드 공통 소재 지식]
- 테라조 라인: 레진 기반, 천연스톤을 하나하나 담은 100% 핸드메이드. "천연스톤과 레진을 조합하여 만든 저희 테라조 제품은" 표현은 설명이 필요한 맥락에서만 사용
- 스테인리스 라인: SUS304, 브러쉬드 피니쉬 기법, 무타공이어도 튼튼하게 고정됨
  - 🚨 **AFC(안티 핑거프린트 코팅) 은 일부 제품에만 적용** — 적용 ✅: 디스펜서/트레이/멀티홀더/욕실선반 / 미적용 ❌: **비누받침대, 칫솔꽂이** (이 두 제품 답변엔 AFC 언급 금지)
- 규조토 발매트: '워셔블 레더 규조토 발매트' — 세탁 가능, 부드럽고 얼룩·세균 걱정 없음. 'KC마크 획득 제품' (KC인증 X)
- 타월: 40수 '슈퍼 코마사' (그냥 '코마사' 아님), GOTS 최고 등급 오가닉 코튼, OEKO-TEX 1등급
- 욕실화(발리콘): 묵직한 무게감으로 미끄럼을 확실히 잡아주는 '안전·착화감' 포지션. '가볍다' 강조 금지
- 미끄럼방지 매트: TPE 소재, 'KC마크가 있는 제품' (KC인증 X). 누적 15만개+ 판매
- 샤워기/필터 라인: **100% 식물 유래 PLA 소재 → 미세플라스틱이 검출되지 않음** (일반 플라스틱 샤워기·필터는 미세플라스틱이 배출됨). 잔류 염소·녹물·세균까지 듀얼 필터로 차단

${productKnowledge ? `[이 제품 관련 지식]\n${productKnowledge}\n` : ''}
[답변 작성 규칙]
- 글자수는 **리뷰 성격에 따라 가변** (기준 ~180자에서 ±30% 범위)
  - 짧게(110~150자): 단순·짧은 긍정 리뷰("좋아요", "잘 받았어요") + 순위 낮음(41위 이하 / 100위 밖)
  - 기본(160~200자): 중간 분량 / 21~40위 / 일반 긍정
  - 길게(210~260자): 상위 노출(1~20위) + 디펜스 필요 / 부정 뉘앙스 정정 / 안전·사용법 안내 / 환불검토 답변 대응시
  - 마무리 인사를 넣을 공간이 부족하면 초과 허용 (마무리 인사 누락 절대 금지)
  - 🚨 모든 답변이 비슷한 분량으로 나오면 안 됨. 짧은 리뷰엔 짧게, 디펜스 필요한 리뷰엔 길게 — **리뷰의 무게에 답변의 무게를 맞출 것**
- 답변은 반드시 "안녕하세요, 고객님." 으로 시작 (구매자 아이디 절대 사용 금지)
- 호칭은 '고객님'으로 통일
- 고객의 리뷰에서 언급된 구체적인 내용 1가지 반드시 포함
- 제품 특성에 맞는 정확한 용어 사용 (타월='슈퍼 코마사', 발매트='워셔블 레더 규조토', 스텐='SUS304')
- '가볍다'는 **욕실화** 리뷰에는 절대 "가볍다"를 강조 X → 묵직한 무게감이 미끄럼방지 설계라고 부드럽게 안내
- '가볍다'는 **욕실매트·발매트** 리뷰에서는 오히려 장점 → "가벼워서 설치와 세척이 편리하다"고 긍정적으로 답변 (욕실화와 정반대 처리)
- 욕실 미끄럼방지 매트의 가격 부담 리뷰에는 고급 TPE + KC마크 + **항곰팡이 테스트 통과** 3가지를 묶어 가격 방어
- 욕실 매트의 촉감/푹신함 리뷰에는 "너무 부드러워 미끄러지지도, 너무 단단하여 발바닥이 아프지도 않은 최적의 경도를 여러 번 테스트하여 개발" 식으로 기능성+촉감 양립을 강조
- 금지 표현: "욕실용으로 딱 맞게", "KC인증"(→'KC마크 획득 제품'), "친환경 PVC", "굴패각 덕분에 항균", "욕실문에 걸리지 않는", "양면 컬러"(→'리버시블 디자인')
- 🚨 **고객센터 전화번호·연락처 임의 표기 절대 금지** (예: "1588-XXXX", "02-XXXX" 같은 번호를 만들지 말 것). 환불·교환·문의 안내가 필요하면 "**고객센터로 연락 주시거나**" 정도만 표현할 것
- [욕실선반 무타공 견고함 강조] "타공이 필수다"·"무타공이라 약할 것 같다" 류 리뷰에는 → 욕실선반 베이직/히든/플랫 모두 무타공 스티커로 설치해도 굉장히 견고하며 펌핑 강하게 해도 흔들림 없음을 적극 강조. 다만 벽면 상태에 따라 무타공 설치 불가한 벽이 있으니 그 경우만 상세페이지 하단을 참조하여 타공으로 설치하라고 안내. 무타공 = 타공만큼 견고함을 반드시 어필
- [3점 + 본문 명백 긍정 케이스] 별점 3점인데 본문이 칭찬 위주("배송 빠르다", "제품 좋다" 등)면 별점이 실수로 잘못 눌린 가능성 → 가벼운 위트 한 줄 가능 (예: "5점이 아니라 3점은 실수로 잘못 누르신거죠?^^;"). 단 너무 무례하지 않게, 부드럽고 친근한 톤으로
- TPE 소재는 "의료용품에 사용되는 소재"로만 언급 ('젖병 젖꼭지 소재' X)
- 테라조 제품 설명이 필요할 때만 "천연스톤과 레진을 조합하여" 표현 활용 (매번 X)
- 세트 구매/인테리어 관련 리뷰는 세트 사용 시 공간이 통일감 있게 살아난다는 점 언급 가능
- 이모티콘 1~2개 사용 (😊 🤍 ✨ 💧 등 내용에 어울리는 것)
- 따뜻하고 진정성 있는 톤, "리뷰 남겨주셔서 감사합니다" 또는 "고맙습니다" 반드시 포함 (마지막 마무리 인사로 배치)
- 별점 3점 이하 + 불만 내용이면 먼저 사과 후 개선/교환/문의 안내
- 별점 4~5점 + 긍정 내용이면 공감·감사·재방문 권유
- 답변 텍스트만 출력 (따옴표, JSON, 헤더 등 불필요)

[문법·자연스러움 주의사항] (반복 지적 사항 — 매우 중요)
- 주어-서술어 관계를 반드시 확인할 것. 다음은 모두 비문(X):
  ✗ "물빠짐을 사용하다" / "설계를 경험하다" / "미끄럼 방지 설계를 경험하다" / "흡수력을 사용하다" / "기능을 사용하다"
- '~을/를 경험하다'는 추상명사(기능·효과·편안함 등)와만 결합. '설계'·'구조'·'디자인'은 경험의 대상이 아님:
  ✓ "미끄럼 방지 기능을 경험하다" / "안전한 착화감을 경험하다"
  ✗ "미끄럼 방지 설계를 경험하다" / "구조를 경험하다"
- '~을/를 사용하다'는 도구·제품 자체가 목적어. 추상명사(물빠짐·흡수력·기능 등)는 사용 대상이 아님.
- 동사·형용사가 어울리는 목적어와 결합되는지 체크. 예: "흡수력이 좋아 만족하셨다니" O, "흡수력을 사용하셔서" X
- 타월 답변에서 "여러 번 세탁해도 변형 없이 오래 사용하실 수 있어요", "GOTS 최고 등급 오가닉 코튼으로 만든 40수 슈퍼 코마사" 등 동일 표현을 반복하지 말 것. 고객 리뷰의 구체적 내용에 맞게 변주하여 작성할 것`;

  const userPrompt = `다음 리뷰에 대한 판매자 답변을 작성해주세요:

제품명: ${productName}
별점: ${rating}점
구매자: ${writer}
리뷰 내용: "${reviewText}"`;

  const msg = await ai.messages.create({
    model:      CONFIG.MODEL,
    max_tokens: 300,
    system:     systemPrompt,
    messages:   [{ role: 'user', content: userPrompt }],
  });
  accumulateUsage(msg.usage);

  return msg.content[0].text.trim();
}

// ─────────────────────────────────────────────────────────
// Slack DM 전송 (Slack Bot Token 필요)
// ─────────────────────────────────────────────────────────
// 오류 발생 시 Slack 단문 알림 (catch 핸들러용)
async function sendSlackError(message) {
  const token     = cfg.SLACK_BOT_TOKEN;
  const channelId = cfg.SLACK_CHANNEL_ID || cfg.SLACK_USER_ID || 'U08KNE04HKK';
  if (!token) return;
  const body = JSON.stringify({ channel: channelId, text: `🚨 *코에르 리뷰 자동화 오류*\n${message}` });
  return new Promise(resolve => {
    const req = require('https').request(
      { hostname: 'slack.com', path: '/api/chat.postMessage', method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json; charset=utf-8' } },
      res => { res.resume(); resolve(); }
    );
    req.on('error', resolve);
    req.write(body);
    req.end();
  });
}

// 워드 파일을 슬랙 채널에 업로드 (files.getUploadURLExternal flow)
// 1) getUploadURLExternal 로 임시 업로드 URL 받음
// 2) URL 에 파일 바이너리 POST
// 3) completeUploadExternal 로 채널에 게시
async function uploadFileToSlack(filepath, channelId, comment = '') {
  const token = cfg.SLACK_BOT_TOKEN;
  if (!token) { log('[슬랙 파일] 토큰 없음 — 스킵'); return; }
  if (!channelId) { log('[슬랙 파일] 채널 ID 없음 — 스킵'); return; }
  if (!fs.existsSync(filepath)) { log(`[슬랙 파일] 파일 없음: ${filepath}`); return; }

  const filename = path.basename(filepath);
  const fileBuf  = fs.readFileSync(filepath);
  const length   = fileBuf.length;

  const httpsReq = (opts, body) => new Promise((resolve, reject) => {
    const req = require('https').request(opts, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });

  try {
    // 1) 업로드 URL 발급
    const step1Body = `filename=${encodeURIComponent(filename)}&length=${length}`;
    const step1 = await httpsReq({
      hostname: 'slack.com', path: '/api/files.getUploadURLExternal', method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(step1Body),
      },
    }, step1Body);
    const step1Json = JSON.parse(step1.body);
    if (!step1Json.ok) { log(`[슬랙 파일] URL 발급 실패: ${step1Json.error}`); return; }
    const { upload_url, file_id } = step1Json;

    // 2) 발급된 URL 에 바이너리 POST
    const url = new URL(upload_url);
    const step2 = await httpsReq({
      hostname: url.hostname, path: url.pathname + url.search, method: 'POST',
      headers: { 'Content-Type': 'application/octet-stream', 'Content-Length': length },
    }, fileBuf);
    if (step2.status >= 400) { log(`[슬랙 파일] 업로드 HTTP ${step2.status}`); return; }

    // 3) completeUploadExternal 로 채널 게시
    const step3Body = JSON.stringify({
      files: [{ id: file_id, title: filename }],
      channel_id: channelId,
      initial_comment: comment || '',
    });
    const step3 = await httpsReq({
      hostname: 'slack.com', path: '/api/files.completeUploadExternal', method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json; charset=utf-8',
        'Content-Length': Buffer.byteLength(step3Body),
      },
    }, step3Body);
    const step3Json = JSON.parse(step3.body);
    if (step3Json.ok) log(`[슬랙 파일] 업로드 완료: ${filename}`);
    else               log(`[슬랙 파일] 게시 실패: ${step3Json.error}`);
  } catch (e) {
    log(`[슬랙 파일] 오류: ${e.message}`);
  }
}

async function sendSlackDM(summary) {
  const token     = cfg.SLACK_BOT_TOKEN;
  const channelId = cfg.SLACK_CHANNEL_ID || cfg.SLACK_USER_ID || 'U08KNE04HKK'; // 채널 우선, 없으면 DM

  if (!token) {
    log('[슬랙] SLACK_BOT_TOKEN 미설정 → 슬랙 전송 생략');
    return;
  }

  // 별점 → 별 이모지 변환
  const stars = n => '⭐'.repeat(Math.max(0, Math.min(5, n || 0)));

  // 날짜 포맷 정리 (예: "2026. 4. 16." → "2026.04.16")
  const dateStr = summary.date.replace(/\s/g, '').replace(/\.$/, '')
    .split('.').map((p, i) => i === 0 ? p : p.padStart(2, '0')).join('.');

  // ── 토큰/비용 요약 문자열 ───────────────────────
  const u = summary.usage || {};
  const usageLine = u.totalTokens != null
    ? `• API 사용: ${u.calls || 0}회 호출 · 총 ${u.totalTokens.toLocaleString()} 토큰 (입력 ${(u.input||0).toLocaleString()} / 출력 ${(u.output||0).toLocaleString()}) · 예상 비용 $${(u.costUSD||0).toFixed(4)} (약 ₩${(u.costKRW||0).toLocaleString()})`
    : '';

  // ── 헤더 ─────────────────────────────────────────
  const lines = [
    `안녕하세요! 오늘(${dateStr}) 코에르 리뷰 자동 답변 작업 완료 보고드립니다 😊`,
    `작업 요약`,
    `• 조건: 1주일 작성 + 답글미등록`,
    `• 총 처리: ${summary.replied}개 완료 ✅  |  환불검토: ${summary.refund}개  |  실패: ${summary.failed}개`,
    `• 실행환경 : ${summary.executionEnv || '백그라운드 자동 (Anthropic API 직접 호출)'}`,
    `• AI 모델 : ${summary.aiModel || (CONFIG.MODEL || 'claude-sonnet-4-5')}`,
    `• 비용 : ${summary.cost || (u.costKRW != null ? `약 ₩${(u.costKRW||0).toLocaleString()} ($${(u.costUSD||0).toFixed(4)})` : '-')}`,
    ...(usageLine ? [usageLine] : []),
    ``,
  ];

  // ── 환불검토 항목 (먼저 표시) ─────────────────────
  const refundList = sortByRank(summary.results.filter(r => r.refundCheck === '검토필요'));
  if (refundList.length > 0) {
    lines.push(`─────────────────────────────`);
    lines.push(`⚠️ 환불검토 필요 항목 (${refundList.length}건)`);
    refundList.forEach((r, i) => {
      lines.push(`No.${i + 1}  ${stars(r.rating)} (${r.rating}점)  |  ${r.writer}`);
      lines.push(`리뷰글번호 : ${r.reviewNo || '-'}`);
      lines.push(`제품명 : ${getMasterProductName(r.productName)}`);
      { const opt = resolveDisplayOption(r); if (opt) lines.push(`구매 옵션 : ${opt}`); }
      lines.push(`별점 : ${stars(r.rating)} (${r.rating}점)`);
      lines.push(`리뷰순위 : ${formatReviewPosition(r.reviewPosition)}`);
      lines.push(`📋 대응 정책 : ${r.refundPolicy || '미확인'}`);
      lines.push(`리뷰 : "${r.reviewText.replace(/\n/g, ' ')}"`);
      if (r.judgeLabel) {
        lines.push(`*🏷️ 판단 : ${r.judgeLabel}*`);
      }
      if (r.judgeReason) {
        lines.push(`*🔴 판단 근거 : ${r.judgeReason}${r.judgeConfidence != null ? ` (confidence ${r.judgeConfidence})` : ''}*`);
      }
      if (r.replyText) {
        lines.push(`💬 답변 대응시 : "${r.replyText}"`);
      }
      lines.push(``);
    });
  }

  // ── 답변 항목을 둘로 분리: 판단체크 필요(근거표시) / 일반 ──
  const answeredAll = summary.results.filter(r => r.replyText && r.refundCheck !== '검토필요');
  const checkList   = sortByRank(answeredAll.filter(r => shouldShowJudgeReason(r)));   // 판단체크 필요
  const repliedList = sortByRank(answeredAll.filter(r => !shouldShowJudgeReason(r)));  // 일반 답변완료

  // ── ② 답변 완료 - 판단체크 필요 (환불검토 다음) ───
  if (checkList.length > 0) {
    lines.push(`─────────────────────────────`);
    lines.push(`⚠️ 답변 완료 - 판단체크 필요 (${checkList.length}건)`);
    checkList.forEach((r, i) => {
      lines.push(`─────────────────────────────`);
      lines.push(`No.${i + 1}`);
      lines.push(`리뷰글번호 : ${r.reviewNo || '-'}`);
      lines.push(`등록자 : ${r.writer}`);
      lines.push(`제품명 : ${getMasterProductName(r.productName)}`);
      { const opt = resolveDisplayOption(r); if (opt) lines.push(`구매 옵션 : ${opt}`); }
      lines.push(`별점 : ${stars(r.rating)} (${r.rating}점)`);
      lines.push(`리뷰순위 : ${formatReviewPosition(r.reviewPosition)}`);
      lines.push(`리뷰 : "${r.reviewText.replace(/\n/g, ' ')}"`);
      if (r.judgeLabel) {
        lines.push(`*🏷️ 판단 : ${r.judgeLabel}*`);
      }
      lines.push(`*🔴 판단 근거 : ${r.judgeReason}${r.judgeConfidence != null ? ` (confidence ${r.judgeConfidence})` : ''}*`);
      lines.push(`답변 : "${r.replyText}"`);
      lines.push(``);
    });
  }

  // ── ③ 답변 완료 항목 (일반) ───────────────────────
  if (repliedList.length > 0) {
    lines.push(`─────────────────────────────`);
    lines.push(`■ 답변 완료 항목 (${repliedList.length}건)`);
    repliedList.forEach((r, i) => {
      lines.push(`─────────────────────────────`);
      lines.push(`No.${i + 1}`);
      lines.push(`리뷰글번호 : ${r.reviewNo || '-'}`);
      lines.push(`등록자 : ${r.writer}`);
      lines.push(`제품명 : ${getMasterProductName(r.productName)}`);
      { const opt = resolveDisplayOption(r); if (opt) lines.push(`구매 옵션 : ${opt}`); }
      lines.push(`별점 : ${stars(r.rating)} (${r.rating}점)`);
      lines.push(`리뷰순위 : ${formatReviewPosition(r.reviewPosition)}`);
      lines.push(`리뷰 : "${r.reviewText.replace(/\n/g, ' ')}"`);
      if (r.judgeLabel) {
        lines.push(`*🏷️ 판단 : ${r.judgeLabel}*`);
      }
      lines.push(`답변 : "${r.replyText}"`);
      lines.push(``);
    });
  }

  // ── 실패 항목 ─────────────────────────────────────
  const failList = summary.results.filter(r => !r.replyText && r.refundCheck === '-');
  if (failList.length > 0) {
    lines.push(`─────────────────────────────`);
    lines.push(`❌ 답변 실패 항목 (${failList.length}건)`);
    failList.forEach((r, i) => {
      lines.push(`No.${i + 1}  ${stars(r.rating)} (${r.rating}점)  |  ${r.writer}`);
      lines.push(`제품명 : ${getMasterProductName(r.productName)}`);
      lines.push(`리뷰 : "${r.reviewText.replace(/\n/g, ' ')}"`);
      lines.push(``);
    });
  }

  const text = lines.join('\n');

  return new Promise(resolve => {
    const body = JSON.stringify({ channel: channelId, text });
    const req  = https.request(
      {
        hostname: 'slack.com',
        path:     '/api/chat.postMessage',
        method:   'POST',
        headers:  {
          'Content-Type':  'application/json; charset=utf-8',
          'Authorization': `Bearer ${token}`,
        },
      },
      res => {
        let data = '';
        res.on('data', chunk => data += chunk);
        res.on('end', () => {
          try {
            const parsed = JSON.parse(data);
            if (parsed.ok) log('[슬랙] DM 전송 완료');
            else           log(`[슬랙] 오류: ${parsed.error}`);
          } catch(e) { log(`[슬랙] 응답 파싱 오류: ${e.message}`); }
          resolve();
        });
      }
    );
    req.on('error', e => { log(`[슬랙] 요청 오류: ${e.message}`); resolve(); });
    req.write(body);
    req.end();
  });
}

// ─────────────────────────────────────────────────────────
// 현재 보이는 AG Grid 행 수집
// ─────────────────────────────────────────────────────────
// ─────────────────────────────────────────────────────────
// Word 문서 생성 (슬랙 메시지와 동일 양식)
// ─────────────────────────────────────────────────────────
async function generateWordDoc(summary) {
  const stars     = n => '⭐'.repeat(Math.max(0, Math.min(5, n || 0)));
  // Word용 별모양 TextRun 배열 — 점수만큼만 금색 ★ 표시 (빈별 없음)
  // bold: 제목줄(No.~)에서 사용, size: 기본 22 (label 기본값과 동일)
  const starRuns = (n, { bold = false, size = 22 } = {}) => {
    const filled = Math.max(0, Math.min(5, n || 0));
    if (filled === 0) return [];
    return [new TextRun({ text: '★'.repeat(filled), bold, size, font: 'Malgun Gothic', color: 'F5A623' })];
  };
  const dateStr   = summary.date.replace(/\s/g, '').replace(/\.$/, '')
    .split('.').map((p, i) => i === 0 ? p : p.padStart(2, '0')).join('.');

  const divider   = () => new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'AAAAAA', space: 1 } },
    spacing: { after: 160 },
    children: [],
  });

  const label = (txt, bold = false, size = 22) => new TextRun({ text: txt, bold, size, font: 'Malgun Gothic' });
  const GRAY  = 'F2F2F2';

  const children = [
    // ── 헤더 ───────────────────────────────────────────
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 80 },
      children: [new TextRun({ text: '코에르 리뷰 자동 답변 보고서', bold: true, size: 32, font: 'Malgun Gothic', color: '1F3864' })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 300 },
      children: [new TextRun({ text: `작업일: ${dateStr}`, size: 22, font: 'Malgun Gothic', color: '666666' })],
    }),

    // ── 작업 요약 ────────────────────────────────────────
    new Paragraph({
      spacing: { before: 100, after: 100 },
      shading: { fill: 'D9E1F2', type: ShadingType.CLEAR },
      children: [label('작업 요약', true, 24)],
    }),
    new Paragraph({
      spacing: { after: 80 },
      children: [label(`• 조건: 1주일 작성 + 답글미등록`)],
    }),
    new Paragraph({
      spacing: { after: 80 },
      children: [label(`• 총 처리: ${summary.replied}개 완료 ✅   |   환불검토: ${summary.refund}개   |   실패: ${summary.failed}개`)],
    }),
    new Paragraph({
      spacing: { after: 80 },
      children: [label(`• 실행환경 : ${summary.executionEnv || '백그라운드 자동 (Anthropic API 직접 호출)'}`)],
    }),
    new Paragraph({
      spacing: { after: 80 },
      children: [label(`• AI 모델 : ${summary.aiModel || (CONFIG.MODEL || 'claude-sonnet-4-5')}`)],
    }),
    new Paragraph({
      spacing: { after: summary.usage ? 80 : 300 },
      children: [label(`• 비용 : ${summary.cost || (summary.usage && summary.usage.costKRW != null ? `약 ₩${summary.usage.costKRW.toLocaleString()} ($${summary.usage.costUSD})` : '-')}`)],
    }),
    ...(summary.usage ? [new Paragraph({
      spacing: { after: 300 },
      children: [label(
        `• API 사용: ${summary.usage.calls}회 호출   |   토큰 ${summary.usage.totalTokens.toLocaleString()} `
        + `(입력 ${summary.usage.input.toLocaleString()} / 출력 ${summary.usage.output.toLocaleString()}`
        + ` / 캐시읽기 ${summary.usage.cacheRead.toLocaleString()} / 캐시쓰기 ${summary.usage.cacheWrite.toLocaleString()})`
        + `   |   비용: $${summary.usage.costUSD} (≈ ₩${summary.usage.costKRW.toLocaleString()})`
      )],
    })] : []),

    divider(),

    // ── 환불검토 항목 (먼저 표시) ─────────────────────────
    ...(() => {
      const list = sortByRank(summary.results.filter(r => r.refundCheck === '검토필요'));
      if (!list.length) return [];
      return [
        new Paragraph({
          spacing: { before: 200, after: 160 },
          children: [label('⚠️ 환불검토 필요 항목', true, 24)],
        }),
        ...list.flatMap((r, i) => [
          new Paragraph({
            spacing: { before: 160, after: 80 },
            shading: { fill: 'FFF2CC', type: ShadingType.CLEAR },
            children: [label(`No.${i + 1}   `, true), ...starRuns(r.rating, { bold: true }), label(` (${r.rating}점)   |   ${r.writer}`, true)],
          }),
          new Paragraph({ spacing: { after: 60 }, children: [label('리뷰글번호 : ', true), label(r.reviewNo || '-')] }),
          new Paragraph({ spacing: { after: 60 }, children: [label('제품명 : ', true), label(getMasterProductName(r.productName))] }),
          ...((opt => opt ? [new Paragraph({ spacing: { after: 60 }, children: [label('구매 옵션 : ', true), label(opt)] })] : [])(resolveDisplayOption(r))),
          new Paragraph({ spacing: { after: 60 }, children: [label('별점 : ', true), ...starRuns(r.rating), label(` (${r.rating}점)`)] }),
          ...(r.reviewPosition > 0 ? [
            new Paragraph({ spacing: { after: 60 }, children: [
              label('📍 리뷰 순위 : ', true),
              new TextRun({ text: `${r.reviewPosition}위 (랭킹순)`, font: 'Malgun Gothic', size: 22, bold: true, color: 'C00000' }),
            ] }),
            new Paragraph({ spacing: { after: 60 }, children: [
              label('📋 대응 정책 : ', true),
              new TextRun({ text: r.refundPolicy || '', font: 'Malgun Gothic', size: 22, bold: true, color: '833C00' }),
            ] }),
          ] : [
            new Paragraph({ spacing: { after: 60 }, children: [label('리뷰순위 : ', true), new TextRun({ text: formatReviewPosition(r.reviewPosition), font: 'Malgun Gothic', size: 22, color: r.reviewPosition === -2 ? '888888' : '000000' })] }),
            ...(r.refundPolicy ? [new Paragraph({ spacing: { after: 60 }, children: [label('📋 대응 정책 : ', true), label(r.refundPolicy)] })] : []),
          ]),
          new Paragraph({ spacing: { after: 60 }, children: [
            label('리뷰 : ', true),
            new TextRun({ text: `"${r.reviewText.replace(/\n/g, ' ')}"`, font: 'Malgun Gothic', size: 22, italics: true }),
          ] }),
          ...(r.judgeLabel ? [new Paragraph({ spacing: { after: 60 }, children: [
            new TextRun({ text: '🏷️ 판단 : ', bold: true, size: 22, font: 'Malgun Gothic' }),
            new TextRun({ text: r.judgeLabel, bold: true, size: 22, font: 'Malgun Gothic', color: r.judgeLabel === '환불검토' ? 'C00000' : '2E7D32' }),
          ] })] : []),
          ...(r.judgeReason ? [new Paragraph({ spacing: { after: 60 }, children: [
            new TextRun({ text: '💭 판단 근거 : ', bold: true, size: 22, font: 'Malgun Gothic', color: 'C00000' }),
            new TextRun({ text: `${r.judgeReason}${r.judgeConfidence != null ? ` (confidence ${r.judgeConfidence})` : ''}`, font: 'Malgun Gothic', size: 22, bold: true, color: 'C00000' }),
          ] })] : []),
          ...(r.replyText ? [new Paragraph({ spacing: { after: 60 }, children: [
            new TextRun({ text: '💬 답변 대응시 : ', bold: true, size: 22, font: 'Malgun Gothic', color: '7030A0' }),
            new TextRun({ text: `"${r.replyText}"`, font: 'Malgun Gothic', size: 22, color: '7030A0' }),
          ] })] : []),
          new Paragraph({ spacing: { after: 80 }, children: [] }),
          new Paragraph({ spacing: { after: 160 }, children: [label('피드백 : ', true)] }),
        ]),
        divider(),
      ];
    })(),

    // ── 답변 항목 렌더 헬퍼 (판단체크/일반 공용) ──────────
    ...(() => {
      const renderAnswered = (r, i, showReason) => [
        new Paragraph({
          spacing: { before: 200, after: 80 },
          shading: { fill: GRAY, type: ShadingType.CLEAR },
          children: [label(`No.${i + 1}`, true, 22)],
        }),
        new Paragraph({ spacing: { after: 60 }, children: [label('리뷰글번호 : ', true), label(r.reviewNo || '-')] }),
        new Paragraph({ spacing: { after: 60 }, children: [label('등록자 : ', true), label(r.writer)] }),
        new Paragraph({ spacing: { after: 60 }, children: [label('제품명 : ', true), label(getMasterProductName(r.productName))] }),
        ...((opt => opt ? [new Paragraph({ spacing: { after: 60 }, children: [label('구매 옵션 : ', true), label(opt)] })] : [])(resolveDisplayOption(r))),
        new Paragraph({ spacing: { after: 60 }, children: [label('별점 : ', true), ...starRuns(r.rating), label(` (${r.rating}점)`)] }),
        new Paragraph({
          spacing: { after: 60 },
          children: [
            label('리뷰순위 : ', true),
            new TextRun({ text: formatReviewPosition(r.reviewPosition), font: 'Malgun Gothic', size: 22, bold: true, color: r.reviewPosition === -2 ? '888888' : '1F3864' }),
          ],
        }),
        new Paragraph({
          spacing: { after: 60 },
          children: [
            label('리뷰 : ', true),
            new TextRun({ text: `"${r.reviewText.replace(/\n/g, ' ')}"`, font: 'Malgun Gothic', size: 22, italics: true }),
          ],
        }),
        ...(r.judgeLabel ? [new Paragraph({
          spacing: { after: 60 },
          children: [
            new TextRun({ text: '🏷️ 판단 : ', bold: true, size: 22, font: 'Malgun Gothic' }),
            new TextRun({ text: r.judgeLabel, bold: true, size: 22, font: 'Malgun Gothic', color: r.judgeLabel === '환불검토' ? 'C00000' : '2E7D32' }),
          ],
        })] : []),
        ...(showReason ? [new Paragraph({
          spacing: { after: 60 },
          children: [
            new TextRun({ text: '💭 판단 근거 : ', bold: true, size: 22, font: 'Malgun Gothic', color: 'C00000' }),
            new TextRun({ text: `${r.judgeReason}${r.judgeConfidence != null ? ` (confidence ${r.judgeConfidence})` : ''}`, font: 'Malgun Gothic', size: 22, bold: true, color: 'C00000' }),
          ],
        })] : []),
        new Paragraph({
          spacing: { after: 80 },
          children: [
            label('답변 : ', true),
            new TextRun({ text: `"${r.replyText}"`, font: 'Malgun Gothic', size: 22, color: '1F497D' }),
          ],
        }),
        new Paragraph({ spacing: { after: 80 }, children: [] }),
        new Paragraph({ spacing: { after: 160 }, children: [label('피드백 : ', true)] }),
      ];

      const answeredAll = summary.results.filter(r => r.replyText && r.refundCheck !== '검토필요');
      const checkList   = sortByRank(answeredAll.filter(r => shouldShowJudgeReason(r)));
      const normalList  = sortByRank(answeredAll.filter(r => !shouldShowJudgeReason(r)));

      const out = [];
      // ② 답변 완료 - 판단체크 필요
      if (checkList.length > 0) {
        out.push(new Paragraph({
          spacing: { before: 200, after: 160 },
          children: [label('⚠️ 답변 완료 - 판단체크 필요', true, 24)],
        }));
        checkList.forEach((r, i) => out.push(...renderAnswered(r, i, true)));
        out.push(divider());
      }
      // ③ 답변 완료 항목 (일반)
      out.push(new Paragraph({
        spacing: { before: 200, after: 160 },
        children: [label('■ 답변 완료 항목', true, 24)],
      }));
      normalList.forEach((r, i) => out.push(...renderAnswered(r, i, false)));
      return out;
    })(),

    divider(),

    // ── 실패 항목 ─────────────────────────────────────────
    ...(() => {
      const list = summary.results.filter(r => !r.replyText && r.refundCheck === '-');
      if (!list.length) return [];
      return [
        new Paragraph({
          spacing: { before: 200, after: 160 },
          children: [label('❌ 답변 실패 항목', true, 24)],
        }),
        ...list.flatMap((r, i) => [
          new Paragraph({
            spacing: { before: 160, after: 80 },
            shading: { fill: 'FCE4D6', type: ShadingType.CLEAR },
            children: [label(`No.${i + 1}   `, true), ...starRuns(r.rating, { bold: true }), label(` (${r.rating}점)   |   ${r.writer}`, true)],
          }),
          new Paragraph({ spacing: { after: 60 }, children: [label('제품명 : ', true), label(getMasterProductName(r.productName))] }),
          new Paragraph({
            spacing: { after: 160 },
            children: [
              label('리뷰 : ', true),
              new TextRun({ text: `"${r.reviewText.replace(/\n/g, ' ')}"`, font: 'Malgun Gothic', size: 22, italics: true }),
            ],
          }),
        ]),
        divider(),
      ];
    })(),

  ];

  const doc = new Document({
    styles: {
      default: { document: { run: { font: 'Malgun Gothic', size: 22 } } },
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children,
    }],
  });

  const replyDir = path.join(__dirname, 'reply');
  if (!fs.existsSync(replyDir)) fs.mkdirSync(replyDir, { recursive: true });

  const now = new Date();
  const dateNum = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;
  const filename = `${dateNum}_reply.docx`;
  const filepath = path.join(replyDir, filename);

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(filepath, buffer);
  return { filename, filepath };
}

async function collectVisibleRows(page) {
  // product_naver_ids.json 의 제품명 목록을 브라우저 컨텍스트로 전달
  let knownProductNames = [];
  let knownProductIds   = [];
  try {
    const ids = getProductNaverIds();
    knownProductNames = Object.values(ids).map(v => v && v.name).filter(Boolean);
    knownProductIds   = Object.keys(ids); // urlId 들 — reviewNo 와 충돌 방지용
  } catch {}
  return page.evaluate((knownNames, knownIds) => {
    window.__COEIR_PRODUCT_NAMES__ = knownNames || [];
    window.__COEIR_PRODUCT_IDS__   = knownIds || [];

    // ── 첫 3행의 col-id + 셀 내용 진단 로그 (옵션 누락 추적용) ──
    // 헤더의 한국어 컬럼명도 함께 출력 → 옵션 컬럼 col-id 즉시 식별 가능
    if (!window.__COEIR_COLID_LOGGED__) {
      // 헤더 col-id → 한국어 라벨 매핑
      const headerMap = {};
      document.querySelectorAll('.ag-header-cell').forEach(h => {
        const colId = h.getAttribute('col-id') || '';
        const text = (h.innerText || h.textContent || '').trim().split('\n')[0];
        if (colId) headerMap[colId] = text;
      });
      console.log('[헤더 col-id ↔ 라벨]', JSON.stringify(headerMap));
      const debugRows = Array.from(document.querySelectorAll('.ag-center-cols-container .ag-row')).slice(0, 3);
      debugRows.forEach((row, idx) => {
        const dump = {};
        row.querySelectorAll('.ag-cell').forEach(c => {
          const k = c.getAttribute('col-id') || '?';
          dump[k] = (c.innerText || '').slice(0, 60);
        });
        console.log(`[row#${idx} col-id 덤프]`, JSON.stringify(dump));
      });
      window.__COEIR_COLID_LOGGED__ = true;
    }
    const pinnedRows  = Array.from(document.querySelectorAll('.ag-pinned-left-cols-container .ag-row'));
    const centerRows  = Array.from(document.querySelectorAll('.ag-center-cols-container .ag-row'));

    return centerRows.map(row => {
      const rowIndex = row.getAttribute('row-index') || '';

      // ── 체크박스 비활성 여부 확인 (pinned left 영역 기준) ──
      const pinnedRow = pinnedRows.find(r => r.getAttribute('row-index') === rowIndex);
      const cb = pinnedRow ? pinnedRow.querySelector('input[type="checkbox"]') : null;
      // 행 자체에 disabled 클래스가 있거나, 체크박스의 disabled 속성, aria-disabled, 부모 클래스 확인
      const rowClass = (pinnedRow || row).getAttribute('class') || '';
      const checkboxDisabled =
        /disabled/i.test(rowClass) ||  // 행 레벨 disabled 클래스
        (cb
          ? (cb.disabled || cb.getAttribute('aria-disabled') === 'true' ||
             cb.closest('[class*="disabled"]') !== null)
          : false);

      const cells = Array.from(row.querySelectorAll('.ag-cell'));
      const texts = cells.map(c => c.textContent.trim()).filter(t => t);

      // col-id 별 셀 매핑 (AG Grid 표준 속성)
      const cellByColId = {};
      cells.forEach(c => {
        const colId = c.getAttribute('col-id') || '';
        if (colId) cellByColId[colId] = (c.innerText || c.textContent || '').trim();
      });

      // ── col-id 키 목록 (아래에서 공용으로 사용) ──
      const colKeys = Object.keys(cellByColId);

      // ── 리뷰글번호 / 채널번호 (productNaverId 충돌 방지) ──
      // 셀러센터 답글미등록 페이지 실측 col-id: id=리뷰글번호, productNo=채널상품번호
      const reviewColKey = colKeys.find(k => k === 'id' || /^review.*(no|id|seq)$/i.test(k) || /reviewNo|reviewId|reviewSeq/.test(k));
      const channelColKey = colKeys.find(k => k === 'productNo' || /channel.*(no|id|product)|productNo|productId/i.test(k));
      const knownPidSet = (typeof window.__COEIR_PRODUCT_IDS__ !== 'undefined') ? new Set(window.__COEIR_PRODUCT_IDS__) : new Set();
      const allNumeric = texts.filter(t => /^\d{8,12}$/.test(t));
      // 알려진 productNaverId 는 제외
      const numericNonPid = allNumeric.filter(n => !knownPidSet.has(n));
      // 10자리(=실제 reviewNo) 우선, 없으면 첫 비-PID 숫자
      const tenDigit = numericNonPid.find(n => n.length === 10);
      const reviewNo = (reviewColKey && cellByColId[reviewColKey])
        ? cellByColId[reviewColKey]
        : (tenDigit || numericNonPid[0] || '');
      const channelNo = (channelColKey && cellByColId[channelColKey])
        ? cellByColId[channelColKey]
        : (allNumeric.find(n => knownPidSet.has(n)) || '');
      const writer      = texts.find(t => t.includes('*') && t.length >= 3) || '';
      const date        = texts.find(t => /\d{4}\.\d{2}\.\d{2}/.test(t)) || '';
      const ratingStr   = texts.find(t => /^[1-5]$/.test(t)) || '';
      const rating      = ratingStr ? parseInt(ratingStr) : 0;

      // ── 제품명 + 구매옵션 분리 ──
      // 1) col-id로 시도: 'productName'/'option' 또는 유사 패턴
      let productName = '';
      let optionName  = '';
      // 헤더 한글 라벨 기반 col-id 찾기 (셀러센터 컬럼명: "옵션", "구매옵션", "상품명", "주문상품명" 등)
      const headerLabels = {};
      document.querySelectorAll('.ag-header-cell').forEach(h => {
        const cid = h.getAttribute('col-id') || '';
        const txt = (h.innerText || h.textContent || '').trim().split('\n')[0];
        if (cid && txt) headerLabels[cid] = txt;
      });
      const optionByLabel  = Object.keys(headerLabels).find(k => /옵션/.test(headerLabels[k]));
      const productByLabel = Object.keys(headerLabels).find(k => /상품명|제품명|주문상품/.test(headerLabels[k]));
      // 옵션 col-id 후보 확장: optionName, prdOption, productOption, optnNm, opt + 한글 라벨 매칭
      const optionColKey  = optionByLabel  || colKeys.find(k => /option|optn(?!o)|prdOpt|productOpt|^opt$/i.test(k));
      const productColKey = productByLabel || colKeys.find(k => /product.*name|productname|productNm|prdNm|goodsName|goodsNm/i.test(k));
      if (productColKey) {
        const raw = cellByColId[productColKey];
        // 같은 셀에 제품명 + 옵션이 줄바꿈/슬래시로 합쳐진 경우 분리
        const lines = raw.split(/\n+|\s\/\s/).map(s => s.trim()).filter(Boolean);
        productName = lines[0] || '';
        if (!optionName && lines.length > 1) optionName = lines.slice(1).join(' / ');
      }
      if (optionColKey) {
        const raw = cellByColId[optionColKey];
        if (raw) optionName = raw.replace(/\s+/g, ' ').trim();
      }

      // 2) col-id 매칭 실패 시: 기존 텍스트 휴리스틱으로 productName 후보 찾기
      if (!productName) {
        const candidates = texts.filter(t =>
          (t.includes('코에르') || t.startsWith('[세트]')) &&
          t.length > 10 &&
          !/(?:요|다|까|군요|네요|겠어요|했어요|이에요|예요|습니다|세요|거요|아요|어요)\s*[.!?♡~]?\s*$/.test(t)
        );
        // product_naver_ids.json 에 등록된 제품명과 정확히 일치하는 후보를 우선 선택
        let chosen = '';
        if (candidates.length && typeof window.__COEIR_PRODUCT_NAMES__ === 'undefined') {
          // 캐시 없음 → 단순히 첫 번째 사용
          chosen = candidates[0];
        } else if (candidates.length) {
          chosen = candidates.find(c => window.__COEIR_PRODUCT_NAMES__.includes(c)) || candidates[0];
        }
        productName = chosen || '';
      }

      // 3) 그래도 옵션 못 찾으면 휴리스틱: productName 셀 안 짧은 두 번째 줄
      if (!optionName && productColKey) {
        const raw = cellByColId[productColKey] || '';
        // "제품명\n옵션" 또는 "제품명 / 옵션" 또는 "제품명: 옵션" 패턴
        const split = raw.split(/\n+|\s\/\s|\s:\s|\s-\s/).map(s => s.trim()).filter(Boolean);
        if (split.length > 1) optionName = split.slice(1).join(' / ');
      }
      // reviewText: 제품명·날짜·숫자·작성자·UI레이블이 아닌 첫 번째 텍스트
      // '코에르' 포함이라도 productName과 다르면 리뷰 내용일 수 있음 (예: "코에르 제품 좋아요")
      // (아래 옵션 휴리스틱(4)에서도 참조하므로 미리 계산)
      const SKIP_LABELS = ['일반','프리미엄','포토','한달사용','베스트','답글있음','답변완료','답변있음'];
      const reviewText  = texts.find(t =>
        t.length > 3 &&
        t !== productName &&               // 제품명 그 자체는 제외
        !/^\d+$/.test(t) &&               // 순수 숫자 제외
        !t.includes('*') &&               // 작성자 패턴 제외
        !/\d{4}\.\d{2}\.\d{2}/.test(t) && // 날짜 제외
        !SKIP_LABELS.includes(t)           // UI 레이블 제외
      ) || '';

      // 4) 마지막 휴리스틱: 색상/사이즈 키워드 포함된 짧은 셀
      if (!optionName) {
        const optCandidate = texts.find(t =>
          t.length > 1 && t.length <= 40 &&
          t !== productName && t !== reviewText &&
          /(?:색상|사이즈|컬러|타입|종류|개입|호선|S$|M$|L$|XL$|소형|중형|대형)/i.test(t)
        );
        if (optCandidate) optionName = optCandidate;
      }

      // 옵션 정리: 제품명과 동일하면 제거, 너무 긴(50자+) 텍스트는 옵션 아님
      if (optionName && (optionName === productName || optionName.length > 80)) optionName = '';
      // 옵션 텍스트가 productName 을 포함하면 그 부분 제거
      if (optionName && productName && optionName.includes(productName)) {
        optionName = optionName.replace(productName, '').replace(/^[\s\/·\-]+|[\s\/·\-]+$/g, '');
      }

      // 최종수정일이 존재하면 답변 있음 (날짜가 2개 이상: 등록일 + 최종수정일)
      const allDates = texts.filter(t => /\d{4}\.\d{2}\.\d{2}/.test(t));
      const hasReply = allDates.length >= 2 ||
        texts.some(t => t.includes('답글있음') || t.includes('답변완료') || t.includes('답변있음'));

      return {
        rowIndex, channelNo, reviewNo, writer, date, rating,
        productName, optionName, reviewText, hasReply,
        checkboxDisabled,  // ← 환불완료 등으로 비활성화된 행
      };
    }).filter(r => r.date || r.reviewNo);
  }, knownProductNames, knownProductIds);
}

// ─────────────────────────────────────────────────────────
// AG Grid 스크롤
// ─────────────────────────────────────────────────────────
async function scrollDown(page) {
  return page.evaluate(() => {
    const sels = ['.ag-body-viewport', '.ag-center-cols-viewport'];
    for (const sel of sels) {
      const vp = document.querySelector(sel);
      if (vp && vp.scrollHeight > vp.clientHeight) {
        const before = vp.scrollTop;
        vp.scrollTop += Math.max(vp.clientHeight * 0.8, 300);
        if (vp.scrollTop !== before) return true;
      }
    }
    return false;
  });
}

// ─────────────────────────────────────────────────────────
// 단건 리뷰 답변 등록 (선택 전/후 검증 포함)
// ─────────────────────────────────────────────────────────
async function postReply(page, row, replyText) {

  // ── STEP 1: 선택 초기화 + 0개 확인 (최대 3회) ──────
  for (let attempt = 1; attempt <= 3; attempt++) {
    await deselectAllRows(page);
    const cnt = await getCheckedCount(page);
    if (cnt === 0) break;
    log(`  [해제 재시도 ${attempt}/3] 아직 ${cnt}개 선택됨`);
    if (attempt === 3) return { success: false, reason: '체크박스 초기화 실패' };
    await sleep(500);
  }

  // ── STEP 2: 대상 행 체크박스 클릭 ──────────────────
  const checked = await clickRowCheckbox(page, row.rowIndex);
  if (!checked) return { success: false, reason: '체크박스 클릭 실패' };
  await sleep(600);

  // ── STEP 3: 정확히 1개 선택됐는지 확인 ─────────────
  const afterCheck = await getCheckedCount(page);
  if (afterCheck !== 1) {
    log(`  [경고] 체크박스 ${afterCheck}개 선택됨 (1개여야 함) → 초기화`);
    await deselectAllRows(page);
    return { success: false, reason: `선택 수 오류 (${afterCheck}개)` };
  }

  // ── STEP 4: 답글작성 버튼 클릭 ─────────────────────
  const btnClicked = await clickReplyButton(page);
  if (!btnClicked) {
    await page.evaluate((idx) => {
      const r = Array.from(document.querySelectorAll('.ag-center-cols-container .ag-row'))
        .find(row => row.getAttribute('row-index') === idx);
      if (r) r.click();
    }, row.rowIndex);
    await sleep(1200);
    const retry = await clickReplyButton(page);
    if (!retry) return { success: false, reason: '답글작성 버튼 없음' };
  }
  await sleep(2500);

  // ── STEP 5: 모달에서 리뷰번호 확인 (1개인지) ────────
  const modalReviewCount = await page.evaluate(() => {
    const modal = document.querySelector('.seller-layer-modal, .modal.show, [role="dialog"]');
    if (!modal) return 0;
    const reviewNoField = modal.querySelector('textarea[readonly], input[readonly]');
    if (!reviewNoField) return 1;
    return reviewNoField.value.split(',').filter(s => s.trim()).length;
  });
  if (modalReviewCount > 1) {
    log(`  [경고] 모달에 ${modalReviewCount}개 리뷰번호 → 초기화 후 중단`);
    await closeModal(page);
    await deselectAllRows(page);
    return { success: false, reason: `모달 다중 선택 (${modalReviewCount}개)` };
  }

  // ── STEP 6: 답변 텍스트 입력 ────────────────────────
  const typed = await typeReplyText(page, replyText);
  if (!typed) {
    await closeModal(page);
    return { success: false, reason: '답변 입력란 없음' };
  }
  await sleep(500);

  // ── STEP 7: 등록 ────────────────────────────────────
  const submitted = await clickSubmit(page);
  if (!submitted) {
    await closeModal(page);
    return { success: false, reason: '등록 버튼 없음' };
  }
  await sleep(1500);

  // ── STEP 8: 오류 팝업 확인 ──────────────────────────
  const popup = await checkAndClosePopup(page);
  if (popup && popup.includes('이미 답글')) {
    await deselectAllRows(page);
    return { success: false, reason: '이미 답변 있음' };
  }

  await sleep(CONFIG.replyDelay - 1500);

  // ── STEP 9: 등록 후 선택 초기화 ────────────────────
  await deselectAllRows(page);
  const afterPost = await getCheckedCount(page);
  if (afterPost > 0) {
    log(`  [경고] 등록 후에도 ${afterPost}개 선택 상태 → 강제 해제`);
    await deselectAllRows(page);
    await sleep(300);
  }

  return { success: true };
}

// ─────────────────────────────────────────────────────────
// 제품 ID 매핑 (brand.naver.com URL ID → originProductNo)
// ─────────────────────────────────────────────────────────
let _productNaverIds = null;
function getProductNaverIds() {
  if (!_productNaverIds) {
    try {
      const p = path.join(__dirname, 'product_naver_ids.json');
      _productNaverIds = fs.existsSync(p) ? JSON.parse(fs.readFileSync(p)) : {};
    } catch(e) { _productNaverIds = {}; }
  }
  return _productNaverIds;
}

// 제품명으로 brand.naver.com URL ID + originProductNo 찾기 (키워드 점수 매칭)
function findProductNaverId(productName) {
  const ids    = getProductNaverIds();
  const pName  = (productName || '').replace(/\s+/g, ' ').trim();

  // 완전 일치 우선
  for (const [urlId, info] of Object.entries(ids)) {
    if (info.name === pName) return { urlId, ...info };
  }

  // 키워드 점수 매칭 (길이 2 이상 단어 기준) + 디스크리미네이터 가중치
  const keywords = pName.split(/[\s[\]()·]+/).filter(w => w.length >= 2);
  // 같은 라인에서 변형을 구분하는 결정 키워드 — 일치하면 가중치 +5
  const DISCRIMINATORS = ['베이직', '히든', '플랫', '스퀘어', '라운드', '웨이브', '듀얼', '올인원', '스타터', 'PLA', 'ACF', 'TypeC', '1P', '4P', '8P'];
  const pDisc = DISCRIMINATORS.filter(d => pName.includes(d));
  let bestMatch = null;
  let bestScore = 0;

  for (const [urlId, info] of Object.entries(ids)) {
    const iName = (info.name || '').replace(/\s+/g, ' ');
    let score = 0;
    for (const kw of keywords) {
      if (iName.includes(kw)) score++;
    }
    // 디스크리미네이터: pName 에 포함된 단어가 후보 이름에도 있으면 +5, 없으면 -3
    for (const d of pDisc) {
      if (iName.includes(d)) score += 5;
      else                   score -= 3;
    }
    if (score > bestScore) { bestScore = score; bestMatch = { urlId, ...info }; }
  }

  return bestScore >= 2 ? bestMatch : null;
}

// ─────────────────────────────────────────────────────────
// brand.naver.com 리뷰 순위 찾기 (REVIEW_RANKING 정렬)
// 환불검토 대상 리뷰가 판매 페이지 몇 번째인지 조회
// ─────────────────────────────────────────────────────────

// 이모지·제어문자 제거 후 공백 정규화 + 말줄임표/반복 구두점 통일
function normalizeReviewText(text) {
  return (text || '')
    .replace(/[\u{1F000}-\u{1FFFF}\u{2300}-\u{27FF}\u{FE00}-\u{FEFF}\u{1FA00}-\u{1FAFF}]/gu, '')
    .replace(/[…⋯⋮]/g, '.')   // 유니코드 말줄임표 → 점
    .replace(/\.{2,}/g, '.')    // 점 2개 이상 → 점 1개
    .replace(/\s+/g, ' ')
    .trim();
}

// 한글(가-힣)만 추출 — 말줄임표/구두점/이모지 차이를 완전히 무시한 비교용
function koreanOnlyKey(text) {
  return (text || '').replace(/[^가-힣]/g, '');
}

async function findReviewPosition(browser, productName, reviewText, writer) {
  if (!reviewText || !reviewText.trim()) {
    log(`  [순위 조회] reviewText 없음 → 순위 조회 불가`);
    return { position: -1, policy: null, productUrl: null };
  }

  const productInfo = findProductNaverId(productName);
  if (!productInfo) {
    log(`  [순위 조회] 제품 ID 매핑 없음: "${productName.substring(0, 30)}..."`);
    return { position: -1, policy: null, productUrl: null };
  }

  const productUrl = `https://brand.naver.com/coeir/products/${productInfo.urlId}`;

  // originProductNo가 없는 제품(욕실화 등) → 페이지에서 자동 추출 시도
  if (!productInfo.originProductNo) {
    log(`  [순위 조회] originProductNo 없음 → 자동 추출 시도...`);
    const tempPage = await browser.newPage();
    try {
      await tempPage.setUserAgent(
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
      );
      // 네트워크 응답 URL에서 originProductNo 패턴 캡처 (fallback #2)
      let networkOrigin = null;
      const onResp = resp => {
        const u = resp.url();
        // 예: /contents/reviews/product-summary/{origin}  /live-product-detail/broadcasts?originProductNo={origin}
        const m1 = u.match(/product-summary\/(\d{6,})/);
        const m2 = u.match(/originProductNo=(\d{6,})/);
        const m3 = u.match(/benefits\/by-products\/(\d{6,})/);
        const n = (m1 && m1[1]) || (m2 && m2[1]) || (m3 && m3[1]);
        if (n && !networkOrigin) networkOrigin = Number(n);
      };
      tempPage.on('response', onResp);
      await tempPage.goto(productUrl, { waitUntil: 'networkidle2', timeout: 30000 });
      let extracted = await tempPage.evaluate(() => {
        try {
          const s = window.__PRELOADED_STATE__;
          if (!s) return null;
          // 여러 경로 시도
          const cands = [
            s?.product?.A?.originProductNo,
            s?.productDetail?.originProductNo,
            s?.productDetailState?.originProductNo,
            Object.values(s?.product || {}).find(v => v?.originProductNo)?.originProductNo,
          ];
          return cands.find(v => v != null) || null;
        } catch(e) { return null; }
      });
      if (!extracted && networkOrigin) extracted = networkOrigin;
      if (extracted) {
        productInfo.originProductNo = extracted;
        // 다음 실행 시 재사용을 위해 JSON 파일 업데이트
        const ids = getProductNaverIds();
        if (ids[productInfo.urlId]) {
          ids[productInfo.urlId].originProductNo = extracted;
          _productNaverIds = ids; // 인메모리 캐시도 갱신
          fs.writeFileSync(
            path.join(__dirname, 'product_naver_ids.json'),
            JSON.stringify(ids, null, 2)
          );
          log(`  [순위 조회] originProductNo 추출 성공: ${extracted} → product_naver_ids.json 저장`);
        }
      } else {
        log(`  [순위 조회] originProductNo 추출 실패, 순위 조회 불가`);
        return { position: -1, policy: null, productUrl };
      }
    } catch(e) {
      log(`  [순위 조회] originProductNo 추출 오류: ${e.message}`);
      return { position: -1, policy: null, productUrl };
    } finally {
      await tempPage.close().catch(() => {});
    }
  }

  const { originProductNo, checkoutMerchantNo } = productInfo;

  log(`  [순위 조회] brand.naver.com 검색 시작...`);
  log(`    매핑 제품: urlId=${productInfo.urlId} / originProductNo=${originProductNo} / 이름="${(productInfo.name || '').slice(0, 35)}"`);
  if (writer) log(`    찾는 작성자: ${writer}`);

  // 이모지 제거 후 앞 30자로 매칭 스니펫 준비
  const matchSnippet = normalizeReviewText(reviewText).substring(0, 30);
  if (!matchSnippet) {
    log(`  [순위 조회] 정규화 후 matchSnippet 비어있음 → 순위 조회 불가`);
    return { position: -1, policy: null, productUrl };
  }

  const reviewPage = await browser.newPage();
  try {
    await reviewPage.setUserAgent(
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36'
    );
    await reviewPage.goto(productUrl + '#REVIEW', { waitUntil: 'networkidle2', timeout: 30000 });
    await sleep(2000);

    const API_URL    = 'https://brand.naver.com/n/v1/contents/reviews/query-pages';
    const PAGE_SIZE  = 20;
    const TOP_LIMIT  = 100;                     // 100위까지만 조회
    const MAX_PAGES  = Math.ceil(TOP_LIMIT / PAGE_SIZE); // = 5
    const matchKor   = koreanOnlyKey(matchSnippet);

    let globalPos  = 0;
    let foundPos   = -1;
    let retryCount = 0;
    let lastSeenText = '';

    for (let p = 1; p <= MAX_PAGES; p++) {
      const body = { checkoutMerchantNo, originProductNo, page: p, pageSize: PAGE_SIZE, reviewSearchSortType: 'REVIEW_RANKING' };

      const result = await reviewPage.evaluate(async (url, reqBody) => {
        try {
          const resp = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(reqBody),
          });
          if (!resp.ok) return { error: `HTTP ${resp.status}` };
          return await resp.json();
        } catch(e) { return { error: e.message }; }
      }, API_URL, body);

      if (result.error) {
        if (retryCount < 2) {
          retryCount++;
          log(`  [순위 조회 API 오류] ${result.error} → ${retryCount}회 재시도...`);
          await sleep(1000 * retryCount);
          p--; // 같은 페이지 재시도
          continue;
        }
        log(`  [순위 조회 API 오류] ${result.error} (3회 실패, 중단)`);
        break;
      }
      retryCount = 0;

      const reviews = result.contents || result.reviews || result.list || result.data || [];
      if (!reviews.length) break;

      // 첫 페이지에서 실제 API 응답 필드명 및 텍스트 샘플 확인
      if (p === 1 && reviews.length > 0) {
        const sampleKeys = Object.keys(reviews[0]).join(', ');
        const sampleText = reviews[0].reviewContent || reviews[0].reviewBody || reviews[0].body
                        || reviews[0].content || reviews[0].reviewText || reviews[0].text || '(없음)';
        log(`  [순위 조회] API 필드: ${sampleKeys}`);
        log(`  [순위 조회] 첫 리뷰 텍스트 샘플: "${sampleText.substring(0, 40)}"`);
        log(`  [순위 조회] 매칭 스니펫: "${matchSnippet}"`);
      }

      for (const rv of reviews) {
        globalPos++;
        if (globalPos > TOP_LIMIT) break; // 100위 컷오프
        const rawText = rv.reviewContent || rv.reviewBody || rv.body || rv.content
                     || rv.reviewText || rv.text || rv.message || '';
        const normText = normalizeReviewText(rawText);
        const normKor  = koreanOnlyKey(normText);
        lastSeenText = normText;

        // 작성자 매칭 (마스킹 ID 동등 비교) — false positive 차단의 1차 게이트
        const apiWriter = (rv.writerMemberId || rv.writerId || rv.maskedWriterId
                        || rv.memberMaskingId || rv.writer || '').toString();
        const writerMatched = writer && apiWriter && (
          apiWriter === writer ||
          apiWriter.replace(/\*/g, '') === writer.replace(/\*/g, '')
        );

        // 본문 매칭: 30자 → 15자 → 한글 시작 일치 fallback
        const textMatched =
          normText.includes(matchSnippet) ||
          matchSnippet.includes(normText.substring(0, 15)) ||
          normText.includes(matchSnippet.substring(0, 15)) ||
          (matchKor.length >= 4 && normKor.length >= 4 && (
            normKor.startsWith(matchKor) ||
            matchKor.startsWith(normKor.substring(0, Math.min(matchKor.length, normKor.length)))
          ));

        // 짧은 리뷰(한글 7자 미만)는 본문만으로 신뢰 어려움 → 작성자도 일치해야 매칭
        // 긴 리뷰는 본문 일치만으로도 매칭 OK (작성자 정보 없는 케이스 호환)
        const matched = textMatched && (matchKor.length >= 7 || writerMatched || !writer);

        if (matched) { foundPos = globalPos; break; }
      }
      if (foundPos > 0 || globalPos >= TOP_LIMIT) break;

      // 마지막 페이지 확인
      const total = result.totalCount || result.totalElements || result.total || result.count || result.totalReviews || 0;
      if (total > 0 && globalPos >= total) break;

      await sleep(300);
    }

    // 매칭 실패 처리:
    //   globalPos === 0 → 진짜 미확인 (API 오류/리뷰 0개) — position = -1
    //   globalPos > 0  → 어쨌든 상위 100개(또는 전체) 안에서 미발견 = 100위 밖 — position = -2
    if (foundPos <= 0) {
      foundPos = (globalPos > 0) ? -2 : -1;
    }
    const outOfTop100 = (foundPos === -2);

    if (foundPos === -1 && lastSeenText) {
      log(`  [순위 조회 디버그] 매칭 실패. 찾던 스니펫: "${matchSnippet}" / 한글키: "${matchKor}"`);
      log(`  [순위 조회 디버그] 마지막 API 텍스트: "${lastSeenText.substring(0, 60)}"`);
    }

    const policy = foundPos === -2 ? '✅ 100위 밖 → 답변으로 충분'
      : foundPos <= 0 ? null
      : foundPos <= 10 ? '⚠️ 1~10위 → 아주 적극 대응 (조건 환불 검토)'
      : foundPos <= 20 ? '🔶 11~20위 → 적극 대응 (답변 or 환불)'
      : foundPos <= 40 ? '🔷 21~40위 → 답변 우선 (필요시 환불)'
      :                  '✅ 41위 이하 → 답변으로 충분';

    if (foundPos > 0) {
      log(`  📍 리뷰 순위: ${foundPos}번째 (총 검색: ${globalPos}개)`);
      log(`  📋 대응 정책: ${policy}`);
    } else if (outOfTop100) {
      log(`  📍 리뷰 순위: 100위 밖 (상위 ${globalPos}개 검색 후 미발견)`);
    } else {
      log(`  [순위 조회] 리뷰를 찾지 못함 (API 오류 또는 매핑 문제. 검색된 리뷰: ${globalPos}개)`);
    }

    return { position: foundPos, policy, productUrl, searchedCount: globalPos };

  } catch(e) {
    log(`  [순위 조회 오류] ${e.message}`);
    return { position: -1, policy: null, productUrl };
  } finally {
    await reviewPage.close().catch(() => {});
  }
}

// ─────────────────────────────────────────────────────────
// 메인
// ─────────────────────────────────────────────────────────
async function main() {
  log('');
  log('╔══════════════════════════════════════════════════════════════╗');
  log('║   코에르 리뷰 미답변 자동 답변 등록 (Claude Sonnet API)     ║');
  log('║   조건: 1주일 이내 작성 + 답글미등록                        ║');
  log('╚══════════════════════════════════════════════════════════════╝');
  log('');

  // 스케줄러 영업일 가드 (--scheduled 플래그 있을 때만)
  await runScheduleGuard();

  // 중복 실행 방지
  const lockFile = path.join(__dirname, '.posting.lock');
  if (fs.existsSync(lockFile)) {
    const age = (Date.now() - fs.statSync(lockFile).mtimeMs) / 60000;
    if (age < 30) { log('[오류] 이미 실행 중입니다. .posting.lock 파일을 삭제 후 재시도하세요.'); process.exit(1); }
    fs.unlinkSync(lockFile);
  }
  fs.writeFileSync(lockFile, String(process.pid));

  const browser = await puppeteer.launch({
    headless:        CONFIG.headless,
    protocolTimeout: 120000,
    args:            ['--no-sandbox', '--disable-setuid-sandbox', '--lang=ko-KR,ko',
                      '--window-size=1600,900'],
    defaultViewport: { width: 1600, height: 900 },
  });

  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36');
  page.on('dialog', async d => { log(`  [알림] ${d.message()}`); await d.accept(); });

  let successCount = 0;
  let refundCount  = 0;
  let failCount    = 0;
  let totalDone    = 0;
  const results    = [];

  try {
    // ── 로그인 ────────────────────────────────────────
    log('[인증] 셀러센터 로그인...');
    await loginToSellerCenter(page);
    log('');

    // ── 스토어 확인/전환 (코에르) ─────────────────────
    log('[스토어] 현재 스토어 확인...');
    try {
      await ensureCoeirStore(page);
    } catch (e) {
      log(`  [경고] 스토어 확인/전환 실패: ${e.message}`);
      log('  [경고] 현재 스토어로 진행하지만, 코에르가 아닐 경우 후속 단계가 실패할 수 있음');
    }
    log('');

    // ── 리뷰 페이지 이동 ──────────────────────────────
    log('[리뷰 페이지] 로드 중...');
    await page.goto(CONFIG.reviewUrl, { waitUntil: 'networkidle2', timeout: 30000 });
    try {
      await page.waitForFunction(
        () => !!Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색'),
        { timeout: 45000 }
      );
    } catch (e) {
      log(`  [실패] 검색 버튼 렌더 대기 타임아웃 (45초)`);
      await dumpDiagnostics(page, 'review_page_load');
      throw e;
    }
    await sleep(1500);

    // ── 날짜 범위: 오늘 ──────────────────────────────
    for (let t = 0; t < 5; t++) {
      const r = await page.evaluate(() => {
        const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '1주일');
        if (btn) { btn.click(); return true; }
        return false;
      });
      if (r) { log('  날짜 범위: 오늘 선택'); break; }
      await sleep(600);
    }
    await sleep(800);

    // ── 답글여부: 답글미등록 ─────────────────────────
    // UI는 Selectize.js 커스텀 드롭다운 사용 (.selectize-control/.selectize-input/.selectize-dropdown-content)
    log('  답글여부 필터 렌더링 대기 중...');
    await page.waitForFunction(() => {
      // Selectize 드롭다운에 '답글미등록' 옵션이 있을 때까지 대기
      const hasSelectize = Array.from(document.querySelectorAll('.selectize-dropdown-content .option'))
        .some(el => el.textContent.trim().includes('답글미등록'));
      const hasLabel = document.body.innerText.includes('답글여부');
      return hasSelectize || hasLabel;
    }, { timeout: 45000 }).catch(async () => {
      log('  [경고] 답글여부 필터 렌더링 타임아웃 (45초)');
      await dumpDiagnostics(page, 'reply_filter_render');
    });
    await sleep(800);

    let replyFilterOk = false;

    // Selectize.js 방식: '답글미등록' 옵션을 포함한 .selectize-control 찾아 클릭
    const selectizeHandle = await page.evaluateHandle(() => {
      const controls = Array.from(document.querySelectorAll('.selectize-control'));
      return controls.find(c => {
        const dropdown = c.querySelector('.selectize-dropdown-content');
        return dropdown && dropdown.textContent.includes('답글미등록');
      }) || null;
    });
    const selectizeEl = selectizeHandle.asElement();

    if (selectizeEl) {
      // 1. Selectize input 영역 클릭 → 드롭다운 열기
      const inputHandle = await selectizeEl.$('.selectize-input');
      if (inputHandle) {
        await inputHandle.click();
        // 드롭다운이 실제로 열릴 때까지 최대 3초 대기 (아침 부하 시 500ms 고정으로는 부족)
        for (let i = 0; i < 15; i++) {
          const open = await page.evaluate(() => {
            const d = document.querySelector('.selectize-dropdown-content');
            return d && d.offsetHeight > 0;
          });
          if (open) break;
          await sleep(200);
        }
        await sleep(200); // 옵션 렌더링 완료 buffer
      }

      // 2. 드롭다운 옵션에서 '답글미등록' 클릭
      replyFilterOk = await page.evaluate(() => {
        const opts = Array.from(document.querySelectorAll('.selectize-dropdown-content .option'));
        const target = opts.find(el => el.textContent.trim().includes('답글미등록'));
        if (target) { target.click(); return true; }
        return false;
      });

      if (replyFilterOk) {
        log('  ✓ 답글여부: 답글미등록 (Selectize)');
      } else {
        // 드롭다운이 안 열렸을 수 있음 → 직접 data-value 세팅 시도
        replyFilterOk = await page.evaluate(() => {
          const controls = Array.from(document.querySelectorAll('.selectize-control'));
          const ctrl = controls.find(c => {
            const d = c.querySelector('.selectize-dropdown-content');
            return d && d.textContent.includes('답글미등록');
          });
          if (!ctrl) return false;
          // Selectize JS API 직접 호출
          const nativeSel = ctrl.previousElementSibling || ctrl.parentElement?.querySelector('select');
          if (nativeSel && nativeSel.selectize) {
            const opt = Array.from(nativeSel.options).find(o => o.text.includes('답글미등록'));
            if (opt) { nativeSel.selectize.setValue(opt.value); return true; }
          }
          // 마지막 수단: 드롭다운 option 강제 클릭
          const allOpts = Array.from(document.querySelectorAll('.selectize-dropdown .option'));
          const t = allOpts.find(el => el.textContent.trim().includes('답글미등록'));
          if (t) { t.click(); return true; }
          return false;
        });
        if (replyFilterOk) log('  ✓ 답글여부: 답글미등록 (Selectize API)');
      }
    }

    if (!replyFilterOk) {
      log('  [오류] 답글미등록 필터 적용 실패 → 안전을 위해 스크립트 종료');
      await browser.close();
      try { fs.unlinkSync(lockFile); } catch(e) {}
      process.exit(1);
    }
    await sleep(500);

    // ── 검색 ────────────────────────────────────────
    await page.evaluate(() => {
      const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색');
      if (btn) btn.click();
    });
    log('  검색 실행');
    await page.waitForFunction(
      () => document.querySelectorAll('.ag-center-cols-container .ag-row').length > 0,
      { timeout: 45000 }
    ).catch(async () => {
      log('  [경고] 그리드 행 렌더 타임아웃 (45초)');
      await dumpDiagnostics(page, 'grid_render');
    });
    await sleep(2000);

    const total = await page.evaluate(() => {
      const m = document.body.innerText.match(/총\s*([\d,]+)개/);
      return m ? parseInt(m[1].replace(/,/g, '')) : '?';
    });
    log(`  총 ${total}개 리뷰 로드됨 (답글미등록 필터 ${replyFilterOk ? '적용됨' : '미적용'})`);

    // AG Grid 렌더링 대기
    await page.waitForFunction(() => {
      const rows = Array.from(document.querySelectorAll('.ag-center-cols-container .ag-row'));
      return rows.filter(row =>
        Array.from(row.querySelectorAll('.ag-cell')).map(c => c.textContent.trim()).filter(t => t).length >= 4
      ).length >= Math.min(rows.length, 1);
    }, { timeout: 20000 }).catch(() => log('  [경고] 그리드 렌더링 타임아웃'));
    await sleep(1000);

    log('');
    log('[처리] 환불검토 판단 → 답변 생성 → 등록 시작...');
    log('(현재 보이는 행을 즉시 처리 — 가상스크롤 대응)');
    log('');

    // ── 스크롤 맨 위로 리셋 ──────────────────────────
    await page.evaluate(() => {
      const vp = document.querySelector('.ag-body-viewport');
      if (vp) vp.scrollTop = 0;
    });
    await sleep(800);

    // ── 발견 즉시 처리 루프 ──────────────────────────
    // AG Grid 가상스크롤: DOM에 현재 화면에 보이는 행만 존재
    // 해결책: 보이는 행 수집 → 첫 번째 미처리 행 즉시 클릭
    // 성공 시 행이 사라짐(답글미등록 필터) → 다음 행이 자동으로 올라옴
    // 환불/실패 행은 processedKeys로 추적 → 다음 루프에서 건너뜀
    const processedKeys = new Set(); // writer_reviewNo 기반 처리완료 행 추적
    let   staleCount    = 0;

    while (totalDone < CONFIG.MAX_REVIEWS) {
      // 현재 DOM에 렌더링된 행 수집
      const visible = await collectVisibleRows(page);

      // 아직 처리하지 않은 행 중 첫 번째 선택
      // (서버 측 '답글미등록' 필터 적용 → DOM의 모든 행은 미답변 상태)
      const nextRow = visible.find(r =>
        !processedKeys.has(`${r.writer}_${r.reviewNo || r.date}`)
      );

      if (!nextRow) {
        // 처리 가능한 행 없음 → 스크롤 다운 시도
        const moved = await scrollDown(page);
        await sleep(800);
        if (!moved) {
          staleCount++;
          if (staleCount >= 5) { log('  모든 행 처리 완료'); break; }
        } else {
          staleCount = 0;
        }
        continue;
      }

      staleCount = 0;
      const key = `${nextRow.writer}_${nextRow.reviewNo || nextRow.date}`;

      // ── 체크박스 비활성 행: 환불완료 등 → 총 개수 제외하고 조용히 패스 ──
      if (nextRow.checkboxDisabled) {
        processedKeys.add(key);
        log(`  [스킵] 체크박스 비활성 행 (환불완료 등) → 제외: ${nextRow.writer} | ${nextRow.productName?.substring(0, 20)}`);
        await sleep(200);
        continue;
      }

      log(`──────────────────────────────────────────`);
      log(`리뷰: ${nextRow.writer} | ${nextRow.date} | 별점 ${nextRow.rating}점`);
      log(`상품: ${nextRow.productName}`);
      log(`내용: "${nextRow.reviewText}"`);
      log('');

      const result = {
        no:             totalDone + 1,
        writer:         nextRow.writer,
        reviewNo:       nextRow.reviewNo,
        productName:    nextRow.productName,
        optionName:     nextRow.optionName || '',
        rating:         nextRow.rating,
        date:           nextRow.date,
        reviewText:     nextRow.reviewText,
        replyText:      '',
        refundCheck:     '-',
        reviewPosition:  0,    // brand.naver.com 랭킹순 순위 (전체 리뷰 공통)
        refundPolicy:    '',   // 대응 정책 텍스트 (환불검토 대상만)
        judgeLabel:      '',   // 판단: '답변' | '환불검토'
        judgeReason:     '',   // 판단 근거 (환불검토 vs 답변)
        judgeConfidence: null, // 판단 확신도 0~100
        productUrl:      '',   // 조회한 brand.naver.com 상품 URL
      };

      // ── 리뷰 순위 조회 (전체 공통) ─────────────────
      log(`  🔍 리뷰 순위 조회 중...`);
      try {
        const posResult = await findReviewPosition(browser, nextRow.productName, nextRow.reviewText, nextRow.writer);
        // position > 0 = 순위, -2 = 100위 밖, 그 외 = 미확인(0)
        result.reviewPosition = posResult.position > 0
          ? posResult.position
          : (posResult.position === -2 ? -2 : 0);
        result.productUrl     = posResult.productUrl || '';
        if (posResult.position > 0) log(`  📍 리뷰 순위: ${posResult.position}번째`);
        else                        log(`  📍 리뷰 순위: 미확인`);
        // 정책은 나중에 환불검토 확정 시 설정
        result._posPolicy = posResult.policy || '';
      } catch(e) {
        log(`  [순위 조회 실패] ${e.message}`);
      }

      // ── 리뷰 종합 판단 (환불검토 vs 답변) ──────────
      log('  [판단] 리뷰 종합 판단 중...');
      const judgment = await judgeReview({
        productName:    nextRow.productName,
        rating:         nextRow.rating,
        reviewText:     nextRow.reviewText,
        reviewPosition: result.reviewPosition,
      });
      result.judgeReason     = judgment.reason || '';
      result.judgeConfidence = judgment.confidence;
      result.judgeLabel      = judgment.needsRefund ? '환불검토' : '답변';
      log(`  🏷️ 판단: ${result.judgeLabel}`);
      log(`  💭 판단 근거: ${result.judgeReason}${result.judgeConfidence != null ? ` (confidence ${result.judgeConfidence})` : ''}`);

      if (judgment.needsRefund) {
        result.refundCheck  = '검토필요';
        result.refundPolicy = result._posPolicy || '순위 미확인';
        refundCount++;
        totalDone++;
        processedKeys.add(key);
        log(`  ⚠️  환불검토 대상 → 답변 생략 (별점: ${nextRow.rating}점)`);
        log(`  📋 대응 정책: ${result.refundPolicy}`);
        results.push(result);
        log('');
        await sleep(300);
        continue;
      }

      // ── Claude API 답변 생성 ───────────────────────
      log('  [Claude API] 답변 생성 중...');
      let replyText;
      try {
        replyText = await generateReply({
          productName: nextRow.productName || '코에르',
          rating:      nextRow.rating,
          reviewText:  nextRow.reviewText,
          writer:      nextRow.writer,
        });
        log(`  [생성된 답변] ${replyText}`);
        log(`  [글자수] ${replyText.length}자`);
      } catch (e) {
        log(`  [오류] Claude API 실패: ${e.message}`);
        processedKeys.add(key);
        failCount++;
        totalDone++;
        results.push(result);
        log('');
        continue;
      }

      // ── 답글 등록 (행이 DOM에 있을 때 즉시 클릭) ───
      log('  [등록] 답글 등록 중...');
      const postResult = await postReply(page, nextRow, replyText);

      if (postResult.success) {
        successCount++;
        result.replyText = replyText;
        processedKeys.add(key); // 성공 행도 반드시 추가 — 재처리 방지
        log(`  ✓ 등록 완료 (성공: ${successCount}개)`);
        // 검색 재실행 → 그리드 강제 새로고침
        // (선택 상태 초기화 + 답변된 행 자동 제거 — 가장 확실한 방법)
        await page.evaluate(() => {
          const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색');
          if (btn) btn.click();
        });
        await page.waitForFunction(
          () => document.querySelectorAll('.ag-center-cols-container .ag-row').length > 0,
          { timeout: 10000 }
        ).catch(() => {});
        // 그리드 스크롤 맨 위로 리셋
        await page.evaluate(() => {
          const vp = document.querySelector('.ag-body-viewport');
          if (vp) vp.scrollTop = 0;
        });
        await sleep(1000);
      } else {
        processedKeys.add(key); // 실패 행 재처리 방지
        failCount++;
        log(`  ✗ 실패: ${postResult.reason}`);
      }

      results.push(result);
      totalDone++;
      log('');
      await sleep(500);
    }

  } finally {
    // ── 결과 요약 ──────────────────────────────────
    log('');
    log('══════════════════════════════════════════');
    log(`최종 결과`);
    log(`  총 리뷰:    ${totalDone}개`);
    log(`  답변 성공:  ${successCount}개`);
    log(`  환불검토:   ${refundCount}개`);
    log(`  실패:       ${failCount}개`);
    log('══════════════════════════════════════════');

    // ── 최종 검증: 답글미등록 필터로 재검색 → 남은 건수로 확인 ─
    const postedCount = results.filter(r => r.replyText).length;
    if (postedCount > 0) {
      log('');
      log('[검증] 답글미등록 재검색으로 남은 건수 확인합니다...');

      // 페이지 완전 재로드
      await page.goto(CONFIG.reviewUrl, { waitUntil: 'networkidle2', timeout: 30000 });
      await page.waitForFunction(
        () => !!Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색'),
        { timeout: 15000 }
      );
      await sleep(1500);

      // 날짜 필터: 오늘
      for (let t = 0; t < 5; t++) {
        const r = await page.evaluate(() => {
          const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '1주일');
          if (btn) { btn.click(); return true; }
          return false;
        });
        if (r) break;
        await sleep(500);
      }
      await sleep(600);

      // 답글여부: 답글미등록 (Selectize.js 방식)
      await sleep(1500); // 렌더링 대기
      const vHandle = await page.evaluateHandle(() => {
        const controls = Array.from(document.querySelectorAll('.selectize-control'));
        return controls.find(c => {
          const d = c.querySelector('.selectize-dropdown-content');
          return d && d.textContent.includes('답글미등록');
        }) || null;
      });
      const vEl = vHandle.asElement();
      if (vEl) {
        const vInput = await vEl.$('.selectize-input');
        if (vInput) { await vInput.click(); await sleep(400); }
        await page.evaluate(() => {
          const opts = Array.from(document.querySelectorAll('.selectize-dropdown-content .option'));
          const t = opts.find(el => el.textContent.trim().includes('답글미등록'));
          if (t) t.click();
        });
      }
      await sleep(600);

      // 검색 실행
      await page.evaluate(() => {
        const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색');
        if (btn) btn.click();
      });
      await page.waitForFunction(
        () => !!document.body.innerText.match(/총\s*[\d,]+개/),
        { timeout: 15000 }
      ).catch(() => {});
      await sleep(2000);

      // 남은 답글미등록 건수 확인
      const remaining = await page.evaluate(() => {
        const m = document.body.innerText.match(/총\s*([\d,]+)개/);
        return m ? parseInt(m[1].replace(/,/g, '')) : -1;
      });

      log(`  처리 전 미답변 건수 기준: 성공 ${postedCount}개 등록 시도`);
      log(`  처리 후 남은 답글미등록: ${remaining}개`);

      if (remaining === 0) {
        log(`  ✅ 검증 완료: 미답변 리뷰 0개 — 모두 답변 등록됨`);
      } else if (remaining > 0) {
        log(`  ⚠️  아직 ${remaining}개 미답변 남음 (환불검토 ${refundCount}개 제외 시 ${Math.max(0, remaining - refundCount)}개 실패 추정)`);
      } else {
        log(`  ⚠️  건수 확인 불가`);
      }
    }

    // ── 리포트 생성 (Word + Slack) ─────────────────
    if (results.length > 0) {
      // 요약 JSON 저장
      const summary = {
        date:         new Date().toLocaleDateString('ko-KR'),
        totalReviews: totalDone,
        replied:      successCount,
        refund:       refundCount,
        failed:       failCount,
        usage:        getUsageSummary(),
        results:      results.map((r, i) => ({
          no:             i + 1,
          reviewNo:       r.reviewNo || '',
          writer:         r.writer,
          rating:         r.rating,
          productName:    r.productName,
          optionName:     r.optionName    || '',
          reviewText:     r.reviewText,
          replyText:      r.replyText || '',
          refundCheck:     r.refundCheck,
          reviewPosition:  r.reviewPosition || 0,
          refundPolicy:    r.refundPolicy   || '',
          judgeLabel:      r.judgeLabel     || '',
          judgeReason:     r.judgeReason    || '',
          judgeConfidence: r.judgeConfidence,
          productUrl:      r.productUrl     || '',
        })),
      };
      const summaryFile = path.join(__dirname, '.review_summary.json');
      fs.writeFileSync(summaryFile, JSON.stringify(summary, null, 2));
      log(`\n[요약] ${summaryFile} 저장됨`);

      // Word 문서 생성
      let docxFilepath = null;
      try {
        const { filename: docxName, filepath } = await generateWordDoc(summary);
        docxFilepath = filepath;
        log(`[Word] 저장 완료: ${docxName}`);
      } catch(e) {
        log(`[Word 오류] ${e.message}`);
      }

      // Slack 채널 자동 전송
      await sendSlackDM(summary);

      // Slack 채널에 워드 파일 첨부 업로드
      if (docxFilepath) {
        const channelId = cfg.SLACK_CHANNEL_ID || cfg.SLACK_USER_ID || 'U08KNE04HKK';
        await uploadFileToSlack(docxFilepath, channelId, `📎 ${summary.date} 보고서 워드 파일`);
      }
    }

    await browser.close();
    try { fs.unlinkSync(lockFile); } catch(e) {}
  }
}

// ─────────────────────────────────────────────────────────
// 모듈 export — 다른 스크립트(`collect_pending.js`, `post_replies.js`,
// 에이전트 routine 등)가 puppeteer/리포트 헬퍼를 재사용할 수 있도록 노출
// ─────────────────────────────────────────────────────────
module.exports = {
  // 설정
  CONFIG,
  // 로그인 / 네비게이션
  loginToSellerCenter,
  ensureCoeirStore,
  switchToCoeirStore,
  dismissBlockingModals,
  // 그리드 / 행 조작
  collectVisibleRows,
  scrollDown,
  deselectAllRows,
  getCheckedCount,
  clickRowCheckbox,
  clickReplyButton,
  typeReplyText,
  clickSubmit,
  checkAndClosePopup,
  closeModal,
  postReply,
  // brand.naver.com 순위 조회
  findReviewPosition,
  getProductNaverIds,
  findProductNaverId,
  normalizeReviewText,
  koreanOnlyKey,
  // 리포트
  generateWordDoc,
  sendSlackDM,
  uploadFileToSlack,
  sendSlackError,
  // 지식 / 분류 (옵션 — 에이전트가 직접 추론할 때 참고용)
  getProductKnowledge,
  // 토큰 사용량 추적 (에이전트 매개 실행에선 미사용)
  accumulateUsage,
  getUsageSummary,
  // 영업일 가드 (주말/공휴일 자동 스킵)
  shouldSkipToday,
  getKoreanHolidays,
  // 표시 헬퍼
  formatReviewPosition,
  resolveDisplayOption,
  rankSortKey,
  sortByRank,
  shouldShowJudgeReason,
};

// CLI 직접 실행 시에만 main() 호출 — `require()` 로 import 할 땐 자동 실행 안 됨
if (require.main === module) {
  main().catch(async err => {
    console.error('\n[오류]', err.message);
    const lockFile = path.join(__dirname, '.posting.lock');
    try { fs.unlinkSync(lockFile); } catch(e) {}
    // 스케줄 실행 중 오류면 Slack으로 즉시 알림
    if (process.argv.includes('--scheduled')) {
      await sendSlackError(err.message).catch(() => {});
    }
    process.exit(1);
  });
}
