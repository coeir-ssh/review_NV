/**
 * 코에르 리뷰 답변 셀러센터 자동 포스팅 (Puppeteer + AG Grid)
 *
 * [구조]
 *  1. 엑셀에서 답변 대상 로드 → (idPrefix+date) 매칭 맵 구성
 *  2. 셀러센터 리뷰 페이지 로드
 *  3. AG Grid 행을 스크롤하며 순회 → 매칭 행마다 바로 답글 등록
 *  4. 페이지네이션 처리 (다음 버튼)
 *
 * 실행: node post_to_naver_puppeteer.js
 */

const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const readline = require('readline');
const path = require('path');
const fs = require('fs');
const cfg = require('./config');

// ─────────────────────────────────────────────────────────
// 설정
// ─────────────────────────────────────────────────────────
const CONFIG = {
  outputDir:   'C:\\Users\\AWESOMATIC\\Desktop\\코에르\\클로드\\리뷰수집프로그램',
  sessionFile: path.join(__dirname, '.seller_session.json'),
  reviewUrl:   'https://sell.smartstore.naver.com/#/review/search',
  headless:    false,
  replyDelay:  2500,
  scrollWait:  1200,     // 스크롤 후 렌더링 대기 (ms)
  maxScrollAttempts: 200,
};

const log = msg => console.log(msg);
const sleep = ms => new Promise(r => setTimeout(r, ms));

// ─────────────────────────────────────────────────────────
// 유틸
// ─────────────────────────────────────────────────────────
function findLatestResponseExcel() {
  const files = fs.readdirSync(CONFIG.outputDir)
    .filter(f => f.startsWith('리뷰답변_자동_') && f.endsWith('.xlsx'))
    .sort().reverse();
  return files.length ? path.join(CONFIG.outputDir, files[0]) : null;
}

function askQuestion(prompt) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise(resolve => rl.question(prompt, ans => { rl.close(); resolve(ans.trim()); }));
}

function saveSession(cookies) { fs.writeFileSync(CONFIG.sessionFile, JSON.stringify(cookies)); }
function loadSession() {
  try { if (fs.existsSync(CONFIG.sessionFile)) return JSON.parse(fs.readFileSync(CONFIG.sessionFile)); }
  catch (e) {}
  return null;
}

// "ssw0****" → "ssw0"
function idPrefix(id) { return (id || '').split('*')[0].trim(); }

// "2026.04.12. 17:59" → "2026.04.12"
function dateOnly(d) {
  return (d || '')
    .replace(/\s*\d{1,2}:\d{2}.*$/, '')
    .replace(/,.*$/, '')
    .replace(/\.\s*$/, '')
    .trim();
}

// ─────────────────────────────────────────────────────────
// 로그인 확인 (콘텐츠 기반)
// ─────────────────────────────────────────────────────────
async function isLoggedIn(page) {
  try {
    await page.waitForFunction(() => {
      const t = document.body.innerText;
      return t.includes('상품관리') || t.includes('리뷰관리') || t.includes('주문관리') ||
             !!document.querySelector('input[placeholder="아이디 또는 이메일 주소"]');
    }, { timeout: 12000 });
  } catch (e) {}
  return page.evaluate(() => {
    const t = document.body.innerText;
    return t.includes('상품관리') || t.includes('리뷰관리') || t.includes('주문관리');
  });
}

// ─────────────────────────────────────────────────────────
// 로그인
// ─────────────────────────────────────────────────────────
async function loginToSellerCenter(page) {
  const savedCookies = loadSession();
  if (savedCookies) {
    log('  저장된 세션 확인 중...');
    await page.setCookie(...savedCookies);
    await page.goto('https://sell.smartstore.naver.com/', { waitUntil: 'networkidle2', timeout: 30000 });
    if (await isLoggedIn(page)) { log('  ✓ 세션 로그인 성공'); return; }
    log('  세션 만료 → 재로그인');
    try { fs.unlinkSync(CONFIG.sessionFile); } catch (e) {}
  }

  log('  커머스 로그인 페이지 접속...');
  await page.goto('https://accounts.commerce.naver.com/login', { waitUntil: 'networkidle2', timeout: 20000 });
  await sleep(1500);

  const idField = await page.$('input[placeholder="아이디 또는 이메일 주소"]');
  const pwField = await page.$('input[type="password"]');
  if (idField && pwField) {
    log('  ID/PW 입력 중...');
    await idField.click({ clickCount: 3 });
    await idField.type(cfg.SELLER_ID, { delay: 80 });
    await sleep(300);
    await pwField.click({ clickCount: 3 });
    await pwField.type(cfg.SELLER_PW, { delay: 80 });
    await sleep(300);
    await page.evaluate(() => {
      const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '로그인');
      if (btn) btn.click();
    });
    await page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 20000 }).catch(() => {});
    await sleep(2000);
  } else {
    log('  [경고] 로그인 폼 없음 - 브라우저에서 직접 로그인 후 Enter');
  }

  if (!(await isLoggedIn(page))) {
    log('');
    log('  ┌──────────────────────────────────────────────────────┐');
    log('  │  [추가 인증] OTP/인증 완료 후 셀러센터가 보이면      │');
    log('  │  터미널에서 Enter 를 누르세요                        │');
    log('  └──────────────────────────────────────────────────────┘');
    await askQuestion('\n  인증 완료 후 Enter ▶ ');
    await sleep(2000);
  }

  const cookies = await page.cookies();
  saveSession(cookies);
  log('  ✓ 로그인 완료 / 세션 저장');
}

// ─────────────────────────────────────────────────────────
// 현재 보이는 AG Grid 행 수집
// ─────────────────────────────────────────────────────────
async function collectVisibleRows(page) {
  return page.evaluate(() => {
    const rows = Array.from(document.querySelectorAll('.ag-center-cols-container .ag-row'));
    return rows.map(row => {
      const cells = Array.from(row.querySelectorAll('.ag-cell'));
      const texts = cells.map(c => c.textContent.trim()).filter(t => t);

      // 리뷰글번호: 8~12자리 숫자
      const reviewNo = texts.find(t => /^\d{8,12}$/.test(t)) || '';
      // 구매자 아이디: 별표 포함 (숫자로 시작하는 ID도 허용: 0907****)
      const writer = texts.find(t => t.includes('*') && t.length >= 3) || '';
      // 날짜
      const date = texts.find(t => /\d{4}\.\d{2}\.\d{2}/.test(t)) || '';
      // 답글 있음 여부 ('Y' 단독은 false positive 가능하므로 제외)
      const hasReply = texts.some(t => t.includes('답글있음') || t.includes('답변완료') || t.includes('답변있음'));
      const top = row.style.top;

      return { reviewNo, writer, date, top, hasReply, allTexts: texts };
    }).filter(r => r.date || r.reviewNo);
  });
}

// ─────────────────────────────────────────────────────────
// 답글여부 필터 설정 (미답변)
// ─────────────────────────────────────────────────────────
async function trySetUnrepliedFilter(page) {
  try {
    // Angular Material mat-select 클릭 방식으로 시도
    const result = await page.evaluate(() => {
      // "답글여부" 라벨 근처의 mat-select 또는 select 찾기
      const allEls = Array.from(document.querySelectorAll('label, mat-label, th, td, span, div'));
      for (const el of allEls) {
        if (el.children.length === 0 && el.textContent.trim() === '답글여부') {
          let container = el.parentElement;
          for (let i = 0; i < 6; i++) {
            if (!container) break;
            // native select
            const sel = container.querySelector('select');
            if (sel) {
              const opt = Array.from(sel.options).find(o =>
                o.value === 'N' || o.text.includes('미답변')
              );
              if (opt) {
                sel.value = opt.value;
                sel.dispatchEvent(new Event('change', {bubbles: true}));
                return 'native-select';
              }
            }
            // mat-select trigger
            const matSel = container.querySelector('mat-select');
            if (matSel) {
              matSel.click();
              return 'mat-select-clicked';
            }
            container = container.parentElement;
          }
        }
      }
      return 'not-found';
    });

    if (result === 'mat-select-clicked') {
      await sleep(800);
      // 옵션 패널에서 미답변 클릭
      const clicked = await page.evaluate(() => {
        const opts = document.querySelectorAll('mat-option');
        for (const opt of opts) {
          if (opt.textContent.trim().includes('미답변') || opt.textContent.trim() === 'N') {
            opt.click();
            return true;
          }
        }
        return false;
      });
      await sleep(500);
      return clicked ? '미답변 설정' : '옵션 없음';
    }
    return result;
  } catch (e) {
    return `오류: ${e.message}`;
  }
}

// ─────────────────────────────────────────────────────────
// 검색 버튼 클릭
// ─────────────────────────────────────────────────────────
async function clickSearchButton(page) {
  await page.evaluate(() => {
    const btn = Array.from(document.querySelectorAll('button'))
      .find(b => b.textContent.trim() === '검색');
    if (btn) btn.click();
  });
}

// ─────────────────────────────────────────────────────────
// AG Grid 스크롤 (한 단계)
// ─────────────────────────────────────────────────────────
async function scrollDown(page) {
  return page.evaluate(() => {
    const selectors = [
      '.ag-body-viewport',
      '.ag-center-cols-viewport',
      '.ag-body-horizontal-scroll-viewport',
    ];
    for (const sel of selectors) {
      const vp = document.querySelector(sel);
      if (vp && vp.scrollHeight > vp.clientHeight) {
        const before = vp.scrollTop;
        vp.scrollTop += Math.max(vp.clientHeight * 0.8, 300);
        if (vp.scrollTop !== before) return { moved: true, sel };
      }
    }
    // 전체 페이지 스크롤도 시도
    const before = window.scrollY;
    window.scrollBy(0, 400);
    if (window.scrollY !== before) return { moved: true, sel: 'window' };
    return { moved: false, sel: null };
  });
}

// ─────────────────────────────────────────────────────────
// 다음 페이지 버튼 클릭 (페이지네이션)
// currentPage: 현재 페이지 번호(1부터 시작)를 직접 전달받아 단순하게 클릭
// ─────────────────────────────────────────────────────────
async function clickNextPage(page, currentPage) {
  const nextPage = currentPage + 1;

  const info = await page.evaluate((nextPageStr) => {
    const allBtns = Array.from(document.querySelectorAll('button, a, [role="button"]'));

    // ag-grid 내장 페이지네이션
    const agNext = document.querySelector('.ag-paging-button[ref="btNext"]:not(.ag-disabled)');
    if (agNext) { agNext.click(); return { clicked: true, method: 'ag-grid' }; }

    // "다음" 텍스트 버튼
    const NEXT_LABELS = ['다음', '>', '›', '>>', '다음 페이지'];
    for (const btn of allBtns) {
      const t = btn.textContent.trim();
      const disabled = btn.disabled || btn.classList.contains('disabled') ||
                       btn.getAttribute('aria-disabled') === 'true' || btn.hasAttribute('disabled');
      if (!disabled && NEXT_LABELS.includes(t)) {
        btn.click();
        return { clicked: true, method: `text:${t}` };
      }
    }

    // 직접 다음 페이지 번호 버튼 클릭 (disabled 체크 없이 텍스트만으로)
    for (const btn of allBtns) {
      const t = btn.textContent.trim();
      // 버튼 텍스트가 정확히 nextPageStr이거나 nextPageStr로 시작하는 경우
      if (t === nextPageStr || t.startsWith(nextPageStr + ' ') || t.startsWith(nextPageStr + '(')) {
        btn.click();
        return { clicked: true, method: `page-btn:${nextPageStr}` };
      }
    }

    // 다음 페이지 버튼 존재 여부 확인 (존재하지 않으면 마지막 페이지)
    const hasAnyPageBtn = allBtns.some(b => {
      const t = b.textContent.trim();
      return /^\d+/.test(t) && b.getBoundingClientRect().top > window.innerHeight * 0.5;
    });

    // 디버그: 하단 버튼 정보 수집
    const bottomBtns = allBtns
      .filter(b => b.getBoundingClientRect().top > window.innerHeight * 0.5)
      .map(b => ({
        text: b.textContent.trim().substring(0, 20),
        ariaLabel: (b.getAttribute('aria-label') || '').substring(0, 20),
        ariaCurrent: b.getAttribute('aria-current'),
        disabled: b.disabled,
        cls: b.className.substring(0, 20),
      }))
      .filter(b => b.text || b.ariaLabel);
    return { clicked: false, nextPageStr, hasAnyPageBtn, bottomBtns: bottomBtns.slice(0, 20) };
  }, String(nextPage));

  if (!info.clicked) {
    if (info.hasAnyPageBtn === false) {
      log(`  [페이지네이션] 마지막 페이지 (${currentPage})`);
    } else {
      log(`  [페이지네이션 디버그] next="${info.nextPageStr}" bottomBtns:${JSON.stringify(info.bottomBtns)}`);
    }
  } else {
    log(`  [다음 페이지] → ${info.method}`);
  }
  return info.clicked;
}

// ─────────────────────────────────────────────────────────
// AG Grid 전체 행 선택 해제 (다음 행 선택 전 반드시 호출)
// ─────────────────────────────────────────────────────────
async function deselectAllRows(page) {
  await page.evaluate(() => {
    // 방법 1: 선택된 행의 체크박스를 모두 클릭해서 해제
    const selectedRows = Array.from(document.querySelectorAll(
      '.ag-pinned-left-cols-container .ag-row-selected, .ag-center-cols-container .ag-row-selected'
    ));
    for (const row of selectedRows) {
      const cb = row.querySelector('input[type="checkbox"]');
      if (cb && cb.checked) { cb.click(); }
      else { row.click(); }
    }
    // 방법 2: 헤더 체크박스로 전체 해제 (선택된 행이 있을 경우)
    if (document.querySelectorAll('.ag-row-selected').length > 0) {
      const headerCb = document.querySelector(
        '.ag-header-cell .ag-checkbox-input, .ag-pinned-left-header .ag-checkbox-input'
      );
      if (headerCb) {
        // 현재 선택된 상태면 한 번 클릭으로 전체 해제 시도
        headerCb.click(); // 전체 선택 or 전체 해제
        if (document.querySelectorAll('.ag-row-selected').length > 0) {
          headerCb.click(); // 상태 반전
        }
      }
    }
  });
  await sleep(300);
}

// ─────────────────────────────────────────────────────────
// 특정 top 위치의 행 체크박스 클릭 (실제 마우스 클릭 사용)
// ─────────────────────────────────────────────────────────
async function clickRowCheckbox(page, rowTop) {
  // evaluateHandle로 체크박스 엘리먼트 핸들 획득
  const handle = await page.evaluateHandle((top) => {
    const pinnedRows = Array.from(document.querySelectorAll('.ag-pinned-left-cols-container .ag-row'));
    const target = pinnedRows.find(r => r.style.top === top);
    if (target) {
      const cb = target.querySelector('input[type="checkbox"]');
      if (cb) return cb;
      const cell = target.querySelector('.ag-cell');
      if (cell) return cell;
    }
    const centerRows = Array.from(document.querySelectorAll('.ag-center-cols-container .ag-row'));
    const centerRow = centerRows.find(r => r.style.top === top);
    if (centerRow) return centerRow;
    return null;
  }, rowTop);

  const el = handle.asElement();
  if (!el) {
    // fallback: evaluate click
    return page.evaluate((top) => {
      const rows = Array.from(document.querySelectorAll('.ag-center-cols-container .ag-row'));
      const row = rows.find(r => r.style.top === top);
      if (row) { row.click(); return 'fallback-click'; }
      return false;
    }, rowTop);
  }

  try {
    // 화면에 스크롤 후 실제 클릭
    await page.evaluate(e => e.scrollIntoView({ behavior: 'instant', block: 'center' }), el);
    await sleep(300);
    await el.click();
    return true;
  } catch (e) {
    // 엘리먼트 클릭 실패 시 evaluate fallback
    return page.evaluate((top) => {
      const rows = Array.from(document.querySelectorAll('.ag-pinned-left-cols-container .ag-row, .ag-center-cols-container .ag-row'));
      const row = rows.find(r => r.style.top === top);
      if (row) { row.click(); return 'eval-fallback'; }
      return false;
    }, rowTop);
  }
}

// ─────────────────────────────────────────────────────────
// 답글작성 버튼 클릭 (실제 마우스 클릭 사용)
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
  } catch (e) {
    // fallback evaluate click
    return page.evaluate(() => {
      const btn = Array.from(document.querySelectorAll('button')).find(b =>
        (b.textContent.trim() === '답글작성' || b.textContent.trim() === '답글 작성') && !b.disabled
      );
      if (btn) { btn.click(); return true; }
      return false;
    });
  }
}

// ─────────────────────────────────────────────────────────
// 열린 모달/다이얼로그 닫기 (Escape)
// ─────────────────────────────────────────────────────────
async function closeModal(page) {
  try {
    await page.keyboard.press('Escape');
    await sleep(800);
    // 닫기(X) 버튼이 있으면 클릭
    await page.evaluate(() => {
      const closeBtn = Array.from(document.querySelectorAll('button')).find(b => {
        const t = b.textContent.trim();
        return t === '닫기' || t === '취소' || t === '×' || t === 'X';
      });
      if (closeBtn) closeBtn.click();
    });
    await sleep(500);
  } catch (e) {}
}

// ─────────────────────────────────────────────────────────
// 모달 textarea/contenteditable에 답변 입력
// Bootstrap .seller-layer-modal 내의 입력 요소를 직접 타겟
// ─────────────────────────────────────────────────────────
async function typeReplyText(page, text) {
  let el = null;
  let isContentEditable = false;

  for (let i = 0; i < 15; i++) {
    const handle = await page.evaluateHandle(() => {
      // 1순위: .seller-layer-modal 내 textarea
      const selLayerModal = document.querySelector('.seller-layer-modal');
      if (selLayerModal) {
        const ta = selLayerModal.querySelector('textarea:not([readonly]):not([disabled])');
        if (ta) return ta;
        const ce = selLayerModal.querySelector('[contenteditable="true"]');
        if (ce) return ce;
      }

      // 2순위: 모달/다이얼로그 내 textarea
      const modalSelectors = [
        '.modal.show', '.modal.fade', '[role="dialog"]',
        'mat-dialog-container', '.cdk-overlay-pane', '.layer-popup', '.popup-wrap'
      ];
      for (const sel of modalSelectors) {
        const modal = document.querySelector(sel);
        if (modal) {
          const ta = modal.querySelector('textarea:not([readonly]):not([disabled])');
          if (ta) return ta;
          const ce = modal.querySelector('[contenteditable="true"]');
          if (ce) return ce;
        }
      }

      // 3순위: rect 크기가 있는 모든 textarea
      const allTA = document.querySelectorAll('textarea:not([readonly]):not([disabled])');
      for (const ta of allTA) {
        const r = ta.getBoundingClientRect();
        if (r.width > 50 && r.height > 20) return ta;
      }

      // 4순위: contenteditable 요소
      const allCE = document.querySelectorAll('[contenteditable="true"]');
      for (const ce of allCE) {
        const r = ce.getBoundingClientRect();
        if (r.width > 50 && r.height > 20) return ce;
      }

      return null;
    });
    el = handle.asElement();
    if (el) {
      isContentEditable = await page.evaluate(e =>
        e.getAttribute('contenteditable') === 'true', el
      ).catch(() => false);
      break;
    }
    if (i < 14) await sleep(400);
  }

  if (!el) return false;

  // 스크롤 + 포커스
  await page.evaluate(e => {
    e.scrollIntoView({ behavior: 'instant', block: 'center' });
    e.focus();
    if (e.tagName === 'TEXTAREA') {
      e.select();
      e.value = '';
    } else {
      // contenteditable
      const range = document.createRange();
      range.selectNodeContents(e);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }
  }, el);
  await sleep(300);

  // 기존 내용 삭제
  await page.keyboard.down('Control');
  await page.keyboard.press('a');
  await page.keyboard.up('Control');
  await page.keyboard.press('Delete');
  await sleep(150);

  // Angular ng-model 감지를 위해 keyboard.type 사용
  await page.keyboard.type(text, { delay: 10 });

  // Angular reactive forms를 위해 input 이벤트도 dispatch
  await page.evaluate(e => {
    e.dispatchEvent(new Event('input', { bubbles: true }));
    e.dispatchEvent(new Event('change', { bubbles: true }));
  }, el);

  return true;
}

// ─────────────────────────────────────────────────────────
// 등록 버튼 클릭 (다이얼로그 내 우선)
// ─────────────────────────────────────────────────────────
async function clickSubmit(page) {
  return page.evaluate(() => {
    const SUBMIT_LABELS = ['등록', '저장', '확인', '작성완료', '등록하기', '저장하기'];
    // 다이얼로그 내 버튼 우선
    const dialogContainers = [
      ...document.querySelectorAll('.seller-layer-modal'),
      ...document.querySelectorAll('mat-dialog-container'),
      ...document.querySelectorAll('[role="dialog"]'),
      ...document.querySelectorAll('.cdk-overlay-pane'),
      ...document.querySelectorAll('.modal.show'),
      ...document.querySelectorAll('.layer-popup'),
    ];
    for (const c of dialogContainers) {
      const btn = Array.from(c.querySelectorAll('button')).find(b =>
        SUBMIT_LABELS.includes(b.textContent.trim()) && !b.disabled &&
        b.getAttribute('aria-disabled') !== 'true'
      );
      if (btn) { btn.click(); return `dialog-btn: ${btn.textContent.trim()}`; }
    }
    // fallback: 페이지 전체에서 찾기
    const btn = Array.from(document.querySelectorAll('button')).find(b =>
      SUBMIT_LABELS.includes(b.textContent.trim()) && !b.disabled &&
      b.getAttribute('aria-disabled') !== 'true'
    );
    if (btn) { btn.click(); return `fallback-btn: ${btn.textContent.trim()}`; }
    return false;
  });
}

// ─────────────────────────────────────────────────────────
// 행 처리: 체크박스 → 답글작성 → 입력 → 등록
// ─────────────────────────────────────────────────────────
async function processRow(page, row, replyText, debugMode) {
  // 0. 이전 선택 모두 해제 (누적 선택 방지)
  await deselectAllRows(page);

  // 1. 해당 행 체크박스 클릭
  const checkResult = await clickRowCheckbox(page, row.top);
  if (!checkResult) return { success: false, reason: '체크박스 클릭 실패' };
  await sleep(800);

  try {
    if (debugMode) {
      await page.screenshot({ path: path.join(CONFIG.outputDir, 'debug_1_checkbox.png') }).catch(() => {});
      const state = await page.evaluate((top) => {
        const rows = Array.from(document.querySelectorAll('.ag-pinned-left-cols-container .ag-row'));
        const r = rows.find(r => r.style.top === top);
        if (!r) return { found: false, searchedTop: top };
        const cb = r.querySelector('input[type="checkbox"]');
        const btns = Array.from(document.querySelectorAll('button'))
          .filter(b => !b.disabled)
          .map(b => b.textContent.trim())
          .filter(t => t.length < 20);
        return { found: true, checked: cb ? cb.checked : 'no-cb', enabledBtns: btns.slice(0, 15) };
      }, row.top);
      log(`  [디버그 체크박스] ${JSON.stringify(state)}`);
    }
  } catch (dbgErr) {
    log(`  [디버그오류] 체크박스: ${dbgErr.message}`);
  }

  // 2. 답글작성 버튼 클릭 (실제 mouse click)
  const replyBtnClicked = await clickReplyButton(page);

  if (debugMode) {
    const allBtns = await page.evaluate(() =>
      Array.from(document.querySelectorAll('button'))
        .map(b => ({ text: b.textContent.trim().substring(0, 20), disabled: b.disabled }))
        .filter(b => b.text)
        .slice(0, 15)
    );
    log(`  [디버그 버튼목록] ${JSON.stringify(allBtns)}`);
    log(`  [디버그 답글클릭] ${replyBtnClicked}`);
  }

  if (!replyBtnClicked) {
    // fallback: 행 클릭 후 재시도
    await page.evaluate((top) => {
      const rows = Array.from(document.querySelectorAll('.ag-center-cols-container .ag-row'));
      const r = rows.find(row => row.style.top === top);
      if (r) r.click();
    }, row.top);
    await sleep(1200);
    const retry = await clickReplyButton(page);
    if (!retry) {
      return { success: false, reason: '답글작성 버튼 없음' };
    }
  }

  // 모달 열릴 때까지 대기 (최대 5초)
  await sleep(2500);

  // 모달 상태 디버그 (항상 캡처)
  const modalDebugInfo = await page.evaluate(() => {
    const allTA = Array.from(document.querySelectorAll('textarea')).map(t => {
      const r = t.getBoundingClientRect();
      return {
        readonly: t.readOnly, disabled: t.disabled,
        classes: t.className.substring(0, 60),
        rect: `${Math.round(r.top)},${Math.round(r.left)},${Math.round(r.width)},${Math.round(r.height)}`,
        value: t.value.substring(0, 30),
      };
    });
    const modals = Array.from(document.querySelectorAll(
      '.seller-layer-modal, .modal.show, .modal.fade, [role="dialog"], mat-dialog-container'
    )).map(m => ({
      tag: m.tagName, cls: m.className.substring(0, 60),
      visible: m.offsetParent !== null || window.getComputedStyle(m).display !== 'none',
    }));
    const editables = Array.from(document.querySelectorAll('[contenteditable="true"]')).map(e => ({
      tag: e.tagName, cls: e.className.substring(0, 40),
      rect: (() => { const r = e.getBoundingClientRect(); return `${Math.round(r.width)},${Math.round(r.height)}`; })(),
    }));
    return { textareas: allTA, modals, editables };
  });
  log(`  [모달상태] TA:${modalDebugInfo.textareas.length}개 모달:${modalDebugInfo.modals.length}개 editable:${modalDebugInfo.editables.length}개`);
  if (modalDebugInfo.textareas.length > 0) {
    log(`    textarea[0]: ${JSON.stringify(modalDebugInfo.textareas[0])}`);
  }
  if (modalDebugInfo.modals.length > 0) {
    log(`    modal[0]: ${JSON.stringify(modalDebugInfo.modals[0])}`);
  }
  if (debugMode) {
    await page.screenshot({ path: path.join(CONFIG.outputDir, `debug_modal.png`) }).catch(() => {});
    log(`  [디버그 전체모달] ${JSON.stringify(modalDebugInfo)}`);
  }

  // 3. textarea 입력
  const typed = await typeReplyText(page, replyText);
  if (!typed) {
    if (debugMode) {
      await page.screenshot({ path: path.join(CONFIG.outputDir, 'debug_3_no_textarea.png') });
    }
    return { success: false, reason: '답글 입력란 없음' };
  }
  await sleep(500);

  // 4. 등록 버튼 클릭
  const submitted = await clickSubmit(page);
  if (!submitted) return { success: false, reason: '등록 버튼 없음' };
  await sleep(1500);

  // 5. "이미 답글 작성" 등 오류 팝업 처리
  const popupMsg = await page.evaluate(() => {
    // 확인 버튼이 있는 경고 팝업 감지
    const alertModal = document.querySelector('.modal.show, .layer-popup, [role="alertdialog"]');
    if (alertModal) {
      const txt = alertModal.textContent.trim().substring(0, 100);
      const confirmBtn = Array.from(alertModal.querySelectorAll('button'))
        .find(b => b.textContent.trim() === '확인' || b.textContent.trim() === '닫기');
      if (confirmBtn) confirmBtn.click();
      return txt;
    }
    return null;
  });
  if (popupMsg) {
    if (popupMsg.includes('이미 답글')) {
      log(`\n    [경고] 이미 답변 있음 → 스킵`);
      await deselectAllRows(page);
      return { success: false, reason: '이미 답변 있음' };
    }
    log(`\n    [팝업] ${popupMsg.substring(0, 50)}`);
  }

  await sleep(CONFIG.replyDelay - 1500);

  // 6. 등록 후 선택 해제
  await deselectAllRows(page);

  return { success: true };
}

// ─────────────────────────────────────────────────────────
// 메인 - 리뷰 순회 및 답글 등록
// ─────────────────────────────────────────────────────────
async function main() {
  // ── 중복 실행 방지 잠금 파일 ──────────────────────────
  const lockFile = path.join(__dirname, '.posting.lock');
  if (fs.existsSync(lockFile)) {
    const lockTime = fs.statSync(lockFile).mtimeMs;
    const ageMin = (Date.now() - lockTime) / 60000;
    if (ageMin < 30) {
      log('[오류] 이미 다른 인스턴스가 실행 중입니다. (.posting.lock)');
      log(`       잠금 파일 생성 후 ${ageMin.toFixed(1)}분 경과`);
      log('       강제 실행하려면 .posting.lock 파일을 삭제하세요.');
      process.exit(1);
    }
    fs.unlinkSync(lockFile); // 30분 이상 된 잠금은 무시
  }
  fs.writeFileSync(lockFile, String(process.pid));

  log('');
  log('╔══════════════════════════════════════════════════════════════╗');
  log('║   코에르 리뷰 답변 셀러센터 자동 포스팅 (AG Grid v2)        ║');
  log('╚══════════════════════════════════════════════════════════════╝');
  log('');

  // ── 엑셀 읽기 ────────────────────────────────────────
  const srcFile = findLatestResponseExcel();
  if (!srcFile) { log('[오류] 리뷰답변_자동_*.xlsx 파일 없음'); return; }
  log(`[파일] ${path.basename(srcFile)}`);

  const wb = XLSX.readFile(srcFile);
  const ws = wb.Sheets['리뷰답변'];
  if (!ws) { log('[오류] "리뷰답변" 시트 없음'); return; }

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const header = rows[0];
  const dataRows = rows.slice(1).filter(r => r.some(c => c));

  const COL = {
    product:  header.indexOf('제품명'),
    id:       header.indexOf('아이디'),
    date:     header.indexOf('날짜'),
    response: header.indexOf('답변'),
    posted:   header.indexOf('포스팅완료'),
  };
  log(`[컬럼] id:${COL.id}, date:${COL.date}, response:${COL.response}, posted:${COL.posted}`);

  // 포스팅완료 컬럼 추가
  let postedColIdx = COL.posted;
  if (postedColIdx < 0) {
    postedColIdx = header.length;
    const hCell = XLSX.utils.encode_cell({ r: 0, c: postedColIdx });
    ws[hCell] = { v: '포스팅완료', t: 's' };
    log(`  포스팅완료 컬럼 추가 (col ${postedColIdx})`);
  }

  // 대상 맵 구성: matchKey → { rowIdx, response, product, posted }
  const targetMap = new Map();
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    const writerId = String(row[COL.id] || '').trim();
    const date     = String(row[COL.date] || '').trim();
    const response = String(row[COL.response] || '').trim();
    const posted   = COL.posted >= 0 ? String(row[COL.posted] || '').trim() : '';

    if (!response || posted) continue;

    // Excel 날짜에서 dateOnly 형식으로 정규화
    const dateCleaned = dateOnly(date);
    const matchKey = `${idPrefix(writerId)}__${dateCleaned}`;

    if (!targetMap.has(matchKey)) {
      targetMap.set(matchKey, {
        rowIdx:   i + 1,   // 헤더 포함 실제 행 번호
        response: response,
        product:  String(row[COL.product] || '').substring(0, 20),
        posted:   false,
      });
    }
  }

  log(`[리뷰] 총 ${dataRows.length}개 / 포스팅 대상: ${targetMap.size}개`);
  if (!targetMap.size) { log('  → 포스팅할 항목 없음'); return; }

  // 맵 샘플 출력 (디버그)
  let sampleCount = 0;
  for (const [k, v] of targetMap.entries()) {
    if (sampleCount++ >= 5) break;
    log(`    - "${k}" → ${v.response.substring(0, 20)}...`);
  }

  // ── 브라우저 실행 ─────────────────────────────────────
  log('');
  log('[브라우저] 실행 중...');
  const browser = await puppeteer.launch({
    headless: CONFIG.headless,
    protocolTimeout: 120000,   // 기본 30초 → 120초로 증가 (callFunctionOn timeout 방지)
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--lang=ko-KR,ko', '--start-maximized'],
    defaultViewport: null,
  });

  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36');
  page.on('dialog', async d => { log(`  [알림] ${d.message()}`); await d.accept(); });

  let successCount = 0, failCount = 0, skipCount = 0;

  try {
    // ── 로그인 ───────────────────────────────────────
    log('[인증] 셀러센터 로그인...');
    await loginToSellerCenter(page);
    log('');

    // ── 리뷰 페이지 이동 ──────────────────────────────
    log('[리뷰 페이지] 로드 중...');
    await page.goto(CONFIG.reviewUrl, { waitUntil: 'networkidle2', timeout: 30000 });
    await page.waitForFunction(
      () => !!Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색'),
      { timeout: 15000 }
    );
    await sleep(1500);

    // ── 날짜 범위 1년으로 확장 ────────────────────────
    // 페이지 완전 로드 후 1년 버튼 클릭 (재시도 3회)
    let dateRange = '실패';
    for (let attempt = 0; attempt < 3; attempt++) {
      dateRange = await page.evaluate(() => {
        const btns = Array.from(document.querySelectorAll('button'));
        const btn = btns.find(b => b.textContent.trim() === '1년');
        if (btn) { btn.click(); return '1년 선택'; }
        // 버튼 텍스트 목록 디버그
        const btnTexts = btns.slice(0, 20).map(b => b.textContent.trim()).filter(t => t).join(', ');
        return `없음 (버튼: ${btnTexts.substring(0, 80)})`;
      });
      if (dateRange === '1년 선택') break;
      await sleep(800);
    }
    log(`  날짜 범위: ${dateRange}`);
    await sleep(800);

    // ── 미답변 필터 적용 시도 ────────────────────────
    const filterResult = await trySetUnrepliedFilter(page);
    log(`  답글여부 필터: ${filterResult}`);
    await sleep(500);

    // ── 검색 ────────────────────────────────────────
    await clickSearchButton(page);
    log('  검색 버튼 클릭');

    await page.waitForFunction(
      () => {
        const rows = document.querySelectorAll('.ag-center-cols-container .ag-row');
        return rows.length > 0;
      },
      { timeout: 15000 }
    ).catch(() => {});
    await sleep(2000);

    // 총 리뷰 수
    const totalText = await page.evaluate(() => {
      const m = document.body.innerText.match(/총\s*([\d,]+)개/);
      return m ? parseInt(m[1].replace(/,/g, '')) : '?';
    });
    log(`  총 ${totalText}개 리뷰`);
    log('');

    // 첫 행 디버그 출력
    const firstRows = await collectVisibleRows(page);
    if (firstRows.length > 0) {
      const r = firstRows[0];
      log(`[디버그] 첫 행: writer="${r.writer}", date="${r.date}", reviewNo="${r.reviewNo}", hasReply=${r.hasReply}`);
      log(`  allTexts: ${JSON.stringify(r.allTexts.slice(0, 8))}`);
      const sampleKey = `${idPrefix(r.writer)}__${dateOnly(r.date)}`;
      log(`  matchKey: "${sampleKey}" → ${targetMap.has(sampleKey) ? '매칭됨' : '매칭없음'}`);
    }
    log('');

    // ── AG Grid 순회 루프 ─────────────────────────────
    log('[포스팅] 리뷰 순회 시작...');
    log('');

    const processedKeys = new Set();
    let scrollAttempt = 0;
    let lastRowCount = 0;
    let staleScrollCount = 0;
    let pageNum = 1;

    while (scrollAttempt < CONFIG.maxScrollAttempts) {
      const visibleRows = await collectVisibleRows(page);

      let foundNewInThisScroll = false;
      for (const row of visibleRows) {
        if (!row.date) continue;
        const matchKey = `${idPrefix(row.writer)}__${dateOnly(row.date)}`;

        if (row.hasReply) {
          if (targetMap.has(matchKey) && !processedKeys.has(matchKey)) {
            log(`    [이미답변] ${matchKey}`);
          }
          continue;
        }
        if (processedKeys.has(matchKey)) continue;

        const target = targetMap.get(matchKey);
        if (!target) {
          // 미매칭 행 로그 (처음 100개)
          if (processedKeys.size < 5 && (successCount + failCount) === 0) {
            // Silent skip
          }
          continue;
        }
        if (target.posted) continue;

        // 처리
        foundNewInThisScroll = true;
        processedKeys.add(matchKey);

        process.stdout.write(`  [${row.writer} / ${dateOnly(row.date)}] ${target.product}... `);

        // 첫 3개는 디버그 모드 ON
        const isDebug = (successCount + failCount) < 3;

        try {
          const result = await processRow(page, row, target.response, isDebug);
          if (result.success) {
            const cellAddr = XLSX.utils.encode_cell({ r: target.rowIdx, c: postedColIdx });
            ws[cellAddr] = { v: '✓', t: 's' };
            target.posted = true;
            successCount++;
            log(`✓`);
          } else {
            failCount++;
            log(`✗ ${result.reason}`);
            // 실패 시 모달 닫기
            await closeModal(page);
          }
        } catch (err) {
          failCount++;
          log(`✗ 오류: ${err.message}`);
          // 예외 발생 시 모달 닫기
          await closeModal(page);
        }

        // 답글 등록 후 페이지가 업데이트됐을 수 있으므로 대기
        await sleep(1000);
        break; // 한 행 처리 후 visibleRows 재수집
      }

      // 새로 처리한 행이 없으면 스크롤
      if (!foundNewInThisScroll) {
        // 첫 스크롤 전 컨테이너 디버그 (최초 1회)
        if (scrollAttempt === 0) {
          const dims = await page.evaluate(() => {
            const info = {};
            ['ag-body-viewport','ag-center-cols-viewport','ag-body-horizontal-scroll-viewport'].forEach(cls => {
              const el = document.querySelector(`.${cls}`);
              if (el) info[cls] = { scrollTop: el.scrollTop, scrollH: el.scrollHeight, clientH: el.clientHeight };
            });
            info.window = { scrollY: window.scrollY, innerH: window.innerHeight, docH: document.body.scrollHeight };
            return info;
          });
          log(`  [스크롤 컨테이너] ${JSON.stringify(dims)}`);
        }

        const { moved, sel } = await scrollDown(page);
        await sleep(CONFIG.scrollWait);

        if (moved) {
          log(`    스크롤 이동: ${sel}`);
        }
        if (!moved) {
          staleScrollCount++;
          log(`    스크롤 정지 (${staleScrollCount}/5)`);
          if (staleScrollCount >= 5) {
            // 다음 페이지 시도
            log(`  [페이지 ${pageNum}] 끝 → 다음 페이지 확인...`);
            const hasNext = await clickNextPage(page, pageNum);
            if (!hasNext) {
              log('  다음 페이지 없음 → 완료');
              break;
            }
            pageNum++;
            log(`  → 페이지 ${pageNum} 로딩 중...`);
            await sleep(3000);
            // 페이지 이동 후 AG Grid 재로드 대기
            await page.waitForFunction(
              () => document.querySelectorAll('.ag-center-cols-container .ag-row').length > 0,
              { timeout: 10000 }
            ).catch(() => {});
            await sleep(1500);
            // 스크롤 맨 위로
            await page.evaluate(() => {
              const vp = document.querySelector('.ag-body-viewport');
              if (vp) vp.scrollTop = 0;
            });
            await sleep(500);
            staleScrollCount = 0;
            scrollAttempt = 0; // 새 페이지에서 스크롤 카운터 리셋
          }
        } else {
          staleScrollCount = 0;
        }
        scrollAttempt++;
      }

      // 진행 상황 (10회마다)
      if (scrollAttempt % 10 === 0 && scrollAttempt > 0) {
        const remaining = targetMap.size - successCount - failCount;
        log(`  [진행] 스크롤 ${scrollAttempt}회 / 성공 ${successCount} / 실패 ${failCount} / 남은 ${remaining}`);
      }

      // 모두 완료했으면 종료
      const allDone = [...targetMap.values()].every(t => t.posted);
      if (allDone) {
        log('  → 모든 대상 처리 완료');
        break;
      }
    }

    // ── 매칭 못 찾은 항목 스킵 리포트 ───────────────
    for (const [k, t] of targetMap.entries()) {
      if (!t.posted) {
        skipCount++;
        log(`  SKIP: "${k}" - ${t.product}`);
      }
    }

    // ── 엑셀 저장 ─────────────────────────────────────
    log('');
    XLSX.writeFile(wb, srcFile);
    log(`[저장] ${path.basename(srcFile)}`);

    log('');
    log('─'.repeat(60));
    log(`  성공: ${successCount}개  실패: ${failCount}개  스킵: ${skipCount}개`);
    log('─'.repeat(60));

  } finally {
    await browser.close();
    try { fs.unlinkSync(lockFile); } catch (e) {}
  }
}

main().catch(err => { console.error('\n[오류]', err.message); process.exit(1); });
