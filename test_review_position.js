/**
 * 리뷰 순위 조회 테스트 스크립트
 * 답글 등록 없이 현재 미답변 리뷰들의 brand.naver.com 순위만 확인
 * 실행: node test_review_position.js
 */
const puppeteer = require('puppeteer');
const path      = require('path');
const fs        = require('fs');
const cfg       = require('./config');

const sleep = ms => new Promise(r => setTimeout(r, ms));

const CONFIG = {
  sessionFile: path.join(__dirname, '.seller_session.json'),
  reviewUrl:   'https://sell.smartstore.naver.com/#/review/search',
};

// ── product_naver_ids.json 로드 ──────────────────────────
let _ids = null;
function getIds() {
  if (!_ids) {
    const p = path.join(__dirname, 'product_naver_ids.json');
    _ids = fs.existsSync(p) ? JSON.parse(fs.readFileSync(p)) : {};
  }
  return _ids;
}

function findProductNaverId(productName) {
  const ids   = getIds();
  const pName = (productName || '').replace(/\s+/g, ' ').trim();

  // 완전 일치
  for (const [urlId, info] of Object.entries(ids)) {
    if (info.name === pName) return { urlId, ...info };
  }
  // 키워드 점수 매칭
  const keywords = pName.split(/[\s[\]()·]+/).filter(w => w.length >= 2);
  let bestMatch = null, bestScore = 0;
  for (const [urlId, info] of Object.entries(ids)) {
    const iName = (info.name || '').replace(/\s+/g, ' ');
    let score = 0;
    for (const kw of keywords) { if (iName.includes(kw)) score++; }
    if (score > bestScore) { bestScore = score; bestMatch = { urlId, ...info }; }
  }
  return bestScore >= 2 ? bestMatch : null;
}

// ── brand.naver.com 리뷰 순위 조회 ──────────────────────
async function findReviewPosition(browser, productName, reviewText) {
  const productInfo = findProductNaverId(productName);
  if (!productInfo) return { position: -1, policy: '제품 ID 매핑 없음' };
  if (!productInfo.originProductNo) return { position: -1, policy: 'originProductNo 없음' };

  const { urlId, originProductNo, checkoutMerchantNo } = productInfo;
  const productUrl = `https://brand.naver.com/coeir/products/${urlId}`;
  const matchSnippet = reviewText.replace(/\s+/g, ' ').trim().substring(0, 20);

  const reviewPage = await browser.newPage();
  try {
    await reviewPage.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36');
    await reviewPage.goto(productUrl + '#REVIEW', { waitUntil: 'networkidle2', timeout: 30000 });
    await sleep(2000);

    const API_URL = 'https://brand.naver.com/n/v1/contents/reviews/query-pages';
    let globalPos = 0, foundPos = -1;

    for (let p = 1; p <= 100; p++) {
      const body = { checkoutMerchantNo, originProductNo, page: p, pageSize: 20, reviewSearchSortType: 'REVIEW_RANKING' };

      const result = await reviewPage.evaluate(async (url, reqBody) => {
        try {
          const resp = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(reqBody) });
          if (!resp.ok) return { error: `HTTP ${resp.status}` };
          return await resp.json();
        } catch(e) { return { error: e.message }; }
      }, API_URL, body);

      if (result.error) { console.log(`    API 오류: ${result.error}`); break; }

      const reviews = result.contents || result.reviews || [];
      if (!reviews.length) break;

      for (const rv of reviews) {
        globalPos++;
        const text = (rv.reviewContent || rv.body || rv.content || '').replace(/\s+/g, ' ').trim();
        if (text.includes(matchSnippet) || (matchSnippet.length > 5 && matchSnippet.includes(text.substring(0, 10)))) {
          foundPos = globalPos;
          break;
        }
      }
      if (foundPos > 0) break;

      const total = result.totalCount || result.totalElements || 0;
      if (total > 0 && globalPos >= total) break;
      await sleep(300);
    }

    const policy = foundPos <= 0 ? '리뷰 미발견 (삭제됐거나 범위 초과)'
      : foundPos <= 10 ? '⚠️  1~10위 → 아주 적극 대응 (조건 환불 검토)'
      : foundPos <= 20 ? '🔶 11~20위 → 적극 대응 (답변 or 환불)'
      : foundPos <= 40 ? '🔷 21~40위 → 답변 우선 (필요시 환불)'
      :                  '✅ 41위 이하 → 답변으로 충분';

    return { position: foundPos, policy, productUrl, searched: globalPos };
  } catch(e) {
    return { position: -1, policy: `오류: ${e.message}`, productUrl };
  } finally {
    await reviewPage.close().catch(() => {});
  }
}

// ── 셀러센터 미답변 리뷰 수집 ───────────────────────────
async function collectReviews(page) {
  return page.evaluate(() => {
    const pinnedRows = Array.from(document.querySelectorAll('.ag-pinned-left-cols-container .ag-row'));
    const centerRows = Array.from(document.querySelectorAll('.ag-center-cols-container .ag-row'));

    return centerRows.map(row => {
      const rowIndex  = row.getAttribute('row-index') || '';
      const pinnedRow = pinnedRows.find(r => r.getAttribute('row-index') === rowIndex);
      const cb        = pinnedRow ? pinnedRow.querySelector('input[type="checkbox"]') : null;
      const checkboxDisabled = cb
        ? (cb.disabled || cb.getAttribute('aria-disabled') === 'true' || cb.closest('[class*="disabled"]') !== null)
        : false;

      const cells  = Array.from(row.querySelectorAll('.ag-cell'));
      const texts  = cells.map(c => c.textContent.trim()).filter(t => t);

      const channelNo   = texts.find(t => /^\d{10}$/.test(t)) || '';
      const reviewNo    = texts.find(t => /^\d{8,12}$/.test(t) && t !== channelNo) || channelNo;
      const writer      = texts.find(t => t.includes('*') && t.length >= 3) || '';
      const date        = texts.find(t => /\d{4}\.\d{2}\.\d{2}/.test(t)) || '';
      const ratingStr   = texts.find(t => /^[1-5]$/.test(t)) || '';
      const rating      = ratingStr ? parseInt(ratingStr) : 0;
      const productName = texts.find(t => t.includes('코에르') && t.length > 10) || '';
      const SKIP_LABELS = ['일반','프리미엄','포토','한달사용','베스트','답글있음','답변완료','답변있음'];
      const reviewText  = texts.find(t =>
        t.length > 3 &&
        t !== productName &&
        !/^\d+$/.test(t) &&
        !t.includes('*') &&
        !/\d{4}\.\d{2}\.\d{2}/.test(t) &&
        !SKIP_LABELS.includes(t)
      ) || '';

      return { rowIndex, reviewNo, writer, date, rating, productName, reviewText, checkboxDisabled };
    }).filter(r => r.date || r.reviewNo);
  });
}

// ── 메인 ────────────────────────────────────────────────
async function main() {
  console.log('\n=== 리뷰 순위 조회 테스트 ===\n');

  const browser = await puppeteer.launch({
    headless: false,
    protocolTimeout: 120000,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--lang=ko-KR,ko', '--window-size=1600,900'],
    defaultViewport: { width: 1600, height: 900 },
  });

  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36');
  page.on('dialog', async d => { await d.accept(); });

  // ── 세션 로드 + 로그인 ───────────────────────────────
  const saved = fs.existsSync(CONFIG.sessionFile)
    ? JSON.parse(fs.readFileSync(CONFIG.sessionFile)) : null;
  if (saved) {
    await page.setCookie(...saved);
    console.log('[세션] 쿠키 로드 완료');
  }

  // 로그인 확인
  await page.goto('https://sell.smartstore.naver.com/', { waitUntil: 'networkidle2', timeout: 30000 });
  await sleep(2000);
  const loggedIn = await page.evaluate(() => {
    const t = document.body.innerText;
    return t.includes('상품관리') || t.includes('리뷰관리') || t.includes('주문관리');
  });

  if (!loggedIn) {
    console.log('[로그인] 자동 로그인 시도...');
    await page.goto('https://accounts.commerce.naver.com/login', { waitUntil: 'networkidle2', timeout: 20000 });
    await sleep(1500);
    const idF = await page.$('input[placeholder="아이디 또는 이메일 주소"]');
    const pwF = await page.$('input[type="password"]');
    if (idF && pwF) {
      await idF.click({ clickCount: 3 }); await idF.type(cfg.SELLER_ID, { delay: 80 });
      await sleep(300);
      await pwF.click({ clickCount: 3 }); await pwF.type(cfg.SELLER_PW, { delay: 80 });
      await sleep(300);
      await page.evaluate(() => {
        const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '로그인');
        if (btn) btn.click();
      });
      await page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 20000 }).catch(() => {});
      await sleep(3000);
    }
    // 추가 인증 필요 시 브라우저에서 직접 처리
    const ok = await page.evaluate(() => {
      const t = document.body.innerText;
      return t.includes('상품관리') || t.includes('리뷰관리') || t.includes('주문관리');
    });
    if (!ok) {
      console.log('  추가 인증이 필요합니다. 브라우저에서 완료 후 Enter...');
      await new Promise(resolve => {
        const { createInterface } = require('readline');
        const rl = createInterface({ input: process.stdin, output: process.stdout });
        rl.question('  완료 후 Enter ▶ ', () => { rl.close(); resolve(); });
      });
    }
    // 세션 저장
    fs.writeFileSync(CONFIG.sessionFile, JSON.stringify(await page.cookies()));
    console.log('[로그인] 완료, 세션 저장');
  } else {
    console.log('[세션] 로그인 유효');
  }

  // ── 리뷰 페이지 이동 ─────────────────────────────────
  console.log('[이동] 셀러센터 리뷰 페이지...');
  await page.goto(CONFIG.reviewUrl, { waitUntil: 'networkidle2', timeout: 30000 });
  await page.waitForFunction(
    () => !!Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색'),
    { timeout: 30000 }
  );
  await sleep(1500);

  // ── 1주일 필터 ───────────────────────────────────────
  for (let t = 0; t < 5; t++) {
    const r = await page.evaluate(() => {
      const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '1주일');
      if (btn) { btn.click(); return true; }
      return false;
    });
    if (r) { console.log('  날짜: 1주일 선택'); break; }
    await sleep(600);
  }
  await sleep(1000);

  // ── 답글미등록 필터 ──────────────────────────────────
  await page.waitForFunction(() => {
    const hasSelectize = Array.from(document.querySelectorAll('.selectize-dropdown-content .option'))
      .some(el => el.textContent.trim().includes('답글미등록'));
    return hasSelectize || document.body.innerText.includes('답글여부');
  }, { timeout: 15000 }).catch(() => {});
  await sleep(600);

  const selectizeHandle = await page.evaluateHandle(() => {
    const controls = Array.from(document.querySelectorAll('.selectize-control'));
    return controls.find(c => {
      const d = c.querySelector('.selectize-dropdown-content');
      return d && d.textContent.includes('답글미등록');
    }) || null;
  });
  const selectizeEl = selectizeHandle.asElement();
  if (selectizeEl) {
    const inputH = await selectizeEl.$('.selectize-input');
    if (inputH) { await inputH.click(); await sleep(500); }
    const ok = await page.evaluate(() => {
      const opts = Array.from(document.querySelectorAll('.selectize-dropdown-content .option'));
      const t = opts.find(el => el.textContent.trim().includes('답글미등록'));
      if (t) { t.click(); return true; }
      return false;
    });
    console.log(ok ? '  답글여부: 답글미등록 ✓' : '  [경고] 답글미등록 선택 실패');
  }
  await sleep(500);

  // ── 검색 ─────────────────────────────────────────────
  await page.evaluate(() => {
    const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색');
    if (btn) btn.click();
  });
  await page.waitForFunction(
    () => document.querySelectorAll('.ag-center-cols-container .ag-row').length > 0,
    { timeout: 15000 }
  ).catch(() => {});
  await sleep(2000);

  // ── 행 수집 ──────────────────────────────────────────
  const rows = await collectReviews(page);
  const active  = rows.filter(r => !r.checkboxDisabled);
  const skipped = rows.filter(r =>  r.checkboxDisabled);

  console.log(`\n  수집된 행: ${rows.length}개`);
  console.log(`  활성(처리대상): ${active.length}개`);
  console.log(`  비활성(환불완료 등 스킵): ${skipped.length}개`);

  if (skipped.length) {
    console.log('\n[스킵 항목]');
    skipped.forEach(r => console.log(`  - ${r.writer} | ${r.productName?.substring(0, 30)} | "${r.reviewText?.substring(0, 30)}"`));
  }

  // ── 각 활성 리뷰 순위 조회 ───────────────────────────
  console.log('\n' + '='.repeat(60));
  console.log('리뷰 순위 조회 시작...');
  console.log('='.repeat(60));

  for (let i = 0; i < active.length; i++) {
    const r = active[i];
    console.log(`\n[${i + 1}/${active.length}]`);
    console.log(`  작성자  : ${r.writer}`);
    console.log(`  별점    : ${r.rating}점`);
    console.log(`  제품명  : ${r.productName}`);
    console.log(`  리뷰    : "${r.reviewText?.substring(0, 60)}"`);
    console.log(`  조회 중...`);

    const pos = await findReviewPosition(browser, r.productName, r.reviewText);

    console.log(`  ─ 결과 ─────────────────────────────────────`);
    if (pos.position > 0) {
      console.log(`  📍 순위   : ${pos.position}번째 (${pos.searched}개 검색)`);
    } else {
      console.log(`  📍 순위   : 미발견`);
    }
    console.log(`  📋 정책   : ${pos.policy}`);
    if (pos.productUrl) console.log(`  🔗 URL    : ${pos.productUrl}`);
  }

  console.log('\n' + '='.repeat(60));
  console.log('조회 완료. 5초 후 종료...');
  await sleep(5000);

  await browser.close();
}

main().catch(e => { console.error(e); process.exit(1); });
