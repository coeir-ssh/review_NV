/**
 * collect_pending.js — 답글미등록 리뷰만 수집해 JSON 으로 출력
 *
 * AI 호출 없음. Puppeteer 만 사용해 셀러센터에서 모든 답글미등록 리뷰를
 * 스크래핑하고, 각 리뷰의 brand.naver.com 랭킹 순위까지 미리 조회한다.
 *
 * 출력: pending_reviews.json
 * 다음 단계: Claude Code 에이전트가 이 JSON 을 읽고 각 리뷰별 판단·답변·검수
 *           를 직접 수행한 뒤 replies.json 으로 작성한다.
 *
 * 사용: node collect_pending.js [--out pending_reviews.json] [--scheduled]
 */

const path = require('path');
const fs   = require('fs');
const puppeteer = require('puppeteer');

const {
  CONFIG,
  loginToSellerCenter,
  ensureCoeirStore,
  collectVisibleRows,
  scrollDown,
  findReviewPosition,
} = require('./post_product_reviews');

const log   = msg => console.log(msg);
const sleep = ms  => new Promise(r => setTimeout(r, ms));

// CLI 옵션
const argv = process.argv.slice(2);
const outArg = argv.find(a => a.startsWith('--out='));
const outFile = outArg ? outArg.split('=')[1] : 'pending_reviews.json';

(async () => {
  log('');
  log('╔══════════════════════════════════════════════════════════════╗');
  log('║   collect_pending.js — 답글미등록 리뷰 수집 (AI 없음)       ║');
  log('╚══════════════════════════════════════════════════════════════╝');
  log('');

  const browser = await puppeteer.launch({
    headless:        CONFIG.headless,
    protocolTimeout: 120000,
    args:            ['--no-sandbox', '--disable-setuid-sandbox', '--lang=ko-KR,ko', '--window-size=1600,900'],
    defaultViewport: { width: 1600, height: 900 },
  });
  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36');
  page.on('dialog', async d => { log(`  [알림] ${d.message()}`); await d.accept(); });

  const collected = [];

  try {
    log('[인증] 셀러센터 로그인...');
    await loginToSellerCenter(page);

    log('[스토어] 코에르 스토어 확인...');
    try { await ensureCoeirStore(page); } catch (e) { log(`  [경고] ${e.message}`); }

    log('[리뷰 페이지] 로드 중...');
    await page.goto(CONFIG.reviewUrl, { waitUntil: 'networkidle2', timeout: 30000 });
    await page.waitForFunction(
      () => !!Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색'),
      { timeout: 45000 }
    );
    await sleep(1500);

    // 1주일 + 답글미등록 필터 적용
    log('[필터] 1주일 + 답글미등록...');
    for (let t = 0; t < 5; t++) {
      const r = await page.evaluate(() => {
        const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '1주일');
        if (btn) { btn.click(); return true; }
        return false;
      });
      if (r) break;
      await sleep(600);
    }
    await sleep(800);

    // 답글여부: 답글미등록 선택 (Selectize.js)
    await page.waitForFunction(() => {
      const hasSelectize = Array.from(document.querySelectorAll('.selectize-dropdown-content .option'))
        .some(el => el.textContent.trim().includes('답글미등록'));
      const hasLabel = document.body.innerText.includes('답글여부');
      return hasSelectize || hasLabel;
    }, { timeout: 45000 }).catch(() => log('  [경고] 답글여부 필터 렌더 타임아웃'));
    await sleep(800);

    const selectizeHandle = await page.evaluateHandle(() => {
      const controls = Array.from(document.querySelectorAll('.selectize-control'));
      return controls.find(c => {
        const dropdown = c.querySelector('.selectize-dropdown-content');
        return dropdown && dropdown.textContent.includes('답글미등록');
      }) || null;
    });
    const selectizeEl = selectizeHandle.asElement();
    if (selectizeEl) {
      const inputHandle = await selectizeEl.$('.selectize-input');
      if (inputHandle) {
        await inputHandle.click();
        for (let i = 0; i < 15; i++) {
          const open = await page.evaluate(() => {
            const d = document.querySelector('.selectize-dropdown-content');
            return d && d.offsetHeight > 0;
          });
          if (open) break;
          await sleep(200);
        }
        await page.evaluate(() => {
          const opt = Array.from(document.querySelectorAll('.selectize-dropdown-content .option'))
            .find(el => el.textContent.trim().includes('답글미등록'));
          if (opt) opt.click();
        });
        await sleep(500);
      }
    }

    // 검색 버튼 클릭
    await page.evaluate(() => {
      const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색');
      if (btn) btn.click();
    });
    await page.waitForFunction(
      () => document.querySelectorAll('.ag-center-cols-container .ag-row').length >= 0,
      { timeout: 15000 }
    ).catch(() => {});
    await sleep(1500);

    // ── 스크롤하며 모든 행 수집 (가상스크롤 한 번 훑기) ──
    log('');
    log('[수집] 답글미등록 행 스크롤 수집 중...');
    const seenKeys = new Set();
    let staleCount = 0;
    while (true) {
      const visible = await collectVisibleRows(page);
      let newOnes = 0;
      for (const r of visible) {
        const key = `${r.writer}_${r.reviewNo || r.date}`;
        if (seenKeys.has(key)) continue;
        seenKeys.add(key);
        if (r.checkboxDisabled) continue; // 환불완료 등은 제외
        collected.push(r);
        newOnes++;
      }
      if (newOnes > 0) {
        log(`  [수집] +${newOnes} (총 ${collected.length})`);
        staleCount = 0;
      } else {
        staleCount++;
      }
      const moved = await scrollDown(page);
      await sleep(600);
      if (!moved && staleCount >= 5) {
        log('  스크롤 종료 — 모든 행 수집 완료');
        break;
      }
    }

    // ── 각 리뷰의 brand.naver.com 순위 미리 조회 ──
    log('');
    log(`[순위 조회] ${collected.length} 건의 brand.naver.com 랭킹 순위 조회...`);
    for (let i = 0; i < collected.length; i++) {
      const r = collected[i];
      log(`──────────────────────────────────────`);
      log(`[${i + 1}/${collected.length}] ${r.writer} | ${r.productName?.substring(0, 30)}`);
      log(`  리뷰: "${(r.reviewText || '').substring(0, 40)}"`);
      try {
        const pos = await findReviewPosition(browser, r.productName, r.reviewText, r.writer);
        r.reviewPosition = pos.position > 0 ? pos.position : (pos.position === -2 ? -2 : 0);
        r.productUrl     = pos.productUrl || '';
        r.posPolicy      = pos.policy || '';
        r.searchedCount  = pos.searchedCount || 0;
      } catch (e) {
        log(`  [순위 조회 실패] ${e.message}`);
        r.reviewPosition = 0; r.productUrl = ''; r.posPolicy = '';
      }
    }

  } finally {
    await browser.close();
  }

  // ── JSON 출력 ──
  const now = new Date();
  const dateStr = `${now.getFullYear()}. ${now.getMonth() + 1}. ${now.getDate()}.`;
  const out = {
    date: dateStr,
    collectedAt: now.toISOString(),
    totalReviews: collected.length,
    reviews: collected.map((r, idx) => ({
      no:             idx + 1,
      writer:         r.writer,
      reviewNo:       r.reviewNo,
      productName:    r.productName,
      optionName:     r.optionName || '',
      rating:         r.rating,
      date:           r.date,
      reviewText:     r.reviewText,
      reviewPosition: r.reviewPosition,
      productUrl:     r.productUrl || '',
      posPolicy:      r.posPolicy || '',
      searchedCount:  r.searchedCount || 0,
      rowIndex:       r.rowIndex,           // 페이지 재로딩 시 행 찾기 보조
      channelNo:      r.channelNo,
    })),
  };
  const outPath = path.isAbsolute(outFile) ? outFile : path.join(__dirname, outFile);
  fs.writeFileSync(outPath, JSON.stringify(out, null, 2), 'utf8');
  log('');
  log(`✅ 완료: ${collected.length} 건 수집 → ${outPath}`);
})().catch(err => {
  console.error('\n[오류]', err.message);
  console.error(err.stack);
  process.exit(1);
});
