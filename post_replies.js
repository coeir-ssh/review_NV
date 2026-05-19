/**
 * post_replies.js — 에이전트가 생성한 답변(replies.json)을 셀러센터에 등록
 *
 * AI 호출 없음. 입력 JSON 의 각 항목을 순회하며:
 *   - judgeLabel === '환불검토' → 답글 등록 안 함 (보고서에만 포함)
 *   - judgeLabel === '답변' + replyText 있음 → 답글 등록
 *
 * 등록 후 검증까지 수행하고 posted_results.json 으로 저장.
 *
 * 입력 형식 (replies.json):
 * {
 *   "date": "2026. 5. 19.",
 *   "results": [
 *     {
 *       "no": 1,
 *       "writer": "abc***",
 *       "reviewNo": "4965448010",
 *       "productName": "...",
 *       "optionName": "...",
 *       "rating": 5,
 *       "date": "2026.05.19. 08:00",
 *       "reviewText": "...",
 *       "reviewPosition": 5,
 *       "productUrl": "...",
 *       "judgeLabel": "답변" | "환불검토",
 *       "judgeReason": "...",
 *       "judgeConfidence": 95,
 *       "replyText": "...",      // 답변일 때만
 *       "refundPolicy": "..."    // 환불검토일 때만
 *     }
 *   ]
 * }
 *
 * 사용: node post_replies.js [--in replies.json] [--out posted_results.json]
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
  postReply,
} = require('./post_product_reviews');

const log   = msg => console.log(msg);
const sleep = ms  => new Promise(r => setTimeout(r, ms));

// CLI 옵션
const argv  = process.argv.slice(2);
const inArg = argv.find(a => a.startsWith('--in='));
const outArg = argv.find(a => a.startsWith('--out='));
const inFile  = inArg  ? inArg.split('=')[1]  : 'replies.json';
const outFile = outArg ? outArg.split('=')[1] : 'posted_results.json';

(async () => {
  log('');
  log('╔══════════════════════════════════════════════════════════════╗');
  log('║   post_replies.js — replies.json 의 답변을 셀러센터 등록    ║');
  log('╚══════════════════════════════════════════════════════════════╝');
  log('');

  const inPath = path.isAbsolute(inFile) ? inFile : path.join(__dirname, inFile);
  if (!fs.existsSync(inPath)) {
    console.error(`[오류] 입력 파일 없음: ${inPath}`);
    process.exit(1);
  }
  const input = JSON.parse(fs.readFileSync(inPath, 'utf8'));
  const replies = input.results || input.reviews || [];
  log(`[입력] ${inPath}`);
  log(`[입력] 총 ${replies.length} 건 (답변=${replies.filter(r => r.judgeLabel === '답변').length}, 환불검토=${replies.filter(r => r.judgeLabel === '환불검토').length})`);

  const browser = await puppeteer.launch({
    headless:        CONFIG.headless,
    protocolTimeout: 120000,
    args:            ['--no-sandbox', '--disable-setuid-sandbox', '--lang=ko-KR,ko', '--window-size=1600,900'],
    defaultViewport: { width: 1600, height: 900 },
  });
  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36');
  page.on('dialog', async d => { log(`  [알림] ${d.message()}`); await d.accept(); });

  let successCount = 0;
  let failCount    = 0;
  let refundCount  = 0;
  const results    = [];

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

    // 동일한 1주일 + 답글미등록 필터 적용
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
    await page.waitForFunction(() => {
      const hasSelectize = Array.from(document.querySelectorAll('.selectize-dropdown-content .option'))
        .some(el => el.textContent.trim().includes('답글미등록'));
      const hasLabel = document.body.innerText.includes('답글여부');
      return hasSelectize || hasLabel;
    }, { timeout: 45000 }).catch(() => {});
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
    await page.evaluate(() => {
      const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색');
      if (btn) btn.click();
    });
    await sleep(1500);

    // ── 환불검토 항목은 등록 없이 결과만 누적 ──
    for (const r of replies) {
      if (r.judgeLabel === '환불검토') {
        refundCount++;
        results.push({ ...r, refundCheck: '검토필요', posted: false });
      }
    }

    // ── 답변 항목만 행 찾아 등록 ──
    const answerList = replies.filter(r => r.judgeLabel === '답변' && r.replyText);
    log('');
    log(`[등록] 답변 ${answerList.length} 건 등록 시작...`);
    const replyByKey = new Map();
    for (const r of answerList) {
      const key = `${r.writer}_${r.reviewNo || r.date}`;
      replyByKey.set(key, r);
    }

    const processedKeys = new Set();
    let   staleCount    = 0;
    while (processedKeys.size < answerList.length) {
      const visible = await collectVisibleRows(page);
      // 답변 대상이며 아직 처리 안 한 행 찾기
      const nextRow = visible.find(v => {
        const k = `${v.writer}_${v.reviewNo || v.date}`;
        return replyByKey.has(k) && !processedKeys.has(k);
      });
      if (!nextRow) {
        const moved = await scrollDown(page);
        await sleep(700);
        if (!moved) {
          staleCount++;
          if (staleCount >= 5) { log('  스크롤 종료 — 더 이상 매칭 행 없음'); break; }
        } else { staleCount = 0; }
        continue;
      }
      staleCount = 0;
      const key = `${nextRow.writer}_${nextRow.reviewNo || nextRow.date}`;
      const target = replyByKey.get(key);

      log(`──────────────────────────────────────`);
      log(`[등록] ${nextRow.writer} | ${nextRow.productName?.substring(0, 30)}`);
      log(`  답변: ${target.replyText.substring(0, 60)}...`);
      const postResult = await postReply(page, nextRow, target.replyText);

      if (postResult.success) {
        successCount++;
        processedKeys.add(key);
        results.push({ ...target, refundCheck: '-', posted: true });
        log(`  ✓ 등록 완료 (${successCount}/${answerList.length})`);
        // 그리드 새로고침
        await page.evaluate(() => {
          const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색');
          if (btn) btn.click();
        });
        await page.waitForFunction(
          () => document.querySelectorAll('.ag-center-cols-container .ag-row').length >= 0,
          { timeout: 10000 }
        ).catch(() => {});
        await page.evaluate(() => {
          const vp = document.querySelector('.ag-body-viewport');
          if (vp) vp.scrollTop = 0;
        });
        await sleep(1000);
      } else {
        failCount++;
        processedKeys.add(key);
        results.push({ ...target, refundCheck: '-', posted: false, failReason: postResult.reason });
        log(`  ✗ 실패: ${postResult.reason}`);
      }
    }

    // ── 등록 대상이었지만 페이지에서 못 찾은 행도 결과에 포함 ──
    for (const r of answerList) {
      const key = `${r.writer}_${r.reviewNo || r.date}`;
      if (!processedKeys.has(key)) {
        failCount++;
        results.push({ ...r, refundCheck: '-', posted: false, failReason: '페이지에서 행을 찾지 못함' });
      }
    }

  } finally {
    await browser.close();
  }

  log('');
  log('══════════════════════════════════════════');
  log(`등록 성공: ${successCount} / 실패: ${failCount} / 환불검토: ${refundCount}`);
  log('══════════════════════════════════════════');

  const out = {
    ...input,
    postedAt: new Date().toISOString(),
    replied:  successCount,
    failed:   failCount,
    refund:   refundCount,
    results,
  };
  const outPath = path.isAbsolute(outFile) ? outFile : path.join(__dirname, outFile);
  fs.writeFileSync(outPath, JSON.stringify(out, null, 2), 'utf8');
  log(`✅ 결과 저장: ${outPath}`);
})().catch(err => {
  console.error('\n[오류]', err.message);
  console.error(err.stack);
  process.exit(1);
});
