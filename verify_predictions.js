/**
 * verify_predictions.js — 전날 판단(환불검토 / 판단근거 답변)이 실제로 맞았는지 검증
 *
 * AI 호출 없음. 셀러센터 리뷰관리에서 "리뷰글번호" 복수검색으로
 * 각 리뷰의 전시상태(블라인드/정상)·답글여부를 확인해 판정한다.
 *
 * 입력: verification_queue.json  (verified:false 항목 = 미검증분)
 * 출력: verification_result.json  + 큐의 해당 항목 verified:true 마킹
 *
 * 판정 규칙:
 *   - 환불검토 건 + 전시상태 '블라인드'  → 적중 (제안대로 환불됨)
 *   - 환불검토 건 + 전시상태 '정상'      → 빗나감 (답변 처리됨, 너무 보수적)
 *   - 답변+판단근거 건 + 전시상태 '정상' + 답글 Y → 정상 처리(답변 달림)
 *   - 답변+판단근거 건 + 전시상태 '블라인드' → 예외(환불됨, 내가 답변으로 본 게 빗나감)
 *
 * 사용: node verify_predictions.js [--force]
 */

const path = require('path');
const fs   = require('fs');
const puppeteer = require('puppeteer');

const {
  CONFIG,
  loginToSellerCenter,
  ensureCoeirStore,
  shouldSkipToday,
} = require('./post_product_reviews');

const log   = msg => console.log(msg);
const sleep = ms  => new Promise(r => setTimeout(r, ms));

const QUEUE_PATH  = path.join(__dirname, 'verification_queue.json');
const RESULT_PATH = path.join(__dirname, 'verification_result.json');

(async () => {
  log('');
  log('╔══════════════════════════════════════════════════════════════╗');
  log('║   verify_predictions.js — 전날 판단 검증 (AI 없음)         ║');
  log('╚══════════════════════════════════════════════════════════════╝');
  log('');

  // 영업일 가드 (--force 로 무시 가능)
  const force = process.argv.includes('--force');
  if (!force) {
    const { skip, reason } = await shouldSkipToday();
    if (skip) {
      log(`[영업일 가드] ${reason} → 검증 건너뜀.`);
      process.exit(0);
    }
  }

  // 큐 로드
  if (!fs.existsSync(QUEUE_PATH)) {
    log('[검증] verification_queue.json 없음 → 검증 대상 없음. 정상 종료.');
    fs.writeFileSync(RESULT_PATH, JSON.stringify({ verifiedAt: new Date().toISOString(), pending: 0, results: [] }, null, 2));
    process.exit(0);
  }
  let queue = JSON.parse(fs.readFileSync(QUEUE_PATH, 'utf8'));
  if (!Array.isArray(queue)) queue = [];
  const pending = queue.filter(q => !q.verified && q.reviewNo);
  log(`[검증] 큐 총 ${queue.length} / 미검증 ${pending.length} 건`);

  if (pending.length === 0) {
    log('[검증] 미검증 대상 없음. 정상 종료.');
    fs.writeFileSync(RESULT_PATH, JSON.stringify({ verifiedAt: new Date().toISOString(), pending: 0, results: [] }, null, 2));
    process.exit(0);
  }

  const reviewNos = pending.map(q => String(q.reviewNo));
  log(`[검증] 조회할 리뷰글번호: ${reviewNos.join(', ')}`);

  const headed = process.argv.includes('--headed');
  const browser = await puppeteer.launch({
    headless:        headed ? false : CONFIG.headless,
    userDataDir:     CONFIG.userDataDir,
    protocolTimeout: 120000,
    args:            ['--no-sandbox', '--disable-setuid-sandbox', '--lang=ko-KR,ko', '--window-size=1600,900'],
    defaultViewport: { width: 1600, height: 900 },
  });
  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36');
  page.on('dialog', async d => { log(`  [알림] ${d.message()}`); await d.accept(); });

  // reviewNo → 셀러센터 상태
  const statusByNo = {};

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

    // ── 날짜 범위 넉넉히 (1개월) — 검증 대상이 며칠 전일 수 있음 ──
    log('[필터] 날짜 1개월...');
    for (let t = 0; t < 5; t++) {
      const r = await page.evaluate(() => {
        const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '1개월');
        if (btn) { btn.click(); return true; }
        return false;
      });
      if (r) break;
      await sleep(600);
    }
    await sleep(800);

    // ── 상세검색: "리뷰글번호" 선택 ──
    log('[상세검색] "리뷰글번호" 선택...');
    // searchKeywordType 드롭다운에서 '리뷰글번호' 선택 (selectize 또는 native select 모두 시도)
    const typeSelected = await page.evaluate(() => {
      // 1) selectize 드롭다운에 '리뷰글번호' 옵션
      const opt = Array.from(document.querySelectorAll('.selectize-dropdown-content .option, option'))
        .find(el => el.textContent.trim() === '리뷰글번호');
      if (opt && opt.tagName === 'OPTION') {
        const sel = opt.closest('select');
        if (sel) { sel.value = opt.value; sel.dispatchEvent(new Event('change', { bubbles: true })); return 'native'; }
      }
      return null;
    });
    // selectize 방식: 컨트롤 클릭 → 옵션 클릭
    if (typeSelected !== 'native') {
      const handle = await page.evaluateHandle(() => {
        const controls = Array.from(document.querySelectorAll('.selectize-control'));
        return controls.find(c => {
          const d = c.querySelector('.selectize-dropdown-content');
          return d && d.textContent.includes('리뷰글번호');
        }) || null;
      });
      const el = handle.asElement();
      if (el) {
        const input = await el.$('.selectize-input');
        if (input) {
          await input.click();
          await sleep(400);
          await page.evaluate(() => {
            const o = Array.from(document.querySelectorAll('.selectize-dropdown-content .option'))
              .find(x => x.textContent.trim() === '리뷰글번호');
            if (o) o.click();
          });
          await sleep(500);
        }
      }
    }
    log(`  → 선택 방식: ${typeSelected || 'selectize'}`);

    // ── 리뷰글번호 textarea 에 복수 입력 ──
    log('[상세검색] 리뷰글번호 복수 입력...');
    const filled = await page.evaluate((nos) => {
      // ng-model searchKeyword 또는 "복수 검색" placeholder textarea
      const ta = Array.from(document.querySelectorAll('textarea'))
        .find(t => /searchKeyword/.test(t.getAttribute('ng-model') || '')
                || /복수 검색/.test(t.getAttribute('placeholder') || ''));
      if (!ta) return false;
      const setter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
      setter.call(ta, nos.join(','));
      ta.dispatchEvent(new Event('input', { bubbles: true }));
      ta.dispatchEvent(new Event('change', { bubbles: true }));
      return true;
    }, reviewNos);
    if (!filled) {
      log('  [경고] 리뷰글번호 textarea 를 찾지 못함 — DOM 구조 확인 필요');
    }
    await sleep(500);

    // ── 검색 ──
    await page.evaluate(() => {
      const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색');
      if (btn) btn.click();
    });
    await page.waitForFunction(
      () => document.querySelectorAll('.ag-center-cols-container .ag-row').length >= 0,
      { timeout: 15000 }
    ).catch(() => {});
    await sleep(2000);

    // ── 첫 행 col-id 덤프 (진단) ──
    const headerDump = await page.evaluate(() => {
      const h = {};
      document.querySelectorAll('.ag-header-cell').forEach(c => {
        const k = c.getAttribute('col-id') || '?';
        h[k] = (c.innerText || '').trim().split('\n')[0];
      });
      return h;
    });
    log(`[진단] 헤더 col-id: ${JSON.stringify(headerDump)}`);

    // ── 결과 행 수집 (스크롤하며) ──
    log('[수집] 검색 결과 전시상태/답글여부 수집...');
    const seen = new Set();
    let stale = 0;
    while (true) {
      const rows = await page.evaluate(() => {
        const out = [];
        document.querySelectorAll('.ag-center-cols-container .ag-row').forEach(row => {
          const cells = {};
          row.querySelectorAll('.ag-cell').forEach(c => {
            const cid = c.getAttribute('col-id') || '';
            if (cid) cells[cid] = (c.innerText || '').trim();
          });
          // id(리뷰글번호) / contentsStatusType(전시상태) / hasComment(답글여부) / modifyDate
          const texts = Array.from(row.querySelectorAll('.ag-cell')).map(c => (c.innerText||'').trim());
          out.push({
            id:        cells.id || texts.find(t => /^\d{8,12}$/.test(t)) || '',
            status:    cells.contentsStatusType || '',
            hasComment:cells.hasComment || '',
            modifyDate:cells.modifyDate || '',
            allTexts:  texts,
          });
        });
        return out;
      });
      let added = 0;
      for (const r of rows) {
        if (!r.id || seen.has(r.id)) continue;
        seen.add(r.id);
        statusByNo[r.id] = r;
        added++;
      }
      if (added > 0) { log(`  [수집] +${added} (총 ${Object.keys(statusByNo).length})`); stale = 0; }
      else stale++;
      const moved = await page.evaluate(() => {
        const vp = document.querySelector('.ag-body-viewport');
        if (!vp) return false;
        const before = vp.scrollTop;
        vp.scrollTop = before + 400;
        return vp.scrollTop !== before;
      });
      await sleep(500);
      if (!moved && stale >= 3) break;
    }

  } finally {
    await browser.close();
  }

  // ── 판정 ──
  const norm = s => (s || '').replace(/\s/g, '');
  const isBlind  = s => /블라인드/.test(norm(s));
  const isNormal = s => /정상/.test(norm(s));

  const results = [];
  for (const q of pending) {
    const found = statusByNo[String(q.reviewNo)];
    let verdict, actualStatus, hasReply;
    if (!found) {
      verdict = '확인불가';        // 셀러센터 검색결과에 없음 (기간 밖/삭제 등)
      actualStatus = '미발견';
      hasReply = '';
    } else {
      actualStatus = found.status || (isBlind(found.allTexts.join(' ')) ? '블라인드' : isNormal(found.allTexts.join(' ')) ? '정상' : '?');
      hasReply = found.hasComment || '';
      const blind = isBlind(actualStatus) || isBlind(found.allTexts.join(' '));
      if (q.judgeLabel === '환불검토') {
        verdict = blind ? '적중' : '빗나감';     // 환불검토인데 블라인드면 적중
      } else { // 답변 + 판단근거
        verdict = blind ? '빗나감(환불됨)' : '정상(답변처리)';
      }
    }
    results.push({
      reviewNo:        q.reviewNo,
      productName:     q.productName,
      judgeLabel:      q.judgeLabel,
      judgeConfidence: q.judgeConfidence,
      judgeReason:     q.judgeReason,
      proposedReply:   q.proposedReply,
      reviewText:      q.reviewText,
      processedDate:   q.processedDate,
      actualStatus,
      hasReply,
      verdict,
    });
    // 큐 마킹
    q.verified = true;
    q.verifiedAt = new Date().toISOString();
    q.verdict = verdict;
    q.actualStatus = actualStatus;
  }

  // 요약 카운트
  const refundHit  = results.filter(r => r.judgeLabel === '환불검토' && r.verdict === '적중').length;
  const refundMiss = results.filter(r => r.judgeLabel === '환불검토' && r.verdict === '빗나감').length;
  const ansOk      = results.filter(r => r.judgeLabel === '답변' && r.verdict === '정상(답변처리)').length;
  const ansMiss    = results.filter(r => r.judgeLabel === '답변' && r.verdict.startsWith('빗나감')).length;
  const unknown    = results.filter(r => r.verdict === '확인불가').length;

  const out = {
    verifiedAt: new Date().toISOString(),
    pending: pending.length,
    counts: { refundHit, refundMiss, ansOk, ansMiss, unknown },
    results,
  };
  fs.writeFileSync(RESULT_PATH, JSON.stringify(out, null, 2), 'utf8');
  fs.writeFileSync(QUEUE_PATH, JSON.stringify(queue, null, 2), 'utf8');

  log('');
  log('══════════════════════════════════════════');
  log(`환불검토 적중 ${refundHit} / 빗나감 ${refundMiss}  |  답변 정상 ${ansOk} / 빗나감 ${ansMiss}  |  확인불가 ${unknown}`);
  log(`✅ 결과 저장: ${RESULT_PATH}`);
  log('══════════════════════════════════════════');
})().catch(err => {
  console.error('\n[오류]', err.message);
  console.error(err.stack);
  process.exit(1);
});
