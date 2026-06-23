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

    // ── 결과 행 수집 (세로+가로 스크롤하며 모든 컬럼 누적) ──
    // AG Grid 는 가로도 가상화 → 전시상태(contentsStatusType)·답글여부(hasComment) 컬럼이
    // 화면 오른쪽 밖이면 DOM 에 없음. 가로로도 스크롤하며 row-index 별 셀을 누적해야 함.
    log('[수집] 검색 결과 전시상태/답글여부 수집 (세로+가로 스크롤)...');
    // rowAccum: rowIndex → { colId: text, ... } (가로 스크롤 단계마다 누적 병합)
    const rowAccum = {};   // rowIndex -> cellByColId
    const headerAccum = {}; // colId -> label (가로 스크롤로 추가 발견되는 헤더 누적)

    // 한 번의 DOM 스냅샷에서 보이는 셀/헤더를 누적
    const harvest = async () => {
      const snap = await page.evaluate(() => {
        const heads = {};
        document.querySelectorAll('.ag-header-cell').forEach(h => {
          const k = h.getAttribute('col-id') || '';
          if (k) heads[k] = (h.innerText || '').trim().split('\n')[0];
        });
        const rows = [];
        // center + pinned-left + pinned-right 모두
        document.querySelectorAll('.ag-row[row-index]').forEach(row => {
          const ri = row.getAttribute('row-index');
          const cells = {};
          row.querySelectorAll('.ag-cell').forEach(c => {
            const cid = c.getAttribute('col-id') || '';
            if (cid) cells[cid] = (c.innerText || '').trim();
          });
          rows.push({ ri, cells });
        });
        return { heads, rows };
      });
      Object.assign(headerAccum, snap.heads);
      for (const { ri, cells } of snap.rows) {
        if (!rowAccum[ri]) rowAccum[ri] = {};
        Object.assign(rowAccum[ri], cells);
      }
    };

    // 가로 스크롤: 한 세로 위치에서 viewport 를 좌→우 끝까지 훑으며 harvest
    // ⚠️ AG Grid 가로 스크롤 컨테이너는 .ag-center-cols-viewport (.ag-body-viewport 아님 — 그건 세로용)
    const HSCROLL = '.ag-center-cols-viewport';
    const sweepHorizontal = async () => {
      await page.evaluate((sel) => { const vp = document.querySelector(sel); if (vp) vp.scrollLeft = 0; }, HSCROLL);
      await sleep(250);
      await harvest();
      for (let i = 0; i < 20; i++) {
        const moved = await page.evaluate((sel) => {
          const vp = document.querySelector(sel);
          if (!vp) return false;
          const before = vp.scrollLeft;
          vp.scrollLeft = before + 500;
          return vp.scrollLeft !== before;
        }, HSCROLL);
        await sleep(250);
        await harvest();
        if (!moved) break;
      }
      await page.evaluate((sel) => { const vp = document.querySelector(sel); if (vp) vp.scrollLeft = 0; }, HSCROLL);
    };

    let stale = 0;
    let prevRowCount = 0;
    while (true) {
      await sweepHorizontal();
      const rowCount = Object.keys(rowAccum).length;
      if (rowCount > prevRowCount) { log(`  [수집] 행 ${rowCount} (누적)`); prevRowCount = rowCount; stale = 0; }
      else stale++;
      const moved = await page.evaluate(() => {
        const vp = document.querySelector('.ag-body-viewport');
        if (!vp) return false;
        const before = vp.scrollTop;
        vp.scrollTop = before + 400;
        return vp.scrollTop !== before;
      });
      await sleep(400);
      if (!moved && stale >= 3) break;
    }
    log(`[진단] 누적 헤더 col-id: ${JSON.stringify(headerAccum)}`);

    // rowAccum → statusByNo (id 기준)
    for (const ri of Object.keys(rowAccum)) {
      const cells = rowAccum[ri];
      const texts = Object.values(cells);
      const id = cells.id || texts.find(t => /^\d{8,12}$/.test(t)) || '';
      if (!id) continue;
      statusByNo[id] = {
        id,
        status:     cells.contentsStatusType || cells.displayStatus || cells.statusType || '',
        hasComment: cells.hasComment || cells.commentYn || cells.hasReply || '',
        modifyDate: cells.modifyDate || '',
        allTexts:   texts,
      };
    }
    log(`[수집] 완료 — 총 ${Object.keys(statusByNo).length} 행`);

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
      const allTxt = found.allTexts.join(' ');
      const blind  = isBlind(found.status) || isBlind(allTxt);
      const normal = isNormal(found.status) || isNormal(allTxt);
      actualStatus = found.status || (blind ? '블라인드' : normal ? '정상' : '?');
      hasReply = found.hasComment || '';
      // 답글 등록 여부: 'Y' 포함이면 답변 달림, 그 외(N/빈값)는 미등록
      const replied = /Y/i.test(hasReply);
      if (actualStatus === '?' || (!blind && !normal)) {
        // 전시상태를 읽지 못함 → 빗나감으로 단정하지 말고 확인불가 처리 (학습 오염 방지)
        verdict = '확인불가(전시상태 못읽음)';
      } else if (q.judgeLabel === '환불검토') {
        // 블라인드 → 적중(환불됨) / 정상+답글Y → 빗나감(담당자가 답변처리) / 정상+답글N → 대기중(담당자 미처리)
        verdict = blind ? '적중' : (replied ? '빗나감' : '대기중');
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
    // 큐 마킹 — '대기중'·'확인불가' 는 아직 확정 아니므로 verified=false 유지(다음날 재검증)
    const settled = (verdict === '적중' || verdict === '빗나감' || verdict === '빗나감(환불됨)' || verdict === '정상(답변처리)');
    if (settled) {
      q.verified = true;
      q.verifiedAt = new Date().toISOString();
      q.verdict = verdict;
      q.actualStatus = actualStatus;
    } else {
      // 대기중/확인불가 → 미확정 유지, 흔적만 기록
      q.verified = false;
      q.lastVerdict = verdict;
      q.lastCheckedAt = new Date().toISOString();
      q.actualStatus = actualStatus;
    }
  }

  // 요약 카운트
  const refundHit     = results.filter(r => r.judgeLabel === '환불검토' && r.verdict === '적중').length;
  const refundMiss    = results.filter(r => r.judgeLabel === '환불검토' && r.verdict === '빗나감').length;
  const refundPending = results.filter(r => r.judgeLabel === '환불검토' && r.verdict === '대기중').length;
  const ansOk      = results.filter(r => r.judgeLabel === '답변' && r.verdict === '정상(답변처리)').length;
  const ansMiss    = results.filter(r => r.judgeLabel === '답변' && r.verdict.startsWith('빗나감')).length;
  const unknown    = results.filter(r => r.verdict.startsWith('확인불가')).length;

  const out = {
    verifiedAt: new Date().toISOString(),
    pending: pending.length,
    counts: { refundHit, refundMiss, refundPending, ansOk, ansMiss, unknown },
    results,
  };
  fs.writeFileSync(RESULT_PATH, JSON.stringify(out, null, 2), 'utf8');
  fs.writeFileSync(QUEUE_PATH, JSON.stringify(queue, null, 2), 'utf8');

  log('');
  log('══════════════════════════════════════════');
  log(`환불검토 적중 ${refundHit} / 빗나감 ${refundMiss} / 대기중 ${refundPending}  |  답변 정상 ${ansOk} / 빗나감 ${ansMiss}  |  확인불가 ${unknown}`);
  log(`✅ 결과 저장: ${RESULT_PATH}`);
  log('══════════════════════════════════════════');
})().catch(err => {
  console.error('\n[오류]', err.message);
  console.error(err.stack);
  process.exit(1);
});
