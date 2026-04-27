/**
 * 특정 리뷰가 상품 리뷰 페이지에서 몇 번째인지 찾는 스크립트
 * 실행: node find_review_position.js
 */
const puppeteer = require('puppeteer');
const sleep     = ms => new Promise(r => setTimeout(r, ms));

const PRODUCT_URL   = 'https://brand.naver.com/coeir/products/5401033074';
const TARGET_TEXT   = '바닥에 고정이 잘 안되네요';
const REVIEW_SEL    = 'span.JnLwAJPsMs';
const MAX_PAGES     = 50;

async function main() {
  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--lang=ko-KR,ko', '--window-size=1400,900'],
    defaultViewport: { width: 1400, height: 900 },
  });

  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36');

  console.log('[이동] 상품 리뷰 페이지...');
  await page.goto(PRODUCT_URL + '#REVIEW', { waitUntil: 'networkidle2', timeout: 30000 });
  await sleep(3000);

  // 현재 정렬 기준 확인
  const sortInfo = await page.evaluate(() => {
    const btns = Array.from(document.querySelectorAll('button, a, span'))
      .filter(el => el.textContent.trim().match(/최신순|추천순|별점순|낮은별점/));
    return btns.map(b => ({ text: b.textContent.trim(), active: b.className }));
  });
  console.log('\n[정렬 기준]', JSON.stringify(sortInfo.slice(0, 4)));

  let globalPos = 0;
  let foundPos  = -1;

  for (let p = 1; p <= MAX_PAGES; p++) {
    await sleep(1500);

    const reviews = await page.evaluate((sel) => {
      return Array.from(document.querySelectorAll(sel)).map(el => el.textContent.trim());
    }, REVIEW_SEL);

    if (reviews.length === 0 && p === 1) {
      console.log('리뷰 셀렉터 미매칭 — 종료');
      break;
    }

    console.log(`\n[${p}페이지] ${reviews.length}개 리뷰`);
    for (let i = 0; i < reviews.length; i++) {
      globalPos++;
      const t = reviews[i].replace(/\n/g, ' ');
      console.log(`  ${globalPos}. ${t.substring(0, 70)}`);
      if (t.includes(TARGET_TEXT)) {
        foundPos = globalPos;
        console.log(`  ✅ 타겟 리뷰 발견! → ${foundPos}번째`);
      }
    }

    if (foundPos > 0) break;

    // 다음 페이지 클릭 ─ 네이버 페이지네이션 구조
    const moved = await page.evaluate(() => {
      // 현재 활성 페이지 번호 파악
      const active = document.querySelector('[class*="pagination"] [class*="active"] button, [class*="pagination"] [class*="on"] button, [class*="pagination"] button[aria-current]');

      // 방법 1: 다음(▶) 버튼
      const paginationBtns = Array.from(document.querySelectorAll('[class*="pagination"] button, [class*="pagination"] a'));
      const nextArrow = paginationBtns.find(b => /다음|next/i.test(b.getAttribute('aria-label') || b.textContent));
      if (nextArrow && !nextArrow.disabled) { nextArrow.click(); return '다음버튼'; }

      // 방법 2: 현재 활성 버튼의 다음 형제
      if (active) {
        const nextLi = active.closest('li, span')?.nextElementSibling;
        const nextBtn = nextLi?.querySelector('button, a');
        if (nextBtn && !nextBtn.disabled) { nextBtn.click(); return '다음형제'; }
      }

      // 방법 3: 숫자 버튼 순서로 다음 번호 클릭
      const numBtns = paginationBtns.filter(b => /^\d+$/.test(b.textContent.trim()));
      const curIdx  = numBtns.findIndex(b => b.classList.toString().includes('active') || b.getAttribute('aria-current'));
      if (curIdx >= 0 && curIdx + 1 < numBtns.length) {
        numBtns[curIdx + 1].click(); return '숫자버튼';
      }
      return false;
    });

    if (!moved) {
      console.log('\n[종료] 다음 페이지 없음');
      break;
    }
    console.log(`  → 다음 페이지 이동 (${moved})`);
  }

  // ── 결과 출력 ──────────────────────────────────────────
  console.log('\n' + '='.repeat(50));
  if (foundPos > 0) {
    const policy = foundPos <= 10 ? '⚠️  1~10위  → 아주 적극 대응 (조건 환불 검토)'
                 : foundPos <= 20 ? '🔶 11~20위 → 적극 대응 (답변 or 환불)'
                 : foundPos <= 40 ? '🔷 21~40위 → 답변 우선 (필요시 환불)'
                 :                  '✅ 41위 이하 → 답변으로 충분';
    console.log(`🎯 결과: "${TARGET_TEXT}" 는 ${foundPos}번째 리뷰`);
    console.log(`   정책: ${policy}`);
  } else {
    console.log(`❌ 타겟 리뷰를 ${MAX_PAGES}페이지 내에서 찾지 못함`);
  }
  console.log('='.repeat(50));

  await browser.close();
}

main().catch(e => { console.error(e); process.exit(1); });
