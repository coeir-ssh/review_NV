/**
 * 코에르 네이버 브랜드스토어 리뷰 수집 프로그램
 * 실행: node collect_reviews.js
 *
 * 수집 대상: 코에르 욕실매트 아기 욕조미끄럼방지패드 화장실앞발판 타일바닥고무깔판 건식논슬립다용도깔개 S
 * 수집 내용: 별점 / 아이디 / 날짜 / 구매옵션 / 리뷰내용
 */

const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const path = require('path');

// ─────────────────────────────────────────────────────────
// 설정
// ─────────────────────────────────────────────────────────
const CONFIG = {
  brandUrl: 'https://brand.naver.com/coeir',
  productKeywords: ['욕실매트', '논슬립', 'S'],
  maxReviewPages: 3,        // 수집할 리뷰 페이지 수 (1페이지 = 20개)
  reviewPageSize: 20,
  outputDir: 'C:\\Users\\AWESOMATIC\\Desktop\\코에르\\클로드\\리뷰수집프로그램',
};

// ─────────────────────────────────────────────────────────
// 유틸
// ─────────────────────────────────────────────────────────
const sleep = ms => new Promise(r => setTimeout(r, ms));

function timestamp() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function log(msg) { console.log(msg); }

function formatDate(isoString) {
  if (!isoString) return '';
  try {
    const d = new Date(isoString);
    const pad = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}.${pad(d.getMonth()+1)}.${pad(d.getDate())}`;
  } catch (_) { return isoString; }
}

// ─────────────────────────────────────────────────────────
// 1단계: 제품 URL 탐색
// ─────────────────────────────────────────────────────────
async function findProductUrl(page) {
  const categoryUrl = CONFIG.brandUrl + '/category/ALL';
  log(`  탐색: ${categoryUrl}`);
  await page.goto(categoryUrl, { waitUntil: 'networkidle2', timeout: 60000 });
  await sleep(2000);

  let url = await findLinkByKeywords(page);
  if (url) return url;

  for (let i = 0; i < 5; i++) {
    await page.evaluate(() => window.scrollBy(0, 600));
    await sleep(1000);
    url = await findLinkByKeywords(page);
    if (url) return url;
  }

  // 브랜드 홈 탐색
  await page.goto(CONFIG.brandUrl, { waitUntil: 'networkidle2', timeout: 60000 });
  await sleep(2000);
  return findLinkByKeywords(page);
}

async function findLinkByKeywords(page) {
  return page.evaluate((keywords) => {
    const normalize = s => s.replace(/\s+/g, '').toLowerCase();
    for (const a of document.querySelectorAll('a[href]')) {
      const text = normalize(a.textContent + ' ' + (a.getAttribute('title') || ''));
      const href = a.getAttribute('href') || '';
      if ((href.includes('products') || href.includes('smartstore')) &&
          keywords.every(kw => text.includes(normalize(kw)))) {
        return href.startsWith('http') ? href : `https://brand.naver.com${href}`;
      }
    }
    for (const img of document.querySelectorAll('img[alt]')) {
      const alt = normalize(img.getAttribute('alt') || '');
      if (keywords.every(kw => alt.includes(normalize(kw)))) {
        const a = img.closest('a[href]');
        if (a) {
          const href = a.getAttribute('href');
          return href.startsWith('http') ? href : `https://brand.naver.com${href}`;
        }
      }
    }
    return null;
  }, CONFIG.productKeywords);
}

// ─────────────────────────────────────────────────────────
// 2단계: 리뷰 API 파라미터 수집
//   - 페이지를 실제로 로드해서 브랜드 서버가 쓰는
//     originProductNo와 checkoutMerchantNo를 가로챕니다.
// ─────────────────────────────────────────────────────────
async function getReviewApiParams(page, productUrl) {
  log(`  제품 페이지 로드: ${productUrl}`);

  // waitForRequest 설정 (페이지 이동 전에 등록)
  const queryPagesPromise = page.waitForRequest(
    req => req.url().includes('query-pages') && req.method() === 'POST',
    { timeout: 30000 }
  ).catch(() => null);

  await page.goto(productUrl, { waitUntil: 'networkidle2', timeout: 60000 });
  await sleep(2000);

  // 리뷰 탭 클릭
  const tabClicked = await page.evaluate(() => {
    const tabs = Array.from(document.querySelectorAll('a, button, [role="tab"], li'));
    for (const el of tabs) {
      const text = (el.textContent || '').replace(/\s+/g, '').trim();
      if ((text.includes('리뷰') || text.includes('구매평')) && el.offsetParent !== null) {
        el.click();
        return true;
      }
    }
    return false;
  });
  if (tabClicked) {
    log('  리뷰 탭 클릭 완료');
    await sleep(2000);
  }

  // 스크롤로 query-pages 호출 유도
  for (let i = 0; i < 20; i++) {
    await page.evaluate(() => window.scrollBy(0, 500));
    await sleep(400);
  }
  await sleep(1000);

  const req = await queryPagesPromise;
  if (!req) return null;

  try {
    const pd = req.postData();
    if (!pd) return null;
    const parsed = JSON.parse(pd);
    return {
      checkoutMerchantNo: parsed.checkoutMerchantNo,
      originProductNo: parsed.originProductNo,
    };
  } catch (_) {
    return null;
  }
}

// ─────────────────────────────────────────────────────────
// 3단계: 리뷰 API 호출 (페이지별)
// ─────────────────────────────────────────────────────────
async function fetchReviewPage(page, params, pageNum) {
  const { checkoutMerchantNo, originProductNo } = params;

  const result = await page.evaluate(async (merchantNo, productNo, pg, pageSize) => {
    try {
      const resp = await fetch('https://brand.naver.com/n/v1/contents/reviews/query-pages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json',
        },
        body: JSON.stringify({
          checkoutMerchantNo: merchantNo,
          originProductNo: productNo,
          page: pg,
          pageSize: pageSize,
          reviewSearchSortType: 'REVIEW_RANKING',
        }),
        credentials: 'include',
      });
      if (!resp.ok) return { error: `HTTP ${resp.status}` };
      const data = await resp.json();
      return { data };
    } catch (e) {
      return { error: e.message };
    }
  }, checkoutMerchantNo, originProductNo, pageNum, CONFIG.reviewPageSize);

  if (result.error) {
    log(`  [경고] 페이지 ${pageNum} API 오류: ${result.error}`);
    return [];
  }

  const contents = result.data?.contents || [];
  return contents.map(item => ({
    rating: String(item.reviewScore || ''),
    userId: item.maskedWriterId || item.writerId || '',
    date: formatDate(item.createDate),
    option: cleanOption(item.productOptionContent || ''),
    content: (item.reviewContent || '').replace(/\n/g, ' ').trim(),
  }));
}

function cleanOption(raw) {
  // "▶오늘출발/5만원 이상 무료배송/무료교환◀: 화이트 M (45x75cm)"
  // → "화이트 M (45x75cm)"

  // ◀ 뒤에 오는 ": " 이후 내용 추출
  const arrowMatch = raw.match(/◀[:\s:：]+(.+)$/);
  if (arrowMatch) return arrowMatch[1].trim();

  // ": " 이후 내용 추출 (일반 옵션 형식)
  const colonMatch = raw.match(/[：:]\s*(.+)$/);
  if (colonMatch) return colonMatch[1].trim();

  return raw.trim();
}

// ─────────────────────────────────────────────────────────
// 4단계: 엑셀 저장
// ─────────────────────────────────────────────────────────
function saveToExcel(reviews) {
  const wb = XLSX.utils.book_new();

  const rows = [
    ['별점', '아이디', '날짜', '구매옵션', '리뷰내용'],
    ...reviews.map(r => [r.rating, r.userId, r.date, r.option, r.content]),
  ];

  const ws = XLSX.utils.aoa_to_sheet(rows);

  ws['!cols'] = [
    { wch: 6 },    // 별점
    { wch: 18 },   // 아이디
    { wch: 14 },   // 날짜
    { wch: 28 },   // 구매옵션
    { wch: 100 },  // 리뷰내용
  ];

  XLSX.utils.book_append_sheet(wb, ws, '리뷰');

  const fileName = `리뷰_코에르욕실매트S_${timestamp()}.xlsx`;
  const outputPath = path.join(CONFIG.outputDir, fileName);
  XLSX.writeFile(wb, outputPath);
  return outputPath;
}

// ─────────────────────────────────────────────────────────
// 메인
// ─────────────────────────────────────────────────────────
async function main() {
  log('');
  log('╔══════════════════════════════════════════════════╗');
  log('║   코에르 리뷰 수집 프로그램                      ║');
  log('║   욕실매트 S | 네이버 브랜드스토어               ║');
  log('╚══════════════════════════════════════════════════╝');
  log('');

  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: { width: 1280, height: 900 },
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });

  const page = await browser.newPage();
  await page.setUserAgent(
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
  );

  const allReviews = [];

  try {
    // ── 1단계: 제품 URL ──────────────────────────
    log('[1단계] 제품 페이지 탐색 중...');
    const productUrl = await findProductUrl(page);

    if (!productUrl) {
      log('[오류] 제품을 찾지 못했습니다.');
      log('  → 브랜드 스토어에서 제품 URL을 직접 복사하여');
      log('     CONFIG.manualProductUrl 에 입력 후 재실행하세요.');
      await browser.close();
      return;
    }
    log(`  ✓ 발견: ${productUrl}\n`);

    // ── 2단계: API 파라미터 획득 ─────────────────
    log('[2단계] 리뷰 API 파라미터 수집 중...');
    const apiParams = await getReviewApiParams(page, productUrl);

    if (!apiParams) {
      log('[오류] 리뷰 API 파라미터를 찾지 못했습니다.');
      log('  → 네이버 로그인이 필요하거나, 페이지 구조가 변경되었을 수 있습니다.');
      await browser.close();
      return;
    }
    log(`  ✓ 상품번호: ${apiParams.originProductNo}, 판매자번호: ${apiParams.checkoutMerchantNo}\n`);

    // ── 3단계: 리뷰 수집 ─────────────────────────
    log(`[3단계] 리뷰 수집 (${CONFIG.maxReviewPages}페이지 × 최대 ${CONFIG.reviewPageSize}개)`);

    for (let pg = 1; pg <= CONFIG.maxReviewPages; pg++) {
      log(`  [${pg}/${CONFIG.maxReviewPages}페이지] 수집 중...`);
      const reviews = await fetchReviewPage(page, apiParams, pg);
      log(`  → ${reviews.length}개 수집`);
      allReviews.push(...reviews);

      if (reviews.length < CONFIG.reviewPageSize) {
        log('  → 마지막 페이지 도달, 수집 종료');
        break;
      }

      if (pg < CONFIG.maxReviewPages) await sleep(500);
    }

    log('');

    // ── 4단계: 엑셀 저장 ─────────────────────────
    if (allReviews.length === 0) {
      log('[결과] 수집된 리뷰가 없습니다.');
    } else {
      log(`[4단계] 총 ${allReviews.length}개 리뷰 → 엑셀 저장 중...`);
      const outputPath = saveToExcel(allReviews);
      log(`  ✓ 저장 완료!`);
      log(`  📄 파일: ${outputPath}`);
    }

  } catch (err) {
    log(`\n[오류] ${err.message}`);
  } finally {
    await browser.close();
    log('');
    log('╔══════════════════════════════════════════════════╗');
    log('║   완료                                           ║');
    log('╚══════════════════════════════════════════════════╝');
  }
}

main().catch(console.error);
