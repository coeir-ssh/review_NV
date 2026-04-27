/**
 * 코에르 전체 제품 리뷰 수집 프로그램 (헤드리스 모드)
 *
 * - 브랜드스토어 전체 제품 자동 탐색
 * - 제품별 리뷰 3페이지씩 수집
 * - 이미 코에르가 답변한 리뷰 자동 제외 (reviewComments 필드)
 * - 환불검토 자동 표시
 * - 엑셀: 제품명 | 별점 | 아이디 | 날짜 | 구매옵션 | 리뷰내용 | 환불
 */

const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const path = require('path');

// ─────────────────────────────────────────────────────────
// 설정
// ─────────────────────────────────────────────────────────
const CONFIG = {
  brandUrl: 'https://brand.naver.com/coeir',
  channelId: '2sWDzdkTnmtC0eKzChuo9', // coeir 브랜드 채널 ID
  maxReviewPages: 3,
  reviewPageSize: 20,
  defaultMerchantNo: 510909573,
  outputDir: 'C:\\Users\\AWESOMATIC\\Desktop\\코에르\\클로드\\리뷰수집프로그램',
};

const sleep = ms => new Promise(r => setTimeout(r, ms));
const log = msg => console.log(msg);

// ─────────────────────────────────────────────────────────
// 유틸
// ─────────────────────────────────────────────────────────
function formatDate(iso) {
  if (!iso) return '';
  try {
    const d = new Date(iso);
    const p = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}.${p(d.getMonth()+1)}.${p(d.getDate())}`;
  } catch { return iso; }
}

function cleanOption(raw) {
  if (!raw) return '';
  const m1 = raw.match(/◀[:\s：]+(.+)$/);
  if (m1) return m1[1].trim();
  const m2 = raw.match(/[：:]\s*([^:]+)$/);
  if (m2) return m2[1].trim();
  return raw.trim();
}

function timestamp() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}

// ─────────────────────────────────────────────────────────
// 답변 여부 확인
//   reviewComments 배열이 있으면 코에르가 답변을 단 리뷰
// ─────────────────────────────────────────────────────────
function hasReply(item) {
  return Array.isArray(item.reviewComments) && item.reviewComments.length > 0;
}

// ─────────────────────────────────────────────────────────
// 환불검토 판단
//   답변으로 해결하기 어려운 리뷰 → '환불검토' 표시
// ─────────────────────────────────────────────────────────
const REFUND_STRONG = [
  '환불', '반품', '불량품', '불량', '파손', '하자', '찢어',
  '끊어', '다쳤', '다침', '부러', '사고',
  '위험해', '위험합', '허위광고', '사기',
];
const REFUND_WEAK = [
  '최악', '쓰레기', '아예 안', '전혀 안', '아예안', '전혀안',
  '사용불가', '사용 불가', '아무 소용', '효과없', '작동안',
];

function shouldRefund(content, ratingStr) {
  if (!content) return false;
  // 강한 신호 - 별점 무관
  if (REFUND_STRONG.some(kw => content.includes(kw))) return true;
  const r = parseInt(ratingStr) || 5;
  // 2점 이하 + 약한 신호
  if (r <= 2 && REFUND_WEAK.some(kw => content.includes(kw))) return true;
  // 1점 + 짧지 않은 내용 + 취향/선호가 아닌 문제
  if (r === 1 && content.length > 30) {
    const notRefund = ['색깔', '색상', '취향', '개인적', '기대보다', '생각보다 작', '생각보다 큰'];
    if (!notRefund.some(kw => content.includes(kw))) return true;
  }
  return false;
}

// ─────────────────────────────────────────────────────────
// 1단계: 전체 제품 목록 수집
//   display-products API + 페이지 스크래핑 병행
// ─────────────────────────────────────────────────────────
async function getAllProducts(page) {
  const products = new Map(); // originProductNo → { name, originProductNo }

  // display-products API 응답 캡처
  const apiHandler = async res => {
    const url = res.url();
    if (url.includes('display-products') && !url.includes('disp-price-range') &&
        !url.includes('display-products/')) {
      try {
        const body = await res.json().catch(() => null);
        if (body && body.simpleProducts) {
          for (const p of body.simpleProducts) {
            if (p.originProductNo && p.name) {
              products.set(p.originProductNo, {
                name: p.name,
                originProductNo: p.originProductNo,
                productNo: p.id || p.productNo,
              });
            }
          }
        }
      } catch (_) {}
    }
  };
  page.on('response', apiHandler);

  log('  카테고리 페이지 탐색 (API 캡처 중)...');
  await page.goto(`${CONFIG.brandUrl}/category/ALL`, { waitUntil: 'networkidle2', timeout: 60000 });
  await sleep(3000);

  // 무한 스크롤 대응
  let prevSize = 0;
  for (let i = 0; i < 15; i++) {
    await page.evaluate(() => window.scrollBy(0, 800));
    await sleep(700);
    if (products.size > prevSize) {
      prevSize = products.size;
    } else if (i > 5) {
      break; // 더 이상 새 제품 없으면 종료
    }
  }

  page.off('response', apiHandler);

  // API로 제품을 못 찾은 경우: 직접 API 호출 시도
  if (products.size === 0) {
    log('  직접 API 호출로 제품 목록 조회...');
    const apiProducts = await page.evaluate(async (channelId) => {
      const results = [];
      let pg = 1;
      while (true) {
        const url = `https://brand.naver.com/n/v2/channels/${channelId}/display-products?` +
          `categoryId=ALL&categoryType=DISPCATG&page=${pg}&pageSize=40&sort=POPULAR`;
        try {
          const resp = await fetch(url, { credentials: 'include' });
          if (!resp.ok) break;
          const data = await resp.json();
          const items = data.simpleProducts || data.products || data.items || [];
          if (!items.length) break;
          items.forEach(p => results.push({
            name: p.name || p.productName || '',
            originProductNo: p.originProductNo || p.id,
            productNo: p.id || p.productNo,
          }));
          if (items.length < 40) break;
          pg++;
        } catch (_) { break; }
      }
      return results;
    }, CONFIG.channelId);

    for (const p of apiProducts) {
      if (p.originProductNo && p.name) {
        products.set(p.originProductNo, p);
      }
    }
  }

  // 여전히 없으면 페이지에서 링크 스크래핑 + originProductNo 개별 조회
  if (products.size === 0) {
    log('  페이지 링크에서 제품 URL 수집...');
    const links = await page.evaluate(() => {
      const seen = new Set();
      const result = [];
      document.querySelectorAll('a[href*="/products/"]').forEach(a => {
        const href = a.getAttribute('href') || '';
        const m = href.match(/\/products\/(\d+)/);
        if (m && !seen.has(m[1])) {
          seen.add(m[1]);
          result.push({
            url: href.startsWith('http') ? href : `https://brand.naver.com${href}`,
            productNo: m[1],
          });
        }
      });
      return result;
    });
    return { type: 'urls', items: links }; // 별도 처리
  }

  return { type: 'products', items: Array.from(products.values()) };
}

// ─────────────────────────────────────────────────────────
// 2단계: 제품 페이지에서 originProductNo 획득
//   review-summary API URL에서 추출 (빠름 - 페이지 완전 로드 불필요)
// ─────────────────────────────────────────────────────────
async function getOriginProductNo(page, productUrl) {
  const reqPromise = page.waitForRequest(
    req => /\/contents\/reviews\/product-summary\/\d+\?/.test(req.url()),
    { timeout: 15000 }
  ).catch(() => null);

  await page.goto(productUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });

  const req = await reqPromise;
  if (!req) return null;

  const m = req.url().match(/product-summary\/(\d+)\?/);
  if (!m) return null;

  // 제품명도 페이지에서 추출
  const productName = await page.evaluate(() => {
    const selectors = ['h3', 'h2', 'h1', '[class*="productTitle"]', '[class*="product_title"]'];
    for (const sel of selectors) {
      const el = document.querySelector(sel);
      if (el && el.textContent.trim().length > 3 && el.textContent.trim().length < 200) {
        return el.textContent.trim();
      }
    }
    return document.title.split('|')[0].trim();
  });

  return { originProductNo: parseInt(m[1]), productName };
}

// ─────────────────────────────────────────────────────────
// 3단계: 리뷰 API 호출 (query-pages POST)
// ─────────────────────────────────────────────────────────
async function fetchReviews(page, originProductNo, maxPages) {
  const allItems = [];

  for (let pg = 1; pg <= maxPages; pg++) {
    const result = await page.evaluate(async (merchantNo, productNo, pgNum, pageSize) => {
      try {
        const resp = await fetch('https://brand.naver.com/n/v1/contents/reviews/query-pages', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            checkoutMerchantNo: merchantNo,
            originProductNo: productNo,
            page: pgNum,
            pageSize: pageSize,
            reviewSearchSortType: 'REVIEW_RANKING',
          }),
          credentials: 'include',
        });
        if (!resp.ok) return { error: `HTTP ${resp.status}`, items: [] };
        const data = await resp.json();
        return { items: data.contents || [], total: data.totalCount };
      } catch (e) {
        return { error: e.message, items: [] };
      }
    }, CONFIG.defaultMerchantNo, originProductNo, pg, CONFIG.reviewPageSize);

    if (result.error) {
      log(`    [경고] 페이지 ${pg} 오류: ${result.error}`);
      break;
    }

    allItems.push(...result.items);
    if (result.items.length < CONFIG.reviewPageSize) break; // 마지막 페이지
    await sleep(300);
  }

  return allItems;
}

// ─────────────────────────────────────────────────────────
// 4단계: 엑셀 저장
// ─────────────────────────────────────────────────────────
function saveToExcel(rows, summary) {
  const wb = XLSX.utils.book_new();

  // ── 리뷰 시트 ──────────────────────────────────────────
  const headers = ['제품명', '별점', '아이디', '날짜', '구매옵션', '리뷰내용', '환불'];
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);

  ws['!cols'] = [
    { wch: 40 },   // 제품명
    { wch: 6 },    // 별점
    { wch: 18 },   // 아이디
    { wch: 14 },   // 날짜
    { wch: 25 },   // 구매옵션
    { wch: 100 },  // 리뷰내용
    { wch: 10 },   // 환불
  ];

  // 환불검토 행 강조 (글자 굵게)
  // xlsx 기본 라이브러리는 스타일 미지원 → 별도 시트로 환불 목록 제공

  XLSX.utils.book_append_sheet(wb, ws, '전체리뷰');

  // ── 환불검토 시트 ──────────────────────────────────────
  const refundRows = rows.filter(r => r[6] === '환불검토');
  if (refundRows.length > 0) {
    const ws2 = XLSX.utils.aoa_to_sheet([headers, ...refundRows]);
    ws2['!cols'] = ws['!cols'];
    XLSX.utils.book_append_sheet(wb, ws2, '환불검토');
  }

  // ── 요약 시트 ─────────────────────────────────────────
  const summaryData = [
    ['제품명', '수집 리뷰', '답변 제외', '환불검토'],
    ...summary,
  ];
  const ws3 = XLSX.utils.aoa_to_sheet(summaryData);
  ws3['!cols'] = [{ wch: 40 }, { wch: 12 }, { wch: 12 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, ws3, '요약');

  const outputPath = path.join(CONFIG.outputDir, `리뷰_코에르_전체_${timestamp()}.xlsx`);
  XLSX.writeFile(wb, outputPath);
  return outputPath;
}

// ─────────────────────────────────────────────────────────
// 5단계: 리뷰 답변 API 조사
// ─────────────────────────────────────────────────────────
async function investigateReplyAPI(page) {
  log('\n[답변 API 조사]');

  const result = await page.evaluate(async () => {
    // Naver Brand Store 댓글(답변) 등록 엔드포인트 후보
    const testReviewId = 4947947260;
    const candidates = [
      { method: 'POST', url: `https://brand.naver.com/n/v1/contents/reviews/${testReviewId}/comments` },
      { method: 'GET', url: `https://brand.naver.com/n/v1/contents/reviews/${testReviewId}/comments` },
      { method: 'POST', url: `https://brand.naver.com/n/v1/review/comment` },
      { method: 'GET', url: 'https://api.commerce.naver.com/external/v1/review/replies' },
    ];

    const results = [];
    for (const c of candidates) {
      try {
        const resp = await fetch(c.url, {
          method: c.method,
          headers: c.method === 'POST' ? { 'Content-Type': 'application/json' } : {},
          body: c.method === 'POST' ? '{}' : undefined,
          credentials: 'include',
        });
        const text = await resp.text();
        results.push({ method: c.method, url: c.url, status: resp.status, body: text.substring(0, 200) });
      } catch (e) {
        results.push({ method: c.method, url: c.url, error: e.message });
      }
    }
    return results;
  });

  result.forEach(r => {
    log(`  ${r.method} ${r.url}`);
    log(`    상태: ${r.status || '오류'} | ${r.body || r.error || ''}`);
  });

  return result;
}

// ─────────────────────────────────────────────────────────
// 메인
// ─────────────────────────────────────────────────────────
async function main() {
  log('');
  log('╔══════════════════════════════════════════════════════╗');
  log('║   코에르 전체 제품 리뷰 수집 프로그램               ║');
  log('║   (헤드리스 - 브라우저 창 없음)                     ║');
  log('╚══════════════════════════════════════════════════════╝');
  log('');

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });

  const allRows = [];    // 엑셀 데이터 행
  const summaryRows = []; // 요약 시트용

  try {
    const page = await browser.newPage();
    await page.setUserAgent(
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
    );

    // ── 1단계: 전체 제품 목록 ──────────────────────────
    log('[1단계] 전체 제품 목록 수집...');
    const productData = await getAllProducts(page);
    log(`  → ${productData.items.length}개 제품 항목 발견`);
    log('');

    const seenOriginNos = new Set();
    let productList = []; // { name, originProductNo, url? }

    if (productData.type === 'products') {
      // API로 받아온 경우: originProductNo 이미 있음
      productList = productData.items;
    } else {
      // URL 스크래핑으로 받아온 경우: 개별 페이지에서 originProductNo 획득 필요
      productList = productData.items.map(item => ({
        url: item.url,
        originProductNo: null,
        name: null,
      }));
    }

    // ── 2단계: 제품별 리뷰 수집 ───────────────────────
    log(`[2단계] 제품별 리뷰 수집 (3페이지씩)`);

    for (let i = 0; i < productList.length; i++) {
      let { name, originProductNo, url } = productList[i];

      // originProductNo가 없으면 제품 페이지에서 획득
      if (!originProductNo && url) {
        const info = await getOriginProductNo(page, url);
        if (!info) {
          log(`  [${i+1}] URL: ${url} → originProductNo 획득 실패, 건너뜀`);
          continue;
        }
        originProductNo = info.originProductNo;
        if (!name) name = info.productName;
      }

      if (!originProductNo) continue;

      // 중복 제품 제외
      if (seenOriginNos.has(originProductNo)) {
        log(`  [${i+1}] 중복 originProductNo(${originProductNo}), 건너뜀`);
        continue;
      }
      seenOriginNos.add(originProductNo);

      log(`  [${i+1}/${productList.length}] ${name || `제품 ${originProductNo}`}`);

      // 리뷰 수집
      const rawItems = await fetchReviews(page, originProductNo, CONFIG.maxReviewPages);
      log(`    원본: ${rawItems.length}개`);

      // 처리 & 필터
      let repliedCount = 0;
      let refundCount = 0;

      for (const item of rawItems) {
        // 코에르 답변 달린 리뷰 제외
        if (hasReply(item)) { repliedCount++; continue; }

        const content = (item.reviewContent || '').replace(/\n/g, ' ').trim();
        const rating = String(item.reviewScore || '');
        const refund = shouldRefund(content, rating) ? '환불검토' : '';
        if (refund) refundCount++;

        const productName = name || item.productName || `제품 ${originProductNo}`;
        allRows.push([
          productName,
          rating,
          item.maskedWriterId || item.writerId || '',
          formatDate(item.createDate),
          cleanOption(item.productOptionContent || ''),
          content,
          refund,
        ]);
      }

      const collected = rawItems.length - repliedCount;
      log(`    수집: ${collected}개 | 답변제외: ${repliedCount}개 | 환불검토: ${refundCount}개`);

      summaryRows.push([
        name || `제품 ${originProductNo}`,
        collected,
        repliedCount,
        refundCount,
      ]);
      log('');
    }

    // ── 3단계: 엑셀 저장 ───────────────────────────────
    if (allRows.length === 0) {
      log('[결과] 수집된 리뷰가 없습니다.');
    } else {
      log(`[3단계] 엑셀 저장 (총 ${allRows.length}개 리뷰)...`);
      const outputPath = saveToExcel(allRows, summaryRows);
      log(`  ✓ 저장 완료: ${outputPath}`);
      log(`  ✓ 시트: 전체리뷰 / 환불검토 / 요약`);
    }

    // ── 4단계: 답변 API 조사 ───────────────────────────
    await investigateReplyAPI(page);

  } catch (err) {
    log(`\n[오류] ${err.message}`);
    log(err.stack);
  } finally {
    await browser.close();
  }

  log('');
  log('╔══════════════════════════════════════════════════════╗');
  log('║   완료                                               ║');
  log('╚══════════════════════════════════════════════════════╝');
}

main().catch(console.error);
