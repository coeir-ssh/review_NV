/**
 * 코에르 리뷰 자동 답변 생성기
 *
 * 1. 브랜드스토어 전체 제품 탐색
 * 2. 제품별 1페이지 리뷰 중 답변 없는 것 상단 3개 수집
 * 3. Claude Sonnet API로 개인화 답변 생성
 * 4. 엑셀 저장 (제품명 | 별점 | 아이디 | 날짜 | 구매옵션 | 리뷰내용 | 답변)
 *
 * 실행 전: config.js에 ANTHROPIC_API_KEY 입력 또는
 *          환경변수 ANTHROPIC_API_KEY 설정
 */

const puppeteer = require('puppeteer');
const Anthropic = require('@anthropic-ai/sdk');
const XLSX = require('xlsx');
const path = require('path');
const cfg = require('./config');

// ─────────────────────────────────────────────────────────
// API 키 확인
// ─────────────────────────────────────────────────────────
const API_KEY = cfg.ANTHROPIC_API_KEY || process.env.ANTHROPIC_API_KEY;
if (!API_KEY) {
  console.error('');
  console.error('[오류] Anthropic API 키가 설정되지 않았습니다.');
  console.error('  ① https://console.anthropic.com 에서 API 키 발급');
  console.error('  ② config.js 파일의 ANTHROPIC_API_KEY 값을 입력하거나');
  console.error('     환경변수: set ANTHROPIC_API_KEY=sk-ant-xxx');
  console.error('');
  process.exit(1);
}

const anthropic = new Anthropic({ apiKey: API_KEY });

// ─────────────────────────────────────────────────────────
// 설정
// ─────────────────────────────────────────────────────────
const CONFIG = {
  brandUrl: 'https://brand.naver.com/coeir',
  channelId: '2sWDzdkTnmtC0eKzChuo9',
  reviewsPerProduct: 3,       // 제품당 수집할 답변 없는 리뷰 수
  reviewPageSize: 20,         // API 한 번에 가져올 리뷰 수 (1페이지)
  defaultMerchantNo: 510909573,
  model: 'claude-sonnet-4-5',  // Claude Sonnet
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

function hasReply(item) {
  return Array.isArray(item.reviewComments) && item.reviewComments.length > 0;
}

// ─────────────────────────────────────────────────────────
// 1단계: 전체 제품 목록 수집
// ─────────────────────────────────────────────────────────
async function getAllProducts(page) {
  const products = new Map();

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
              });
            }
          }
        }
      } catch (_) {}
    }
  };
  page.on('response', apiHandler);

  log('  카테고리 페이지 탐색 중...');
  await page.goto(`${CONFIG.brandUrl}/category/ALL`, { waitUntil: 'networkidle2', timeout: 60000 });
  await sleep(3000);

  let prevSize = 0;
  for (let i = 0; i < 15; i++) {
    await page.evaluate(() => window.scrollBy(0, 800));
    await sleep(700);
    if (products.size > prevSize) {
      prevSize = products.size;
    } else if (i > 5) break;
  }

  page.off('response', apiHandler);

  // 방법 2: 직접 API 호출
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
          const items = data.simpleProducts || data.products || [];
          if (!items.length) break;
          items.forEach(p => results.push({
            name: p.name || '',
            originProductNo: p.originProductNo || p.id,
          }));
          if (items.length < 40) break;
          pg++;
        } catch (_) { break; }
      }
      return results;
    }, CONFIG.channelId);

    for (const p of apiProducts) {
      if (p.originProductNo && p.name) products.set(p.originProductNo, p);
    }
  }

  // 방법 3: 페이지 링크 스크래핑 후 개별 제품 페이지 방문
  if (products.size === 0) {
    log('  페이지 링크 스크래핑으로 제품 URL 수집...');
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

    log(`  → ${links.length}개 제품 URL 발견, 개별 페이지에서 originProductNo 획득 중...`);
    for (const link of links) {
      const info = await getOriginProductNo(page, link.url);
      if (info && info.originProductNo) {
        products.set(info.originProductNo, {
          name: info.productName,
          originProductNo: info.originProductNo,
        });
      }
      await sleep(500);
    }
  }

  return Array.from(products.values());
}

// ─────────────────────────────────────────────────────────
// 제품 페이지에서 originProductNo 획득
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
// 2단계: 제품별 1페이지 리뷰 수집 (답변 없는 것 상단 N개)
// ─────────────────────────────────────────────────────────
async function fetchTopUnrepliedReviews(page, originProductNo) {
  const result = await page.evaluate(async (merchantNo, productNo, pageSize) => {
    try {
      const resp = await fetch('https://brand.naver.com/n/v1/contents/reviews/query-pages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          checkoutMerchantNo: merchantNo,
          originProductNo: productNo,
          page: 1,
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
  }, CONFIG.defaultMerchantNo, originProductNo, CONFIG.reviewPageSize);

  if (result.error) {
    log(`    [경고] 리뷰 수집 오류: ${result.error}`);
    return [];
  }

  // 1페이지 상단부터 순서대로 답변 없는 리뷰만 최대 N개
  const unreplied = result.items
    .filter(item => !hasReply(item))
    .slice(0, CONFIG.reviewsPerProduct);

  return unreplied;
}

// ─────────────────────────────────────────────────────────
// 3단계: Claude Sonnet으로 개인화 답변 생성
// ─────────────────────────────────────────────────────────
const SYSTEM_PROMPT = `당신은 코에르(COEIR) 브랜드의 고객 서비스 담당자입니다.

【브랜드 소개】
코에르는 고품질 욕실 용품 전문 한국 브랜드입니다.
- 욕실 소품: 스텐/테라조 디스펜서·멀티홀더·트레이·비누받침대·칫솔꽂이 (무타공 설치 가능)
- 욕실 매트: 규조토 발매트, 미끄럼방지 매트
- 샤워용품: 온오프 샤워헤드, PLA 정수필터 (절수·수압 개선)
- 수건: 40수 코마사 프리미엄 타올 (호텔 퀄리티)
- 욕실선반 (플랫/히든 타입), 욕실화

【답변 작성 규칙】
1. 구조: 인사 한 줄 → 리뷰 내용 기반 개인화 문단 → 마무리 한 줄
2. 인사: "안녕하세요, 고객님!" 또는 "안녕하세요, 고객님." 등 자연스럽게 변형
3. 개인화 문단 (핵심):
   - 리뷰에서 고객이 언급한 구체적인 내용을 반드시 1~2가지 직접 언급
   - 고객의 상황(아이 있음, 전세집, 리모델링 중, 재구매, 선물 등)을 파악해 맞춤 응대
   - "선물용으로도 좋을 것 같아요" 같은 추천 표현은 선물 구매로 오해하지 말 것
   - AI가 작성한 티가 나지 않게, 실제 사람이 리뷰를 읽고 공감하며 쓴 것처럼 자연스럽게
   - 100자 이상 작성
4. 마무리: "리뷰 감사드리며, 앞으로도 코에르 제품 많이 사랑해주세요. 좋은 하루 보내세요!"
   (약간의 변형 허용: "편안한 하루 되세요!", "행복한 하루 되세요!" 등)
5. 별점 3점 이하 리뷰는 불편함에 공감하고 개선 안내를 포함할 것
6. 각 리뷰마다 반드시 다른 톤·표현 사용 (복붙 느낌 금지)

【출력 형식】
반드시 유효한 JSON 배열만 출력하세요. 다른 텍스트 없이 JSON만:
[
  {"idx": 0, "response": "답변 전체 내용"},
  {"idx": 1, "response": "답변 전체 내용"}
]`;

async function generateResponses(reviews) {
  if (reviews.length === 0) return [];

  const input = reviews.map((r, i) => ({
    idx: i,
    product: r.productName,
    rating: r.rating,
    option: r.option,
    content: r.content,
  }));

  const resp = await anthropic.messages.create({
    model: CONFIG.model,
    max_tokens: 4096,
    system: SYSTEM_PROMPT,
    messages: [{
      role: 'user',
      content: `다음 리뷰들에 대해 각각 개인화된 답변을 작성해주세요:\n\n${JSON.stringify(input, null, 2)}`,
    }],
  });

  const text = resp.content[0].text.trim();
  const match = text.match(/\[[\s\S]*\]/);
  if (!match) throw new Error(`JSON 파싱 실패: ${text.substring(0, 200)}`);
  return JSON.parse(match[0]);
}

// ─────────────────────────────────────────────────────────
// 4단계: 엑셀 저장
// ─────────────────────────────────────────────────────────
function saveToExcel(rows) {
  const wb = XLSX.utils.book_new();
  // reviewId 컬럼 추가 (post_to_naver.js에서 활용)
  const HEADERS = ['reviewId', '제품명', '별점', '아이디', '날짜', '구매옵션', '리뷰내용', '답변', '포스팅완료'];
  const COL_WIDTHS = [
    { wch: 18 },  // reviewId
    { wch: 40 },  // 제품명
    { wch: 5 },   // 별점
    { wch: 15 },  // 아이디
    { wch: 12 },  // 날짜
    { wch: 25 },  // 구매옵션
    { wch: 60 },  // 리뷰내용
    { wch: 80 },  // 답변
    { wch: 12 },  // 포스팅완료
  ];

  const ws = XLSX.utils.aoa_to_sheet([HEADERS, ...rows]);
  ws['!cols'] = COL_WIDTHS;
  XLSX.utils.book_append_sheet(wb, ws, '리뷰답변');

  const outPath = path.join(CONFIG.outputDir, `리뷰답변_자동_${timestamp()}.xlsx`);
  XLSX.writeFile(wb, outPath);
  return outPath;
}

// ─────────────────────────────────────────────────────────
// 메인
// ─────────────────────────────────────────────────────────
async function main() {
  log('');
  log('╔══════════════════════════════════════════════════════════════╗');
  log('║   코에르 리뷰 자동 답변 생성기                              ║');
  log('║   수집: 제품별 1페이지 상단 / 답변없는 리뷰 3개            ║');
  log(`║   AI: Claude Sonnet (${CONFIG.model})  ║`);
  log('╚══════════════════════════════════════════════════════════════╝');
  log('');

  const browser = await puppeteer.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });

  const allReviews = [];   // { productName, rating, id, date, option, content }
  const allResponses = []; // 엑셀 행 데이터

  try {
    const page = await browser.newPage();
    await page.setUserAgent(
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
    );

    // ── 1단계: 제품 목록 수집 ─────────────────────────────
    log('[1단계] 전체 제품 목록 수집...');
    const products = await getAllProducts(page);
    log(`  → ${products.length}개 제품 발견`);
    log('');

    // ── 2단계: 제품별 리뷰 수집 ──────────────────────────
    log('[2단계] 제품별 리뷰 수집 (1페이지, 답변없는 리뷰 최대 3개)');

    const seenOriginNos = new Set();

    for (let i = 0; i < products.length; i++) {
      const { name, originProductNo } = products[i];
      if (!originProductNo || seenOriginNos.has(originProductNo)) continue;
      seenOriginNos.add(originProductNo);

      process.stdout.write(`  [${String(i+1).padStart(2)}/${products.length}] ${name.substring(0, 30)}... `);

      const items = await fetchTopUnrepliedReviews(page, originProductNo);
      log(`${items.length}개 수집`);

      for (const item of items) {
        const content = (item.reviewContent || '').replace(/\n/g, ' ').trim();
        allReviews.push({
          reviewId: item.id || item.reviewId || '',   // 네이버 답변 posting용
          productName: name,
          rating: String(item.reviewScore || ''),
          id: item.maskedWriterId || item.writerId || '',
          date: formatDate(item.createDate),
          option: cleanOption(item.productOptionContent || ''),
          content,
        });
      }

      await sleep(200);
    }

    log('');
    log(`  총 ${allReviews.length}개 리뷰 수집 완료`);
    log('');

    if (allReviews.length === 0) {
      log('[결과] 수집된 리뷰가 없습니다. (모든 리뷰에 답변이 달렸거나 리뷰 없음)');
      return;
    }

    // ── 3단계: API 답변 생성 (배치 10개씩) ───────────────
    log(`[3단계] Claude Sonnet API 답변 생성 (총 ${allReviews.length}개)`);

    const BATCH_SIZE = 5;   // 배치 크기 줄여서 rate limit 방지
    const responsesMap = new Map(); // idx → response

    for (let start = 0; start < allReviews.length; start += BATCH_SIZE) {
      const batch = allReviews.slice(start, start + BATCH_SIZE);
      const batchNum = Math.floor(start / BATCH_SIZE) + 1;
      const totalBatches = Math.ceil(allReviews.length / BATCH_SIZE);

      process.stdout.write(`  배치 ${batchNum}/${totalBatches} (${start+1}~${start+batch.length}번)... `);

      let aiResults = null;
      for (let retry = 0; retry < 5; retry++) {
        try {
          aiResults = await generateResponses(batch);
          break;
        } catch (err) {
          const waitSec = 10 * (retry + 1); // 10초, 20초, 30초, 40초, 50초
          process.stdout.write(`재시도${retry+1}(${waitSec}s대기) `);
          await sleep(waitSec * 1000);
        }
      }

      if (aiResults) {
        for (const r of aiResults) {
          responsesMap.set(start + r.idx, r.response || '');
        }
        log('✓');
      } else {
        log('(실패 - 기본값 적용)');
        for (let j = 0; j < batch.length; j++) {
          responsesMap.set(start + j, '리뷰 감사합니다. 앞으로도 코에르 제품 많이 사랑해주세요 :)');
        }
      }

      // 배치 사이 대기 (rate limit 방지)
      if (start + BATCH_SIZE < allReviews.length) await sleep(3000);
    }

    // ── 4단계: 엑셀 저장 ─────────────────────────────────
    log('');
    log('[4단계] 엑셀 저장...');

    for (let i = 0; i < allReviews.length; i++) {
      const r = allReviews[i];
      const response = responsesMap.get(i) || '';
      allResponses.push([
        r.reviewId,    // 네이버 답변 posting용 ID
        r.productName,
        r.rating,
        r.id,
        r.date,
        r.option,
        r.content,
        response,
        '',            // 포스팅완료 (post_to_naver.js 실행 후 체크됨)
      ]);
    }

    const outPath = saveToExcel(allResponses);
    log(`  ✓ 저장 완료: ${outPath}`);

    // ── 통계 출력 ──────────────────────────────────────
    log('');
    log('─'.repeat(60));
    log(`  수집 제품 수  : ${seenOriginNos.size}개`);
    log(`  수집 리뷰 수  : ${allReviews.length}개`);
    log(`  답변 생성 수  : ${allResponses.filter(r => r[6]).length}개`);
    log('─'.repeat(60));

  } catch (err) {
    log(`\n[오류] ${err.message}`);
    log(err.stack);
  } finally {
    await browser.close();
  }

  log('');
  log('╔══════════════════════════════════════════════════════════════╗');
  log('║   완료! 엑셀 파일을 열어 답변을 확인하세요.                 ║');
  log('╚══════════════════════════════════════════════════════════════╝');
}

main().catch(err => {
  console.error('\n[오류]', err.message);
  process.exit(1);
});
