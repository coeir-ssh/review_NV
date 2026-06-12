/**
 * collect_pending_cafe24.js — 자사몰(카페24) 최근 리뷰 수집 (나쁜 리뷰 모니터링용)
 *
 * AI 호출 없음. 카페24 Admin API 로 최근 N일 리뷰 게시글을 수집하되,
 * 이미 처리(알림)한 글(cafe24_seen.json)은 제외하고 pending_reviews_cafe24.json 으로 출력.
 *
 * ⚠️ 방향 전환: 자사몰은 카페24 답글이 공식몰에 노출되지 않아 자동답변이 무의미.
 *    대신 "나쁜 리뷰 감지 → 슬랙 알림 → (추후) 삭제" 파이프라인을 사용한다.
 *
 * 다음 단계: 에이전트가 이 JSON 을 읽고 각 리뷰를 '삭제대상' / '유지' 로 판별 →
 *           screened_reviews_cafe24.json 작성 → alert_bad_reviews_cafe24.js 가 슬랙 알림.
 *
 * 사용:
 *   node collect_pending_cafe24.js                 # 최근 N일 미처리 리뷰 수집
 *   node collect_pending_cafe24.js --baseline      # 현재 리뷰를 모두 '처리됨'으로 시드(최초 1회, 알림 폭주 방지)
 *   node collect_pending_cafe24.js --days=3 --force
 */

const path = require('path');
const fs   = require('fs');

const cafe24 = require('./cafe24_api');

const log = msg => console.log(msg);

const SEEN_FILE = path.join(__dirname, 'cafe24_seen.json');

// CLI 옵션
const argv = process.argv.slice(2);
const getArg = (k, d) => {
  const a = argv.find(x => x.startsWith(`--${k}=`));
  return a ? a.split('=')[1] : d;
};
const baseline = argv.includes('--baseline');
const outFile  = getArg('out', 'pending_reviews_cafe24.json');
const days     = parseInt(getArg('days', baseline ? '30' : '3'), 10) || (baseline ? 30 : 3);
const force    = argv.includes('--force');

// ── seen-set (이미 처리한 article_no) ───────────────────────
function loadSeen() {
  if (!fs.existsSync(SEEN_FILE)) return { seen: {}, updatedAt: null };
  try { const j = JSON.parse(fs.readFileSync(SEEN_FILE, 'utf8')); return { seen: j.seen || {}, updatedAt: j.updatedAt || null }; }
  catch (_) { return { seen: {}, updatedAt: null }; }
}
function saveSeen(state) {
  fs.writeFileSync(SEEN_FILE, JSON.stringify({ seen: state.seen, updatedAt: new Date().toISOString() }, null, 2), 'utf8');
}

// ── 유틸 ───────────────────────────────────────────────────
function ymd(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}
function stripHtml(html) {
  if (!html) return '';
  return String(html)
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/<\/p>/gi, ' ')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')
    .replace(/&amp;/gi, '&').replace(/&quot;/gi, '"').replace(/&#39;/gi, "'")
    .replace(/\s+/g, ' ')
    .trim();
}
function extractRating(a) {
  const cand = a.rating ?? a.star ?? a.point ?? a.review_rating ?? a.rate;
  const n = parseInt(cand, 10);
  return Number.isFinite(n) ? n : 0;
}

(async () => {
  log('');
  log('╔══════════════════════════════════════════════════════════════╗');
  log('║   collect_pending_cafe24.js — 자사몰 리뷰 수집 (모니터링)    ║');
  log('╚══════════════════════════════════════════════════════════════╝');
  log('');

  // 모니터링은 매일(주말·공휴일 포함) 실행 — 나쁜 리뷰는 언제 올라올지 모르므로 영업일 가드 없음.
  cafe24.assertConfigured();

  const now = new Date();
  const since = new Date(now.getTime() - days * 24 * 60 * 60 * 1000);
  const sinceStr = ymd(since);
  const untilStr = ymd(now);

  log(`[수집] 리뷰 게시판(board_no=${cafe24.BOARD_NO}) | 기간 ${sinceStr} ~ ${untilStr}`);
  const articles = await cafe24.listReviewArticles({ since: sinceStr, until: untilStr });
  log(`[수집] 기간 내 리뷰 ${articles.length} 건`);

  const state = loadSeen();

  // ── baseline 시드: 현재 리뷰를 모두 처리됨으로 표시하고 종료 ──
  if (baseline) {
    let added = 0;
    for (const a of articles) {
      const no = a.article_no ?? a.no;
      if (no != null && !state.seen[no]) { state.seen[no] = true; added++; }
    }
    saveSeen(state);
    const outPath = path.isAbsolute(outFile) ? outFile : path.join(__dirname, outFile);
    fs.writeFileSync(outPath, JSON.stringify({ channel: 'cafe24', baseline: true,
      date: `${now.getFullYear()}. ${now.getMonth()+1}. ${now.getDate()}.`, collectedAt: now.toISOString(),
      totalReviews: 0, reviews: [] }, null, 2), 'utf8');
    log(`[baseline] 현재 리뷰 ${added} 건을 '처리됨'으로 시드 → 이후 새 리뷰만 알림 대상.`);
    process.exit(0);
  }

  // ── 미처리 리뷰만 수집 ──
  const collected = [];
  for (const a of articles) {
    const articleNo = a.article_no ?? a.no ?? a.article_id;
    if (articleNo == null) continue;
    if (state.seen[articleNo]) continue; // 이미 처리(알림)함 → 스킵

    const productNo = a.product_no ?? a.product_code ?? null;
    let productName = '';
    try { productName = await cafe24.getProductName(productNo); } catch (_) {}
    if (!productName) productName = a.product_name || (productNo ? `상품 ${productNo}` : '미상');

    collected.push({
      articleNo,
      writer:     a.writer || a.member_id || a.user_name || '익명',
      rating:     extractRating(a),
      date:       a.created_date || a.write_date || a.created_at || '',
      title:      a.title || '',
      reviewText: stripHtml(a.content || a.contents || a.title || ''),
      productNo,
      productName,
    });
  }

  log(`[수집] 미처리 신규 리뷰 ${collected.length} 건`);

  // ── JSON 출력 (네이버 스키마 호환) ──
  const dateStr = `${now.getFullYear()}. ${now.getMonth() + 1}. ${now.getDate()}.`;
  const out = {
    channel:      'cafe24',
    date:         dateStr,
    collectedAt:  now.toISOString(),
    totalReviews: collected.length,
    reviews: collected.map((r, idx) => ({
      no:             idx + 1,
      writer:         r.writer,
      reviewNo:       String(r.articleNo),
      articleNo:      r.articleNo,
      productName:    r.productName,
      productNo:      r.productNo,
      optionName:     '',
      rating:         r.rating,
      date:           r.date,
      reviewText:     r.reviewText,
      reviewPosition: 0,   // 자사몰은 순위 없음
      productUrl:     '',
      posPolicy:      '',
      searchedCount:  0,
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
