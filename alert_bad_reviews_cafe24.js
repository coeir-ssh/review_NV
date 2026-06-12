/**
 * alert_bad_reviews_cafe24.js — 자사몰(카페24) 나쁜 리뷰 슬랙 알림
 *
 * 에이전트가 판별한 screened_reviews_cafe24.json 을 읽어:
 *   - judgeLabel === '삭제대상' 인 리뷰를 슬랙으로 알림 (글번호·별점·본문·근거)
 *   - 점검한 모든 리뷰의 article_no 를 cafe24_seen.json 에 기록 (재알림 방지)
 *
 * ⚠️ 현재는 알림만. 자동 삭제는 하지 않음(추후 cafe24_api.deleteArticle 로 추가 예정).
 *
 * 사용: node alert_bad_reviews_cafe24.js [--in=screened_reviews_cafe24.json] [--no-slack]
 */

const path = require('path');
const fs   = require('fs');
const https = require('https');
const cfg  = require('./config');
const cafe24 = require('./cafe24_api');

const log = msg => console.log(msg);
const SEEN_FILE = path.join(__dirname, 'cafe24_seen.json');

const argv = process.argv.slice(2);
const getArg = (k, d) => { const a = argv.find(x => x.startsWith(`--${k}=`)); return a ? a.split('=')[1] : d; };
const inFile = getArg('in', 'screened_reviews_cafe24.json');
const skipSlack = argv.includes('--no-slack');

function loadSeen() {
  if (!fs.existsSync(SEEN_FILE)) return { seen: {} };
  try { const j = JSON.parse(fs.readFileSync(SEEN_FILE, 'utf8')); return { seen: j.seen || {} }; }
  catch (_) { return { seen: {} }; }
}
function saveSeen(state) {
  fs.writeFileSync(SEEN_FILE, JSON.stringify({ seen: state.seen, updatedAt: new Date().toISOString() }, null, 2), 'utf8');
}

const stars = n => '⭐'.repeat(Math.max(0, Math.min(5, n || 0)));

function sendSlack(text) {
  const token = cfg.SLACK_BOT_TOKEN;
  const channelId = cfg.SLACK_CHANNEL_ID || cfg.SLACK_USER_ID;
  if (!token || !channelId) { log('[슬랙] 토큰/채널 미설정 → 전송 생략'); return Promise.resolve(); }
  return new Promise(resolve => {
    const body = JSON.stringify({ channel: channelId, text });
    const req = https.request(
      { hostname: 'slack.com', path: '/api/chat.postMessage', method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=utf-8', Authorization: `Bearer ${token}` } },
      res => { let d=''; res.on('data',c=>d+=c); res.on('end',()=>{ try{const p=JSON.parse(d); log(p.ok?'[슬랙] 전송 완료':`[슬랙] 오류: ${p.error}`);}catch(e){log('[슬랙] 응답 파싱 오류');} resolve(); }); }
    );
    req.on('error', e => { log(`[슬랙] 요청 오류: ${e.message}`); resolve(); });
    req.write(body); req.end();
  });
}

(async () => {
  const inPath = path.isAbsolute(inFile) ? inFile : path.join(__dirname, inFile);
  if (!fs.existsSync(inPath)) { console.error(`[오류] 입력 파일 없음: ${inPath}`); process.exit(1); }
  const data = JSON.parse(fs.readFileSync(inPath, 'utf8'));

  if (data.skipped === true) { log(`[skip] skipped=true (${data.skipReason||''}) → 알림 생략.`); process.exit(0); }

  const results = data.results || data.reviews || [];
  const bad = results.filter(r => r.judgeLabel === '삭제대상');
  const dateStr = (data.date || '').replace(/\s/g, '').replace(/\.$/, '');

  log(`[입력] ${inPath} | 점검 ${results.length}건 / 나쁜 리뷰 ${bad.length}건`);

  // ── 삭제대상 각 건의 상품 내 노출 순위(최신순) 계산 ──
  for (const r of bad) {
    try {
      const pos = await cafe24.getProductReviewPosition(r.productNo ?? r.product_no, r.articleNo ?? r.reviewNo);
      r._rankPos = pos.position;
      r._rankTotal = pos.total;
    } catch (e) {
      log(`  [경고] 순위 계산 실패(article ${r.articleNo}): ${e.message}`);
      r._rankPos = 0; r._rankTotal = 0;
    }
  }
  const fmtRank = r => r._rankPos > 0
    ? `최신순 ${r._rankPos}번째${r._rankTotal ? ` (전체 ${r._rankTotal}개 중)` : ''}`
    : '순위 미확인';

  // ── 슬랙 메시지 작성 (삭제 검토 대상이 있을 때만 전송) ──
  if (!skipSlack) {
    if (bad.length === 0) {
      log('[알림] 삭제 검토 대상 0건 → 전송 생략(조용히 종료).');
    } else {
      const lines = [
        `🚨 자사몰(카페24) 삭제 검토 대상 리뷰 ${bad.length}건 (${dateStr})`,
        `(자동 삭제 안 함 — 아래 글번호로 카페24 관리자 > 게시판 > 리뷰 에서 확인 후 삭제)`,
        ``,
      ];
      bad.forEach((r, i) => {
        lines.push(`─────────────────────────────`);
        lines.push(`No.${i + 1}  ${stars(r.rating)} (${r.rating}점)  |  ${r.writer}`);
        lines.push(`글번호 : ${r.reviewNo || r.articleNo || '-'}`);
        lines.push(`상품 : ${r.productName || '-'}`);
        lines.push(`📍 리뷰 순위 : ${fmtRank(r)}`);
        lines.push(`작성일 : ${(r.date || '').toString().slice(0, 16).replace('T', ' ')}`);
        lines.push(`리뷰 : "${(r.reviewText || '').replace(/\n/g, ' ')}"`);
        lines.push(`🔴 판단 : ${r.judgeLabel}${r.judgeReason ? ` — ${r.judgeReason}` : ''}${r.judgeConfidence != null ? ` (confidence ${r.judgeConfidence})` : ''}`);
        lines.push(``);
      });
      await sendSlack(lines.join('\n'));
    }
  } else {
    log('[슬랙] --no-slack 옵션 — 전송 생략');
  }

  // ── 점검한 모든 리뷰를 seen 에 기록 (재알림 방지) ──
  const state = loadSeen();
  let marked = 0;
  for (const r of results) {
    const no = r.articleNo ?? r.reviewNo;
    if (no != null && !state.seen[no]) { state.seen[no] = true; marked++; }
  }
  saveSeen(state);
  log(`[seen] ${marked}건 처리 기록 (cafe24_seen.json)`);
  log('✅ 완료');
})().catch(err => {
  console.error('\n[오류]', err.message);
  console.error(err.stack);
  process.exit(1);
});
