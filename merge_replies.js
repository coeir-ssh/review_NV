/**
 * merge_replies.js — 에이전트의 최소 판단 출력(judgements.json)을
 *   pending_reviews.json 의 메타데이터와 병합해 완전한 replies.json 생성.
 *
 * 목적: 에이전트가 리뷰 원문·제품명 등 큰 필드를 다시 출력하지 않게 해
 *   출력 토큰량을 절반 이하로 줄임(대량 리뷰일 때 32K/64K 한도 초과 방지).
 *
 * judgements.json 형식 (에이전트가 작성):
 *   { "date":"...", "results":[ { "no":1, "judgeLabel":"답변|환불검토",
 *       "judgeReason":"...", "judgeConfidence":92, "replyText":"...",
 *       "refundPolicy":"" } , ... ] }
 *
 * 출력: replies.json (post_replies.js 가 기대하는 전체 필드 형식)
 * 사용: node merge_replies.js
 */
const fs = require('fs');
const path = require('path');

const pendingPath = path.join(__dirname, 'pending_reviews.json');
const judgePath   = path.join(__dirname, 'judgements.json');
const outPath     = path.join(__dirname, 'replies.json');

if (!fs.existsSync(judgePath)) {
  console.error('[merge] judgements.json 없음 → 병합 중단 (에이전트 판단 실패 추정). replies.json 생성 안 함.');
  process.exit(1);
}
const pending = JSON.parse(fs.readFileSync(pendingPath, 'utf8'));
const J = JSON.parse(fs.readFileSync(judgePath, 'utf8'));

const byNo = new Map();
pending.reviews.forEach((r, i) => byNo.set(r.no != null ? r.no : (i + 1), r));

const results = (J.results || []).map(j => {
  const r = byNo.get(j.no);
  if (!r) { console.error('[merge] pending 에 없는 no=' + j.no + ' → 건너뜀'); return null; }
  return {
    no: j.no, writer: r.writer, reviewNo: r.reviewNo, productName: r.productName,
    optionName: r.optionName || '', rating: r.rating, date: r.date,
    reviewText: r.reviewText, reviewPosition: r.reviewPosition, productUrl: r.productUrl || '',
    judgeLabel: j.judgeLabel, judgeReason: j.judgeReason || '',
    judgeConfidence: j.judgeConfidence != null ? j.judgeConfidence : null,
    replyText: j.replyText || '',
    refundPolicy: j.judgeLabel === '환불검토' ? (j.refundPolicy || r.posPolicy || '순위 미확인') : '',
  };
}).filter(Boolean);

// 판단 누락(수집됐는데 에이전트가 빠뜨린 리뷰) 경고
const judged = new Set(results.map(x => x.no));
const missing = pending.reviews.map((r, i) => r.no != null ? r.no : (i + 1)).filter(n => !judged.has(n));
if (missing.length) console.error('[merge] ⚠️ 판단 누락 no: ' + missing.join(', '));

fs.writeFileSync(outPath, JSON.stringify({ date: pending.date, results }, null, 2), 'utf8');
console.log('[merge] replies.json 생성: ' + results.length + '건 (답변 '
  + results.filter(x => x.judgeLabel === '답변').length + ' / 환불검토 '
  + results.filter(x => x.judgeLabel === '환불검토').length + ' / 판단누락 ' + missing.length + ')');
