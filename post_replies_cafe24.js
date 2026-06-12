/**
 * post_replies_cafe24.js — 에이전트가 생성한 답변(replies_cafe24.json)을 카페24에 등록
 *
 * AI 호출 없음. 입력 JSON 의 각 항목을 순회하며:
 *   - judgeLabel === '환불검토' → 답글 등록 안 함 (보고서에만 포함)
 *   - judgeLabel === '답변' + replyText 있음 → 카페24 게시글에 댓글(답글) 등록
 *
 * 결과는 posted_results_cafe24.json 으로 저장 (네이버 posted_results.json 과 동일 구조).
 *
 * 옵션:
 *   --in=replies_cafe24.json        입력 파일
 *   --out=posted_results_cafe24.json 출력 파일
 *   --limit=N                       앞에서 N 건만 등록 (드라이런/단계 검증용)
 *
 * 사용: node post_replies_cafe24.js [--in=...] [--out=...] [--limit=1]
 */

const path = require('path');
const fs   = require('fs');
const cafe24 = require('./cafe24_api');

const log   = msg => console.log(msg);
const sleep = ms  => new Promise(r => setTimeout(r, ms));

const argv = process.argv.slice(2);
const getArg = (k, d) => {
  const a = argv.find(x => x.startsWith(`--${k}=`));
  return a ? a.split('=')[1] : d;
};
const inFile  = getArg('in', 'replies_cafe24.json');
const outFile = getArg('out', 'posted_results_cafe24.json');
const limit   = parseInt(getArg('limit', '0'), 10) || 0;  // 0 = 전체

(async () => {
  log('');
  log('╔══════════════════════════════════════════════════════════════╗');
  log('║   post_replies_cafe24.js — 자사몰 답변 등록                   ║');
  log('╚══════════════════════════════════════════════════════════════╝');
  log('');

  const inPath = path.isAbsolute(inFile) ? inFile : path.join(__dirname, inFile);
  if (!fs.existsSync(inPath)) {
    console.error(`[오류] 입력 파일 없음: ${inPath}`);
    process.exit(1);
  }
  const input = JSON.parse(fs.readFileSync(inPath, 'utf8'));

  // skip 마커가 전파된 경우 등록 없이 그대로 통과
  if (input.skipped === true) {
    const outPath = path.isAbsolute(outFile) ? outFile : path.join(__dirname, outFile);
    fs.writeFileSync(outPath, JSON.stringify({ ...input, postedAt: new Date().toISOString(), replied: 0, failed: 0, refund: 0, results: [] }, null, 2), 'utf8');
    log('[skip] skipped=true → 등록 생략, 결과 파일만 저장.');
    process.exit(0);
  }

  const replies = input.results || input.reviews || [];
  log(`[입력] ${inPath}`);
  log(`[입력] 총 ${replies.length} 건 (답변=${replies.filter(r => r.judgeLabel === '답변').length}, 환불검토=${replies.filter(r => r.judgeLabel === '환불검토').length})`);

  cafe24.assertConfigured();

  let successCount = 0;
  let failCount    = 0;
  let refundCount  = 0;
  const results    = [];

  // ── 환불검토 항목: 등록 없이 결과만 누적 ──
  for (const r of replies) {
    if (r.judgeLabel === '환불검토') {
      refundCount++;
      results.push({ ...r, refundCheck: '검토필요', posted: false });
    }
  }

  // ── 답변 항목 등록 ──
  let answerList = replies.filter(r => r.judgeLabel === '답변' && r.replyText);
  if (limit > 0) {
    log(`[제한] --limit=${limit} → 앞 ${limit} 건만 등록`);
    answerList = answerList.slice(0, limit);
  }
  log('');
  log(`[등록] 답변 ${answerList.length} 건 등록 시작...`);

  for (const r of answerList) {
    const articleNo = r.articleNo ?? r.reviewNo;
    log(`──────────────────────────────────────`);
    log(`[등록] article ${articleNo} | ${r.writer} | ${(r.productName || '').substring(0, 24)}`);
    log(`  답변: ${r.replyText.substring(0, 60)}...`);
    try {
      await cafe24.postComment(articleNo, r.replyText);
      successCount++;
      results.push({ ...r, refundCheck: '-', posted: true });
      log(`  ✓ 등록 완료 (${successCount}/${answerList.length})`);
    } catch (e) {
      failCount++;
      results.push({ ...r, refundCheck: '-', posted: false, failReason: e.message });
      log(`  ✗ 실패: ${e.message}`);
    }
    await sleep(400); // API 레이트 보호
  }

  log('');
  log('══════════════════════════════════════════');
  log(`등록 성공: ${successCount} / 실패: ${failCount} / 환불검토: ${refundCount}`);
  log('══════════════════════════════════════════');

  const out = {
    ...input,
    channel:  'cafe24',
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
