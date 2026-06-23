/**
 * generate_report.js — posted_results.json 으로 Word + Slack DM + 파일 업로드 수행
 *
 * 사용:
 *   node generate_report.js [--in posted_results.json] [--no-slack]
 */

const path = require('path');
const fs   = require('fs');
const cfg  = require('./config');

const {
  generateWordDoc,
  sendSlackDM,
  uploadFileToSlack,
} = require('./post_product_reviews');

const log = msg => console.log(msg);

const argv = process.argv.slice(2);
const inArg = argv.find(a => a.startsWith('--in='));
const inFile = inArg ? inArg.split('=')[1] : 'posted_results.json';
const skipSlack = argv.includes('--no-slack');

(async () => {
  const inPath = path.isAbsolute(inFile) ? inFile : path.join(__dirname, inFile);
  if (!fs.existsSync(inPath)) {
    console.error(`[오류] 입력 파일 없음: ${inPath}`);
    process.exit(1);
  }
  const summary = JSON.parse(fs.readFileSync(inPath, 'utf8'));
  // generateWordDoc / sendSlackDM 는 summary.results 를 기대
  // posted_results.json 의 results 가 이미 호환 형식임 (writer/productName/judgeLabel/...)
  log(`[입력] ${inPath} (${summary.results?.length || 0} 건)`);

  // ── 영업일 가드: 주말/공휴일 등 스킵 데이는 보고서·슬랙 모두 전송 안 함 ──
  // pending_reviews.json 에서 skipped:true 가 posted_results.json 으로 전파되거나,
  // 에이전트가 명시적으로 skipped 마커를 둔 경우 모두 처리.
  if (summary.skipped === true) {
    log(`[영업일 가드] skipped=true (${summary.skipReason || '사유 미상'}) → Word·Slack 전송 생략. 정상 종료.`);
    process.exit(0);
  }

  // ── 에이전트 매개 실행 디폴트 메타 주입 (값 없으면 채움) ──
  // 환경변수로 모델명 오버라이드 가능: COEIR_AI_MODEL
  if (!summary.executionEnv) summary.executionEnv = '윈도우 Claude Code 클라이언트 앱';
  if (!summary.aiModel)      summary.aiModel      = process.env.COEIR_AI_MODEL || 'Claude Opus 4.7';
  if (!summary.cost)         summary.cost         = '₩0';
  // 에이전트 매개 모드에선 API 토큰/비용 추적이 없으므로 summary.usage 는 미설정.
  // 위 3개 디폴트 라인이 작업 요약에 표시된다.

  // ── 전날 판단 검증 요약 주입 (환불검토 섹션 앞에 표시) ──
  // 우선순위: 에이전트가 분석해 적은 verification_summary.txt > 자동 생성(verification_result.json)
  if (!summary.verificationSummary) {
    const txtPath = path.join(__dirname, 'verification_summary.txt');
    const resPath = path.join(__dirname, 'verification_result.json');
    if (fs.existsSync(txtPath)) {
      const t = fs.readFileSync(txtPath, 'utf8').trim();
      if (t) summary.verificationSummary = t;
    } else if (fs.existsSync(resPath)) {
      try {
        const vr = JSON.parse(fs.readFileSync(resPath, 'utf8'));
        if (vr.results && vr.results.length) {
          const c = vr.counts || {};
          const short = s => (s || '').replace(/^코에르\s*/, '').slice(0, 18);
          const refunds = vr.results.filter(r => r.judgeLabel === '환불검토');
          const lines = [
            `(전 영업일 처리분이 실제로 어떻게 됐는지 셀러센터에서 확인한 결과)`,
          ];
          // 2일 이상 미처리 대기중 건 → 최우선 체크 경고 (맨 위에)
          const overdue = refunds.filter(r => r.verdict === '대기중' && (r.daysPending || 0) >= 2);
          if (overdue.length > 0) {
            lines.push(`🚨 환불검토 ${overdue.length}건이 2일 이상 처리되지 않고 있습니다 — 최우선 체크 요청!`);
            overdue.forEach(r => {
              lines.push(`🚨 ${r.reviewNo} ${short(r.productName)} — ${r.daysPending}일째 미처리`);
            });
          }
          // 환불검토 검증 — 대상이 있으면 카운트+상세, 없으면 명시
          if (refunds.length > 0) {
            lines.push(`· 환불검토 적중 ${c.refundHit || 0}건 (실제 블라인드 처리됨)`);
            lines.push(`· 환불검토 빗나감 ${c.refundMiss || 0}건 (담당자가 답변으로 처리 → 보수적 판단)`);
            if (c.refundPending || 0) lines.push(`· 환불검토 대기중 ${c.refundPending}건 (담당자 처리 전 — 익일 재확인)`);
            refunds.forEach(r => {
              const mark = r.verdict === '적중' ? '✅ 적중'
                         : r.verdict === '빗나감' ? '❌ 빗나감'
                         : r.verdict === '대기중' ? '⏳ 대기중'
                         : `· ${r.verdict}`;
              const tail = r.verdict === '대기중' ? `담당자 처리 전${r.daysPending != null ? ` (${r.daysPending}일째)` : ''}` : `실제 전시상태 '${r.actualStatus}'`;
              lines.push(`[환불검토 ${mark}] ${r.reviewNo} ${short(r.productName)} — ${tail}`);
            });
          } else {
            lines.push(`· 전일 환불검토 항목 없었음`);
          }
          // 답변 검증
          lines.push(`· 답변 정상등록 ${c.ansOk || 0}건 (답변으로 본 건이 실제 답글 정상 등록)`);
          if (c.ansMiss || 0) lines.push(`· 답변 빗나감 ${c.ansMiss}건 (답변으로 봤으나 실제 환불·블라인드)`);
          if (c.unknown || 0)  lines.push(`· 확인불가 ${c.unknown}건`);
          // 답변 빗나간 건은 구체적으로
          vr.results.filter(r => r.judgeLabel === '답변' && (r.verdict || '').startsWith('빗나감')).forEach(r => {
            lines.push(`[답변 ❌ 빗나감] ${r.reviewNo} ${short(r.productName)} — 실제 '${r.actualStatus}' (환불됐어야)`);
          });
          summary.verificationSummary = lines.join('\n');
        }
      } catch {}
    }
    // 검증 결과가 전혀 없을 때도 영역은 항상 표시 (사용자 요청)
    if (!summary.verificationSummary) {
      summary.verificationSummary = '전일 환불검토 항목 없었음 (검증 대상 없음)';
    }
  }

  // .review_summary.json 으로도 저장 (resend_with_rank.js 등과 호환)
  const stdSummaryPath = path.join(__dirname, '.review_summary.json');
  fs.writeFileSync(stdSummaryPath, JSON.stringify(summary, null, 2), 'utf8');

  log('[Word] 생성 중...');
  const { filename, filepath } = await generateWordDoc(summary);
  log(`  → ${filename}`);

  if (!skipSlack) {
    log('[Slack] DM 전송 중...');
    try {
      await sendSlackDM(summary);
      const channelId = cfg.SLACK_CHANNEL_ID || cfg.SLACK_USER_ID;
      if (channelId) {
        log('[Slack] Word 파일 업로드 중...');
        await uploadFileToSlack(filepath, channelId, `📎 ${summary.date} 보고서 워드 파일`);
      }
    } catch (e) {
      log(`[Slack 오류] ${e.message}`);
    }
  } else {
    log('[Slack] --no-slack 옵션 — 전송 생략');
  }

  log('');
  log('✅ 완료');
})().catch(err => {
  console.error('\n[오류]', err.message);
  console.error(err.stack);
  process.exit(1);
});
