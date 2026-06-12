/**
 * generate_report_cafe24.js — 자사몰(카페24) 결과로 Word + Slack 별도 리포트 생성
 *
 * posted_results_cafe24.json 을 읽어 네이버와 동일한 보고서 함수를 재사용하되,
 * channelLabel='자사몰(카페24)' 를 주입해 제목/요약에 채널을 표기하고
 * 파일명에 _cafe24 접미사를 붙여 네이버 보고서와 분리한다.
 *
 * 사용: node generate_report_cafe24.js [--in=posted_results_cafe24.json] [--no-slack]
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
const getArg = (k, d) => {
  const a = argv.find(x => x.startsWith(`--${k}=`));
  return a ? a.split('=')[1] : d;
};
const inFile = getArg('in', 'posted_results_cafe24.json');
const skipSlack = argv.includes('--no-slack');

(async () => {
  const inPath = path.isAbsolute(inFile) ? inFile : path.join(__dirname, inFile);
  if (!fs.existsSync(inPath)) {
    console.error(`[오류] 입력 파일 없음: ${inPath}`);
    process.exit(1);
  }
  const summary = JSON.parse(fs.readFileSync(inPath, 'utf8'));
  log(`[입력] ${inPath} (${summary.results?.length || 0} 건)`);

  // 영업일 가드: skip 데이는 보고서·슬랙 모두 생략
  if (summary.skipped === true) {
    log(`[영업일 가드] skipped=true (${summary.skipReason || '사유 미상'}) → Word·Slack 생략. 정상 종료.`);
    process.exit(0);
  }

  // ── 자사몰 채널 메타 주입 ──
  summary.channelLabel = '자사몰(카페24)';
  summary.fileSuffix   = '_cafe24';
  if (!summary.executionEnv) summary.executionEnv = '윈도우 Claude Code 클라이언트 앱';
  if (!summary.aiModel)      summary.aiModel      = process.env.COEIR_AI_MODEL || 'Claude Opus 4.7';
  if (!summary.cost)         summary.cost         = '₩0';

  // .review_summary_cafe24.json 으로도 저장
  const stdSummaryPath = path.join(__dirname, '.review_summary_cafe24.json');
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
        await uploadFileToSlack(filepath, channelId, `📎 ${summary.date} 자사몰(카페24) 보고서`);
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
