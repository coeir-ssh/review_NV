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
