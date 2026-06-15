/**
 * seed_login.js — 셀러센터 로그인을 전용 크롬 프로필(.chrome_profile)에 "심는" 도구.
 *
 * 네이버는 봇 자동 로그인을 차단하므로 로그인은 사람이 해야 한다.
 * 이 스크립트를 한 번 실행해 사람이 로그인하면, 그 로그인이 프로필에 저장되어
 * 이후 collect_pending.js / post_replies.js / 자동 실행이 재로그인 없이 재사용한다.
 * (세션이 정말 만료될 때만 — 보통 수일~수주 뒤 — 다시 한 번 실행하면 됨)
 *
 * 사용 (반드시 보이는 브라우저로):
 *   $env:HEADLESS="false"; node seed_login.js
 *
 * 동작: 브라우저 창 뜸 → 사람이 [로그인하기]로 로그인 → 자동 감지 →
 *       브라우저 "정상 종료"(쿠키를 프로필에 flush) → 완료.
 */
const puppeteer = require('puppeteer');
const { CONFIG, loginToSellerCenter } = require('./post_product_reviews');

(async () => {
  console.log('');
  console.log('╔══════════════════════════════════════════════════════════════╗');
  console.log('║   seed_login.js — 셀러센터 로그인 심기 (전용 프로필)        ║');
  console.log('╚══════════════════════════════════════════════════════════════╝');
  console.log('');
  if (CONFIG.headless) {
    console.log('⚠️  headless 모드입니다. 로그인하려면 보이는 브라우저가 필요합니다.');
    console.log('   PowerShell 에서:  $env:HEADLESS="false"; node seed_login.js');
    process.exit(1);
  }
  const browser = await puppeteer.launch({
    headless:        false,
    protocolTimeout: 120000,
    args:            ['--no-sandbox', '--disable-setuid-sandbox', '--lang=ko-KR,ko', '--window-size=1600,900'],
    defaultViewport: { width: 1600, height: 900 },
  });
  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36');
  try {
    await loginToSellerCenter(page);     // 이미 로그인돼 있으면 즉시 통과, 아니면 사람 로그인 대기
    console.log('');
    console.log('✅ 로그인이 프로필에 저장되었습니다. 이제 자동 실행/수집/등록이 재로그인 없이 동작합니다.');
    // 쿠키가 프로필 디스크에 확실히 기록되도록 잠시 대기 후 "정상 종료"
    await new Promise(r => setTimeout(r, 3000));
    await browser.close();               // ★ 정상 종료 = 프로필에 쿠키 flush (강제종료 금지)
    process.exit(0);
  } catch (e) {
    console.error('[오류]', e.message);
    await browser.close().catch(() => {});
    process.exit(1);
  }
})();
