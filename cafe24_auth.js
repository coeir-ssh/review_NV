/**
 * cafe24_auth.js — 카페24 Open API 최초 OAuth 인증 (1회 실행)
 *
 * [순서]
 *  1. 아래 출력되는 인증 URL 을 브라우저에서 연다 (카페24 관리자 로그인 필요)
 *  2. 권한 동의 → Redirect URI(예: https://localhost)로 이동하며 주소창에 ?code=... 가 붙음
 *  3. 그 주소창 전체 URL(또는 code 값)을 복사해 이 콘솔에 붙여넣기
 *  4. access_token / refresh_token 이 .cafe24_token.json 에 저장됨
 *
 * 토큰 만료(access 2시간 / refresh 2주)는 cafe24_api.js 가 자동 갱신.
 * refresh 마저 만료되면(2주 이상 미사용) 이 스크립트를 다시 실행.
 *
 * 사용: node cafe24_auth.js
 */

const readline = require('readline');
const {
  getAuthUrl, exchangeCodeForToken, assertConfigured, MALL_ID, REDIRECT_URI, SCOPES,
} = require('./cafe24_api');

function ask(prompt) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise(resolve => rl.question(prompt, ans => { rl.close(); resolve(ans.trim()); }));
}

// 붙여넣은 입력(전체 URL 또는 code)에서 code 추출
function extractCode(input) {
  if (!input) return '';
  // 전체 URL 이면 code 파라미터 파싱
  const m = input.match(/[?&]code=([^&\s]+)/);
  if (m) return decodeURIComponent(m[1]);
  return input; // 이미 code 값만 붙여넣은 경우
}

(async () => {
  try {
    assertConfigured();
  } catch (e) {
    console.error(`\n[설정 오류] ${e.message}`);
    console.error('config.js 의 CAFE24_MALL_ID / CAFE24_CLIENT_ID / CAFE24_CLIENT_SECRET / CAFE24_REDIRECT_URI 를 먼저 채우세요.\n');
    process.exit(1);
  }

  const url = getAuthUrl();
  console.log('');
  console.log('╔══════════════════════════════════════════════════════════════╗');
  console.log('║   카페24 Open API 최초 인증                                   ║');
  console.log('╚══════════════════════════════════════════════════════════════╝');
  console.log('');
  console.log(`몰 ID      : ${MALL_ID}`);
  console.log(`Redirect   : ${REDIRECT_URI}`);
  console.log(`요청 scope : ${SCOPES}`);
  console.log('');
  console.log('① 아래 URL 을 브라우저에 붙여넣어 열고, 카페24 관리자로 로그인 후 권한에 동의하세요:');
  console.log('');
  console.log(url);
  console.log('');
  console.log(`② 동의하면 ${REDIRECT_URI} 로 이동하며 주소창에 ?code=... 가 붙습니다.`);
  console.log('   (페이지가 안 열려도 OK — 주소창의 전체 URL 만 복사하면 됩니다)');
  console.log('');

  const pasted = await ask('③ 리다이렉트된 전체 URL(또는 code 값)을 여기에 붙여넣고 Enter: ');
  const code = extractCode(pasted);
  if (!code) {
    console.error('\n[오류] code 를 찾지 못했습니다. 전체 URL 또는 code 값을 정확히 붙여넣으세요.\n');
    process.exit(1);
  }

  console.log('\n[토큰 발급] 중...');
  const rec = await exchangeCodeForToken(code);
  console.log('');
  console.log('✅ 인증 완료 — .cafe24_token.json 저장됨');
  console.log(`   access_token  만료: ${new Date(rec.expires_at).toLocaleString('ko-KR')}`);
  console.log(`   refresh_token 만료: ${new Date(rec.refresh_expires_at).toLocaleString('ko-KR')}`);
  console.log('');
  console.log('다음: node collect_pending_cafe24.js 로 리뷰 수집을 테스트하세요.');
})().catch(err => {
  console.error('\n[오류]', err.message);
  process.exit(1);
});
