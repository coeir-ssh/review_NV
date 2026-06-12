/**
 * post_run_check.js — 데일리 루틴 실행 후 "완주 여부" 감시 (워치독)
 *
 * 스케줄러가 claude --print 에이전트 실행을 끝낸 직후 이 스크립트를 돌린다.
 * 에이전트가 중간에 죽어도(헤드리스 종료 등) 죽은 에이전트는 스스로 실패 보고를
 * 못 하므로, 에이전트 "바깥"에서 오늘자 산출물이 생겼는지 객관적으로 검사한다.
 *
 * 판정:
 *   - pending_reviews.json 이 오늘자이고 skipped=true        → 영업일 아님. 정상(조용히 종료).
 *   - pending_reviews.json 이 오늘자이고 totalReviews=0       → 처리할 리뷰 없음. 정상.
 *   - 오늘자 posted_results.json 이 존재 + 오늘자 보고서 docx → 정상 완주.
 *   - 그 외 (산출물 누락/과거 날짜)                          → 실패로 간주 → Slack 알림.
 *
 * 사용: node post_run_check.js
 */

const path = require('path');
const fs   = require('fs');
const { sendSlackError, sendSlackText, sessionAgeDays } = require('./post_product_reviews');

const DIR = __dirname;

// 세션이 만료에 가까우면(=마지막 전체 로그인 후 N일 경과) 사전 리마인드.
// 네이버 셀러 세션은 보통 1~2주마다 만료 + 추가 인증을 요구하므로,
// 7시 자동 실행이 깜짝 실패하기 전에 미리 한가할 때 수동 로그인하도록 안내한다.
const SESSION_WARN_DAYS = 6;   // 이 일수 넘으면 경고 시작
async function maybeWarnSessionAge() {
  try {
    const age = sessionAgeDays();
    if (age != null && age >= SESSION_WARN_DAYS) {
      await sendSlackText(
        `🔑 *코에르 셀러 세션 갱신 권장*\n` +
        `• 마지막 로그인 후 약 ${Math.floor(age)}일 경과 — 곧 세션이 만료되어 자동 실행이 멈출 수 있어요.\n` +
        `• 한가하실 때 터미널에서 \`node collect_pending.js\` 한 번 실행해 로그인(세션 갱신)해 주세요.\n` +
        `• 미리 갱신해두면 아침 7시 자동 실행이 세션 만료로 실패하는 일을 예방할 수 있습니다.`
      );
      console.log(`[워치독] 세션 ${Math.floor(age)}일 경과 → 갱신 권장 알림 전송`);
    }
  } catch (e) { console.log('[워치독] 세션 점검 실패:', e.message); }
}

function todayStr() {
  const n = new Date();
  return `${n.getFullYear()}${String(n.getMonth() + 1).padStart(2, '0')}${String(n.getDate()).padStart(2, '0')}`;
}

// 파일 mtime 이 오늘인지
function isModifiedToday(file) {
  try {
    const st = fs.statSync(file);
    const m = st.mtime;
    const n = new Date();
    return m.getFullYear() === n.getFullYear() && m.getMonth() === n.getMonth() && m.getDate() === n.getDate();
  } catch { return false; }
}

function readJsonSafe(file) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return null; }
}

(async () => {
  const ymd = todayStr();
  const pendingPath = path.join(DIR, 'pending_reviews.json');
  const postedPath  = path.join(DIR, 'posted_results.json');
  const reportPath  = path.join(DIR, 'reply', `${ymd}_reply.docx`);

  const pending = readJsonSafe(pendingPath);
  const pendingFresh = isModifiedToday(pendingPath);

  // 1) 수집 자체가 오늘 안 됨 → 가장 앞단계 실패
  if (!pending || !pendingFresh) {
    await sendSlackError(
      `데일리 리뷰 루틴 실패 (${ymd})\n` +
      `• 증상: 오늘자 pending_reviews.json 이 없거나 갱신 안 됨 (Step 1 수집 실패 추정)\n` +
      `• 조치: 셀러센터 로그인/스토어 전환/그리드 렌더 확인 필요\n` +
      `• 로그: logs\\agent_run_${ymd}_*.log`
    );
    console.log('[워치독] 실패 알림 전송: 수집 미완료');
    process.exit(0);
  }

  // 2) 영업일 아님(주말/공휴일) → 정상. 조용히 종료
  if (pending.skipped === true) {
    console.log(`[워치독] 정상: 영업일 아님 (${pending.skipReason || ''}) — 알림 없음`);
    process.exit(0);
  }

  // 3) 처리할 리뷰 0건 → 정상
  if (!pending.totalReviews || pending.totalReviews === 0) {
    console.log('[워치독] 정상: 답글미등록 0건 — 알림 없음');
    await maybeWarnSessionAge();
    process.exit(0);
  }

  // 4) 리뷰가 있는데 등록 결과/보고서가 오늘자가 아니면 실패
  const postedOk = readJsonSafe(postedPath) && isModifiedToday(postedPath);
  const reportOk = fs.existsSync(reportPath);

  if (!postedOk || !reportOk) {
    const miss = [];
    if (!postedOk) miss.push('posted_results.json(등록결과) 누락/과거');
    if (!reportOk) miss.push(`${ymd}_reply.docx(보고서) 없음`);
    await sendSlackError(
      `데일리 리뷰 루틴 미완주 (${ymd})\n` +
      `• 수집: ${pending.totalReviews}건 정상\n` +
      `• 누락: ${miss.join(' / ')}\n` +
      `• 추정 원인: Step 3(등록) 또는 Step 4(보고서) 미실행 — 에이전트가 중간 종료됐을 가능성\n` +
      `• 로그: logs\\agent_run_${ymd}_*.log`
    );
    console.log('[워치독] 실패 알림 전송: 등록/보고서 미완료');
    process.exit(0);
  }

  console.log('[워치독] 정상 완주 확인 — 알림 없음');
  await maybeWarnSessionAge();
  process.exit(0);
})().catch(async (e) => {
  // 워치독 자체 오류도 알림 (조용히 죽지 않게)
  try {
    await sendSlackError(`데일리 리뷰 워치독(post_run_check.js) 자체 오류: ${e.message}`);
  } catch {}
  console.error('[워치독 오류]', e.message);
  process.exit(0);
});
