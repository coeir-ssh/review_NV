/**
 * collect_status.js — 수집 결과 판정 (.bat 이 분기에 사용).
 * pending_reviews.json 의 상태 + "오늘자 갱신 여부(freshness)"까지 확인해 출력:
 *   SKIP  = 주말/공휴일 (skipped=true)
 *   ZERO  = 오늘 수집됐고 답글미등록 0건
 *   GO    = 오늘 수집됐고 처리할 리뷰 있음
 *   STALE = 파일이 오늘자가 아님 (수집 실패로 어제 파일이 남은 상태) → 진행 금지·재시도 대상
 *   ERR   = 파일 없음/파싱 오류
 * (STALE 을 GO 로 오판해 어제 데이터를 재처리하던 버그 방지)
 */
const fs = require('fs');
const path = require('path');
const p = path.join(__dirname, 'pending_reviews.json');
try {
  const st = fs.statSync(p);
  const now = new Date();
  const m = st.mtime;
  const fresh = m.getFullYear() === now.getFullYear() && m.getMonth() === now.getMonth() && m.getDate() === now.getDate();
  const d = JSON.parse(fs.readFileSync(p, 'utf8'));
  if (d.skipped === true) process.stdout.write('SKIP');
  else if (!fresh)        process.stdout.write('STALE');
  else if (d.totalReviews > 0) process.stdout.write('GO');
  else                    process.stdout.write('ZERO');
} catch (e) {
  process.stdout.write('ERR');
}
