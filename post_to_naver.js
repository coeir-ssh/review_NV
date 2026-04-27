/**
 * 코에르 리뷰 답변 네이버 스마트스토어 자동 포스팅
 *
 * 1. 최신 리뷰답변_자동_*.xlsx 파일 읽기
 * 2. 네이버 커머스 API로 액세스 토큰 발급
 * 3. reviewId별로 답변 POST
 * 4. 포스팅 완료 여부를 엑셀에 업데이트
 *
 * 실행: node post_to_naver.js
 */

const crypto = require('crypto');
const https = require('https');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const cfg = require('./config');

// ─────────────────────────────────────────────────────────
// 설정
// ─────────────────────────────────────────────────────────
const CONFIG = {
  outputDir: 'C:\\Users\\AWESOMATIC\\Desktop\\코에르\\클로드\\리뷰수집프로그램',
  clientId:     cfg.NAVER_CLIENT_ID,
  clientSecret: cfg.NAVER_CLIENT_SECRET,
  tokenUrl:     'https://api.commerce.naver.com/external/v1/oauth2/token',
  replyBaseUrl: 'https://api.commerce.naver.com/external/v1/reviews',
  delayMs: 1000,  // 답변 간격 (1초) - 너무 빠르면 rate limit
};

const log = msg => console.log(msg);
const sleep = ms => new Promise(r => setTimeout(r, ms));

// ─────────────────────────────────────────────────────────
// 유틸
// ─────────────────────────────────────────────────────────
function timestamp() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}

// 최신 리뷰답변 파일 찾기
function findLatestResponseExcel() {
  const files = fs.readdirSync(CONFIG.outputDir)
    .filter(f => f.startsWith('리뷰답변_자동_') && f.endsWith('.xlsx'))
    .sort()
    .reverse();
  return files.length ? path.join(CONFIG.outputDir, files[0]) : null;
}

// ─────────────────────────────────────────────────────────
// 네이버 커머스 API 인증
// ─────────────────────────────────────────────────────────
function makeSignature(clientId, clientSecret, timestamp) {
  const message = `${clientId}_${timestamp}`;
  return crypto.createHmac('sha256', clientSecret)
    .update(message)
    .digest('base64');
}

async function getAccessToken() {
  const ts = Date.now();
  const sign = makeSignature(CONFIG.clientId, CONFIG.clientSecret, ts);

  const body = [
    'grant_type=client_credentials',
    `client_id=${encodeURIComponent(CONFIG.clientId)}`,
    `timestamp=${ts}`,
    `client_secret_sign=${encodeURIComponent(sign)}`,
    'type=SELF',
  ].join('&');

  return new Promise((resolve, reject) => {
    const url = new URL(CONFIG.tokenUrl);
    const options = {
      hostname: url.hostname,
      path: url.pathname,
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body),
      },
    };

    const req = https.request(options, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (parsed.access_token) {
            resolve(parsed.access_token);
          } else {
            reject(new Error(`토큰 발급 실패: ${data}`));
          }
        } catch (e) {
          reject(new Error(`토큰 응답 파싱 실패: ${data}`));
        }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// ─────────────────────────────────────────────────────────
// 리뷰 답변 POST
// ─────────────────────────────────────────────────────────
async function postReply(accessToken, reviewId, replyText) {
  const body = JSON.stringify({ content: replyText });
  const url = new URL(`${CONFIG.replyBaseUrl}/${reviewId}/reply`);

  return new Promise((resolve, reject) => {
    const options = {
      hostname: url.hostname,
      path: url.pathname,
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(body),
      },
    };

    const req = https.request(options, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        resolve({ status: res.statusCode, body: data });
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

// ─────────────────────────────────────────────────────────
// 메인
// ─────────────────────────────────────────────────────────
async function main() {
  log('');
  log('╔══════════════════════════════════════════════════════════════╗');
  log('║   코에르 리뷰 답변 네이버 자동 포스팅                       ║');
  log('║   (스마트스토어 커머스 API)                                  ║');
  log('╚══════════════════════════════════════════════════════════════╝');
  log('');

  // ── 엑셀 파일 찾기 ───────────────────────────────────
  const srcFile = findLatestResponseExcel();
  if (!srcFile) {
    log('[오류] 리뷰답변_자동_*.xlsx 파일이 없습니다.');
    log('  → 먼저 node auto_review_response.js 를 실행하세요.');
    return;
  }
  log(`[파일] ${path.basename(srcFile)}`);

  // ── 엑셀 읽기 ────────────────────────────────────────
  const wb = XLSX.readFile(srcFile);
  const ws = wb.Sheets['리뷰답변'];
  if (!ws) { log('[오류] "리뷰답변" 시트를 찾을 수 없습니다.'); return; }

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const header = rows[0]; // ['reviewId', '제품명', '별점', ...]
  const dataRows = rows.slice(1).filter(r => r[0] || r[1]);

  // 컬럼 인덱스 파악 (헤더 기반으로 유연하게)
  const COL = {
    reviewId:  header.indexOf('reviewId'),
    product:   header.indexOf('제품명'),
    response:  header.indexOf('답변'),
    posted:    header.indexOf('포스팅완료'),
  };

  log(`[리뷰] 총 ${dataRows.length}개 행`);

  // 포스팅 대상: 답변 있고 포스팅 미완료이며 reviewId 있는 것
  const targets = dataRows
    .map((row, i) => ({ row, rowIdx: i + 1 }))  // rowIdx: 헤더 포함 실제 행 번호
    .filter(({ row }) => {
      const reviewId = String(row[COL.reviewId] || '').trim();
      const response = String(row[COL.response] || '').trim();
      const posted   = String(row[COL.posted]   || '').trim();
      return reviewId && response && !posted;
    });

  log(`[대상] 포스팅 필요: ${targets.length}개`);

  if (targets.length === 0) {
    log('  → 포스팅할 항목이 없습니다. (이미 완료되었거나 reviewId/답변 없음)');
    return;
  }

  // ── 액세스 토큰 발급 ─────────────────────────────────
  log('');
  log('[인증] 네이버 커머스 API 액세스 토큰 발급 중...');
  let accessToken;
  try {
    accessToken = await getAccessToken();
    log('  ✓ 토큰 발급 성공');
  } catch (err) {
    log(`  ✗ 토큰 발급 실패: ${err.message}`);
    return;
  }

  // ── 답변 포스팅 ───────────────────────────────────────
  log('');
  log('[포스팅] 답변 등록 시작...');
  log('');

  let successCount = 0;
  let failCount = 0;

  for (const { row, rowIdx } of targets) {
    const reviewId  = String(row[COL.reviewId]).trim();
    const product   = String(row[COL.product] || '').substring(0, 20);
    const response  = String(row[COL.response]).trim();

    process.stdout.write(`  [${reviewId}] ${product}... `);

    try {
      const result = await postReply(accessToken, reviewId, response);

      if (result.status === 200 || result.status === 201 || result.status === 204) {
        // 성공 → 엑셀 해당 행 포스팅완료 체크
        const cellAddr = XLSX.utils.encode_cell({ r: rowIdx, c: COL.posted });
        ws[cellAddr] = { v: '✓', t: 's' };
        successCount++;
        log(`✓ (${result.status})`);
      } else {
        failCount++;
        log(`✗ (${result.status}) ${result.body.substring(0, 100)}`);
      }
    } catch (err) {
      failCount++;
      log(`✗ 오류: ${err.message}`);
    }

    await sleep(CONFIG.delayMs);
  }

  // ── 엑셀 업데이트 저장 ───────────────────────────────
  log('');
  log('[저장] 포스팅 결과 엑셀 업데이트...');
  XLSX.writeFile(wb, srcFile);
  log(`  ✓ 업데이트 완료: ${path.basename(srcFile)}`);

  // ── 최종 통계 ─────────────────────────────────────────
  log('');
  log('─'.repeat(60));
  log(`  성공: ${successCount}개`);
  log(`  실패: ${failCount}개`);
  log('─'.repeat(60));
  log('');
  log('╔══════════════════════════════════════════════════════════════╗');
  log('║   완료! 엑셀에서 포스팅완료 컬럼을 확인하세요.              ║');
  log('╚══════════════════════════════════════════════════════════════╝');
}

main().catch(err => {
  console.error('\n[오류]', err.message);
  process.exit(1);
});
