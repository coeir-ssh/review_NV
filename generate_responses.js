/**
 * 코에르 리뷰 AI 처리 프로그램
 *
 * 1. 최신 수집 엑셀에서 리뷰 읽기
 * 2. Claude AI로 환불 여부 판단 + 답변 생성
 * 3. 새 엑셀 저장 (전체리뷰 / 환불검토 / 요약 시트)
 *
 * 실행 전: config.js에 ANTHROPIC_API_KEY 입력
 */

const Anthropic = require('@anthropic-ai/sdk');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const cfg = require('./config');

// ─────────────────────────────────────────────────────────
// API 키 확인
// ─────────────────────────────────────────────────────────
const API_KEY = cfg.ANTHROPIC_API_KEY;
if (!API_KEY) {
  console.error('');
  console.error('[오류] Anthropic API 키가 설정되지 않았습니다.');
  console.error('');
  console.error('  ① https://console.anthropic.com 에서 API 키 발급');
  console.error('  ② config.js 파일을 열어 ANTHROPIC_API_KEY 값을 입력');
  console.error('');
  process.exit(1);
}

const client = new Anthropic({ apiKey: API_KEY });

// ─────────────────────────────────────────────────────────
// 설정
// ─────────────────────────────────────────────────────────
const CONFIG = {
  outputDir: 'C:\\Users\\AWESOMATIC\\Desktop\\코에르\\클로드\\리뷰수집프로그램',
  batchSize: 10,
  model: 'claude-haiku-4-5-20251001',
};

// ─────────────────────────────────────────────────────────
// 코에르 브랜드 컨텍스트 + AI 프롬프트
// ─────────────────────────────────────────────────────────
const SYSTEM_PROMPT = `당신은 코에르(COEIR) 브랜드의 고객 서비스 담당자입니다.

【브랜드 소개】
코에르는 고품질 욕실 용품 전문 한국 브랜드입니다.
- 욕실 소품: 스텐/테라조 디스펜서·멀티홀더·트레이·비누받침대·칫솔꽂이 (무타공 설치 가능)
- 욕실 매트: 규조토 발매트, 미끄럼방지 매트 (S/M/L, 화이트·그레이·베이지 등)
- 샤워용품: 온오프 샤워헤드, PLA·ACF 정수필터, 샤워호스 (절수·수압 개선 기능)
- 수건: 40수 코마사 프리미엄 타올 (호텔 퀄리티)
- 욕실선반 (플랫/히든 타입), 욕실화

【환불 판단 기준 - 엄격히 적용】

■ 별점 1~2점 → 기본 환불검토
  단, 아래는 답변 처리:
  · 배송만 늦었거나 포장 아쉬움 (제품 자체는 만족)
  · 고객이 사이즈·색상을 잘못 선택한 경우
  · "별점 실수로 낮게 줬어요" 유형
  · 전반적으로 내용이 긍정적인 경우

■ 별점 3점 → 내용 기반 판단
  환불검토: 제품 불량·파손·기능 결함이 명확한 경우
  답변 처리: 아쉬운 점은 있지만 사용 가능, 취향 차이, 설치 어려움

■ 별점 4~5점 → 기본 답변 처리
  환불검토: 내용에 명백한 제품 불량·파손·기능 오작동이 있을 때만

【⚠️ 중요 - 문맥 파악】
  "파손 없이", "파손되지 않고", "불량 없이" → 환불 아님 (부정 표현!)
  "필터가 있어서 안심" → 긍정 표현, 환불 아님
  "포장을 잘해주셔서 파손되지 않고" → 완전 긍정, 환불 아님

【답변 작성 기준 (환불검토가 아닐 때만)】
  · 정확히 100자 이내 (한글 기준, 절대 초과 금지)
  · 이모티콘 2~3개 포함 (내용에 어울리는 것)
  · 리뷰에서 언급한 구체적 내용 1가지 언급
  · 따뜻하고 친근한 톤, "감사합니다" 또는 "고맙습니다" 필수 포함
  · 답변 예시 톤: "고객님 덕분에 저희도 기뻐요! 😊 [구체적 내용] 코에르가 늘 함께하겠습니다 🤍"

【출력 형식 - 반드시 유효한 JSON 배열만 출력】
[
  {"idx": 0, "refund": false, "response": "답변 내용 (100자 이내)"},
  {"idx": 1, "refund": true, "response": null}
]`;

// ─────────────────────────────────────────────────────────
// 유틸
// ─────────────────────────────────────────────────────────
const log = msg => console.log(msg);
const sleep = ms => new Promise(r => setTimeout(r, ms));

function timestamp() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}

// ─────────────────────────────────────────────────────────
// 최신 수집 파일 탐색
// ─────────────────────────────────────────────────────────
function findLatestExcel() {
  const files = fs.readdirSync(CONFIG.outputDir)
    .filter(f => f.startsWith('리뷰_코에르_전체_') && f.endsWith('.xlsx'))
    .sort()
    .reverse();
  return files.length ? path.join(CONFIG.outputDir, files[0]) : null;
}

// ─────────────────────────────────────────────────────────
// 엑셀 읽기
// ─────────────────────────────────────────────────────────
function readReviews(filePath) {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets['전체리뷰'];
  if (!ws) throw new Error('"전체리뷰" 시트를 찾을 수 없습니다.');

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  // 헤더 제외, 빈 행 제외
  const data = rows.slice(1).filter(r => r[0] || r[5]); // 제품명이나 리뷰내용이 있는 행
  return data;
}

// ─────────────────────────────────────────────────────────
// Claude API 호출 (배치 단위)
// ─────────────────────────────────────────────────────────
async function callClaude(batch) {
  const input = batch.map((row, i) => ({
    idx: i,
    product: String(row[0] || ''),
    rating: row[1],
    content: String(row[5] || ''),
  }));

  const resp = await client.messages.create({
    model: CONFIG.model,
    max_tokens: 2048,
    system: SYSTEM_PROMPT,
    messages: [{ role: 'user', content: JSON.stringify(input) }],
  });

  const text = resp.content[0].text.trim();
  const match = text.match(/\[[\s\S]*\]/);
  if (!match) throw new Error(`JSON 파싱 실패:\n${text.substring(0, 300)}`);
  return JSON.parse(match[0]);
}

// ─────────────────────────────────────────────────────────
// 전체 리뷰 처리
// ─────────────────────────────────────────────────────────
async function processAll(reviews) {
  const totalBatches = Math.ceil(reviews.length / CONFIG.batchSize);
  const results = new Array(reviews.length);

  for (let b = 0; b < totalBatches; b++) {
    const start = b * CONFIG.batchSize;
    const batch = reviews.slice(start, start + CONFIG.batchSize);

    process.stdout.write(`  배치 ${String(b+1).padStart(3)}/${totalBatches} (${start+1}~${start+batch.length}번 리뷰)... `);

    let aiResults = null;
    for (let retry = 0; retry < 3; retry++) {
      try {
        aiResults = await callClaude(batch);
        break;
      } catch (err) {
        process.stdout.write(`재시도${retry+1} `);
        await sleep(2000 * (retry + 1));
      }
    }

    if (!aiResults) {
      // 실패 시 기본값
      aiResults = batch.map((_, i) => ({ idx: i, refund: false, response: '리뷰 감사합니다 😊' }));
      process.stdout.write('(기본값 적용) ');
    }

    for (const r of aiResults) {
      const row = batch[r.idx];
      if (!row) continue;
      const globalIdx = start + r.idx;
      results[globalIdx] = {
        data: row.slice(0, 6), // 제품명~리뷰내용 (기존 환불 컬럼 제거)
        refund: r.refund ? '환불검토' : '',
        response: r.refund ? '' : (r.response || ''),
      };
    }

    process.stdout.write('✓\n');

    if (b < totalBatches - 1) await sleep(300); // 레이트리밋 방지
  }

  return results.filter(Boolean);
}

// ─────────────────────────────────────────────────────────
// 엑셀 저장
// ─────────────────────────────────────────────────────────
function saveExcel(results) {
  const wb = XLSX.utils.book_new();
  const HEADERS = ['제품명', '별점', '아이디', '날짜', '구매옵션', '리뷰내용', '환불', '답변'];
  const COL_WIDTHS = [
    { wch: 42 },   // 제품명
    { wch: 6 },    // 별점
    { wch: 18 },   // 아이디
    { wch: 14 },   // 날짜
    { wch: 25 },   // 구매옵션
    { wch: 80 },   // 리뷰내용
    { wch: 10 },   // 환불
    { wch: 55 },   // 답변
  ];

  const allRows = results.map(r => [...r.data, r.refund, r.response]);

  // 전체리뷰 시트
  const ws1 = XLSX.utils.aoa_to_sheet([HEADERS, ...allRows]);
  ws1['!cols'] = COL_WIDTHS;
  XLSX.utils.book_append_sheet(wb, ws1, '전체리뷰');

  // 환불검토 시트 (환불 대상만)
  const refundRows = allRows.filter(r => r[6] === '환불검토');
  if (refundRows.length > 0) {
    const ws2 = XLSX.utils.aoa_to_sheet([HEADERS, ...refundRows]);
    ws2['!cols'] = COL_WIDTHS;
    XLSX.utils.book_append_sheet(wb, ws2, '환불검토');
  }

  // 요약 시트 (제품별 통계)
  const byProduct = {};
  for (const r of results) {
    const name = r.data[0] || '기타';
    if (!byProduct[name]) byProduct[name] = { total: 0, refund: 0, withResp: 0 };
    byProduct[name].total++;
    if (r.refund) byProduct[name].refund++;
    if (r.response) byProduct[name].withResp++;
  }
  const summaryRows = Object.entries(byProduct)
    .sort((a, b) => b[1].total - a[1].total)
    .map(([name, s]) => [name, s.total, s.refund, s.withResp]);

  const ws3 = XLSX.utils.aoa_to_sheet([
    ['제품명', '리뷰 수', '환불검토', '답변 생성'],
    ...summaryRows,
    [],
    ['합계', results.length, refundRows.length, results.filter(r => r.response).length],
  ]);
  ws3['!cols'] = [{ wch: 42 }, { wch: 10 }, { wch: 10 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, ws3, '요약');

  const outPath = path.join(CONFIG.outputDir, `리뷰_AI처리_${timestamp()}.xlsx`);
  XLSX.writeFile(wb, outPath);
  return { outPath, refundCount: refundRows.length, responseCount: results.filter(r => r.response).length };
}

// ─────────────────────────────────────────────────────────
// 메인
// ─────────────────────────────────────────────────────────
async function main() {
  log('');
  log('╔══════════════════════════════════════════════════════╗');
  log('║   코에르 리뷰 AI 처리 프로그램                      ║');
  log('║   환불 판단 + 답변 생성 (Claude Haiku)              ║');
  log('╚══════════════════════════════════════════════════════╝');
  log('');

  // 파일 탐색
  const srcFile = findLatestExcel();
  if (!srcFile) {
    log('[오류] 수집된 리뷰 파일이 없습니다.');
    log('  → 먼저 전체리뷰수집_실행.bat 을 실행하세요.');
    return;
  }
  log(`[원본] ${path.basename(srcFile)}`);

  // 리뷰 읽기
  const reviews = readReviews(srcFile);
  log(`[리뷰] ${reviews.length}개`);

  const estSec = Math.ceil(reviews.length / CONFIG.batchSize * 1.5);
  log(`[예상] 약 ${estSec}초 (${Math.ceil(estSec/60)}분)\n`);

  // AI 처리
  log('[AI 처리]');
  const results = await processAll(reviews);

  const refundCount = results.filter(r => r.refund).length;
  const responseCount = results.filter(r => r.response).length;
  log(`\n  완료: 총 ${results.length}개`);
  log(`  환불검토: ${refundCount}개`);
  log(`  답변 생성: ${responseCount}개`);

  // 저장
  log('\n[엑셀 저장]');
  const { outPath } = saveExcel(results);
  log(`  ✓ ${outPath}`);
  log(`  ✓ 시트: 전체리뷰 / 환불검토 / 요약`);

  log('');
  log('╔══════════════════════════════════════════════════════╗');
  log('║   완료 - 엑셀을 열어 답변을 확인하세요              ║');
  log('╚══════════════════════════════════════════════════════╝');
}

main().catch(err => {
  console.error('\n[오류]', err.message);
  process.exit(1);
});
