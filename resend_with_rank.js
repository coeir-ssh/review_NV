/**
 * 오늘 작업 결과에 리뷰 순위를 추가해서 Excel/Word/Slack 재발송
 * 실행: node resend_with_rank.js
 */
const puppeteer = require('puppeteer');
const Anthropic = require('@anthropic-ai/sdk');
const XLSX      = require('xlsx');
const path      = require('path');
const fs        = require('fs');
const https     = require('https');
const cfg       = require('./config');
const { getMasterProductName, inferOption } = require('./master_product_names');
function resolveDisplayOption(r) {
  return (r.optionName && r.optionName.trim()) || inferOption(r.productName) || '';
}
const REASON_CONFIDENCE_THRESHOLD = 90;
function shouldShowJudgeReason(r) {
  if (!r.judgeReason) return false;
  if (r.judgeLabel === '환불검토') return true;
  if (r.judgeConfidence != null && r.judgeConfidence < REASON_CONFIDENCE_THRESHOLD) return true;
  return false;
}
function rankSortKey(p) {
  if (typeof p === 'number' && p > 0) return p;
  if (p === -2) return 9999;
  return 99999;
}
function sortByRank(list) {
  return [...list].sort((a, b) => rankSortKey(a.reviewPosition) - rankSortKey(b.reviewPosition));
}

function normalizeReviewText(text) {
  return (text || '')
    .replace(/[\u{1F000}-\u{1FFFF}\u{2300}-\u{27FF}\u{FE00}-\u{FEFF}\u{1FA00}-\u{1FAFF}]/gu, '')
    .replace(/[…⋯⋮]/g, '.')
    .replace(/\.{2,}/g, '.')
    .replace(/\s+/g, ' ')
    .trim();
}
function koreanOnlyKey(text) {
  return (text || '').replace(/[^가-힣]/g, '');
}
function formatReviewPosition(p) {
  if (p > 0) return `${p}위`;
  if (p === -2) return '100위 밖';
  return '미확인';
}
const {
  Document, Packer, Paragraph, TextRun,
  AlignmentType, BorderStyle, ShadingType,
} = require('docx');

const sleep = ms => new Promise(r => setTimeout(r, ms));
const log   = msg => console.log(msg);

// ── product_naver_ids 로드 ───────────────────────────────
let _ids = null;
function getIds() {
  if (!_ids) {
    const p = path.join(__dirname, 'product_naver_ids.json');
    _ids = fs.existsSync(p) ? JSON.parse(fs.readFileSync(p)) : {};
  }
  return _ids;
}
function findProductNaverId(productName) {
  const ids   = getIds();
  const pName = (productName || '').replace(/\s+/g, ' ').trim();
  for (const [urlId, info] of Object.entries(ids)) {
    if (info.name === pName) return { urlId, ...info };
  }
  const keywords = pName.split(/[\s[\]()·]+/).filter(w => w.length >= 2);
  let best = null, bestScore = 0;
  for (const [urlId, info] of Object.entries(ids)) {
    const iName = (info.name || '').replace(/\s+/g, ' ');
    let score = 0;
    for (const kw of keywords) { if (iName.includes(kw)) score++; }
    if (score > bestScore) { bestScore = score; best = { urlId, ...info }; }
  }
  return bestScore >= 2 ? best : null;
}

// ── brand.naver.com 순위 조회 ────────────────────────────
async function findReviewPosition(browser, productName, reviewText, writer) {
  if (!reviewText || reviewText.trim().length < 3) {
    return { position: -1, policy: '리뷰 내용 없음', productUrl: null };
  }
  const productInfo = findProductNaverId(productName);
  if (!productInfo || !productInfo.originProductNo) {
    return { position: -1, policy: '제품 ID 없음', productUrl: null };
  }

  const { urlId, originProductNo, checkoutMerchantNo } = productInfo;
  const productUrl   = `https://brand.naver.com/coeir/products/${urlId}`;
  const matchSnippet = normalizeReviewText(reviewText).substring(0, 30);
  const matchKor     = koreanOnlyKey(matchSnippet);

  const reviewPage = await browser.newPage();
  try {
    await reviewPage.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36');
    await reviewPage.goto(productUrl + '#REVIEW', { waitUntil: 'networkidle2', timeout: 30000 });
    await sleep(2000);

    const API_URL   = 'https://brand.naver.com/n/v1/contents/reviews/query-pages';
    const PAGE_SIZE = 20;
    const TOP_LIMIT = 100;
    const MAX_PAGES = Math.ceil(TOP_LIMIT / PAGE_SIZE);
    let globalPos = 0, foundPos = -1;

    for (let p = 1; p <= MAX_PAGES; p++) {
      const body = { checkoutMerchantNo, originProductNo, page: p, pageSize: PAGE_SIZE, reviewSearchSortType: 'REVIEW_RANKING' };
      const result = await reviewPage.evaluate(async (url, reqBody) => {
        try {
          const resp = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(reqBody) });
          if (!resp.ok) return { error: `HTTP ${resp.status}` };
          return await resp.json();
        } catch(e) { return { error: e.message }; }
      }, API_URL, body);

      if (result.error) break;
      const reviews = result.contents || result.reviews || [];
      if (!reviews.length) break;

      for (const rv of reviews) {
        globalPos++;
        if (globalPos > TOP_LIMIT) break;
        const rawText = rv.reviewContent || rv.reviewBody || rv.body || rv.content || rv.reviewText || rv.text || '';
        const normText = normalizeReviewText(rawText);
        const normKor  = koreanOnlyKey(normText);
        const apiWriter = (rv.writerMemberId || rv.writerId || rv.maskedWriterId || rv.memberMaskingId || rv.writer || '').toString();
        const writerMatched = writer && apiWriter && (
          apiWriter === writer || apiWriter.replace(/\*/g, '') === writer.replace(/\*/g, '')
        );
        const textMatched =
          normText.includes(matchSnippet) ||
          matchSnippet.includes(normText.substring(0, 15)) ||
          normText.includes(matchSnippet.substring(0, 15)) ||
          (matchKor.length >= 4 && normKor.length >= 4 && (
            normKor.startsWith(matchKor) ||
            matchKor.startsWith(normKor.substring(0, Math.min(matchKor.length, normKor.length)))
          ));
        const matched = textMatched && (matchKor.length >= 7 || writerMatched || !writer);
        if (matched) { foundPos = globalPos; break; }
      }
      if (foundPos > 0 || globalPos >= TOP_LIMIT) break;
      const total = result.totalCount || result.totalElements || 0;
      if (total > 0 && globalPos >= total) break;
      await sleep(300);
    }

    if (foundPos <= 0) foundPos = (globalPos > 0) ? -2 : -1;
    const outOfTop100 = (foundPos === -2);

    const policy = foundPos === -2 ? '✅ 100위 밖 → 답변으로 충분'
      : foundPos <= 0 ? null
      : foundPos <= 10 ? '⚠️ 1~10위 → 아주 적극 대응 (조건 환불 검토)'
      : foundPos <= 20 ? '🔶 11~20위 → 적극 대응 (답변 or 환불)'
      : foundPos <= 40 ? '🔷 21~40위 → 답변 우선 (필요시 환불)'
      :                  '✅ 41위 이하 → 답변으로 충분';

    log(`    → ${foundPos > 0 ? `${foundPos}위` : outOfTop100 ? '100위 밖' : '미발견'} / 검색: ${globalPos}개`);
    return { position: foundPos, policy, productUrl };
  } catch(e) {
    return { position: -1, policy: null, productUrl };
  } finally {
    await reviewPage.close().catch(() => {});
  }
}

// ── Excel 생성 ───────────────────────────────────────────
function generateExcel(results, dateStr) {
  const data = results.map(r => ({
    '리뷰글번호': r.reviewNo,
    '상품명':     r.productName,
    '리뷰순위':   formatReviewPosition(r.reviewPosition),
    '구매자평점': r.rating,
    '리뷰등록일': r.date,
    '리뷰내용':   r.reviewText,
    '답변내용':   r.replyText || '',
    '환불검토':   r.refundCheck,
    '대응정책':   r.refundCheck === '검토필요' ? (r.refundPolicy || '미확인') : '-',
    '판단 근거': r.judgeReason || '-',
  }));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = [
    { wch: 15 }, { wch: 30 }, { wch: 10 }, { wch: 10 },
    { wch: 20 }, { wch: 50 }, { wch: 60 }, { wch: 12 }, { wch: 35 }, { wch: 50 },
  ];
  XLSX.utils.book_append_sheet(wb, ws, '리뷰답변');
  const replyDir = path.join(__dirname, 'reply');
  if (!fs.existsSync(replyDir)) fs.mkdirSync(replyDir, { recursive: true });
  const filename = `${dateStr}_reply.xlsx`;
  const filepath = path.join(replyDir, filename);
  XLSX.writeFile(wb, filepath);
  return { filename, filepath };
}

// ── Word 문서 생성 ───────────────────────────────────────
async function generateWordDoc(summary, dateStr) {
  const stars   = n => '⭐'.repeat(Math.max(0, Math.min(5, n || 0)));
  const divider = () => new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'AAAAAA', space: 1 } },
    spacing: { after: 160 }, children: [],
  });
  const label = (txt, bold = false, size = 22) =>
    new TextRun({ text: txt, bold, size, font: 'Malgun Gothic' });
  const GRAY = 'F2F2F2';

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { after: 80 },
      children: [new TextRun({ text: '코에르 리뷰 자동 답변 보고서', bold: true, size: 32, font: 'Malgun Gothic', color: '1F3864' })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { after: 300 },
      children: [new TextRun({ text: `작업일: ${summary.date}`, size: 22, font: 'Malgun Gothic', color: '666666' })],
    }),
    new Paragraph({
      spacing: { before: 100, after: 100 }, shading: { fill: 'D9E1F2', type: ShadingType.CLEAR },
      children: [label('작업 요약', true, 24)],
    }),
    new Paragraph({ spacing: { after: 80 }, children: [label(`• 조건: 1주일 작성 + 답글미등록`)] }),
    new Paragraph({
      spacing: { after: 80 },
      children: [label(`• 총 처리: ${summary.replied}개 완료 ✅   |   환불검토: ${summary.refund}개   |   실패: ${summary.failed}개`)],
    }),
    new Paragraph({
      spacing: { after: 80 },
      children: [label(`• 실행환경 : ${summary.executionEnv || '백그라운드 자동 (Anthropic API 직접 호출)'}`)],
    }),
    new Paragraph({
      spacing: { after: 80 },
      children: [label(`• AI 모델 : ${summary.aiModel || 'claude-sonnet-4-5'}`)],
    }),
    new Paragraph({
      spacing: { after: summary.usage ? 80 : 300 },
      children: [label(`• 비용 : ${summary.cost || (summary.usage && summary.usage.costKRW != null ? `약 ₩${summary.usage.costKRW.toLocaleString()} ($${summary.usage.costUSD})` : '-')}`)],
    }),
    ...(summary.usage ? [new Paragraph({
      spacing: { after: 300 },
      children: [label(
        `• API 사용: ${summary.usage.calls}회 호출   |   토큰 ${summary.usage.totalTokens.toLocaleString()} `
        + `(입력 ${summary.usage.input.toLocaleString()} / 출력 ${summary.usage.output.toLocaleString()}`
        + ` / 캐시읽기 ${summary.usage.cacheRead.toLocaleString()} / 캐시쓰기 ${summary.usage.cacheWrite.toLocaleString()})`
        + `   |   비용: $${summary.usage.costUSD} (≈ ₩${summary.usage.costKRW.toLocaleString()})`
      )],
    })] : []),
    divider(),

    // ── 환불검토 (먼저 표시)
    ...(() => {
      const list = sortByRank(summary.results.filter(r => r.refundCheck === '검토필요'));
      if (!list.length) return [];
      return [
        new Paragraph({ spacing: { before: 200, after: 160 }, children: [label('⚠️ 환불검토 필요 항목', true, 24)] }),
        ...list.flatMap((r, i) => [
          new Paragraph({ spacing: { before: 160, after: 80 }, shading: { fill: 'FFF2CC', type: ShadingType.CLEAR }, children: [label(`No.${i + 1}   ${stars(r.rating)} (${r.rating}점)   |   ${r.writer}`, true)] }),
          new Paragraph({ spacing: { after: 60 }, children: [label('리뷰글번호 : ', true), label(r.reviewNo || '-')] }),
          new Paragraph({ spacing: { after: 60 }, children: [label('제품명 : ', true), label(getMasterProductName(r.productName))] }),
          ...((opt => opt ? [new Paragraph({ spacing: { after: 60 }, children: [label('구매 옵션 : ', true), label(opt)] })] : [])(resolveDisplayOption(r))),
          new Paragraph({ spacing: { after: 60 }, children: [label('별점 : ', true), label(`${stars(r.rating)} (${r.rating}점)`)] }),
          ...(r.reviewPosition > 0 ? [
            new Paragraph({ spacing: { after: 60 }, children: [label('📍 리뷰 순위 : ', true), new TextRun({ text: `${r.reviewPosition}위 (랭킹순)`, font: 'Malgun Gothic', size: 22, bold: true, color: 'C00000' })] }),
            new Paragraph({ spacing: { after: 60 }, children: [label('📋 대응 정책 : ', true), new TextRun({ text: r.refundPolicy || '', font: 'Malgun Gothic', size: 22, bold: true, color: '833C00' })] }),
          ] : [
            new Paragraph({ spacing: { after: 60 }, children: [label('리뷰순위 : ', true), new TextRun({ text: formatReviewPosition(r.reviewPosition), font: 'Malgun Gothic', size: 22, color: r.reviewPosition === -2 ? '888888' : '000000' })] }),
            ...(r.refundPolicy ? [new Paragraph({ spacing: { after: 60 }, children: [label('📋 대응 정책 : ', true), label(r.refundPolicy)] })] : []),
          ]),
          new Paragraph({ spacing: { after: 60 }, children: [label('리뷰 : ', true), new TextRun({ text: `"${r.reviewText.replace(/\n/g, ' ')}"`, font: 'Malgun Gothic', size: 22, italics: true })] }),
          ...(r.judgeLabel ? [new Paragraph({ spacing: { after: 60 }, children: [
            new TextRun({ text: '🏷️ 판단 : ', bold: true, size: 22, font: 'Malgun Gothic' }),
            new TextRun({ text: r.judgeLabel, bold: true, size: 22, font: 'Malgun Gothic', color: r.judgeLabel === '환불검토' ? 'C00000' : '2E7D32' }),
          ] })] : []),
          ...(r.judgeReason ? [new Paragraph({ spacing: { after: 60 }, children: [
            new TextRun({ text: '💭 판단 근거 : ', bold: true, size: 22, font: 'Malgun Gothic', color: 'C00000' }),
            new TextRun({ text: `${r.judgeReason}${r.judgeConfidence != null ? ` (confidence ${r.judgeConfidence})` : ''}`, font: 'Malgun Gothic', size: 22, bold: true, color: 'C00000' }),
          ] })] : []),
          ...(r.replyText ? [new Paragraph({ spacing: { after: 60 }, children: [
            new TextRun({ text: '💬 답변 대응시 : ', bold: true, size: 22, font: 'Malgun Gothic', color: '7030A0' }),
            new TextRun({ text: `"${r.replyText}"`, font: 'Malgun Gothic', size: 22, color: '7030A0' }),
          ] })] : []),
          new Paragraph({ spacing: { after: 80 }, children: [] }),
          new Paragraph({ spacing: { after: 160 }, children: [label('피드백 : ', true)] }),
        ]),
        divider(),
      ];
    })(),

    // ── 답변 항목 (판단체크 필요 / 일반 분리) ──
    ...(() => {
      const renderAnswered = (r, i, showReason) => [
        new Paragraph({ spacing: { before: 200, after: 80 }, shading: { fill: GRAY, type: ShadingType.CLEAR }, children: [label(`No.${i + 1}`, true, 22)] }),
        new Paragraph({ spacing: { after: 60 }, children: [label('리뷰글번호 : ', true), label(r.reviewNo || '-')] }),
        new Paragraph({ spacing: { after: 60 }, children: [label('등록자 : ', true), label(r.writer)] }),
        new Paragraph({ spacing: { after: 60 }, children: [label('제품명 : ', true), label(getMasterProductName(r.productName))] }),
        ...((opt => opt ? [new Paragraph({ spacing: { after: 60 }, children: [label('구매 옵션 : ', true), label(opt)] })] : [])(resolveDisplayOption(r))),
        new Paragraph({ spacing: { after: 60 }, children: [label('별점 : ', true), label(`${stars(r.rating)} (${r.rating}점)`)] }),
        new Paragraph({ spacing: { after: 60 }, children: [label('리뷰순위 : ', true), new TextRun({ text: formatReviewPosition(r.reviewPosition), font: 'Malgun Gothic', size: 22, bold: true, color: r.reviewPosition === -2 ? '888888' : '1F3864' })] }),
        new Paragraph({ spacing: { after: 60 }, children: [label('리뷰 : ', true), new TextRun({ text: `"${r.reviewText.replace(/\n/g, ' ')}"`, font: 'Malgun Gothic', size: 22, italics: true })] }),
        ...(r.judgeLabel ? [new Paragraph({ spacing: { after: 60 }, children: [
          new TextRun({ text: '🏷️ 판단 : ', bold: true, size: 22, font: 'Malgun Gothic' }),
          new TextRun({ text: r.judgeLabel, bold: true, size: 22, font: 'Malgun Gothic', color: r.judgeLabel === '환불검토' ? 'C00000' : '2E7D32' }),
        ] })] : []),
        ...(showReason ? [new Paragraph({ spacing: { after: 60 }, children: [
          new TextRun({ text: '💭 판단 근거 : ', bold: true, size: 22, font: 'Malgun Gothic', color: 'C00000' }),
          new TextRun({ text: `${r.judgeReason}${r.judgeConfidence != null ? ` (confidence ${r.judgeConfidence})` : ''}`, font: 'Malgun Gothic', size: 22, bold: true, color: 'C00000' }),
        ] })] : []),
        new Paragraph({ spacing: { after: 80 }, children: [label('답변 : ', true), new TextRun({ text: `"${r.replyText}"`, font: 'Malgun Gothic', size: 22, color: '1F497D' })] }),
        new Paragraph({ spacing: { after: 80 }, children: [] }),
        new Paragraph({ spacing: { after: 160 }, children: [label('피드백 : ', true)] }),
      ];
      const answeredAll = summary.results.filter(r => r.replyText && r.refundCheck !== '검토필요');
      const checkList   = sortByRank(answeredAll.filter(r => shouldShowJudgeReason(r)));
      const normalList  = sortByRank(answeredAll.filter(r => !shouldShowJudgeReason(r)));
      const out = [];
      if (checkList.length > 0) {
        out.push(new Paragraph({ spacing: { before: 200, after: 160 }, children: [label('⚠️ 답변 완료 - 판단체크 필요', true, 24)] }));
        checkList.forEach((r, i) => out.push(...renderAnswered(r, i, true)));
        out.push(divider());
      }
      out.push(new Paragraph({ spacing: { before: 200, after: 160 }, children: [label('■ 답변 완료 항목', true, 24)] }));
      normalList.forEach((r, i) => out.push(...renderAnswered(r, i, false)));
      return out;
    })(),
    divider(),

    // ── 실패
    ...(() => {
      const list = summary.results.filter(r => !r.replyText && r.refundCheck === '-');
      if (!list.length) return [];
      return [
        new Paragraph({ spacing: { before: 200, after: 160 }, children: [label('❌ 답변 실패 항목', true, 24)] }),
        ...list.flatMap((r, i) => [
          new Paragraph({ spacing: { before: 160, after: 80 }, shading: { fill: 'FCE4D6', type: ShadingType.CLEAR }, children: [label(`No.${i + 1}   ${stars(r.rating)} (${r.rating}점)   |   ${r.writer}`, true)] }),
          new Paragraph({ spacing: { after: 60 }, children: [label('제품명 : ', true), label(getMasterProductName(r.productName))] }),
          new Paragraph({ spacing: { after: 160 }, children: [label('리뷰 : ', true), new TextRun({ text: `"${r.reviewText.replace(/\n/g, ' ')}"`, font: 'Malgun Gothic', size: 22, italics: true })] }),
        ]),
        divider(),
      ];
    })(),

    new Paragraph({ spacing: { before: 200 }, alignment: AlignmentType.RIGHT, children: [label(`📊 Excel: reply/${summary.excelFile} 저장 완료`, false, 20)] }),
  ];

  const doc = new Document({
    styles: { default: { document: { run: { font: 'Malgun Gothic', size: 22 } } } },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children }],
  });

  const replyDir = path.join(__dirname, 'reply');
  if (!fs.existsSync(replyDir)) fs.mkdirSync(replyDir, { recursive: true });
  const filename = `${dateStr}_reply.docx`;
  const filepath = path.join(replyDir, filename);
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(filepath, buffer);
  return { filename, filepath };
}

// ── Slack 전송 ───────────────────────────────────────────
// 워드 파일을 슬랙 채널에 업로드
async function uploadFileToSlack(filepath, channelId, comment = '') {
  const token = cfg.SLACK_BOT_TOKEN;
  if (!token || !channelId || !fs.existsSync(filepath)) { log('[슬랙 파일] 스킵'); return; }
  const filename = path.basename(filepath);
  const fileBuf  = fs.readFileSync(filepath);
  const length   = fileBuf.length;
  const httpsReq = (opts, body) => new Promise((resolve, reject) => {
    const req = https.request(opts, res => {
      let data=''; res.on('data', c => data+=c); res.on('end', () => resolve({status:res.statusCode, body:data}));
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
  try {
    const step1Body = `filename=${encodeURIComponent(filename)}&length=${length}`;
    const step1 = await httpsReq({
      hostname:'slack.com', path:'/api/files.getUploadURLExternal', method:'POST',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/x-www-form-urlencoded','Content-Length':Buffer.byteLength(step1Body)},
    }, step1Body);
    const j1 = JSON.parse(step1.body);
    if (!j1.ok) { log(`[슬랙 파일] URL 발급 실패: ${j1.error}`); return; }
    const url = new URL(j1.upload_url);
    const step2 = await httpsReq({
      hostname:url.hostname, path:url.pathname+url.search, method:'POST',
      headers:{'Content-Type':'application/octet-stream','Content-Length':length},
    }, fileBuf);
    if (step2.status >= 400) { log(`[슬랙 파일] 업로드 HTTP ${step2.status}`); return; }
    const step3Body = JSON.stringify({files:[{id:j1.file_id,title:filename}],channel_id:channelId,initial_comment:comment||''});
    const step3 = await httpsReq({
      hostname:'slack.com', path:'/api/files.completeUploadExternal', method:'POST',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json; charset=utf-8','Content-Length':Buffer.byteLength(step3Body)},
    }, step3Body);
    const j3 = JSON.parse(step3.body);
    if (j3.ok) log(`[슬랙 파일] 업로드 완료: ${filename}`);
    else       log(`[슬랙 파일] 게시 실패: ${j3.error}`);
  } catch (e) { log(`[슬랙 파일] 오류: ${e.message}`); }
}

async function sendSlack(summary) {
  const token     = cfg.SLACK_BOT_TOKEN;
  const channelId = cfg.SLACK_CHANNEL_ID || cfg.SLACK_USER_ID;
  if (!token) { log('[슬랙] 토큰 없음'); return; }

  const stars = n => '⭐'.repeat(Math.max(0, Math.min(5, n || 0)));

  const lines = [
    `안녕하세요! 오늘(${summary.date}) 코에르 리뷰 자동 답변 작업 완료 보고드립니다 😊`,
    `작업 요약`,
    `• 조건: 1주일 작성 + 답글미등록 / Claude Sonnet API`,
    `• 총 처리: ${summary.replied}개 완료 ✅  |  환불검토: ${summary.refund}개  |  실패: ${summary.failed}개`,
  ];
  if (summary.usage) {
    lines.push(
      `• API 사용: ${summary.usage.calls}회 호출  |  토큰 ${summary.usage.totalTokens.toLocaleString()} `
      + `(입력 ${summary.usage.input.toLocaleString()} / 출력 ${summary.usage.output.toLocaleString()}`
      + ` / 캐시읽기 ${summary.usage.cacheRead.toLocaleString()} / 캐시쓰기 ${summary.usage.cacheWrite.toLocaleString()})`
      + `  |  비용: $${summary.usage.costUSD} (≈ ₩${summary.usage.costKRW.toLocaleString()})`
    );
  }
  lines.push(``);

  // 환불검토 (먼저 표시)
  const refundList = sortByRank(summary.results.filter(r => r.refundCheck === '검토필요'));
  if (refundList.length > 0) {
    lines.push(`─────────────────────────────`);
    lines.push(`⚠️ 환불검토 필요 항목 (${refundList.length}건)`);
    refundList.forEach((r, i) => {
      lines.push(`No.${i + 1}  ${stars(r.rating)} (${r.rating}점)  |  ${r.writer}`);
      lines.push(`리뷰글번호 : ${r.reviewNo || '-'}`);
      lines.push(`제품명 : ${getMasterProductName(r.productName)}`);
      { const opt = resolveDisplayOption(r); if (opt) lines.push(`구매 옵션 : ${opt}`); }
      lines.push(`별점 : ${stars(r.rating)} (${r.rating}점)`);
      lines.push(`리뷰순위 : ${formatReviewPosition(r.reviewPosition)}`);
      lines.push(`📋 대응 정책 : ${r.refundPolicy || '미확인'}`);
      lines.push(`리뷰 : "${r.reviewText.replace(/\n/g, ' ')}"`);
      if (r.judgeLabel) {
        lines.push(`*🏷️ 판단 : ${r.judgeLabel}*`);
      }
      if (r.judgeReason) {
        lines.push(`*🔴 판단 근거 : ${r.judgeReason}${r.judgeConfidence != null ? ` (confidence ${r.judgeConfidence})` : ''}*`);
      }
      if (r.replyText) {
        lines.push(`💬 답변 대응시 : "${r.replyText}"`);
      }
      lines.push(``);
    });
  }

  // 답변 항목 (판단체크 필요 / 일반 분리)
  const answeredAll = summary.results.filter(r => r.replyText && r.refundCheck !== '검토필요');
  const checkList   = sortByRank(answeredAll.filter(r => shouldShowJudgeReason(r)));
  const repliedList = sortByRank(answeredAll.filter(r => !shouldShowJudgeReason(r)));

  // ② 답변 완료 - 판단체크 필요
  if (checkList.length > 0) {
    lines.push(`─────────────────────────────`);
    lines.push(`⚠️ 답변 완료 - 판단체크 필요 (${checkList.length}건)`);
    checkList.forEach((r, i) => {
      lines.push(`─────────────────────────────`);
      lines.push(`No.${i + 1}`);
      lines.push(`리뷰글번호 : ${r.reviewNo || '-'}`);
      lines.push(`등록자 : ${r.writer}`);
      lines.push(`제품명 : ${getMasterProductName(r.productName)}`);
      { const opt = resolveDisplayOption(r); if (opt) lines.push(`구매 옵션 : ${opt}`); }
      lines.push(`별점 : ${stars(r.rating)} (${r.rating}점)`);
      lines.push(`리뷰순위 : ${formatReviewPosition(r.reviewPosition)}`);
      lines.push(`리뷰 : "${r.reviewText.replace(/\n/g, ' ')}"`);
      if (r.judgeLabel) {
        lines.push(`*🏷️ 판단 : ${r.judgeLabel}*`);
      }
      lines.push(`*🔴 판단 근거 : ${r.judgeReason}${r.judgeConfidence != null ? ` (confidence ${r.judgeConfidence})` : ''}*`);
      lines.push(`답변 : "${r.replyText}"`);
      lines.push(``);
    });
  }

  // ③ 답변 완료 항목 (일반)
  if (repliedList.length > 0) {
    lines.push(`─────────────────────────────`);
    lines.push(`■ 답변 완료 항목 (${repliedList.length}건)`);
    repliedList.forEach((r, i) => {
      lines.push(`─────────────────────────────`);
      lines.push(`No.${i + 1}`);
      lines.push(`리뷰글번호 : ${r.reviewNo || '-'}`);
      lines.push(`등록자 : ${r.writer}`);
      lines.push(`제품명 : ${getMasterProductName(r.productName)}`);
      { const opt = resolveDisplayOption(r); if (opt) lines.push(`구매 옵션 : ${opt}`); }
      lines.push(`별점 : ${stars(r.rating)} (${r.rating}점)`);
      lines.push(`리뷰순위 : ${formatReviewPosition(r.reviewPosition)}`);
      lines.push(`리뷰 : "${r.reviewText.replace(/\n/g, ' ')}"`);
      if (r.judgeLabel) {
        lines.push(`*🏷️ 판단 : ${r.judgeLabel}*`);
      }
      lines.push(`답변 : "${r.replyText}"`);
      lines.push(``);
    });
  }

  // 실패
  const failList = summary.results.filter(r => !r.replyText && r.refundCheck === '-');
  if (failList.length > 0) {
    lines.push(`─────────────────────────────`);
    lines.push(`❌ 답변 실패 항목 (${failList.length}건)`);
    failList.forEach((r, i) => {
      lines.push(`No.${i + 1}  ${stars(r.rating)} (${r.rating}점)  |  ${r.writer}`);
      lines.push(`제품명 : ${getMasterProductName(r.productName)}`);
      lines.push(`리뷰 : "${r.reviewText.replace(/\n/g, ' ')}"`);
      lines.push(``);
    });
  }

  lines.push(`─────────────────────────────`);
  lines.push(`📊 Excel: reply/${summary.excelFile} 저장 완료`);

  const text = lines.join('\n');
  return new Promise(resolve => {
    const body = JSON.stringify({ channel: channelId, text });
    const req  = https.request(
      { hostname: 'slack.com', path: '/api/chat.postMessage', method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=utf-8', 'Authorization': `Bearer ${token}` } },
      res => {
        let data = '';
        res.on('data', c => data += c);
        res.on('end', () => {
          try {
            const p = JSON.parse(data);
            log(p.ok ? '[슬랙] 전송 완료' : `[슬랙] 오류: ${p.error}`);
          } catch(e) {}
          resolve();
        });
      }
    );
    req.on('error', e => { log(`[슬랙] 오류: ${e.message}`); resolve(); });
    req.write(body); req.end();
  });
}

// ── 메인 ────────────────────────────────────────────────
async function main() {
  const summaryFile = path.join(__dirname, '.review_summary.json');
  if (!fs.existsSync(summaryFile)) {
    log('[오류] .review_summary.json 없음. 먼저 post_product_reviews.js를 실행하세요.');
    process.exit(1);
  }

  const summary = JSON.parse(fs.readFileSync(summaryFile));
  log(`\n=== 리뷰 순위 추가 재발송 ===`);
  log(`대상: ${summary.results.length}개 리뷰 / 작업일: ${summary.date}`);

  const browser = await puppeteer.launch({
    headless: false,
    protocolTimeout: 120000,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--lang=ko-KR,ko', '--window-size=1400,900'],
    defaultViewport: { width: 1400, height: 900 },
  });

  // 각 리뷰 순위 조회
  for (let i = 0; i < summary.results.length; i++) {
    const r = summary.results[i];
    log(`\n[${i + 1}/${summary.results.length}] ${r.writer} | ${r.productName?.substring(0, 25)}`);
    log(`  리뷰: "${r.reviewText?.substring(0, 40)}"`);

    const pos = await findReviewPosition(browser, r.productName, r.reviewText, r.writer);
    r.reviewPosition = pos.position > 0 ? pos.position : (pos.position === -2 ? -2 : 0);
    r.productUrl     = pos.productUrl || '';

    // 정책은 환불검토 대상만
    if (r.refundCheck === '검토필요') {
      r.refundPolicy = pos.policy || '미확인';
    } else {
      r.refundPolicy = '';
    }
  }

  await browser.close();

  // 날짜 추출 (summary.date 기반)
  const now = new Date();
  const dateStr = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;

  // Excel 재생성
  log('\n[Excel] 재생성 중...');
  const { filename: excelFile } = generateExcel(summary.results, dateStr);
  summary.excelFile = excelFile;
  log(`  → ${excelFile}`);

  // Word 재생성
  log('[Word] 재생성 중...');
  const { filename: docxFile, filepath: docxFilepath } = await generateWordDoc(summary, dateStr);
  log(`  → ${docxFile}`);

  // Slack 재전송
  log('[슬랙] 전송 중...');
  await sendSlack(summary);

  // Slack 채널에 워드 파일 첨부 업로드
  const channelId = cfg.SLACK_CHANNEL_ID || cfg.SLACK_USER_ID;
  if (docxFilepath && channelId) {
    await uploadFileToSlack(docxFilepath, channelId, `📎 ${summary.date} 보고서 워드 파일`);
  }

  // 요약 JSON 업데이트
  fs.writeFileSync(summaryFile, JSON.stringify(summary, null, 2));
  log('\n✅ 완료');
}

module.exports = { generateWordDoc, generateExcel, sendSlack };

if (require.main === module) {
  main().catch(e => { console.error(e); process.exit(1); });
}
