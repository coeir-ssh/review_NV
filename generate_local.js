/**
 * 코에르 리뷰 환불판단 + 답변생성 (로컬 룰 기반)
 * API 키 없이 Claude가 직접 설계한 룰로 처리
 */
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const OUTPUT_DIR = 'C:\\Users\\AWESOMATIC\\Desktop\\코에르\\클로드\\리뷰수집프로그램';

// ─────────────────────────────────────────────────────────
// 환불 판단
// ─────────────────────────────────────────────────────────
function judgeRefund(rating, content) {
  const r = parseInt(rating) || 5;
  const c = content || '';

  // 명백한 불량/사고 키워드 → 별점 무관 환불검토
  const STRONG_REFUND = ['타일깨져', '타일 깨져', '욕조깨져', '욕조 깨져', '파손됐', '파손되었', '다쳤', '다침', '사고', '불량품', '하자있', '하자 있'];
  if (STRONG_REFUND.some(k => c.includes(k))) return true;

  // 1점 → 환불검토 (단 긍정적 내용 제외)
  if (r === 1) {
    const POSITIVE = ['너무 좋', '정말 좋', '만족', '예쁘', '이쁘', '추천', '재구매'];
    if (!POSITIVE.some(k => c.includes(k))) return true;
  }

  // 2점 → 환불검토 (기능 결함·탈락 언급 있을 때)
  if (r === 2) {
    const FUNC_FAIL = ['떨어져', '접착력', '안돼', '안됩', '고장', '작동안', '작동 안', '쉽게 떨어'];
    if (FUNC_FAIL.some(k => c.includes(k))) return true;
    // 2점인데 긍정적 → 답변
    const POSITIVE = ['잘받았', '좋아요', '좋습니다', '만족', '예쁘', '이쁘'];
    if (POSITIVE.some(k => c.includes(k))) return false;
    return true; // 2점 기본 환불
  }

  // 3점 → 기능 결함 명확할 때만 환불
  if (r === 3) {
    const FUNC_FAIL = ['떨어져', '또 떨어', '무섭', '기능 못', '작동안', '파손', '깨져'];
    if (FUNC_FAIL.some(k => c.includes(k))) return true;
    return false;
  }

  // 4~5점 → 파손/불량만 환불
  if (r >= 4) {
    const HARD_FAIL = ['파손', '불량', '깨졌', '고장', '작동안', '하자'];
    // 단, "파손 없이", "파손되지 않고" 같은 부정문은 제외
    const NEG_CONTEXT = ['파손 없', '파손되지', '파손없', '불량 없', '불량없'];
    if (HARD_FAIL.some(k => c.includes(k)) && !NEG_CONTEXT.some(k => c.includes(k))) return true;
    return false;
  }

  return false;
}

// ─────────────────────────────────────────────────────────
// 제품 카테고리 분류
// ─────────────────────────────────────────────────────────
function getCategory(productName) {
  const p = productName || '';
  if (p.includes('샤워기') || p.includes('샤워 헤드') || p.includes('온오프')) return 'shower';
  if (p.includes('필터') || p.includes('PLA') || p.includes('ACF')) return 'filter';
  if (p.includes('호스')) return 'hose';
  if (p.includes('욕실선반') || p.includes('선반')) return 'shelf';
  if (p.includes('트레이') || p.includes('드라이 트레이')) return 'tray';
  if (p.includes('디스펜서')) return 'dispenser';
  if (p.includes('멀티홀더') || p.includes('칫솔꽂이') || p.includes('칫솔걸이')) return 'holder';
  if (p.includes('비누받침')) return 'soap';
  if (p.includes('규조토')) return 'diatomite';
  if (p.includes('욕실매트') || p.includes('미끄럼방지')) return 'mat';
  if (p.includes('욕실화') || p.includes('슬리퍼')) return 'shoes';
  if (p.includes('수건') || p.includes('타월')) return 'towel';
  if (p.includes('부속품') || p.includes('쇼핑백') || p.includes('기프트')) return 'accessory';
  if (p.includes('세트')) return 'set';
  return 'general';
}

// ─────────────────────────────────────────────────────────
// 키워드 추출
// ─────────────────────────────────────────────────────────
function extractKeywords(content) {
  const c = content || '';
  return {
    hasDesign:    /디자인|예쁘|이쁘|깔끔|고급|인테리어|색상|화이트|그레이|베이지/.test(c),
    hasPressure:  /수압|물줄기|물살|세차게|강하게|개운/.test(c),
    hasSafety:    /안전|미끄럼|안미끄|아기|아이|어린이|안심/.test(c),
    hasEasy:      /설치|간단|쉽게|편리|편해|사용하기/.test(c),
    hasQuality:   /품질|고급|튼튼|견고|단단|스텐|스테인/.test(c),
    hasDelivery:  /배송|빠르|배달|도착/.test(c),
    hasRepurchase:/재구매|다시 구매|또 구매|추가 구매|두번째|세번째/.test(c),
    hasAbsorb:    /흡수|빨리 마|건조|수분/.test(c),
    hasFilter:    /필터|정수|깨끗|녹물/.test(c),
    hasFit:       /딱 맞|맞아요|맞습|사이즈|크기/.test(c),
    hasStick:     /붙여|부착|접착|떨어/.test(c),
    hasSatisfy:   /만족|좋아요|좋습니다|최고|1등|추천|굿|good/.test(c),
    hasPackage:   /포장|박스|패키지/.test(c),
  };
}

// ─────────────────────────────────────────────────────────
// 답변 생성
// ─────────────────────────────────────────────────────────
function generateResponse(rating, content, productName) {
  const r = parseInt(rating) || 5;
  const cat = getCategory(productName);
  const kw = extractKeywords(content);
  const c = content || '';

  // 1~2점 만족 유형 (환불 아닌 케이스)
  if (r <= 2) {
    if (kw.hasSatisfy) return trim100('소중한 리뷰 감사해요 😊 제품은 만족하셨다니 다행이에요! 불편하신 점은 고객센터로 알려주시면 빠르게 도움드리겠습니다 🤍');
    return trim100('불편함을 드려 죄송합니다 😔 고객센터로 연락주시면 최대한 빠르게 도움드리겠습니다 🙏 더 나은 코에르가 되겠습니다');
  }

  // 3점 답변
  if (r === 3) {
    if (cat === 'shelf' && kw.hasStick) return trim100('무게 걱정 말씀 충분히 이해해요 🙏 전용 접착제로 시공 시 20kg 이상도 버팁니다! 혹시 불안하시면 CS로 연락주세요 😊');
    if (cat === 'soap') return trim100('비누 크기나 종류에 따라 흘러내릴 수 있어요 🧼 액상 비누나 각진 형태 비누가 더 잘 맞답니다! 참고해 주세요 😊');
    if (cat === 'tray' && c.includes('물자국')) return trim100('물자국은 마른 천으로 자주 닦아주시면 훨씬 깔끔하게 유지되더라고요 ✨ 앞으로 더 좋은 제품으로 보답하겠습니다 🤍');
    if (kw.hasSatisfy) return trim100('소중한 리뷰 감사해요 😊 조금 더 만족드릴 수 있도록 더 노력하겠습니다! 코에르를 선택해 주셔서 고맙습니다 🤍');
    if (c.includes('광고') || c.includes('비싼')) return trim100('가격 부분 말씀해 주셔서 감사해요 🙏 더 좋은 가격에 더 좋은 품질로 보답하도록 노력하겠습니다 😊');
    return trim100('소중한 의견 감사합니다 🙏 더 만족드릴 수 있도록 계속 개선해 나가겠습니다. 코에르를 선택해 주셔서 고맙습니다 🤍');
  }

  // 4~5점 답변 - 카테고리 + 키워드별
  let base = '';

  if (cat === 'shower') {
    if (kw.hasPressure) base = '수압이 마음에 드셨다니 정말 기뻐요 🚿 온오프 기능으로 물도 절약되고 편리하시죠! 늘 함께해 주세요 🤍';
    else if (kw.hasFilter) base = '필터 덕분에 더 깨끗한 샤워 즐기고 계시다니 감사해요 ✨ 아이 있는 가정에서 더욱 든든하게 사용하실 수 있어요 🤍';
    else if (kw.hasDesign) base = '무광 디자인이 욕실과 잘 어울리신다니 기뻐요 😊 코에르 샤워기로 매일 개운한 샤워 하세요 🚿🤍';
    else if (kw.hasPackage) base = '꼼꼼한 포장으로 안전하게 받아보셨다니 다행이에요 😊 코에르 샤워기 오래오래 애용해 주세요 🚿🤍';
    else base = '코에르 샤워기 선택해 주셔서 감사합니다 🚿 매일 상쾌한 샤워 되시길 바랍니다 😊🤍';
  } else if (cat === 'filter') {
    if (kw.hasFilter) base = '깨끗한 물로 더 건강하게 샤워하고 계시다니 기뻐요 ✨ 필터 교체 시기도 확인하시면서 오래 사용해 주세요 😊🤍';
    else base = '코에르 필터 사용해 주셔서 감사합니다 ✨ 깨끗하고 건강한 샤워 생활 되시길 바랍니다 😊🤍';
  } else if (cat === 'shelf') {
    if (kw.hasEasy) base = '무타공 설치가 간편하셨다니 정말 기뻐요 🔧 욕실을 더 깔끔하게 사용하세요! 코에르가 늘 함께합니다 😊🤍';
    else if (kw.hasDesign) base = '욕실 분위기가 달라지셨다니 저도 기뻐요 ✨ 무타공 선반으로 깔끔한 욕실 즐겨주세요 😊🤍';
    else base = '코에르 욕실선반 선택해 주셔서 감사합니다 😊 깔끔하고 실용적인 욕실 즐겨주세요 ✨🤍';
  } else if (cat === 'tray') {
    if (kw.hasDesign) base = '세면대 위가 훨씬 고급스러워 보이시죠 ✨ 코에르 트레이로 매일 기분 좋은 아침 맞이하세요 😊🤍';
    else if (kw.hasQuality) base = '튼튼하고 세련된 트레이 마음에 드셨다니 기뻐요 😊 욕실 인테리어까지 완성해 주셔서 감사합니다 🤍';
    else base = '코에르 트레이 사용해 주셔서 감사합니다 😊 깔끔한 욕실 라이프 즐겨주세요 ✨🤍';
  } else if (cat === 'dispenser') {
    if (kw.hasDesign) base = '욕실이 훨씬 깔끔해지셨다니 기뻐요 🧴 리필도 편리하시죠! 오래오래 사용해 주세요 😊🤍';
    else base = '코에르 디스펜서 사용해 주셔서 감사합니다 🧴 깔끔하고 실용적인 욕실 라이프 되세요 😊🤍';
  } else if (cat === 'holder') {
    if (kw.hasDesign) base = '칫솔도 예쁘게 정리되니 욕실이 더 깔끔해 보이죠 😊 코에르 홀더 오래오래 함께해 주세요 🤍✨';
    else base = '코에르 홀더 사용해 주셔서 감사합니다 😊 깔끔한 욕실 라이프 즐겨주세요 ✨🤍';
  } else if (cat === 'soap') {
    if (kw.hasDesign) base = '비누받침까지 예쁘게 꾸며주셨군요 🧼 코에르로 욕실이 더 멋스러워졌을 것 같아요 😊🤍';
    else base = '코에르 비누받침대 사용해 주셔서 감사합니다 🧼 물빠짐 좋게 오래오래 쓰세요 😊🤍';
  } else if (cat === 'diatomite') {
    if (kw.hasAbsorb) base = '흡수력이 마음에 드셨다니 기뻐요 😊 규조토 발매트로 욕실 바닥도 청결하게 유지하세요 ✨🤍';
    else if (kw.hasDesign) base = '테라조 패턴이 욕실에 잘 어울리시죠 😊 오염도 잘 안 생기고 청결하게 사용하세요 ✨🤍';
    else base = '코에르 규조토 발매트 사용해 주셔서 감사합니다 😊 항상 뽀송뽀송한 욕실 즐겨주세요 ✨🤍';
  } else if (cat === 'mat') {
    if (kw.hasSafety) base = '아이 목욕할 때 더 안심되셨다니 정말 기뻐요 👶 미끄럼 없이 안전하게 사용해 주세요 😊🤍';
    else if (kw.hasFit) base = '사이즈가 딱 맞으셨다니 다행이에요 😊 미끄럼 없이 안전하게 욕실 생활 하세요 🤍✨';
    else base = '코에르 욕실매트 사용해 주셔서 감사합니다 😊 미끄럼 없이 안전한 욕실 생활 되세요 🤍✨';
  } else if (cat === 'shoes') {
    if (kw.hasSafety) base = '미끄럼 없이 안전하게 사용하고 계시다니 기뻐요 😊 코에르 욕실화로 편안한 욕실 생활 되세요 🤍';
    else base = '코에르 욕실화 사용해 주셔서 감사합니다 😊 편안하고 안전한 욕실 생활 되세요 🤍✨';
  } else if (cat === 'towel') {
    if (kw.hasQuality) base = '고급스러운 수건 마음에 드셨다니 기뻐요 🛁 40수 코마사 퀄리티 오래오래 즐겨주세요 😊🤍';
    else base = '코에르 수건 사용해 주셔서 감사합니다 🛁 호텔 같은 욕실 라이프 즐겨주세요 😊🤍';
  } else if (cat === 'set') {
    if (kw.hasDesign) base = '세트로 욕실이 통일감 있게 예뻐지셨겠어요 ✨ 코에르 세트 오래오래 함께해 주세요 😊🤍';
    else base = '코에르 세트 선택해 주셔서 감사합니다 😊 욕실이 더 특별해지셨길 바랍니다 ✨🤍';
  } else {
    // 일반
    if (kw.hasRepurchase) base = '재구매까지 해주셨군요! 정말 감사합니다 😍 코에르를 믿어주셔서 더 열심히 하겠습니다 🤍✨';
    else if (kw.hasDelivery) base = '빠른 배송으로 만족하셨다니 기뻐요 😊 코에르는 앞으로도 신속하게 배송드리겠습니다 🤍';
    else base = '소중한 리뷰 감사합니다 😊 코에르가 늘 만족드릴 수 있도록 더 노력하겠습니다 🤍✨';
  }

  // 재구매 특별 응답 (4~5점 공통)
  if (kw.hasRepurchase && r >= 4 && cat !== 'general') {
    base = '재구매까지 해주셨군요! 정말 감사해요 😍 코에르를 믿어주시는 만큼 더 좋은 제품으로 보답하겠습니다 🤍✨';
  }

  return trim100(base);
}

function trim100(text) {
  // 100자 이내로 자르기 (완전한 문장 단위 유지)
  if (text.length <= 100) return text;
  const cut = text.substring(0, 97);
  const lastSpace = Math.max(cut.lastIndexOf(' '), cut.lastIndexOf('!'), cut.lastIndexOf('.'));
  return (lastSpace > 50 ? cut.substring(0, lastSpace) : cut) + '..';
}

// ─────────────────────────────────────────────────────────
// 유틸
// ─────────────────────────────────────────────────────────
function timestamp() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}

// ─────────────────────────────────────────────────────────
// 메인
// ─────────────────────────────────────────────────────────
function main() {
  console.log('\n╔══════════════════════════════════════════════════════╗');
  console.log('║   코에르 리뷰 환불판단 + 답변생성                   ║');
  console.log('╚══════════════════════════════════════════════════════╝\n');

  // 최신 수집 파일 탐색
  const srcFile = fs.readdirSync(OUTPUT_DIR)
    .filter(f => f.startsWith('리뷰_코에르_전체_') && f.endsWith('.xlsx'))
    .sort().reverse()[0];
  if (!srcFile) { console.log('[오류] 수집 파일 없음'); return; }

  const srcPath = path.join(OUTPUT_DIR, srcFile);
  console.log(`[원본] ${srcFile}`);

  // 읽기
  const wb_in = XLSX.readFile(srcPath);
  const ws = wb_in.Sheets['전체리뷰'];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const data = rows.slice(1).filter(r => r[0] || r[5]);
  console.log(`[리뷰] ${data.length}개\n`);

  // 처리
  const results = [];
  let refundCount = 0, responseCount = 0;

  const byProduct = {};

  for (const row of data) {
    const [product, rating, userId, date, option, content] = row;
    const productName = String(product || '');
    const ratingStr = String(rating || '5');
    const contentStr = String(content || '');

    const isRefund = judgeRefund(ratingStr, contentStr);
    const response = isRefund ? '' : generateResponse(ratingStr, contentStr, productName);

    if (isRefund) refundCount++;
    else responseCount++;

    results.push([productName, ratingStr, userId||'', date||'', option||'', contentStr, isRefund ? '환불검토' : '', response]);

    // 요약용
    if (!byProduct[productName]) byProduct[productName] = { total: 0, refund: 0 };
    byProduct[productName].total++;
    if (isRefund) byProduct[productName].refund++;
  }

  console.log(`처리 완료:`);
  console.log(`  환불검토: ${refundCount}개`);
  console.log(`  답변생성: ${responseCount}개`);
  console.log(`  답변 길이 초과(100자 이상): ${results.filter(r => r[7].length > 100).length}개`);

  // 엑셀 저장
  const wb = XLSX.utils.book_new();
  const HEADERS = ['제품명', '별점', '아이디', '날짜', '구매옵션', '리뷰내용', '환불', '답변'];
  const COLS = [{ wch: 42 }, { wch: 6 }, { wch: 18 }, { wch: 14 }, { wch: 25 }, { wch: 80 }, { wch: 10 }, { wch: 55 }];

  // 전체리뷰 시트
  const ws1 = XLSX.utils.aoa_to_sheet([HEADERS, ...results]);
  ws1['!cols'] = COLS;
  XLSX.utils.book_append_sheet(wb, ws1, '전체리뷰');

  // 환불검토 시트
  const refundRows = results.filter(r => r[6] === '환불검토');
  if (refundRows.length) {
    const ws2 = XLSX.utils.aoa_to_sheet([HEADERS, ...refundRows]);
    ws2['!cols'] = COLS;
    XLSX.utils.book_append_sheet(wb, ws2, '환불검토');
  }

  // 별점별 샘플 시트 (검토용 - 별점 1~4 전체 + 5점 50개)
  const sampleRows = [
    ...results.filter(r => ['1','2','3','4'].includes(r[1])),
    ...results.filter(r => r[1] === '5').slice(0, 50),
  ];
  const ws3 = XLSX.utils.aoa_to_sheet([HEADERS, ...sampleRows]);
  ws3['!cols'] = COLS;
  XLSX.utils.book_append_sheet(wb, ws3, '검토용샘플');

  // 요약 시트
  const summaryRows = Object.entries(byProduct)
    .sort((a, b) => b[1].total - a[1].total)
    .map(([name, s]) => [name, s.total, s.refund, s.total - s.refund]);
  const ws4 = XLSX.utils.aoa_to_sheet([
    ['제품명', '전체', '환불검토', '답변생성'],
    ...summaryRows,
    [],
    ['합계', data.length, refundCount, responseCount],
  ]);
  ws4['!cols'] = [{ wch: 42 }, { wch: 8 }, { wch: 10 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, ws4, '요약');

  const outPath = path.join(OUTPUT_DIR, `리뷰_AI처리_${timestamp()}.xlsx`);
  XLSX.writeFile(wb, outPath);

  console.log(`\n[저장] ${outPath}`);
  console.log('[시트] 전체리뷰 / 환불검토 / 검토용샘플 / 요약');
  console.log('\n※ 검토용샘플 시트: 별점 1~4점 전체 + 5점 50개');
  console.log('  답변 품질과 환불 판단을 먼저 이 시트에서 확인하세요.\n');
}

main();
