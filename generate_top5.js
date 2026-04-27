/**
 * 코에르 리뷰 답변 생성기 - 제품별 상위 5개 테스트
 *
 * 구조:
 *   Part1: 공통 인사 (고정)
 *   Part2: 리뷰 내용 기반 개인화 (100자 이상, 각 리뷰마다 다름)
 *   Closing: 고정 마무리
 */

const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const DIR = 'C:\\Users\\AWESOMATIC\\Desktop\\코에르\\클로드\\리뷰수집프로그램';
const PART1 = '안녕하세요, 고객님. 저희 제품을 구매해주시고 리뷰 남겨주셔서 감사합니다 :)';
const CLOSING = '리뷰 감사드리며, 앞으로도 코에르 제품 많이 사랑해주세요. 좋은 하루 보내세요!';

// ────────────────────────────────────────────────────────────────
// 유틸
// ────────────────────────────────────────────────────────────────
function timestamp() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}

function findLatestExcel() {
  const files = fs.readdirSync(DIR)
    .filter(f => f.startsWith('리뷰_코에르_전체_') && f.endsWith('.xlsx'))
    .sort().reverse();
  return files.length ? path.join(DIR, files[0]) : null;
}

/** 리뷰 내용 기반 결정론적 번호 */
function seed(content) {
  const c = (content || '').replace(/\s/g, '');
  let s = 0;
  for (let i = 0; i < Math.min(12, c.length); i++) s += c.charCodeAt(i) * (i + 5);
  return Math.abs(s);
}

function pick(arr, s) { return arr[s % arr.length]; }

// ────────────────────────────────────────────────────────────────
// 환불 판단
// ────────────────────────────────────────────────────────────────
function judgeRefund(rating, content) {
  const c = content || '';
  const r = Number(rating) || 5;
  const NEG_CTX = ['파손 없', '파손되지', '파손없', '불량 없', '불량없', '이상 없', '이상없', '잘 도착'];
  if (NEG_CTX.some(kw => c.includes(kw))) return false;
  const HARD = ['파손됐', '파손되었', '부서졌', '깨졌', '깨져서', '금이 갔', '부러졌', '불량품', '작동이 안', '전혀 안 됩', '먹통'];
  if (HARD.some(kw => c.includes(kw))) return true;
  if (r <= 2) {
    if (['배송만', '포장만', '별점 실수', '제품은 좋', '제품이 좋'].some(kw => c.includes(kw))) return false;
    const pos = ['예쁘', '만족', '좋아요', '좋습니다', '완벽', '최고'].filter(kw => c.includes(kw)).length;
    const neg = ['별로', '아쉬', '실망', '나빠', '나쁘'].filter(kw => c.includes(kw)).length;
    if (pos >= 2 && neg === 0) return false;
    return true;
  }
  if (r === 3) return ['안 됩니다', '안됩니다', '작동 안', '불량', '고장', '제대로 안'].some(kw => c.includes(kw));
  return HARD.some(kw => c.includes(kw));
}

// ────────────────────────────────────────────────────────────────
// 리뷰 내용을 분석해 개인화 문장들 생성
// ────────────────────────────────────────────────────────────────
function buildPersonalized(content, rating) {
  const c = content || '';
  const r = Number(rating) || 5;
  const pts = []; // 감지된 포인트별 문장

  // ── 저별점 / 아쉬움 먼저 처리
  if (r <= 2) {
    pts.push(
      '기대하셨던 만큼의 만족을 드리지 못한 것 같아 정말 죄송합니다 🙏 ' +
      '남겨주신 소중한 의견은 코에르가 더 나은 제품과 서비스를 만들어 나가는 데 소중히 반영하겠습니다.'
    );
  }

  // ── 패키지/언박싱 고급스러움
  if (/패키지|박스.{0,5}(예쁘|고급|깔끔)|포장.{0,5}(예쁘|고급|깔끔|좋)/.test(c)) {
    pts.push(
      '패키지부터 제품까지 고급스럽게 느껴주셔서 코에르 팀 모두 정말 기뻐요 😊 ' +
      '언박싱 순간부터 특별하게 느껴지도록 세심하게 준비하고 있는데, 그 마음이 잘 전달된 것 같아 뿌듯합니다.'
    );
  } else if (/고급스러워|고급스럽|고급진|럭셔리/.test(c)) {
    pts.push(
      '고급스럽다고 느껴주셔서 코에르 팀 모두 정말 기쁩니다 😊 ' +
      '디자인부터 소재까지 품격 있는 욕실 제품을 만들기 위해 늘 노력하고 있는데, 알아봐 주셔서 감사해요.'
    );
  }

  // ── 색상 + 욕실 어울림
  const colorTests = [
    [/화이트.{0,20}(어울|예쁘|좋|잘 맞|딱)/, '화이트',
      '화이트 컬러가 욕실 인테리어에 자연스럽게 어울리신다니 정말 다행이에요! ' +
      '밝고 깔끔한 느낌이 욕실 전체 분위기를 환하게 만들어드렸으면 좋겠습니다 🤍'],
    [/그레이.{0,20}(어울|예쁘|좋|잘 맞|딱)/, '그레이',
      '그레이 컬러가 욕실에 모던하게 어울리신다니 기쁩니다! ' +
      '차분하면서도 세련된 분위기가 욕실을 한층 고급스럽게 만들어드렸으면 해요 🤍'],
    [/베이지.{0,20}(어울|예쁘|좋|잘 맞|딱)/, '베이지',
      '베이지 컬러가 욕실에 따뜻하게 어울리신다니 반갑습니다! ' +
      '편안하고 내추럴한 분위기가 욕실을 더 아늑하게 만들어드렸으면 해요 🤍'],
    [/(색상|색깔|컬러).{0,20}(어울|예쁘|좋|잘 맞|딱)/, '색상',
      '색상이 욕실 인테리어와 잘 어울리신다니 정말 다행이에요! ' +
      '욕실 분위기에 잘 녹아들 수 있도록 신경 써서 개발한 컬러라 더욱 보람차요 🤍'],
  ];
  for (const [re, , txt] of colorTests) {
    if (re.test(c)) { pts.push(txt); break; }
  }

  // ── 인테리어 변화 (색상 언급 없을 때)
  if (/인테리어.{0,20}(좋|예쁘|어울|꾸미|바꾸|달라|업그레이드)/.test(c) && !pts.some(s => s.includes('어울'))) {
    pts.push(
      '욕실 인테리어에 코에르가 함께해서 분위기가 달라지셨다니 정말 기쁩니다! ' +
      '욕실도 생활 공간이니까, 눈에 들어올 때마다 기분 좋아지는 공간이 되셨으면 해요 😊'
    );
  }

  // ── 필터 + 아이 + 구축 조합 (가장 구체적인 경우)
  const hasFilter = /필터|정수|불순물|녹물/.test(c);
  const hasKids = /아이|어린이|아기|자녀|아들|딸/.test(c);
  const hasOldBldg = /구축|오래된 집|오래된 아파트/.test(c);
  if (hasFilter) {
    if (hasKids && hasOldBldg) {
      pts.push(
        '구축 주택에서 아이를 키우시면서 수질 걱정이 많으셨을 텐데, 코에르 정수 필터가 그 걱정을 덜어드릴 수 있어 정말 보람차요 💧 ' +
        '아이 건강까지 함께 지킬 수 있도록 코에르가 늘 노력하겠습니다!'
      );
    } else if (hasKids) {
      pts.push(
        '아이를 키우시는 가정에 코에르 샤워 필터가 도움이 됐으면 해서 더욱 기쁩니다 💧 ' +
        '피부가 민감한 아이에게도 깨끗한 물로 안심하고 샤워시킬 수 있도록 코에르가 함께할게요!'
      );
    } else if (hasOldBldg) {
      pts.push(
        '구축 주택에서는 수질 걱정이 생길 수 있는데, 코에르 필터 덕분에 깨끗한 물로 안심하고 샤워하실 수 있겠죠 💧 ' +
        '이중 필터로 보이지 않는 불순물까지 걸러드릴게요!'
      );
    } else {
      pts.push(
        '정수 필터 덕분에 불순물 걱정 없이 깨끗하고 건강한 물로 샤워하실 수 있겠죠 💧 ' +
        '코에르가 욕실 위생까지 꼼꼼하게 챙겨드리겠습니다!'
      );
    }
  }

  // ── 수압
  if (/수압.{0,15}(좋|세|강|시원|만족|굿|올라|높아)|물줄기.{0,15}(좋|세|강|시원)/.test(c)) {
    pts.push(
      '수압도 마음에 드셨다니 정말 다행이에요 🚿 ' +
      '기존 샤워기보다 확실히 시원하게 느껴지도록 설계했는데, 그 부분이 만족스러우셨다니 코에르 팀 모두 기쁩니다!'
    );
  }

  // ── 그립감
  if (/그립감|잡기.{0,5}(편|쉽)|손에 잘 잡/.test(c)) {
    pts.push(
      '그립감도 좋으셨군요 😊 ' +
      '샤워기를 오래 잡고 있어도 미끄러지지 않도록 표면 처리에 신경 쓴 부분인데, 불편함 없이 사용하시고 계신 것 같아 다행이에요!'
    );
  }

  // ── 절수 기능
  if (/절수|물 절약|온오프.{0,10}(좋|편|절약)/.test(c)) {
    pts.push(
      '절수 기능으로 물도 아끼고 환경까지 함께 지킬 수 있어서 더욱 의미있는 제품이죠 😊 ' +
      '알뜰하게 사용하시면서 코에르의 온오프 버튼이 매일 든든한 역할을 해드리길 바랍니다!'
    );
  }

  // ── 무타공 + 전세/임대
  const hasNoHole = /무타공|타공 없이|구멍 없이|구멍 뚫지/.test(c);
  const hasRent = /전세|월세|임대|이사/.test(c);
  if (hasNoHole) {
    if (hasRent) {
      pts.push(
        '전세집에서도 벽에 구멍 뚫지 않고 깔끔하게 설치하실 수 있어서 정말 다행이에요 😊 ' +
        '이사 후에도 함께 가져가실 수 있으니 오래오래 코에르와 함께하세요!'
      );
    } else {
      pts.push(
        '무타공으로 벽 손상 없이 깔끔하게 설치하셨군요 😊 ' +
        '나중에 위치를 바꾸거나 이사하실 때도 부담 없이 사용하실 수 있어서 더욱 편리한 제품이에요!'
      );
    }
  }

  // ── 설치 간편 (무타공 언급 없을 때)
  if (/설치.{0,10}(쉽|편|간편|간단|빠르|금방|1분)/.test(c) && !pts.some(s => s.includes('설치') || s.includes('타공'))) {
    pts.push(
      '설치가 간편하셨다니 다행이에요 😊 ' +
      '공구 없이도 손쉽게 달 수 있도록 설계한 제품인데, 바로 사용하실 수 있어서 더 편리하게 느껴지셨겠어요!'
    );
  }

  // ── 흡착/고정력
  if (/흡착.{0,10}(좋|강|잘)|잘 붙|떨어지지 않/.test(c)) {
    pts.push(
      '흡착력이 확실해서 안심하고 사용하실 수 있겠죠 😊 ' +
      '시간이 지나도 떨어지지 않도록 접착 방식을 꼼꼼하게 설계했는데, 튼튼하게 잘 버텨주고 있어서 다행이에요!'
    );
  }

  // ── 무게감 / 안정감
  if (/묵직|무게감|안정.{0,5}(있|감)|흔들리지|움직이지 않/.test(c)) {
    pts.push(
      '묵직하고 안정감 있다고 느껴주셨군요 😊 ' +
      '가볍고 플라스틱 느낌이 나지 않도록 소재와 무게에 신경을 많이 쓴 부분인데, 그 차이를 알아봐 주셔서 감사합니다!'
    );
  }

  // ── 사이즈 딱 맞음
  if (/사이즈.{0,8}(잘|딱|맞|좋|적당)|크기.{0,8}(딱|맞|좋|적당)/.test(c)) {
    pts.push(
      '사이즈가 딱 맞으셨다니 정말 기쁩니다 😊 ' +
      '다양한 욕실 환경에 잘 맞도록 크기를 고민해서 만든 제품인데, 고객님 공간에 잘 어울리셔서 다행이에요!'
    );
  }

  // ── 가성비 / 값어치
  if (/가성비|값어치|돈값|가격 대비|이 가격에|이 값에/.test(c)) {
    pts.push(
      '가성비가 좋다고 느껴주셔서 정말 감사합니다 😊 ' +
      '합리적인 가격에 좋은 품질을 드리기 위해 코에르가 늘 고민하고 있는데, 그 노력이 전달되어서 기쁩니다!'
    );
  }

  // ── 스텐/내구성/녹 방지
  if (/스텐|스테인리스|녹.{0,8}(없|안|걱정|강)|부식/.test(c)) {
    pts.push(
      '스텐 소재라 습기 많은 욕실에서도 녹이나 변색 걱정 없이 오래 사용하실 수 있어요 😊 ' +
      '욕실은 특히 습기와 물에 노출이 많은 공간이라 소재 선택에 더 신경을 쓴 부분이랍니다!'
    );
  }

  // ── 규조토 매트 + 건조
  if (/규조토/.test(c)) {
    if (/건조|잘 마|빨리 마|금방 마/.test(c)) {
      pts.push(
        '규조토 특유의 빠른 건조 덕분에 욕실이 항상 쾌적하고 위생적으로 유지될 거예요 🌿 ' +
        '물기가 바닥에 오래 남지 않아서 욕실 위생 관리도 훨씬 편리해지셨을 것 같아 기쁩니다!'
      );
    } else {
      pts.push(
        '규조토 매트로 발을 딛는 순간부터 편안하고 위생적인 욕실 생활을 즐기실 수 있겠죠 🌿 ' +
        '자연 소재라 안심하고 사용하실 수 있다는 것도 코에르 규조토 매트의 큰 장점이에요!'
      );
    }
  }

  // ── 미끄럼 방지
  if (/미끄럼|미끄러짐/.test(c)) {
    pts.push(
      '미끄럼 없이 안전하게 사용하실 수 있어서 저희도 마음이 놓여요 😊 ' +
      '특히 물기 있는 욕실에서 안전은 정말 중요하니까, 가족 모두 안심하고 사용하세요!'
    );
  }

  // ── 수건 흡수력
  if (/흡수|흡수력/.test(c) && /수건|타올|타월/.test(c)) {
    pts.push(
      '흡수력이 좋은 수건이라 샤워 후 물기도 빠르게 닦이고 매번 포근하게 사용하실 수 있겠죠 🤍 ' +
      '코에르 타올은 세탁을 반복해도 촉감이 유지되도록 신경 써서 만들었어요!'
    );
  }

  // ── 호텔 타올 느낌
  if (/호텔/.test(c)) {
    pts.push(
      '호텔 타올 같은 느낌이 나신다니 정말 기뻐요 🤍 ' +
      '40수 코마사 원단으로 만든 코에르 타올이라 가능한 퀄리티인데, 집에서도 매일 호캉스 기분 즐기세요!'
    );
  }

  // ── 포근함/부드러움
  if (/포근|부드러워|부드러운 감촉|폭신|푹신/.test(c)) {
    pts.push(
      '포근하고 부드러운 감촉에 만족해 주셔서 감사해요 💕 ' +
      '샤워 후에 포근하게 감싸주는 느낌이 매일 욕실 시간을 더 기분 좋게 만들어드렸으면 좋겠어요!'
    );
  }

  // ── 공간 활용 / 수납
  if (/공간.{0,10}(활용|넉넉|정리|수납)|수납.{0,10}(좋|편|넉넉)|정리가 잘/.test(c)) {
    pts.push(
      '좁은 욕실 공간도 알차게 활용하실 수 있어서 정말 다행이에요 😊 ' +
      '깔끔하게 정리된 욕실을 볼 때마다 코에르와 함께한 선택이 만족스러우셨으면 좋겠습니다!'
    );
  }

  // ── 재구매
  if (/재구매|또 구매|다시 구매|재주문/.test(c)) {
    pts.push(
      '재구매까지 해주셨다니 진심으로 감사드려요 😊 ' +
      '처음 구매하신 후 만족하셔서 다시 찾아주신 거잖아요, 그 신뢰에 변함없는 품질로 항상 보답하겠습니다!'
    );
  }

  // ── 실제 선물 구매 (추천 표현 제외)
  //    "선물로 샀어요", "선물 드렸어요", "선물해줬어요" 등만 감지
  if (/선물(로|로도)?\s*(샀|구매했|줬어|드렸|보냈|해줬|할게)/.test(c) || /선물\s*(드렸|줬어)/.test(c)) {
    pts.push(
      '소중한 분께 선물해 주셨군요 🎁 ' +
      '받으신 분도 분명 기뻐하셨을 거예요, 특별한 마음을 담은 선물에 코에르가 함께할 수 있어서 저희도 기쁩니다!'
    );
  }

  // ── 고민 끝 구매
  if (/고민.{0,20}(했|끝에|하다가)|망설.{0,10}(였|이다가)/.test(c)) {
    pts.push(
      '고민 끝에 코에르를 선택해 주셔서 감사합니다 😊 ' +
      '믿고 선택해 주신 만큼 후회 없는 제품이 됐으면 좋겠고, 오래오래 만족스럽게 사용하시길 바랍니다!'
    );
  }

  // ── 배송 빠름
  if (/빠른 배송|빨리 왔|당일|새벽배송/.test(c)) {
    pts.push(
      '빠르게 받아보셨군요 📦 기다리지 않고 바로 사용하실 수 있어서 더 좋으셨겠어요! ' +
      '앞으로도 빠르고 안전하게 배송될 수 있도록 노력하겠습니다.'
    );
  }

  // ── 다양한 용도 활용 (구체적 언급)
  if (/면봉|칫솔|클렌징|로션|향수|샴푸|바디워시/.test(c)) {
    pts.push(
      '다양한 용도로 활용해 주시니 정말 기쁩니다 😊 ' +
      '코에르 제품이 욕실 생활 곳곳에서 편리하게 사용되고 있다니, 더 실용적인 제품으로 보답하겠습니다!'
    );
  }

  // ── 문의/AS 응대 칭찬
  if (/문의|AS|애프터|답변.{0,10}(빠르|친절)|응대/.test(c)) {
    pts.push(
      '문의에 빠르게 응해드렸다니 다행이에요 😊 ' +
      '제품 구매 후에도 불편함이 없으시도록 코에르가 늘 옆에서 도와드리겠습니다!'
    );
  }

  // ── 디자인 (다른 언급 없을 때)
  const hasDesign = /디자인.{0,15}(예쁘|좋|깔끔|심플|마음|예)|예뻐요|예쁘네요|예쁩니다/.test(c);
  if (hasDesign && !pts.some(s => s.includes('어울') || s.includes('고급') || s.includes('인테리어'))) {
    pts.push(
      '깔끔하고 예쁜 디자인에 만족해 주셔서 감사합니다 🤍 ' +
      '보기에도 좋고 사용하기에도 편리한 제품을 만들기 위해 디자인부터 기능까지 꼼꼼하게 고민했어요!'
    );
  }

  // ── 퀄리티/소재 (기타 언급 없을 때)
  if (/퀄리티|질감|소재가|재질이|두께/.test(c) && !pts.some(s => s.includes('소재') || s.includes('고급') || s.includes('스텐'))) {
    pts.push(
      '제품 소재와 퀄리티까지 꼼꼼하게 느껴봐 주셔서 감사해요 😊 ' +
      '오래 사용하실수록 코에르 제품의 내구성과 품질이 더욱 믿음직하게 느껴지실 거예요!'
    );
  }

  // 너무 많은 포인트는 앞에서 3개로 제한 (응답이 너무 길어지지 않도록)
  const MAX_PTS = 3;
  const usedPts = pts.slice(0, MAX_PTS);

  // ────────────────────────────────────────────────────────────
  // 결과 조합
  // ────────────────────────────────────────────────────────────
  let result = '';

  if (usedPts.length === 0) {
    // 완전 fallback - 적어도 따뜻하고 100자 이상
    const s = seed(c);
    result = pick([
      '고객님의 따뜻한 리뷰 덕분에 코에르 팀 모두 오늘 하루 힘이 났어요 😊 앞으로도 고객님이 매일 욕실에서 만족하실 수 있도록 더 좋은 제품과 서비스로 보답하겠습니다! 코에르와 함께 더욱 쾌적한 욕실 생활 되세요 🤍',
      '소중한 경험을 공유해 주셔서 정말 감사합니다 🤍 코에르가 고객님의 욕실을 더욱 편안하고 아름답게 만들어 드릴 수 있도록 끊임없이 노력하겠습니다. 앞으로도 코에르와 함께하는 욕실 생활이 즐거우시길 바랍니다!',
      '코에르 제품을 선택해 주셔서 진심으로 감사드려요 😊 고객님 한 분 한 분의 만족이 저희가 더 좋은 제품을 만들어 나가는 가장 큰 원동력이랍니다! 앞으로도 코에르가 늘 든든하게 함께하겠습니다 🤍',
      '따뜻한 리뷰를 남겨주셔서 코에르 팀이 큰 힘을 얻었어요 🤍 고객님의 욕실 생활이 코에르와 함께 더욱 편안하고 즐거워지시길 바라며, 앞으로도 좋은 품질의 제품으로 항상 보답하겠습니다 😊',
    ], s);
  } else {
    result = usedPts.join(' ');
    // 100자 미만이면 보완 (pts가 하나이고 짧은 경우 등)
    if (result.replace(/[\uD800-\uDFFF]/g, 'X').length < 100) {
      const s = seed(c);
      const extras = [
        ' 코에르와 함께 더욱 쾌적하고 편안한 욕실 생활 즐기세요 🤍',
        ' 앞으로도 오래오래 만족스럽게 사용하시길 바랍니다 😊',
        ' 고객님의 소중한 만족이 코에르의 가장 큰 보람이에요 🤍',
        ' 더 좋은 제품으로 항상 보답할게요, 감사합니다 😊',
        ' 코에르가 늘 곁에서 더 나은 욕실 경험을 드릴게요 🤍',
      ];
      result += extras[s % extras.length];
    }
  }

  return result;
}

// ────────────────────────────────────────────────────────────────
// 전체 답변 조합
// ────────────────────────────────────────────────────────────────
function generateResponse(rating, content) {
  const r = Number(rating) || 5;
  if (judgeRefund(r, content)) return null;
  const part2 = buildPersonalized(content, r);
  return `${PART1}\n${part2}\n${CLOSING}`;
}

// ────────────────────────────────────────────────────────────────
// 메인
// ────────────────────────────────────────────────────────────────
function main() {
  const srcFile = findLatestExcel();
  if (!srcFile) {
    console.error('[오류] 리뷰 파일이 없습니다. 먼저 전체리뷰수집_실행.bat을 실행하세요.');
    return;
  }
  console.log(`[원본] ${path.basename(srcFile)}`);

  const wb = XLSX.readFile(srcFile);
  const ws = wb.Sheets['전체리뷰'];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const data = rows.slice(1).filter(r => r[0] || r[5]);

  const byProduct = {};
  for (const row of data) {
    const name = String(row[0] || '기타');
    if (!byProduct[name]) byProduct[name] = [];
    byProduct[name].push(row);
  }

  const top5Rows = [];
  for (const [, reviews] of Object.entries(byProduct)) top5Rows.push(...reviews.slice(0, 5));

  console.log(`[선택] ${Object.keys(byProduct).length}개 제품 × 상위 5개 = ${top5Rows.length}개 리뷰`);

  let refundCount = 0;
  let responseCount = 0;
  let under100 = 0;

  const resultRows = top5Rows.map(row => {
    const rating = row[1];
    const content = String(row[5] || '');
    const response = generateResponse(rating, content);

    if (response === null) {
      refundCount++;
      return [...row.slice(0, 6), '환불검토', ''];
    } else {
      responseCount++;
      const part2 = response.split('\n')[1] || '';
      const len = part2.replace(/[\uD800-\uDFFF]/g, 'X').length;
      if (len < 100) under100++;
      return [...row.slice(0, 6), '', response];
    }
  });

  // 엑셀 저장
  const HEADERS = ['제품명', '별점', '아이디', '날짜', '구매옵션', '리뷰내용', '환불', '답변'];
  const COL_WIDTHS = [
    { wch: 42 }, { wch: 6 }, { wch: 18 }, { wch: 14 },
    { wch: 25 }, { wch: 80 }, { wch: 10 }, { wch: 100 },
  ];

  const newWb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.aoa_to_sheet([HEADERS, ...resultRows]);
  ws1['!cols'] = COL_WIDTHS;
  ws1['!rows'] = resultRows.map(() => ({ hpt: 100 }));
  XLSX.utils.book_append_sheet(newWb, ws1, '상위5개리뷰');

  const refundRows = resultRows.filter(r => r[6] === '환불검토');
  if (refundRows.length > 0) {
    const ws2 = XLSX.utils.aoa_to_sheet([HEADERS, ...refundRows]);
    ws2['!cols'] = COL_WIDTHS;
    XLSX.utils.book_append_sheet(newWb, ws2, '환불검토');
  }

  const outPath = path.join(DIR, `리뷰_TOP5테스트_${timestamp()}.xlsx`);
  XLSX.writeFile(newWb, outPath);

  console.log(`\n[결과]`);
  console.log(`  총 ${top5Rows.length}개 리뷰`);
  console.log(`  환불검토: ${refundCount}개`);
  console.log(`  답변 생성: ${responseCount}개`);
  console.log(`  Part2 100자 미만: ${under100}개`);
  console.log(`\n[저장] ${path.basename(outPath)}`);
}

main();
