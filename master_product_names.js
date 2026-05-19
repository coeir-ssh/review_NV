/**
 * 네이버 풀네임(검색 노출용 긴 제목) → 마스터테이블 대표 상품명 매핑
 * 매핑 근거: products_knowledge.json 의 name 필드 + 사용자 피드백
 *
 * - getMasterProductName(fullName): 풀네임 → 워드/슬랙에 표기할 짧은 마스터명
 * - inferOption(fullName): 풀네임에 인코딩된 변형(베이직/웨이브, 라운드/스퀘어, 1P 패턴 등)을 옵션으로 추출
 *   AG Grid 에서 optionName 이 비어있을 때 fallback 으로 사용
 */
function getMasterProductName(fullName) {
  const n = (fullName || '').replace(/\s+/g, ' ').trim();
  if (!n) return n;

  // ── 매트 / 발매트 ──────────────────────────────
  if (/빨아쓰는규조토|규조토발매트|규조토.*발매트|워셔블/.test(n)) return '워셔블레더 규조토 발매트';
  if (/욕실매트|미끄럼방지\s*패드|미끄럼방지\s*매트|욕조미끄럼방지|타일바닥|화장실바닥패드/.test(n)) return '욕실 미끄럼방지 매트';

  // ── 욕실화 ────────────────────────────────────
  if (/욕실화|욕실슬리퍼|화장실슬리퍼|실내화/.test(n)) return '발리콘 욕실화';

  // ── 테라조 ────────────────────────────────────
  if (/테라조.*5종/.test(n)) return '테라조 5종 세트\n(디스펜서 · 멀티홀더 · 트레이 · 비누받침대 · 칫솔꽂이)';
  if (/테라조.*3종/.test(n)) return '테라조 3종 세트\n(디스펜서 · 멀티홀더 · 트레이)';
  if (/테라조.*2종/.test(n)) return '테라조 2종 세트\n(비누받침대 · 칫솔꽂이)';
  // 드라이 트레이는 "테라조 [고급 욕실] 드라이 트레이" 처럼 사이에 단어가 있을 수 있음 → 트레이 룰보다 먼저
  if (/테라조.*드라이.*트레이/.test(n))             return '테라조 드라이 트레이';
  if (/테라조\s*(화장솜|면봉|캐니스터)/.test(n))    return '테라조 캐니스터';
  if (/테라조\s*비누받침대/.test(n))                return '테라조 비누받침대'; // 베이직/웨이브 옵션은 inferOption 으로 분리
  if (/테라조\s*칫솔꽂이/.test(n))                  return '테라조 칫솔꽂이';   // 라운드/스퀘어 옵션은 inferOption 으로 분리
  if (/테라조\s*멀티홀더/.test(n))                  return '테라조 멀티홀더';
  if (/테라조\s*디스펜서/.test(n))                  return '테라조 디스펜서';
  if (/테라조.*트레이/.test(n))                     return '테라조 트레이';

  // ── 스테인리스 (스텐) ─────────────────────────
  if (/(스테인리스|스텐).*5종/.test(n)) return '스테인리스 5종 세트\n(칫솔꽂이 · 비누받침대 · 멀티홀더 · 디스펜서 · 트레이)';
  if (/(스테인리스|스텐).*4종/.test(n)) return '스테인리스 4종 세트\n(멀티홀더 · 디스펜서 · 트레이 · 욕실선반)';
  if (/(스테인리스|스텐).*3종/.test(n)) return '스테인리스 3종 세트\n(멀티홀더 · 디스펜서 · 트레이)';
  if (/(스테인리스|스텐).*2종/.test(n)) return '스테인리스 2종 세트\n(칫솔꽂이 · 비누받침대)';
  if (/(스테인리스|스텐)\s*욕실선반.*히든/.test(n)) return '스테인리스 욕실선반 히든';
  if (/(스테인리스|스텐)\s*욕실선반.*플랫/.test(n)) return '스테인리스 욕실선반 플랫';
  if (/(스테인리스|스텐)\s*욕실선반/.test(n))       return '스테인리스 욕실선반 베이직';
  if (/(스테인리스|스텐)\s*비누받침대/.test(n))     return '스테인리스 비누받침대';
  if (/(스테인리스|스텐)\s*칫솔꽂이/.test(n))       return '스테인리스 칫솔꽂이';
  if (/(스테인리스|스텐)\s*멀티홀더/.test(n))       return '스테인리스 멀티홀더';
  if (/(스테인리스|스텐)\s*디스펜서/.test(n))       return '스테인리스 디스펜서';
  if (/(스테인리스|스텐)\s*드라이/.test(n))         return '스테인리스 드라이트레이';
  if (/(스테인리스|스텐).*트레이/.test(n))          return '스테인리스 트레이';

  // ── 타월 / 수건 ───────────────────────────────
  if (/수건|타월/.test(n)) {
    if (/8P/i.test(n)) return '[벌크] 오가닉 에버닌 그린 타월 (8P)';
    if (/4P/i.test(n)) return '[벌크] 오가닉 에버닌 그린 타월 (4P)';
    if (/1P/i.test(n)) return '페이스 타월'; // 패턴(스트라이프/체커드/플로라/블룸)은 inferOption 으로 분리
    return '오가닉 에버닌 그린 타월';
  }

  // ── 샤워기 라인 ───────────────────────────────
  if (/올인원.*샤워|샤워.*올인원/.test(n)) return '올인원 샤워 세트';
  if (/샤워.*듀얼.*필터/.test(n))          return '샤워 듀얼 필터 세트';
  if (/샤워.*스타터/.test(n))              return '샤워 스타터 세트';
  if (/온오프.*샤워|샤워.*온오프/.test(n)) return '샤워기'; // 클라우드 화이트 등 색상은 inferOption 으로 분리
  if (/ACF.*헤드필터|헤드필터/.test(n))    return 'ACF 헤드필터';
  if (/PLA.*필터|샤워기필터.*PLA/.test(n)) return 'PLA 필터';
  if (/샤워기호스|샤워.*호스/.test(n))     return '샤워호스';

  // ── 기타 ──────────────────────────────────────
  if (/부속품/.test(n))    return '부속품 모음';
  if (/기프트.*박스/.test(n)) return '기프트 박스';
  if (/쇼핑백/.test(n))    return '기프트백'; // 쇼핑백 → 기프트백 으로 표기

  return n; // fallback: 원본 그대로
}

/**
 * 풀네임에 인코딩된 변형 옵션을 추출.
 * AG Grid 에서 buyer 옵션을 못 잡았을 때 fallback 으로 사용한다.
 * (실제 buyer 옵션 — 색상/사이즈 등 — 이 잡히면 그것을 우선 사용)
 */
function inferOption(fullName) {
  const n = (fullName || '').replace(/\s+/g, ' ').trim();
  if (!n) return '';

  // 테라조 비누받침대: 베이직 / 웨이브
  if (/테라조\s*비누받침대/.test(n)) {
    if (/웨이브/.test(n)) return '웨이브';
    if (/베이직/.test(n)) return '베이직';
  }
  // 테라조 칫솔꽂이: 라운드 / 스퀘어
  if (/테라조\s*칫솔꽂이/.test(n)) {
    if (/스퀘어/.test(n)) return '스퀘어';
    if (/라운드/.test(n)) return '라운드';
  }
  // 테라조 트레이/드라이트레이: S/M/L
  if (/테라조.*트레이/.test(n)) {
    const m = n.match(/받침\s*(S|M|L)\b/);
    if (m) return m[1];
  }
  // 스테인리스 트레이: 베이직 S/M/L
  if (/스텐|스테인리스/.test(n) && /트레이/.test(n)) {
    const m = n.match(/베이직\s*(S|M|L)/);
    if (m) return `베이직 ${m[1]}`;
  }
  // 스테인리스 욕실선반: 베이직 S/M, 플랫 M/L
  if (/(스텐|스테인리스).*욕실선반/.test(n)) {
    if (/플랫/.test(n))   { const m = n.match(/플랫\s*(M|L)/); return m ? `플랫 ${m[1]}` : '플랫'; }
    if (/히든/.test(n))   return '히든';
    if (/베이직/.test(n)) { const m = n.match(/베이직\s*(S|M)/); return m ? `베이직 ${m[1]}` : '베이직'; }
  }
  // 스테인리스 비누받침대: TypeA/B/C
  if (/(스텐|스테인리스)\s*비누받침대/.test(n)) {
    const m = n.match(/Type\s*([A-C])/i);
    if (m) return `Type ${m[1].toUpperCase()}`;
  }
  // 페이스 타월(1P) 패턴
  if ((/수건|타월/.test(n)) && /1P/i.test(n)) {
    if (/스트라이프/.test(n)) return '스트라이프';
    if (/체커드/.test(n))     return '체커드';
    if (/플로라/.test(n))     return '플로라';
    if (/블룸/.test(n))       return '블룸';
  }
  // 욕실매트 사이즈 (풀네임 끝에 S/M/L)
  if (/욕실매트|미끄럼방지/.test(n)) {
    const m = n.match(/\b(S|M|L)\s*$/);
    if (m) return m[1];
  }
  // 발매트: S/M
  if (/발매트|규조토/.test(n)) {
    const m = n.match(/\b(S|M)\s*$/);
    if (m) return m[1];
  }
  // 욕실화 사이즈는 풀네임에 없음 → AG Grid optionName 의존 (M/L)
  // 샤워기 색상
  if (/온오프.*샤워|샤워.*온오프/.test(n)) {
    if (/클라우드\s*화이트|화이트/.test(n)) return '클라우드 화이트';
    if (/그레이/.test(n))                   return '그레이';
  }
  return '';
}

module.exports = { getMasterProductName, inferOption };
