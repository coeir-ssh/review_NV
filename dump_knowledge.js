/**
 * dump_knowledge.js — 전 제품 지식·룰을 한 번에 출력 (Step 0 덤프).
 * .bat 이 실행해 knowledge_dump.txt 로 저장 → 에이전트가 읽어 답변 품질 일관성 확보.
 * (에이전트가 node 를 직접 안 돌려도 되도록 분리)
 *
 * 사용:  node dump_knowledge.js > knowledge_dump.txt
 */
const m = require('./post_product_reviews');
const products = [
  '욕실 미끄럼방지 매트', '워셔블레더 규조토 발매트', '발리콘 욕실화',
  '스테인리스 비누받침대', '스테인리스 칫솔꽂이', '스테인리스 멀티홀더',
  '스테인리스 디스펜서', '스테인리스 트레이', '스테인리스 욕실선반 베이직',
  '테라조 비누받침대', '테라조 칫솔꽂이', '테라조 멀티홀더',
  '테라조 디스펜서', '테라조 트레이', '테라조 드라이 트레이', '테라조 캐니스터',
  '페이스 타월', 'PLA 필터 온오프 샤워기', 'PLA 필터', 'ACF 헤드필터', '샤워호스',
  '샤워 듀얼 필터 세트', '샤워 스타터 세트', '올인원 샤워 세트',
];
for (const p of products) {
  const k = m.getProductKnowledge(p);
  if (!k) continue;
  console.log('======== ' + p + ' ========');
  console.log(k);
  console.log();
}
