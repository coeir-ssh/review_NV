# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

코에르(COEIR) 욕실용품 브랜드의 **네이버 스마트스토어 리뷰 자동 답변** 시스템입니다. Puppeteer로 sell.smartstore.naver.com(셀러센터)과 brand.naver.com(브랜드스토어)를 제어하고, Anthropic API(Claude Sonnet 4.5)로 답변 생성과 검수를 수행합니다.

모든 사용자 노출 텍스트(로그, Slack 메시지, Word 보고서, Excel 시트)는 한국어입니다.

## 실행 명령어

```bash
# 메인 자동화 스크립트 — 오늘의 답글미등록 리뷰에 자동 답변
node post_product_reviews.js

# 부가 스크립트
node collect_all_reviews.js        # 전체 제품 리뷰 수집 → xlsx
node collect_reviews.js            # 단일 제품 리뷰 수집
node find_review_position.js       # 특정 리뷰의 랭킹순 순위 단독 조회
node test_review_position.js       # 순위 조회 테스트 하네스
node resend_with_rank.js           # .review_summary.json 읽어 xlsx/docx/Slack 재생성 (재등록 없음)

# 최종 사용자용 .bat 래퍼
#   AI답변생성_실행.bat, 리뷰수집_실행.bat, 전체리뷰수집_실행.bat

# 테스트 러너·린터 없음. 문법 검사만:
node -c post_product_reviews.js
```

최초 실행 시 headed 브라우저에서 수동 로그인 요청 → 쿠키가 `.seller_session.json`에 저장되어 이후 자동 로그인.

## 민감 정보 (`config.js`)

`config.js`에 셀러 로그인, Anthropic 키, Slack Bot Token, 네이버 API 키의 하드코딩 fallback이 포함되어 있습니다. 환경변수가 있으면 덮어씁니다. **새 비밀값을 커밋하지 말 것**, `config.js` 전체 내용을 로그에 덤프하지 말 것.

## 아키텍처

### 메인 스크립트: `post_product_reviews.js`

약 1,800줄 단일 파일. `main()` 내부 실행 단계:

1. **로그인 / 세션** — `loginToSellerCenter()`가 `.seller_session.json` 쿠키 복원, 실패 시 수동 로그인 폴백.
2. **필터 설정** — 리뷰 검색 페이지에서 `1주일` 날짜 버튼과 `답글미등록` 드롭다운(Selectize.js) 선택.
3. **행 스캔 루프** — AG Grid는 가상 스크롤을 사용. `collectVisibleRows()`는 `.ag-pinned-left-cols-container`(체크박스)와 `.ag-center-cols-container`(데이터)를 `row-index`로 매칭. **체크박스 비활성** 행(환불완료)은 총 개수에서 제외하고 스킵.
4. **리뷰별 파이프라인**:
   - `findReviewPosition()` — 새 탭에서 brand.naver.com의 `/n/v1/contents/reviews/query-pages`를 `REVIEW_RANKING` 정렬로 호출, 본문 스니펫 매칭으로 순위 산출. `product_naver_ids.json`(URL id → `originProductNo`) 매핑 필수.
   - `checkNeedsRefund()` — 1~2점 자동 환불검토, 3점 Claude 판단, 4~5점 스킵.
   - `generateReply()` — 브랜드 시스템 프롬프트 + `getProductKnowledge(productName)`가 제품명 키워드(`샤워기`, `테라조`, `타월` 등)로 분기한 제품별 지식 주입.
   - **`auditReply()` (검수 서브에이전트)** — 동일 제품 지식을 재활용, 10개 체크리스트(시작문 / 100~200자 / 금지 표현 / 제품 USP / 리뷰 본문 반영 / 이모티콘 / 감사 표현 / 별점별 톤 / 욕실화 "가볍다" 금지 / 자연스러움)로 심사. JSON `{passed, score, issues, revisedReply}` 반환. 실패 + 수정본 있으면 **수정본으로 교체 후 재검수 1회** → 수정본 등록. **fail-open**: 파싱/API 오류로는 등록을 절대 막지 않음.
   - `postReply()` — 행 클릭 → 모달 textarea 입력 → 등록 → 검색 재실행으로 그리드 갱신.
5. **리포트 생성** — `generateExcel()` + `generateWordDoc()` + `sendSlackDM()`. Word/Slack은 "답변 완료 / 환불검토 / 실패" 3 블록으로 구성. **리뷰 순위는 모든 행**(제품명 다음 줄)에 표시, **대응 정책은 환불검토 항목에만** 표시.
6. **검증** — 답글미등록 필터로 재검색하여 남은 건수를 기준으로 등록 누락 여부 보고.

`result` 객체에 담긴 검수 필드: `originalReply`, `auditPassed`, `auditScore`, `auditIssues`, `auditRevised`. 이 필드들은 `.review_summary.json` 덤프 → Slack 메시지 → Word 보고서로 전달. **Excel에는 의도적으로 포함하지 않음**.

### 주요 데이터 파일

- `product_naver_ids.json` — brand.naver.com URL id → `{name, originProductNo, checkoutMerchantNo}` 매핑. 순위 조회 필수. `originProductNo: null`(예: 욕실화 `8863030705`)은 순위 조회 불가.
- `products_knowledge.json`, `product_urls.json` — 구형 `generate_*.js` / `collect_*.js` 스크립트용.
- `.seller_session.json` — 로그인 쿠키 저장.
- `.review_summary.json` — 최근 실행 결과 전체 덤프. `resend_with_rank.js`가 재등록 없이 리포트만 재생성할 때 사용.
- `.posting.lock` — `post_product_reviews.js` 중복 실행 방지.
- `reply/YYYYMMDD_reply.{xlsx,docx}` — 일별 리포트 출력.

### 프롬프트에 내장된 브랜드 규칙

`generateReply()`와 `auditReply()` 양쪽 모두에서 동일한 규칙을 강제합니다. **한쪽만 수정하면 반드시 다른 쪽도 같이 수정**할 것:

- `안녕하세요, 고객님.` 으로 시작 (구매자 아이디 사용 금지)
- 100~200자 (한글 기준)
- 금지 표현: `KC인증`(→`KC마크 획득 제품`), `친환경 PVC`, `굴패각 덕분에 항균`, `욕실문에 걸리지 않는`, `욕실용으로 딱 맞게`, TPE를 `젖병 젖꼭지 소재`로 표현
- 제품별 핵심 USP:
  - PLA 샤워기·필터 = **미세플라스틱 미검출** (최우선)
  - 타월 = `슈퍼 코마사` (그냥 `코마사` 아님)
  - 욕실화 = 묵직한 무게 = 안전·미끄럼방지 포지션 (**절대 "가볍다" 칭찬 금지**)
  - 발매트 = `워셔블 레더 규조토`
  - 스테인리스 = `SUS304`
- 이모티콘 1~2개, `감사합니다` 또는 `고맙습니다` 필수
- 별점 ≤3점 → 사과 먼저 / 4~5점 → 공감·감사

### 기타 스크립트 (레거시·일회성, 메인 흐름 아님)

- `collect_all_reviews.js`, `collect_reviews.js`, `auto_review_response.js` — 브랜드 공개 API를 직접 쓰던 초기 수집 파이프라인.
- `generate_local.js`, `generate_opus.js`, `generate_top5.js`, `generate_responses.js`, `make_excel.js` — 초기 답변 생성 실험(로컬 규칙, Opus 모델, TOP5 테스트 등). 현재 흐름과 연결 없음.
- `post_to_naver.js`, `post_to_naver_puppeteer.js` — 초기 등록 실험.
- `debug_filter.js` — UI 디버깅 헬퍼.

동작을 확장할 때는 `post_product_reviews.js`를 수정할 것. 레거시 스크립트는 참조용으로만 남겨둠.

## 이 저장소 특화 편집 팁

- `post_product_reviews.js`의 기본값은 `headless: true`. 그리드 인터랙션 디버깅 시 `false`로 전환.
- AG Grid는 가상 스크롤이라 보이는 행만 DOM에 존재함. 행 수집은 반드시 pinned-left(체크박스) + center(데이터)를 `row-index`로 매칭하는 패턴 유지.
- Slack/Word 출력 구조를 변경할 때는 `resend_with_rank.js`도 같이 수정 — summary → 리포트 매핑이 거기서 동일하게 재현되기 때문.
- 네이버 brand API 호출은 올바른 `checkoutMerchantNo` 헤더가 필수. 새로 요청을 만들지 말고 기존 `page.evaluate(fetch(...))` 패턴을 복사해서 쓸 것.
