# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 프로젝트 개요

코에르(COEIR) 욕실용품 브랜드의 **네이버 스마트스토어 리뷰 자동 답변** 시스템입니다. Puppeteer로 sell.smartstore.naver.com(셀러센터)과 brand.naver.com(브랜드스토어)를 제어하고, Anthropic API(Claude Sonnet 4.5)로 답변 생성과 검수를 수행합니다.

모든 사용자 노출 텍스트(로그, Slack 메시지, Word 보고서, Excel 시트)는 한국어입니다.

## 실행 명령어

```bash
# ★ 현재 메인 = 데일리 자동화 오케스트레이터 (스케줄 작업이 이걸 실행)
#   .bat 이 node 단계를 직접(포그라운드) 실행하고, 에이전트는 "판단·답변"만 담당.
#   순서: 수집 → (지식덤프) → 에이전트 판단(replies.json) → 등록 → 검증 → 보고서 → 워치독
daily_review_auto.bat
#   세부 스크립트(개별 실행/디버그용):
node collect_pending.js [--scheduled]   # 답글미등록 수집 → pending_reviews.json (--scheduled=영업일 가드)
node dump_knowledge.js > knowledge_dump.txt   # 제품지식 덤프(에이전트가 읽음)
node post_replies.js                    # replies.json 의 답변 등록 → posted_results.json
node verify_predictions.js              # 전날 판단 사후 검증 → verification_result.json
node generate_report.js                 # Word + Slack 발송 (posted_results.json 기반)
node post_run_check.js                  # 워치독: 미완주/세션만료 시 Slack 알림 + 세션 사전 경고

# 세션 만료 시 수동 로그인 1회 (보이는 브라우저) — 평소엔 자동 로그인이라 거의 불필요
$env:HEADLESS="false"; node seed_login.js

# 레거시 메인 (구 직접 API 방식 — 현재 스케줄에서 미사용, 참조용)
node post_product_reviews.js

# 부가 스크립트
node collect_all_reviews.js        # 전체 제품 리뷰 수집 → xlsx
node collect_reviews.js            # 단일 제품 리뷰 수집
node find_review_position.js       # 특정 리뷰의 랭킹순 순위 단독 조회
node test_review_position.js       # 순위 조회 테스트 하네스
node resend_with_rank.js           # .review_summary.json 읽어 xlsx/docx/Slack 재생성 (재등록 없음)

# 최종 사용자용 .bat 래퍼
#   AI답변생성_실행.bat, 리뷰수집_실행.bat, 전체리뷰수집_실행.bat

# ── 자사몰(카페24) 나쁜 리뷰 모니터링 파이프라인 ──
#   ⚠️ 자사몰은 카페24 답글이 공식몰에 노출 안 됨 → 자동답변 폐기.
#      대신 "나쁜 리뷰 감지 → 슬랙 알림"(자동삭제는 추후).
node cafe24_auth.js                       # 최초 1회 OAuth 인증 (.cafe24_token.json)
node collect_pending_cafe24.js --baseline # 최초 1회 기존 리뷰를 '처리됨' 시드(알림 폭주 방지)
node collect_pending_cafe24.js            # 최근 N일 미처리 신규 리뷰 → pending_reviews_cafe24.json
#   → 에이전트가 삭제대상/유지 판별 → screened_reviews_cafe24.json
node alert_bad_reviews_cafe24.js          # 나쁜 리뷰 슬랙 알림 + cafe24_seen.json 기록
#   래퍼: 자사몰리뷰모니터링_실행.bat (daily_review_cafe24_prompt.txt → claude --print)
#   스케줄: Windows 작업 "코에르_자사몰리뷰모니터링" 매일 07:30 (네이버 코에르_리뷰_에이전트매개 07:00과 분리)
#   (레거시·미사용: post_replies_cafe24.js / generate_report_cafe24.js — 답변 등록 방식)

# 테스트 러너·린터 없음. 문법 검사만:
node -c post_product_reviews.js
```

## 로그인 / 세션 (중요 — 과거 반복 장애 지점)

- **자동 로그인**: `tryAutoLogin()`(post_product_reviews.js)이 셀러센터 홈 → [로그인하기] → `config.js`의 ID/PW 입력으로 로그인. **headless 에서도 사람 개입 없이 동작.**
- **봇 감지 우회 필수**: 모든 puppeteer launch 에 `--disable-blink-features=AutomationControlled` + `ignoreDefaultArgs:['--enable-automation']`, 페이지에 `navigator.webdriver` 숨김. 이게 없으면 네이버가 로그인을 막음.
- **로그인 확인은 반드시 대시보드 주소(`#/home/dashboard`)로**. 루트(`/`)는 네이버가 공개 `/home`(로그아웃 화면)으로 리다이렉트해서 "세션 만료"로 오판함 (며칠간 장애의 실제 원인이었음).
- **세션 저장**: `saveSession(await page.cookies())` — 2개월 검증된 방식. ⚠️ `getAllCookies`(CDP 전 도메인)·`userDataDir`(전용 프로필)은 시도했다가 **오히려 복원이 깨져서 제거**함. 네이버 세션 쿠키는 단명(~1일)이라 매일 자동 로그인으로 갱신하는 게 정상.
- 세션 만료로 자동 로그인까지 실패하면(드묾): `$env:HEADLESS="false"; node seed_login.js` 로 1회 수동 로그인.

## 실행 구조 (왜 .bat 오케스트레이터인가)

- **~5/19**: `node post_product_reviews.js` 직접 실행(단일 포그라운드, API 유료). 2개월 안정.
- **5/20~6/16**: "에이전트 매개"로 전환(claude --print 가 node 실행, 구독 무비용). 그러나 **에이전트가 node 를 백그라운드로 던지고 대기 → --print 세션 종료 시 프로세스 고아·멈춤**으로 자동 실행만 반복 실패(수동은 포그라운드라 정상).
- **6/16~**: `daily_review_auto.bat` 가 node 단계를 **직접(포그라운드)** 실행, 에이전트(`daily_judge_prompt.txt`)는 **판단·답변(replies.json)만** 작성. 직접 실행의 안정성 + 에이전트 무비용 판단을 결합. **에이전트가 node 를 직접 실행하지 않게 하는 것이 핵심.**

### .bat / 스케줄러 함정 (6/17 장애로 확정 — 반드시 지킬 것)
- 🚨 **.bat 안에서 `chcp 65001` 금지**. chcp 로 코드페이지를 바꾸면 cmd 가 .bat 파일 읽는 바이트 위치를 잃어 **이후 라인이 단어 중간에서 잘려 파싱**됨(스케줄러=CP949 환경에서 터짐, 수동 실행=이미 UTF-8이라 무증상). → `daily_review_auto.bat` 는 **ASCII 전용 + chcp 없음**. (로그의 한글이 깨져 보여도 JSON·Slack 은 UTF-8 이라 기능 무관)
- **.bat 줄바꿈은 CRLF**. (Write 도구는 LF 로 저장하므로 저장 후 CRLF 변환 필요)
- **스케줄러 액션은 ASCII 경로로**: 한글 경로(`코에르\클로드`)를 PowerShell `Set-ScheduledTask` 로 넣으면 작업 XML 인코딩이 깨져 0xFF. → ASCII junction 사용: `C:\coeir_review` → 프로젝트 폴더 (`New-Item -ItemType Junction`). 스케줄 액션 = `cmd /c "C:\coeir_review\daily_review_auto.bat"`.
- **스케줄 작업**: `코에르_리뷰_에이전트매개` 매일 07:00, WakeToRun=True. + 별도 `코에르_리뷰_워치독` 08:30 (2차 안전망).

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

### 자사몰(카페24) 나쁜 리뷰 모니터링 (네이버와 완전 별도)

**🚨 자사몰은 네이버와 목적이 다름**: 카페24 답글은 코에르 공식몰에 노출되지 않아 자동답변이 무의미. 그래서 자사몰은 **"최근 리뷰 수집 → 나쁜 리뷰 판별 → 슬랙 알림"** 모니터링 파이프라인이다 (자동 삭제는 추후). 네이버 셀러센터 Puppeteer 크롤링과 달리 카페24는 **Open API(Admin API)** 사용.

- `cafe24_api.js` — API 클라이언트 코어. built-in `https`만 사용(새 의존성 없음).
  - OAuth 2.0(액세스 2시간/리프레시 2주 자동 갱신, `.cafe24_token.json` 캐시). **스코프 = community(게시판)·product 등 전체** (`mall.read_community`/`mall.write_community`/… — 게시판은 `board`가 아니라 `community`).
  - `listBoards`/`listReviewArticles`/`listComments`/`postComment`/`getProductName`/`deleteArticle`/`deleteComment`.
  - 레이트리밋 대응: 모든 요청 직렬화 + 최소 550ms 간격 + 429 지수 백오프 (카페24는 ~2req/s).
  - 게시글 날짜 필터는 **`start_date`/`end_date`** (created_* 아님). 댓글 등록은 `writer`·`password` 필수.
  - 상품후기 게시판 `CAFE24_REVIEW_BOARD_NO`(기본 4 = "리뷰", `listBoards`로 확인됨).
- `cafe24_auth.js` — 최초 1회 OAuth CLI (인증 URL → 리다이렉트 code 붙여넣기 → 토큰 저장).
- `collect_pending_cafe24.js` — 최근 N일(기본 3) 리뷰 중 **미처리(cafe24_seen.json에 없는)** 건만 수집 → `pending_reviews_cafe24.json`. `--baseline` 으로 기존 리뷰를 '처리됨' 시드(최초 1회). `shouldSkipToday` 영업일 가드 재사용.
  - ⚠️ **답변완료 판정은 reply 필드 쓰지 말 것**: 게시글 `reply`('T'/'F')는 "답변글(threaded)"용이라 댓글을 달아도 'F' 유지(실측). 댓글 존재 여부는 `listComments`로만 판정. (단 현재 모니터링 흐름은 댓글 판정 자체를 안 쓰고 seen-set으로 중복관리.)
- 에이전트 판별: `routines/daily_review_cafe24.md`. **별점 1~2점 = 무조건 삭제대상 / 3~5점 = 본문 내용 판단(네이버 환불검토 로직, 순위 분기 없음)**. 출력 `screened_reviews_cafe24.json`(점검한 전체 포함, judgeLabel='삭제대상'|'유지').
- `alert_bad_reviews_cafe24.js` — '삭제대상'을 슬랙으로 알림 + 점검한 전체 article_no를 `cafe24_seen.json`에 기록(재알림 방지). 자동 삭제는 안 함.
- 설정: `config.js` 의 `CAFE24_*` 키(mall_id, client_id/secret, redirect_uri, review_board_no, api_version, comment_writer/password) + `node cafe24_auth.js` 선행.

**레거시(미사용)**: `post_replies_cafe24.js`/`generate_report_cafe24.js` 는 초기 "자동답변+보고서" 방식. 답글 미노출로 폐기됐으나 코드는 참조용으로 남김. `generateWordDoc`/`sendSlackDM`(post_product_reviews.js)의 `summary.channelLabel`/`fileSuffix` 분기도 이때 추가된 것(네이버 동작엔 영향 없음).

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
