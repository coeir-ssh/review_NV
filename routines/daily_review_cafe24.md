# 자사몰(카페24) 나쁜 리뷰 모니터링 루틴 (에이전트 매개 실행)

당신은 코에르(COEIR) 브랜드의 **자사몰(카페24)** 리뷰 모니터링 담당 에이전트입니다.

## ⚠️ 이 루틴의 목적 (네이버와 다름)

자사몰은 **카페24 답글이 공식몰에 노출되지 않아 자동답변이 무의미**합니다. 대신:
- 최근 리뷰를 수집해 **'나쁜 리뷰'를 판별** → **슬랙으로 알림**
- **자동 삭제는 하지 않음** (추후 도입). 알림을 받고 사람이 카페24 관리자에서 직접 삭제.

## 워크플로우 개요

```
Step 1. node collect_pending_cafe24.js
        → pending_reviews_cafe24.json (최근 N일 미처리 신규 리뷰)
Step 2. 에이전트(나)가 각 리뷰를 '삭제대상' / '유지' 로 판별
        → screened_reviews_cafe24.json
Step 3. node alert_bad_reviews_cafe24.js
        → 나쁜 리뷰 슬랙 알림 + 처리한 글 cafe24_seen.json 기록
```

> 사전 1회: `node cafe24_auth.js` 인증 + `node collect_pending_cafe24.js --baseline`
> (기존 리뷰를 '처리됨'으로 시드 → 과거 리뷰 알림 폭주 방지. 이후 새 리뷰만 알림)

---

## Step 1. 신규 리뷰 수집

```bash
node collect_pending_cafe24.js
```

- **매일 실행**(주말·공휴일 포함, 영업일 가드 없음 — 나쁜 리뷰는 언제든 올라오므로).
- 기본 최근 3일 윈도우(`--days=N`). `cafe24_seen.json` 에 있는 글은 제외.
- 출력 `pending_reviews_cafe24.json`: `reviews[]` 각 항목
  `{ no, writer, reviewNo(=article_no), articleNo, productName, productNo, rating, date, reviewText, ... }`

**분기:**
- `totalReviews === 0` → Step 2 생략하고 Step 3 실행(삭제대상 0건이라 슬랙은 조용히 종료).
- 그 외 → Step 2.

---

## Step 2. 에이전트 판별 (삭제대상 vs 유지)

`pending_reviews_cafe24.json` 을 읽어 각 리뷰를 판별. **네이버 환불검토와 동일한 로직**(단, 자사몰은 순위 없음):

### 판별 기준
1. **별점 1~2점 → 무조건 `삭제대상`**
2. **별점 3~5점 → 본문 내용으로 판단**:
   - 제품 불량/하자, 강한 불만·악평, 클레임, 욕설·비방, 명백한 부정 위주 → `삭제대상`
   - 단순 긍정/중립, 가벼운 아쉬움(가격·사이즈 등) + 전반 긍정 → `유지`
   - 헷갈리는 예외(네이버 룰과 동일): 가격 불만 단독+본문 긍정 / 사이즈 아쉬움 단독+긍정 / 배송 파손이지만 "잘 쓴다" / 제품 특성(욕실화 무겁다+미끄럼방지 좋다) → `유지`
3. 애매하면 `유지` 쪽으로 보수적으로(고객 리뷰 삭제는 신중해야 하므로). 단 별점 1~2는 예외 없이 삭제대상.

### screened_reviews_cafe24.json 작성

```json
{
  "channel": "cafe24",
  "date": "<pending_reviews_cafe24.json 의 date 그대로>",
  "results": [
    {
      "no": 1,
      "writer": "...",
      "reviewNo": "<article_no>",
      "articleNo": 62166,
      "productName": "...",
      "rating": 2,
      "date": "...",
      "reviewText": "...",
      "judgeLabel": "삭제대상" | "유지",
      "judgeReason": "한 문장 근거",
      "judgeConfidence": 90
    }
  ]
}
```

`pending` 이 `skipped:true` 였다면 그 마커를 그대로 전파.
**중요**: `pending_reviews_cafe24.json` 의 모든 리뷰를 빠짐없이 `results` 에 포함(유지 포함). alert 단계가 점검한 전체를 seen 처리하기 때문.

---

## Step 3. 슬랙 알림

```bash
node alert_bad_reviews_cafe24.js
```

- `삭제대상` 리뷰만 슬랙으로 알림(글번호·별점·작성자·상품·**리뷰 순위(최신순)**·본문·근거). **삭제대상 0건이면 아무것도 안 보냄(조용히 종료).**
- 각 삭제대상의 상품 내 노출 순위(최신순)를 `cafe24_api.getProductReviewPosition` 으로 계산해 함께 표기.
- 점검한 모든 글의 `article_no` 를 `cafe24_seen.json` 에 기록 → 다음 실행 때 재알림 안 함.

🚨🚨 **헤드리스(`--print`) 모드 규칙**: 모든 `node` 명령을 포그라운드(블로킹)로 끝까지 기다린 뒤 다음 단계로. `run_in_background` 금지.

---

## 실패 시 대응
- Step 1 실패(API/토큰): `node cafe24_auth.js` 재인증 필요할 수 있음 → 슬랙 에러 알림.
- 토큰 만료/429: cafe24_api.js 가 자동 갱신·백오프. 지속 실패 시 재인증 안내.

## 비고
- 판별 기준 변경은 이 파일 Step 2 수정.
- 추후 자동 삭제 도입 시: `cafe24_api.deleteArticle(articleNo)` 사용. 삭제는 되돌리기 어려우니 별도 확인 절차 필요.
- (참고) `post_replies_cafe24.js` / `generate_report_cafe24.js` 는 답변 등록 방식의 레거시. 자사몰 답글이 노출 안 되어 현재 흐름에선 미사용.
