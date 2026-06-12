/**
 * cafe24_api.js — 카페24 Open API(Admin API) 클라이언트
 *
 * 새 의존성 없음 (built-in https 만 사용).
 *
 * [기능]
 *  - OAuth 2.0: getAuthUrl / exchangeCodeForToken / refreshAccessToken / getValidAccessToken
 *    · 토큰은 .cafe24_token.json 에 캐시. 만료(2시간) 임박 시 refresh_token(2주)로 자동 갱신.
 *  - 도메인 헬퍼: listBoards / listReviewArticles / listComments / postComment / getProductName
 *
 * [참고] 카페24 Admin API
 *  - 게시글:  GET  /api/v2/admin/boards/{board_no}/articles
 *  - 댓글:    GET|POST /api/v2/admin/boards/{board_no}/articles/{article_no}/comments
 *  - 상품:    GET  /api/v2/admin/products/{product_no}
 *  - 토큰:    POST https://{mall_id}.cafe24api.com/api/v2/oauth/token
 *  - 인증:    GET  https://{mall_id}.cafe24api.com/api/v2/oauth/authorize
 */

const https = require('https');
const fs    = require('fs');
const path  = require('path');
const cfg   = require('./config');

const TOKEN_FILE   = path.join(__dirname, '.cafe24_token.json');
const PRODUCT_CACHE = path.join(__dirname, 'cafe24_products.json');

const MALL_ID    = cfg.CAFE24_MALL_ID;
const CLIENT_ID  = cfg.CAFE24_CLIENT_ID;
const CLIENT_SECRET = cfg.CAFE24_CLIENT_SECRET;
const REDIRECT_URI  = cfg.CAFE24_REDIRECT_URI;
const BOARD_NO   = String(cfg.CAFE24_REVIEW_BOARD_NO || '4');
const API_VERSION = cfg.CAFE24_API_VERSION || '2024-06-01';
// 전체 권한 (사용자 요청). 17개 분류 ×읽기/쓰기. 매출통계·접속통계는 읽기 전용.
// 리뷰 자동답변에 실제 필요한 건 community(게시판) + product 뿐이나, 향후 확장 위해 전체 부여.
const SCOPES = [
  'mall.read_application',  'mall.write_application',
  'mall.read_category',     'mall.write_category',
  'mall.read_product',      'mall.write_product',
  'mall.read_collection',   'mall.write_collection',
  'mall.read_supply',       'mall.write_supply',
  'mall.read_personal',     'mall.write_personal',
  'mall.read_order',        'mall.write_order',
  'mall.read_community',    'mall.write_community',
  'mall.read_customer',     'mall.write_customer',
  'mall.read_notification', 'mall.write_notification',
  'mall.read_store',        'mall.write_store',
  'mall.read_promotion',    'mall.write_promotion',
  'mall.read_design',       'mall.write_design',
  'mall.read_salesreport',
  'mall.read_shipping',     'mall.write_shipping',
  'mall.read_translation',  'mall.write_translation',
  'mall.read_analytics',
].join(',');

const HOST = `${MALL_ID}.cafe24api.com`;

const log = msg => console.log(msg);

function assertConfigured() {
  if (!MALL_ID || MALL_ID === 'YOUR_MALL_ID')
    throw new Error('CAFE24_MALL_ID 미설정 — config.js 를 확인하세요.');
  if (!CLIENT_ID || CLIENT_ID === 'YOUR_CAFE24_CLIENT_ID')
    throw new Error('CAFE24_CLIENT_ID 미설정 — config.js 를 확인하세요.');
}

// ─────────────────────────────────────────────────────────
// 저수준 HTTPS 요청 (JSON in/out)
// ─────────────────────────────────────────────────────────
function httpRequest({ method, hostname, pathName, headers = {}, body = null }) {
  return new Promise((resolve, reject) => {
    const payload = body == null ? null
      : (typeof body === 'string' ? body : JSON.stringify(body));
    const allHeaders = { ...headers };
    if (payload != null) allHeaders['Content-Length'] = Buffer.byteLength(payload);

    const req = https.request(
      { hostname, path: pathName, method, headers: allHeaders },
      res => {
        let data = '';
        res.on('data', chunk => (data += chunk));
        res.on('end', () => {
          let parsed = null;
          try { parsed = data ? JSON.parse(data) : null; } catch (_) { parsed = data; }
          resolve({ status: res.statusCode, body: parsed, raw: data });
        });
      }
    );
    req.on('error', reject);
    if (payload != null) req.write(payload);
    req.end();
  });
}

// ─────────────────────────────────────────────────────────
// OAuth 2.0
// ─────────────────────────────────────────────────────────
function getAuthUrl(state = 'coeir') {
  assertConfigured();
  const q = new URLSearchParams({
    response_type: 'code',
    client_id:     CLIENT_ID,
    state,
    redirect_uri:  REDIRECT_URI,
    scope:         SCOPES,
  });
  return `https://${HOST}/api/v2/oauth/authorize?${q.toString()}`;
}

function basicAuthHeader() {
  const token = Buffer.from(`${CLIENT_ID}:${CLIENT_SECRET}`).toString('base64');
  return `Basic ${token}`;
}

function saveToken(tok) {
  // expires_at: 만료 시각(ms). 응답의 expires_at(ISO) 우선, 없으면 2시간 가정.
  const issued = Date.now();
  const expiresAt = tok.expires_at ? new Date(tok.expires_at).getTime()
                                   : issued + 2 * 60 * 60 * 1000;
  const refreshExpiresAt = tok.refresh_token_expires_at
    ? new Date(tok.refresh_token_expires_at).getTime()
    : issued + 14 * 24 * 60 * 60 * 1000;
  const record = {
    access_token:  tok.access_token,
    refresh_token: tok.refresh_token,
    expires_at:    expiresAt,
    refresh_expires_at: refreshExpiresAt,
    issued_at:     issued,
    mall_id:       MALL_ID,
    scopes:        tok.scopes || SCOPES,
  };
  fs.writeFileSync(TOKEN_FILE, JSON.stringify(record, null, 2), 'utf8');
  return record;
}

function loadToken() {
  if (!fs.existsSync(TOKEN_FILE)) return null;
  try { return JSON.parse(fs.readFileSync(TOKEN_FILE, 'utf8')); }
  catch (_) { return null; }
}

async function exchangeCodeForToken(code) {
  assertConfigured();
  const body = new URLSearchParams({
    grant_type:   'authorization_code',
    code,
    redirect_uri: REDIRECT_URI,
  }).toString();
  const res = await httpRequest({
    method: 'POST', hostname: HOST, pathName: '/api/v2/oauth/token',
    headers: {
      'Authorization': basicAuthHeader(),
      'Content-Type':  'application/x-www-form-urlencoded',
    },
    body,
  });
  if (res.status !== 200 || !res.body || !res.body.access_token) {
    throw new Error(`토큰 발급 실패 (status ${res.status}): ${res.raw}`);
  }
  return saveToken(res.body);
}

async function refreshAccessToken(refreshToken) {
  assertConfigured();
  const body = new URLSearchParams({
    grant_type:    'refresh_token',
    refresh_token: refreshToken,
  }).toString();
  const res = await httpRequest({
    method: 'POST', hostname: HOST, pathName: '/api/v2/oauth/token',
    headers: {
      'Authorization': basicAuthHeader(),
      'Content-Type':  'application/x-www-form-urlencoded',
    },
    body,
  });
  if (res.status !== 200 || !res.body || !res.body.access_token) {
    throw new Error(`토큰 갱신 실패 (status ${res.status}): ${res.raw}`);
  }
  return saveToken(res.body);
}

/** 유효한 access_token 반환. 만료 임박이면 refresh, refresh도 만료면 안내 에러. */
async function getValidAccessToken() {
  assertConfigured();
  const tok = loadToken();
  if (!tok || !tok.access_token) {
    throw new Error('카페24 토큰이 없습니다. 먼저 `node cafe24_auth.js` 로 인증하세요.');
  }
  // 만료 5분 전이면 갱신
  const fiveMin = 5 * 60 * 1000;
  if (Date.now() < tok.expires_at - fiveMin) {
    return tok.access_token;
  }
  // refresh 토큰 만료 여부
  if (tok.refresh_expires_at && Date.now() >= tok.refresh_expires_at) {
    throw new Error('refresh_token 만료(2주 초과). `node cafe24_auth.js` 로 재인증하세요.');
  }
  log('[카페24] access_token 만료 임박 → refresh_token 으로 갱신');
  const updated = await refreshAccessToken(tok.refresh_token);
  return updated.access_token;
}

// ─────────────────────────────────────────────────────────
// Admin API 호출 래퍼
// ─────────────────────────────────────────────────────────
// 카페24 레이트리밋: 초당 ~2건(누수 버킷 40). 요청 간 최소 간격을 두고 직렬화.
const sleep = ms => new Promise(r => setTimeout(r, ms));
const MIN_INTERVAL_MS = 550;   // ≈ 2req/s 이하 유지
let _lastReqAt = 0;
let _chain = Promise.resolve();
async function throttle() {
  // 모든 요청을 직렬화해 동시 호출이 버킷을 한꺼번에 비우지 않도록 함
  const run = _chain.then(async () => {
    const now = Date.now();
    const wait = _lastReqAt + MIN_INTERVAL_MS - now;
    if (wait > 0) await sleep(wait);
    _lastReqAt = Date.now();
  });
  _chain = run.catch(() => {});
  return run;
}

async function apiRequest(method, apiPath, { query = null, body = null, retry = true, attempt = 0 } = {}) {
  const accessToken = await getValidAccessToken();
  let pathName = apiPath;
  if (query) {
    const q = new URLSearchParams(
      Object.fromEntries(Object.entries(query).filter(([, v]) => v != null && v !== ''))
    ).toString();
    if (q) pathName += `?${q}`;
  }
  await throttle();
  const res = await httpRequest({
    method, hostname: HOST, pathName,
    headers: {
      'Authorization':         `Bearer ${accessToken}`,
      'Content-Type':          'application/json',
      'X-Cafe24-Api-Version':  API_VERSION,
    },
    body,
  });
  // 429 → 지수 백오프 후 재시도 (최대 5회)
  if (res.status === 429 && attempt < 5) {
    const backoff = 1000 * Math.pow(2, attempt); // 1s,2s,4s,8s,16s
    log(`[카페24] 429 레이트리밋 — ${backoff}ms 대기 후 재시도 (${attempt + 1}/5)`);
    await sleep(backoff);
    return apiRequest(method, apiPath, { query, body, retry, attempt: attempt + 1 });
  }
  // 401 → 토큰 갱신 후 1회 재시도
  if (res.status === 401 && retry) {
    const tok = loadToken();
    if (tok && tok.refresh_token) {
      await refreshAccessToken(tok.refresh_token);
      return apiRequest(method, apiPath, { query, body, retry: false, attempt });
    }
  }
  if (res.status < 200 || res.status >= 300) {
    const errMsg = res.body && res.body.error ? JSON.stringify(res.body.error) : res.raw;
    throw new Error(`카페24 API ${method} ${apiPath} 실패 (status ${res.status}): ${errMsg}`);
  }
  return res.body;
}

const apiGet  = (p, query) => apiRequest('GET', p, { query });
const apiPost = (p, body)  => apiRequest('POST', p, { body });

// ─────────────────────────────────────────────────────────
// 도메인 헬퍼
// ─────────────────────────────────────────────────────────
/** 게시판 목록 — 상품후기 board_no 확인용 */
async function listBoards() {
  const res = await apiGet('/api/v2/admin/boards');
  return (res && res.boards) || [];
}

/**
 * 상품후기 게시글 목록 (날짜 윈도우, 페이지네이션 자동)
 * @param {object} opts { since: 'YYYY-MM-DD', until: 'YYYY-MM-DD', boardNo }
 * @returns {Array} articles
 */
async function listReviewArticles({ since, until, boardNo = BOARD_NO } = {}) {
  const all = [];
  const limit = 100;
  let offset = 0;
  while (true) {
    const res = await apiGet(`/api/v2/admin/boards/${boardNo}/articles`, {
      start_date: since,   // 카페24 게시글 날짜 필터는 start_date/end_date (created_* 아님)
      end_date:   until,
      limit,
      offset,
    });
    const batch = (res && res.articles) || [];
    all.push(...batch);
    if (batch.length < limit) break;
    offset += limit;
    if (offset > 5000) break; // 안전장치
  }
  return all;
}

/** 특정 게시글의 댓글(답글) 목록 */
async function listComments(articleNo, boardNo = BOARD_NO) {
  const res = await apiGet(`/api/v2/admin/boards/${boardNo}/articles/${articleNo}/comments`);
  return (res && res.comments) || [];
}

/** 게시글에 답글(댓글) 등록. writer·password 필수 (config 기본값 사용) */
async function postComment(articleNo, content, opts = {}) {
  const writer   = opts.writer   || cfg.CAFE24_COMMENT_WRITER   || '코에르';
  const password = opts.password || cfg.CAFE24_COMMENT_PASSWORD || 'coeir2026!';
  const boardNo  = opts.boardNo  || BOARD_NO;
  const res = await apiPost(
    `/api/v2/admin/boards/${boardNo}/articles/${articleNo}/comments`,
    { shop_no: 1, request: { content, writer, password } }
  );
  return (res && res.comment) || res;
}

/**
 * 상품 내 리뷰 순위(노출 위치) 계산 — 최신순 기준(가장 최근=1위).
 * 자사몰 리뷰는 보통 최신순으로 노출되므로 "몇 번째로 보이는 리뷰인지"의 근사치.
 * @returns {{position:number, total:number}} position=1부터, 못 찾으면 0
 */
async function getProductReviewPosition(productNo, articleNo, boardNo = BOARD_NO) {
  if (!productNo) return { position: 0, total: 0 };
  const all = [];
  const limit = 100;
  let offset = 0;
  while (true) {
    const res = await apiGet(`/api/v2/admin/boards/${boardNo}/articles`, { product_no: productNo, limit, offset });
    const batch = (res && res.articles) || [];
    all.push(...batch);
    if (batch.length < limit) break;
    offset += limit;
    if (offset > 5000) break;
  }
  // 삭제글 제외 후 최신순 정렬
  const live = all.filter(a => String(a.deleted || 'F').toUpperCase() !== 'T');
  live.sort((a, b) => new Date(b.created_date).getTime() - new Date(a.created_date).getTime());
  const idx = live.findIndex(a => String(a.article_no) === String(articleNo));
  return { position: idx >= 0 ? idx + 1 : 0, total: live.length };
}

/** 게시글(리뷰) 삭제 — 추후 자동삭제용. 되돌리기 어려우니 호출 측에서 신중히. */
async function deleteArticle(articleNo, boardNo = BOARD_NO) {
  return apiRequest('DELETE', `/api/v2/admin/boards/${boardNo}/articles/${articleNo}`);
}

/** 댓글 삭제 */
async function deleteComment(articleNo, commentNo, boardNo = BOARD_NO) {
  return apiRequest('DELETE', `/api/v2/admin/boards/${boardNo}/articles/${articleNo}/comments/${commentNo}`);
}

// 상품명 캐시 (product_no → name)
function loadProductCache() {
  if (!fs.existsSync(PRODUCT_CACHE)) return {};
  try { return JSON.parse(fs.readFileSync(PRODUCT_CACHE, 'utf8')); }
  catch (_) { return {}; }
}
function saveProductCache(cache) {
  fs.writeFileSync(PRODUCT_CACHE, JSON.stringify(cache, null, 2), 'utf8');
}

/** product_no → 상품명 (캐시 우선) */
async function getProductName(productNo) {
  if (!productNo) return '';
  const key = String(productNo);
  const cache = loadProductCache();
  if (cache[key]) return cache[key];
  try {
    const res = await apiGet(`/api/v2/admin/products/${productNo}`, { fields: 'product_no,product_name' });
    const name = (res && res.product && res.product.product_name) || '';
    if (name) { cache[key] = name; saveProductCache(cache); }
    return name;
  } catch (e) {
    log(`[카페24] 상품명 조회 실패(product_no=${productNo}): ${e.message}`);
    return '';
  }
}

module.exports = {
  // 설정
  MALL_ID, BOARD_NO, API_VERSION, SCOPES, REDIRECT_URI, HOST,
  TOKEN_FILE, PRODUCT_CACHE,
  assertConfigured,
  // OAuth
  getAuthUrl, exchangeCodeForToken, refreshAccessToken, getValidAccessToken,
  loadToken, saveToken,
  // 호출 래퍼
  apiGet, apiPost, apiRequest,
  // 도메인
  listBoards, listReviewArticles, listComments, postComment, getProductName,
  getProductReviewPosition, deleteArticle, deleteComment,
};
