/**
 * edit_reply.js — 이미 등록된 리뷰 답변을 수정 (셀러센터 리뷰관리).
 *
 * 셀러센터 목록의 '답글작성' 버튼은 이미 답변된 건에 "추가 답글 불가" 팝업만 띄운다.
 * 실제 수정 경로: 리뷰내용 셀 클릭 → 인라인 editor(ng-model=vm.viewData.inputCommentContent,
 * 기존 답변이 채워진 편집가능 textarea) → "답글 수정" 버튼(=저장) → "답글이 수정되었습니다" 팝업.
 *
 * 사용:
 *   node edit_reply.js <리뷰글번호>                       # 드라이런: 현재 답변만 출력
 *   node edit_reply.js <리뷰글번호> --text "새 답변" --apply   # 실제 저장
 *   node edit_reply.js <리뷰글번호> --text-file reply.txt --apply
 *   옵션: --headed (브라우저 표시)
 *
 * 안전장치: 검색 결과가 정확히 1행일 때만 진행. 저장 전 제품명·현재 답변을 출력.
 */
const fs = require('fs');
const puppeteer = require('puppeteer');
const { CONFIG, loginToSellerCenter, ensureCoeirStore } = require('./post_product_reviews');

const argv = process.argv.slice(2);
const reviewNo = argv.find(a => /^\d{6,}$/.test(a));
const APPLY  = argv.includes('--apply');
const headed = argv.includes('--headed');
const textArg = (() => {
  const i = argv.indexOf('--text');
  if (i >= 0 && argv[i + 1]) return argv[i + 1];
  const fi = argv.indexOf('--text-file');
  if (fi >= 0 && argv[fi + 1]) return fs.readFileSync(argv[fi + 1], 'utf8').trim();
  return null;
})();

const log = (...a) => console.log(...a);
const sleep = ms => new Promise(r => setTimeout(r, ms));

if (!reviewNo) { log('사용법: node edit_reply.js <리뷰글번호> [--text "새 답변" | --text-file f] [--apply] [--headed]'); process.exit(1); }
if (APPLY && !textArg) { log('[중단] --apply 에는 --text 또는 --text-file 로 새 답변이 필요합니다.'); process.exit(1); }

async function searchReview(page, no) {
  await page.goto(CONFIG.reviewUrl, { waitUntil: 'networkidle2', timeout: 30000 });
  await page.waitForFunction(() => !!Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색'), { timeout: 45000 });
  await sleep(1500);
  for (let t = 0; t < 5; t++) { const r = await page.evaluate(() => { const b = Array.from(document.querySelectorAll('button')).find(x => x.textContent.trim() === '1개월'); if (b) { b.click(); return true; } return false; }); if (r) break; await sleep(600); }
  await sleep(800);
  const native = await page.evaluate(() => { const opt = Array.from(document.querySelectorAll('option')).find(el => el.textContent.trim() === '리뷰글번호'); if (opt) { const sel = opt.closest('select'); if (sel) { sel.value = opt.value; sel.dispatchEvent(new Event('change', { bubbles: true })); return true; } } return false; });
  if (!native) {
    const handle = await page.evaluateHandle(() => { const cs = Array.from(document.querySelectorAll('.selectize-control')); return cs.find(c => { const d = c.querySelector('.selectize-dropdown-content'); return d && d.textContent.includes('리뷰글번호'); }) || null; });
    const el = handle.asElement();
    if (el) { const input = await el.$('.selectize-input'); if (input) { await input.click(); await sleep(400); await page.evaluate(() => { const o = Array.from(document.querySelectorAll('.selectize-dropdown-content .option')).find(x => x.textContent.trim() === '리뷰글번호'); if (o) o.click(); }); await sleep(500); } }
  }
  await page.evaluate((n) => { const ta = Array.from(document.querySelectorAll('textarea')).find(t => /searchKeyword/.test(t.getAttribute('ng-model') || '') || /복수 검색/.test(t.getAttribute('placeholder') || '')); if (!ta) return; const s = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set; s.call(ta, n); ta.dispatchEvent(new Event('input', { bubbles: true })); ta.dispatchEvent(new Event('change', { bubbles: true })); }, no);
  await sleep(500);
  await page.evaluate(() => { const b = Array.from(document.querySelectorAll('button')).find(x => x.textContent.trim() === '검색'); if (b) b.click(); });
  await sleep(2500);
  return page.evaluate(() => document.querySelectorAll('.ag-center-cols-container .ag-row[row-index]').length);
}

(async () => {
  const browser = await puppeteer.launch({ headless: headed ? false : CONFIG.headless, userDataDir: CONFIG.userDataDir, protocolTimeout: 120000, args: ['--no-sandbox', '--disable-setuid-sandbox', '--lang=ko-KR,ko', '--window-size=1600,900'], defaultViewport: { width: 1600, height: 900 } });
  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36');
  page.on('dialog', async d => { log(`  [알림] ${d.message()}`); await d.accept(); });
  try {
    log('[인증] 셀러센터 로그인...'); await loginToSellerCenter(page);
    try { await ensureCoeirStore(page); } catch (e) { log('  [경고]', e.message); }
    const n = await searchReview(page, reviewNo);
    log(`[검색] 리뷰글번호 ${reviewNo} → 결과 ${n}행`);
    if (n !== 1) { log('[중단] 결과가 정확히 1행이 아니라 안전상 중단.'); await browser.close(); process.exit(2); }

    const rowProduct = await page.evaluate(() => { const c = document.querySelector('.ag-row[row-index="0"] .ag-cell[col-id="productName"]'); return c ? (c.innerText || '').trim() : ''; });
    log('[대상] 제품명:', rowProduct);

    // 리뷰내용 셀 클릭 → 인라인 editor 열기
    await page.evaluate(() => { const v = document.querySelector('.ag-center-cols-viewport'); if (v) v.scrollLeft = 0; });
    await sleep(300);
    await page.evaluate(() => { const c = document.querySelector('.ag-row[row-index="0"] .ag-cell[col-id="reviewContent"]'); if (c) { const a = c.querySelector('a'); (a || c).click(); } });
    await sleep(2000);

    const cur = await page.evaluate(() => { const ta = Array.from(document.querySelectorAll('textarea')).find(t => (t.getAttribute('ng-model') || '') === 'vm.viewData.inputCommentContent'); return ta ? ta.value || '' : null; });
    if (cur === null) { log('[중단] 답변 editor(textarea) 를 찾지 못함.'); await browser.close(); process.exit(3); }
    log('[현재 답변]', JSON.stringify(cur));

    if (!APPLY) { log('\n[드라이런] --apply 없이 실행 → 저장하지 않고 종료. 새 답변은 --text 로 전달.'); await browser.close(); return; }
    if (!cur.trim()) { log('[중단] 현재 답변이 비어있음 — 등록된 답변이 아닐 수 있어 저장 보류.'); await browser.close(); process.exit(4); }

    // 새 답변 입력
    const set = await page.evaluate((txt) => {
      const ta = Array.from(document.querySelectorAll('textarea')).find(t => (t.getAttribute('ng-model') || '') === 'vm.viewData.inputCommentContent');
      if (!ta) return false;
      const s = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value').set;
      ta.focus(); s.call(ta, txt); ta.dispatchEvent(new Event('input', { bubbles: true })); ta.dispatchEvent(new Event('change', { bubbles: true }));
      return ta.value === txt;
    }, textArg);
    log('[입력] 새 답변 반영:', set);
    await sleep(600);

    // "답글 수정" 버튼(=저장) 클릭
    const saved = await page.evaluate(() => {
      const ta = Array.from(document.querySelectorAll('textarea')).find(t => (t.getAttribute('ng-model') || '') === 'vm.viewData.inputCommentContent');
      let box = ta; for (let i = 0; i < 5 && box.parentElement; i++) box = box.parentElement;
      const cands = Array.from(box.querySelectorAll('button, a'));
      for (const label of ['답글 수정', '수정', '저장', '등록']) { const b = cands.find(x => (x.innerText || '').trim() === label); if (b) { b.click(); return label; } }
      return null;
    });
    log('[저장] 버튼 클릭:', saved);
    await sleep(2500);
    const popup = await page.evaluate(() => { const m = document.querySelector('.seller-layer-modal, .modal.show, [role="dialog"]'); const t = m ? (m.innerText || '').trim().replace(/\s+/g, ' ').slice(0, 120) : ''; if (m) { const ok = Array.from(m.querySelectorAll('button')).find(b => /확인|예/.test((b.innerText || '').trim())); if (ok) ok.click(); } return t; });
    log('[결과] 팝업:', JSON.stringify(popup));
    await sleep(1500);
    const ok = /수정되었습니다|수정 완료|등록되었습니다/.test(popup);
    log(ok ? '\n✅ 수정 완료' : '\n⚠️ 성공 팝업을 확인하지 못함 — 셀러센터에서 직접 확인 필요');
    await browser.close();
  } catch (e) { log('[오류]', e.message); try { await browser.close(); } catch {} process.exit(1); }
})();
