/**
 * 답글여부 필터 디버그 스크립트
 * 실행: node debug_filter.js
 */
const puppeteer = require('puppeteer');
const fs        = require('fs');
const path      = require('path');
const cfg       = require('./config');

const sleep = ms => new Promise(r => setTimeout(r, ms));

const CONFIG = {
  sessionFile: path.join(__dirname, '.seller_session.json'),
  reviewUrl:   'https://sell.smartstore.naver.com/#/review/search',
};

function loadSession() {
  try { if (fs.existsSync(CONFIG.sessionFile)) return JSON.parse(fs.readFileSync(CONFIG.sessionFile)); }
  catch(e) {}
  return null;
}

async function main() {
  const browser = await puppeteer.launch({
    headless:        false,  // 화면 보이게
    protocolTimeout: 120000,
    args:            ['--no-sandbox', '--disable-setuid-sandbox', '--lang=ko-KR,ko', '--window-size=1600,900'],
    defaultViewport: { width: 1600, height: 900 },
  });

  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36');
  page.on('dialog', async d => { console.log('  [알림]', d.message()); await d.accept(); });

  // 세션 로드
  const saved = loadSession();
  if (saved) {
    await page.setCookie(...saved);
    console.log('[세션] 쿠키 로드 완료');
  }

  // 리뷰 페이지 이동
  console.log('[이동] 리뷰 페이지...');
  await page.goto(CONFIG.reviewUrl, { waitUntil: 'networkidle2', timeout: 30000 });
  await page.waitForFunction(
    () => !!Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '검색'),
    { timeout: 15000 }
  );
  await sleep(2000);

  // 날짜: 1주일
  for (let t = 0; t < 5; t++) {
    const r = await page.evaluate(() => {
      const btn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.trim() === '1주일');
      if (btn) { btn.click(); return true; }
      return false;
    });
    if (r) { console.log('  날짜: 1주일 선택'); break; }
    await sleep(600);
  }
  await sleep(3000);  // 더 길게 대기 - Angular 렌더링

  // 스크린샷 먼저
  await page.screenshot({ path: path.join(__dirname, 'debug_filter_page.png'), fullPage: false });
  console.log('  스크린샷 저장: debug_filter_page.png');

  // 페이지 상태 디버그
  console.log('\n=== DOM 디버그 ===');
  const debug = await page.evaluate(() => {
    const info = {};

    // 1. 모든 select 요소
    info.selectCount = document.querySelectorAll('select').length;
    info.selects = Array.from(document.querySelectorAll('select')).map(s => ({
      name: s.name || '',
      id: s.id || '',
      class: s.className || '',
      options: Array.from(s.options).map(o => o.text),
    }));

    // 2. mat-select 요소
    info.matSelectCount = document.querySelectorAll('mat-select').length;
    info.matSelects = Array.from(document.querySelectorAll('mat-select')).map(ms => ({
      class: ms.className || '',
      ariaLabel: ms.getAttribute('aria-label') || '',
      id: ms.id || '',
      html: ms.outerHTML.substring(0, 200),
    }));

    // 3. combobox 요소
    info.comboboxCount = document.querySelectorAll('[role="combobox"]').length;
    info.comboboxes = Array.from(document.querySelectorAll('[role="combobox"]')).map(cb => ({
      tag: cb.tagName,
      class: cb.className || '',
      text: cb.textContent.trim().substring(0, 50),
      html: cb.outerHTML.substring(0, 200),
    }));

    // 4. '답글여부' 텍스트 존재 여부
    info.hasReplyFilter = document.body.innerText.includes('답글여부');
    info.hasUnregistered = document.body.innerText.includes('답글미등록');
    info.hasUnregisteredInDOM = document.body.innerHTML.includes('답글미등록');

    // 5. '답글여부' 포함 요소 찾기
    info.replyFilterEls = Array.from(document.querySelectorAll('*'))
      .filter(el => {
        const txt = el.textContent.trim();
        return txt === '답글여부' || (txt.includes('답글여부') && txt.length < 50);
      })
      .slice(0, 3)
      .map(el => ({
        tag: el.tagName,
        class: el.className || '',
        text: el.textContent.trim(),
        html: el.outerHTML.substring(0, 300),
        parentHTML: el.parentElement?.outerHTML?.substring(0, 400) || '',
      }));

    // 6. '답글미등록' 포함 요소 찾기
    info.unregisteredEls = Array.from(document.querySelectorAll('*'))
      .filter(el => el.textContent.trim().includes('답글미등록') && el.textContent.trim().length < 30)
      .slice(0, 5)
      .map(el => ({
        tag: el.tagName,
        class: el.className || '',
        text: el.textContent.trim(),
        html: el.outerHTML.substring(0, 200),
      }));

    // 7. 클래스에 select 포함된 가시적 요소
    info.selectClassEls = Array.from(document.querySelectorAll('[class*="select"]'))
      .filter(el => {
        const r = el.getBoundingClientRect();
        return r.width > 20 && r.height > 10;
      })
      .slice(0, 10)
      .map(el => ({
        tag: el.tagName,
        class: el.className || '',
        text: el.textContent.trim().substring(0, 40),
        html: el.outerHTML.substring(0, 150),
      }));

    // 8. 현재 페이지 URL
    info.url = window.location.href;

    // 9. 버튼 목록 (첫 20개)
    info.buttons = Array.from(document.querySelectorAll('button'))
      .filter(b => {
        const r = b.getBoundingClientRect();
        return r.width > 0 && r.height > 0;
      })
      .slice(0, 20)
      .map(b => b.textContent.trim());

    return info;
  });

  // 결과 출력
  console.log('현재 URL:', debug.url);
  console.log('hasReplyFilter:', debug.hasReplyFilter);
  console.log('hasUnregistered:', debug.hasUnregistered);
  console.log('hasUnregisteredInDOM:', debug.hasUnregisteredInDOM);
  console.log('selectCount:', debug.selectCount);
  console.log('matSelectCount:', debug.matSelectCount);
  console.log('comboboxCount:', debug.comboboxCount);
  console.log('');
  console.log('selects:', JSON.stringify(debug.selects, null, 2));
  console.log('matSelects:', JSON.stringify(debug.matSelects, null, 2));
  console.log('comboboxes:', JSON.stringify(debug.comboboxes, null, 2));
  console.log('');
  console.log('replyFilterEls:', JSON.stringify(debug.replyFilterEls, null, 2));
  console.log('unregisteredEls:', JSON.stringify(debug.unregisteredEls, null, 2));
  console.log('');
  console.log('selectClassEls:', JSON.stringify(debug.selectClassEls, null, 2));
  console.log('');
  console.log('buttons:', JSON.stringify(debug.buttons));

  // 파일로도 저장
  fs.writeFileSync(path.join(__dirname, 'debug_filter_result.json'), JSON.stringify(debug, null, 2));
  console.log('\n결과 저장: debug_filter_result.json');

  console.log('\n=== 10초 후 종료 ===');
  await sleep(10000);

  await browser.close();
}

main().catch(e => { console.error(e); process.exit(1); });
