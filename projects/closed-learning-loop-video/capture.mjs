// 실제 사이트(ai-secretary-memory.vercel.app)를 단계별로 조작하며 30fps 프레임으로 캡처한다.
// 출력: capture/<clip>/00000.jpg ... , capture/manifest.json (클립별 프레임 수)
import { chromium } from '/opt/node22/lib/node_modules/playwright/index.mjs';
import fs from 'node:fs';
import path from 'node:path';

const URL = 'https://ai-secretary-memory.vercel.app/';
const FPS = 30;
const VW = 1100, VH = 820;
const OUT = path.resolve('capture');

const ease = t => (t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2);

async function recordClip(browser, name, script) {
  const dir = path.join(OUT, name);
  fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(dir, { recursive: true });
  const ctx = await browser.newContext({ viewport: { width: VW, height: VH }, deviceScaleFactor: 1 });
  await ctx.grantPermissions(['clipboard-read', 'clipboard-write'], { origin: URL });
  const page = await ctx.newPage();
  await page.goto(URL, { waitUntil: 'networkidle' });
  // 즉시 스크롤 + 가짜 커서/클릭 파문 주입
  await page.addStyleTag({ content: `html{scroll-behavior:auto!important}
    #__cur{position:fixed;left:0;top:0;width:28px;height:28px;z-index:2147483647;pointer-events:none}
    #__rip{position:fixed;width:56px;height:56px;margin:-28px 0 0 -28px;border-radius:50%;border:4px solid #ff9f1c;z-index:2147483646;pointer-events:none;opacity:0}` });
  await page.evaluate(() => {
    const c = document.createElement('div'); c.id = '__cur';
    c.innerHTML = '<svg viewBox="0 0 28 28" width="28" height="28"><path d="M3 2 L3 22 L8.5 17 L12.5 26 L16 24.5 L12 15.5 L19.5 15.5 Z" fill="#111" stroke="#fff" stroke-width="1.8" stroke-linejoin="round"/></svg>';
    const r = document.createElement('div'); r.id = '__rip';
    document.body.append(c, r);
  });

  let n = 0;
  const events = [];
  const cur = { x: VW * 0.62, y: VH * 0.55 };
  const shot = async () => {
    await page.evaluate(([x, y]) => { document.getElementById('__cur').style.transform = `translate(${x}px,${y}px)`;
      return new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))); }, [cur.x, cur.y]);
    await page.screenshot({ path: path.join(dir, String(n++).padStart(5, '0') + '.jpg'), type: 'jpeg', quality: 88 });
  };
  const frames = s => Math.max(1, Math.round(s * FPS));
  const api = {
    page,
    async hold(s) { for (let i = 0; i < frames(s); i++) await shot(); },
    async scrollTo(y, s) {
      const y0 = await page.evaluate(() => scrollY);
      const k = frames(s);
      for (let i = 1; i <= k; i++) { await page.evaluate(v => scrollTo(0, v), y0 + (y - y0) * ease(i / k)); await shot(); }
    },
    async scrollToEl(sel, offset, s) {
      const y = await page.evaluate(([q, o]) => document.querySelector(q).getBoundingClientRect().top + scrollY - o, [sel, offset]);
      await api.scrollTo(y, s);
    },
    async moveTo(x, y, s) {
      const x0 = cur.x, y0 = cur.y, k = frames(s);
      for (let i = 1; i <= k; i++) { const e = ease(i / k); cur.x = x0 + (x - x0) * e; cur.y = y0 + (y - y0) * e; await shot(); }
    },
    async moveToEl(sel, fx = 0.5, fy = 0.5, s = 0.9) {
      const b = await page.evaluate(q => { const r = document.querySelector(q).getBoundingClientRect(); return [r.left, r.top, r.width, r.height]; }, sel);
      await api.moveTo(b[0] + b[2] * fx, b[1] + b[3] * fy, s);
    },
    // 파문 효과와 함께 실제 클릭을 수행하고, 펼쳐진 이미지가 모두 로드될 때까지 기다린다
    async click(sel) {
      const k = frames(0.45);
      events.push({ f: n + 2, x: Math.round(cur.x), y: Math.round(cur.y) });
      for (let i = 0; i < k; i++) {
        const t = i / k;
        await page.evaluate(([x, y, t]) => { const r = document.getElementById('__rip'); r.style.left = x + 'px'; r.style.top = y + 'px'; r.style.opacity = String(1 - t); r.style.transform = `scale(${0.4 + t})`; }, [cur.x, cur.y, t]);
        if (i === 2) {
          await page.click(sel);
          await page.evaluate(q => { const d = document.querySelector(q).closest('details'); d?.querySelectorAll('img').forEach(im => { im.loading = 'eager'; }); }, sel);
          await page.waitForFunction(q => { const d = document.querySelector(q).closest('details'); return !d || [...d.querySelectorAll('img')].every(im => im.complete && im.naturalWidth > 0); }, sel, { timeout: 15000 });
          await page.waitForTimeout(150);
        }
        await shot();
      }
      await page.evaluate(() => { document.getElementById('__rip').style.opacity = '0'; });
    },
  };
  await script(api);
  await ctx.close();
  console.log(name, n, 'frames', events.length, 'clicks');
  return { frames: n, clicks: events };
}

// 공통 패턴: 단계 제목으로 스크롤 → "실제 화면 보기" 클릭 → 캡처 이미지로 스크롤
async function openShot(a, step) {
  const sum = `#${step} details.screenshot > summary`;
  await a.moveToEl(sum, 0.18, 0.5, 0.9);
  await a.hold(0.2);
  await a.click(sum);
  await a.scrollToEl(`#${step} details.screenshot figure`, 30, 1.3);
}

const clips = {
  step1: async a => {
    await a.scrollToEl('#steps', 20, 0.01);
    await a.hold(0.6);
    await a.moveToEl('#steps .step-nav a[href="#step1"]', 0.5, 0.5, 0.8);
    await a.scrollToEl('#step1', 30, 1.2);
    await a.moveToEl('#step1 p strong', 0.5, 0.6, 0.8);
    await a.hold(0.8);
    await openShot(a, 'step1');
    await a.moveTo(560, 330, 1.0);
    await a.hold(2.2);
  },
  step2: async a => {
    await a.scrollToEl('#step2', 30, 0.01);
    await a.hold(0.4);
    await a.moveToEl('#step2 p strong', 0.5, 0.6, 0.8);
    await a.hold(0.8);
    await openShot(a, 'step2');
    await a.moveTo(620, 420, 1.0);
    await a.hold(2.2);
  },
  step3: async a => {
    await a.scrollToEl('#step3', 30, 0.01);
    await a.hold(0.4);
    await openShot(a, 'step3');
    await a.hold(1.6);
    await a.scrollToEl('#create-prompt-box', 60, 1.3);
    await a.moveToEl('#create-prompt-box button[data-copy]', 0.5, 0.5, 0.9);
    await a.click('#create-prompt-box button[data-copy]');
    await a.hold(1.6);
  },
  step4: async a => {
    await a.scrollToEl('#step4', 30, 0.01);
    await a.hold(0.4);
    await openShot(a, 'step4');
    await a.hold(1.4);
    await a.scrollToEl('#step4 .file-list', 160, 1.3);
    await a.moveToEl('#step4 .file-list li:first-child', 0.2, 0.5, 0.9);
    await a.hold(1.8);
  },
  step5: async a => {
    await a.scrollToEl('#step5', 30, 0.01);
    await a.hold(0.4);
    await openShot(a, 'step5');
    await a.hold(1.4);
    await a.scrollToEl('#step5 details.skill-document', 220, 1.3);
    await a.moveToEl('#step5 details.skill-document > summary', 0.15, 0.5, 0.8);
    await a.click('#step5 details.skill-document > summary');
    await a.scrollToEl('#step5 details.skill-document pre', 120, 1.2);
    await a.hold(1.6);
  },
  step6: async a => {
    await a.scrollToEl('#call-skill', 30, 0.01);
    await a.hold(0.4);
    await a.moveToEl('#call-skill .command-pattern code', 0.5, 0.6, 0.9);
    await a.hold(1.0);
    await a.moveToEl('#call-skill button[data-copy]', 0.5, 0.5, 0.8);
    await a.click('#call-skill button[data-copy]');
    await a.hold(0.6);
    await a.scrollToEl('#step6 details.screenshot', 200, 1.0);
    await openShot(a, 'step6');
    await a.hold(1.2);
    await a.scrollToEl('#step6 .quote', 300, 1.2);
    await a.moveToEl('#step6 .quote', 0.3, 0.25, 0.8);
    await a.hold(1.8);
  },
};

const only = process.argv.slice(2);
const browser = await chromium.launch();
const manifestPath = path.join(OUT, 'manifest.json');
const manifest = fs.existsSync(manifestPath) ? JSON.parse(fs.readFileSync(manifestPath, 'utf8')) : {};
for (const [name, fn] of Object.entries(clips)) {
  if (only.length && !only.includes(name)) continue;
  manifest[name] = await recordClip(browser, name, fn);
}
const full = { ...manifest, fps: FPS, width: VW, height: VH };
fs.writeFileSync(manifestPath, JSON.stringify(full, null, 2));
// composition.html이 file://에서도 읽을 수 있도록 JS로도 기록
fs.writeFileSync(path.join(OUT, 'manifest.js'), 'window.MANIFEST = ' + JSON.stringify(full) + ';\n');
await browser.close();
