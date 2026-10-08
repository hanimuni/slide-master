// composition.html을 프레임 단위로 seek → 스크린샷 → MP4, audio/timing.js가 있으면 나레이션·효과음·배경 음악을 합성
// 사용: node render.mjs                 (전체 렌더)
//       node render.mjs --still 3.5 12  (지정 초의 정지 화면만 PNG로)
import { chromium } from '/opt/node22/lib/node_modules/playwright/index.mjs';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';

const args = process.argv.slice(2);
const FRAMES = path.resolve('render-frames');
const SILENT = path.resolve('render-silent.mp4');
const OUTMP4 = path.resolve('closed-learning-loop-chalkboard.mp4');
const WORKERS = 4;
const url = 'file://' + path.resolve('composition.html') + '?render';
const ff = a => execFileSync('ffmpeg', ['-loglevel', 'error', '-y', ...a], { stdio: 'inherit' });

const browser = await chromium.launch();
async function openPage() {
  const page = await browser.newPage({ viewport: { width: 1920, height: 1080 }, deviceScaleFactor: 1 });
  await page.goto(url);
  await page.evaluate(() => document.fonts.ready);
  return page;
}

if (args[0] === '--still') {
  const page = await openPage();
  fs.mkdirSync('stills', { recursive: true });
  for (const s of args.slice(1)) {
    await page.evaluate(t => window.seek(t), +s);
    await page.screenshot({ path: `stills/t${s}.png` });
  }
  console.log('stills:', args.slice(1).join(' '), 'total', (await page.evaluate(() => window.TOTAL)).toFixed(2));
  await browser.close();
  process.exit(0);
}

const page0 = await openPage();
const { total, fps, events, timing } = await page0.evaluate(() => ({ total: window.TOTAL, fps: window.FPS, events: window.AUDIO_EVENTS, timing: window.TIMING }));
await page0.close();
const N = Math.round(total * fps);
console.log(`total ${total.toFixed(2)}s, ${N} frames, audio: ${timing ? (timing.placeholder ? 'placeholder' : 'elevenlabs') : 'none'}`);
fs.rmSync(FRAMES, { recursive: true, force: true });
fs.mkdirSync(FRAMES);
let next = 0, done = 0;
await Promise.all(Array.from({ length: WORKERS }, async () => {
  const page = await openPage();
  while (next < N) {
    const i = next++;
    await page.evaluate(t => window.seek(t), i / fps);
    await page.screenshot({ path: path.join(FRAMES, String(i).padStart(5, '0') + '.jpg'), type: 'jpeg', quality: 92 });
    if (++done % 500 === 0) console.log(done, '/', N);
  }
  await page.close();
}));
await browser.close();
ff(['-framerate', String(fps), '-i', path.join(FRAMES, '%05d.jpg'), '-c:v', 'libx264', '-preset', 'slow', '-crf', '20', '-pix_fmt', 'yuv420p', SILENT]);

if (!timing) { fs.renameSync(SILENT, OUTMP4); console.log('wrote (silent)', OUTMP4); process.exit(0); }

// ── 오디오 합성 ──
const inputs = [], parts = [];
const add = f => { inputs.push('-i', f); return inputs.length / 2; }; // 0번은 영상
const ms = t => Math.max(0, Math.round(t * 1000));
const narr = events.filter(e => e.type === 'narr' && timing.scenes[e.id]);
const narrLabels = narr.map((e, k) => {
  const idx = add(timing.scenes[e.id].narr);
  parts.push(`[${idx}:a]aformat=sample_rates=44100:channel_layouts=stereo,adelay=${ms(e.t)}:all=1[n${k}]`);
  return `[n${k}]`;
});
parts.push(`${narrLabels.join('')}amix=inputs=${narrLabels.length}:normalize=0,apad=whole_dur=${total.toFixed(3)},asplit=2[narr][narrsc]`);
const VOL = { whoosh: 0.35, pop: 0.45, click: 0.6, ding: 0.5 };
const sfxLabels = [];
for (const [type, file] of Object.entries(timing.sfx)) {
  const ev = events.filter(e => e.type === type);
  if (!ev.length) continue;
  const idx = add(file);
  parts.push(`[${idx}:a]aformat=sample_rates=44100:channel_layouts=stereo,volume=${VOL[type] ?? 0.5},asplit=${ev.length}${ev.map((_, k) => `[${type}${k}s]`).join('')}`);
  ev.forEach((e, k) => { parts.push(`[${type}${k}s]adelay=${ms(e.t)}:all=1[${type}${k}]`); sfxLabels.push(`[${type}${k}]`); });
}
const mixIn = ['[narr]'];
if (sfxLabels.length) { parts.push(`${sfxLabels.join('')}amix=inputs=${sfxLabels.length}:normalize=0[sfx]`); mixIn.push('[sfx]'); }
if (timing.bgm) {
  inputs.push('-stream_loop', '-1');
  const idx = add(timing.bgm);
  parts.push(`[${idx}:a]aformat=sample_rates=44100:channel_layouts=stereo,atrim=0:${total.toFixed(3)},volume=0.30,afade=t=in:d=1,afade=t=out:st=${(total - 2.5).toFixed(2)}:d=2.5[bgmraw]`);
  parts.push(`[bgmraw][narrsc]sidechaincompress=threshold=0.02:ratio=8:attack=30:release=350[bgm]`);
  mixIn.push('[bgm]');
} else parts.push('[narrsc]anullsink');
parts.push(`${mixIn.join('')}amix=inputs=${mixIn.length}:normalize=0,alimiter=limit=0.95,atrim=0:${total.toFixed(3)}[aout]`);
ff(['-i', SILENT, ...inputs, '-filter_complex', parts.join(';'), '-map', '0:v', '-map', '[aout]',
  '-c:v', 'copy', '-c:a', 'aac', '-b:a', '192k', '-movflags', '+faststart', OUTMP4]);
fs.rmSync(SILENT);
console.log('wrote', OUTMP4, `(${narrLabels.length} narration, ${sfxLabels.length} sfx, bgm: ${!!timing.bgm})`);
