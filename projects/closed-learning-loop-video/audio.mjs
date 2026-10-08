// ElevenLabs로 나레이션(여성)·배경 음악·효과음을 만들고 audio/timing.js(장면 길이)를 기록한다.
// 사용: ELEVENLABS_API_KEY=... node audio.mjs           (실제 생성)
//       node audio.mjs --dry                             (API 없이 예상 길이의 자리표시 음원으로 파이프라인 점검)
// 선택: ELEVENLABS_VOICE_ID (기본: Sarah), ELEVENLABS_MODEL (기본: eleven_multilingual_v2)
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';

const DRY = process.argv.includes('--dry');
const KEY = process.env.ELEVENLABS_API_KEY;
const VOICE = process.env.ELEVENLABS_VOICE_ID || 'EXAVITQu4vr4xnSDxMaL';
const MODEL = process.env.ELEVENLABS_MODEL || 'eleven_multilingual_v2';
const API = 'https://api.elevenlabs.io/v1';
const LEAD = 0.35, TAIL = 0.55; // 장면 시작 후 나레이션 시작까지, 나레이션 끝난 뒤 여유
const OUT = path.resolve('audio');
const N = JSON.parse(fs.readFileSync('narration.json', 'utf8'));

if (!DRY && !KEY) { console.error('ELEVENLABS_API_KEY가 없습니다. --dry로 점검만 할 수 있습니다.'); process.exit(1); }
fs.mkdirSync(OUT, { recursive: true });

const dur = f => +execFileSync('ffprobe', ['-v', 'error', '-show_entries', 'format=duration', '-of', 'csv=p=0', f]).toString().trim();
const tone = (file, sec, freq, vol) => execFileSync('ffmpeg', ['-loglevel', 'error', '-y', '-f', 'lavfi', '-i',
  `sine=frequency=${freq}:duration=${sec}`, '-af', `volume=${vol}`, '-ar', '44100', '-ac', '2', file]);

async function post(url, body) {
  const r = await fetch(url, { method: 'POST', headers: { 'xi-api-key': KEY, 'content-type': 'application/json', accept: 'audio/mpeg' }, body: JSON.stringify(body) });
  if (!r.ok) throw new Error(`${r.status} ${url.replace(API, '')}: ${(await r.text()).slice(0, 300)}`);
  return Buffer.from(await r.arrayBuffer());
}

const timing = { placeholder: DRY, scenes: {}, sfx: {}, bgm: null };

for (const s of N.scenes) {
  const file = path.join(OUT, `narr-${s.id}.mp3`);
  if (DRY) tone(file, s.est, 330, 0.08);
  else {
    fs.writeFileSync(file, await post(`${API}/text-to-speech/${VOICE}?output_format=mp3_44100_128`, {
      text: s.text, model_id: MODEL, language_code: 'ko',
      voice_settings: { stability: 0.45, similarity_boost: 0.8, style: 0.35, use_speaker_boost: true, speed: 1.05 },
    }));
  }
  const d = dur(file);
  timing.scenes[s.id] = { narr: `audio/narr-${s.id}.mp3`, narrDur: +d.toFixed(3), dur: +(LEAD + d + TAIL).toFixed(3), lead: LEAD };
  console.log('narration', s.id, d.toFixed(2) + 's');
}

for (const [id, x] of Object.entries(N.sfx)) {
  const file = path.join(OUT, `sfx-${id}.mp3`);
  try {
    if (DRY) tone(file, Math.min(x.dur, 0.15), 880, 0.15);
    else fs.writeFileSync(file, await post(`${API}/sound-generation`, { text: x.text, duration_seconds: x.dur, prompt_influence: 0.5 }));
    timing.sfx[id] = `audio/sfx-${id}.mp3`;
    console.log('sfx', id, 'ok');
  } catch (e) { console.error('sfx', id, 'FAILED:', e.message); }
}

const total = Object.values(timing.scenes).reduce((a, s) => a + s.dur, 0);
if (!DRY) {
  try {
    fs.writeFileSync(path.join(OUT, 'bgm.mp3'), await post(`${API}/music`, {
      prompt: N.music_prompt, music_length_ms: Math.ceil(total + 2) * 1000, force_instrumental: true,
    }));
    timing.bgm = 'audio/bgm.mp3';
    console.log('bgm ok', dur(path.join(OUT, 'bgm.mp3')).toFixed(1) + 's');
  } catch (e) { console.error('bgm FAILED (배경 음악 없이 진행):', e.message); }
}

fs.writeFileSync(path.join(OUT, 'timing.js'), 'window.TIMING = ' + JSON.stringify(timing) + ';\n');
console.log(`total ${total.toFixed(1)}s → audio/timing.js${DRY ? ' (자리표시)' : ''}`);
