# 닫힌 학습 고리 · 칠판 영상

`https://ai-secretary-memory.vercel.app/`의 "실제 화면으로 따라 하는 6단계"를 실제로 조작하며 녹화한 화면과, 손그림 칠판 장면을 HTML로 합성한 영상이다.

| 파일 | 역할 |
|---|---|
| `capture.mjs` | 사이트를 Playwright로 조작(스크롤·커서·클릭)하며 단계별 30fps 프레임과 클릭 위치를 `capture/`에 저장 |
| `narration.json` | 장면별 나레이션 대본(여성, 아이에게 설명하는 톤), 배경 음악·효과음 프롬프트 |
| `audio.mjs` | ElevenLabs로 나레이션·배경 음악·효과음 생성 → `audio/timing.js`(장면 길이 = 음성 길이 + 여유) |
| `composition.html` | 칠판 장면 11개 타임라인. 브라우저로 열면 실시간 미리보기(스페이스: 일시정지, ←/→: 5초 이동) |
| `render.mjs` | 프레임마다 `seek(t)` → 스크린샷 → MP4, `audio/timing.js`가 있으면 나레이션·효과음·배경 음악(나레이션 때 자동으로 작아짐)을 합성 |
| `fonts/` | Gaegu (SIL OFL, Google Fonts) |

```bash
node capture.mjs                                   # 1. 녹화 (약 1분)
ELEVENLABS_API_KEY=... node audio.mjs              # 2. 음원 생성 (점검만: node audio.mjs --dry)
node render.mjs --still 12 40                      # (선택) 정지 화면 확인
node render.mjs                                    # 3. closed-learning-loop-chalkboard.mp4
```

필요한 네트워크 허용 도메인: `ai-secretary-memory.vercel.app`, `api.elevenlabs.io`, `fonts.gstatic.com`(폰트를 새로 받을 때만).
선택 환경변수: `ELEVENLABS_VOICE_ID`(기본 Sarah `EXAVITQu4vr4xnSDxMaL`), `ELEVENLABS_MODEL`(기본 `eleven_multilingual_v2`).

효과: stamp(쾅 찍히는 글자 + 분필 가루 + 화면 흔들림), 지우개 장면 전환, 클릭 순간 녹화 화면 확대 + 주황 원.
GitHub 화면(1~2단계)은 네트워크 정책상 직접 열 수 없어서, 사이트에 실린 실제 GitHub 캡처를 펼쳐 보는 장면으로 녹화했다.
