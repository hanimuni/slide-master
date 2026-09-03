---
deck_id: growup
kind: deck
native_structure_mode: structured
category: brand
summary: GrowUp 도입제안서 스타일 — 사내 서비스/앱 도입 제안, 제품 소개, 파일럿 성과 보고, 따뜻한 베이지 위 보라 시그니처
keywords: [growup, 제안서, 서비스소개, purple, 폰목업]
primary_color: "#7C6BFF"
canvas_format: ppt169
canvas_width: 1280
canvas_height: 720
canvas_viewbox: "0 0 1280 720"
source_canvas_width: 1280
source_canvas_height: 720
source_viewbox: "0 0 1280 720"
replication_mode: fidelity
page_count: 14
placeholders:
  01_cover: ["{{TITLE}}", "{{SUBTITLE}}", "{{DATE}}", "{{BRAND_MARK}}"]
  02_chapter: ["{{CHAPTER_NUM}}", "{{CHAPTER_TITLE}}", "{{CHAPTER_DESC}}"]
  03a_content_key_value: ["{{PAGE_TITLE}}", "{{EYEBROW}}", "{{PAGE_NUM}}"]
  03b_content_three_card: ["{{PAGE_TITLE}}", "{{EYEBROW}}", "{{PAGE_NUM}}"]
  03c_content_transition: ["{{PAGE_TITLE}}", "{{EYEBROW}}", "{{PAGE_NUM}}"]
  03d_content_five_topic: ["{{PAGE_TITLE}}", "{{EYEBROW}}", "{{PAGE_NUM}}"]
  03e_content_timeline: ["{{PAGE_TITLE}}", "{{EYEBROW}}", "{{PAGE_NUM}}"]
  03f_content_feature_left: ["{{PAGE_TITLE}}", "{{CONTENT_AREA}}", "{{PAGE_NUM}}"]
  03g_content_feature_right: ["{{PAGE_TITLE}}", "{{CONTENT_AREA}}", "{{PAGE_NUM}}"]
  03h_content_persona: ["{{PAGE_TITLE}}", "{{KEY_MESSAGE}}", "{{PAGE_NUM}}"]
  03i_content_impact: ["{{PAGE_TITLE}}", "{{EYEBROW}}", "{{PAGE_NUM}}"]
  03j_content_stat_grid: ["{{PAGE_TITLE}}", "{{EYEBROW}}", "{{PAGE_NUM}}"]
  04_ending: ["{{THANK_YOU}}", "{{ENDING_SUBTITLE}}"]
  05_action_dark: ["{{PAGE_TITLE}}", "{{CONTACT_INFO}}", "{{BRAND_MARK}}"]
---

# GrowUp Proposal Deck — Design Specification

> Authored 2026-09-03 from `growup-도입제안서.pptx` (18 slides, 1280×720, pptxgenjs build) plus its
> companion style guide `design.md`. Visual fidelity: literal — page geometry, card radii, shadow
> colors and type scale reproduce the source deck. The source package carried a single blank
> master/layout, so the Master/Layout/slot system here is newly authored.

---

## I. Template Overview

| Property | Description |
| --- | --- |
| **Display Name** | GrowUp 도입제안서 |
| **Use Cases** | 사내 서비스·앱 도입 제안, 제품/기능 소개, 파일럿 운영 성과 보고, 팀 대상 설득형 발표 |
| **Design Tone** | Warm, friendly, product-forward — 따뜻한 베이지 바탕 위에서 보라 한 톤이 지면을 끌고 간다 |
| **Theme Mode** | Mixed — 표지·섹션·마무리는 warm/보라, 본문은 화이트, 마지막 액션 페이지만 다크 |

**Warm + Lilac 샌드위치**: 표지(warm) → 본문(white) → 섹션 디바이더(보라 풀) → 본문(white) → 마무리(warm) → 액션(dark). 덱 전체를 이 리듬으로 배열할 때 이 템플릿이 가장 자연스럽다.

**Anti-mood** (reject at design time): "consulting 밀도 그리드", "채도 높은 무지개 카테고리 팔레트", "제목 밑 데코 라인", "텍스트만 있는 슬라이드".

**Litmus test**: 아무 본문 페이지에서 보라를 빼도 위계가 남는가 — 보라는 강조이지 구조가 아니다. 모든 페이지에 시각 요소(폰 목업 / 아이콘 디스크 / 카드 / 통계 타일) 중 최소 하나가 있어야 한다.

---

## II. Color Scheme

| Role | HEX | Purpose |
| --- | --- | --- |
| Primary | `#7C6BFF` | 강조·CTA·번호 배지·솔루션 카드·아이콘 디스크·알약 키 |
| Primary dark | `#5A49C9` | 카드 제목, 보라 텍스트 강조 |
| Primary light | `#EDEAFB` | 보조 아이콘 디스크, 옅은 배경 |
| Warm | `#FCEFE6` | 표지·섹션·마무리 배경, 카드 서피스 |
| Lilac | `#EEEAF7` | 코너 클립 디스크, 체크 서브카드 배경 |
| Card / White | `#FFFFFF` | 본문 배경, 통계 타일, 토픽 카드 |
| Dark | `#241C45` | 액션 페이지 배경, 폰 베젤 |
| Dark card | `#33295E` | 다크 페이지 위 카드 |
| Ink | `#1F1B2E` | 헤드라인·본문 기본 |
| Body | `#4A4658` | 본문 설명 |
| Gray | `#8A8696` | 캡션·페이지 번호·비활성 라벨 |

**Dark-page 보조 팔레트** (05_action_dark 전용): eyebrow `#B7A9FF`, 본문 `#C9C2E8`, 보조 `#9A8FD0`, 구분선 `#5A4E8C`, 데코 디스크 `#2E2456`.

**적용 규칙**

- 보라 한 톤이 지면의 60–70%를 지배한다. 두 번째 강조색을 도입하지 않는다.
- 다단계 구분이 필요한 페이지(`03d`, `03e`)는 **보라 명도 램프**로만 단계를 만든다: `#7C6BFF → #8F80FF → #A395FF → #B7AAFF → #CBC1FF`, 진행형은 `#EEEAF7 → #E2DCF7 → #D5CDF6 → #C7BCF6 → #B7A9FF → #9A8CFF → #7C6BFF`.
- 원본 덱의 5색 도메인 팔레트(파랑/코랄/민트/라벤더/옐로우)는 **의도적으로 이식하지 않았다**. 그 색은 원 서비스의 5근육 도메인 전용이며, 재사용 템플릿에 들어가면 주제와 무관한 색이 지면을 흐린다.
- 카드 그림자 색은 배경에 맞춘다: 보라 카드 `#9B86FF` @0.28, 흰/베이지 카드 `#C9BEEA` @0.20–0.25, **다크 배경에서는 그림자를 쓰지 않고 명도 대비로만 분리한다.**

---

## III. Typography

전 페이지 Pretendard 단일. 위계는 굵기와 크기로만 만든다 (스택: `Pretendard, "Malgun Gothic", sans-serif`).

| Role | px | Weight | Color |
| --- | --- | --- | --- |
| 표지 헤드라인 | 66.7 | bold | `#1F1B2E` (강조어만 `#7C6BFF`) |
| 섹션 디바이더 제목 | 53.3 | bold | `#FFFFFF` |
| 마무리 헤드라인 | 58.7 | bold | `#1F1B2E` |
| 다크 페이지 제목 | 42.7 | bold | `#FFFFFF` |
| 본문 페이지 제목 | 40 | bold | `#1F1B2E` |
| 기능 페이지 제목 | 33.3 | bold | `#1F1B2E` |
| 통계 큰 숫자 | 50.7 | bold | `#7C6BFF` |
| 카드 제목 (대/중) | 24–25.3 / 17.3–22 | bold | `#5A49C9` / `#1F1B2E` |
| eyebrow | 16.7 | bold, `letter-spacing="4"` | `#7C6BFF` (다크 페이지 `#B7A9FF`) |
| 본문 설명 | 18–20.7 | regular | `#4A4658` |
| 페이지 번호 | 12 | regular | `#8A8696` |

**eyebrow 규칙**: 모든 본문 페이지 헤더는 영문 대문자 한 줄 + `letter-spacing="4"` + 보라. 이 한 줄이 이 템플릿의 지문이다.
**금칙**: 제목 밑 장식 라인을 넣지 않는다. 본문은 좌정렬, 중앙 정렬은 제목과 짧은 라벨에만 쓴다.

폰트 설치 필요: 수신자 PC에 Pretendard가 없으면 대체 폰트로 렌더된다.

---

## IV. Signature Design Elements

**① 코너에서 잘리는 큰 원** — 페이지 밖으로 흘러나가는 디스크가 여백을 채운다. Layout 레이어의 고정 요소다.

| Page | Geometry | Fill |
| --- | --- | --- |
| `01_cover` | `cx=1195.2 cy=206.4 r=360` | `#EEEAF7` |
| `02_chapter` | `cx=144 cy=48 r=384` / `cx=1248 cy=720 r=336` | `#8A7BFF` / `#6E5CF0` |
| `03a` | `cx=105.6 cy=432 r=316.8` | `#EEEAF7` |
| `03f` / `03g` | `cx=1233.6 cy=638.4 r=312` / `cx=33.6 cy=638.4 r=312` | `#FCEFE6` |
| `04_ending` | 세로 타원 `cx=1147.2 cy=331.2 rx=360 ry=456` | `#EEEAF7` |
| `05_action_dark` | `cx=96 cy=768 r=336` | `#2E2456` |

**② 라운드 카드 + 배경 맞춤 그림자** — 반경은 카드 크기에 비례한다: 대형 카드 `rx≈17.3`, 스탯/토픽 타일 `rx≈15.4`, 라벨 카드 `rx≈11.5–13.4`, 알약형 `rx = h/2`. 카드 테두리(`stroke`)는 기본으로 쓰지 않고 `03c` 상단 문제 칩과 `04_ending` 값 칩에서만 의도적으로 쓴다.

**③ 컬러 원 안의 아이콘 글리프** — 아이콘은 항상 원 안에 놓인다. 옅은 배경(`#EDEAFB`)에 잉크 글리프, 또는 솔리드 보라에 흰 글리프. 페르소나 카드만 흰 원 + 보라 2.7px 테두리를 쓴다.

**④ 폰 목업** — 다크 베젤(`#241C45`) 라운드 프레임 + 안쪽 스크린 이미지. 스크린 높이는 537.6px 고정, 베젤 여백 8.2px. 화면 이미지의 원본 픽셀 비율을 지켜야 하며 (`preserveAspectRatio="none"`으로 강제 스트레치할 때는 소스 비율을 먼저 맞춘다), 목업과 텍스트 칼럼 사이는 72px를 띄운다.

**레이아웃 문법**

- 좌우 마진 `86.4px`(=0.9in), 본문 폭 `1107.2px`. 헤더 블록은 eyebrow `y=78.5` → 제목 `y=145.4`로 고정.
- 3열 카드 그리드: `x = 86.4 / 471.5 / 856.5`, `w = 337.1`, 갭 `48px`. 5열은 `w = 196.9`, 갭 `30.7px`.
- 페이지 번호는 우하단 `x=1232 y=689.4` 우정렬. 표지·마무리에는 넣지 않는다.
- 기능 페이지는 `03f`(폰 좌) → `03g`(폰 우)를 교대로 배치해 리듬을 만든다. 같은 패턴이 4장 이하면 한쪽으로 통일해도 된다.

**밀도 리듬**: 표지·섹션·마무리는 크게 비우고, 본문은 카드 3–6개로 채운다. 한 페이지에 카드 6개(통계 그리드)가 상한이다.

---

## V. Page Roster

| File | Layout key / picker name | What it looks like | Fits |
| --- | --- | --- | --- |
| `01_cover.svg` | `01_cover` / Cover | warm 바탕, 우상단 라일락 디스크, 우측 캐릭터 키비주얼, 좌측 66.7px 헤드라인, 좌하단 보라 점 + 브랜드 + 태그라인 | 표지 |
| `02_chapter.svg` | `02_chapter` / Chapter Divider | 보라 풀블리드, 좌상·우하 톤 디스크, 중앙 정렬 eyebrow/제목/설명, 하단 브랜드 | 섹션 전환 |
| `03a_content_key_value.svg` | `03a_content_key_value` / Content Key Value | warm 바탕, 좌측 폰 목업, 우측 보라 알약 키 + 값 5행 | 서비스 개요, 스펙 요약 |
| `03b_content_three_card.svg` | `03b_content_three_card` / Content Three Card | 화이트 바탕, 베이지 카드 3장, 라일락 아이콘 디스크 + 보라 카드 제목 + 2줄 본문 | 문제 정의, 핵심 3요소 |
| `03c_content_transition.svg` | `03c_content_transition` / Content Problem To Solution | 상단 회색 문제 칩 3 → 보라 화살표 디스크 → 하단 솔리드 보라 해결 카드 3 | 문제→해결 전환 |
| `03d_content_five_topic.svg` | `03d_content_five_topic` / Content Five Topic | warm 바탕, 흰 카드 5장, 보라 명도 램프 아이콘 디스크, 하단 캡션 | 5요소 분해, 구성 요소 소개 |
| `03e_content_timeline.svg` | `03e_content_timeline` / Content Progression | 하단 정렬된 7개 디스크가 좌→우로 커지고 `›`로 이어짐, 단계명·코드·수치 3단 라벨 | 성장 단계, 로드맵, 레벨 체계 |
| `03f_content_feature_left.svg` | `03f_content_feature_left` / Content Feature Phone Left | 좌측 폰 목업 + 우측 번호 배지·제목·설명·체크 카드 3 | 기능 상세 (홀수 번째) |
| `03g_content_feature_right.svg` | `03g_content_feature_right` / Content Feature Phone Right | `03f`의 좌우 반전 | 기능 상세 (짝수 번째) |
| `03h_content_persona.svg` | `03h_content_persona` / Content Persona Pair | 베이지 카드 2장(테두리 아바타 원 + 보라 알약 라벨 + 불릿 3) + 하단 보라 풀폭 메시지 바 | 타깃 사용자, 이해관계자 |
| `03i_content_impact.svg` | `03i_content_impact` / Content Impact Cards | 세로로 긴 베이지 카드 3장, 좌정렬 보라 제목 + 본문 + 불릿 3 | 기대 효과, 도입 이점 |
| `03j_content_stat_grid.svg` | `03j_content_stat_grid` / Content Stat Grid | warm 바탕, 흰 타일 3×2, 50.7px 보라 숫자 + 캡션, 하단 요약 | 파일럿 성과, 지표 요약 |
| `04_ending.svg` | `04_ending` / Ending | warm 바탕, 우측 세로 라일락 타원 + 캐릭터, 좌측 58.7px 헤드라인 + 설명 + 아웃라인 값 칩 4 | 마무리 메시지 |
| `05_action_dark.svg` | `05_action_dark` / Action Dark | 다크 바탕, 좌측 요청사항 카드 3 + 우측 번호 원 스텝 사다리 3 + 하단 담당자 정보 | 요청 사항 / 다음 단계 / 컨택트 |

---

## VI. Assets

`images/` (SVG에서 `../images/<name>`으로 참조):

| File | Size | Used by | Intended usage |
| --- | --- | --- | --- |
| `mascot_hero.png` | 768×768 (투명 배경) | `01_cover` | 표지 우측 키비주얼. 주제에 맞는 캐릭터/키 이미지로 교체하는 자리다. |
| `mascot_closing.png` | 768×768 (투명 배경) | `04_ending` | 마무리 페이지 키비주얼. 표지와 짝을 이루되 다른 포즈/표정을 쓴다. |
| `phone_screen.png` | 390×844 | `03a`, `03f`, `03g` | 폰 목업 스크린. 새 앱 스크린샷으로 교체할 때 **원본 픽셀 비율을 먼저 확인**한다. |

원본 덱의 나머지 자산(7단계 마스코트 세트, 앱 스크린샷 6종)은 그 서비스 고유의 샘플 콘텐츠이므로 번들하지 않았다. `03e`의 7단계는 이미지 대신 보라 램프 디스크로 구현했다.

**아이콘 자리 규약**: 원본이 이모지를 쓰던 자리는 전부 **중립 링 글리프**(`<circle fill="none" stroke>`)로 두었다. 링의 반지름은 원본 글리프 크기의 0.42배이고 색은 그 자리의 글자색을 그대로 쓴다. 실제 덱을 만들 때 이 `<circle>`을 이모지 `<text>`나 아이콘 `<path>`로 교체한다 — 링의 지름이 곧 글리프가 차지해야 할 크기다. 긴 `{{...}}` 토큰을 이 자리에 넣으면 원을 넘치므로 두지 않았다.

---

## VII. Placeholder Overrides

프론트매터 `placeholders:` 블록 참조. 이 템플릿은 표준 어휘(`{{TITLE}}` / `{{PAGE_TITLE}}` / `{{CONTENT_AREA}}` / `{{PAGE_NUM}}` / `{{THANK_YOU}}`)를 그대로 쓰되, 두 가지를 추가한다:

- `{{EYEBROW}}` — 모든 본문 페이지의 영문 대문자 섹션 라벨. 이 덱의 헤더는 eyebrow 없이 성립하지 않으므로 별도 슬롯으로 노출한다.
- 패턴별 인덱스 자리(`{{CARD_n_TITLE}}`, `{{TOPIC_n}}`, `{{Sn_VAL}}`, `{{STAT_n}}`, `{{ASK_n_BODY}}` 등) — 고정 개수 그리드가 곧 페이지 정체성이라 자유 본문 영역으로 뭉뚱그리지 않는다. 좁은 칸(5열 토픽 카드 `196.9px`, 7단 진행 레인 `~160px`, 통계 타일 숫자 `50.7px`)에 들어가는 토큰은 **칸 폭 안에서 렌더되도록 짧게** 지었다. 이름을 늘리려면 그 칸의 폭도 함께 키워야 한다.
- `03f`/`03g`의 번호 배지와 `05_action_dark`의 스텝 번호는 토큰이 아니라 리터럴 숫자(`01`/`02`, `1`/`2`/`3`)다. 배지 원(`r=34.6`/`r=24`)에 두 글자 이상은 들어가지 않는다.
