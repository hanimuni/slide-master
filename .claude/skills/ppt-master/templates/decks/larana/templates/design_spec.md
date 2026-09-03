---
deck_id: larana
kind: deck
native_structure_mode: structured
category: brand
summary: 라라나마켓 스타일 — 구독·생활 서비스 런칭 제안, D2C 브랜드 소개, 앱 서비스 기획, 둥글고 친근한 그린-화이트
keywords: [larana, 라라나마켓, 제안서, 구독서비스, green]
primary_color: "#16A647"
canvas_format: ppt169
canvas_width: 1280
canvas_height: 720
canvas_viewbox: "0 0 1280 720"
source_canvas_width: 1920
source_canvas_height: 1080
source_viewbox: "0 0 1920 1080"
replication_mode: standard
page_count: 10
placeholders:
  01_cover: ["{{TITLE_LINE_1}}", "{{HL}}", "{{TITLE_LINE_2}}", "{{SUBTITLE}}", "{{DATE}}"]
  02_chapter: ["{{CHAPTER_NUM}}", "{{CHAPTER_TITLE}}", "{{CHAPTER_DESC}}"]
  03a_content_key_value: ["{{PAGE_TITLE}}", "{{EYEBROW}}"]
  03b_content_three_card: ["{{PAGE_TITLE}}", "{{EYEBROW}}"]
  03c_content_problem_solution: ["{{PAGE_TITLE}}", "{{EYEBROW}}"]
  03d_content_feature_callout: ["{{FEATURE_1_TITLE}}", "{{FEATURE_2_TITLE}}"]
  03e_content_persona: ["{{PAGE_TITLE}}", "{{KEY_MESSAGE}}"]
  03f_content_model_card: ["{{PAGE_TITLE}}", "{{EYEBROW}}"]
  04_ending: ["{{THANK_YOU}}", "{{ENDING_SUBTITLE}}"]
  05_contact: ["{{LABEL_1}}", "{{VALUE_1}}", "{{TAGLINE}}"]
---

# Larana Market Deck — Design Specification

> Authored 2026-09-03 from the "라라나마켓 Design System" bundle: `colors_and_type.css` tokens,
> ten `slides/*.html` slide templates authored at 1920×1080, six brand SVG assets, and eleven
> `source_slides/*.png` reference renders. Visual fidelity: literal — every coordinate is the
> CSS pixel value scaled by exactly 2/3 into the `ppt169` 1280×720 canvas. The source bundle
> had no PowerPoint package, so the Master/Layout/slot system here is newly authored.

---

## I. Template Overview

| Property | Description |
| --- | --- |
| **Display Name** | 라라나마켓 (LalanaMarket) |
| **Use Cases** | 구독·생활 서비스 런칭 제안, D2C 브랜드 소개, 앱 서비스 기획안, 소비자 대상 신규 사업 제안 |
| **Design Tone** | Soft, round, reassuring — 흰 여백 위 잔디빛 그린 하나, 모서리는 전부 둥글게 |
| **Theme Mode** | Mixed — 본문은 화이트/민트, 섹션 전환은 그린 풀필드, 컨택트만 다크 |

이 덱을 한눈에 알아보게 하는 것: **제목 안의 단어를 감싸는 그린 알약**, **영문 대문자 그린 eyebrow + 한글 대형 제목**, 그리고 **2px 그린 테두리의 흰 카드**. 색은 그린 하나뿐이고, 위계는 채움(솔리드 그린)과 테두리(아웃라인)의 차이로 만든다.

**보이스**: 한국어 `~요` 체. `~다` 종결이나 명령형을 쓰지 않는다. 동글동글 / 부드럽게 / 편하게 / 쉽게 / 알아서 — 이 어휘가 브랜드의 성격이다. 카피를 쓸 때 주어는 생략한다.

**Anti-mood**: "각진 모서리", "그린 외 두 번째 강조색", "그라디언트 배경", "빽빽한 표 중심 슬라이드".

---

## II. Color Scheme

| Role | HEX | Purpose |
| --- | --- | --- |
| Brand | `#16A647` | 유일한 강조색 — 알약, 솔리드 카드, 아이콘 디스크, 테두리, eyebrow |
| Brand hover/press | `#13923D` / `#0F7A35` | 그린 위 그린이 필요할 때만 (거의 쓰지 않음) |
| Mint pill | `#D4EFDA` | 옅은 민트 알약 |
| Mint wash | `#E8F5EC` | 대형 블롭 배경, 이미지 자리 표시 밴드 |
| Whisper mint | `#F3FAF5` | 넓은 패널 |
| Ink | `#1F1F1F` | 본문 기본 |
| Ink soft | `#4A4A4A` | 보조 설명 |
| Ink mute | `#9A9A9A` | 캡션, 자리 표시 라벨 |
| Ink dark | `#2A2A2A` | 컨택트 페이지 풀블리드 배경, 디바이스 베젤 |
| Line | `#E5E5E5` | 헤어라인 |
| Surface | `#FAFAFA` | 섹션 디바이더 상단 밴드 |
| White | `#FFFFFF` | 기본 지면 |

**적용 규칙**

- 그린은 **단일 강조색**이다. 두 번째 색을 도입하지 않는다. 여러 항목을 구분해야 하면 색이 아니라 **아웃라인 카드 vs 솔리드 그린 카드**로 나눈다.
- 문제/현상은 아웃라인 카드, 해결/제안은 솔리드 그린 카드. `03c`가 이 규칙의 원형이다.
- 다크 지면(`05_contact`)에서도 그린은 그대로 쓰고 텍스트만 흰색으로 뒤집는다.

---

## III. Typography

Pretendard 단일. 위계는 굵기 컷과 크기로만 만든다.

| Role | px | Family | Color |
| --- | --- | --- | --- |
| 표지 헤드라인 | 85.3 | `'Pretendard Black'` | `#1F1F1F` (하이라이트 단어만 알약 안 흰색) |
| 마무리 헤드라인 | 64 | `'Pretendard Black'` | `#1F1F1F` |
| 섹션 디바이더 제목 | 58.7 | `'Pretendard Black'` | `#1F1F1F` |
| 중앙 정렬 본문 제목 | 50.7 | `'Pretendard Black'` | `#1F1F1F` |
| 좌정렬 본문 제목 | 45.3 | `'Pretendard Black'` | `#1F1F1F` |
| 기능 블록 제목 | 28 | `'Pretendard Black'` | `#16A647` |
| 카드 제목 | 21.3–24 | `'Pretendard Black'` | `#16A647` |
| eyebrow | 21.3 / 24 | `'Pretendard ExtraBold'` + `letter-spacing="2.6/2.9"` | `#16A647` |
| 알약 라벨 | 20 | `'Pretendard ExtraBold'` / `'Pretendard Black'` | `#FFFFFF` |
| 본문 | 13.3–22.7 | `'Pretendard Medium'` | `#1F1F1F` / `#4A4A4A` |
| 컨택트 라벨 | 16 | `'Pretendard ExtraBold'` + `letter-spacing="1.3"` | `#16A647` |

큰 제목에는 `letter-spacing`을 음수로 준다 (`-0.9` ~ `-2.1`). 이 타이트한 자간이 한글 900 웨이트를 덩어리로 보이게 만드는 핵심이다.

**원본 대비 변경**: 소스 CSS는 영문 eyebrow에 Montserrat을 쓰지만 이 설치 환경에 없어 `'Pretendard ExtraBold'` + 자간으로 대체했다. 대문자 트래킹 인상은 유지된다. Pretendard가 없는 PC에서는 대체 폰트로 렌더된다.

---

## IV. Signature Design Elements

**① 하이라이트 알약** — 헤드라인 안의 한 단어를 완전 라운드(`rx = h/2`) 그린 알약으로 감싸고 글자를 흰색으로 뒤집는다. **표지와 마무리에서만** 쓴다. 본문에서 반복하면 힘이 죽는다.

**② eyebrow + 한글 H1** — 모든 본문 페이지 헤더는 영문 대문자 그린 eyebrow 한 줄 + 그 아래 큰 한글 제목. 좌정렬(`03b`/`03c`/`03e`/`03f`)과 중앙정렬(`03a`) 두 변형이 있다.

**③ 아웃라인 카드 vs 솔리드 카드** — 카드 반경은 전부 `rx=16`. 아웃라인은 흰 배경 + `#16A647` 1.3px 테두리, 솔리드는 `#16A647` 채움 + 흰 텍스트. 그림자를 쓰지 않는다.

**④ 아이콘 디스크 2종** — 아웃라인 디스크(`r=18–20`, 흰 배경 + 그린 1.7px 테두리 + 그린 글리프)와 솔리드 디스크(`r=28`, 그린 채움 + 흰 글리프). 번호 배지는 솔리드(`r=21.3`).

**⑤ 민트 블롭** — 화면 밖으로 크게 흘러나가는 `#E8F5EC` 타원. `03a`(`cx=886.7 cy=666.7 rx=644 ry=448`)와 `04_ending`(`cx=400 cy=360 rx=533.3 ry=493.3`)에서 디바이스 뒤를 받친다.

**⑥ 점선 커넥터** — 디바이스의 핫스팟에서 오른쪽 콜아웃 블록으로 이어지는 8px 간격 그린 점 12개. `03d` 전용.

**⑦ 3점/2점 브랜드 마크** — 지름이 같은 원 3개(또는 2개)를 겹치듯 나란히. 그린 지면 위에서는 흰색으로 뒤집는다. 벡터 원본을 파일로 링크하지 않고 각 SVG 안에 네이티브 원으로 그려 넣었다 — PPTX로 나갈 때 편집 가능한 도형이 된다.

**레이아웃 문법**

- 좌우 마진 `86.7px`(소스 130px), 상단 `66.7–73.3px`. 표지·클로징만 좌측 `93.3px`.
- 3열 그리드: `x = 86.7 / 463.6 / 840.5`, `w = 352.9`, 갭 `24px`. 2열은 `w = 537.3`, 갭 `32px`.
- 디바이스 목업: 베젤 `#2D2D2D`/`#2A2A2A`, 스크린 `#1C1C1C`/`#1A1A1A`, 반경 40, 노치 `93.3×16` 라운드.
- 페이지 번호를 쓰지 않는다. 하단 우측(또는 중앙) 태그라인 한 줄이 그 자리를 대신한다.

**밀도 리듬**: 표지·디바이더·클로징·컨택트는 크게 비운다. 본문은 3열 카드가 기본 상한이고, 카드 안은 아이콘 → 제목 → 2줄 본문 → (구분선) → 불릿 순서를 지킨다.

---

## V. Page Roster

| File | Layout key / picker name | What it looks like | Fits |
| --- | --- | --- | --- |
| `01_cover.svg` | `01_cover` / Cover | 흰 지면, 좌측 2줄 85.3px 헤드라인 + 그린 하이라이트 알약, 720px 그린 룰, 부제, 우하단 2점 마크 | 표지 |
| `02_chapter.svg` | `02_chapter` / Section Divider | 상단 오프화이트 밴드에 eyebrow+제목, 하단 그린 풀필드에 리드 카피 + 반전 3점 마크 | 섹션 전환 |
| `03a_content_key_value.svg` | `03a_content_key_value` / Content Key Value | 중앙 정렬 헤더, 우하단 민트 블롭, 좌측 디바이스, 우측 그린 알약 키 + 값 5행 | 서비스 개요, 스펙 요약 |
| `03b_content_three_card.svg` | `03b_content_three_card` / Content Three Card | 아웃라인 카드 3장, 아웃라인 아이콘 디스크 + 그린 제목 + 본문 + 민트 이미지 밴드 | 문제 정의, 3요소 소개 |
| `03c_content_problem_solution.svg` | `03c_content_problem_solution` / Content Problem To Solution | 열마다 아웃라인 문제 카드 → 그린 `»` → 솔리드 그린 해결 카드 | 문제→해결 대응 |
| `03d_content_feature_callout.svg` | `03d_content_feature_callout` / Content Feature Callout | 좌측 세로 디바이스 + 핫스팟 2개, 점선 커넥터로 이어진 번호 블록 2개, 각 블록에 아웃라인 칩 3개 | 기능 상세 2개 묶음 |
| `03e_content_persona.svg` | `03e_content_persona` / Content Persona Pair | 아웃라인 카드 2장(선화 그린 인물 + 그린 알약 이름 + 2개 라벨 불릿 그룹), 하단 그린 요약 바 | 타깃 고객, 사용자 정의 |
| `03f_content_model_card.svg` | `03f_content_model_card` / Content Model Card | 좌정렬 카드 3장, 솔리드 그린 아이콘 디스크 + 제목 + 본문 + 헤어라인 + 불릿 3 | 비즈니스 구조, 수익 모델 |
| `04_ending.svg` | `04_ending` / Closing | 대형 민트 블롭, 우측 15° 기울인 디바이스, 좌측 64px 헤드라인 + 3점 마크 + 설명 + 아웃라인 칩 4 | 마무리 메시지 |
| `05_contact.svg` | `05_contact` / Contact Dark | 다크 풀블리드, 좌하단 그린 대문자 라벨 + 흰 값 4행, 우하단 그린 3점 마크 + 태그라인 | 컨택트 |

---

## VI. Assets

번들 파일 없음 — 이 템플릿은 `images/`도 `icons/`도 쓰지 않는다.

소스 번들의 벡터 6종은 **각 SVG 안에 네이티브 도형으로 그려 넣었다**:

| 소스 파일 | 어디에 들어갔나 | 형태 |
| --- | --- | --- |
| `brand-mark.svg` / `brand-mark-2dot.svg` / `brand-mark-white.svg` | `01_cover`(2점), `02_chapter`(반전 3점), `04_ending`(3점), `05_contact`(3점) | `<circle>` 2–3개 |
| `hero-blob.svg` | `03a`, `04_ending` | `<ellipse>` 1개 |
| `persona-busy-pro.svg` / `persona-solo.svg` | `03e` | `<g transform="scale(0.65)">` 안의 선화 `<path>` 세트 |

외부 파일 링크 대신 인라인 도형으로 둔 이유: PPTX로 내보낼 때 링크 이미지가 아니라 **편집 가능한 DrawingML 도형**이 되고, 템플릿을 옮길 때 끊어질 참조가 없다. 색을 바꾸려면 해당 SVG의 `fill="#16A647"`만 교체하면 된다.

`03b`의 이미지 밴드와 `03a`·`03d`·`04`의 디바이스 스크린은 비어 있는 자리다. 실제 사진/스크린샷은 프로젝트 단계에서 `images/`로 넣고 `<image href="../images/...">`로 채운다.

---

## VII. Placeholder Overrides

프론트매터 `placeholders:` 블록 참조. 표준 어휘를 쓰되 이 템플릿 고유의 세 가지를 더한다:

- `{{TITLE_LINE_1}}` / `{{HL}}` / `{{TITLE_LINE_2}}` — 표지 헤드라인은 하이라이트 알약이 정체성이라 단일 `{{TITLE}}`로 합치지 않았다. `{{HL}}`은 알약 안에 들어갈 **한 단어**(한글 2–4자)다. 알약 폭(`300px`)은 넣는 단어 길이에 맞춰 조정하고, 같은 만큼 `{{TITLE_LINE_2}}`의 `x`도 옮긴다. 이름이 짧은 이유도 이것 — 85.3px 크기에서 긴 토큰은 알약을 넘친다.
- 아이콘 자리는 토큰이 아니라 **중립 링 글리프**(`<circle fill="none" stroke>`)로 두었다. 링 지름이 그 자리 글리프가 차지해야 할 크기이며, 실제 덱에서는 `<path>` 아이콘이나 이모지 `<text>`로 교체한다. 원본 번들의 Lucide 계열 선 아이콘을 쓰면 스트로크 굵기 2 안팎이 이 시스템과 맞는다.
- `03d`의 번호 배지는 리터럴 `01` / `02`다. 배지 원(`r=21.3`)에 두 글자 이상은 들어가지 않는다.
- `{{EYEBROW}}` — 모든 본문 페이지의 영문 대문자 섹션 라벨.
- `{{TAGLINE}}` — 페이지 번호 대신 쓰는 하단 한 줄 브랜드 서술.

`03d`는 페이지 제목이 없고 번호 블록 두 개가 곧 제목이라 `{{PAGE_TITLE}}`을 두지 않았다.
