---
name: pptx-com-effects
description: >
  파워포인트를 직접 조작(COM)해서 "도형 병합·모핑·3D·그림 효과·애니메이션" 같은 파워포인트 고유 기능이 들어간
  PPTX 장표를 만들고, 슬라이드쇼를 영상으로 뽑아 움직임까지 검증하는 작업 방식. 검증된 레시피 65종 동봉 —
  글자 안에 비율만큼 차오르기, 로고·글자 안에서 GIF 재생, 롱 섀도 제목, 같은 색 그림자로 빛나는 글자,
  글자 하나를 빈칸으로 만든 워드마크, 글자 뒤에 피사체 겹치기(입체감), 배경 흐림 + 피사체 강조,
  글래스모피즘 카드, 글자 갈라지기(모핑), 글자가 뚫린 색 패널 슬라이드 인(모핑), 사진 한 장 확대 이동(모핑),
  3D 모델 회전(모핑), 모핑으로 요소 등장, 사진 위 그라데이션 가림막, 비네팅, 시상식·종무식 금빛 제목(금빛 그라데이션·영상 배경·글자 안 영상), 뉴모피즘 글자,
  인물이 다가오는 시차 확대(모핑), 숫자 넘기기(모핑), 도형으로 그린 그래프(막대·원형·도넛·링), 글자 꾸미기 8종(윤곽선·그림자), 층층이 솟아오르는 사진(모핑), 마우스 오버 메뉴, 사내 템플릿(표지+내지), 회전 다이얼 전환(모핑), 글자 안 사진 바꾸기, 선을 덜어 낸 표, 카드 배열, 사진 불규칙 배치(10종), 막대에 사진 채우기, 콰직 충격 줌, 색 패널 닫힘, 창이 열리는 사진, 커튼 인트로, 줌 맵, 절반 띠 뒤집기, 이미지 슬라이더, 용어 풀이 말풍선. GIF 내보내기 지원. PPT 작업 방식·유의사항 참고 문서(30장) 포함.
  고를 수 있는 그림 카탈로그(215종) 포함 — "PPT 효과 뭐가 있어", "효과 카탈로그/메뉴 보여줘", "이 장표에 어울리는 효과 추천해줘"에도 쓴다.
  기획안이나 이미 만든 덱을 주고 "이 스킬 효과로 바꿔줘", "이 PPT에 효과 입혀줘", "초안을 이 효과들로 다시 만들어줘"라고 하면 장표별로 어울리는 효과를 제안해 승인받은 장만, 원래 디자인(색·서체·DESIGN.md)을 유지한 채 사본에서 바꾼다.
  트리거 — "숫자 안에 색 채우기", "퍼센트만큼 차오르는 글자", "로고 안에 GIF", "글자 안에 영상/패턴",
  "롱 섀도", "긴 그림자 표지", "네온/빛나는 글자", "글래스모피즘", "유리 카드", "글자 뒤로 사람이 지나가게",
  "배경 흐리게 하고 인물 강조", "글자 가르기", "텍스트 쪼개기", "사진 위 글자 가독성", "그라데이션으로 어둡게", "비네팅", "주변을 어둡게", "시상식 PPT", "종무식", "금빛 글자", "골드 그라데이션", "배경에 영상 깔기", "글자 안에서 영상", "뉴모피즘",
  "인물이 다가오게", "숫자가 바뀌는 효과", "그래프/차트 예쁘게", "도넛 그래프", "달성률 링", "글자 테두리/윤곽선", "글자 그림자", "표 깔끔하게", "사내 템플릿", "회사 PPT 표지", "메뉴 누르면 이동", "마우스 올리면", "돌아가는 전환", "글자 안에 사진", "GIF로 내보내기", "카드 배열", "사진 여러 장 배치", "사진을 원형으로/비스듬히/불규칙하게", "콰직/폭발/임팩트 있게", "확대/축소(줌) 발표", "인트로 영상처럼", "이미지 슬라이더", "용어 설명 뜨게", "2024에서 2025로", "PPT 잘 만드는 법/유의사항", "모핑 전환", "모핑으로 움직이게", "확대되며 이동",
  "3D 모델 넣고 돌리기", "도형 병합으로", "빼기/교차/조각", "유튜브 PPT 강의·쇼츠처럼 만들어줘",
  "이 영상에 나온 PPT 기법 따라 해줘", "움직이는 장표", "슬라이드쇼에서 움직이게".
  python-pptx·pptxgenjs로는 못 만드는 효과(도형 병합, SVG 도형 변환, 잘린 GIF, 모핑, 글자 3D, 그림 흐림,
  3D 모델)가 필요하면 이 스킬을 쓴다. 일반 발표자료 여러 장을 만드는 일은 pptx-creator-review가 맡는다.
  같은 효과를 HTML 웹 슬라이드로 만들 때는 html-slide-effects를 쓴다.
---

# PPTX COM Effects

파워포인트 튜토리얼 영상에 나오는 효과는 대부분 **파워포인트 프로그램 안에서만 되는 기능**
(도형 병합, SVG→도형 변환, 모핑, 글자 3D, 그림 꾸밈 효과, 3D 모델)으로 만든다. 파일 생성 라이브러리로는
흉내만 낼 수 있고 결과물이 파워포인트에서 편집되지 않는다. 그래서 이 스킬은 설치된 파워포인트를 스크립트로
조작해 사람이 손으로 한 것과 같은 도형을 만든다.

움직임이 핵심인 장표는 정지 이미지로 검수할 수 없다. 슬라이드쇼를 mp4로 내보내 프레임을 보는 것이
유일한 검증 방법이다. 정지 효과는 슬라이드를 PNG로 내보내 본다. 모든 레시피가 그 단계를 포함한다.

## 전제

- Windows + 데스크톱 PowerPoint 설치 (2026-10 기준 PowerPoint 16에서 검증). 없으면 이 스킬은 쓸 수 없다고 알린다.
- 프레임 확인용 ffmpeg. 패턴 GIF 생성용 Python + Pillow.
- 스크립트는 `scripts/`에 있다. PowerShell 도구에서 `& "<SKILL_DIR>\scripts\<이름>.ps1" ...`로 실행한다.
- 슬라이드는 960×540pt. 좌표·크기 매개변수는 모두 이 기준이다.

## 먼저 카탈로그를 보여 준다

사용자는 이 스킬에 무엇이 들어 있는지 모른다. 효과 이름만으로는 모습을 떠올릴 수 없으니, 고를 수 있게 메뉴부터 내민다.

- **효과를 지정하지 않고 스킬을 불렀거나 "뭐가 있어?"라고 물으면**: `catalog/catalog.md`를 읽어 번호가 붙은 표를 답변에 그대로 보여 주고,
  그림 카탈로그를 기본 브라우저로 연다(`Start-Process "<SKILL_DIR>\catalog\catalog.html"`). 그런 다음 번호나 이름으로 고르게 한다.
- **장표 내용만 주면**("매출 75% 달성 장표"): 카탈로그의 "어울리는 장표"를 보고 맞는 효과 2~3개를 번호·이유와 함께 추천하고 고르게 한다.
- **기획안이나 만든 덱을 주며 효과를 입혀 달라고 하면**: 아래 "받은 기획안·초안에 효과 입히기"로 간다(제안하고 승인받은 뒤 실행).
- **기획안으로 덱 전체를 새로 만들어 달라고 하면**: 아래 "덱 한 벌을 통째로 만들 때"로 간다(마스터 규칙을 정해 초안을 끝까지 만든 뒤 고쳐 나간다).
- **번호나 이름을 말하면**("7번으로", "비네팅으로"): 바로 아래 작업 순서로 간다. 카탈로그를 다시 보여 주지 않는다.

카탈로그는 html-slide-effects와 함께 쓴다(215종. ‘PPTX 전용’·‘웹 전용’ 표시가 없으면 양쪽 가능. 2026-10-09 기준 웹으로도 되는 것이 182종, PPTX 전용 33종, 웹 전용 1종).
항목은 **쓰임새** 기준 12개 범주로 나뉜다: 표지·오프닝, 목차·간지, 글자·타이포 효과, 사진·인물, 수치·KPI, 차트·그래프, 지도, 흐름·일정·로드맵,
구조·조직·체계, 비교·카드·목록, 행사·마무리, 전환·인터랙션·도구. 추천할 때는 사용자의 장표가 어느 범주인지부터 잡고 그 안에서 고른다.
`catalog.html`은 위쪽에 검색창(이름·용도·재료·번호)과 범주 칩, ‘움직임/정지/웹 가능’ 필터가 있다. 움직이는 항목은 카드에서 바로 재생되고, 위쪽 ‘▶ 슬라이드 쇼’ 단추를 누르면 지금 걸러진 항목 중 움직이는 것만 크게, 차례로 재생한다(← → 로 넘김, 정지 항목은 빠진다). 번호는 한번 붙으면 바뀌지 않는다.

레시피를 추가하면 견본을 `catalog/thumbs/<이름>.jpg`(640×360) 또는 `<이름>.mp4`(움직임, 10초 이내)로 넣고, `scripts/build_catalog.py`의 `ITEMS` **맨 끝에** 한 줄 더한 뒤
`CATEGORIES`의 맞는 범주에 그 번호를 넣고 실행한다(범주에 빠진 번호가 있으면 실행이 멈춘다). 영상의 정지 그림은 `thumbs/posters/`에 저절로 만들어진다.

## 받은 기획안·초안에 효과 입히기

사용자가 기획안이나 이미 만든 덱을 주고 "이 스킬 효과로 바꿔 줘"라고 하면 `references/apply-to-deck.md`를 읽고 그 절차를 따른다. 뼈대는 이렇다.

1. **읽기** — `deck_inspect.ps1 -Deck a.pptx -OutDir <폴더>`로 장별 그림과 `deck.json`(쓰인 색·서체, 장별 글)을 뽑아 직접 본다.
2. **유지할 디자인 정리** — DESIGN.md가 있으면 그것, 없으면 원본에서 실제로 쓰인 배경·글자·강조색과 서체.
3. **장별 판단** — 목표는 "효과를 넣었다"가 아니라 **누가 봐도 인정할 만큼 임팩트 있는 결과물**이다. 덱의 승부처(표지·간지·핵심 수치·마무리)를 먼저 찾고, 장이 하는 일에 맞는 효과를 카탈로그·레퍼런스 사례에서 우선 고른다. 맞는 사례가 없는데 임팩트가 필요한 장이면 "그대로 둠"으로 끝내지 말고 그 장을 위한 장면을 **새로 기획**해 제안한다(기준은 `references/ppt-craft.md` 28장). 임팩트가 필요 없는 장은 "그대로 둠".
4. **제안표를 보여 주고 멈춘다** — 장 / 지금 내용 / 제안(카탈로그 번호, 또는 ‘새 기획’) / 바뀌는 점 / 필요한 재료. 표 아래에 **‘한 걸음 더’ 아이디어를 꼭 한 가지** 덧붙인다: 카탈로그 사례를 그대로 쓰는 것보다 더 나아질 발전안 하나(무엇을, 어느 장에, 왜 더 센지). **승인 전에는 빌드도 사본도 만들지 않는다.** ‘한 걸음 더’와 ‘새 기획’도 승인된 것만 만든다. 만든 결과를 사용자가 컨펌하면 그때 스킬에 레시피·카탈로그 항목으로 올린다(컨펌 없이 스킬을 바꾸지 않는다).
5. **승인된 장만 실행** — 레시피를 원본의 색·서체로 빌드하고 `deck_replace.ps1 -Deck a.pptx -Replace '3=fx3.pptx','4=fx4.pptx' -Out a_fx.pptx -PreviewDir <폴더>`로 **사본에서** 그 장만 바꾼다. 원본 파일은 쓰지 않는다.
6. **원본과 나란히 놓고 확인** — 문구·색·서체가 유지됐는지, 바꾸지 않은 장이 그대로인지 본 뒤 보고한다.

원본 장의 제목 띠·로고·쪽 번호는 `deck_replace.ps1`의 `-KeepRecurring`(장마다 같은 자리에 있는 도형), `-KeepTitle`(원본 제목), `-MasterShapes`(마스터에 있는 것)로 새 장에 가져온다. 이때 레시피는 자기 제목 없이 빌드한다. 복사는 클립보드를 거치므로 사용자의 클립보드가 지워진다는 점을 미리 알린다. 웹 덱(HTML)을 받았으면 html-slide-effects의 같은 절을 따른다.

## 덱 한 벌을 통째로 만들 때

기획안만 있고 디자인이 아직 없는 덱을 처음부터 만드는 경우다(이미 디자인된 덱에 효과를 입히는 위 절차와 다르다). `references/ppt-craft.md` 31장을 먼저 읽는다. 뼈대는 이렇다.

1. **마스터 규칙부터 정한다** — 헤더(구분 라벨·제목·거버닝 메시지·내비게이션) / 본문 / 푸터의 자리, 상하좌우 여백, 색 세 값, 서체와 글자 단계, 카드 한 가지, 본문 움직임 한 가지. 답변 첫머리에 이 규칙을 몇 줄로 밝힌다.
2. **본문은 그 틀 하나로 짓는다** — 레시피를 장마다 불러 합치지 않는다(레시피마다 제목 위치·바탕·카드가 달라 "일관성이 없다"는 평가를 받았다). `scripts/deck_helpers.ps1`의 틀 함수(`Set-DeckFrame`, `Chrome`)와 원형 함수(`StatWide`, `BA`, `RowCard`, `RowBA`, `TCard`, `Timeline`, `Photo`, `Emph`)로 덱 스크립트 하나에서 차례로 짓는다. `scripts/deck_demo.ps1`이 그대로 복사해 시작할 수 있는 견본이다(6장표, `-Out`·`-PreviewDir`·`-Video`).
3. **차별화는 몇 곳에만** — 표지, 목차, 마무리, 그리고 본문에서 특별히 강조할 성과·과제 2~4곳. 모핑과 레시피 장면은 이 자리에만 쓴다. 강조는 한 곳으로 몰지 않는다.
4. **같은 패턴을 잇지 않는다** — 같은 원형은 세 장까지. 글자 카드만 나란한 장에는 주인공(큰 핵심어, 높이 차, 어두운 닻 카드, 아이콘·사진)을 하나 세운다.
5. **초안을 끝까지 만들고 나서 고친다** — 장마다 승인받느라 멈추지 않는다. 초안을 보여 준 뒤 글자 크기, 밋밋한 장, 강조 위치를 고쳐 나간다. 고칠 때마다 장별 그림 묶음과 영상 프레임을 다시 보고, 버전마다 파일 이름을 나눈다.

## 작업 순서

1. **기법 파악.** 사용자가 영상 링크를 주면 `watch` 스킬(또는 yt-dlp + ffmpeg 프레임 시트)로 자막과 화면을 보고,
   메뉴 경로와 클릭 순서까지 단계로 적는다(도형 병합은 선택 순서에 따라 결과가 달라진다).
   영상이 없으면 요청을 아래 레시피 표에 대응시킨다.
2. **레시피 선택.** 표에 맞으면 스크립트를 그대로 쓴다. 안 맞으면 "새 기법 만들기"를 따른다.
3. **재료 준비.** "재료와 권리" 참고. 사진·배경 제거 이미지·3D 모델은 사용자가 준 것을 쓰고, 없으면 먼저 묻는다.
4. **빌드.** 레시피 스크립트를 `-Out`(pptx)과 `-Preview`(mp4 또는 png)와 함께 실행한다.
5. **검증.** 미리보기를 Read로 직접 본다. 정지 효과는 png 한 장, 움직임은 시간 순 프레임 시트와 마지막 프레임.
   ```bash
   ffmpeg -v error -y -i preview.mp4 -vf "fps=3,scale=640:-1,tile=4x4" -frames:v 1 sheet.png
   ffmpeg -v error -y -sseof -0.3 -i preview.mp4 -frames:v 1 last.png
   ```
6. **전달.** pptx와 미리보기를 사용자 작업 폴더에 두고, 무엇을 눈으로 확인했는지와 확인하지 못한 것을 구분해 알린다.
   편집 화면과 슬라이드쇼의 모습이 다른 경우(표의 "알릴 점")는 반드시 말해 준다.

수정 요청이 오면 매개변수만 바꿔 다시 빌드하고 5번을 반복한다. 같은 출력 경로를 덮어쓰려면 `-Force`.

## 레시피 표

매개변수의 뜻과 기본값은 각 스크립트 맨 위 `param(...)` 블록의 주석에 있다. 쓰기 전에 그 블록만 읽는다.
여러 줄 문구는 `|`로 줄을 나눈다. `-Preview`가 `.png`면 정지 이미지, 그 밖이면 슬라이드쇼 영상이다.

| 효과 | 스크립트 | 필수 재료 | 미리보기 | 알릴 점 |
|---|---|---|---|---|
| 글자 안에 비율만큼 차오르기 | `text_fill.ps1 -Label TWOSOME -Number 75` | — | mp4 | 편집 화면에는 빈 윤곽선만 보인다. 둥근 윗면(`-Dome`)이면 가운데가 비율보다 높게 찬다 |
| 로고·글자 안에서 GIF 재생 | `gif_in_shape.ps1 -Svg logo.svg -Gif p.gif` 또는 `-Text SALE` | SVG(PNG 불가), GIF | mp4 | 편집 화면에서는 GIF가 멈춰 보인다. 파일 크기 ≈ GIF 크기 |
| 롱 섀도 제목 | `long_shadow.ps1 -Text 'PPT\|PROJECT'` | — | png | 글자 3D 서식이라 문구는 그대로 편집된다 |
| 같은 색 그림자로 발광 | `glow.ps1 -Text PPTPRO` 또는 `-Image a.jpg -Color '#FF4FA0'` | — | png | 어두운 배경에서만 효과가 난다. 그림은 `-Color`를 그림의 주된 색으로 준다 |
| 글자 하나를 빈칸으로 | `negative_letter.ps1 -Text WATER -Index 2 -Cover 0.72` | — | png | 결과는 도형이라 문구를 더는 편집할 수 없다. 가운데가 뚫린 글자(A, O, D, P)가 잘 읽힌다. 둥근 글자(O)는 사각 블록이 도드라져 A·삼각형 계열보다 투박하다 |
| 글자 뒤에 피사체(입체감) | `subject_layers.ps1 -Image p.jpg -Cutout cut.png -Text RUN` | 사진 + 같은 구도의 배경 제거 PNG | png | 피사체의 일부만 글자 앞에 두려면 컷아웃에서 나머지를 지운 이미지를 받는다 |
| 배경 흐림 + 피사체 강조 | `subject_layers.ps1 -Image p.jpg -Cutout cut.png -Blur 20 -Brightness -0.2 -CutoutScale 1.05` | 위와 같음 | png | 제목을 인물 앞에 두려면 `-Text T -TextOnTop -TextLeft 70` (기본은 인물 뒤) |
| 글래스모피즘 카드 | `glass_card.ps1 -Image p.jpg -Title T -Body 'a\|b' -Card 540,110,360,320` | 사진 | png | 카드 뒤 사진이 단색에 가까우면 유리 느낌이 약하다 |
| 글자 갈라지기 | `text_slice.ps1 -Text SALE` | — | mp4 | 2장짜리(온전한 글자 → 모핑 0.3초 → 갈라진 글자) |
| 글자 구멍으로 열리는 표지 | `panel_open_cover.ps1 -Image p.jpg -Title '가치 있는\|삶의 시작' -Kicker K -Sub s -Foot f` (`-Side Right`, `-Sound boom.mp3`) | 사진 (효과음 선택) | mp4 | 한 장짜리 자동 애니메이션. 글자 구멍은 처음에 (제목 자리 + 슬라이드 폭 − 패널 폭) 위치의 사진을 비추므로 그 자리가 밝은 사진을 고른다. 굵은 서체일수록 잘 보인다. 제목 마지막 줄이 강조색 |
| 글자 뚫린 색 패널 슬라이드 인 | `cutout_panel.ps1 -Image p.jpg -Text 'YOU\|HAD ME\|AT HELLO'` | 사진 | mp4 | 2장짜리. 글자 구멍으로 사진이 비친다 |
| 사진 한 장 확대 이동 | `image_zoom.ps1 -Image p.jpg -Title T -Stops '0.72,0.5,2.4,SUN','0.2,0.85,3,CITY'` | 고해상도 사진 | mp4 | 지점마다 한 장씩 늘어난다. 확대 배율만큼 원본 해상도가 필요하다 |
| 3D 모델 회전 이동 | `model3d_morph.ps1 -Model m.glb -Title 'NEW\|PHONE'` | GLB | mp4 | 2장짜리. PowerPoint 2019 이상에서만 3D가 보인다 |
| 모핑으로 요소 등장 | `morph_reveal.ps1 -Title 'A\|B' -Items 'X\|Y\|Z'` | — | mp4 | 2장짜리. 슬라이드 밖에 둔 패널이 들어온다 |
| 사진 위 그라데이션 가림막 | `gradient_scrim.ps1 -Image p.jpg -Title T -Body b -Side Bottom` | 사진 | png | 가림막은 한 방향 그라데이션 도형이라 `-Cover`를 줄이면 사진이 더 밝게 남는다 |
| 비네팅(주변 어둡게, 가운데 집중) | `gradient_scrim.ps1 -Image p.jpg -Side Center -Title T -Body b` | 사진 | png | 방사형 그라데이션 70/25/0%. 원래 어두운 사진에서는 전체가 너무 어두워지니 `-Transparency 0.85,0.45,0.1`처럼 올린다 |
| 시상식·종무식 금빛 제목 | `gold_award.ps1 -Title '2025\|브랜드 종무식' -Sub S -Frame` (+ `-Video bg.mp4`, `-TextVideo glitter.mp4`) | 영상은 선택 | png / mp4 | 영상은 자동 실행 + 반복으로 들어간다. `-TextVideo`를 쓰면 제목이 도형이 되어 문구를 못 고친다. 편집 화면에서는 영상이 멈춰 보인다 |
| 뉴모피즘 글자 | `neumorph_text.ps1 -Text PPTPRO` | — | png | 배경과 글자가 같은 회색이어야 한다. 흰색·검정 배경에서는 안 된다 |
| 인물이 다가오는 시차 확대 | `subject_layers.ps1 -Image p.jpg -Cutout cut.png -MorphZoom 1.35` | 사진 + 배경 제거 PNG | mp4 | 2장짜리. 피사체는 크게, 배경은 조금만 커진다 |
| 숫자 넘기기 | `number_roll.ps1 -Prefix 202 -From 4 -To 5` | — | mp4 | 2장짜리. 배경과 같은 색 도형으로 가리므로 배경이 단색이어야 한다 |
| 도형으로 그린 그래프 | `shape_chart.ps1 -Type Column -Values 120,450,950 -Labels 'a\|b\|c'` (`-Type Bar/Pie/Donut/Ring`) | — | png | 차트 개체가 아니라 도형이다. 수치를 바꾸려면 다시 빌드한다 |
| 글자 꾸미기 8종(채우기·윤곽선·그림자) | `text_style.ps1 -Text PPTPRO -Style FillOutlineShadow` (`-Style All` = 견본 한 장) | — | png | Double 두 종은 글상자가 둘이라 문구를 두 번 고쳐야 한다 |
| 층층이 솟아오르는 사진 | `layered_rise.ps1 -Background sky.jpg -Layers far.png,near.png -Text TOKYO -TextAfter 1` | 배경 + 층별 투명 PNG | mp4 | 2장짜리. 층을 나누는 배경 제거는 직접 해야 한다 |
| 마우스 오버 메뉴(웹사이트식 이동) | `hover_menu.ps1 -Title T -Sections 'a\|b\|c'` | — | png | 슬라이드쇼에서만 동작한다. 영상·PDF에는 남지 않는다 |
| 사내 템플릿(표지 + 내지) | `brand_template.ps1 -Image g.jpg -Sub s -Title t -Contact 'a\|b\|c' -BodyTitle b` | 그래픽 이미지 (로고는 선택) | png | 미리보기는 내지, 표지는 `*_cover.png` |
| 회전 다이얼 전환 | `wheel_menu.ps1 -Sections 'a\|b\|c\|d' -Images 1.jpg,2.jpg` | 섹션별 배경 사진(선택) | mp4 | 섹션 수만큼 장이 생긴다. 다이얼 안에는 어둡게 하지 않은 원래 사진이 비친다 |
| 글자 안 사진이 닦이며 바뀌기 | `text_picture_wipe.ps1 -Text WELCOME -Sub s -Images a.jpg,b.jpg` | 사진 2장 이상 | mp4 | 한 장짜리 애니메이션. 글자는 그림이 되어 수정할 수 없다 |
| 선을 덜어 낸 표 | `clean_table.ps1 -Rows 'h1\|h2','a\|b','c\|d' -Highlight 1` | — | png | 진짜 표라 내용은 그대로 고칠 수 있다 |
| 핵심 문장 + 카드 배열 | `card_row.ps1 -Key 문장 -Cards 'a\|b\|c' -Bodies 'x\|y\|z'` | — | png | 본문 안 줄바꿈은 `;` |
| 사진을 불규칙한 모양으로 배치 | `photo_layout.ps1 -Images a.jpg,b.jpg,c.jpg -Layout Circles -Title T` (`Columns/Mosaic/Circles/Slant/Scatter/Brush/Blob/Diagonal/Frames/Arc`) | 사진 1~4장 | png | 사진은 모양대로 잘린다. 비율은 깨지지 않지만 가장자리는 잘려 나간다 |
| 같은 크기 사진 격자 / 사진 한 장 + 설명 띠 | `photo_layout.ps1 -Images ... -Layout Grid -Grid 5,3 -Title T` / `-Layout Callouts -Captions '머리;설명\|...'` | 사진 여러 장 / 1장 | png | Grid는 사진이 칸보다 적으면 돌려 가며 채운다. Callouts의 띠는 넷까지 |
| 굵은 외곽선 패널(네오브루탈리즘) | `brutal_panel.ps1 -Title SITUATION -Items '머리;줄;줄\|머리;줄' -Image p.jpg` | — (사진 선택) | png | 그림자는 흐림 0·투명도 0. 항목은 셋까지 |
| 원이 퍼지며 다음 장으로 | `circle_reveal.ps1 -Title '표지\|제목' -Heading H -Items '머리;설명\|...'` | — | mp4 | 2장짜리. 뒤 장에만 있는 글은 원이 커지는 동안 서서히 나타난다 |
| 세로 연혁 | `timeline.ps1 -Items '2016;내용\|2019;내용' -Highlight 2 -Title T` | — (투명 PNG 선택) | png | 강조는 한두 곳만 |
| 뉴모피즘 카드 | `neumorph_cards.ps1 -Title OBJECTIVE -Cards '머리;설명\|...'` | — | png | 카드 하나가 도형 두 겹이다(도형은 그림자를 하나만 가진다). 배경은 밝은 회색 |
| 이름표 상자(식순·목차) | `label_box.ps1 -Label '식 순' -Items 'a\|b\|c' -Double` | — (배경 사진 선택) | png | 테두리는 제목 자리를 비운 열린 선이라 사진 위에서도 된다 |
| 완성 장표 양식 38종 | `design_page.ps1 -Layout Radar -Title T -Items '기획;80\|디자인;95'` (`Window/Lines/Index/Dots/Radar/HalfDonut/Leaders/Process/Plus/Stack/Pills/Ring/GlassFade/Runner/Scribble/Sparkle/Bento/NeonCard/Rows/Select/Focus/FadePhoto/Magazine/Slash/Poster/Emblem/SideBands/Pastel/Diamonds/GlassBar/Split/Pierce/Drawer/Launch/TiltCover/SlantPage/Swot/Profile`) | 양식에 따라 사진 1장 | png (Drawer·Launch·`-Steps`는 mp4) | Select·Focus에 `-Steps`를 주면 항목이 차례로 켜진다(항목마다 한 장, 모핑). `-Items`의 뜻은 양식마다 다르다(스크립트 머리 주석). 항목은 `\|`, 항목 안은 `;`로 나눈다 |
| 마우스 오버 미로 게임 | `hover_maze.ps1 -Title MAZE` (`-Points 'x,y x,y …'`로 길 모양, `-Width`로 난이도) | — | png | 4장짜리(시작·미로·실패·성공). 시작·재시도 버튼은 출발점 자리에 있다(누른 자리가 벽 위면 바로 실패하므로). 벽 도형의 마우스 오버 링크라 스크립트로는 동작을 시험할 수 없다(링크 설정만 확인) |
| 왼쪽에서 날아오는 카드 | `fly_panels.ps1 -Title 'A\|B' -Items '북미;설명;도달:83;전환:42\|…' -Lead 88` | — | mp4 | 한 장 안의 애니메이션(날아오기, 이전 효과 다음에). 카드마다 그룹으로 묶여 함께 움직인다 |
| 색인 탭 목차 | `tab_chapters.ps1 -Items '기획;설명\|디자인;설명\|…' -Title SLIDES` | — | mp4 | 챕터 수만큼 장이 생긴다(모핑). 탭은 카드와 통합한 한 도형이다 |
| 색으로 칠한 지도 | `python make_choropleth.py countries-110m.json data.csv out --label China India` → `map_chart.ps1 -Svg out.svg -Title T -Tag Total -Unit Mt` | 나라 이름·값 CSV + world-atlas TopoJSON | png | 나라마다 이름 붙은 도형(세계 지도 약 280개)이라 빌드가 1~2분 걸린다. SVG를 넣고 `도형으로 변환`을 스크립트로 누르면 멈췄다(2026-10-08) — 그래서 처음부터 도형으로 그린다. 좁은 나라의 값 글자는 겹칠 수 있다 |
| 빈 곳 찾기 점검 | `balance_check.ps1 -Deck a.pptx -Slide 3 -Fill icon.png` | 점검할 pptx | png | 원본은 건드리지 않는다. 내용의 바깥 테두리를 3×3으로 나눠 빈 칸을 표시하고, `-Fill`이면 첫 빈 칸에 그림을 놓은 장을 앞에 둔다. 겹친 도형은 면적을 두 번 센다 |
| 한 장 인포그래픽 17종 | `info_page.ps1 -Layout Scale -Title T -Items '자동화 업무;75\|수작업 업무;25'` (`Pictogram/DiamondGrid/Phone/Scale/AreaChart/FadeBars/ProcessArrow/SlantChart/OrgChart/Donut3D/PhotoTimeline/ProCon/SignPost/FadeCards/StepArrows/Compare/Folders`) | 양식별(스크립트 머리말). 사진·아이콘은 모두 선택 | png (`Folders`는 mp4) | 색·배경 기본값이 양식마다 다르다(`-Accent`, `-Palette`로 바꾼다). `SlantChart`의 막대는 도형이고 차트 개체가 아니다. `SignPost`의 기둥은 3차원 서식이 아니라 면 세 개다(영상의 회전값을 넣으면 기둥이 기운다) |
| 발표용 장면 7종 (자동 애니메이션) | `scene_page.ps1 -Layout Ring -Kicker 'WEEKLY ACTIVE' -Title '매주 10명 중 8명이\|앱을 엽니다' -Items 83 -Sub '5주 평균'` (`DayFlow/Dashboard/HeroChart/Ring/People/Voice/Roadmap`. 항목 형식은 스크립트 머리말) | Voice만 사진 | mp4 | 한 장 안의 애니메이션이 ‘이전 효과 다음에’로 이어져 클릭 없이 재생된다(Voice는 2장 모핑). 색은 `-Accent`·`-Light`·`-Dark` 세 값에서 섞어 쓴다. 본문 15pt 이상, 핵심 숫자 76~190pt. `-HeavyFont`에 Black 굵기 서체를 주면 숫자가 더 세다 |
| 임팩트 장면 4종 (모핑) | `impact_scene.ps1 -Layout Kinetic -Items '보이게\|칭찬하게\|자라게' -Cards '한 줄\|한 줄\|한 줄'` (`Spotlight/Kinetic/Carousel/Bento`. 항목 형식은 스크립트 머리말) | Spotlight만 사진 | mp4 | 단계마다 한 장(Bento는 한 장). 모든 도형을 첫 장에 만들고 복제한 뒤 장마다 위치·크기·색만 바꾼다. Kinetic은 저절로 넘어가고 Spotlight·Carousel은 클릭마다 한 단계. 색·서체 매개변수는 `scene_page.ps1`과 같다(공용 `scene_helpers.ps1`) |
| 임팩트 장면 8종 더 (모핑) | `impact_morph.ps1 -Layout StatSeries -Items '6;명;매일 씁니다\|55;건;5주 동안'` (`Flip/Parallax/CutAway/Accordion/Wave/Orbit/ArcList/StatSeries`. 항목 형식은 스크립트 머리말) | CutAway 그림 2장, Accordion 사진 3~5장 | mp4 | Flip·CutAway·Wave는 첫 클릭 뒤 저절로 끝까지 재생, 나머지는 클릭마다 한 단계. 그림의 보이는 폭은 `PictureFormat.Crop.ShapeWidth`로 바꾼다(모핑이 부드럽게 잇는다). 모서리 둥근 사각형의 폭을 크게 바꾸면 둥글기 조절값을 다시 줘야 한다 |
| 영상에서 옮긴 장면 19종 | `video_scene.ps1 -Layout ZoomThrough/Bubbles/Showcase|CircleFan|FanTimeline`, `video_morph.ps1 -Layout Cast/FoldPanels/RingDial/StepFan/YearReel`, `video_craft.ps1 -Layout SplitTitle/GlassRise/Loupe/Lotus/Bars3D/FlightPath/RegionZoom/Interlock/Ladder` (항목 형식은 각 스크립트 머리말. 이 중 9종은 html-slide-effects의 `scene`에 웹 버전이 있다) | 레이아웃마다 다름(머리말). Cast는 배경 없는 인물 PNG, 지도 두 종은 `assets/world_110m.json`(Natural Earth, 퍼블릭 도메인) 동봉 | mp4 | 조각은 돌아도 사진은 바로 서게 하려면 `Fill.RotateWithObject = 0`. 회전한 도형에 이동 경로를 주면 경로도 같이 돌아가므로 처음부터 향할 방향으로 그린다. 소문자 `$w`·`$h`는 슬라이드 크기 `$W`·`$H`를 덮어쓴다(PowerShell은 대소문자 구분 없음) |
| 쇼츠에서 옮긴 장면 14종 | `shorts_scene.ps1 -Layout EraRoll -Title T -Items '2026;기반 구축;설명\|2028;시범 운영;설명'` (`WaveSplit/TabDeck/WaveWindow/RingWindows/KeywordLinks/GridCells/MosaicRise/EraRoll/ArchColumns/LensTour/DiamondBricks/PaperPeel/TornReveal/FoldArrows`. 항목 형식은 스크립트 머리말) | 레이아웃마다 다름(머리말). 배경 없는 그림은 필요 없다 | mp4 | 모두 클릭 없이 재생된다. 14종 모두 html-slide-effects의 `scene --layout <소문자 이름>`에 웹 버전이 있다. 영상의 기법에 한 가지씩 더한 것이 머리말의 ‘-’ 뒤에 적혀 있다. RingWindows는 슬라이드 배경 사진을 화면에 맞춰 늘리므로 16:9 사진을 쓴다. LensTour·WaveWindow의 위치(fx, fy)는 슬라이드 기준 0~1이니 빌드한 그림에서 확인한다. 매개변수와 같은 이름의 변수(`$Cards`와 `$cards`)는 같은 변수다 — 새 레이아웃을 넣을 때 주의 |
| 공공·행사·표지·차트 장면 53종 | `gov_frame.ps1`(비전 체계도·과제 개요·조직도·로드맵 8), `gov_data.ps1`(성과·보고·AI 정책 9), `cover_motion.ps1`(정책 표지·목차·해외 진출·국내 지도 핀·웹앱 소개·제안서 8), `event_page.ps1`(백드롭·대기·식순·환영·연사·시상·세션 8), `morph_cover.ps1`(모핑 표지·목차·마무리 9), `info_motion.ps1`(인포그래픽 5), `chart_motion.ps1`(차트·대표 수치 6). 모두 `-Layout <이름>`, 항목 형식은 각 스크립트 머리말. AppIntro·Tilt3D를 뺀 51종은 html-slide-effects의 `scene --layout <소문자 이름>`으로 웹 버전을 만든다 | 레이아웃마다 다름(머리말). 지도는 `assets/world_110m.json`·`assets/korea_sido.json`(통계청 2013 경계, 독도 없음) 동봉 | mp4 | 샘플마다 배색·서체가 다르다: `-Accent/-Light/-Dark/-Ink/-Font/-HeavyFont`로 바꾼다. 공공 장표의 기관명·수치는 예시이니 실제 자료로 바꾸고 `-Note`에 출처를 적는다. 사용자 지정 효과는 `Behaviors.Add` 뒤에 `Timing`을 건드린다(순서가 바뀌면 재생 중 도형이 안 보인다). 좌표 인자는 `[double]`로 선언한다(음수가 문자열로 들어간다) |
| 찢어진 사진 컷 | `torn_photo.ps1 -Style Split -Images p.jpg -Title 'BEFORE / AFTER'` (`Edge/Split/Strips/Scraps`, `-Seed`로 찢긴 모양, `-Rough`로 거칠기, `-Pieces`로 조각 수) | 사진 1장(`Scraps`는 3장까지) | png | 자유형 도형과 사진의 교차라 결과는 그림이다. 흰 종이 테두리는 그림 윤곽선(3pt), 들뜬 느낌은 그림자. 조각은 자른 뒤에 옮긴다(먼저 옮기면 사진이 이어지지 않는다) |
| 막대 여러 개에 사진 한 장 | `stripe_photo.ps1 -Image p.jpg -Style Diagonal -Title CITY -Body 'a\|b'` (`-Style Vertical`) | 사진 | png | 막대들을 통합한 뒤 사진과 교차한 것이다 |
| 콰직 충격 줌(폭발) | `impact_punch.ps1 -Image p.jpg -Text RUN -Cutout cut.png` | 사진 (+ 배경 제거 PNG) | mp4 | 3장짜리. 둘째 장은 0.12초 전환 뒤 저절로 넘어간다. 제목은 도형이라 수정할 수 없다 |
| 색 패널이 닫히며 인물이 다가오기 | `color_close.ps1 -Image p.jpg -Cutout cut.png -Text FUTURE` | 사진 + 배경 제거 PNG | mp4 | 2장짜리. 제목은 도형이라 수정할 수 없다 |
| 창이 열리는 사진(알약 창 → 전면 → 반쪽) | `window_morph.ps1 -Image sky.jpg -Object plane.png -Title T -Sub s -Bullets 'a\|b'` | 사진 (+ 투명 PNG 개체) | mp4 | 3장짜리. 사진의 자르기 틀만 바뀐다 |
| 커튼 인트로(색 패널이 쓸고 지나감) | `curtain_sweep.ps1 -Image p.jpg -Words 'TRAVEL\|TO\|THAILAND' -Object plane.png` | 사진 (+ 투명 PNG 개체) | mp4 | 단어 수 + 1장. 마지막 단어가 사진 위 제목으로 남는다 |
| 줌 맵(전체 그림 → 주제로 확대) | `zoom_map.ps1 -Image p.jpg -Title T -Nodes 'a\|b\|c' -Bodies 'x\|y\|z'` | 사진 | mp4 | 주제 수 + 1장. 쇼에서 원을 클릭하면 확대, 아무 데나 클릭하면 돌아온다. 파워포인트의 확대/축소 개체가 아니라 링크 + 모핑이다 |
| 절반 띠 뒤집기 | `half_band.ps1 -Images a.png,b.png,c.png -Title T -Sub s -Captions 'a\|b\|c'` | 투명 PNG 2~5장 | mp4 | 2장짜리 |
| 이미지 슬라이더(알약 창 + 썸네일) | `image_slider.ps1 -Images a.jpg,b.jpg,c.jpg -Title T -Body 'a\|b'` | 사진 3~6장 | mp4 | 사진 수만큼 장이 생긴다. 어두운 사진은 창이 잘 안 보인다 |
| 용어 풀이 말풍선(마우스 오버) | `hover_note.ps1 -Before 앞 -Word 용어 -After 뒤 -Note '풀이\|둘째 줄'` | — | png | 슬라이드쇼에서만 동작한다. 미리보기는 풀이가 뜬 장이다 |

```powershell
$sk = "<SKILL_DIR>\scripts"
& "$sk\long_shadow.ps1" -Text '피피티|고수 만들기' -Font '맑은 고딕' -Color '#FFB020' -Background '#0A63D8' -Out C:\work\cover.pptx -Preview C:\work\cover.png
& "$sk\text_slice.ps1" -Text '반값' -Font '맑은 고딕' -Out C:\work\slice.pptx -Preview C:\work\slice.mp4
```

패턴 GIF가 필요하면 `python "<SKILL_DIR>/scripts/make_pattern_gif.py" C:/work/pattern.gif` (640×480, 3초 무한 반복).
색·개수·속도는 파일 위쪽 상수를 고친다. 모든 움직임이 한 주기에 정수 번 돌아야 이음매가 없다.

### 여러 레시피를 한 번에 빌드할 때

PowerShell 프로세스 하나에서는 빌드를 하나만 돌린다(이유는 "함정"). 연속으로 만들려면 빌드마다 자식 프로세스를 띄운다.

```powershell
foreach ($cmd in "& '$sk\glow.ps1' -Text A -Out C:\work\a.pptx", "& '$sk\glow.ps1' -Text B -Out C:\work\b.pptx") {
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command $cmd
}
```

`-File` 방식은 배열 매개변수(`-Stops`, `-Card`)를 못 넘기므로 `-Command`를 쓴다. 한글 문구는 자식 프로세스에
문자열로 넘기지 말고, 그 빌드만 도구 호출에서 직접 `& "$sk\..."`로 실행한다.

## 모핑 — 애니메이션 대신 쓰는 기본기

모핑은 **앞 슬라이드와 뒤 슬라이드에서 이름이 같은 요소**의 위치·크기·회전·색 차이를 자동으로 이어 준다.
그래서 만드는 순서는 늘 같다: 한 장을 완성한다 → `Add-MorphSlide <초>`로 복제한다 → 둘 중 한쪽의 요소를 옮기거나 바꾼다.

- 등장시키려면 앞 슬라이드에서 그 요소를 슬라이드 밖에 둔다. 사라지게 하려면 뒤 슬라이드에서 밖으로 뺀다.
- 한쪽에만 있는 요소는 서서히 나타나거나 사라진다(`image_zoom.ps1`의 지점별 문구가 이 방식).
- 복제하면 이름이 그대로 유지되므로 짝짓기는 저절로 된다. 따로 만든 두 요소를 짝지으려면 양쪽 이름을 `!!같은이름`으로 맞춘다.
- 전환 시간: 글자 갈라짐 같은 타격감은 0.3초, 3D 회전 0.5초, 사진 이동·패널 등장은 1초가 영상들의 값이다.

## 작업 방식과 유의사항 — `references/ppt-craft.md`

효과를 고르기 전에, 그리고 사용자에게 손으로 만드는 법이나 주의점을 설명해야 할 때 읽는다.
겹침 순서로 깊이를 만드는 법, 도형 병합의 선택 순서, 모핑이 어색해지는 지점(겹침 순서 불일치, 배경이 안 움직임),
사진 위에 글자를 얹는 세 가지 방법, 그림자의 세 가지 쓰임, 가상의 네모 배치, 포스터에서 색 뽑기,
재료 구하는 곳, 그리고 사용자에게 꼭 알려야 하는 한계(수정 불가, 버전·서체 의존)가 들어 있다.
여러 장짜리 덱을 스크립트 하나로 짓거나 효과음을 넣을 때는 29장(덱 뼈대·시간 설계·효과음·파일 관리)을 먼저 읽는다.
항목마다 직접 빌드해 확인한 것(✅)과 영상 내용 그대로인 것(📺)을 구분해 두었으니, 📺 항목을 근거로 말할 때는 그렇게 밝힌다.

## 새 기법 만들기

레시피에 없는 효과는 `ppt_helpers.ps1`을 dot-source한 새 스크립트로 만든다. 기존 레시피가 본보기다.
영상의 클릭 순서를 그대로 옮기면 된다.

| 파워포인트 조작 | COM 호출 |
|---|---|
| 도형 병합(통합 1 / 교차 3 / 빼기 4) | `Merge-Shapes <모드> @(먼저 선택한 도형 이름, 나머지…)` — 결과 도형을 돌려준다 |
| 도형 병합 → 조각 | 조각(5) 대신 같은 도형 쌍에 교차 한 번, 빼기 한 번을 한다(조각을 골라 지울 필요가 없다. `text_slice.ps1`) |
| 글자를 도형으로 | `New-GlyphShape $textbox` (글자와 큰 사각형의 교차). 정확한 글자 경계도 이걸로 잰다 |
| 글자 한 자의 위치 | `$t.TextFrame.TextRange.Characters(i,1).BoundLeft / .BoundWidth` |
| SVG `도형으로 변환` | 창 활성화 → `$pic.Select()` → `$app.CommandBars.ExecuteMso('SVGEdit')` → 그룹 해제 |
| 사진을 슬라이드 가득 | `Add-Cover <경로>` |
| 그림을 도형 모양으로 자르기 | `Merge-Shapes 3 @(그림 이름, 도형 이름)` — 그림을 먼저 줘야 결과가 그림(Type 13)으로 남는다 |
| 꾸밈 효과 → 흐리게 | `$e = $pic.Fill.PictureEffects.Insert(2); $e.EffectParameters(1).Value = <반경>` |
| 그림 밝기 | `$pic.PictureFormat.Brightness` (0.5가 원본) |
| 그라데이션 채우기 | `$r.Fill.TwoColorGradient(1,1)` → `.GradientStops`의 `Insert(색, 위치, 투명도, 순번)`, 각 중지점 `.Color.RGB/.Position/.Transparency`, 방향은 `.GradientAngle` |
| 방사형(원) 그라데이션 | COM으로는 사각형 경로(`TwoColorGradient(7,1)` → `path="rect"`)만 된다. 저장 → 슬라이드 XML의 `path="rect"`를 `"circle"`로 바꿈 → 다시 열기 (`gradient_scrim.ps1`의 Center 분기) |
| 글자에 그라데이션 | `$t.TextFrame2.TextRange.Font.Fill`에 도형과 같은 방식(`TwoColorGradient` → `GradientStops` → `GradientAngle = 0`이 선형 오른쪽) |
| 영상 삽입·자동 반복 | `$sl.Shapes.AddMediaObject2(경로,0,-1,l,t,-1,-1)` → `.AnimationSettings.PlaySettings`의 `PlayOnEntry`, `LoopUntilStopped` = -1, `AdvanceMode = 2`. 여러 개면 `TimeLine.MainSequence`의 둘째부터 `Timing.TriggerType = 2`(이전 효과와 함께) |
| 영상을 글자·도형 모양으로 | `Merge-Shapes 3 @(영상 이름, 도형 이름)` — 영상을 먼저. 결과에 재생 설정을 다시 건다 |
| 그림자(바깥 / 안쪽) | `$shape.Shadow` — `.Style = 1`이면 안쪽. 글자 그림자는 `$t.TextFrame2.TextRange.Font.Shadow` |
| 글자 3차원 서식·회전 | `$t.TextFrame2.ThreeD` — 설정 순서는 `long_shadow.ps1` 주석 참고(순서가 틀리면 카메라가 초기화된다) |
| 3D 모델 | `$sl.Shapes.Add3DModel(경로,0,-1,l,t,w,h)` → `.Model3D.RotationX/Y/Z` |
| 맨 뒤 / 맨 앞 | `$shape.ZOrder(1)` / `$shape.ZOrder(0)` |
| 모션 경로 | `AddEffect($s,0,0,2)` 후 `Behaviors.Add(1).MotionEffect.Path = "M 0 0 L 0 -0.15 E"` (슬라이드 크기 대비 비율) |
| 나타내기 닦아내기(아래→위) | `AddEffect($s,22,0,2)` + `EffectParameters.Direction = 3` |
| 슬라이드 복제 + 모핑 | `Add-MorphSlide <초>` (전환 `EntryEffect = 3954`) |
| 부분 원형(원 그래프 조각) | `AddShape(142, …)` → `.Adjustments(1)`=시작각, `(2)`=끝각 (도, 시계 방향, 0=3시) |
| 글자 윤곽선 / 채우기 없음 | `$t.TextFrame2.TextRange.Font.Line`(`.Visible/.ForeColor/.Weight`), `.Font.Fill.Visible = 0` |
| 글자 그림자(선명 / 부드럽게) | `.Font.Shadow` — `.Blur = 0` + `.OffsetX/.OffsetY` / `.Blur`를 키우고 오프셋 0 |
| 도형 안 글자 줄바꿈 해제 | `$shape.TextFrame.WordWrap = 0` |
| 슬라이드 배경 채우기 | `$shape.Fill.Background()`, 배경 사진은 `$sl.Background.Fill.UserPicture(경로)` |
| 그룹 + 회전 | `$sl.Shapes.Range(@(이름…)).Group()` → `.Rotation` (모핑에서는 장마다 늘어나는 값으로) |
| 실행 설정(클릭 1 / 마우스 오버 2) | `$s.ActionSettings(2).Action = 7; .Hyperlink.SubAddress = "<SlideID>,<번호>,"` · 마지막으로 본 슬라이드는 `.Action = 5` |
| 효과음을 애니메이션에 맞춰 재생 | `AddMediaObject2(mp3,0,-1,0,0,-1,-1)` → 아이콘은 `.Left = [single](-80)` → `MainSequence.AddEffect($m, 83, 0, 2)` + `Timing.TriggerDelayTime`. 음량을 먼저 맞추고, 전환 소리는 영상에 안 실린다(`ppt-craft.md` 29장) |
| 사용자 지정 효과로 시작하는 도형을 처음부터 보이게 | 0초에 나타내기를 먼저: `AddEffect($s, 1, 0, 2)` (없으면 이동·크기 효과가 시작될 때까지 화면에 없다) |
| 책장 넘김 전환 / 영상에서 장 끝에 머물기 | `SlideShowTransition.EntryEffect = 3946` / 보이지 않는 도형에 지연을 준 효과를 마지막에 하나 |
| 표 | `AddTable(행,열,…).Table` → `.ApplyStyle('{2D5ABB26-0587-4C30-8999-92F81FD0307C}')`(스타일 없음) → `.Cell(r,c).Borders(1 위/3 아래)` |
| 로고의 단색 배경 지우기 | `$pic.PictureFormat.TransparentBackground = -1; .TransparencyColor = <BGR>` (실제 로고로는 미검증) |
| 저장 + 미리보기 | `Save-Deck $out $preview <슬라이드당 초>` — 미리보기가 `.png`면 정지 그림, `.gif`면 애니메이션 GIF, 그 밖은 mp4 |

값을 모르는 열거형 상수는 추측하지 않는다. 후보 값을 차례로 넣고 저장한 pptx의 XML
(`ppt/slides/slide1.xml`의 `prst="..."`)을 읽으면 어떤 번호가 어떤 이름인지 바로 나온다.
롱 섀도 카메라 번호(40/42/45/47)는 그렇게 찾았다.

새 스크립트도 끝에 `Save-Deck`으로 미리보기를 뽑고, 5번 검증을 통과한 뒤에 전달한다.

## 함정

- **한 PowerShell 프로세스에서 빌드를 두 번 돌리지 않는다.** 파워포인트는 프로세스가 끝나야 완전히 종료되는데,
  같은 세션의 두 번째 빌드는 반쯤 닫힌 인스턴스에 붙어 엉뚱한 오류로 죽는다. `New-Deck`이 이전 인스턴스가
  꺼지기를 최대 10초 기다리지만, 프로세스를 나누는 것이 원칙이다.
- **좌표를 넣다가 "Specified cast is not valid"가 나면 `[single](...)`로 감싼다.** `if` 식의 결과,
  다른 도형에서 읽은 값, 그림의 `Width`에 넣는 실수에서 실제로 났다(2026-10-08). 같은 줄에서 매번 나므로
  우연이 아니다. 리터럴과 단순 산술 결과는 대개 그냥 들어간다.
- **도형 병합의 결과는 맨 위 도형이 아닐 수 있다.** 기준 도형의 순서 자리를 물려받는 경우가 있어서
  `Shapes(Shapes.Count)`로 집으면 다른 도형을 건드린다. `Merge-Shapes`는 새로 생긴 Id로 찾는다.
- **복제한 도형은 직접 붙인 이름까지 물려받는다.** 이름으로 `Range()`를 잡으면 원본이 잡혀서 원본이 병합에 빨려 들어간다
  (`gold_award.ps1`에서 겪음). 복제 직후 고유한 이름을 준다. `New-GlyphShape`는 그렇게 한다.
- 내려받은 mp4가 "PowerPoint cannot insert the file"로 거부되면 ffmpeg로 오디오 없이 다시 인코딩한다
  (`-an -c:v libx264 -pix_fmt yuv420p`). 같은 영상이 그렇게 하자 들어갔다. 원인은 확인하지 못했다.
- PowerShell 변수는 대소문자를 구분하지 않는다. `$S`와 `$s`는 같은 변수다. `Try`처럼 예약어인 함수 이름도 못 쓴다.
- **한글 서체는 `Font.NameFarEast`로 들어간다.** `Font.Name`만 주면 영문·숫자에만 적용되고 한글은 얇은 기본 서체로 나온다
  (2026-10-08 실제로 겪음, `New-Text`가 둘 다 설정하도록 고침). 직접 글상자를 만들 때도 둘 다 넣는다.
  이 PC에 있는 두꺼운 한글 서체: `Paperlogy 9 Black`, `Gmarket Sans TTF Bold`, `Pretendard Black`, `NanumSquare Neo Heavy`, `Noto Sans KR Black`.
- 획이 겹치게 그려진 서체(Paperlogy 등)는 글자를 도형으로 바꾸면 겹친 자리에 가는 선이 남는다. `text_fill.ps1`의 윤곽선에서
  글자 안쪽에 실선이 보이고, `text_slice.ps1`은 갈라지기 전 장에 실금이 비친다. 거슬리면 서체를 바꾼다(`Gmarket Sans TTF Bold`는 깨끗했다).
  실제로 `$h = New-FitText …`가 슬라이드 높이 `$H`를 덮어써 뒤의 좌표 계산이 죽었다(2026-10-08). 헬퍼의 `$W`, `$H`, `$sl`, `$pres`, `$app`과 겹치는 한 글자 변수를 쓰지 않는다.
- **마우스 오버 동작은 스크립트로 시험하지 못한다.** 백그라운드에서 띄운 슬라이드쇼 창은 활성화되지 않아 포인터를 옮겨도 반응이 없었다.
  저장된 pptx의 슬라이드 XML에서 `hlinkHover`/`hlinkClick`을 읽어 설정이 들어갔는지만 확인하고, 실제 동작은 사용자가 쇼에서 봐야 한다고 알린다.
- **확대/축소(Zoom) 개체는 COM으로 만들 수 없다**(`Shapes`에 해당 메서드가 없음, 2026-10-08). `zoom_map.ps1`은 같은 모습을 링크와 모핑으로 만든다. 진짜 확대/축소가 필요하면 손으로 넣는 절차를 ppt-craft 24장에서 안내한다.
- `$home`은 PowerShell 예약 변수라 슬라이드를 담을 수 없다. 슬라이드를 `Duplicate()`하면 사본은 맨 끝이 아니라 원본 바로 뒤에 생긴다.
- COM 색은 BGR이다. `ConvertTo-Bgr '#RRGGBB'`를 쓴다.
- `.ps1`은 ASCII로만 쓴다(PowerShell 5.1이 BOM 없는 파일을 ANSI로 읽는다). 한글 문구는 매개변수로 넘긴다.
- 스크립트가 중간에 죽으면 저장 안 된 프레젠테이션이 남을 수 있다. 레시피는 `catch`에서 닫지만, 직접 짠 스크립트도 같은 처리를 한다.
- 사용자가 파워포인트를 열어 둔 상태일 수 있다. 내가 만들지 않은 프레젠테이션은 닫지 않고, 열린 문서가 남아 있으면 프로그램을 종료하지 않는다.

## 재료와 권리

- **배경 제거 이미지는 이 스킬이 만들지 않는다.** 영상들도 외부 도구(Adobe 배경 제거 등)를 쓴다. 사용자가
  원본과 같은 크기·같은 구도의 투명 PNG를 주어야 하고, 없으면 그렇게 요청한다. 파워포인트의 배경 제거는
  사람이 영역을 손봐야 해서 스크립트로 돌리지 않는다.
  사용자가 "사이트에서 받아서 해 달라"고 하면 Aside 브라우저로 할 수 있다(2026-10-08 확인, Adobe 계정이 로그인된 상태):
  adobe.com의 Firefly "Remove Background" 페이지 파일 입력에 사진을 넣으면 편집기 탭이 새로 열리고 15~25초 뒤 배경이 지워진다.
  오른쪽 위 다운로드 버튼은 텍스트로 안 잡혀 좌표로 눌렀다. 결과 PNG는 원본과 같은 크기·같은 구도였다.
  Aside REPL의 `setInputFiles`는 세션 폴더 안 상대 경로(`./tmp/...`)만 받으므로 사진을 먼저 REPL 안에서 `fetch`해 저장한다.
- 무료 사진은 Unsplash에서 로그인 없이 받을 수 있다(검색 `/napi/search/photos?query=...&plus=none`, 받기
  `images.unsplash.com/photo-<id>?w=1920&h=1080&fit=crop`). 16:9로 잘라 받으면 슬라이드에 그대로 맞는다. 출처를 기록해 함께 전달한다.
- 3D 모델은 GLB로 받는다. Sketchfab 등은 모델마다 라이선스가 다르니 사용자가 고른 것을 쓴다.
  Sketchfab과 Freepik은 내려받기에 로그인이 필요하다. 로그인되어 있지 않으면 대신 로그인하지 말고 사용자에게 알린 뒤,
  로그인 없이 받을 수 있는 공개 자료(예: Khronos glTF-Sample-Assets, CC BY)로 대체할지 묻거나 그렇게 했다고 밝힌다.
  가로로 긴 모델은 `-To`의 크기 값을 줄이지 않으면 슬라이드 밖으로 나간다(크기 값은 높이다).
- 남의 로고는 상표다. 따라 하기 연습용으로는 만들되, 외부 공개용이면 사용자가 권리를 확인하도록 알린다.
- GIPHY 등에서 가져온 GIF는 대부분 작가 저작물이고 사용 조건이 불분명하다. 사용자가 직접 골라 준 것이 아니면
  라이선스가 적힌 자료(예: 위키미디어 커먼즈)나 직접 만든 패턴을 쓰고, 영상과 다른 출처를 썼다면 그 사실을 먼저 말한다.
- 출처 표기가 필요한 라이선스는 슬라이드 노트에 적는다(`gif_in_shape.ps1 -Notes`).
- 브랜드 색은 공식 값을 확인하지 못했으면 근사값이라고 밝힌다.
