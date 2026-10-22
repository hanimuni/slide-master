> See [`strategist.md`](strategist.md) §6.2 for where these rules enter the outline and [`executor-base.md`](executor-base.md) §2 for where they enter each page.

# Report Deck Craft Specification

Canonical page-level rule set for Korean report decks: the title / governing-message pair, governing-message wording, evidence integrity, list limits, chart emphasis, footer, machine-readable structure, and the final read-through. Strategist applies it while authoring `design_spec.md` §V / §VII / §IX; Executor applies it on every content page.

**Trigger**: every main-pipeline deck. **Exempt**: verbatim preservation routes ([`beautify-pptx`](../workflows/beautify-pptx.md), [`ppt-template-fill`](../../ppt-template-fill/SKILL.md)) and user-authored titles / wording, which stay as given.

**Precedence**: this file owns the rules below. Other references point here instead of restating them; when a mode file or template default differs on these points, this file wins.

---

## 1. Title Structure — Topic Title + Governing Message

**Hard rule**: every content and data page uses the two-layer Korean report structure, deck-wide. Never mix it with single action-title pages in one deck.

| Element | Role | Spec source | Rule |
|---|---|---|---|
| Topic title (주제 제목) | Names the page's subject, tied to the contents numbering | §IX `Title` | Noun phrase, no conclusion, one line (e.g. `Ⅱ-1. 제도 이용 실태`) |
| Governing message (거버닝 메시지) | The page's one message | §IX `Core message`, rendered verbatim | One sentence directly under the title in the header zone; wording per §2 |

- **Terminology**: *governing message* is the page-level line and is the same object as the §IX `Core message`. The deck-level answer confirmed at strategist §c is the *deck core message* (덱 핵심 메시지). Do not introduce a separate "headline" or "lead" element — the `lead` typography slot is where the governing message renders.
- **Voice by mode** — the title stays a topic label in every mode; the mode sets the governing message's voice:

| Mode | Governing message voice |
|---|---|
| `pyramid` | Assertion the body proves (§2 pattern) |
| `briefing` | Coverage statement with the key figure ("부처별 정원 현황 — 총 1,240명, 전년 대비 3% 증가") — no persuasion |
| `narrative` | The story beat this page advances |
| `instructional` | What the page teaches, as one sentence |
| `showcase` | One evocative sentence about the featured item |
| `custom` | As written in `mode_behavior` |

- **Exempt pages** (header pair and footer alike): cover, contents, section divider, closing, and `breathing` pages. They may carry the message as hero text instead of the header pair. A section divider may show the section's key-line claim as its subtitle.

---

## 2. Governing Message Wording

**Hard rule — length**: follow the copy-length budget in [`strategist.md`](strategist.md) §6.2 — one line by default; a balanced two-line message only when §IX plans it; never three lines. Over budget → split the page or move qualifiers into the body. Never shrink the type.

**Default — pattern (may override for `briefing` coverage lines)**: indicator + direction + magnitude + period + cause or implication.

| Weak (topic / activity / deferral) | Strong (governing message) |
|---|---|
| ○○제도 운영 현황 | ○○제도 이용률 3년간 12%→5% 하락, 신청 절차 복잡성이 주원인 |
| 관계부처 의견 수렴 결과 | 5개 부처 중 4곳 찬성, 기재부는 재원 분담 방식 보완 요구 |
| 개선방안은 다음과 같음 | 신청 서류 7종→3종 축소 시 처리 기간 절반으로 단축 가능 |

**Forbidden — governing-message wording**:
- Deferral phrases: "다음과 같음", "아래와 같음", "~에 대해 보고드림"
- Topic-only phrases standing in for a message: "○○ 현황", "○○ 추진 상황", "검토 결과"
- Undefined abstract targets: "최적화", "활성화", "효율화", "고도화" without the measurable definition fixed at strategist §c
- Two claims joined by "및" / "그리고" / "~하고" → split into two pages
- Weak predicates where a change verb exists: "있음", "보여줌", "나타남", "관련됨" → 증가·감소·상승·하락·초과·미달·단축·확대·축소·전환·차단
- Stacked hedges without the evidence ("~할 수도 있을 것으로 보임") — state the evidence, or name the uncertainty specifically
- Any message stronger than the source or the confirmed `evidence_rigor`

**Default — one grammatical form per peer set**: pages answering the same parent question end the same way (all noun-ending "~감소" or all sentence-ending "~감소함"). Record the deck's default form in §I; a peer set may deviate only as a whole.

---

## 3. Evidence Integrity — No Missing, No Surplus

**Hard rule** — on every content and data page:
- **No missing evidence**: each claim element of the governing message (every number, cause, implication) has visible support in the body.
- **No surplus evidence**: each body block supports the governing message. Anything else moves to another page, the appendix, or the speaker notes.

**Default — evidence mix**: combine at least two of quantitative data, an external case, analysis, and an authoritative source (institution or named expert). A claim without a source is marked `Needs-Manual`, never filled in.

**Summary page rule**: a summary page may list reasons without their evidence only when each reason reappears as a later page's governing message. Wording and numbers must match exactly.

---

## 4. Lists and 개조식

**Default — list limits (may override for reference tables and appendix pages)**: at most 4 items per list and at most 2 rendered lines per item. Beyond that, group the items under 2–3 labeled heads or convert them to a diagram or table. (Kosslyn et al., 2012: 91% of analyzed slideshows broke these limits; working memory holds about four units.)

**Hard rule — 개조식 keeps context**: each item keeps its period, magnitude, comparison base, and subject — "최근 3개월 매출 18% 감소(전년 동기 대비)", not "매출 감소". Changing endings to "~함" is not compression.

**Default — one symbol hierarchy**: record one hierarchy in §V and use it deck-wide (e.g. `□ → ○ → - → ※`), at most three levels on one page.

---

## 5. Takeaway Box and Sticker

**Default — takeaway box (on no more than ~20% of content pages)**: a bottom or side box adds exactly one of an implication, the next question, or a detail the chart hides. It never restates the governing message.

**Mandatory — sticker** when the page shows provisional, estimated, or illustrative data, or is circulated for discussion: a small tag at the top-right of the header zone reading `예비` / `추정` / `예시(가상 수치)` / `논의용`. Its position is part of the deck frame (§V).

---

## 6. Charts

**Mandatory — data intake from images and PDFs**: when chart values exist only as a chart image, table image, or PDF page, transcribe them into a data table (category, series, value, unit, period, source) in `design_spec.md` §VII before drawing. Transcribe during source analysis (before the Step 4 confirmation), print the table in the chat confirmation bundle for verification, then copy the confirmed values into §VII. Mark illegible or ambiguous values `Needs-Manual`; never estimate them.

**Default — single emphasis**: one accent color marks the data point or series that proves the governing message. Every other series uses the locked neutral tiers (light gray steps). When several series matter, keep one accent and separate the rest by neutral lightness.

**Forbidden — chart rendering**: 3D effects, exploded pies, tilted axes, more than 5 series (regroup the tail as "기타"), and decorative gradients on data marks.

**Mandatory — chart labeling**: the chart caption names the metric, unit, and period. The governing message carries the claim, so the caption does not repeat it. Axes carry units. Use a legend when there are 2 or more series. Add a source line per §7.

**Default — data-deck pattern (decks with 3 or more chart pages)**: open with a summary page (every chart as a small multiple with its page's governing message), then one chart per page — chart on the left (~60% width), and a right column titled `이슈 포인트` with 3 numbered points. Group chart pages by the conclusion they support, not by metric type.

**Hard rule — no color-only meaning**: every encoded difference also reads in grayscale (direct labels, value labels, or patterns).

---

## 7. Footer

**Hard rule — footer layout** (recorded as part of the §V deck frame; the §1 exempt pages may omit it):

| Position | Content | Format |
|---|---|---|
| Bottom-left | Notes in the order general note → numbered footnotes → source | `출처: 자료명(연도)`; several sources joined by `; `; own work `자체 분석`. Smallest structural size (`footnote` slot) |
| Bottom-right | Deck short title + page number | `분기별 사업 실적 · 01/05` — zero-padded current / total, no spaces around `/`; total = numbered pages |

**Hard rule — external data is always sourced**: every page with external figures carries its source line on the slide, never only in speaker notes.

---

## 8. Machine-Readable Structure

**Hard rule**: titles, governing messages, and every conclusion are live text — never baked into an image.

**Default — tables**: no merged cells unless the source table's meaning requires them; one header row (plus a header column when needed); one value per cell.

**Default — numbering**: one numbering scheme deck-wide that matches the contents page (e.g. `Ⅰ. → 1. → 가.`).

**Default — appendix for live questions** (when `desired_outcome` is `approve` / `decide` / `act`, or the briefing is live or discussion-format): plan an appendix after the closing page headed `참고` — expected questions, calculation basis, option comparison, detailed figures. The export has no hidden-slide flag; tell the user to hide these slides in PowerPoint if needed.

**Reference — not a constraint**: Korean central-government reporting is moving toward AI-readable documents (2026 MOIS pilot: no merged cells, Markdown attachments) and time-boxed briefings (15-minute reports). Details are published only through press reports; translate a stated time box into `time_budget_minutes`.

---

## 9. Delivery Versions and Fonts

**Default — read version first**: Korean public reporting is mostly read while being briefed, so recommend `delivery_purpose` `text` or `balanced`. When an oral version is also needed, derive a separate `presentation` deck afterward instead of one file serving both.

**Mandatory — PDF companion**: at handoff, tell the user to also save a PDF from PowerPoint with fonts embedded. PPTX does not embed fonts, and recipient PCs may not have Pretendard.

---

## 10. Final Read-Through (logic → facts → form)

**Mandatory — after `verify_deck.py` passes, before handoff.** Check in this order; fixing form first wastes the work when logic changes.

1. **Logic**: read only the governing messages from the first content page to the last. Check for missing steps, duplicates, and leaps. For `approve` / `decide` / `act`, the decision request appears early and on the closing page.
2. **Facts**: summary and detail pages agree to the number. Totals, percentage bases, units, and decimals are consistent. Every external figure has a source line. Figures you derived trace back to the source.
3. **Form**: no template sample text or placeholder residue, stickers where §5 requires them, grayscale legibility, footers and page numbers present.

**Output**: one line per stage in the handoff message — passed, or the page and issue. List anything left unverified.
