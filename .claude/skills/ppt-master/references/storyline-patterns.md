> See [`modes/pyramid.md`](modes/pyramid.md) for the argument rules and [`report-craft.md`](report-craft.md) for page-level wording.

# Storyline Patterns Reference Manual

Deck-level skeletons for six recurring Korean report types. Strategist picks one while authoring `design_spec.md` §IX when no user-supplied outline exists; the skeleton fixes page roles and order, then §IX fills each governing message from the source.

**Reference — not a constraint**: page counts and stubs show a complete storyline, not a quota. Keep the deck core message before its evidence, and drop any slot the source cannot support — never invent content to fill one. User-authored outlines and preservation routes are exempt.

---

## 1. Selection

| Report type | Use when | Logic order | Typical mode |
|---|---|---|---|
| 정책 제안형 (recommendation) | One recommendation needs approval | Inductive key line; reasons by strength or domain | `pyramid` |
| 대안 비교형 (decision brief) | Two or more options; leadership chooses | Recommendation first, options compared on fixed criteria | `pyramid` |
| 실태·분석 결과형 (findings) | Survey, audit, or data analysis results | Findings grouped by insight, not by method | `pyramid` / `briefing` |
| 중장기 전략·계획형 (strategy / plan) | Multi-year direction or annual plan | Vision → tasks → roadmap → resources | `pyramid` / `narrative` |
| 성과 데이터형 (data story) | Metric or performance readout, mostly charts | Charts grouped by the conclusion they support | `pyramid` / `briefing` |
| 업무보고형 (status / board) | Periodic report to minister, board, or committee | Domain order; flags and decisions isolated | `briefing` |

**Default — Korean contents order**: the common public-report contents (추진배경 → 현황 및 문제점 → 개선방안 → 향후계획) is an SCR variant. Keep it as the contents order when the user expects it, but write each section's opening page as a governing message, and check that the messages read as one chain (§3).

---

## 2. Skeletons

### 2.1 정책 제안형

| # | Page role | Governing-message stub |
|---|---|---|
| 1 | Situation | The accepted current state |
| 2 | Complication | What changed, or the gap that makes the current approach insufficient |
| 3 | Deck core message | "○○를 △△까지 추진해야 함 — [핵심 이유]" |
| 4 | Key line | The 3–4 reasons, as parallel short claims |
| 5–8 | One section per reason | Evidence pages under each reason's claim |
| 9 | Risks and mitigation | "위험 2건은 기존 자원으로 관리 가능" |
| 10 | Implementation | Steps with owner, timing, budget |
| 11 | Closing + decision request | The recommendation restated, then what to approve and by when |

### 2.2 대안 비교형

| # | Page role | Governing-message stub |
|---|---|---|
| 1 | Situation | The decision to be made and the deadline |
| 2 | Complication | Constraints and trade-offs that make it non-trivial |
| 3 | Deck core message | "A안 채택 권고 — 세 가지 제약을 모두 충족하는 유일한 안" |
| 4 | Criteria | The comparison criteria, fixed before the options |
| 5–7 | One page per option | Each option against the same criteria |
| 8 | Comparison | The comparison table plus the inference ("A안은 비용·위험 우위, B안은 속도만 우위") |
| 9 | Closing + decision request | The decision, the owner, the deadline |

**Hard rule**: the comparison page states the inference. A neutral table that leaves the comparison to the audience is incomplete.

### 2.3 실태·분석 결과형

| # | Page role | Governing-message stub |
|---|---|---|
| 1 | Context | What was studied, the scope, and why now |
| 2 | Deck core message | "세 가지 결과가 하나의 원인을 가리킴: [핵심 통찰]" |
| 3 | Key line | The findings, each as a claim |
| 4–12 | One section per finding | The section divider states the finding, then evidence pages |
| 13 | Implications | What the findings change |
| 14 | Next steps | Only when actions follow |

**Default**: section dividers name the inference ("현장 담당자는 지침보다 선례를 따름"), never the method ("심층면접 결과").

### 2.4 중장기 전략·계획형

| # | Page role | Governing-message stub |
|---|---|---|
| 1 | Situation + complication | Environment change and the current trajectory's limit |
| 2 | Deck core message | The strategy thesis in one directional claim |
| 3 | Framework | Vision → goals → strategic tasks (one structure page) |
| 4–10 | One section per strategic task | Each task's target and key projects |
| 11 | Roadmap | Phases named by what they deliver, not by activity |
| 12 | Resources | Bounded headcount and budget |
| 13 | Expected effect | Measurable outcomes, stated as numbers |

**Default**: "1단계: 기반 조성" is an activity; "1단계(~'27.6.): 표준 모델 확정·시범기관 5곳 선정" is a deliverable. Use the deliverable form.

### 2.5 성과 데이터형

Apply the data-deck pattern in [`report-craft.md`](report-craft.md) §6: a summary page of every chart with its governing message, then one chart per page with the `이슈 포인트` column. Order the sections by the conclusion each chart group supports, and put methodology and raw tables in the appendix.

### 2.6 업무보고형

| # | Page role | Governing-message stub |
|---|---|---|
| 1 | Period headline | The single most important status claim plus the one decision needed |
| 2 | Key line | 3 domains (e.g. 주요 성과 / 추진 현황 / 현안), each with a one-line status claim |
| 3–8 | One section per domain | Current position plus the one forward signal that matters |
| 9 | Flags | One dedicated page per risk that needs the reader's attention |
| 10 | Decisions required | "이번 보고에서 결정이 필요한 사항 N건" |

**Hard rule**: flags and decisions get their own pages. A risk sentence inside a performance page goes unseen.

**Default**: avoid "정상 추진 중" as a message. State the measured position ("12개 과제 중 10개 일정 내, 2개 1개월 지연 — 예산 이월 필요").

---

## 3. Headline-Track Check

**Mandatory — run on the finished §IX before Step 4 ends** (see `SKILL.md` Step 4):

1. List every content page as `P<NN> [topic title] — governing message`.
2. Read the governing messages alone, top to bottom. Fix missing steps, duplicates, and leaps between adjacent messages.
3. Within each peer set, apply the rules in [`modes/pyramid.md`](modes/pyramid.md) §1 (Peer sets): one parent question, one logical order, one grammatical form.
4. Summary pages: each summarized point reappears verbatim as a later governing message, with matching numbers ([`report-craft.md`](report-craft.md) §3).
5. Pages whose governing message cannot be stated → merge or cut. A transition sentence never gets its own page.
