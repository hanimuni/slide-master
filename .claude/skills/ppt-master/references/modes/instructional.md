# Mode: instructional

Teaching-led exposition. Decompose a concept into ordered, digestible parts and build understanding step by step. For training, tutorials, explainers, onboarding, science / knowledge sharing.

---

## 1. Narrative skeleton

**Decompose, then sequence**: break the subject into parts and present them in a deliberate order (simple → complex, prerequisite → dependent, overview → detail).

**One concept per page**: each page teaches a single idea well; do not stack unrelated concepts.

**Parallel exposition**: sibling concepts get parallel structure — same shape, same depth — so the audience can compare and map them.

**Show, then tell**: lead with a concrete example or analogy, then state the principle. A worked example beats an abstract definition.

**Signpost**: orient the learner — what we covered, what comes next.

Topic titles name the step or concept; the governing message states what the page teaches ("어텐션 가중치는 질의와 각 토큰의 유사도로 계산됨") — clear over clever ([`report-craft.md`](../report-craft.md) §1).

---

## 2. Page-structure tendencies

- Numbered steps / ordered flows for processes; parallel cards for sibling concepts.
- Diagrams that build incrementally; annotate the part currently being explained.
- A concrete example anchors each abstract point.

> Step / flow / diagram geometry lives in [`templates/charts/`](../../templates/charts/); this mode decides *the learning order and granularity*.

---

## 3. Speaker-notes register

Patient, explanatory. Define before using; analogy then principle. Anticipate the learner's question and answer it. Steady pace; signpost transitions ("now that we have X, we can ask Y"). Conversational data. (Common framework: [`speaker-notes.md`](../speaker-notes.md).)

---

## 4. Page skeleton example

```
Title:      "Step 2 — 토큰 점수 계산"
Governing:  "질의와 각 토큰의 내적이 곧 점수 — 클수록 주목도가 높음"
Body:   concrete example (3 tokens) → the rule it illustrates → one diagram
Notes:  "Remember the query from the last page? Here's what it does next…"
```
