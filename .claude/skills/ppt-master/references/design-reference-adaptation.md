# DESIGN.md Reference Adaptation

Use this reference only when the user explicitly names a particular `DESIGN.md`
as visual inspiration. A file that merely happens to be named `DESIGN.md` remains
ordinary source content unless the user gives it this role.

## 1. Output contract

Reduce the reference to:

- **3–5 slide invariants**: the few visual principles that must remain true.
- **0–3 prohibitions**: concrete visual moves the deck must avoid.
- **One style mapping**: the closest existing `visual_style`, or `custom` with a
  neutral `visual_style_behavior` paragraph when no preset fits.

Do not copy a corporate catalog into the project, create a runtime dependency on
an external selector, or invent a company-named visual style.

## 2. Translation lens

Translate only the rules that affect a slide:

| Reference signal | Slide expression |
|---|---|
| Color system | background, primary, accent, secondary accent, body-text roles |
| Typography | title/body contrast, weight span, casing, density |
| Spacing | whitespace posture, grouping distance, dense/breathing rhythm |
| Corners and lines | sharp/soft geometry, radius character, border/stroke behavior |
| Imagery | subject treatment, crop, realism/abstraction, lighting, mood |
| Information density | page load, comparison depth, chart/detail complexity |

Discard interaction-only rules such as hover, navigation, buttons, breakpoints,
responsive rearrangement, touch targets, focus states, and cursor behavior.

## 3. Safety and precedence

- Do not reproduce a specific company's trade dress. Apply transferable visual
  principles, not a brand-identical composition.
- Do not automatically copy logos, proprietary fonts, illustrations, icon sets,
  or distinctive graphic assets. Use an asset exactly only when the user supplied
  it or confirmed the right to use it.
- Safety and rights override every visual request. Installed template or brand
  locks override this reference. An explicit user-fixed value overrides your
  recommendation unless it violates safety or rights.
- When the reference conflicts with a template or brand lock, keep the lock and
  adapt the remaining principles around it. Do not create a second conflict gate.

## 4. Recording the result

Record the source path and the normalized 3–5 invariants / 0–3 prohibitions in
`design_spec.md` as design rationale. Realize them through the existing color,
typography, `visual_style_behavior`, image strategy, spacing, and `page_rhythm`
fields. Do not add new fields to `spec_lock.md`.

For `create-brand`, retain only identity-level tokens and asset permissions; do
not import page layout or web interaction behavior. For `create-template`, the
normalized principles may guide the reusable slide system, but a textual
`DESIGN.md` alone never proves pixel geometry or authorizes high-fidelity copying.
