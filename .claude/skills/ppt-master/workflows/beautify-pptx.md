---
description: Content-faithful PPT beautification — re-layout an existing deck while preserving its text verbatim and inheriting its palette and body size (typography is locked to Pretendard), so regenerated elements blend with it when pasted back.
---

# Beautify PPTX (Re-layout) Workflow

> Mirror of the [`ppt-template-fill` skill](../../ppt-template-fill/SKILL.md): template-fill reuses a deck's design and swaps in new content; beautify keeps a deck's content and redoes its layout.

Re-lays-out an existing `.pptx`: the text is preserved **verbatim**, the source deck's palette and body size are **inherited as truth** (fonts are locked to Pretendard like every SVG-authored deck), and only layout, hierarchy, and whitespace are redesigned. Output is a brand-new native deck generated through the standard SVG pipeline — not a patch over the original.

**Trigger**: the user supplies a `.pptx` and asks to beautify / re-layout / 重新排版 / 美化 while keeping the content. Explicit intent + a provided file only; never auto-infer.

---

## 1. When to Run

| Pattern | Example |
|---|---|
| Existing `.pptx` + beautify intent | "把这份 PPT 美化一下" / "make this deck look better" |
| Existing `.pptx` + re-layout intent | "重新排版这份 PPT，内容别动" / "re-layout this, keep the wording" |
| Existing `.pptx` + paste-back intent | "重排后我要把元素贴回原来的模板" |

**Hard rule — content is frozen**: every text string from the source is preserved exactly (no add / remove / reword / reorder). Beautification freedom lives only in layout, hierarchy, spacing, and visual rhythm.

**Hard rule — not a patch, not a fill**: this regenerates a native deck through Strategist → Executor → export (SKILL.md Steps 4–7). It does **not** edit the source file in place, and it is **not** the [`ppt-template-fill` skill](../../ppt-template-fill/SKILL.md) (which clones source slides and replaces text). It also does not parse an arbitrary third-party template for text-only substitution (the rejected #53 direction) — it builds every page from scratch.

**Distinct from mirror templates**: `replication_mode: mirror` (executor §1.1) keeps layout + visuals verbatim and edits text. Beautify is the inverse — content verbatim, layout redone, identity inherited.

**When this is the wrong route — re-architecture belongs to the main pipeline**: beautify preserves the source's page count and page order 1:1. It is for "keep this deck, just lay it out better". When the user instead wants the original page breakdown reconsidered — merge / split / reorder pages, re-outline the structure, build a *better deck* from the same content rather than a prettier version of the same pages — that is not beautify. This includes re-pagination for fit: "keep every word but split a crowded page so it reads better" changes page count, so it is the main pipeline, not beautify. Run the main SKILL.md pipeline from Step 1 (`source_to_md.py` converts the deck; Step 2 `import-sources` imports it), where the Strategist re-architects the outline freely from the extracted content. The deciding question: is the source's page split information to preserve, or just the previous author's structure to improve? Preserve → beautify (here); improve → main pipeline Step 1.

---

## 2. Inputs

🚧 **GATE**: the user has provided:

| Input | Required | Notes |
|---|---:|---|
| Source PPTX | Yes | The deck to re-lay-out |
| Beautify scope | Optional | Density / emphasis preference — never content rewrites, and never page drops (v1 is strict 1:1) |

---

## 3. Create the Project Workspace

Match the canvas to the source so 1:1 pages and paste-back align. Determine the source aspect first — before the project exists, run `beautify_identity.py <source.pptx>` to **stdout** and read `canvas.aspect` (the formal standard intake bundle is written in Step 4, after `init`) — then `init` with the matching format:

| Source aspect | Format |
|---|---|
| ≈1.778 (16:9) | `ppt169` |
| ≈1.333 (4:3) | `ppt43` |
| other | nearest format in [`canvas-formats.md`](../references/canvas-formats.md); record the source pixel size in the spec |

```bash
python3 ${SKILL_DIR}/scripts/project_manager.py init <title> --format <format>
python3 ${SKILL_DIR}/scripts/project_manager.py import-sources <project_path> <source.pptx> --move
```

`<project_path>` is the `projects/<YYYYMMDD>_<title>` path `init` prints — copy it from the output.

---

## 4. Extract Identity and Data; Assemble Inventory

Use the standard PPTX intake bundle from Step 3. `project_manager.py import-sources` already writes it under `analysis/` for PPTX-family inputs. If the bundle is missing because the project predates this workflow, generate it once:

```bash
python3 ${SKILL_DIR}/scripts/pptx_intake.py <project_path>/sources/<source.pptx> -o <project_path>/analysis
```

**Content + images — already produced by Step 3.** `import-sources` ran `ppt_to_md` on the deck, so the **frozen content contract** is `sources/<stem>.md` (one source slide per block, in order). If the source deck contains pictures, they are already propagated to `images/` with per-slide binding in `images/image_manifest.json` (`occurrences[].slide_index`). Do **not** re-run `ppt_to_md` — it would duplicate the conversion and write images to `analysis/<stem>_files/` instead of `images/`.

**Visual identity (theme + observed sample + canvas)**: read `<project_path>/analysis/<stem>.identity.json` (intake prefixes per-deck artifacts by source-file stem).

| Field | Use |
|---|---|
| `theme.palette.background` / `text` / `primary` / `accent1..6` | the deck's *declared* colors |
| `theme.fonts.title` / `body` (`ea` = CJK, `latin`) | the deck's *declared* fonts — reference only; output typography is the locked Pretendard plan |
| `theme.sizes.title` / `body` (pt) | the deck's *declared* placeholder sizes (master `txStyles`) — the size a run inherits when it sets no explicit `sz`; `body` is the **level-1** default (coarsest, commonly over-reads) |
| `theme.sizes.body_levels` (pt list) | the full master `bodyStyle` ramp (lvl1..lvl9, e.g. `[32, 28, 24, 20, …]`) — **reference context** so you can read a deeper level than the over-reading level-1, not an auto-seed |
| `observed.colors` / `observed.fonts` (`latin` / `ea`, frequency-ranked) | a usage **sample / frequency hint** — run-level fonts + explicit `srgbClr` fills across slides |
| `observed.sizes_pt` (pt, frequency-ranked) | a usage **sample** of run-level explicit point sizes — the **size the deck actually renders at** when it overrides the placeholder default; the source for the Step 5 `body_size` recommendation |
| `layout_sizes_pt` (pt, frequency-ranked) | **reference fact only**, NOT an auto-seed — the level-1 sizes that the in-use slide layouts' body placeholders declare. Usually empty (decks rely on runs / master) and ambiguous when present; use it as a hint when judging the body size, never as the authoritative seed |
| `canvas.aspect` | drives the Step 3 format choice |

> Note: `theme` is what the deck declares; `observed` is a frequency sample of run-level overrides (not a complete style resolution — it misses `schemeClr` and master/layout inheritance, and counts chart/gradient fills). A hand-edited deck can diverge from `theme` — Step 5 recommends which to inherit and the user confirms.

**Chart + table data (for regeneration)**: read `<project_path>/analysis/<stem>.slide_library.json`. It contains the source chart and table *data* so they can be redrawn natively in the inherited style:

| `<stem>.slide_library.json` field | Use |
|---|---|
| `slides[].charts[]` (`chart_type` / `categories` / `series[].values`) | regenerate as a native SVG chart via the `§VII` `templates/charts/` path |
| `slides[].tables[]` (`row_count` / `column_count` / cell text) | regenerate as a native SVG table |

**Hard rule — regenerate visuals, do not carry them over**: charts / tables / images are rebuilt from their data in the inherited style, never spliced in byte-for-byte. This keeps the deck style-consistent and natively editable. **Data values are frozen** (categories / series / cell text / numbers unchanged); only their rendering is the deck's own. Pictures (`ppt_to_md`-extracted files) are reused but re-laid-out — position / crop / size follow the new layout, not the source slot. A user who wants an original element verbatim copies it across themselves.

**Optional source-SVG visual reference**: when the source deck has complex vector decoration, distinctive page chrome, or a visual language that cannot be captured by `<stem>.identity.json` colors/fonts alone, create a read-only SVG reference package under `analysis/`. This is for understanding style only; it is not a carry-over asset path.

```bash
python3 ${SKILL_DIR}/scripts/pptx_to_svg.py <project_path>/sources/<source.pptx> -o <project_path>/analysis/source_svg_import
python3 ${SKILL_DIR}/scripts/extract_svg_assets.py <project_path>/analysis/source_svg_import/svg-flat \
    --icons-dir <project_path>/analysis/source_svg_import/icons \
    --inplace --id-prefix source_flat --min-decoration-bytes 3000 --clean-stale
```

Use the cleaned `analysis/source_svg_import/svg-flat/slide_*.svg` files plus `analysis/source_svg_import/svg-flat_vector_asset_inventory.json` in Step 5/Strategist. Extraction is required for inspection when complex vectors exist: it creates a candidate pool the AI can index, compare, and judge for possible reuse without reading every heavy vector body. Read an individual `analysis/source_svg_import/icons/*.svg` only when the cleaned page and inventory indicate that candidate may be promoted or materially affects the style decision. These candidates are analysis artifacts first, not automatic output assets.

Default: do **not** copy these candidates into the project `icons/`, do **not** list them as reusable output assets, and do **not** preserve original vector decorations byte-for-byte in the beautified deck. The Executor still regenerates fresh native shapes from the confirmed plan.

Optional reuse gate: if a candidate is a non-text brand/logo/motif/decorative asset that should survive the beautification, list it in the Step 5 plan with source slide, candidate filename, intended reuse, and dependency notes from the inventory. Wait for user confirmation. Only confirmed candidates may be promoted into `<project_path>/icons/` and referenced from generated SVGs with `<use data-icon="..."/>`; native export expands those references in memory, while an on-demand `finalize_svg.py` run re-inlines them only for the self-contained SVG preview. Never promote text-bearing groups, charts/tables, source page layouts, or dense slide composites as reusable assets.

**Assemble the inventory** — the deterministic join into one per-slide ledger, `analysis/beautify_inventory.json`, the contract Step 5 confirms and Step 7 verifies against:

```bash
python3 ${SKILL_DIR}/scripts/beautify_inventory.py <project_path>/analysis/<stem>.slide_library.json \
    --images <project_path>/images/image_manifest.json -o <project_path>/analysis/beautify_inventory.json
```

If `images/image_manifest.json` does not exist because the source deck has no extracted pictures, omit `--images`. The script joins per slide: `text_blocks` (slot text + geometry), `tables` (cell grid), `charts` (categories + series values), `diagrams` (SmartArt nodes + hierarchy/connections + source layout), and `images` (bound via `image_manifest` `occurrences[].slide_index`, with geometry / `usage_count`). The **frozen source values are inlined**, so the inventory is a self-contained contract, not a pointer back to `slide_library.json`. It emits `ignored` and `needs_confirmation` as **empty arrays** — fill them with judgment before Step 5:

| Field | Fill with |
|---|---|
| `ignored` | hidden slides / shapes, master-only text, image crop / opacity / rotation / mask (not captured upstream) |
| `needs_confirmation` | unreadable SmartArt data; combo / dual-axis / waterfall charts; merged-cell or multi-header tables; density-outlier pages — **either** overcrowded **or** near-empty / title-only |

**SmartArt output boundary**: Preserve its extracted wording and semantic relationships, then redraw it through SVG as ordinary editable PowerPoint shapes. Do not attempt to regenerate a native SmartArt object or reuse persisted-drawing text as a second content source.

```markdown
## ✅ Extraction Complete

- [x] `sources/<stem>.md` (from Step 3) holds every source slide's text, in order; extracted pictures, if any, are in `images/` + `images/image_manifest.json`
- [x] `analysis/<stem>.identity.json` has theme + observed identity + canvas aspect
- [x] `analysis/<stem>.slide_library.json` holds chart + table data and SmartArt semantic structure for regeneration
- [x] `analysis/source_profile.json` (multi-deck index) summarizes the source facts in its `decks[]` entry
- [x] `analysis/beautify_inventory.json` ledgers per-slide text / images / data + ignored + needs-confirmation
- [ ] **Next**: Step 5 — Beautify Plan (recommend & confirm)
```

---

## 5. Beautify Plan — Recommend & Confirm

⛔ **BLOCKING**: the scope is not hard-coded — same spirit as the Strategist confirmation stage. Recommend each item below from what the deck actually contains (the Step 4 inventory), present the plan, and **wait for the user to confirm or adjust** before writing any spec. Chat is the canonical channel; the confirm UI below is the visual convenience surface over it for the palette + typography review (its result is honored identically to a chat reply).

This step has two halves:
- **Visual re-confirm via the confirm UI** — the **full** Step 4 confirm page (below), seeded from the source so every targeted-confirmation field (canvas, mode, visual style, palette, icons, the locked Pretendard typography plan incl. body size, image strategy, generation mode) is **pre-filled with the inherited / source-derived default and left editable**. Beautify *recommends* keeping the source's identity, but never removes the user's place to override any field — you may choose not to change a value, but you must not deny the place to change it. This is also where the deck's text size is confirmed: `<stem>.identity.json` now carries size hints — `observed.sizes_pt` (the point sizes the deck actually renders at) and `theme.sizes` (the declared placeholder defaults) — so the locked Pretendard plan's `body_size` **follows the source's own font size** rather than a blind canvas default; the user still confirms or overrides it here.
- **Structural scope** — the inventory-driven list decisions below (ignored, reuse, needs-confirmation, verification level) stay in **chat**; they have no confirm-UI widget.

| Plan item | Recommend from | Default lean |
|---|---|---|
| Identity source | `<stem>.identity.json` `theme` vs `observed` | present **both as color candidates in the confirm UI** so the user picks the one that looks right (theme first when the deck is theme-driven; observed first when slides override heavily) — recommend a default ordering and say why. Typography is not a choice here: one locked Pretendard plan carrying the source body size |
| Preserve scope | inventory `text_blocks` / `images` / `charts` / `tables` / `diagrams` | all text verbatim; data values and SmartArt relationships frozen; pictures reused |
| Ignored | inventory `ignored` | name them so the user sees what drops (hidden / master-only text / image crop / rotation) |
| Needs confirmation | inventory `needs_confirmation` | flag complex charts + overcrowded pages explicitly; ask how to handle |
| Verification level | deck size / risk | recommend the Step 7 per-page checks; user sets strictness |

**Hard rule — content is frozen, not the scope decisions**: text strings and chart/table/table-cell data values are non-negotiable (verbatim). *Which* identity to inherit, what to ignore, and how to treat flagged items are recommend-then-confirm, never silently decided.

**Recommend honestly — name the v1 ceiling**:

| Item | What v1 delivers |
|---|---|
| Overcrowded source page | layout / hierarchy / whitespace improve **within the page as-is** — v1 does **not** relieve information overload (that needs re-pagination / rewrite, deferred). Flag such pages; the user may accept or note them for manual split |
| Paste-back into the original | regenerated elements share the inherited palette and body size, so they **blend visually** when pasted (text is set in Pretendard, not the source fonts). v1 does **not** guarantee a seamless coordinate-level drop-in (slide coordinates, master placeholders, font availability are the original deck's, not ours) |
| Complex charts / merged-cell tables | best-effort from the captured data; combo / dual-axis / waterfall lose the un-captured plots — flagged for the user |

**Visual re-confirm — full confirm UI seeded from the source**:

Write `<project_path>/confirm_ui/recommendations.json` and launch the same confirm server SKILL.md Step 4 uses. Do **not** hide fields: seed **every** targeted-confirmation field with the inherited / source-derived default so the user sees the recommendation and keeps the place to change it. Schema → [`scripts/docs/confirm_ui.md`](../scripts/docs/confirm_ui.md).

```json
{
  "lang": "ko",
  "recommend": {
    "canvas": "<step3-canvas-id>",
    "mode": "briefing",
    "visual_style": "<closest visual-style id to the source look>",
    "icons": "<sensible default icon library>",
    "image_usage": ["provided"]
  },
  "page_count": { "value": "<source-slide-count>" },
  "audience": { "value": "<carry over from the deck's apparent audience, or leave blank>" },
  "color": { "selected": 0, "candidates": [
    { "name_ko": "원본 PPT 재현(추천)", "name_zh": "复刻源 PPT（推荐）", "name_en": "Source replica (recommended)", "name_ja": "元PPTを再現（推奨）", "palette": { "background": "#...", "secondary_bg": "#...", "primary": "#...", "accent": "#...", "secondary_accent": "#...", "body_text": "#..." } },
    { "name_ko": "실제 사용 색상(observed)", "name_zh": "实际用色（observed）", "name_en": "Observed palette", "name_ja": "実際の使用色（observed）", "palette": { "background": "#...", "secondary_bg": "#...", "primary": "#...", "accent": "#...", "secondary_accent": "#...", "body_text": "#..." } },
    { "name_ko": "대안 배색 A", "name_zh": "备选配色 A", "name_en": "Alternative palette A", "name_ja": "代替配色A", "palette": { "background": "#...", "secondary_bg": "#...", "primary": "#...", "accent": "#...", "secondary_accent": "#...", "body_text": "#..." } }
  ] },
  "typography": { "selected": 0, "candidates": [
    { "name_ko": "Pretendard 고정(원본 본문 크기 유지)", "name_zh": "Pretendard 锁定（保留源正文字号）", "name_en": "Pretendard lock (source body size kept)", "name_ja": "Pretendard固定（元の本文サイズを維持）", "heading": { "cjk": "Pretendard", "latin": "Pretendard", "css": "Pretendard, \"Malgun Gothic\", sans-serif" }, "body": { "cjk": "Pretendard", "latin": "Pretendard", "css": "Pretendard, \"Malgun Gothic\", sans-serif" }, "body_size": <dominant observed.sizes_pt × 4/3, as px> }
  ] }
}
```

- **Recommend keep, allow override**: pre-fill canvas / mode / visual style / icons / image strategy with the source-faithful default (canvas = Step 3 format, mode = `briefing`, image_usage = `provided` since pictures are reused). Enumerable fields already list every catalog option with the source-faithful one badged, so the user can switch. Beautify's only true non-choices are the frozen text and the strict 1:1 page count (changing those means routing to the main pipeline instead — see CLAUDE.md). The §c material-divergence field is therefore not surfaced here — beautify never reshapes content (text is verbatim).
- **Color — our recommendation is the pre-selected default = the source replica**: author **several color candidates** like the from-scratch flow. The pre-selected default (`selected: 0`, the first card) is what beautify recommends — the palette that **best replicates the source deck's colors** (the truest reading of `theme` / `observed`). Replicate-by-default.
- **Color — judge the other alternatives exactly as the from-scratch flow does**: author each non-replica palette with the **same content-driven judgment the Strategist uses when generating from scratch** (color §e), applied to the material this project provides — the source document's content and subject, the company's own theme colors, and any brand signal. Reach **≥3 color candidates total** so a user who departs from the replica still lands on a considered, content-fitting direction — depart-by-choice.
- **Hard rule — typography is the single locked Pretendard plan**: the typography card carries exactly **one** candidate — `Pretendard, "Malgun Gothic", sans-serif` for heading and body (the install-local font lock, [`strategist.md`](../references/strategist.md) §g) — with the source deck's body **size**. Forbidden — ≥3 font candidates, a "source replica" font set, or inheriting `theme.fonts` / `observed.fonts`. Deviate only when the user explicitly names another font in the current conversation.
- **`body_size` is the load-bearing field, and the locked plan follows the source's own size**: seed the Pretendard plan's `body_size` from the source's actual body size — take the dominant `observed.sizes_pt` value (the most frequent run-level size, the **body proxy**) and **convert it to px (`× 4/3`)** before seeding, since the system is px-only and the source measures in pt: a source 20pt body becomes `26.67`px, so the locked plan renders at the source's true size (seeding the bare `20` as px would shrink it ~25% — the pt-as-px trap). Whichever source value you land on below (observed mode, or `theme.sizes.body`) gets the same `× 4/3` conversion. The confirm page (and chat fallback) then writes that px to `result.json` (`body_size`) **directly — no further conversion, no `body_size_pt` provenance** (pt never enters the contract). The "most frequent = body" read is a proxy, not a guarantee — `observed.sizes_pt` counts every explicit run size (titles, captions, footnotes, chart/label text included, no placeholder-type resolution), so a deck dense with small labels can let a caption size outrank true body; cross-check the proxy against the page's actual body blocks and the sanity range below before trusting it, and prefer the size the body paragraphs visibly render at over the raw mode when the two disagree. Fall back to `theme.sizes.body` (the declared placeholder size) when `observed.sizes_pt` is empty, and to a PPT delivery-purpose baseline (`text` 20 / `balanced` 24 / `presentation` 32 px — one fixed value per purpose) only when neither is present. Note `theme.sizes.body` is the master `bodyStyle` **level-1 declared default** — a coarse value that commonly **over-reads** the real body density (decks often render body at a deeper outline level or override it smaller), so when you land on this fallback treat it as an upper-ish guess and run it through the sanity check below, never as a precise body size. `theme.sizes.body_levels` and `layout_sizes_pt` are **reference context, not extra fallback tiers**: consult them to judge a saner body value when the deck is theme-driven (`observed` empty) — e.g. a deeper `body_levels` entry or a `layout_sizes_pt` hint may read truer than level-1 — but do not auto-seed from them; the seed chain stays `observed → theme.sizes.body → delivery-purpose baseline`, and a theme-driven deck whose body size genuinely can't be pinned cleanly is exactly the case the sanity check is for. The canvas hint stays a **sanity range**, not the seed: if the source's own size lands far outside it (a dense source doc reads tiny on a projection canvas), surface that to the user rather than silently snapping — the recommendation is the source's size, the user confirms or overrides. This is what prevents the deck from exporting at an unintentionally small size while still honoring the source.

```bash
python3 ${SKILL_DIR}/scripts/confirm_ui/server.py <project_path> --daemon --wait
```

Read the confirmed canvas + palette + typography (incl. `body_size`) and any other overrides from `<project_path>/confirm_ui/result.json`. Chat is the canonical fallback when the page cannot open (remote / headless) — present the same fields in chat and honor the reply identically. Always run `--shutdown` on exit (page-confirm or chat-fallback) so port 5050 is free for Step 6 live preview.

**Hard rule — this §5 confirmation replaces the SKILL.md Step 4 confirmation gate**: do not launch a second Step 4 confirm page or chat bundle. On confirmation, hand off to SKILL.md Step 4 **spec writing** as Strategist (gate already satisfied) → `validate_spec.py` → Step 5 (only if `image_usage` includes `ai` / `web`) → Step 6 → Step 7. The two beautify invariants always hold: the content-faithful clause ([`strategist.md`](../references/strategist.md) §d Layer 1) and page count = source slide count (strict 1:1). Everything else comes from the **confirmed** `result.json` — `mode` (recommended `briefing`), canvas, `visual_style`, color (e) + the locked Pretendard typography (g) incl. `body_size` (the reviewed values; skip both recommendation flows) — honoring whatever the user kept or overrode. §VII = chart/table data → `templates/charts/`, §VIII = source pictures for re-layout.

**Hard rule — §IX is verbatim and 1:1**: each source slide becomes exactly one page, in source order, its text transcribed word-for-word from `sources/<stem>.md`. Do not merge, split, drop, or rewrite. Write `design_spec.md` + `spec_lock.md` per `strategist.md` §6 and run `python3 ${SKILL_DIR}/scripts/validate_spec.py <project_path>`, then continue at Step 5 / Step 6.

---

## 6. Executor + Export

Run the standard pipeline (SKILL.md Steps 6–7). The Executor re-lays-out each page — hierarchy, spacing, alignment, page rhythm — using **only** the inherited palette + locked Pretendard typography from `spec_lock.md`, regenerates charts / tables as native SVG from the extracted data, and re-lays-out the source pictures.

```bash
python3 ${SKILL_DIR}/scripts/finalize_svg.py <project_path>  # only when the user asks for svg_final/ previews or svg_final/ already contains SVGs
python3 ${SKILL_DIR}/scripts/svg_to_pptx.py <project_path>
python3 ${SKILL_DIR}/scripts/verify_deck.py <project_path>
```

---

## 7. Validate Output

```bash
python3 ${SKILL_DIR}/scripts/source_to_md/ppt_to_md.py <project_path>/exports/<title>_verN.pptx -o <project_path>/validation/readback.md
```

| Check | Expected |
|---|---|
| Text fidelity | every source text string appears in the output, unaltered |
| Data fidelity | chart categories / series / table cells match the source exactly |
| Page count | output slide count equals the source slide count |
| Regenerated visuals | charts / tables are native SVG re-themed to the inherited palette |
| Identity | generated text / shapes use only the confirmed palette + locked Pretendard typography at the confirmed body size |
| Paste-back | copying a beautified element into the original deck looks native |

```markdown
## ✅ Beautify Complete

- [x] Content + data values verbatim (read-back Markdown matches the source)
- [x] 1:1 page count preserved
- [x] Source colors + body size inherited as locked truth; text set in Pretendard
- [x] Charts / tables regenerated as native SVG in the inherited style
- [x] Native PPTX exported to `exports/`
```

---

## Current Boundary

| Capability | Status |
|---|---|
| Re-layout with verbatim text | Supported |
| Inherit source palette / body size as truth (fonts locked to Pretendard) | Supported |
| Strict 1:1 page mapping | Supported |
| Regenerate charts / tables as native SVG from extracted data | Supported |
| Re-lay-out source pictures | Supported |
| Re-pagination (split dense / merge sparse) | Not in v1 |
| Carry source charts / tables / images over byte-for-byte | Out of scope — user copies originals manually if wanted |
| Free visual-style application / cleanup deviating from source identity | Not in v1 |
| Batch / multi-deck beautification | Not in v1 |
