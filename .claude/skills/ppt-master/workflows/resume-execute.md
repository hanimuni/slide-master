---
description: Execution-session entry — resume PPT execution in a fresh chat after the planning session (SKILL.md Step 1-5) completed in a previous chat. Reads project state from disk and runs Step 6 + Step 7 with no prior-chat context carry-over.
---

# Resume Execute Workflow

> Standalone execution-session entry. Run when the planning session (SKILL.md Step 1–5) completed in a previous chat and the user wants to continue with SVG generation + export. Loads project state from disk and runs Step 6 + Step 7 in a clean session.

This workflow is **independent**: it owns the execution session starting from a fresh chat — no upstream conversation context required. By isolating SVG generation in its own session, the model gains 20–40K context headroom by not carrying the planning session's Strategist confirmation dialogue, image search/fetch results, or Strategist references.

## When to Run

The user opens a new chat and gives a phrase that names a project path and signals continuation. Recognize any of:

| Pattern | Example |
|---|---|
| "이어서 생성 projects/<YYYYMMDD>_<title>" | "이어서 생성 projects/20260315_joe_hisaishi" |
| "继续生成 projects/<YYYYMMDD>_<title>" | "继续生成 projects/20260315_joe_hisaishi" |
| "resume execution projects/<YYYYMMDD>_<title>" | "resume execution projects/20260315_joe_hisaishi" |
| Project path + any "이어서 / 계속 / 继续 / 恢复 / 接着做" semantic | "projects/20260315_joe_hisaishi 이어서 끝내줘" |

**Prerequisite**: the planning session must have completed in the named project. Verified by file presence in Step 1; do NOT auto-trigger planning on missing state.

---

## Step 1: Sanity check

Verify the project's planning-session artifacts before doing anything else:

| File / Directory | Required when | Reason |
|---|---|---|
| `<project_path>/spec_lock.md` | Always | Strategist's execution contract; Executor re-reads it at P01/P05/P09 milestones and after context compaction |
| `<project_path>/design_spec.md` | Always | Section IX page outline; Executor cross-references it |
| `<project_path>/images/` | `spec_lock images` references any image | Images must exist for embedding |
| `<project_path>/templates/` | `spec_lock page_layouts` references any | Layout SVGs needed for batch read |
| `${SKILL_DIR}/templates/charts/<chart_name>.svg` | `spec_lock page_charts` references any | Chart SVGs are read from the skill library, not the project `templates/` |

If any required artifact is missing → report which one(s) and stop. Do NOT auto-fall-back into planning; the user must either complete the planning session in the original chat or explicitly restart.

---

## Step 2: Load SKILL.md, proceed from Step 6

```
Read ${SKILL_DIR}/SKILL.md
```

Then jump to `### Step 6: Executor Phase` and follow SKILL.md Step 6 verbatim — its reference reading list, gates, and generation order are authoritative; do not substitute a summary from this workflow. Continue into Step 7 (Post-processing & Export) exactly as SKILL.md documents it, including the mandatory final `python3 ${SKILL_DIR}/scripts/verify_deck.py <project_path>` gate.

**Step 7.2 is on demand**: run `finalize_svg.py` only when the user asks for `svg_final/` previews or `svg_final/` already contains SVGs.

The fresh session pays the cost of re-reading references (~14K tokens) but earns back substantially more headroom by dropping the planning session's accumulated context. Net win in both window pressure and reasoning budget per page.

**Source materials**: the execution session is fresh; `<project_path>/sources/<file>.md` is NOT in context. The Executor SHOULD read the relevant `sources/` files when crafting per-page content — they hold the concrete facts, quotes, names, and details that turn skeleton outlines into substantive slides. `design_spec.md §IX` only carries the per-page intent; the source materials carry the texture. The split-mode handoff is designed to free context budget precisely for this kind of high-quality enrichment.

> Note: this workflow does NOT duplicate Step 6 / Step 7 content. SKILL.md is the authoritative procedure; resume-execute only adds the resumption entry (When to Run + Step 1 sanity check above) and the source-materials guidance above.

---

## Step 3: Hand-back

When Step 7 completes and `exports/<title>_ver<N>.pptx` is produced, the workflow ends. Report the export path to the user.

If the deck contains data charts, the [`verify-charts`](verify-charts.md) workflow runs between Step 6 and Step 7 as documented in SKILL.md — resume mode handles it the same way the continuous mode does.
