---
name: cursor_ppt_skill
description: A professional presentation generator. Build structured, high-impact PowerPoint decks (.pptx) from raw ideas, concepts, or technical specifications. Use when the user asks to brainstorm a presentation, create slides, build a PPT/PPTX, or present a project/change. Enforces a narrative-first workflow with assertion-evidence slides and a density validator to ensure the output is professional and communicative.
---

# cursor_ppt_skill

A professional, narrative-driven presentation tool for Cursor. Transform raw concepts into structured `.pptx` decks by following a rigorous storytelling framework. The core idea is **two-stage generation**: first an `outline.yaml` (the structured prompt), then a deterministic build through `python-pptx` plus optional mermaid rendering.

`outline.yaml` is the source of truth. Edit it to refine your deck and rebuild. Avoid manual `.pptx` structural edits.

## Entry conditions

Use this skill when the user wants to brainstorm a presentation, create a deck, explain a concept, or present any technical/business topic. If the user asks for a simple chart or a text document, this is the wrong tool.

## Required reading before drafting

Before writing any outline:

1. Read [methodology.md](methodology.md) — narrative arc, assertion-evidence, billboard test, slide-type taxonomy, pacing rule.
2. Read [outline-schema.md](outline-schema.md) — the exact YAML schema with all 11 slide types and their fields.

Both are short. Skipping them is the failure mode that produces bad decks.

## The 8 phases

Walk the user through these in order. Do not skip ahead.

```
[ ] Phase 0: Setup check (always run first, silently)
[ ] Phase 1: Discovery
[ ] Phase 2: Template inspection (if template supplied)
[ ] Phase 3: Narrative arc (plain prose, user confirms)
[ ] Phase 4: Draft outline.yaml
[ ] Phase 5: Validate
[ ] Phase 6: Build
[ ] Phase 7: Iterate
```

### Phase 0 — Setup check

Run this silently before doing anything else. Do not ask the user whether to run it — just do it. Tell the user what you are installing only if something is missing.

**Locate the skill directory.** This file (`SKILL.md`) lives inside the skill directory. All script paths below are relative to that directory.

**1. Python packages.** Run the following and install anything that fails:

```bash
python -c "import pptx, pydantic, yaml, PIL, pygments; print('ok')"
```

If that line errors, run:

```bash
pip install python-pptx pydantic pyyaml pillow pygments
```

Tell the user: `Installing required Python packages...` then confirm when done.

**2. Default template.** Check whether `templates/default.pptx` exists next to this file. If it does not:

```bash
python scripts/build_default_template.py
```

Tell the user: `Generating default slide template...` then confirm when done.

**3. Mermaid.** Check whether `mmdc` is available:

```bash
npx --no -- mmdc --version
```

- If it succeeds → mermaid is ready, proceed silently.
- If it errors → **Mermaid is mandatory for this skill.** You MUST guide the user to install it.
  - Run `npm install` from the skill directory.
  - If Node.js is not installed at all, tell the user forcefully: "Node.js is required to render architecture diagrams. Please install it from https://nodejs.org, then we will continue." Do not give them an option to skip.

**NEVER proceed with empty placeholder diagram boxes, and NEVER suggest skipping diagrams.** The power of this tool relies on visual diagrams.

Once all three checks pass (or are explicitly resolved), proceed to Phase 1.

---

### Phase 1 — Discovery (MANDATORY checklist)

Present ALL of the following to the user. Use Cursor's interactive selection features (e.g. numbered options the user can click) when possible. DO NOT proceed to Phase 3 until every item has a user-confirmed answer.

**Required confirmations (MUST ask, MUST get answer):**

1. **Audience** — Who will see this? Suggest options:
   - Technical team (engineers)
   - Mixed audience (some technical, some not)
   - Leadership / executives
   - External / clients
   - [Let user type custom]

2. **Duration** — How long is the talk? Suggest options:
   - 5 minutes (lightning talk, ~3-4 content slides)
   - 15 minutes (standard, ~7-10 content slides)
   - 30 minutes (deep dive, ~15-20 content slides)
   - 60 minutes (workshop, ~25-35 content slides)
   - [Let user type custom]

3. **Goal** — What should the audience do after? (free text, MUST ask)

4. **Template** — Suggest options:
   - Use default template (clean 16:9)
   - I have a company .pptx template → ask for file path

5. **Visual style** — Suggest options:
   - Editorial — clean and professional with accent bars and surface cards
   - Corporate — bold headings, strong color blocks, authoritative
   - Minimal — lots of whitespace, understated elegance
   - Modern — gradients, rounded cards, contemporary feel
   - Vibrant — colorful, energetic, eye-catching
   - Surprise me (let the tool pick based on content)

6. **Color theme** — Suggest options:
   - Light (white background, dark text)
   - Dark (dark background, light text)
   - Ocean (deep blue palette)
   - Warm (amber and earth tones)
   - Forest (green and gold)

7. **Images/assets** — Suggest options:
   - I have images/screenshots to include → ask for folder path, then run `scripts/scan_assets.py`
   - Find suitable images from the web for me
   - No images needed — diagrams and text only

8. **Source material** — What is this deck about? (free text — repo diff, concept, architecture, etc.)

**CRITICAL RULES:**
- Be decisive. Do NOT guess the answers, but forcefully ask the user to provide them if missing.
- Do NOT skip the style/theme selection.
- Do NOT proceed until all 8 items are confirmed.
- If the user says "just make something about X", respond: "I will generate this for you, but I MUST confirm the following parameters first to ensure high quality." Then present items 1–8.
- After all 8 items are confirmed, present a **firm confirmation summary** as a table and ask the user to approve before moving to Phase 2.

### Phase 2 — Template inspection & onboarding

Only when the user supplied a template:

1. **Auto-calculate Safe Areas**:
   ```bash
   python scripts/onboard_template.py path/to/company-template.pptx --auto
   ```
   This script performs geometric collision detection on the template to automatically calculate "Safe Areas" that avoid corporate logos and master slide decorations. It creates a `.meta.yaml` file next to the template. When `build_deck.py` runs, it will read this file and constrain complex slides (Code, Comparison) within these safe areas to perfectly blend with the template.

2. **Inspect layouts** (optional but recommended):
   ```bash
   python scripts/inspect_template.py path/to/company-template.pptx -o template-spec.yaml
   ```

Read `template-spec.yaml`. It tells you:

- Slide dimensions (16:9, 4:3, custom).
- All slide layouts available, by index and name.
- A `suggested_mapping` of slide types -> layouts based on layout-name heuristics.

If the template lacks an expected layout (e.g. no "Comparison"), `build_deck.py` will still pick the closest match. Make a note for the user if a heuristic match looks wrong.

### Phase 3 — Narrative arc

Write the talk **as plain prose**, not slides. Use the five beats:

1. **Problem** — what is broken, missing, or worth talking about
2. **Stakes** — why does anyone care
3. **Solution** — what we propose / built / discovered
4. **Evidence** — data, code, diagrams, before/after
5. **Ask** — what you want from the audience

Show the prose to the user and **explicitly ask for confirmation** before drafting slides. This is the single highest-leverage step. Do not skip it even for short decks.

For repo-change decks, the typical mapping is:

- Problem: the bug, scaling issue, feature gap behind the PR
- Stakes: metrics, incident logs, customer impact
- Solution: the design / the diff (read it from the repo)
- Evidence: benchmarks, before/after diagrams, tests
- Ask: merge approval, rollout signoff, follow-up tasks

### Phase 4 — Draft `outline.yaml`

Copy the structure from one of the `examples/*.outline.yaml` files that matches the talk shape (repo change / concept / architecture review). Fill it in by translating the prose narrative into slides.

Hard rules to apply while drafting:

- Every **content slide** (`bullets`, `comparison`, `code`, `diagram`, `image`, `table`, `statement`) must have a full-sentence `headline`. Not a noun. Not a phrase. A sentence.
- **One idea per slide**. If two bullets do not share a subject, split the slide.
- **Max 5 bullets** per list, **max 15 words** per bullet. Push extra detail to `notes:` instead.
- **Show, don't list.** Don't make every slide `bullets`. Convert flows/architecture to `diagram`, "A vs B" to `comparison`, charts/screenshots to `image`, the single key claim to `statement`. If three consecutive slides are bullets, change one.
- **Code blocks max 20 lines**. Use `highlight_lines:` to focus attention. Code is syntax-highlighted automatically.
- **Diagrams max ~10 nodes**.
- **Speaker notes** (`notes:`) carry depth, citations, transitions.
- **Kicker**: add a short `kicker:` eyebrow to group slides into beats ("The problem", "Evidence", "Ask").
- **Pacing**: 1.5-2 min per content slide for technical talks.

If the user has a template, set `meta.template:` to the template path. Otherwise leave it unset and the built-in `templates/default.pptx` is used. Set `meta.theme:` to the user's chosen color theme from Phase 1 item 6.

**Set `meta.style:` to match the user's choice from Phase 1 item 5.** Valid values: `editorial`, `corporate`, `minimal`, `modern`, `vibrant`. If the user chose "Surprise me", pick based on content type: technical → `editorial` or `minimal`, business → `corporate`, teaching → `modern`, creative → `vibrant`.

Always set `meta.output:` to the desired `.pptx` path (relative to the outline file).

**Image workflow.** When a slide needs an image:

1. If the user has images, have them sit in an `assets/` folder beside the outline. Run:
   ```bash
   python scripts/scan_assets.py assets/
   ```
   It reports each image's dimensions and a suggested `layout`. Reference files with `path: "./assets/<name>"`.
2. If no image exists and a diagram won't do, **find one on the web**, download it into `assets/` (e.g. `curl -o assets/x.png <url>`), and set `source:` to the URL. The source is written into the speaker notes automatically.
3. Never stretch an image — the renderer fits to aspect ratio. Use `layout: right` to put bullets beside a tall image.

The tool does not generate images; it places and attributes them.

### Phase 5 — Validate

```bash
python scripts/validate_outline.py outline.yaml
```

If there are schema errors, fix them. If there are density warnings, fix them too unless there is a concrete reason not to. Use `--strict` if you want warnings to fail the run.

### Phase 6 — Build

**Before running build**, check the outline for `type: diagram` slides:

```bash
python -c "import yaml,sys; d=yaml.safe_load(open(sys.argv[1])); print(sum(1 for s in d['slides'] if s.get('type')=='diagram'),'diagram slides')" outline.yaml
```

- If the count is **0** → build normally:
  ```bash
  python scripts/build_deck.py outline.yaml
  ```
- If the count is **> 0 AND mmdc is not available** → **do not build yet**. Tell the user:
  > "This outline contains diagram slides, but mermaid (mmdc) is not installed. I am running `npm install` to install it now."
  Run the installation. Do not ask for permission to skip it.
- If the count is **> 0 AND mmdc is available** → build normally; diagrams will render.

The output path comes from `meta.output`.

### Phase 7 — Iterate

User reviews the `.pptx`. For any change, **edit `outline.yaml` and rebuild**. Do not hand-edit the `.pptx` for structural changes (last-mile cosmetic tweaks are fine).

Loop: validate -> build -> review -> edit outline -> validate -> build.

## Utility scripts

Execute these; do not paste their contents into chat.

- **`scripts/onboard_template.py <template.pptx> --auto`** — headless geometric safe-area calculation for company templates. Run in Phase 2 to generate `.meta.yaml`.
- **`scripts/inspect_template.py <template.pptx>`** — dump a template's layouts and a suggested slide-type-to-layout mapping. Run in Phase 2.
- **`scripts/scan_assets.py <folder>`** — list images in a folder with dimensions and a suggested layout. Run in Phase 4 before writing image slides.
- **`scripts/validate_outline.py <outline.yaml>`** — schema + density + pacing checks. Run in Phase 5 and before any rebuild.
- **`scripts/build_deck.py <outline.yaml>`** — build the deck. Honours `meta.template` and `meta.output`. Run in Phase 6.
- **`scripts/render_mermaid.py <input.mmd> -o <out.png>`** — standalone mermaid rendering. Usually called internally by `build_deck.py`.
- **`scripts/build_default_template.py`** — regenerate `templates/default.pptx`. Run once after install.
- **`scripts/smoke_test.py`** — round-trip every file in `examples/` through validate -> build. Run before shipping changes to the skill itself.

## Examples to copy from

- [`examples/repo-change.outline.yaml`](examples/repo-change.outline.yaml) — explaining a PR / migration
- [`examples/concept-explainer.outline.yaml`](examples/concept-explainer.outline.yaml) — teaching a concept
- [`examples/architecture-review.outline.yaml`](examples/architecture-review.outline.yaml) — presenting a system design

## Common pitfalls

- **Skipping Phase 3.** You will produce a structurally fine deck that doesn't flow. Always write prose first and confirm with the user.
- **Headline-as-noun.** "Latency" is not a headline. "Cache layer cut p99 latency by 38%" is.
- **Six-bullet slides.** Either delete two bullets or split the slide. Density validator catches it; do not bypass.
- **Pasting whole functions into code slides.** Show the 5-10 lines that matter. The rest goes in `notes:` or the linked source.
- **Editing `.pptx` for structural changes.** You lose regenerability. Edit `outline.yaml` and rebuild.
