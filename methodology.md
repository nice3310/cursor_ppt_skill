# Methodology: how to make a deck that doesn't suck

This file is read by the agent **before** drafting any `outline.yaml`. It is the "soft skill" layer. The schema and validator enforce some of these rules mechanically; the rest depend on judgement.

The default failure mode of AI-generated decks is "wall of bullets on a white slide." Every rule below exists to push back against that.

---

## 1. Audience first, slides last

Before naming a single slide, the agent must answer four questions out loud and confirm with the user:

- **Who** is the audience? (junior engineers? VPs? mixed?)
- **How long** is the slot? (5 / 15 / 30 / 60 min)
- **What action** should they take after? (approve, decide, learn, align)
- **What do they already know?** (skip anything they know; explain anything they don't)

If the user hasn't told you, **you MUST ask**. Do not guess. Do not infer. Do not proceed. These questions are a hard gate — the agent must have explicit user answers for all of them before writing a single slide. A deck written for the wrong audience is worthless no matter how pretty.

- **Visual style**: which of the 5 built-in styles (`editorial`, `corporate`, `minimal`, `modern`, `vibrant`) does the user want? If unsure, offer "Surprise me" and pick based on content type.

---

## 2. Narrative arc before slide list

Write the story in **plain prose** first. Five beats:

1. **Problem** — what is broken, missing, or worth talking about
2. **Stakes** — why does anyone care; what happens if we ignore it
3. **Solution** — what we propose / built / discovered
4. **Evidence** — data, code, diagrams, before/after
5. **Ask** — what you want from the audience (approval, feedback, action)

Show the user the prose. Get confirmation. **Only then** translate to slides. This single step prevents 80% of "the deck doesn't flow" complaints.

For repo-change decks, the mapping is usually:

| Beat | Repo source |
|------|-------------|
| Problem | the bug / scaling issue / feature gap that motivated the PR |
| Stakes | metrics, incident logs, customer impact |
| Solution | the design / the diff |
| Evidence | benchmarks, before/after diagrams, test results |
| Ask | merge approval, rollout plan signoff, follow-up tasks |

---

## 3. Assertion-Evidence: every slide has a thesis

Every content slide carries a **full-sentence `headline`**, not a topic noun.

- Bad: `title: "Latency"`
- Good: `headline: "Cache layer cut p99 latency by 38%."`

The headline is the slide's *thesis*. The body is the *evidence* (numbers, diagram, bullets, code). If you can't write the thesis as a complete sentence, you don't yet know what the slide is about — go back to step 2.

The schema requires `headline` on `bullets`, `comparison`, `code`, `diagram`, `image`, `table`, and `statement` slide types. `title`, `section`, `summary`, `qa`, and `quote` slides are exempt.

---

## 4. One idea per slide

If a slide has two theses, it must become two slides. Symptoms of a two-idea slide:

- Two bullets that don't share a verb or a subject
- A bullet list that needs a "meanwhile" or "separately"
- A diagram that needs two captions

Split it. Slides are cheap. Confusion is expensive.

---

## 5. Billboard test (5-second rule)

A viewer in row 12 with their phone out should be able to absorb the slide's point in **five seconds**. To pass:

- Headline is one sentence, < 12 words ideally, hard cap 60 chars (validator enforces).
- Bullets: max 5 per slide, max 15 words per bullet (validator enforces).
- Code: max 20 lines, monospace, no inline comments unless essential (validator enforces).
- Diagrams: max 10 nodes/labels. If it needs more, it's an appendix slide.
- Tables: max 5 columns × 7 rows. Anything bigger goes in a doc, not a slide.

If a slide can't pass the billboard test, push the detail to **speaker notes**.

---

## 6. Speaker notes carry the depth

The slide is the **visual aid**. The presenter's mouth is the **narration**. The deck on screen and the spoken talk are *not the same content*.

Use `notes:` for:

- The full argument, not just the headline
- Citations, links, caveats
- Numbers behind the highlighted number
- Transitions to the next slide ("This brings us to...")

If you find yourself wanting to add a sixth bullet, put it in notes instead.

---

## 6b. Show, don't list: prefer visuals to bullets

Bullets are the lowest-information slide type. Before reaching for `bullets`, ask whether the point is better carried by:

- a **diagram** (`type: diagram`) for any flow, architecture, sequence, or relationship
- an **image** (`type: image`) for a chart, screenshot, or before/after
- a **comparison** card for any "A vs B"
- a **statement** for the single most important claim

A deck that is 80% bullets reads as a document. Aim for a mix. If three consecutive slides are all `bullets`, convert at least one to a diagram, image, or comparison.

### Images

The tool cannot generate images. Two ways to get them:

1. The user drops files into an `assets/` folder. Run `scripts/scan_assets.py assets/` to see dimensions and a suggested layout, then reference them with `type: image`.
2. The agent finds a suitable image on the web, downloads it into `assets/`, and records the URL in the slide's `source` field. Always attribute the source — it lands in the speaker notes automatically.

Pick the image `layout` from its shape: wide images → `full`; tall/square images → `right` (with a few bullets beside them).

### Diagrams (mermaid)

Diagrams render through mermaid with the deck's colors. To keep them legible:

- **Max ~10 nodes.** More than that belongs in an appendix or a doc.
- Prefer `flowchart LR` for pipelines, `sequenceDiagram` for request flows, `flowchart TB` for hierarchies.
- Keep node labels to 1-4 words. The headline carries the explanation; the diagram carries the structure.
- One diagram = one relationship. Don't cram a whole system onto one slide.

## 7. Slide-type taxonomy forces structure

Picking a `type` is itself a thinking aid:

- **`title`** — opens the deck
- **`section`** — a divider; reset attention before a new beat
- **`statement`** — one big sentence; use for the central claim of the talk
- **`bullets`** — list of evidence supporting the headline
- **`comparison`** — two columns: before/after, us/them, option-A/option-B
- **`code`** — a code change or API surface
- **`diagram`** — mermaid source rendered to image; architecture, flow, sequence
- **`image`** — screenshot, chart, photo with a caption
- **`table`** — small structured comparison (≤5 cols × 7 rows)
- **`quote`** — customer/PM/exec quote that grounds the problem
- **`summary`** — the 3 takeaways slide near the end
- **`qa`** — final Q&A slide

If you're tempted to invent a new type, you're probably overpacking a slide. Split it.

---

## 8. Pacing rule

Roughly **1.5–2 minutes per content slide** for technical talks. Title, section, summary, and qa slides don't count.

For a 15-minute talk: 7–10 content slides. The validator warns when slide count is wildly off the budget.

---

## 9. Templates: respect them

If the user supplied a company template, **inspect it first** (`scripts/inspect_template.py`). It already encodes:

- Brand colors and fonts (use them; don't override)
- Approved slide layouts with placeholder positions
- Logo placement, page numbers, footer rules

The builder maps each slide `type` to the **closest matching layout name** in the template. If the template lacks (say) a `comparison` layout, the builder falls back to the closest two-content layout. Do not invent extra design — let the template speak.

---

## 10. Source of truth is `outline.yaml`, not the `.pptx`

When the user asks for a change, edit `outline.yaml` and rebuild. Hand-editing the `.pptx` is allowed for last-mile polish (drag a box, fix a typo), but every structural change goes through the outline so it stays regenerable and diffable.

---

## Quick mental checklist before handing back the outline

- [ ] Audience, duration, goal, and ask are explicit in `narrative:`
- [ ] Each content slide has a full-sentence `headline`
- [ ] No slide has more than 5 bullets or more than ~15 words per bullet
- [ ] Speaker notes carry the depth the slide can't hold
- [ ] Slide count matches the pacing rule for the duration
- [ ] Validator runs clean before calling `build_deck.py`
