# ppt-builder — Cursor skill

**AI-driven Presentation-as-Code.** Transform raw ideas, system architectures, or code changes into structured, professional `.pptx` decks.

Say goodbye to wall-of-text white slides. This skill enforces a narrative-first workflow and "Assertion-Evidence" slide design to create presentations that actually communicate.

---

## Installation (3 steps)

### Step 1: Copy to Cursor personal skills folder

Cursor reads personal skills from this path. Put the entire `ppt-builder/` folder there:

**Windows**
```powershell
Copy-Item -Recurse "ppt-builder" "$env:USERPROFILE\.cursor\skills\ppt-builder"
```

**macOS / Linux**
```bash
cp -r ppt-builder ~/.cursor/skills/ppt-builder
```

After restarting Cursor, typing "help me make a ppt" in chat will automatically trigger this skill.

### Step 2: Install Python packages

Requires Python 3.9+.

```powershell
pip install python-pptx pydantic pyyaml pillow pygments
```

### Step 3: (Optional) Install mermaid support

Only needed if you want to draw architecture/flow diagrams. Requires Node.js.

```powershell
cd ~/.cursor/skills/ppt-builder   # For Windows: $env:USERPROFILE\.cursor\skills\ppt-builder
npm install
```

If not installed, `type: diagram` slides will show placeholder text instead of breaking.

---

## Usage

### 1. Talk to the Cursor Agent (Recommended)

In any project's Cursor chat, simply start a conversation. The Agent will guide you through discovery, narrative drafting, and final generation. Try:

- "I want to explain the concept of [Topic] to [Audience], help me brainstorm a 10-slide deck."
- "Create a pitch deck for a new feature that does [X]."
- "Turn this technical design doc into a presentation for the executive team."
- "Use `./company-template.pptx` to make a presentation about our Q3 goals."

### 2. Manual Workflow (for Power Users)

If you prefer to define your slides directly, you can use the `outline.yaml` format.

1. **Pick a starting point**: Choose a file from the `examples/` folder.
2. **Modify**: Edit the YAML to fit your content.
3. **Build**:
   ```powershell
   # Validate the structure and content density
   python scripts/validate_outline.py my-deck.outline.yaml

   # Generate the PowerPoint file
   python scripts/build_deck.py my-deck.outline.yaml
   ```

---

## 📂 Examples

The `examples/` directory contains pre-configured outlines for common scenarios. You can use these as templates:

- `concept-explainer.outline.yaml`: Best for teaching new ideas or primers.
- `architecture-review.outline.yaml`: Best for system designs or structural proposals.
- `repo-change.outline.yaml`: Best for technical walkthroughs or PR reviews.

---

## What outline.yaml looks like

Core concept: Each slide must have a `headline` (a full assertion sentence, not a noun title). The schema prevents overpacking.

```yaml
meta:
  title: "JWT Migration"
  output: "./deck.pptx"
  # template: "./company-template.pptx"  # Uncomment if using a company template

narrative:
  audience: "backend team"
  duration_minutes: 15
  goal: "Get consensus on phased rollout"
  key_takeaways:
    - "Sessions don't scale past 50k users"
    - "JWT cuts auth latency by 38%"

slides:
  - type: title
    title: "JWT Migration"

  - type: statement
    headline: "Session storage is our biggest scalability bottleneck."

  - type: bullets
    title: "Why JWT"
    headline: "JWT removes the DB read on every request."   # Full assertion sentence, required
    bullets:
      - "No session lookup per request"
      - "Horizontal scale without sticky sessions"
      - "Standard libraries, easy key rotation"

  - type: code
    title: "Verification middleware"
    headline: "8 lines, stateless."
    language: "python"
    highlight_lines: [3, 4]
    code: |
      def verify(token):
          try:
              claims = jwt.decode(token, PUBLIC_KEY, algorithms=["RS256"])
          except jwt.InvalidTokenError as e:
              raise Unauthorized(str(e))
          return Claims(**claims)

  - type: summary
    title: "Takeaways"
    points:
      - "Sessions hit their scaling wall."
      - "JWT cuts p99 by 38%."
      - "Approve 3-week phased rollout."

  - type: qa
    title: "Questions?"
```

See `outline-schema.md` for full specifications of all 11 slide types.

---

## How to use a company template

```powershell
# Inspect what layouts the template has first
python "$env:USERPROFILE\.cursor\skills\ppt-builder\scripts\inspect_template.py" company.pptx

# Set template path in outline.yaml
# meta:
#   template: "./company.pptx"

# Build (will inherit master slide colors/fonts/logos)
python "$env:USERPROFILE\.cursor\skills\ppt-builder\scripts\build_deck.py" my-talk.outline.yaml
```

If no template is set, it defaults to `templates/default.pptx` (16:9 clean white background).

---

## How to read Validator output

```
3 warning(s):
  - slides[4] (bullets): 7 bullets (>5); split the slide
  - slides[7].bullets[2]: 22 words (>15); tighten
  - pacing: 12 content slides for 15 min = 1.2 min/slide (<1.5); too rushed
```

Each line is an action instruction: if there are more than 5 bullets, split the slide; if a bullet is too long, shorten the detail and move it to `notes:`; if pacing is too tight, remove slides or ask for more time.

---

## Directory Structure

```
ppt-builder/
├── SKILL.md                # Cursor skill entry point
├── methodology.md          # Narrative arc / billboard test / speaker notes philosophy
├── outline-schema.md       # Full YAML spec for 11 slide types
├── pyproject.toml
├── package.json
├── examples/
│   ├── repo-change.outline.yaml
│   ├── concept-explainer.outline.yaml
│   └── architecture-review.outline.yaml
├── templates/
│   └── default.pptx        # Fallback when no company template is provided
└── scripts/
    ├── build_deck.py         # outline.yaml → .pptx
    ├── validate_outline.py   # schema + density + pacing validation
    ├── inspect_template.py   # Analyze company .pptx template
    ├── render_mermaid.py     # mermaid → PNG (requires Node)
    ├── build_default_template.py
    ├── smoke_test.py
    └── lib/
        ├── schema.py   # Pydantic models
        ├── layouts.py  # slide type → layout mapping
        ├── renderers.py
        └── theme.py
```
.py   # Analyze company .pptx template
    ├── render_mermaid.py     # mermaid → PNG (requires Node)
    ├── build_default_template.py
    ├── smoke_test.py
    └── lib/
        ├── schema.py   # Pydantic models
        ├── layouts.py  # slide type → layout mapping
        ├── renderers.py
        └── theme.py
```
