# outline.yaml schema reference

This is the canonical reference for the structured prompt that the agent produces. The Pydantic models in `scripts/lib/schema.py` are the source of truth; this file documents them in human-readable form.

## Top-level shape

```yaml
meta:        # required
narrative:   # required
slides:      # required, list, length >= 2
```

## `meta`

```yaml
meta:
  title: "Migrating auth to JWT"         # required, str, 1-120 chars
  subtitle: "Backend sync, 2026-Q2"      # optional, str
  presenter: "Y.J. Chen"                 # optional, str
  date: "2026-05-29"                     # optional, ISO date string
  template: "./company-template.pptx"    # optional, path; if absent -> templates/default.pptx
  output: "./deck.pptx"                  # required, path where build_deck.py writes
  style: "editorial"                     # optional: "editorial" (default), "corporate", "minimal", "modern", "vibrant"
  theme: "light"                         # optional: "light", "dark", "ocean", "warm", "forest" (default: "light")
  footer: "Migrating auth to JWT"        # optional footer text; defaults to meta.title
  assets_dir: "./assets"                 # optional; where image files live (see image type)
```

`style` controls the visual design language — layout positions, decorations, bullet markers, card treatments. It is independent of the color `theme`. If omitted, defaults to `"editorial"` for backward compatibility.

## Design fields (shared)

Most content slides accept an optional `kicker` — a short uppercase eyebrow
shown above the title (e.g. "The problem", "Evidence", "Trade-off"). Keep it
under ~3 words. It groups slides into narrative beats and adds visual rhythm.

```yaml
- type: bullets
  kicker: "The problem"     # optional eyebrow above the title
  title: "..."
  headline: "..."
  bullets: ["..."]
```

## `narrative`

```yaml
narrative:
  audience: "backend team, mixed seniority"   # required, str
  duration_minutes: 15                        # required, int, 1-180
  goal: "Get team consensus on JWT migration" # required, str, one sentence
  key_takeaways:                              # required, 1-5 items, each <= 80 chars
    - "Sessions don't scale past 50k users"
    - "JWT cuts auth latency by ~70%"
    - "Migration is a 3-week phased rollout"
```

## `slides`

Every slide is one of 11 `type` values. Common fields:

| Field | Type | Required | Notes |
|-------|------|----------|-------|
| `type` | enum | yes | one of the 11 types below |
| `title` | str | depends | <= 60 chars; required on most types |
| `headline` | str | depends | full-sentence thesis; required on content types |
| `notes` | str | no | speaker notes; unlimited length |

### `type: title`

```yaml
- type: title
  title: "Migrating auth to JWT"
  subtitle: "Why, how, when"
  presenter: "Y.J. Chen"
  date: "2026-05-29"
```

`subtitle`, `presenter`, `date` are optional and fall back to `meta`.

### `type: section`

```yaml
- type: section
  title: "Part 1 — Why sessions hurt"
  number: "01"          # optional; auto-numbered (01, 02, ...) if omitted
```

A dark divider with a large ghosted number. No `headline` required.

### `type: statement`

```yaml
- type: statement
  headline: "Session storage is our biggest scalability bottleneck."
  notes: "Source: 2026-Q1 incident review."
```

One big sentence, centered, large type. No `title`, no body — the headline *is* the slide.

### `type: bullets`

```yaml
- type: bullets
  title: "Why JWT"
  headline: "JWT removes the session table as a single point of failure."
  bullets:
    - "No server-side session lookup on each request"
    - "Horizontal scale without sticky load balancing"
    - "Standard libraries; rotate keys without re-login"
  numbered: false        # optional; true renders 1/2/3 chips instead of dots
  notes: "Cover the trade-off: revocation gets harder."
```

`bullets` is a list, **1–5 items**, each **<= 15 words**. Validator warns over either limit. Each bullet renders as a designed row with an accent marker, not raw text.

### `type: comparison`

```yaml
- type: comparison
  title: "Sessions vs JWT"
  headline: "JWT trades a DB read for a CPU verify."
  columns:
    - label: "Sessions"
      bullets:
        - "DB lookup per request"
        - "Easy revocation"
    - label: "JWT"
      bullets:
        - "Local signature verify"
        - "Revocation via short TTL + denylist"
  notes: "..."
```

`columns` is 2 or 3 entries. Each `bullets` follows the same caps as the `bullets` type.

### `type: code`

```yaml
- type: code
  title: "Token verification middleware"
  headline: "Verification is 8 lines and stateless."
  language: "python"
  code: |
    def verify(token: str) -> Claims:
        try:
            claims = jwt.decode(
                token, PUBLIC_KEY, algorithms=["RS256"], audience="api"
            )
        except jwt.InvalidTokenError as e:
            raise Unauthorized(str(e)) from e
        return Claims(**claims)
  highlight_lines: [3, 4, 5]   # optional, 1-indexed line numbers
  notes: "..."
```

`code` is a literal block. Validator warns if > 20 lines. `language` accepts any Pygments lexer alias.

### `type: diagram`

```yaml
- type: diagram
  title: "Auth flow with JWT"
  headline: "The gateway verifies once; downstream services trust the claim."
  mermaid: |
    sequenceDiagram
      Client->>Gateway: POST /login
      Gateway->>AuthService: verify creds
      AuthService-->>Gateway: signed JWT
      Gateway-->>Client: JWT cookie
      Client->>Gateway: GET /api/x (+JWT)
      Gateway->>ServiceX: forward request + claims
  notes: "..."
```

`mermaid` is rendered to PNG via `mmdc` and inserted. Validator warns if `mermaid` is empty or if the rendered diagram fails.

### `type: image`

```yaml
- type: image
  title: "p99 latency drop"
  headline: "p99 fell from 220 ms to 137 ms after rollout."
  path: "./assets/p99-chart.png"
  caption: "Production, 2026-04-12 to 2026-04-26"
  source: "Internal Grafana dashboard"   # citation/URL -> goes into speaker notes
  layout: "full"                          # auto | full | right | left
  bullets:                                # optional; shows text beside the image
    - "38% reduction at peak"
    - "No regression at p50"
  notes: "..."
```

- `path` is resolved relative to `meta.assets_dir` is not used here; paths are relative to the `outline.yaml` directory (use `./assets/...`).
- Images are fitted **preserving aspect ratio** (never stretched), with a hairline border and a soft shadow.
- `source` is required-by-convention: the validator warns if it is missing. It is appended to the slide's speaker notes as `Image source: ...`. Always attribute web images.
- `layout`: `auto` picks `full` (no bullets) or `right` (with bullets). `right`/`left` place text and image side by side.
- Use `scripts/scan_assets.py <folder>` to list available images with dimensions and a suggested layout before writing image slides.

The tool does **not** generate images. Either the user drops files into the assets folder, or the agent finds an image on the web, downloads it into the assets folder, and records the URL in `source`.

### `type: table`

```yaml
- type: table
  title: "Migration plan"
  headline: "Three weeks, two reversible cutovers."
  columns: ["Week", "Service", "Mode", "Owner"]
  rows:
    - ["1", "auth-svc", "dual-write", "Y.J."]
    - ["2", "gateway", "JWT preferred", "R.L."]
    - ["3", "session DB", "decommission", "Y.J."]
  notes: "..."
```

Validator warns if > 5 columns or > 7 rows.

### `type: quote`

```yaml
- type: quote
  quote: "Auth latency is the top complaint in our NPS survey."
  attribution: "PM, 2026-Q1 retro"
```

No `title`, no `headline`. The quote is the content.

### `type: summary`

```yaml
- type: summary
  title: "Takeaways"
  points:
    - "Sessions hit their scaling wall."
    - "JWT cut p99 latency by 38%."
    - "Approve phased rollout starting next sprint."
  notes: "Pause here. Invite questions."
```

`points` is 1–5 items, each <= 100 chars. Should mirror `narrative.key_takeaways`.

### `type: qa`

```yaml
- type: qa
  title: "Questions?"
  contact: "y.j.chen@example.com"
```

Closing slide. `contact` optional.

---

## Full skeleton

```yaml
meta:
  title: "..."
  template: "./company-template.pptx"
  style: "editorial"
  output: "./deck.pptx"

narrative:
  audience: "..."
  duration_minutes: 15
  goal: "..."
  key_takeaways: ["...", "...", "..."]

slides:
  - type: title
    title: "..."
  - type: section
    title: "Part 1"
  - type: statement
    headline: "..."
  - type: bullets
    title: "..."
    headline: "..."
    bullets: ["...", "..."]
  - type: comparison
    title: "..."
    headline: "..."
    columns:
      - { label: "A", bullets: ["..."] }
      - { label: "B", bullets: ["..."] }
  - type: code
    title: "..."
    headline: "..."
    language: "python"
    code: |
      ...
  - type: diagram
    title: "..."
    headline: "..."
    mermaid: |
      flowchart LR
        A --> B
  - type: image
    title: "..."
    headline: "..."
    path: "./assets/chart.png"
  - type: table
    title: "..."
    headline: "..."
    columns: ["A", "B"]
    rows: [["1", "2"]]
  - type: quote
    quote: "..."
    attribution: "..."
  - type: summary
    title: "Takeaways"
    points: ["...", "...", "..."]
  - type: qa
    title: "Questions?"
```
