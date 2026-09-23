# PPT Brand Converter

Uploads an arbitrary PowerPoint file and returns it rebuilt in the
Learners Point brand template — same visual language, same component
patterns, same judgment calls developed by hand across dozens of real
conversions.

## How it works

```
Upload .pptx
    │
    ▼
extract.py (python-pptx)        deterministic: title, text, tables,
    │                            images, column layout hints
    ▼
planner.js (Claude API)         semantic: reads each source slide,
    │                            picks a brand layout, fills its fields —
    │                            using lib/system-prompt.md, which encodes
    │                            the full component catalog + design rules
    ▼
render.js + components.js       deterministic: dispatches the plan onto
    │                            the real pptxgenjs component library
    ▼
patch_bullets.py                deterministic: XML post-process for
    │                            bullet colors (pptxgenjs has no native
    │                            buClr option)
    ▼
qa.js                           structural scan (rejects any invalid
    │                            zero/negative shape geometry) + Claude
    │                            vision pass on rendered slide images,
    │                            with targeted re-render on flagged slides
    ▼
Final .pptx, fully editable
```

Deterministic code owns every pixel; Claude only ever chooses *which*
already-built, already-tested component fits a slide's content, and
supplies the content for it. That split is what keeps output reliable
at scale — the visual system can't drift, because it's fixed code, not
something the model generates fresh each time.

## Project layout

```
server/
  server.js              Express app: upload, job orchestration, download
  lib/
    components.js        The brand component library (the core asset —
                          every slide pattern proven out in the source
                          conversation, consolidated and parameterized)
    system-prompt.md      Claude's design-decision instructions: the full
                          layout catalog, when to use each, content rules
    planner.js            Batches the extracted deck to Claude, gets back
                          a structured slide plan
    extract.py             Reads a source .pptx into structured JSON
                          (python-pptx)
    render.js              Executes a slide plan against components.js
    patch_bullets.py       Bullet-color XML post-processor
    qa.js                  Structural + Claude-vision QA
  assets/                 Brand images (gradient background, logo,
                          footer decoration) — the actual files used
                          throughout the source conversation
  public/index.html       Upload UI
  Dockerfile              Node 20 + Python3/python-pptx + LibreOffice
                          (LibreOffice powers both the QA render and the
                          structural validator)
```

## Local development

```bash
cd server
npm install
export ANTHROPIC_API_KEY=sk-ant-...
node server.js
# open http://localhost:3000
```

Requires `python3` with `python-pptx` installed, and `soffice`
(LibreOffice) on PATH for the QA/render-to-image step. The Dockerfile
installs both for you in production.

## Extending the brand system

Almost everything about how a conversion turns out lives in two files:

- **`lib/components.js`** — add a new visual pattern here (as a new
  function following the existing ones' style: fixed brand colors/fonts,
  dynamic height-from-character-count sizing so nothing overflows, using
  `safeLine()` for any shape whose dimensions are computed at runtime).
- **`lib/system-prompt.md`** — then add a matching entry to the Layout
  Catalog so Claude knows the new component exists and when to reach for
  it, and add a case for it in `render.js`'s dispatcher.

The brand constants (colors, fonts, canvas size) are exported from
`components.js` — change them once there and every layout picks it up.

## Known limitations (v1)

- Layouts don't place inline images from the source deck (kept out so
  every slide's text stays fully editable, per the "no image slides"
  requirement) — `image_full` is the deliberate exception, used only
  when a source slide is essentially a diagram/photo with no clean
  structured equivalent.
- The planner batches 12 source slides per Claude call; very large decks
  (200+ slides) take several minutes end-to-end, most of it the batched
  planning calls plus the visual QA pass.
- Visual QA re-renders a flagged slide once; it doesn't loop indefinitely
  chasing a fix.
