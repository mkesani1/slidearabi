# SlideArabi Architecture

This document describes the internal design of the SlideArabi 1.0.0 engine: the pipeline phases, concurrency model, LLM integration, file flow, error handling, and module responsibilities.

---

## Table of Contents

1. [System Overview](#system-overview)
2. [Pipeline DAG](#pipeline-dag)
3. [Phase Descriptions](#phase-descriptions)
4. [Translation Stack](#translation-stack)
5. [Visual QA System](#visual-qa-system)
6. [Deterministic vs AI Phases](#deterministic-vs-ai-phases)
7. [Concurrency Model](#concurrency-model)
8. [File Processing Flow](#file-processing-flow)
9. [Data Flow Through Phases](#data-flow-through-phases)
10. [Error Handling Strategy](#error-handling-strategy)
11. [Multi-Tenant Approach](#multi-tenant-approach)
12. [Cost Model](#cost-model)
13. [Module Reference](#module-reference)

---

## System Overview

```
┌─────────────────────────────────────────────────────────────────────┐
│  Frontend — slidearabi.com                                          │
│  Next.js (Vercel)                                                   │
│  - File upload UI                                                   │
│  - Slide preview carousel                                           │
│  - Stripe checkout redirect                                         │
│  - Download page                                                    │
└──────────────────────────┬──────────────────────────────────────────┘
                           │ HTTPS / REST
                           ▼
┌─────────────────────────────────────────────────────────────────────┐
│  Backend API — api.slidearabi.com                                   │
│  FastAPI + uvicorn (Docker / Railway)                               │
│                                                                     │
│  ┌─────────────────────────────────────────────────────────────┐   │
│  │  Job Manager                                                │   │
│  │  - In-memory job registry (job_id → JobState)              │   │
│  │  - Semaphore(1) per replica for pipeline execution          │   │
│  │  - Background task dispatch via ThreadPoolExecutor          │   │
│  └────────────────────────────┬────────────────────────────────┘   │
│                               │                                     │
│  ┌────────────────────────────▼────────────────────────────────┐   │
│  │  SlideArabiPipeline                                         │   │
│  │  Phase 0 → Phase 1 → Phase 2 → Phase 3 →                  │   │
│  │  Phase 4 → Phase 5 → Phase 6                               │   │
│  └────────────────────────────┬────────────────────────────────┘   │
│                               │                                     │
│  ┌────────────────────────────▼────────────────────────────────┐   │
│  │  Storage                                                    │   │
│  │  uploads/    — raw uploaded .pptx files                     │   │
│  │  outputs/    — converted .pptx files                        │   │
│  └─────────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────────┘
                           │
          ┌────────────────┼────────────────┐
          ▼                ▼                ▼
   OpenAI API       Anthropic API    Google AI API
   (GPT-5.2)     (Claude Sonnet 4.6) (Gemini 3.1 Pro)
```

---

## Pipeline DAG

```
INPUT .pptx
    │
    ▼
┌──────────────────────────────────┐
│  Phase 0: Parse & Resolve        │  property_resolver.py
│  OOXML parse + inheritance walk  │  layout_analyzer.py
└──────────────────┬───────────────┘
                   │  ResolvedPresentation
                   ▼
┌──────────────────────────────────┐
│  Phase 1: Translate              │  llm_translator.py
│  GPT-5.2 primary                 │
│  Claude Sonnet 4.6 QA            │
└──────────────────┬───────────────┘
                   │  TranslatedPresentation
                   ▼
┌──────────────────────────────────┐
│  Phase 2: Master & Layout RTL    │  rtl_transforms.py
│  Mirror masters, layouts, themes │  template_registry.py
└──────────────────┬───────────────┘
                   │  TransformedPresentation (masters/layouts)
                   ▼
┌──────────────────────────────────┐
│  Phase 3: Slide Content RTL      │  rtl_transforms.py
│  Shapes, tables, charts, runs    │
└──────────────────┬───────────────┘
                   │  TransformedPresentation (slides)
                   ▼
┌──────────────────────────────────┐
│  Phase 4: Typography             │  typography.py
│  Font substitution, lang attrs   │
└──────────────────┬───────────────┘
                   │  TypographyNormalizedPresentation
                   ▼
┌──────────────────────────────────┐
│  Phase 5: Structural Validation  │  structural_validator.py
│  Read-only XML integrity checks  │  vqa_engine.py
└──────────────────┬───────────────┘
                   │  ValidationReport (no mutations)
                   ▼
┌──────────────────────────────────┐
│  Phase 6: Visual QA              │  visual_qa.py
│  Gemini 3.1 Pro → Claude 4.6     │
│  Parallelized across slides      │
└──────────────────┬───────────────┘
                   │
                   ▼
            OUTPUT .pptx
```

---

## Phase Descriptions

### Phase 0: Parse & Resolve

**Module:** `property_resolver.py`, `layout_analyzer.py`

The input `.pptx` is opened with `python-pptx`. `property_resolver.py` walks the full OOXML inheritance cascade for every text run in the file:

```
theme defaults → slide master → slide layout → slide → shape → paragraph → run
```

Each run's effective properties (font name, size, bold, color, language) are fully resolved before any transformation begins. This prevents ambiguity in later phases when properties are written back.

`layout_analyzer.py` classifies each slide by template type: `title`, `title_content`, `two_column`, `blank`, `section_header`, `picture_with_caption`, `custom`. The classification guides phase-specific logic in Phases 2 and 3.

### Phase 1: Translate

**Module:** `llm_translator.py`

All text runs extracted in Phase 0 are batched and sent to the translation layer. Translation runs in two passes:

1. **GPT-5.2** receives each text batch with context (slide title, shape type, surrounding runs) and returns Arabic translations.
2. **Claude Sonnet 4.6** receives the original English and the GPT-5.2 output and performs a QA pass, flagging or correcting mistranslations, incorrect proper nouns, and formatting artifacts.

The final translation for each run is the Claude-reviewed output. If Claude finds no issues, the GPT-5.2 output is used as-is.

### Phase 2: Master & Layout Transformation

**Module:** `rtl_transforms.py`, `template_registry.py`

All slide masters and slide layouts are processed before any individual slides. This phase:

- Mirrors shape X positions within the master/layout bounds.
- Sets `rtl="1"` on all `<p:sp>` body properties at the master level.
- Applies the correct `<a:bodyPr>` anchor and alignment for RTL text containers.
- Rewrites tab stop lists in reverse order.
- Consults `template_registry.py` for canonical RTL positions for each layout type.

### Phase 3: Slide Content Transformation

**Module:** `rtl_transforms.py`

The core RTL transformation applied to every slide. Operations include:

- **Shape mirroring:** X coordinate of each shape is recalculated as `slide_width - shape.left - shape.width`.
- **Text frame alignment:** All paragraph alignment attributes flipped (`left` → `right`, `right` → `left`, `center` unchanged).
- **Text direction:** `<a:bodyPr>` direction set to `rtl`; `vert` attribute reviewed per shape type.
- **Table reflow:** Column order reversed; cell alignment flipped; table anchor mirrored.
- **Chart axis reversal:** Category axes flagged for RTL rendering via `<c:plotVisOnly>` and axis crossing overrides.
- **Run-level attributes:** `<a:rPr lang>` set to `ar-SA` for translated runs; `cs` (complex script) font references updated.
- **Bullet lists:** Indent levels and tab stops reversed.

### Phase 4: Typography Normalization

**Module:** `typography.py`

All font references in the transformed file are reviewed:

- Latin typeface families (e.g., `Calibri`, `Arial`, `Times New Roman`) are mapped to their Arabic equivalents (e.g., `Calibri`, `Arial`, `Traditional Arabic`).
- Complex-script (`<a:cs>`) and East-Asian (`<a:ea>`) font slots are updated.
- `<a:latin>` slots retain the original font only where that font has confirmed Arabic glyph coverage.
- The theme font scheme is updated to set Arabic-compatible `majorFont` and `minorFont` complex-script entries.

### Phase 5: Structural Validation

**Module:** `structural_validator.py`, `vqa_engine.py`

A read-only validation pass. No content is modified. Checks include:

- All shapes remain within slide extents (`cx + x ≤ slide_width`, `cy + y ≤ slide_height`).
- No orphaned `<a:r>` (run) elements exist outside `<a:p>` (paragraph) containers.
- XML is well-formed (no unclosed tags, no duplicate attribute keys).
- Each text frame that had content in Phase 0 still contains at least one run.
- No shape GUIDs are duplicated within a slide.

If validation failures are detected, the job transitions to `FAILED` with a structured error report. The output is not written.

### Phase 6: Visual QA

**Module:** `visual_qa.py`

Each slide is rendered to a PNG thumbnail. The dual-pass VQA system processes slides in parallel:

1. **Gemini 3.1 Pro** receives the slide thumbnail and a structured prompt asking it to evaluate: text overflow, misaligned shapes, missing content, incorrect text direction, and visual artifacts.
2. **Claude Sonnet 4.6** receives the same thumbnail plus Gemini's findings and either confirms or escalates each issue.

Issues escalated by both models cause the job to transition to `QA_WARNING` (the file is still returned but flagged). Issues raised only by Gemini and dismissed by Claude are logged but do not affect job status.

---

## Translation Stack

```
English text runs (batched by slide)
        │
        ▼
┌───────────────────────────┐
│  GPT-5.2                  │
│  Primary translation      │
│  Context: slide title,    │
│  shape type, run position │
└───────────────┬───────────┘
                │  Arabic (draft)
                ▼
┌───────────────────────────┐
│  Claude Sonnet 4.6        │
│  QA pass                  │
│  Input: English original  │
│       + GPT-5.2 Arabic    │
│  Output: corrected Arabic │
│        or PASS            │
└───────────────┬───────────┘
                │  Arabic (final)
                ▼
        TranslatedPresentation
```

Translation batches are sized to stay within model context windows. Each batch includes slide context to help the model handle slide titles, bullet points, and labels coherently.

---

## Visual QA System

```
Slide thumbnails (PNG, one per slide)
        │
        ├─── Slide 1 ──┐
        ├─── Slide 2   │  ThreadPoolExecutor(max_workers=5)
        ├─── Slide 3   │  (parallel across slides)
        │   ...        │
        └─── Slide N ──┘
                │
                │  Per slide (sequential within each worker):
                ▼
┌───────────────────────────────┐
│  Gemini 3.1 Pro               │
│  Visual analysis              │
│  Structured issue report      │
└───────────────┬───────────────┘
                │  Issue list (or empty)
                ▼
┌───────────────────────────────┐
│  Claude Sonnet 4.6            │
│  Confirmation / escalation    │
│  Input: thumbnail + issues    │
│  Output: CONFIRMED / DISMISS  │
└───────────────┬───────────────┘
                │
                ▼
        VQA result per slide
        (aggregated → job status)
```

VQA runs after the output `.pptx` is assembled. The file is returned regardless of QA warnings; the `status` field in the job response indicates `COMPLETE` or `QA_WARNING`.

---

## Deterministic vs AI Phases

| Phase | Type | Reason |
|-------|------|--------|
| 0 — Parse & Resolve | Deterministic | Pure OOXML parsing; no ambiguity |
| 1 — Translate | AI | Natural language translation requires LLM judgment |
| 2 — Master & Layout RTL | Deterministic | Geometric transformations with exact formulae |
| 3 — Slide Content RTL | Deterministic | Geometric transformations with exact formulae |
| 4 — Typography | Deterministic | Font mapping table; no inference needed |
| 5 — Structural Validation | Deterministic | Rule-based XML assertion |
| 6 — Visual QA | AI | Visual judgment and issue interpretation require LLM |

Deterministic phases are fast (milliseconds to low seconds for typical decks) and produce identical output for identical input. AI phases introduce latency and API cost but are bounded: Phase 1 scales with slide count and text density; Phase 6 scales with slide count only.

---

## Concurrency Model

```
Incoming request
      │
      ▼
FastAPI endpoint (async)
      │
      ├─ Non-pipeline work (I/O, Stripe, status reads) — handled on event loop
      │
      └─ Pipeline execution
            │
            ▼
      job_semaphore.acquire()   ← Semaphore(1) per Railway replica
            │
       ┌────┴────┐
       │ ACQUIRED │
       └────┬────┘
            │
      ThreadPoolExecutor.submit(run_pipeline, job_id)
            │                      │
            │              pipeline executes synchronously
            │              (python-pptx is not async-safe)
            │
      background task monitors and updates job state
            │
      semaphore.release()
```

**Why `ThreadPoolExecutor` instead of `asyncio`:**
`python-pptx` makes synchronous file I/O and in-memory XML mutations. Running it inside an `asyncio` coroutine would block the event loop. `ThreadPoolExecutor` isolates the synchronous workload to worker threads, keeping the FastAPI event loop free for health checks, status polls, and Stripe webhooks.

**Why `Semaphore(1)` instead of a queue:**
For v1.0.0, each Railway replica processes one deck at a time. Concurrent requests receive a `503 Service Unavailable` response with a `Retry-After: 30` header. Horizontal scaling is achieved by increasing Railway replica count. A persistent job queue (Redis, etc.) is deferred to v1.1.0.

**Phase 6 internal parallelism:**
Within a single pipeline execution, Visual QA uses `ThreadPoolExecutor(max_workers=5)` to process up to 5 slides concurrently. This is nested inside the outer pipeline thread.

---

## File Processing Flow

```
1. POST /convert
   - Validate .pptx MIME type and file size
   - Save to uploads/{job_id}.pptx
   - Create JobState(status=PENDING)
   - Return {job_id}

2. GET /preview/{job_id}   [no payment required]
   - Run Phase 0 only (parse and resolve)
   - Render first 3 slides to PNG thumbnails
   - Return thumbnail URLs

3. POST /create-checkout-session
   - Calculate price: slide_count × per_slide_price[country]
   - Create Stripe checkout session
   - Return {checkout_url}

4. POST /verify-payment  (or Stripe webhook)
   - Verify payment_intent status with Stripe
   - Update JobState(status=PAID)
   - Dispatch pipeline to ThreadPoolExecutor

5. Pipeline runs (background thread):
   - Phase 0–6 execute sequentially
   - JobState.progress updated after each phase (0–100)
   - On success: save outputs/{job_id}.pptx, set status=COMPLETE
   - On failure: set status=FAILED, attach error report

6. GET /status/{job_id}
   - Returns {status, progress, phase, error?}

7. GET /download/{job_id}
   - Verify status=COMPLETE and payment verified
   - Stream outputs/{job_id}.pptx
```

---

## Data Flow Through Phases

```
uploads/{job_id}.pptx
      │
      │  python-pptx Presentation object
      ▼
Phase 0: ResolvedPresentation
  - prs: Presentation
  - resolved_runs: Dict[run_id, ResolvedRunProps]
  - slide_classifications: Dict[slide_idx, LayoutType]
      │
      ▼
Phase 1: TranslatedPresentation
  - prs: Presentation  (text runs mutated in-place with Arabic)
  - translation_map: Dict[run_id, str]
  - qa_corrections: List[QACorrection]
      │
      ▼
Phase 2–3: TransformedPresentation
  - prs: Presentation  (shapes, frames, XML attrs mutated for RTL)
  - transform_log: List[TransformRecord]
      │
      ▼
Phase 4: TypographyNormalizedPresentation
  - prs: Presentation  (font refs updated)
  - font_substitutions: Dict[original_font, arabic_font]
      │
      ▼
Phase 5: ValidationReport
  - passed: bool
  - errors: List[ValidationError]
  - warnings: List[ValidationWarning]
  (prs unchanged — read-only phase)
      │
      ▼
Phase 6: VQAReport
  - slide_results: List[SlideVQAResult]
  - overall_status: PASS | QA_WARNING
  (prs unchanged — read-only phase)
      │
      ▼
prs.save(outputs/{job_id}.pptx)
```

---

## Error Handling Strategy

Each phase raises a typed exception on failure:

| Exception | Phase | Meaning |
|-----------|-------|---------|
| `ParseError` | 0 | File is not a valid .pptx |
| `TranslationError` | 1 | LLM API failure after retries |
| `TransformError` | 2, 3 | Unexpected OOXML structure prevented transform |
| `TypographyError` | 4 | Font substitution mapping failure |
| `ValidationError` | 5 | Post-transform structural integrity failure |
| `VQAError` | 6 | Both VQA models returned API errors |

Exceptions bubble up to `SlideArabiPipeline.run()`, which catches them, sets `JobState.status = FAILED`, attaches the exception message and phase identifier to the job, and releases the semaphore.

**Retry policy:**
- LLM API calls (Phases 1 and 6) retry up to 3 times with exponential backoff on rate-limit or transient errors.
- No retries for validation failures (Phases 5) — these indicate a deterministic transform produced invalid output and must be investigated.
- No fix loops — if a phase fails, the pipeline stops. The client is expected to retry the full job after the issue is resolved.

**No fix loops by design:**
Automatic fix loops (detect issue → attempt fix → re-validate) were considered during the architecture review and rejected. Reasons: unpredictable latency, risk of compounding errors, and difficulty auditing what the system changed. Each phase is run once and produces auditable, immutable output.

---

## Multi-Tenant Approach

For v1.0.0, multi-tenancy is implemented via Railway horizontal replicas:

```
Load Balancer (Railway)
      │
      ├── Replica 1  [Semaphore(1) — processing job A]
      ├── Replica 2  [Semaphore(1) — processing job B]
      └── Replica 3  [Semaphore(1) — idle, accepts next request]
```

Each replica:
- Holds its own in-memory job registry.
- Processes one deck at a time (Semaphore(1)).
- Returns `503 Service Unavailable` with `Retry-After: 30` if busy.

The client (frontend) polls `/status/{job_id}` and retries `503` responses automatically. Because job state is in-memory (not shared across replicas), the client's `job_id` is sticky to the replica that accepted the `/convert` POST — Railway's session affinity ensures subsequent status polls route to the same replica.

**v1.1.0 roadmap:** Replace in-memory job registry with Redis; introduce a persistent job queue (Redis Streams or Celery); remove session affinity requirement.

---

## Cost Model

Approximate API costs per 30-slide deck with moderate text density:

| Component | Model | Estimated cost |
|-----------|-------|---------------|
| Phase 1 — Translation (primary) | GPT-5.2 | ~$0.25 |
| Phase 1 — Translation (QA) | Claude Sonnet 4.6 | ~$0.10 |
| Phase 6 — Visual QA (Gemini pass) | Gemini 3.1 Pro | ~$0.08 |
| Phase 6 — Visual QA (Claude pass) | Claude Sonnet 4.6 | ~$0.07 |
| **Total API cost** | | **~$0.50** |

Revenue at $10/slide for a 30-slide deck (USD): **$300**.

Gross margin per deck (API cost only): **~99.8%**. Infrastructure costs (Railway, Vercel, Stripe fees) reduce this, but the API cost component is negligible relative to revenue.

---

## Module Reference

### `server.py`

FastAPI application entry point. Responsibilities:

- Defines all HTTP endpoints (`/convert`, `/status`, `/preview`, `/download`, `/create-checkout-session`, `/verify-payment`, `/health`).
- Manages the job registry (`Dict[str, JobState]`).
- Handles file upload validation and storage.
- Creates and verifies Stripe checkout sessions.
- Dispatches pipeline execution to the `ThreadPoolExecutor`.
- Enforces the `Semaphore(1)` concurrency limit.

### `pipeline.py`

`SlideArabiPipeline` orchestrator. Responsibilities:

- Instantiates and sequences all phase modules.
- Passes output from each phase as input to the next.
- Updates `JobState.progress` after each phase completes.
- Catches phase exceptions and surfaces structured failure reports.
- Calls `prs.save()` on successful completion.

### `rtl_transforms.py`

Core RTL transformation engine (~3,700 lines). Responsibilities:

- Shape coordinate mirroring for masters, layouts, and slides.
- Text frame `bodyPr` direction, anchor, and alignment writes.
- Paragraph alignment and tab stop reversal.
- Table column order reversal and cell property updates.
- Chart axis RTL configuration.
- Run-level `lang` and `cs` attribute management.

### `llm_translator.py`

Dual-LLM translation layer. Responsibilities:

- Extracts text run batches from the resolved presentation.
- Submits batches to GPT-5.2 with structured prompts.
- Submits GPT-5.2 output to Claude Sonnet 4.6 for QA.
- Writes final Arabic text back into the presentation object.
- Implements retry logic with exponential backoff.

### `visual_qa.py`

Dual-pass VQA system. Responsibilities:

- Renders slide thumbnails via `python-pptx` + Pillow (or LibreOffice headless).
- Submits thumbnails to Gemini 3.1 Pro with structured evaluation prompts.
- Submits thumbnails + Gemini findings to Claude Sonnet 4.6 for confirmation.
- Parallelizes across slides using `ThreadPoolExecutor(max_workers=5)`.
- Aggregates per-slide results into a job-level `VQAReport`.

### `vqa_engine.py`

XML structural validation engine. Responsibilities:

- Parses the OOXML tree post-transformation.
- Checks shape bounds against slide extents.
- Validates run containment, attribute uniqueness, and XML well-formedness.
- Returns structured `ValidationError` and `ValidationWarning` lists.

### `typography.py`

Arabic font handling. Responsibilities:

- Maintains the Latin-to-Arabic font substitution map.
- Updates `<a:latin>`, `<a:cs>`, and `<a:ea>` font references per run.
- Sets `lang="ar-SA"` on all Arabic runs.
- Updates the presentation theme font scheme for Arabic complex-script fonts.

### `property_resolver.py`

OOXML property inheritance resolver. Responsibilities:

- Walks the full OOXML property cascade for each text run.
- Produces `ResolvedRunProps` records with no unresolved inheritance references.
- Handles `<a:defRPr>` (default run properties) at paragraph and list levels.

### `layout_analyzer.py`

Slide layout classifier. Responsibilities:

- Inspects each slide's placeholder types and counts.
- Classifies slides into `LayoutType` enum values.
- Provides classification metadata to Phases 2 and 3 for type-specific transform paths.

### `template_registry.py`

Layout pattern registry. Responsibilities:

- Stores canonical RTL shape positions and sizes for each `LayoutType`.
- Used by Phase 2 to set reference positions for transformed masters and layouts.
- Extensible for custom deck templates.

### `structural_validator.py`

Post-transform structural validator (Phase 5). Responsibilities:

- Asserts shape count consistency vs. Phase 0 baseline.
- Verifies text frame population (no unintended content loss).
- Checks for duplicate shape IDs within slides.
- Runs XML schema-level assertions where applicable.
