# Changelog

All notable changes to SlideArabi will be documented in this file.

This project adheres to [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.0.0] — 2026-03-07

### Added

- **`slidearabi/server.py`** — FastAPI application with job lifecycle management, Stripe checkout integration, file upload/download endpoints, and health check.
- **`slidearabi/pipeline.py`** — `SlideArabiPipeline` orchestrator: sequences all seven phases, manages phase I/O, and surfaces structured errors.
- **`slidearabi/rtl_transforms.py`** — Core RTL engine (~3,700 lines). Covers shape mirroring, text frame alignment, table RTL reflow, chart axis reversal, and all OOXML attribute writes required for right-to-left rendering.
- **`slidearabi/llm_translator.py`** — Dual-LLM translation layer. GPT-5.2 performs primary translation; Claude Sonnet 4.6 runs a QA pass to catch mistranslations and formatting artifacts.
- **`slidearabi/visual_qa.py`** — Dual-pass visual QA system. Gemini 3.1 Pro renders each slide and emits structured issue reports; Claude Sonnet 4.6 confirms or escalates each finding. Slides processed in parallel via `ThreadPoolExecutor(max_workers=5)`.
- **`slidearabi/vqa_engine.py`** — XML structural validation engine. Performs read-only checks for shape bounds overflow, orphaned run properties, and malformed OOXML after RTL transformation.
- **`slidearabi/typography.py`** — Arabic font substitution and Unicode directional attribute management. Maps Latin typeface families to their Arabic equivalents and sets `<a:rPr lang>` and `<a:bodyPr>` direction attributes.
- **`slidearabi/property_resolver.py`** — OOXML property inheritance resolver. Walks the theme → layout master → slide master → slide layout → slide cascade to produce fully resolved property sets before transformation.
- **`slidearabi/layout_analyzer.py`** — Slide layout classifier. Categorizes slides by structural template type (title, content, two-column, blank, etc.) to guide phase-specific transformation logic.
- **`slidearabi/template_registry.py`** — Layout pattern registry. Stores canonical RTL layout configurations for each slide template category.
- **`slidearabi/structural_validator.py`** — Post-transform structural validator (Phase 5). Asserts XML well-formedness, verifies shape presence, and confirms text frame integrity without modifying any content.
- **`slidearabi/config.py`** — Centralised configuration: model identifiers, pricing table (USD, GBP, SAR, AED, EGP, EUR), concurrency limits, and directory defaults.
- **Seven-phase deterministic + AI pipeline** — Phases 0, 2, 3, 4, 5 are fully deterministic; Phases 1 and 6 call external LLM APIs.
- **Geo-based pricing** — Per-slide pricing resolved by country at checkout: USD $1 (international), SAR 5, AED 5, EGP 50. GCC countries default to AED.
- **Payment-gated download flow** — Upload triggers a free preview; Stripe checkout gates full processing and download.
- **Docker deployment configuration** — `Dockerfile` and `railway.toml` for Railway backend deployment.
- **Multi-tenant concurrency model** — `Semaphore(1)` per Railway replica; 503 with `Retry-After` header for concurrent requests; horizontal scaling via Railway replica count.
- **3-model architecture review** — Pipeline design validated against Claude Opus 4.6, Codex 5.3, and Gemini 3.1 Pro recommendations for VQA sequencing, concurrency strategy, and payment flow.

### Changed

- Renamed package from `slideshift_v2` to `slidearabi` — unified branding with SemVer. All internal imports, Docker image tags, Railway service name, and environment variable prefixes updated accordingly.
- Railway backend replaces the previous worker-based architecture with the new deterministic pipeline. The Stripe account and keys are shared with the existing production system.

### Architecture

- VQA runs sequentially per slide (Gemini → Claude) but slides are processed in parallel across a `ThreadPoolExecutor(max_workers=5)` pool.
- `python-pptx` OOXML operations are synchronous; the server uses `ThreadPoolExecutor` rather than `asyncio` for pipeline execution to avoid blocking the event loop without introducing async complexity.
- No phase produces side effects on upstream phase output. Each phase consumes the output of the previous phase as an immutable input.
- The frontend (Next.js on Vercel) is unchanged in v1.0.0. Only the Railway backend is replaced.

---

[1.0.0]: https://github.com/mkesani1/slidearabi/releases/tag/v1.0.0
