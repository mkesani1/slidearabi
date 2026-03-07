## Description

<!-- Provide a clear summary of what this PR does and why it is needed. -->
<!-- If this closes a GitHub issue, include: Closes #<issue_number> -->

---

## Type of Change

Check all that apply:

- [ ] Bug fix — non-breaking change that resolves an issue
- [ ] New feature — non-breaking change that adds functionality
- [ ] Breaking change — fix or feature that would cause existing behavior to change
- [ ] Refactor — code restructuring with no functional change
- [ ] Documentation update
- [ ] Dependency update

---

## Affected Components

Check all modules touched by this PR:

- [ ] `server.py` — API endpoints, job management, Stripe
- [ ] `pipeline.py` — Phase orchestration
- [ ] `rtl_transforms.py` — Core RTL engine
- [ ] `llm_translator.py` — Translation stack
- [ ] `visual_qa.py` — Visual QA system
- [ ] `vqa_engine.py` — XML structural checks
- [ ] `typography.py` — Font handling
- [ ] `property_resolver.py` — OOXML inheritance
- [ ] `layout_analyzer.py` — Slide classification
- [ ] `template_registry.py` — Layout patterns
- [ ] `structural_validator.py` — Post-transform validation
- [ ] `config.py` — Model names, pricing, constants
- [ ] Tests
- [ ] Documentation
- [ ] Deployment configuration (Dockerfile, railway.toml)

---

## Checklist

### Code Quality

- [ ] All tests pass locally (`python -m pytest`)
- [ ] No syntax errors (`python -m py_compile slidearabi/*.py`)
- [ ] Type hints are present on all new public functions and methods
- [ ] No bare `except:` clauses — all exceptions are typed

### Model and Pricing

- [ ] No LLM model names are hardcoded in business logic — model identifiers are referenced from `config.py`
- [ ] No pricing values are hardcoded outside `config.py`
- [ ] If models were changed: `ARCHITECTURE.md` and `CHANGELOG.md` are updated to reflect the new model versions

### Pipeline Correctness

- [ ] No fix loops have been introduced — each phase still runs exactly once
- [ ] Phase output remains immutable with respect to prior phases (no upstream mutation)
- [ ] If new OOXML attributes are written: they have been tested against both PowerPoint and Google Slides

### Security

- [ ] No API keys, secrets, or credentials appear anywhere in the diff
- [ ] No new environment variables are required without a corresponding entry in `.env.example` and `README.md`

### Breaking Changes

- [ ] If this is a breaking change: the PR title starts with `[BREAKING]`
- [ ] If the API contract changed: `API.md` is updated
- [ ] If the pipeline phases changed: `ARCHITECTURE.md` is updated
- [ ] `CHANGELOG.md` has an entry for this change under the correct version

---

## Testing Notes

<!-- Describe how you tested this change. Include: -->
<!-- - What .pptx files were used (e.g., slide count, layout types, text density) -->
<!-- - Which phases were exercised -->
<!-- - Any edge cases tested -->

---

## Additional Context

<!-- Screenshots, example outputs, benchmark results, or anything else reviewers should know. -->
