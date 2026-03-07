# SlideArabi

**AI-powered English-to-Arabic PowerPoint RTL conversion engine.**

![Python 3.12](https://img.shields.io/badge/python-3.12-blue.svg)
![Version 1.0.0](https://img.shields.io/badge/version-1.0.0-green.svg)
![License MIT](https://img.shields.io/badge/license-MIT-lightgrey.svg)

SlideArabi accepts English `.pptx` files, translates all text to Arabic using a dual-LLM pipeline, applies deterministic right-to-left (RTL) layout transformations, and returns a fully RTL-compliant Arabic `.pptx` — ready to open in PowerPoint or Google Slides.

---

## How It Works

The conversion runs as a seven-phase pipeline:

| Phase | Name | Type | Description |
|-------|------|------|-------------|
| 0 | Parse & Resolve | Deterministic | Parses the OOXML file and resolves all inherited text/shape properties via the OOXML cascade (theme → layout → master → slide) |
| 1 | Translate | AI | Translates all extracted text using GPT-5.2 (primary) with Claude Sonnet 4.6 as a QA pass |
| 2 | Master & Layout Transformation | Deterministic | Mirrors and re-aligns all slide masters and layouts for RTL rendering |
| 3 | Slide Content Transformation | Deterministic | Applies RTL transforms to every shape, table, chart, and text frame on every slide |
| 4 | Typography Normalization | Deterministic | Substitutes Latin fonts with appropriate Arabic typefaces; sets correct Unicode directional attributes |
| 5 | Structural Validation | Deterministic | Read-only checks that verify shape bounds, text frame integrity, and XML well-formedness after transformation |
| 6 | Visual QA | AI | Dual-pass visual quality assurance: Gemini 3.1 Pro renders each slide and flags issues; Claude Sonnet 4.6 confirms or escalates |

Each phase produces immutable output. There are no fix loops — phases run exactly once.

See [ARCHITECTURE.md](ARCHITECTURE.md) for the full technical breakdown.

---

## Quick Start (Local Development)

### Prerequisites

- Python 3.12
- Docker (optional, for parity with Railway deployment)
- OpenAI, Anthropic, and Google AI API keys
- Stripe secret key

### 1. Clone the repository

```bash
git clone https://github.com/mkesani1/slidearabi.git
cd slidearabi
```

### 2. Create a virtual environment

```bash
python3.12 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 3. Configure environment variables

```bash
cp .env.example .env
# Edit .env with your keys — see Environment Variables below
```

### 4. Run the development server

```bash
uvicorn slidearabi.server:app --reload --port 8000
```

The API will be available at `http://localhost:8000`.

### 5. Run with Docker

```bash
docker build -t slidearabi .
docker run --env-file .env -p 8000:8000 slidearabi
```

---

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| `POST` | `/convert` | Upload a `.pptx` file to start a conversion job |
| `GET` | `/status/{job_id}` | Poll job progress |
| `GET` | `/preview/{job_id}` | Retrieve slide preview images (free, pre-payment) |
| `GET` | `/download/{job_id}` | Download the converted `.pptx` (requires payment) |
| `POST` | `/create-checkout-session` | Create a Stripe checkout session |
| `POST` | `/verify-payment` | Verify Stripe payment completion |
| `GET` | `/health` | Health check |

Full request/response schemas, error codes, and `curl` examples are in [API.md](API.md).

---

## Architecture Overview

```
Frontend (Vercel / slidearabi.com)
        |
        | HTTPS
        v
Backend API (FastAPI / Railway)
        |
        |-- Job Queue (in-memory, Semaphore-guarded)
        |
        v
SlideArabiPipeline
  Phase 0  →  Phase 1  →  Phase 2  →  Phase 3
                                           |
                                     Phase 4  →  Phase 5  →  Phase 6
        |
        v
Output .pptx (stored, download link returned)
```

For the full system diagram, DAG, concurrency model, and module responsibilities, see [ARCHITECTURE.md](ARCHITECTURE.md).

---

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `OPENAI_API_KEY` | Yes | GPT-5.2 translation |
| `ANTHROPIC_API_KEY` | Yes | Claude Sonnet 4.6 QA and VQA |
| `GOOGLE_AI_API_KEY` | Yes | Gemini 3.1 Pro VQA |
| `STRIPE_SECRET_KEY` | Yes | Stripe payments (shared with production account) |
| `STRIPE_WEBHOOK_SECRET` | Yes | Stripe webhook signature verification |
| `UPLOAD_DIR` | No | Directory for uploaded files (default: `./uploads`) |
| `OUTPUT_DIR` | No | Directory for converted files (default: `./outputs`) |
| `PORT` | No | Server port (default: `8000`) |
| `LOG_LEVEL` | No | Logging level (default: `INFO`) |

Do not commit `.env` to version control. A `.env.example` with placeholder values is provided.

---

## Deployment

See [DEPLOY.md](DEPLOY.md) for full instructions covering:

- Railway (production backend)
- Vercel (frontend)
- Docker build and configuration
- Environment variable management
- Stripe webhook registration

---

## Contributing

1. Fork the repository and create a feature branch from `main`.
2. Follow the existing code style (PEP 8, type hints throughout).
3. Run `python -m pytest` and `python -m py_compile` before opening a PR.
4. Use the [pull request template](.github/PULL_REQUEST_TEMPLATE.md).
5. Do not hardcode model names or pricing in business logic — use the constants defined in `slidearabi/config.py`.

See [ARCHITECTURE.md](ARCHITECTURE.md) for module responsibilities before making structural changes.

---

## License

MIT License. See [LICENSE](LICENSE) for full text.

---

## Repository

GitHub: [github.com/mkesani1/slidearabi](https://github.com/mkesani1/slidearabi)
