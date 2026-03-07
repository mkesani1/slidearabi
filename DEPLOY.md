# SlideArabi 1.0.0 — Deployment Guide

## Architecture

```
┌─────────────────────────┐         ┌──────────────────────────────┐
│   SlideArabi Frontend   │  HTTPS  │   SlideArabi Backend (API)   │
│   (Vercel / Next.js)    │ ◄─────► │   (Railway / Docker)         │
│   slidearabi.com        │         │   api.slidearabi.com         │
└─────────────────────────┘         └──────────────────────────────┘
```

---

## 1. Backend — Railway

### Step-by-step

1. **Create a GitHub repo** named `slidearabi` (private recommended):
   ```bash
   cd slidearabi
   git init
   git add -A
   git commit -m "SlideArabi 1.0.0 — initial deploy"
   git remote add origin git@github.com:mkesani1/slidearabi.git
   git push -u origin main
   ```

2. **Create a Railway project** at [railway.app](https://railway.app):
   - "New Project" → "Deploy from GitHub Repo" → select `slidearabi`
   - Railway auto-detects the Dockerfile

3. **Set environment variables** in Railway → Variables tab:

   | Variable | Value | Source |
   |----------|-------|--------|
   | `OPENAI_API_KEY` | Your OpenAI key | Same key used in current production |
   | `ANTHROPIC_API_KEY` | Your Anthropic key | New — for Claude Sonnet 4.6 QA + VQA |
   | `GEMINI_API_KEY` | Your Gemini key | New — for Gemini 3.1 Pro VQA |
   | `STRIPE_SECRET_KEY` | `sk_live_...` | **Reuse from current Vercel env vars** (see below) |

   Railway auto-sets `PORT` — no need to add it manually.

4. **Deploy** — Railway builds the Docker image and starts the service.

5. **Get your Railway URL** from Settings → Networking:
   - Default: `slidearabi-production-XXXX.up.railway.app`
   - Or add custom subdomain: `api.slidearabi.com`

6. **Verify**:
   ```bash
   curl https://YOUR-RAILWAY-URL/health
   # → {"status":"ok","version":"1.0.0"}
   ```

---

## 2. Stripe — Reusing Your Existing Account

Your current production site (`slidearabi.com` on Vercel, repo `mkesani1/slideshift`) already has Stripe integrated. The v1.0.0 backend uses the **same Stripe account and keys**.

### Where to find your keys

Your Stripe keys are stored as Vercel environment variables for the `slideshift` project:

| Key | Vercel variable name | Where to copy it |
|-----|---------------------|------------------|
| Secret key | `STRIPE_SECRET_KEY` | Railway env vars |
| Publishable key | `NEXT_PUBLIC_STRIPE_PUBLISHABLE_KEY` | Frontend `index.html` |
| Webhook secret | `STRIPE_WEBHOOK_SECRET` | Railway env vars (optional for v1) |

**To retrieve them:**
1. Go to [Vercel Dashboard](https://vercel.com) → your `slideshift` project → Settings → Environment Variables
2. Copy `STRIPE_SECRET_KEY` and `STRIPE_WEBHOOK_SECRET`
3. Add them to Railway environment variables

Alternatively, get them directly from [Stripe Dashboard](https://dashboard.stripe.com/apikeys) → Developers → API Keys.

### Pricing alignment

The v1.0.0 backend geo-pricing matches your production system exactly:

| Region | Currency | Per Slide | Stripe unit amount |
|--------|----------|-----------|-------------------|
| Saudi Arabia | SAR | 5 | 500 |
| UAE | AED | 5 | 500 |
| Egypt | EGP | 50 | 5000 |
| Other GCC (BH, KW, OM, QA) | AED | 5 | 500 |
| International (default) | USD | $1 | 100 |



### Webhook setup (optional for v1)

If you want Stripe to push payment confirmations to the new backend:

1. Stripe Dashboard → Webhooks → Add endpoint
2. URL: `https://YOUR-RAILWAY-URL/stripe-webhook`
3. Events: `checkout.session.completed`
4. Copy the webhook signing secret → add as `STRIPE_WEBHOOK_SECRET` in Railway

For v1.0.0, webhooks are optional — the `/verify-payment` endpoint handles synchronous payment verification.

---

## 3. Frontend — Vercel

The existing `slidearabi.com` frontend (repo `mkesani1/slideshift`) continues to work as-is. The v1.0.0 changes only replace the Railway backend.

### To connect the existing frontend to the new backend

Update `PROCESSING_SERVICE_URL` in your Vercel environment variables to point to the new Railway URL:

```
PROCESSING_SERVICE_URL=https://YOUR-RAILWAY-URL
```

### For the new standalone frontend (tier1-mockup-v2)

If deploying the new static frontend separately:

1. Update `api-integration.js` line 10 with your Railway URL:
   ```javascript
   constructor(baseUrl = 'https://YOUR-RAILWAY-URL') {
   ```

2. Push to a separate GitHub repo (e.g., `slidearabi-web`)
3. Import in Vercel → "Other" framework preset → deploy

### CORS (pre-configured)

The backend allows these origins:
- `https://slidearabi.com`
- `https://www.slidearabi.com`
- `http://localhost:3000`

---

## 4. Quick Test (No Stripe Needed)

The `/convert` endpoint processes immediately without payment:

```bash
# Upload a deck
curl -X POST https://YOUR-RAILWAY-URL/convert \
  -F "file=@presentation.pptx"
# → {"job_id":"abc-123","status":"queued","slide_count":15}

# Poll status
curl https://YOUR-RAILWAY-URL/status/abc-123
# → {"status":"processing","progress_pct":45,"current_phase":"translating"}

# Download when done
curl -OJ https://YOUR-RAILWAY-URL/download/abc-123
```

Payment gating can be enabled later by modifying `_start_pipeline_if_allowed()` in `server.py` to check `job.paid`.

---

## 5. Version Tracking

- **Current release:** 1.0.0
- Version is returned by `/health` and set in `server.py`
- Follow [SemVer](https://semver.org):
  - `1.0.1` — bug fixes, font tweaks
  - `1.1.0` — new features (batch upload, more languages)
  - `2.0.0` — breaking API changes
- All changes logged in [CHANGELOG.md](CHANGELOG.md)

---

## 6. Cost Model

| Component | Est. Cost Per 30-Slide Deck |
|-----------|---------------------------|
| GPT-5.2 translation | ~$0.20 |
| Claude Sonnet 4.6 QA | ~$0.10 |
| Gemini 3.1 Pro VQA | ~$0.15 |
| Railway hosting | ~$5/month |
| **Total API cost** | **~$0.50** |
| **Revenue (30 slides, SA)** | **SAR 150 (~$40)** |
| **Gross margin** | **~98.7%** |
