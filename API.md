# SlideArabi API Reference

**Version:** 1.0.0  
**Base URL:** `https://api.slidearabi.com`  

For local development the base URL is `http://localhost:8000`.

---

## Authentication

There is no API key authentication in v1.0.0. Payment authorization is handled entirely through Stripe. Endpoints that require payment (specifically `GET /download/{job_id}`) verify payment status against the internal job record before serving the file.

---

## Table of Contents

1. [POST /convert](#post-convert)
2. [GET /status/{job_id}](#get-statusjob_id)
3. [GET /preview/{job_id}](#get-previewjob_id)
4. [POST /create-checkout-session](#post-create-checkout-session)
5. [POST /verify-payment](#post-verify-payment)
6. [GET /download/{job_id}](#get-downloadjob_id)
7. [GET /health](#get-health)
8. [Error Codes](#error-codes)
9. [Job Status Values](#job-status-values)

---

## POST /convert

Upload a `.pptx` file and create a conversion job.

Payment is not required at upload time. A free slide preview is available after upload via `GET /preview/{job_id}`.

### Request

- **Content-Type:** `multipart/form-data`

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `file` | file | Yes | The `.pptx` file to convert. Max size: 50 MB. |
| `country` | string | No | ISO 3166-1 alpha-2 country code for pricing (e.g., `US`, `GB`, `SA`, `AE`, `EG`). Defaults to `US` if omitted or unrecognized. |

### Response

**Status:** `200 OK`

```json
{
  "job_id": "j_8f3d2a1c",
  "status": "PENDING",
  "slide_count": 24,
  "price_per_slide": 10,
  "currency": "USD",
  "total_price": 240,
  "total_price_display": "$240.00"
}
```

| Field | Type | Description |
|-------|------|-------------|
| `job_id` | string | Unique identifier for this conversion job. |
| `status` | string | Initial job status. Always `PENDING` on creation. |
| `slide_count` | integer | Number of slides detected in the uploaded file. |
| `price_per_slide` | integer | Per-slide price in the smallest currency unit (e.g., cents for USD). |
| `currency` | string | ISO 4217 currency code. |
| `total_price` | integer | Total price in smallest currency unit. |
| `total_price_display` | string | Human-readable formatted price. |

### Errors

| Status | Code | Description |
|--------|------|-------------|
| `400` | `INVALID_FILE_TYPE` | Uploaded file is not a `.pptx` file. |
| `400` | `FILE_TOO_LARGE` | File exceeds the 50 MB limit. |
| `400` | `EMPTY_FILE` | File has zero slides. |
| `503` | `SERVER_BUSY` | All replicas are currently processing jobs. Retry after the indicated delay. |

### curl Example

```bash
curl -X POST https://api.slidearabi.com/convert \
  -F "file=@presentation.pptx" \
  -F "country=US"
```

---

## GET /status/{job_id}

Poll the current status and progress of a conversion job.

### Path Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `job_id` | string | The job ID returned by `POST /convert`. |

### Response

**Status:** `200 OK`

```json
{
  "job_id": "j_8f3d2a1c",
  "status": "PROCESSING",
  "progress": 42,
  "current_phase": 2,
  "current_phase_name": "Master & Layout Transformation",
  "slide_count": 24,
  "paid": true,
  "error": null
}
```

| Field | Type | Description |
|-------|------|-------------|
| `job_id` | string | Job identifier. |
| `status` | string | Current job status. See [Job Status Values](#job-status-values). |
| `progress` | integer | Completion percentage, 0–100. |
| `current_phase` | integer \| null | Index of the currently executing phase (0–6), or `null` if not yet started or complete. |
| `current_phase_name` | string \| null | Human-readable phase name, or `null`. |
| `slide_count` | integer | Slide count for this job. |
| `paid` | boolean | Whether payment has been confirmed. |
| `error` | object \| null | Populated only when `status` is `FAILED`. Contains `phase`, `code`, and `message` fields. |

### Error Object (when status is FAILED)

```json
{
  "phase": 1,
  "phase_name": "Translate",
  "code": "TRANSLATION_ERROR",
  "message": "GPT-5.2 returned a rate limit error after 3 retries."
}
```

### Errors

| Status | Code | Description |
|--------|------|-------------|
| `404` | `JOB_NOT_FOUND` | No job exists with the given `job_id`. |

### curl Example

```bash
curl https://api.slidearabi.com/status/j_8f3d2a1c
```

---

## GET /preview/{job_id}

Retrieve slide preview images for the uploaded deck. No payment is required. Previews are generated from the original English file (pre-conversion).

Previews are available after `POST /convert` returns, regardless of payment status. Typically the first 3 slides are rendered.

### Path Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `job_id` | string | The job ID returned by `POST /convert`. |

### Query Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `slides` | string | `1-3` | Comma-separated slide indices or a range (e.g., `1,2,3` or `1-3`). Maximum of 10 slides per request. |

### Response

**Status:** `200 OK`

```json
{
  "job_id": "j_8f3d2a1c",
  "slide_count": 24,
  "previews": [
    {
      "slide_index": 1,
      "url": "https://api.slidearabi.com/preview-image/j_8f3d2a1c/1.png",
      "width_px": 1280,
      "height_px": 720
    },
    {
      "slide_index": 2,
      "url": "https://api.slidearabi.com/preview-image/j_8f3d2a1c/2.png",
      "width_px": 1280,
      "height_px": 720
    },
    {
      "slide_index": 3,
      "url": "https://api.slidearabi.com/preview-image/j_8f3d2a1c/3.png",
      "width_px": 1280,
      "height_px": 720
    }
  ]
}
```

### Errors

| Status | Code | Description |
|--------|------|-------------|
| `404` | `JOB_NOT_FOUND` | No job with the given ID. |
| `400` | `INVALID_SLIDE_RANGE` | Requested slide indices are out of range or malformed. |

### curl Example

```bash
curl "https://api.slidearabi.com/preview/j_8f3d2a1c?slides=1-3"
```

---

## POST /create-checkout-session

Create a Stripe checkout session for a conversion job. The price is calculated from the slide count and the country code provided at upload.

### Request

- **Content-Type:** `application/json`

```json
{
  "job_id": "j_8f3d2a1c",
  "success_url": "https://slidearabi.com/processing?job_id=j_8f3d2a1c",
  "cancel_url": "https://slidearabi.com/upload"
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `job_id` | string | Yes | The job ID for which to create a checkout. |
| `success_url` | string | Yes | URL to redirect to after successful payment. |
| `cancel_url` | string | Yes | URL to redirect to if the user cancels checkout. |

### Response

**Status:** `200 OK`

```json
{
  "checkout_url": "https://checkout.stripe.com/pay/cs_live_...",
  "session_id": "cs_live_a1b2c3d4",
  "amount": 2400,
  "currency": "usd",
  "amount_display": "$240.00"
}
```

| Field | Type | Description |
|-------|------|-------------|
| `checkout_url` | string | Stripe-hosted checkout URL. Redirect the user here. |
| `session_id` | string | Stripe checkout session ID. Store this to verify payment. |
| `amount` | integer | Total charge in smallest currency unit. |
| `currency` | string | Lowercase ISO 4217 currency code. |
| `amount_display` | string | Human-readable formatted amount. |

### Errors

| Status | Code | Description |
|--------|------|-------------|
| `404` | `JOB_NOT_FOUND` | No job with the given ID. |
| `400` | `ALREADY_PAID` | This job has already been paid. |
| `502` | `STRIPE_ERROR` | Stripe API returned an error. Retry after a short delay. |

### curl Example

```bash
curl -X POST https://api.slidearabi.com/create-checkout-session \
  -H "Content-Type: application/json" \
  -d '{
    "job_id": "j_8f3d2a1c",
    "success_url": "https://slidearabi.com/processing?job_id=j_8f3d2a1c",
    "cancel_url": "https://slidearabi.com/upload"
  }'
```

---

## POST /verify-payment

Verify that a Stripe payment has been completed and trigger pipeline execution. This endpoint is called by the frontend after the user returns from Stripe checkout, as a belt-and-suspenders check alongside the Stripe webhook.

### Request

- **Content-Type:** `application/json`

```json
{
  "job_id": "j_8f3d2a1c",
  "session_id": "cs_live_a1b2c3d4"
}
```

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `job_id` | string | Yes | The job ID to verify payment for. |
| `session_id` | string | Yes | The Stripe checkout session ID from `POST /create-checkout-session`. |

### Response

**Status:** `200 OK`

```json
{
  "job_id": "j_8f3d2a1c",
  "paid": true,
  "status": "PROCESSING",
  "message": "Payment confirmed. Conversion pipeline started."
}
```

| Field | Type | Description |
|-------|------|-------------|
| `job_id` | string | Job identifier. |
| `paid` | boolean | `true` if payment is confirmed. |
| `status` | string | Updated job status after payment confirmation. |
| `message` | string | Human-readable status message. |

### Errors

| Status | Code | Description |
|--------|------|-------------|
| `404` | `JOB_NOT_FOUND` | No job with the given ID. |
| `400` | `PAYMENT_NOT_COMPLETE` | Stripe session exists but payment has not completed. |
| `400` | `SESSION_MISMATCH` | The `session_id` does not match the session created for this job. |
| `400` | `ALREADY_PAID` | Payment was already verified; pipeline is already running or complete. |
| `502` | `STRIPE_ERROR` | Could not verify payment status with Stripe. |

### curl Example

```bash
curl -X POST https://api.slidearabi.com/verify-payment \
  -H "Content-Type: application/json" \
  -d '{
    "job_id": "j_8f3d2a1c",
    "session_id": "cs_live_a1b2c3d4"
  }'
```

---

## GET /download/{job_id}

Download the converted Arabic `.pptx` file. Requires the job to be in `COMPLETE` or `QA_WARNING` status and payment to be confirmed.

### Path Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `job_id` | string | The job ID returned by `POST /convert`. |

### Response

**Status:** `200 OK`

- **Content-Type:** `application/vnd.openxmlformats-officedocument.presentationml.presentation`
- **Content-Disposition:** `attachment; filename="{original_filename}_arabic.pptx"`
- **Body:** Binary `.pptx` file stream.

An additional response header is included:

```
X-QA-Status: PASS
```

or, when visual QA detected warnings:

```
X-QA-Status: QA_WARNING
X-QA-Warning-Count: 3
```

### Errors

| Status | Code | Description |
|--------|------|-------------|
| `404` | `JOB_NOT_FOUND` | No job with the given ID. |
| `402` | `PAYMENT_REQUIRED` | Job exists but payment has not been confirmed. |
| `409` | `JOB_NOT_COMPLETE` | Job is still processing or has failed. Check `/status/{job_id}`. |
| `410` | `FILE_EXPIRED` | Output file has been deleted (jobs expire after 24 hours). |

### curl Example

```bash
curl -OJ https://api.slidearabi.com/download/j_8f3d2a1c
```

---

## GET /health

Health check endpoint. Returns the service operational status and current load. Used by Railway for liveness and readiness probes.

### Response

**Status:** `200 OK` (healthy) or `503 Service Unavailable` (degraded)

```json
{
  "status": "ok",
  "version": "1.0.0",
  "active_jobs": 1,
  "semaphore_available": false,
  "uptime_seconds": 84623
}
```

| Field | Type | Description |
|-------|------|-------------|
| `status` | string | `"ok"` or `"degraded"`. |
| `version` | string | API version string. |
| `active_jobs` | integer | Number of jobs currently in progress on this replica. |
| `semaphore_available` | boolean | Whether this replica can accept a new pipeline job. |
| `uptime_seconds` | integer | Seconds since the server started. |

### curl Example

```bash
curl https://api.slidearabi.com/health
```

---

## Error Codes

All error responses use the following envelope:

```json
{
  "error": {
    "code": "JOB_NOT_FOUND",
    "message": "No job found with ID j_8f3d2a1c.",
    "status": 404
  }
}
```

| HTTP Status | Code | Description |
|-------------|------|-------------|
| 400 | `INVALID_FILE_TYPE` | File is not a `.pptx`. |
| 400 | `FILE_TOO_LARGE` | File exceeds 50 MB. |
| 400 | `EMPTY_FILE` | File contains zero slides. |
| 400 | `INVALID_SLIDE_RANGE` | Requested slide indices are invalid. |
| 400 | `PAYMENT_NOT_COMPLETE` | Stripe payment not yet confirmed. |
| 400 | `SESSION_MISMATCH` | Stripe session ID does not match job. |
| 400 | `ALREADY_PAID` | Job has already been paid. |
| 402 | `PAYMENT_REQUIRED` | Download attempted before payment. |
| 404 | `JOB_NOT_FOUND` | No job exists with the given ID. |
| 409 | `JOB_NOT_COMPLETE` | Download attempted before pipeline completed. |
| 410 | `FILE_EXPIRED` | Output file expired and has been deleted. |
| 503 | `SERVER_BUSY` | Replica is processing another job. Retry after `Retry-After` header value (seconds). |
| 502 | `STRIPE_ERROR` | Stripe API call failed. Retry after a short delay. |

---

## Job Status Values

| Status | Description |
|--------|-------------|
| `PENDING` | Job created, awaiting payment. |
| `PAID` | Payment confirmed; pipeline has been queued. |
| `PROCESSING` | Pipeline is actively running. |
| `COMPLETE` | Pipeline finished successfully. File is ready to download. |
| `QA_WARNING` | Pipeline finished but Visual QA flagged issues. File is still available. |
| `FAILED` | Pipeline encountered an unrecoverable error. See the `error` field in `/status/{job_id}`. |

---

## Pricing Reference

Per-slide prices by country (applied at checkout):

| Country | Currency | Per Slide | Stripe Amount Unit |
|---------|----------|-----------|-------------------|
| United States | USD | $10.00 | 1000 |
| United Kingdom | GBP | £10.00 | 1000 |
| Saudi Arabia | SAR | 40 | 4000 |
| United Arab Emirates | AED | 40 | 4000 |
| Egypt | EGP | 50 | 5000 |
| EU member states | EUR | €10.00 | 1000 |
| All others | USD | $10.00 | 1000 |
