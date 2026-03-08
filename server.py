import base64
import json
import logging
import os
import shutil
import subprocess
import threading
import time
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Dict, List, Optional

import stripe
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

from slidearabi.llm_translator import DualLLMTranslator, TranslatorConfig
from slidearabi.pipeline import PipelineConfig, SlideArabiPipeline


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
)
logger = logging.getLogger("slidearabi.server")


MAX_FILE_SIZE = 150 * 1024 * 1024  # 150 MB
BASE_DIR = Path("/tmp/slideshift_jobs")
BASE_DIR.mkdir(parents=True, exist_ok=True)
JOB_TTL_HOURS = 24

PIPELINE_SEMAPHORE = threading.Semaphore(1)
JOBS_LOCK = threading.Lock()

PHASE_EXTRACTING = "extracting"
PHASE_TRANSLATING = "translating"
PHASE_RTL = "rtl_transforms"
PHASE_QC = "quality_check"
PHASE_DONE = "done"


@dataclass
class JobState:
    job_id: str
    status: str = "queued"
    progress_pct: int = 0
    current_phase: str = PHASE_EXTRACTING
    input_path: str = ""
    output_path: str = ""
    total_slides: int = 0
    paid: bool = False
    created_at: datetime = field(default_factory=lambda: datetime.now(timezone.utc))
    error: Optional[str] = None
    preview_slides: List[dict] = field(default_factory=list)


JOBS: Dict[str, JobState] = {}


class CheckoutRequest(BaseModel):
    job_id: str
    slide_count: int


class VerifyPaymentRequest(BaseModel):
    session_id: str
    job_id: str


app = FastAPI(title="SlideArabi API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://slidearabi.com",
        "https://www.slidearabi.com",
        "http://localhost:3000",
        "*",  # TODO: remove wildcard once frontend is on slidearabi.com
    ],
    allow_methods=["*"],
    allow_headers=["*"],
)


def _cleanup_expired_jobs() -> None:
    cutoff = datetime.now(timezone.utc) - timedelta(hours=JOB_TTL_HOURS)
    remove_ids = []
    with JOBS_LOCK:
        for job_id, job in JOBS.items():
            if job.created_at < cutoff:
                remove_ids.append(job_id)

        for job_id in remove_ids:
            job = JOBS.pop(job_id, None)
            if job:
                job_dir = BASE_DIR / job_id
                try:
                    if job_dir.exists():
                        shutil.rmtree(job_dir, ignore_errors=True)
                except Exception as exc:
                    logger.warning("Failed to remove expired job dir for %s: %s", job_id, exc)

    if remove_ids:
        logger.info("Cleaned up %d expired jobs", len(remove_ids))


def _job_cleanup_loop() -> None:
    while True:
        try:
            _cleanup_expired_jobs()
        except Exception as exc:
            logger.exception("Cleanup loop error: %s", exc)
        time.sleep(3600)


threading.Thread(target=_job_cleanup_loop, daemon=True).start()


def _count_slides(pptx_path: Path) -> int:
    try:
        import zipfile
        import xml.etree.ElementTree as ET

        with zipfile.ZipFile(pptx_path, "r") as zf:
            with zf.open("ppt/presentation.xml") as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
                sld_id_lst = root.find("p:sldIdLst", ns)
                if sld_id_lst is None:
                    return 0
                return len(sld_id_lst.findall("p:sldId", ns))
    except Exception as exc:
        logger.exception("Failed counting slides for %s: %s", pptx_path, exc)
        raise HTTPException(status_code=400, detail="Unable to read PPTX slide structure")


def _render_preview_slides(input_path: Path, preview_dir: Path, max_slides: int = 3) -> List[dict]:
    preview_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to",
        "jpg",
        "--outdir",
        str(preview_dir),
        str(input_path),
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    generated = sorted(preview_dir.glob("*.jpg"))
    previews = []
    for idx, img_path in enumerate(generated[:max_slides], start=1):
        with open(img_path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("ascii")
        previews.append(
            {
                "slide_num": idx,
                "image_url": f"data:image/jpeg;base64,{encoded}",
            }
        )
    return previews


def _build_translate_fn():
    """Build a translate_fn matching the pipeline's signature: List[str] → Dict[str, str]."""
    openai_key = os.getenv("OPENAI_API_KEY", "")
    anthropic_key = os.getenv("ANTHROPIC_API_KEY", "")

    config = TranslatorConfig(
        openai_api_key=openai_key,
        anthropic_api_key=anthropic_key,
    )
    translator = DualLLMTranslator(config)

    def translate_fn(texts: list) -> dict:
        return translator.translate(texts)

    return translate_fn


def _set_job_state(job_id: str, **kwargs) -> None:
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return
        for k, v in kwargs.items():
            setattr(job, k, v)


def _run_pipeline_worker(job_id: str) -> None:
    acquired = PIPELINE_SEMAPHORE.acquire(blocking=False)
    if not acquired:
        _set_job_state(
            job_id,
            status="failed",
            error="Pipeline busy. Please retry later.",
            progress_pct=0,
        )
        logger.warning("Semaphore unavailable for job %s", job_id)
        return

    try:
        with JOBS_LOCK:
            job = JOBS.get(job_id)
            if not job:
                return

        _set_job_state(job_id, status="processing", current_phase=PHASE_EXTRACTING, progress_pct=5)

        input_path = Path(job.input_path)
        output_path = Path(job.output_path)

        translate_fn = _build_translate_fn()
        cfg = PipelineConfig(
            input_path=str(input_path),
            output_path=str(output_path),
            translate_fn=translate_fn,
        )

        pipeline = SlideArabiPipeline(config=cfg)

        _set_job_state(job_id, current_phase=PHASE_TRANSLATING, progress_pct=25)
        pipeline.run()

        _set_job_state(job_id, current_phase=PHASE_RTL, progress_pct=70)
        _set_job_state(job_id, current_phase=PHASE_QC, progress_pct=90)

        if not output_path.exists():
            raise RuntimeError("Pipeline finished but output file missing")

        _set_job_state(
            job_id,
            status="completed",
            progress_pct=100,
            current_phase=PHASE_DONE,
            error=None,
        )
        logger.info("Job %s completed", job_id)

    except Exception as exc:
        logger.exception("Pipeline failed for job %s: %s", job_id, exc)
        _set_job_state(
            job_id,
            status="failed",
            error=str(exc),
            current_phase=PHASE_DONE,
        )
    finally:
        PIPELINE_SEMAPHORE.release()


def _start_pipeline_if_allowed(job_id: str) -> None:
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Job not found")
        if job.status in {"processing", "completed"}:
            return

    t = threading.Thread(target=_run_pipeline_worker, args=(job_id,), daemon=True)
    t.start()


def _is_pptx(upload: UploadFile) -> bool:
    filename = (upload.filename or "").lower()
    return filename.endswith(".pptx")


@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    try:
        if not _is_pptx(file):
            raise HTTPException(status_code=400, detail="Only .pptx files are allowed")

        content = await file.read()
        if len(content) > MAX_FILE_SIZE:
            raise HTTPException(status_code=413, detail="File too large (max 50MB)")

        if PIPELINE_SEMAPHORE._value <= 0:  # best-effort guard before queueing
            raise HTTPException(status_code=503, detail="Server busy", headers={"Retry-After": "30"})

        job_id = str(uuid.uuid4())
        job_dir = BASE_DIR / job_id
        preview_dir = job_dir / "preview"
        job_dir.mkdir(parents=True, exist_ok=True)

        input_path = job_dir / "input.pptx"
        output_path = job_dir / "output_ar.pptx"
        with open(input_path, "wb") as f:
            f.write(content)

        slide_count = _count_slides(input_path)

        preview_slides = []
        try:
            preview_slides = _render_preview_slides(input_path, preview_dir, max_slides=3)
        except Exception as exc:
            logger.warning("Preview generation failed for %s: %s", job_id, exc)

        job = JobState(
            job_id=job_id,
            status="queued",
            progress_pct=0,
            current_phase=PHASE_EXTRACTING,
            input_path=str(input_path),
            output_path=str(output_path),
            total_slides=slide_count,
            paid=False,
            preview_slides=preview_slides,
        )

        with JOBS_LOCK:
            JOBS[job_id] = job

        # Start immediately for MVP flow (can be gated by payment later via verify-payment)
        _start_pipeline_if_allowed(job_id)

        return {"job_id": job_id, "status": "queued", "slide_count": slide_count}
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("/convert failed: %s", exc)
        raise HTTPException(status_code=500, detail="Internal server error")


@app.get("/status/{job_id}")
def status(job_id: str):
    try:
        with JOBS_LOCK:
            job = JOBS.get(job_id)
            if not job:
                raise HTTPException(status_code=404, detail="Job not found")

            return {
                "status": job.status,
                "progress_pct": max(0, min(100, job.progress_pct)),
                "current_phase": job.current_phase,
                "total_slides": job.total_slides,
            }
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("/status failed for %s: %s", job_id, exc)
        raise HTTPException(status_code=500, detail="Internal server error")


@app.get("/preview/{job_id}")
def preview(job_id: str):
    try:
        with JOBS_LOCK:
            job = JOBS.get(job_id)
            if not job:
                raise HTTPException(status_code=404, detail="Job not found")
            if not job.preview_slides:
                raise HTTPException(status_code=404, detail="Preview not ready")
            return {
                "preview_slides": job.preview_slides,
                "total_slides": job.total_slides,
            }
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("/preview failed for %s: %s", job_id, exc)
        raise HTTPException(status_code=500, detail="Internal server error")


@app.get("/download/{job_id}")
def download(job_id: str):
    try:
        with JOBS_LOCK:
            job = JOBS.get(job_id)
            if not job:
                raise HTTPException(status_code=404, detail="Job not found")
            if job.status != "completed":
                raise HTTPException(status_code=400, detail="Job not completed")
            output_path = Path(job.output_path)

        if not output_path.exists():
            raise HTTPException(status_code=404, detail="Output not found")

        return FileResponse(
            path=str(output_path),
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=f"slideshift_{job_id}.pptx",
        )
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("/download failed for %s: %s", job_id, exc)
        raise HTTPException(status_code=500, detail="Internal server error")


def _country_from_ip() -> str:
    try:
        cmd = [
            "curl",
            "-s",
            "https://ipapi.co/json/",
        ]
        out = subprocess.check_output(cmd, text=True, timeout=10)
        data = json.loads(out)
        return str(data.get("country_code", "")).upper()
    except Exception as exc:
        logger.warning("Geo lookup failed, defaulting to USD: %s", exc)
        return ""


# Geo-pricing — per-slide in smallest currency unit (cents/halalas/fils)
GEO_PRICING = {
    "SA": ("sar", 500),     # SAR 5/slide
    "AE": ("aed", 500),     # AED 5/slide
    "EG": ("egp", 5000),    # EGP 50/slide
    # GCC countries default to AED
    "BH": ("aed", 500),
    "KW": ("aed", 500),
    "OM": ("aed", 500),
    "QA": ("aed", 500),
}


def _price_for_country(country_code: str, slide_count: int):
    cc = (country_code or "").upper()
    if cc in GEO_PRICING:
        return GEO_PRICING[cc]
    return "usd", 100       # Default: $1/slide


@app.post("/create-checkout-session")
def create_checkout_session(payload: CheckoutRequest):
    try:
        with JOBS_LOCK:
            job = JOBS.get(payload.job_id)
            if not job:
                raise HTTPException(status_code=404, detail="Job not found")

        if payload.slide_count <= 0:
            raise HTTPException(status_code=400, detail="slide_count must be positive")

        stripe_key = os.getenv("STRIPE_SECRET_KEY", "")
        if not stripe_key:
            raise HTTPException(status_code=500, detail="Stripe key not configured")

        stripe.api_key = stripe_key

        country = _country_from_ip()
        currency, unit_amount = _price_for_country(country, payload.slide_count)

        session = stripe.checkout.Session.create(
            mode="payment",
            payment_method_types=["card"],
            line_items=[
                {
                    "quantity": payload.slide_count,
                    "price_data": {
                        "currency": currency,
                        "unit_amount": unit_amount,
                        "product_data": {
                            "name": "SlideArabi PPT Translation",
                            "description": f"{payload.slide_count} slides",
                        },
                    },
                }
            ],
            success_url="https://slidearabi.com/success?session_id={CHECKOUT_SESSION_ID}",
            cancel_url="https://slidearabi.com/cancel",
            metadata={"job_id": payload.job_id},
        )

        return {"checkout_url": session.url, "session_id": session.id}
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("/create-checkout-session failed: %s", exc)
        raise HTTPException(status_code=500, detail="Internal server error")


@app.post("/verify-payment")
def verify_payment(payload: VerifyPaymentRequest):
    try:
        stripe_key = os.getenv("STRIPE_SECRET_KEY", "")
        if not stripe_key:
            raise HTTPException(status_code=500, detail="Stripe key not configured")

        stripe.api_key = stripe_key

        with JOBS_LOCK:
            job = JOBS.get(payload.job_id)
            if not job:
                raise HTTPException(status_code=404, detail="Job not found")

        session = stripe.checkout.Session.retrieve(payload.session_id)
        paid = session.payment_status == "paid"

        if paid:
            _set_job_state(payload.job_id, paid=True)
            _start_pipeline_if_allowed(payload.job_id)
            return {"verified": True, "status": "processing"}

        return {"verified": False, "status": "unpaid"}
    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("/verify-payment failed: %s", exc)
        raise HTTPException(status_code=500, detail="Internal server error")


@app.get("/health")
def health():
    try:
        return {"status": "ok", "version": "1.0.0"}
    except Exception as exc:
        logger.exception("/health failed: %s", exc)
        raise HTTPException(status_code=500, detail="Internal server error")
