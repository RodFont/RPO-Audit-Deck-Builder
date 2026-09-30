"""
FastAPI wrapper around audit_to_deck.py.

Exposes a single upload endpoint that accepts an audit spreadsheet
(.xlsx/.xlsm/.csv), runs the existing analysis/deck-building logic
unmodified, and streams back the generated PowerPoint (and optionally
the JSON summary) as a zip.

Run locally:
    pip install -r app/requirements.txt
    uvicorn app.main:app --reload --port 8000

The original CLI script (audit_to_deck.py) is imported as-is from the
project root, so any changes made there are picked up automatically
without touching this file.
"""
import json
import os
import shutil
import sys
import tempfile
import zipfile
from datetime import date
from types import SimpleNamespace

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

# Make the sibling audit_to_deck.py importable regardless of CWD.
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

import audit_to_deck as audit  # noqa: E402

ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".csv"}
MAX_UPLOAD_BYTES = int(os.environ.get("MAX_UPLOAD_BYTES", 25 * 1024 * 1024))  # 25 MB default

app = FastAPI(title="RPO Report Audit Service", version="1.0.0")

# Same-origin uploads from the Akamai-fronted domain; adjust as needed.
app.add_middleware(
    CORSMiddleware,
    allow_origins=os.environ.get("ALLOWED_ORIGINS", "*").split(","),
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)


def _build_args(input_path: str, out_dir: str, title: str, client: str | None) -> SimpleNamespace:
    """Recreate the argparse.Namespace the CLI's main() would build,
    using the same defaults, so analyze()/build_deck() behave identically."""
    ns = {
        "input": input_path,
        "sheet": None,
        "status_map": None,
        "evidence_pattern": r"\b([A-Z][A-Z0-9]*(?:_[A-Z0-9]+)*_(?:EXPORT|REPORT|EVIDENCE))\b",
        "high_below": 0.60,
        "medium_below": 0.80,
        "title": title,
        "client": client,
        "date": None,
        "out_dir": out_dir,
        "json": None,
        "no_deck": False,
    }
    for role in audit.ROLE_ORDER:
        ns[f"{role}_col"] = None
    return SimpleNamespace(**ns)


STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/")
def index():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))


@app.get("/healthz")
def healthz():
    return {"status": "ok"}


@app.post("/api/audit")
async def create_audit(
    file: UploadFile = File(...),
    title: str = "Findings & Recommended Plan",
    client: str | None = None,
):
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(400, f"Unsupported file type '{ext}'. Allowed: {', '.join(sorted(ALLOWED_EXTENSIONS))}")

    workdir = tempfile.mkdtemp(prefix="audit_")
    try:
        input_path = os.path.join(workdir, file.filename)
        size = 0
        with open(input_path, "wb") as f:
            while chunk := await file.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_UPLOAD_BYTES:
                    raise HTTPException(413, "File too large")
                f.write(chunk)

        try:
            args = _build_args(input_path, workdir, title, client)
            analysis = audit.analyze(input_path, args)
        except Exception as exc:  # noqa: BLE001
            raise HTTPException(422, f"Could not parse spreadsheet: {exc}") from exc

        stem = os.path.splitext(os.path.basename(file.filename))[0]
        deck_path = os.path.join(workdir, f"{stem}_Summary.pptx")
        json_path = os.path.join(workdir, f"{stem}_Summary.json")

        entities = analysis["entities"]
        subtitle = client or (", ".join(entities) if 0 < len(entities) <= 2 else stem.replace("_", " "))
        audit.build_deck(analysis, deck_path, title, subtitle, date.today().strftime("%B %Y"))

        slim = {k: v for k, v in analysis.items() if k not in ("rows", "passes")}
        with open(json_path, "w") as f:
            json.dump(slim, f, indent=2, default=str)

        zip_path = os.path.join(workdir, f"{stem}_Summary.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(deck_path, arcname=os.path.basename(deck_path))
            zf.write(json_path, arcname=os.path.basename(json_path))

        return FileResponse(
            zip_path,
            media_type="application/zip",
            filename=os.path.basename(zip_path),
            background=_cleanup_task(workdir),
        )
    except HTTPException:
        shutil.rmtree(workdir, ignore_errors=True)
        raise
    except Exception:
        shutil.rmtree(workdir, ignore_errors=True)
        raise


def _cleanup_task(workdir: str):
    from starlette.background import BackgroundTask

    return BackgroundTask(shutil.rmtree, workdir, ignore_errors=True)


@app.exception_handler(Exception)
async def unhandled_exception_handler(request, exc):  # noqa: ARG001
    return JSONResponse(status_code=500, content={"detail": "Internal server error"})
