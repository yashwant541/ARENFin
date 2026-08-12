"""
backend.py — Python backend for the Arabic→English financial converter webapp.

Works in two places with the same code:

  • Dataiku DSS standard web app — paste this into the backend "Python" tab.
    DSS already injects a Flask `app`; this file detects it and only registers
    routes. Put the `afx/` package (and its financial_dictionary.csv) in the
    project's Libraries → python so `import afx` resolves, OR set env AFX_HOME to
    the folder that contains `afx/`, and optionally AFX_DICT to the glossary file.

  • Standalone — `python standalone_app.py` runs the exact same routes plus the
    front-end for local testing (see standalone_app.py).

Endpoints (all under the backend URL DSS gives the JS via getWebAppBackendUrl):

  GET  /api/health                    -> glossary stats + service status
  POST /api/convert   (multipart)     -> convert every uploaded file, one job
  GET  /api/download/<job>/<fid>      -> one converted .xlsx (attachment)
  GET  /api/download_all/<job>        -> all outputs of a job as a .zip

Each uploaded file is converted independently: one file failing (e.g. a legacy
.xls with no xlrd) never blocks the others; its row just reports the error.
"""
from __future__ import annotations
import io
import json
import os
import pathlib
import re
import sys
import tempfile
import traceback
import uuid
import zipfile


# --------------------------------------------------------------------------
# 1. Make the afx package importable (DSS lib, AFX_HOME, or a repo checkout)
# --------------------------------------------------------------------------
def _bootstrap_afx() -> None:
    try:
        import afx  # noqa: F401  (already on the path — DSS project lib, etc.)
        return
    except Exception:
        pass
    here = pathlib.Path(__file__).resolve()
    candidates = [os.environ.get("AFX_HOME"), *[str(p) for p in here.parents]]
    for c in candidates:
        if c and (pathlib.Path(c) / "afx" / "__init__.py").exists():
            sys.path.insert(0, c)
            return


_bootstrap_afx()
from afx import FinancialTranslator, ExcelConverter  # noqa: E402


def _find_dictionary() -> str:
    env = os.environ.get("AFX_DICT")
    if env and os.path.exists(env):
        return env
    import afx as _afx
    base = pathlib.Path(_afx.__file__).parent
    for name in ("financial_dictionary.csv", "financial_dictionary.json"):
        p = base / name
        if p.exists():
            return str(p)
    raise FileNotFoundError(
        "Glossary not found. Put financial_dictionary.csv next to the afx package "
        "or set the AFX_DICT environment variable."
    )


# --------------------------------------------------------------------------
# 2. Load the engine once (cheap; ~1500 variants) and prepare a job store
# --------------------------------------------------------------------------
DICT_PATH = _find_dictionary()
TRANSLATOR = FinancialTranslator(DICT_PATH)
CONVERTER = ExcelConverter(TRANSLATOR)

JOBS_ROOT = pathlib.Path(tempfile.gettempdir()) / "afx_webapp_jobs"
JOBS_ROOT.mkdir(parents=True, exist_ok=True)

ALLOWED_EXT = {".xlsx", ".xlsm", ".xls", ".csv", ".xltx"}
_SAFE = re.compile(r"[^A-Za-z0-9._\u0600-\u06FF\- ]+")


def _safe_name(name: str) -> str:
    name = os.path.basename(name or "file")
    name = _SAFE.sub("_", name).strip() or "file"
    return name[:150]


# --------------------------------------------------------------------------
# 3. Flask app — provided by DSS, or created here for standalone use
# --------------------------------------------------------------------------
from flask import request, jsonify, send_file  # noqa: E402

try:
    app  # type: ignore[used-before-def]  # DSS injects this
except NameError:
    from flask import Flask
    app = Flask(__name__)


# --------------------------------------------------------------------------
# 4. Conversion of one uploaded file -> a result dict for the UI
# --------------------------------------------------------------------------
def _convert_one(job_dir: pathlib.Path, fid: str, filename: str, data: bytes) -> dict:
    display = filename or "file"
    ext = pathlib.Path(display).suffix.lower()
    result = {"fid": fid, "name": display, "status": "error", "message": "",
              "sheets": 0, "rtl": 0, "reordered": 0, "review": 0, "output_name": None,
              "sample_terms": []}

    if ext not in ALLOWED_EXT:
        result["message"] = f"Unsupported type '{ext}'. Allowed: .xlsx .xlsm .xls .csv"
        return result

    in_dir = job_dir / "in"
    out_dir = job_dir / "out"
    in_dir.mkdir(parents=True, exist_ok=True)
    in_path = in_dir / _safe_name(display)
    in_path.write_bytes(data)

    try:
        rep = CONVERTER.convert_file(in_path, out_dir)
    except Exception as exc:  # per-file isolation — never blocks siblings
        result["message"] = str(exc)
        return result

    out_path = pathlib.Path(rep.output_path)
    review = sum(s.fuzzy + s.untranslated for s in rep.sheets)
    rtl = sum(1 for s in rep.sheets if s.rtl_source)
    reordered = sum(1 for s in rep.sheets if s.columns_reversed)

    # a few untranslated Arabic terms surfaced as "add these to the glossary"
    terms, seen = [], set()
    for s in rep.sheets:
        for (_sheet, _coord, src, _eng, method, _conf) in s.review_items:
            if method == "untranslated" and src not in seen:
                seen.add(src)
                terms.append(src)
            if len(terms) >= 6:
                break

    result.update(status="done", message="Converted",
                  sheets=len(rep.sheets), rtl=rtl, reordered=reordered, review=review,
                  output_name=out_path.name, sample_terms=terms)
    return result


# --------------------------------------------------------------------------
# 5. Routes
# --------------------------------------------------------------------------
@app.route("/api/health", methods=["GET"])
def health():
    stats = TRANSLATOR.stats()
    return jsonify({
        "status": "ok",
        "glossary": {
            "entries": stats.get("entries"),
            "arabic_variants": stats.get("arabic_variants"),
            "categories": len(stats.get("categories") or []),
        },
        "dictionary_path": os.path.basename(DICT_PATH),
        "accepted": sorted(ALLOWED_EXT),
    })


@app.route("/api/convert", methods=["POST"])
def convert():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded (field name must be 'files')."}), 400

    job_id = uuid.uuid4().hex[:12]
    job_dir = JOBS_ROOT / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    results, manifest = [], {}
    for i, storage in enumerate(files):
        data = storage.read()
        res = _convert_one(job_dir, f"f{i}", storage.filename, data)
        if res["status"] == "done":
            manifest[res["fid"]] = {
                "path": str((job_dir / "out" / res["output_name"]).resolve()),
                "name": res["output_name"],
            }
        results.append(res)

    (job_dir / "manifest.json").write_text(json.dumps(manifest), encoding="utf-8")

    done = sum(1 for r in results if r["status"] == "done")
    return jsonify({
        "job_id": job_id,
        "count": len(results),
        "succeeded": done,
        "failed": len(results) - done,
        "total_review": sum(r["review"] for r in results),
        "results": results,
    })


def _manifest(job_id: str) -> dict:
    mf = JOBS_ROOT / job_id / "manifest.json"
    if not mf.exists():
        return {}
    return json.loads(mf.read_text(encoding="utf-8"))


@app.route("/api/download/<job_id>/<fid>", methods=["GET"])
def download(job_id, fid):
    entry = _manifest(job_id).get(fid)
    if not entry or not os.path.exists(entry["path"]):
        return jsonify({"error": "Output not found or expired."}), 404
    return send_file(
        entry["path"], as_attachment=True, download_name=entry["name"],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/download_all/<job_id>", methods=["GET"])
def download_all(job_id):
    manifest = _manifest(job_id)
    if not manifest:
        return jsonify({"error": "Job not found or expired."}), 404
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for entry in manifest.values():
            if os.path.exists(entry["path"]):
                zf.write(entry["path"], arcname=entry["name"])
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"converted_{job_id}.zip",
                     mimetype="application/zip")


# convenience for standalone: a plain health check at root of the API
@app.route("/api/", methods=["GET"])
def _api_root():
    return jsonify({"service": "afx-webapp", "see": "/api/health"})
