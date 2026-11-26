"""App local: PDF scan và ảnh chụp bảng → Excel."""

from __future__ import annotations

import sys

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")

import subprocess
import threading
import uuid
from pathlib import Path

import io

from flask import Flask, jsonify, request, send_file, send_from_directory

from engine import convert_files, get_ocr, ocr_status

ROOT = Path(__file__).resolve().parent
app = Flask(__name__, static_folder=str(ROOT / "static"), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024

ALLOWED = {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".gif"}
jobs: dict[str, dict] = {}
jobs_lock = threading.Lock()
work_lock = threading.Lock()


def _job(job_id: str) -> dict | None:
    with jobs_lock:
        return jobs.get(job_id)


def _update(job_id: str, **fields) -> None:
    with jobs_lock:
        if job_id in jobs:
            jobs[job_id].update(fields)


def _run(job_id: str, files: list[tuple[str, bytes]], borderless: bool, keep_outside: bool) -> None:
    with work_lock:
        try:
            _update(job_id, status="running", progress=8, message="Đang mở bộ đọc tiếng Việt…")
            get_ocr()

            def on_page(message: str) -> None:
                job = _job(job_id) or {}
                progress = min(92, int(job.get("progress", 12)) + 4)
                _update(job_id, status="running", progress=progress, message=message)

            result = convert_files(
                files,
                borderless=borderless,
                keep_outside=keep_outside,
                on_page=on_page,
            )
            sheets = [
                {k: v for k, v in sheet.items() if k != "kind"}
                for sheet in result["sheets"]
            ]
            _update(
                job_id,
                status="done",
                progress=100,
                message="Đã đọc xong.",
                sheets=sheets,
                tables=result["tables"],
                pages=result["pages"],
                note=result["note"],
                xlsx=result["xlsx"],
            )
        except Exception as exc:
            _update(job_id, status="error", progress=100, message=str(exc))


@app.get("/")
def home():
    return send_from_directory(app.static_folder, "index.html")


@app.get("/api/ready")
def ready():
    return jsonify(ocr_status())


@app.post("/api/jobs")
def create_job():
    uploads = request.files.getlist("files")
    picked = []
    for item in uploads:
        if not item or not item.filename:
            continue
        suffix = Path(item.filename).suffix.lower()
        if suffix not in ALLOWED:
            return jsonify(error=f"Không đọc được đuôi {suffix}. Dùng PDF hoặc ảnh."), 400
        data = item.read()
        if not data:
            continue
        if len(data) > 50 * 1024 * 1024:
            return jsonify(error=f"{item.filename} lớn hơn 50MB."), 400
        picked.append((Path(item.filename).name, data))
    if not picked:
        return jsonify(error="Chưa chọn file."), 400
    if len(picked) > 20:
        return jsonify(error="Mỗi lần tối đa 20 file."), 400

    borderless = request.form.get("borderless", "1") != "0"
    keep_outside = request.form.get("keep_outside", "1") != "0"
    job_id = uuid.uuid4().hex
    with jobs_lock:
        jobs[job_id] = {
            "status": "queued",
            "progress": 2,
            "message": "Đã nhận file.",
            "sheets": [],
            "tables": 0,
            "pages": 0,
            "note": "",
            "xlsx": b"",
        }
    threading.Thread(
        target=_run,
        args=(job_id, picked, borderless, keep_outside),
        daemon=True,
    ).start()
    return jsonify(id=job_id)


@app.get("/api/jobs/<job_id>")
def job_status(job_id: str):
    job = _job(job_id)
    if not job:
        return jsonify(error="Không thấy lượt đọc này."), 404
    return jsonify(
        status=job["status"],
        progress=job["progress"],
        message=job["message"],
        sheets=job["sheets"],
        tables=job["tables"],
        pages=job["pages"],
        note=job["note"],
    )


@app.get("/api/jobs/<job_id>/xlsx")
def download(job_id: str):
    job = _job(job_id)
    if not job or job["status"] != "done" or not job["xlsx"]:
        return jsonify(error="File Excel chưa sẵn sàng."), 404
    return send_file(
        io.BytesIO(job["xlsx"]),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="bang-ra-excel.xlsx",
    )


def _warm() -> None:
    try:
        get_ocr()
    except Exception:
        pass


def _edge_path() -> Path | None:
    candidates = [
        Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
        Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
    ]
    for path in candidates:
        if path.exists():
            return path
    return None


def open_app_window() -> None:
    """Cửa sổ riêng, không thanh địa chỉ — giống app cài trên máy."""
    edge = _edge_path()
    if edge is None:
        import webbrowser

        webbrowser.open("http://127.0.0.1:8765")
        return
    profile = ROOT / ".edge-profile"
    profile.mkdir(exist_ok=True)
    subprocess.Popen(
        [
            str(edge),
            "--app=http://127.0.0.1:8765",
            "--window-size=1120,840",
            f"--user-data-dir={profile}",
            "--no-first-run",
            "--disable-extensions",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def main() -> None:
    import socket

    sock = socket.socket()
    try:
        sock.bind(("127.0.0.1", 8765))
    except OSError:
        open_app_window()
        print("App đang mở.")
        return
    finally:
        sock.close()

    threading.Thread(target=_warm, daemon=True).start()
    threading.Timer(0.8, open_app_window).start()
    print("Đang mở Bảng ra Excel.")
    app.run(host="127.0.0.1", port=8765, threaded=True, debug=False)


if __name__ == "__main__":
    main()
