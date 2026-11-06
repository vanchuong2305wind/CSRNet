"""Web app (Flask) đếm người & cảnh báo đám đông bất thường.

Chạy:   python server.py            → mở http://127.0.0.1:5000
Biến môi trường (tùy chọn):
    CROWD_WEIGHTS   đường dẫn trọng số (mặc định weights/csrnet_partB.pth)
    CROWD_HOST      địa chỉ lắng nghe (mặc định 127.0.0.1 — chỉ máy này truy cập được)
    CROWD_PORT      cổng (mặc định 5000)
    CONTACT_EMAIL   email hiện ở thanh liên hệ (bấm để sao chép); để trống thì ẩn

API:
    GET  /api/status                 trạng thái mô hình, ngưỡng, các mức cảnh báo
    POST /api/image                  phân tích một ảnh (multipart: file + tham số)
    POST /api/jobs                   tạo phiên phân tích video/camera → {id}
    GET  /api/jobs/<id>/events       luồng kết quả theo thời gian (Server-Sent Events)
    POST /api/jobs/<id>/stop         dừng phiên

Ảnh/video chỉ được xử lý trong bộ nhớ (video tải lên được ghi tạm rồi xóa ngay khi xong);
chỉ số liệu thống kê được trả về trình duyệt.
"""
import base64
import json
import os
import tempfile
import threading
import time
import uuid

import cv2
import numpy as np
import torch
from flask import Flask, Response, abort, jsonify, render_template, request

from crowd_core import (LEVELS, THRESHOLDS, CrowdMonitor, density_level, density_overlay,
                        load_model, predict, privacy_blur)

BASE = os.path.dirname(os.path.abspath(__file__))
WEIGHTS = os.environ.get("CROWD_WEIGHTS", os.path.join(BASE, "weights", "csrnet_partB.pth"))
DEVICE = "cuda" if torch.cuda.is_available() else "cpu"
CONTACT_EMAIL = os.environ.get("CONTACT_EMAIL", "")
VIDEO_EXT = {".mp4", ".avi", ".mov", ".mkv"}
IMAGE_EXT = {".jpg", ".jpeg", ".png", ".bmp"}
DISPLAY_SIDE = 960  # cạnh dài tối đa của ảnh gửi về trình duyệt

# Màu mức cảnh báo theo bảng màu đất của giao diện (Xanh, Vàng, Cam, Đỏ)
UI_LEVEL_COLORS = ["#325825", "#A8741A", "#C2531F", "#A3261A"]

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024

# ---------------------------------------------------------------------------- Mô hình
MODEL = None
if os.path.exists(WEIGHTS):
    MODEL = load_model(WEIGHTS, DEVICE)  # (model, density_scale, meta)
_infer_lock = threading.Lock()  # nhiều yêu cầu cùng lúc → suy luận lần lượt


def model_meta():
    if MODEL is None:
        return None
    meta = MODEL[2]
    mae = meta.get("val_mae")
    return {"dataset": str(meta.get("dataset", "CSRNet")), "epoch": meta.get("epoch", "?"),
            "val_mae": None if mae is None else round(float(mae), 2), "device": DEVICE}


def analyse(img_rgb, p):
    model, dscale, _ = MODEL
    with _infer_lock:
        count, den, img = predict(model, img_rgb, dscale, DEVICE, p["max_side"])
    shown = privacy_blur(img) if p["blur"] else img
    return count, density_overlay(shown, den), img


def to_data_url(img_rgb, quality=82):
    h, w = img_rgb.shape[:2]
    s = min(1.0, DISPLAY_SIDE / max(h, w))
    if s < 1:
        img_rgb = cv2.resize(img_rgb, (int(w * s), int(h * s)), interpolation=cv2.INTER_AREA)
    ok, buf = cv2.imencode(".jpg", cv2.cvtColor(img_rgb, cv2.COLOR_RGB2BGR), [cv2.IMWRITE_JPEG_QUALITY, quality])
    return "data:image/jpeg;base64," + base64.b64encode(buf.tobytes()).decode()


def read_params(src):
    """Đọc & giới hạn tham số người dùng gửi lên."""
    def num(key, default, lo, hi):
        try:
            v = float(src.get(key, default))
        except (TypeError, ValueError):
            v = default
        return min(hi, max(lo, v))

    crowd_type = src.get("crowd_type", "di chuyển")
    if crowd_type not in THRESHOLDS:
        crowd_type = "di chuyển"
    max_side = int(num("max_side", 1024, 256, 1600))
    return {
        "area": num("area", 100.0, 1.0, 100000.0),
        "crowd_type": crowd_type,
        "blur": str(src.get("blur", "1")).lower() in ("1", "true", "on", "yes"),
        "max_side": max_side,
        "stride": num("stride", 1.0, 0.2, 5.0),
        "surge_ratio": num("surge_ratio", 50, 10, 200) / 100,
        "surge_window": num("surge_window", 10, 2, 60),
        "min_persist": int(num("min_persist", 3, 1, 10)),
        "max_secs": num("max_secs", 60, 10, 3600),
    }


def require_model():
    if MODEL is None:
        abort(Response(json.dumps({"error": f"Không tìm thấy trọng số {WEIGHTS}. Hãy chép csrnet_partB.pth "
                                             "vào thư mục weights/."}, ensure_ascii=False),
                       status=503, mimetype="application/json"))


def ext_of(name):
    return os.path.splitext(name or "")[1].lower()


# ---------------------------------------------------------------------------- Trang
@app.get("/")
def index():
    return render_template("index.html", contact_email=CONTACT_EMAIL)


@app.get("/api/status")
def status():
    return jsonify({
        "ok": MODEL is not None,
        "model": model_meta(),
        "weights": os.path.relpath(WEIGHTS, BASE),
        "thresholds": THRESHOLDS,
        "levels": [dict(lv, color=UI_LEVEL_COLORS[i]) for i, lv in enumerate(LEVELS)],
    })


# ---------------------------------------------------------------------------- Ảnh
@app.post("/api/image")
def api_image():
    require_model()
    f = request.files.get("file")
    if f is None or ext_of(f.filename) not in IMAGE_EXT:
        return jsonify({"error": "Hãy chọn một ảnh JPG, PNG hoặc BMP."}), 400
    img = cv2.imdecode(np.frombuffer(f.read(), np.uint8), cv2.IMREAD_COLOR)
    if img is None:
        return jsonify({"error": "Không đọc được ảnh."}), 400
    p = read_params(request.form)
    t0 = time.time()
    count, vis, resized = analyse(cv2.cvtColor(img, cv2.COLOR_BGR2RGB), p)
    ms = (time.time() - t0) * 1000
    rho = count / p["area"]
    return jsonify({
        "count": count, "density": rho, "level": density_level(rho, p["crowd_type"]), "ms": ms,
        "overlay": to_data_url(vis),
        "original": None if p["blur"] else to_data_url(resized),
    })


# ---------------------------------------------------------------------------- Video / camera
JOBS = {}
_jobs_lock = threading.Lock()
JOB_TTL = 600  # phiên chưa được mở luồng sau 10 phút thì bị hủy


def _cleanup_job(job):
    if job.get("tmp") and os.path.exists(job["tmp"]):
        try:
            os.remove(job["tmp"])
        except OSError:
            pass


def _gc_jobs():
    now = time.time()
    with _jobs_lock:
        stale = [k for k, j in JOBS.items() if not j["started"] and now - j["created"] > JOB_TTL]
        for k in stale:
            _cleanup_job(JOBS.pop(k))


@app.post("/api/jobs")
def create_job():
    require_model()
    _gc_jobs()
    p = read_params(request.form)
    source = request.form.get("source", "file")
    job = {"params": p, "stop": False, "started": False, "created": time.time(), "tmp": None}
    if source == "file":
        f = request.files.get("file")
        if f is None or ext_of(f.filename) not in VIDEO_EXT:
            return jsonify({"error": "Hãy chọn video MP4, AVI, MOV hoặc MKV."}), 400
        fd, tmp = tempfile.mkstemp(suffix=ext_of(f.filename))
        with os.fdopen(fd, "wb") as out:
            f.save(out)
        job.update(kind="file", src=tmp, tmp=tmp)
    else:
        url = (request.form.get("url") or "0").strip()
        job.update(kind="camera", src=int(url) if url.isdigit() else url)
    jid = uuid.uuid4().hex
    with _jobs_lock:
        JOBS[jid] = job
    return jsonify({"id": jid})


@app.post("/api/jobs/<jid>/stop")
def stop_job(jid):
    with _jobs_lock:
        job = JOBS.get(jid)
    if job:
        job["stop"] = True
    return jsonify({"ok": True})


def _sse(event, data):
    return f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n"


def run_job(jid, job):
    p = job["params"]
    is_file = job["kind"] == "file"
    cap = cv2.VideoCapture(job["src"])
    summary = {"samples": 0, "max_count": 0.0, "max_level": 0, "stopped": False}
    try:
        if not cap.isOpened():
            yield _sse("fail", {"error": "Không mở được nguồn video."})
            return
        fps = cap.get(cv2.CAP_PROP_FPS) or 25
        total = cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0
        mon = CrowdMonitor(p["area"], p["crowd_type"], p["surge_ratio"], p["surge_window"],
                           min_persist=p["min_persist"])
        max_secs = None if is_file else p["max_secs"]
        next_t, t_start, idx = 0.0, time.time(), 0
        yield _sse("start", {"kind": job["kind"], "fps": fps, "duration": total / fps if total else None})
        while not job["stop"]:
            ok, frame = cap.read()
            if not ok:
                break
            tsec = idx / fps if is_file else time.time() - t_start
            idx += 1
            if max_secs and tsec > max_secs:
                break
            if tsec < next_t:
                continue
            next_t = tsec + p["stride"]
            count, vis, _ = analyse(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB), p)
            r = mon.update(tsec, count)
            summary["samples"] += 1
            summary["max_count"] = max(summary["max_count"], count)
            summary["max_level"] = max(summary["max_level"], r["level"])
            progress = min(1.0, idx / total) if is_file and total else (min(1.0, tsec / max_secs) if max_secs else None)
            yield _sse("tick", {"t": tsec, "count": count, "smoothed": r["smoothed"], "density": r["density"],
                                "level": r["level"], "surge": r["surge"], "events": r["events"],
                                "progress": progress, "frame": to_data_url(vis, 72)})
        summary["stopped"] = job["stop"]
        yield _sse("done", summary)
    except Exception as e:  # báo lỗi cho trình duyệt thay vì cắt luồng im lặng
        yield _sse("fail", {"error": f"Lỗi khi phân tích: {e}"})
    finally:
        cap.release()
        _cleanup_job(job)
        with _jobs_lock:
            JOBS.pop(jid, None)


@app.get("/api/jobs/<jid>/events")
def job_events(jid):
    with _jobs_lock:
        job = JOBS.get(jid)
        if job is None or job["started"]:
            return jsonify({"error": "Phiên không tồn tại hoặc đã chạy."}), 404
        job["started"] = True
    return Response(run_job(jid, job), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.errorhandler(413)
def too_large(_):
    return jsonify({"error": "Tệp quá lớn (tối đa 500 MB)."}), 413


if __name__ == "__main__":
    host = os.environ.get("CROWD_HOST", "127.0.0.1")
    port = int(os.environ.get("CROWD_PORT", "5000"))
    print(f"Mô hình: {'đã nạp' if MODEL else 'CHƯA có trọng số'} · {DEVICE} · http://{host}:{port}")
    app.run(host=host, port=port, threaded=True, debug=False)
