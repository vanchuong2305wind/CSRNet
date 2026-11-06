"""App demo: đếm người & cảnh báo đám đông bất thường.

Chạy:  streamlit run app.py
Cần file trọng số do notebook xuất ra: weights/csrnet_partB.pth
"""
import os
import tempfile
import time

import cv2
import numpy as np
import pandas as pd
import streamlit as st
import torch

from crowd_core import (LEVELS, THRESHOLDS, CrowdMonitor, density_level, density_overlay,
                        load_model, predict, privacy_blur)

st.set_page_config(page_title="Crowd Monitor", page_icon="👥", layout="wide")
DEVICE = "cuda" if torch.cuda.is_available() else "cpu"


@st.cache_resource
def get_model(path):
    return load_model(path, DEVICE)


def level_badge(level, extra=""):
    lv = LEVELS[level]
    st.markdown(
        f"<div style='background:{lv['color']};color:white;padding:14px 18px;border-radius:8px'>"
        f"<div style='font-size:26px;font-weight:700'>MỨC {lv['name']} – {lv['label']}</div>"
        f"<div>{lv['action']}{extra}</div></div>",
        unsafe_allow_html=True)


# ---------------------------------------------------------------- Sidebar
st.sidebar.header("Cấu hình")
weights = st.sidebar.text_input("File trọng số", "weights/csrnet_partB.pth")
area = st.sidebar.number_input("Diện tích vùng camera quan sát (m²)", 1.0, 100000.0, 100.0, 10.0,
                               help="Diện tích thực tế của khu vực trong khung hình, đo/hiệu chỉnh khi lắp camera.")
crowd_type = st.sidebar.radio("Loại đám đông", list(THRESHOLDS), horizontal=True)
t = THRESHOLDS[crowd_type]
st.sidebar.caption(f"Ngưỡng (người/m²): Vàng ≥ {t[0]:g} · Cam ≥ {t[1]:g} · Đỏ ≥ {t[2]:g}")
blur = st.sidebar.checkbox("Chế độ riêng tư (làm mờ ảnh hiển thị)", value=True)
max_side = st.sidebar.select_slider("Độ phân giải xử lý (cạnh dài)", [512, 768, 1024, 1280], 1024)

st.sidebar.subheader("Phát hiện bất thường (video)")
stride = st.sidebar.slider("Phân tích mỗi (giây)", 0.2, 5.0, 1.0, 0.2)
surge_ratio = st.sidebar.slider("Ngưỡng tăng đột biến (%)", 10, 200, 50, 10) / 100
surge_window = st.sidebar.slider("Cửa sổ thời gian tăng đột biến (giây)", 2, 60, 10)
min_persist = st.sidebar.slider("Số lần đo liên tiếp để đổi mức", 1, 10, 3,
                                help="Chống cảnh báo chập chờn: mức chỉ thay đổi khi được xác nhận nhiều lần liên tiếp.")

st.title("👥 Hệ thống đếm người & cảnh báo đám đông bất thường")
st.caption("Pipeline: khung hình → CSRNet (bản đồ mật độ) → số người → mật độ người/m² → luật cảnh báo. "
           "Ảnh/video chỉ được xử lý trong bộ nhớ, không lưu lại; chỉ số liệu thống kê được ghi nhật ký.")

if not os.path.exists(weights):
    st.error(f"Không tìm thấy `{weights}`. Hãy chạy notebook huấn luyện, tải `csrnet_partB.pth` "
             "và chép vào thư mục `weights/`.")
    st.stop()
model, dscale, meta = get_model(weights)
st.sidebar.success(f"Mô hình: {meta.get('dataset', 'CSRNet')} · epoch {meta.get('epoch', '?')} · "
                   f"val MAE {meta.get('val_mae', float('nan')):.2f} · {DEVICE}")


def analyse(img_rgb):
    count, den, img = predict(model, img_rgb, dscale, DEVICE, max_side)
    shown = privacy_blur(img) if blur else img
    return count, density_overlay(shown, den)


tab_img, tab_vid = st.tabs(["🖼️ Ảnh", "🎞️ Video / Camera"])

# ---------------------------------------------------------------- Ảnh
with tab_img:
    up = st.file_uploader("Chọn ảnh đám đông", type=["jpg", "jpeg", "png", "bmp"])
    if up:
        img = cv2.imdecode(np.frombuffer(up.read(), np.uint8), cv2.IMREAD_COLOR)
        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        t0 = time.time()
        count, vis = analyse(img)
        rho = count / area
        c1, c2 = st.columns([3, 2])
        c1.image(vis, caption="Bản đồ nhiệt mật độ", use_container_width=True)
        with c2:
            m1, m2, m3 = st.columns(3)
            m1.metric("Số người", f"{count:.0f}")
            m2.metric("Mật độ (người/m²)", f"{rho:.2f}")
            m3.metric("Thời gian", f"{(time.time() - t0) * 1000:.0f} ms")
            level_badge(density_level(rho, crowd_type))
            if not blur:
                st.image(img, caption="Ảnh gốc", use_container_width=True)

# ---------------------------------------------------------------- Video
with tab_vid:
    src = st.radio("Nguồn", ["Tải video lên", "Camera (webcam / RTSP)"], horizontal=True)
    cap, tmp_path, max_secs = None, None, None
    if src == "Tải video lên":
        vf = st.file_uploader("Chọn video", type=["mp4", "avi", "mov", "mkv"])
        if vf and st.button("▶️ Phân tích video", type="primary"):
            fd, tmp_path = tempfile.mkstemp(suffix=os.path.splitext(vf.name)[1])
            with os.fdopen(fd, "wb") as f:
                f.write(vf.read())
            cap = cv2.VideoCapture(tmp_path)
    else:
        url = st.text_input("Chỉ số webcam (0, 1…) hoặc URL RTSP", "0")
        max_secs = st.number_input("Thời gian theo dõi (giây)", 10, 3600, 60)
        if st.button("▶️ Bắt đầu", type="primary"):
            cap = cv2.VideoCapture(int(url) if url.isdigit() else url)

    if cap is not None:
        if not cap.isOpened():
            st.error("Không mở được nguồn video.")
        else:
            is_file = tmp_path is not None
            fps = cap.get(cv2.CAP_PROP_FPS) or 25
            mon = CrowdMonitor(area, crowd_type, surge_ratio, surge_window, min_persist=min_persist)
            c1, c2 = st.columns([3, 2])
            frame_ph = c1.empty()
            with c2:
                metrics_ph, badge_ph, chart_ph = st.empty(), st.empty(), st.empty()
            prog = st.progress(0.0) if is_file else None
            total = cap.get(cv2.CAP_PROP_FRAME_COUNT) or 1
            rows, log = [], []
            next_t, t_start, idx = 0.0, time.time(), 0
            try:
                while True:
                    ok, frame = cap.read()
                    if not ok:
                        break
                    tsec = idx / fps if is_file else time.time() - t_start
                    idx += 1
                    if max_secs and tsec > max_secs:
                        break
                    if tsec < next_t:
                        continue
                    next_t = tsec + stride
                    count, vis = analyse(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
                    r = mon.update(tsec, count)
                    rows.append({"t (s)": round(tsec, 1), "Số người": round(count, 1), "Mật độ (/m²)": round(r["density"], 3),
                                 "Mức": LEVELS[r["level"]]["name"], "Tăng đột biến": r["surge"]})
                    for e in r["events"]:
                        log.append({"t (s)": round(tsec, 1), "Mức": LEVELS[r["level"]]["name"], "Sự kiện": e})

                    frame_ph.image(vis, caption=f"t = {tsec:.1f}s", use_container_width=True)
                    with metrics_ph.container():
                        m1, m2 = st.columns(2)
                        m1.metric("Số người", f"{count:.0f}")
                        m2.metric("Mật độ (người/m²)", f"{r['density']:.2f}")
                    with badge_ph.container():
                        level_badge(r["level"], " · ⚠️ Tăng đột biến!" if r["surge"] else "")
                    chart_ph.line_chart(pd.DataFrame(rows).set_index("t (s)")[["Số người"]], height=200)
                    if prog:
                        prog.progress(min(1.0, idx / total))
            finally:
                cap.release()
                if tmp_path:
                    os.remove(tmp_path)

            st.subheader("Nhật ký cảnh báo")
            if log:
                st.dataframe(pd.DataFrame(log), use_container_width=True, hide_index=True)
            else:
                st.info("Không có sự kiện bất thường.")
            if rows:
                df = pd.DataFrame(rows)
                st.download_button("⬇️ Tải số liệu (CSV)", df.to_csv(index=False).encode("utf-8-sig"),
                                   "crowd_log.csv", "text/csv")
