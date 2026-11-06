/* Crowd Monitor: giao diện phía trình duyệt (không phụ thuộc thư viện) */
(() => {
  "use strict";
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const fmt = (v, d = 0) => Number(v).toLocaleString("vi-VN", { minimumFractionDigits: d, maximumFractionDigits: d });

  let STATUS = { ok: false, levels: [], thresholds: {} };

  // ------------------------------------------------------------ Cấu hình (lưu trên trình duyệt)
  const MAX_SIDES = [512, 768, 1024, 1280];
  const DEFAULTS = { area: 100, crowd_type: "di chuyển", blur: true, max_side: 1024, stride: 1, surge_ratio: 50, surge_window: 10, min_persist: 3 };
  let S = { ...DEFAULTS };
  try { Object.assign(S, JSON.parse(localStorage.getItem("crowd-settings") || "{}")); } catch (_) { /* bỏ qua */ }
  const save = () => { try { localStorage.setItem("crowd-settings", JSON.stringify(S)); } catch (_) { /* bỏ qua */ } };

  function bindSettings() {
    const area = $("#s-area"), blur = $("#s-blur");
    const ranges = [
      ["#s-stride", "stride", "#o-stride", (v) => `${fmt(v, 1)} giây`],
      ["#s-surge", "surge_ratio", "#o-surge", (v) => `${v}%`],
      ["#s-window", "surge_window", "#o-window", (v) => `${v} giây`],
      ["#s-persist", "min_persist", "#o-persist", (v) => `${v} lần`],
    ];
    area.value = S.area;
    area.addEventListener("change", () => {
      const v = Math.min(100000, Math.max(1, parseFloat(area.value) || DEFAULTS.area));
      area.value = S.area = v; save();
    });
    blur.checked = !!S.blur;
    blur.addEventListener("change", () => { S.blur = blur.checked; save(); syncPrivacyTag(); });
    for (const [sel, key, out, f] of ranges) {
      const el = $(sel);
      el.value = S[key];
      const show = () => { $(out).textContent = f(Number(el.value)); };
      el.addEventListener("input", () => { S[key] = Number(el.value); show(); save(); });
      show();
    }
    const ms = $("#s-maxside");
    ms.value = Math.max(0, MAX_SIDES.indexOf(S.max_side));
    const showMs = () => { S.max_side = MAX_SIDES[ms.value]; $("#o-maxside").textContent = `${S.max_side}px`; };
    ms.addEventListener("input", () => { showMs(); save(); });
    showMs();
    $$("#s-type button").forEach((b) => b.addEventListener("click", () => { S.crowd_type = b.dataset.type; save(); syncType(); }));
    syncType(); syncPrivacyTag();
  }
  function syncType() {
    $$("#s-type button").forEach((b) => b.classList.toggle("active", b.dataset.type === S.crowd_type));
    const t = STATUS.thresholds[S.crowd_type];
    $("#s-thr").textContent = t ? `Ngưỡng (người/m²): Vàng ≥ ${t[0]} · Cam ≥ ${t[1]} · Đỏ ≥ ${t[2]}` : "";
  }
  function syncPrivacyTag() { $("#tag-privacy").classList.toggle("off", !S.blur); }
  function formData() {
    const fd = new FormData();
    for (const [k, v] of Object.entries(S)) fd.append(k, typeof v === "boolean" ? (v ? "1" : "0") : v);
    return fd;
  }

  // ------------------------------------------------------------ Điều hướng
  const nav = $("#navpill"), flyout = $("#settings"), toggle = $("#nav-toggle");
  function go(view) {
    $$(".view").forEach((v) => { v.hidden = v.dataset.view !== view; });
    $$(".nav-item[data-go]").forEach((b) => b.classList.toggle("active", b.dataset.go === view));
    closeMenu();
    history.replaceState(null, "", `#${view}`);
  }
  function setSettings(open) {
    flyout.hidden = !open;
    $$("[data-open-settings].nav-item").forEach((b) => b.setAttribute("aria-expanded", String(open)));
    if (open) $("#s-area").focus({ preventScroll: true });
  }
  function closeMenu() { nav.classList.remove("open"); toggle.setAttribute("aria-expanded", "false"); }
  $$("[data-go]").forEach((b) => b.addEventListener("click", (e) => { e.preventDefault(); go(b.dataset.go); }));
  $$("[data-open-settings]").forEach((b) => b.addEventListener("click", (e) => { e.preventDefault(); setSettings(flyout.hidden); }));
  $("#settings-close").addEventListener("click", () => setSettings(false));
  toggle.addEventListener("click", () => {
    const open = !nav.classList.contains("open");
    nav.classList.toggle("open", open);
    toggle.setAttribute("aria-expanded", String(open));
    if (!open) setSettings(false);
  });
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") { setSettings(false); closeMenu(); } });
  document.addEventListener("click", (e) => {
    if (!nav.contains(e.target) && !e.target.closest("[data-open-settings]")) { setSettings(false); closeMenu(); }
  });

  // ------------------------------------------------------------ Toast + sao chép email
  let toastTimer;
  function toast(msg) {
    const t = $("#toast");
    t.textContent = msg; t.classList.add("show");
    clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove("show"), 2200);
  }
  $$(".copy-email").forEach((b) => b.addEventListener("click", async () => {
    try { await navigator.clipboard.writeText(b.dataset.email); toast(`Đã sao chép ${b.dataset.email}`); }
    catch (_) { toast(b.dataset.email); }
  }));

  // ------------------------------------------------------------ Thẻ mức cảnh báo
  function levelCard(level, extra = "") {
    const L = STATUS.levels;
    const lv = L[level]; if (!lv) return "";
    const scale = L.map((x, i) => `<span style="${i <= level ? `background:${L[i].color}` : ""}"></span>`).join("");
    return `<div class="level-card" style="--lv:${lv.color}">
      <div class="lv-top"><span class="lv-chip">MỨC ${esc(lv.name)}</span><span class="lv-title">${esc(lv.label)}</span></div>
      <div class="lv-action">${esc(lv.action)}</div>
      ${extra ? `<div class="lv-extra">${esc(extra)}</div>` : ""}
      <div class="lv-scale">${scale}</div></div>`;
  }

  // ------------------------------------------------------------ Vùng thả tệp
  function dropzone(zone, input, onFile) {
    input.addEventListener("change", () => { if (input.files[0]) onFile(input.files[0]); input.value = ""; });
    ["dragenter", "dragover"].forEach((ev) => zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.add("drag"); }));
    ["dragleave", "drop"].forEach((ev) => zone.addEventListener(ev, (e) => { e.preventDefault(); zone.classList.remove("drag"); }));
    zone.addEventListener("drop", (e) => { const f = e.dataTransfer.files[0]; if (f) onFile(f); });
  }
  async function readError(res) {
    try { return (await res.json()).error || `Lỗi ${res.status}`; } catch (_) { return `Lỗi ${res.status}`; }
  }

  // ------------------------------------------------------------ 01 · Ảnh
  let imgData = null;
  async function analyseImage(file) {
    const err = $("#img-error");
    err.hidden = true; $("#img-loading").hidden = false;
    const fd = formData(); fd.append("file", file);
    try {
      const res = await fetch("/api/image", { method: "POST", body: fd });
      if (!res.ok) throw new Error(await readError(res));
      imgData = await res.json();
      $("#img-empty").hidden = true; $("#img-result").hidden = false;
      $("#img-count").textContent = fmt(imgData.count);
      $("#img-density").textContent = fmt(imgData.density, 2);
      $("#img-level").innerHTML = levelCard(imgData.level);
      $("#img-time").textContent = `${file.name} · xử lý trong ${fmt(imgData.ms)} ms`;
      $("#img-seg").hidden = !imgData.original;
      showImg("overlay");
    } catch (e) {
      err.textContent = e.message; err.hidden = false;
    } finally { $("#img-loading").hidden = true; }
  }
  function showImg(which) {
    if (!imgData) return;
    $("#img-out").src = imgData[which] || imgData.overlay;
    $("#img-caption").textContent = which === "original" ? "Ảnh gốc" : (S.blur ? "Bản đồ nhiệt mật độ (đã làm mờ)" : "Bản đồ nhiệt mật độ");
    $$("#img-seg button").forEach((b) => b.classList.toggle("active", b.dataset.show === which));
  }
  $$("#img-seg button").forEach((b) => b.addEventListener("click", () => showImg(b.dataset.show)));
  dropzone($("#img-drop"), $("#img-input"), analyseImage);

  // ------------------------------------------------------------ 02 · Video / camera
  let source = "file", videoFile = null, job = null, rows = [], logRows = [];
  $$("#src-seg button").forEach((b) => b.addEventListener("click", () => {
    if (job) return;
    source = b.dataset.src;
    $$("#src-seg button").forEach((x) => x.classList.toggle("active", x === b));
    $("#src-file").hidden = source !== "file"; $("#src-camera").hidden = source !== "camera";
    $("#run-btn").textContent = source === "file" ? "Phân tích video" : "Bắt đầu theo dõi";
  }));
  dropzone($("#vid-drop"), $("#vid-input"), (f) => { videoFile = f; $("#vid-name").textContent = f.name; });

  function setRunning(on, text = "") {
    $("#run-btn").disabled = on; $("#stop-btn").hidden = !on;
    $("#vid-loading").hidden = !on; $("#vid-loading-text").textContent = text;
  }
  function fail(msg) {
    const err = $("#vid-error"); err.textContent = msg; err.hidden = false;
  }

  $("#run-btn").addEventListener("click", async () => {
    $("#vid-error").hidden = true;
    if (!STATUS.ok) return fail("Chưa nạp được mô hình.");
    const fd = formData(); fd.append("source", source);
    if (source === "file") {
      if (!videoFile) return fail("Hãy chọn một video trước.");
      fd.append("file", videoFile);
    } else {
      fd.append("url", $("#cam-url").value.trim() || "0");
      fd.append("max_secs", $("#cam-secs").value || 60);
    }
    rows = []; logRows = [];
    $("#log-wrap").hidden = true; $("#live").hidden = true;
    setRunning(true, source === "file" ? "Đang tải video lên…" : "Đang mở camera…");
    let id;
    try {
      const res = await fetch("/api/jobs", { method: "POST", body: fd });
      if (!res.ok) throw new Error(await readError(res));
      id = (await res.json()).id;
    } catch (e) { setRunning(false); return fail(e.message); }

    $("#vid-loading-text").textContent = "Đang phân tích…";
    const es = new EventSource(`/api/jobs/${id}/events`);
    job = { id, es };
    es.addEventListener("start", () => { $("#live").hidden = false; });
    es.addEventListener("tick", (e) => onTick(JSON.parse(e.data)));
    es.addEventListener("done", (e) => finish(JSON.parse(e.data)));
    es.addEventListener("fail", (e) => { fail(JSON.parse(e.data).error); finish(null); });
    es.onerror = () => { if (job && job.es === es) { fail("Mất kết nối tới máy chủ."); finish(null); } };
  });
  $("#stop-btn").addEventListener("click", () => {
    if (!job) return;
    fetch(`/api/jobs/${job.id}/stop`, { method: "POST" });
    $("#vid-loading-text").textContent = "Đang dừng…";
  });

  function onTick(d) {
    $("#live").hidden = false;
    rows.push(d);
    const L = STATUS.levels;
    for (const ev of d.events) logRows.push({ t: d.t, level: d.level, text: ev });
    $("#live-frame").src = d.frame;
    $("#live-caption").textContent = `t = ${fmt(d.t, 1)}s`;
    $("#live-count").textContent = fmt(d.count);
    $("#live-density").textContent = fmt(d.density, 2);
    $("#live-level").innerHTML = levelCard(d.level, d.surge ? "⚠ Tăng đột biến!" : "");
    $("#live-progress span").style.width = d.progress == null ? "100%" : `${(d.progress * 100).toFixed(1)}%`;
    drawChart();
    if (d.events.length && L[d.level]) toast(`${d.events[d.events.length - 1]}`);
  }

  function finish(summary) {
    if (!job) return;
    job.es.close(); job = null;
    setRunning(false);
    if (!rows.length && !summary) return;
    $("#log-wrap").hidden = false;
    const L = STATUS.levels;
    $("#log-body").innerHTML = logRows.map((r) => `<tr><td>${fmt(r.t, 1)}</td>
      <td><span class="chip" style="--lv:${L[r.level].color}">${esc(L[r.level].name)}</span></td><td>${esc(r.text)}</td></tr>`).join("");
    $("#log-empty").hidden = logRows.length > 0;
    $(".log").hidden = logRows.length === 0;
    if (summary) {
      $("#summary").textContent = `${summary.samples} lần đo · tối đa ${fmt(summary.max_count)} người · mức cao nhất ${L[summary.max_level].name}${summary.stopped ? " · đã dừng" : ""}`;
    }
    $("#csv-btn").disabled = rows.length === 0;
  }

  $("#csv-btn").addEventListener("click", () => {
    const L = STATUS.levels;
    const head = ["t (s)", "Số người", "Số người (làm trơn)", "Mật độ (/m²)", "Mức", "Tăng đột biến"];
    const lines = [head.join(",")].concat(rows.map((r) => [r.t.toFixed(1), r.count.toFixed(1), r.smoothed.toFixed(1),
      r.density.toFixed(3), L[r.level].name, r.surge ? "True" : "False"].join(",")));
    const blob = new Blob(["﻿" + lines.join("\n")], { type: "text/csv;charset=utf-8" });
    const a = Object.assign(document.createElement("a"), { href: URL.createObjectURL(blob), download: "crowd_log.csv" });
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  });

  // ------------------------------------------------------------ Biểu đồ (canvas)
  function drawChart() {
    const cv = $("#chart"), ctx = cv.getContext("2d");
    const dpr = window.devicePixelRatio || 1, W = cv.clientWidth, H = cv.clientHeight;
    if (cv.width !== W * dpr) { cv.width = W * dpr; cv.height = H * dpr; }
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, W, H);
    if (!rows.length) return;
    const pad = { l: 36, r: 8, t: 10, b: 22 };
    const tMax = Math.max(rows[rows.length - 1].t, 1), tMin = rows[0].t;
    const thr = (STATUS.thresholds[S.crowd_type] || []).map((x) => x * S.area);
    const peak = Math.max(...rows.map((r) => Math.max(r.count, r.smoothed)));
    let yMax = Math.max(peak * 1.15, 5);
    const visibleThr = thr.filter((v) => v <= yMax * 1.6);
    if (visibleThr.length) yMax = Math.max(yMax, Math.max(...visibleThr) * 1.08);
    const x = (t) => pad.l + ((t - tMin) / Math.max(tMax - tMin, 1e-6)) * (W - pad.l - pad.r);
    const y = (v) => H - pad.b - (v / yMax) * (H - pad.t - pad.b);
    ctx.font = "11px Manrope, sans-serif"; ctx.fillStyle = "rgba(75,45,21,.6)";
    // lưới
    ctx.strokeStyle = "rgba(75,45,21,.1)"; ctx.lineWidth = 1;
    for (let i = 0; i <= 4; i++) {
      const v = (yMax * i) / 4, yy = Math.round(y(v)) + .5;
      ctx.beginPath(); ctx.moveTo(pad.l, yy); ctx.lineTo(W - pad.r, yy); ctx.stroke();
      ctx.textAlign = "right"; ctx.fillText(fmt(v), pad.l - 6, yy + 4);
    }
    ctx.textAlign = "left"; ctx.fillText(`${fmt(tMin, 0)}s`, pad.l, H - 6);
    ctx.textAlign = "right"; ctx.fillText(`${fmt(tMax, 0)}s`, W - pad.r, H - 6);
    // ngưỡng mức cảnh báo (quy ra số người)
    ctx.setLineDash([4, 4]);
    let lastLabelY = Infinity;
    thr.forEach((v, i) => {
      if (v > yMax) return;
      const c = STATUS.levels[i + 1].color, yy = Math.round(y(v)) + .5;
      ctx.strokeStyle = c; ctx.globalAlpha = .7;
      ctx.beginPath(); ctx.moveTo(pad.l, yy); ctx.lineTo(W - pad.r, yy); ctx.stroke();
      ctx.globalAlpha = 1;
      if (lastLabelY - yy >= 13) {  // bỏ nhãn nếu chồng lên nhãn trước
        ctx.fillStyle = c; ctx.textAlign = "right";
        ctx.fillText(STATUS.levels[i + 1].name, W - pad.r - 2, yy - 4);
        lastLabelY = yy;
      }
    });
    ctx.setLineDash([]);
    // vùng + đường số đo
    const line = (key) => { ctx.beginPath(); rows.forEach((r, i) => (i ? ctx.lineTo : ctx.moveTo).call(ctx, x(r.t), y(r[key]))); };
    line("smoothed"); ctx.lineTo(x(rows[rows.length - 1].t), y(0)); ctx.lineTo(x(rows[0].t), y(0)); ctx.closePath();
    ctx.fillStyle = "rgba(82,43,18,.08)"; ctx.fill();
    line("count"); ctx.strokeStyle = "rgba(82,43,18,.35)"; ctx.lineWidth = 1.2; ctx.stroke();
    line("smoothed"); ctx.strokeStyle = "#522B12"; ctx.lineWidth = 2.2; ctx.lineJoin = "round"; ctx.stroke();
    const last = rows[rows.length - 1];
    ctx.beginPath(); ctx.arc(x(last.t), y(last.smoothed), 4, 0, Math.PI * 2);
    ctx.fillStyle = STATUS.levels[last.level].color; ctx.fill();
  }
  window.addEventListener("resize", () => { if (rows.length) drawChart(); });

  // ------------------------------------------------------------ Hồ cá Koi
  function koi(cls, body, head, spot) {
    return `<svg class="koi ${cls}" viewBox="0 0 92 40" aria-hidden="true">
      <g class="tail"><path d="M24 20 C 14 10, 6 4, 0 6 C 6 14, 6 26, 0 34 C 6 36, 14 30, 24 20 Z" fill="${head}" opacity=".85"/></g>
      <g class="fin"><path d="M50 13 C 44 4, 38 2, 34 4 C 40 8, 44 12, 48 15 Z" fill="#FFF3E2" opacity=".75"/>
        <path d="M50 27 C 44 36, 38 38, 34 36 C 40 32, 44 28, 48 25 Z" fill="#FFF3E2" opacity=".75"/></g>
      <path d="M20 20 C 30 9, 62 6, 80 12 C 92 16, 92 24, 80 28 C 62 34, 30 31, 20 20 Z" fill="${body}"/>
      <path d="M58 9 C 66 8, 76 9, 82 13 C 78 17, 68 16, 60 14 Z" fill="#FFF6E8"/>
      <path d="M38 22 C 46 20, 54 24, 56 30 C 48 31, 40 29, 36 26 Z" fill="#FFF6E8"/>
      <ellipse cx="46" cy="16" rx="7" ry="3.4" fill="${spot}"/>
      <circle cx="86" cy="17" r="1.3" fill="#2a1508"/><circle cx="86" cy="23" r="1.3" fill="#2a1508"/></svg>`;
  }
  $("#pond-stage").insertAdjacentHTML("beforeend",
    koi("k1", "#FF7A3D", "#FF4433", "#E2341F") + koi("k2", "#FFF6E8", "#FF7A3D", "#FF4433") +
    koi("k3", "#FF4433", "#E2341F", "#FF7A3D") + koi("k4", "#FF7A3D", "#FF4433", "#E2341F"));

  // ------------------------------------------------------------ Khởi động
  async function init() {
    try {
      STATUS = await (await fetch("/api/status")).json();
    } catch (_) { STATUS = { ok: false, levels: [], thresholds: {} }; }
    const badge = $("#status-badge");
    badge.classList.toggle("off", !STATUS.ok);
    $("#status-text").textContent = STATUS.ok ? "Hệ thống đang giám sát" : "Chưa nạp mô hình";
    if (STATUS.ok && STATUS.model) {
      const m = STATUS.model;
      $("#meta").innerHTML = [["Thiết bị", m.device.toUpperCase()], ["Epoch", m.epoch], ["MAE", m.val_mae ?? "–"]]
        .map(([k, v]) => `<span class="meta-link">${esc(k)} <b>${esc(v)}</b></span>`).join("");
    } else {
      const n = $("#model-missing");
      n.textContent = `Không tìm thấy trọng số ${STATUS.weights || "weights/csrnet_partB.pth"}. Hãy chạy notebook huấn luyện, tải csrnet_partB.pth và chép vào thư mục weights/, rồi khởi động lại server.`;
      n.hidden = false;
    }
    bindSettings();
    const start = location.hash.slice(1);
    go(start === "video" ? "video" : "image");
  }
  init();
})();
