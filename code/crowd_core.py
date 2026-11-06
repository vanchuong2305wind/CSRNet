"""Thành phần dùng chung cho notebook huấn luyện và app demo.

- CSRNet: mạng ước lượng bản đồ mật độ (Li et al., CVPR 2018).
- Tiền xử lý / suy luận / vẽ bản đồ nhiệt.
- CrowdMonitor: luật cảnh báo đám đông bất thường (quá tải mật độ + tăng đột biến).
"""
from collections import deque

import cv2
import numpy as np
import torch
import torch.nn as nn

IMAGENET_MEAN = np.array([0.485, 0.456, 0.406], dtype=np.float32)
IMAGENET_STD = np.array([0.229, 0.224, 0.225], dtype=np.float32)
DOWNSAMPLE = 8  # CSRNet xuất bản đồ mật độ ở 1/8 độ phân giải ảnh vào


# ----------------------------------------------------------------------------
# Mô hình
# ----------------------------------------------------------------------------
def _make_layers(cfg, in_channels=3, dilation=1):
    layers = []
    for v in cfg:
        if v == "M":
            layers.append(nn.MaxPool2d(kernel_size=2, stride=2))
        else:
            layers += [nn.Conv2d(in_channels, v, 3, padding=dilation, dilation=dilation),
                       nn.ReLU(inplace=True)]
            in_channels = v
    return nn.Sequential(*layers)


class CSRNet(nn.Module):
    FRONTEND = [64, 64, "M", 128, 128, "M", 256, 256, 256, "M", 512, 512, 512]
    BACKEND = [512, 512, 512, 256, 128, 64]

    def __init__(self, load_vgg=False):
        super().__init__()
        self.frontend = _make_layers(self.FRONTEND)                      # 10 lớp conv đầu của VGG-16
        self.backend = _make_layers(self.BACKEND, 512, dilation=2)       # 6 lớp conv giãn (dilation=2)
        self.output = nn.Conv2d(64, 1, kernel_size=1)
        for m in list(self.backend.modules()) + [self.output]:
            if isinstance(m, nn.Conv2d):
                nn.init.normal_(m.weight, std=0.01)
                nn.init.zeros_(m.bias)
        if load_vgg:
            import torchvision
            vgg = torchvision.models.vgg16(weights=torchvision.models.VGG16_Weights.IMAGENET1K_V1)
            self.frontend.load_state_dict(vgg.features[:23].state_dict())

    def forward(self, x):
        return self.output(self.backend(self.frontend(x)))


def load_model(weights_path, device="cpu"):
    """Đọc checkpoint do notebook xuất ra. Trả về (model, density_scale, meta)."""
    # weights_only=True: chỉ giải nén tensor/kiểu cơ bản, không thực thi mã pickle tùy ý trong file trọng số
    ckpt = torch.load(weights_path, map_location=device, weights_only=True)
    model = CSRNet(load_vgg=False)
    model.load_state_dict(ckpt["model"])
    model.to(device).eval()
    return model, float(ckpt.get("density_scale", 1.0)), {k: v for k, v in ckpt.items() if k != "model"}


# ----------------------------------------------------------------------------
# Tiền xử lý & suy luận
# ----------------------------------------------------------------------------
def resize_for_model(img_rgb, max_side=1024):
    """Thu nhỏ để cạnh dài <= max_side và hai chiều chia hết cho 8."""
    h, w = img_rgb.shape[:2]
    s = min(1.0, max_side / max(h, w))
    nh = max(DOWNSAMPLE, int(h * s) // DOWNSAMPLE * DOWNSAMPLE)
    nw = max(DOWNSAMPLE, int(w * s) // DOWNSAMPLE * DOWNSAMPLE)
    if (nh, nw) != (h, w):
        img_rgb = cv2.resize(img_rgb, (nw, nh), interpolation=cv2.INTER_AREA)
    return img_rgb


def to_tensor(img_rgb):
    x = (img_rgb.astype(np.float32) / 255.0 - IMAGENET_MEAN) / IMAGENET_STD
    return torch.from_numpy(x.transpose(2, 0, 1)).unsqueeze(0)


@torch.no_grad()
def predict(model, img_rgb, density_scale=1.0, device="cpu", max_side=1024):
    """Trả về (số người ước lượng, bản đồ mật độ 1/8, ảnh đã resize)."""
    img = resize_for_model(img_rgb, max_side)
    den = model(to_tensor(img).to(device))[0, 0].float().cpu().numpy() / density_scale
    den = np.clip(den, 0, None)
    return float(den.sum()), den, img


def density_overlay(img_rgb, den, alpha=0.5):
    h, w = img_rgb.shape[:2]
    d = cv2.resize(den, (w, h), interpolation=cv2.INTER_CUBIC)
    d = np.clip(d / (d.max() + 1e-8), 0, 1)
    heat = cv2.applyColorMap((d * 255).astype(np.uint8), cv2.COLORMAP_JET)
    heat = cv2.cvtColor(heat, cv2.COLOR_BGR2RGB)
    return cv2.addWeighted(img_rgb, 1 - alpha, heat, alpha, 0)


def privacy_blur(img_rgb, ksize=31):
    """Làm mờ ảnh hiển thị: người xem chỉ thấy phân bố đám đông, không nhận diện được cá nhân."""
    return cv2.GaussianBlur(img_rgb, (ksize, ksize), 0)


# ----------------------------------------------------------------------------
# Luật cảnh báo bất thường
# ----------------------------------------------------------------------------
LEVELS = [
    {"name": "XANH", "label": "Bình thường", "color": "#2e7d32", "action": "Giám sát thông thường"},
    {"name": "VÀNG", "label": "Chú ý", "color": "#f9a825", "action": "Theo dõi sát, chuẩn bị phương án phân luồng"},
    {"name": "CAM", "label": "Cảnh báo", "color": "#ef6c00", "action": "Hạn chế người vào, mở thêm lối thoát, điều lực lượng"},
    {"name": "ĐỎ", "label": "Nguy hiểm", "color": "#c62828", "action": "Dừng dòng người vào, giải tỏa khẩn cấp, thông báo loa"},
]
# Ngưỡng mật độ (người/m²) chuyển sang mức Vàng / Cam / Đỏ — tổng hợp từ Fruin và Still.
THRESHOLDS = {"di chuyển": (2.0, 3.0, 4.0), "đứng": (3.0, 4.0, 5.0)}


def density_level(rho, crowd_type="di chuyển"):
    return int(sum(rho >= t for t in THRESHOLDS[crowd_type]))


class CrowdMonitor:
    """Theo dõi chuỗi số đếm theo thời gian và phát cảnh báo.

    - Quá tải: mật độ rho = C / A vượt ngưỡng; mức chỉ đổi khi giữ nguyên `min_persist` lần đo liên tiếp.
    - Tăng đột biến: (C_t - C_{t-W}) / C_{t-W} >= surge_ratio với C_{t-W} >= min_count; khi đó nâng mức thêm một bậc.
    """

    def __init__(self, area_m2, crowd_type="di chuyển", surge_ratio=0.5, surge_window=10.0,
                 min_count=10, min_persist=3, ema_alpha=0.5):
        self.area = float(area_m2)
        self.crowd_type = crowd_type
        self.surge_ratio = surge_ratio
        self.surge_window = surge_window
        self.min_count = min_count
        self.min_persist = min_persist
        self.ema_alpha = ema_alpha
        self.history = deque()  # (t, số đếm đã làm trơn)
        self.smoothed = None
        self.level = 0
        self._cand, self._n = 0, 0
        self.surge = False
        self.final_level = 0

    def update(self, t, count):
        a = self.ema_alpha
        self.smoothed = count if self.smoothed is None else a * count + (1 - a) * self.smoothed
        rho = self.smoothed / self.area
        events = []

        raw = density_level(rho, self.crowd_type)
        if raw == self.level:
            self._n = 0
        else:
            self._n = self._n + 1 if raw == self._cand else 1
            self._cand = raw
            if self._n >= self.min_persist:
                if raw > self.level:
                    events.append(f"Mật độ tăng lên mức {LEVELS[raw]['name']} ({rho:.2f} người/m²)")
                else:
                    events.append(f"Mật độ giảm về mức {LEVELS[raw]['name']} ({rho:.2f} người/m²)")
                self.level, self._n = raw, 0

        while self.history and self.history[0][0] < t - self.surge_window:
            self.history.popleft()
        base = self.history[0][1] if self.history else None
        surge = (base is not None and base >= self.min_count
                 and (self.smoothed - base) / base >= self.surge_ratio)
        if surge and not self.surge:
            events.append(f"Số người tăng đột biến {base:.0f} → {self.smoothed:.0f} trong {self.surge_window:.0f}s")
        self.surge = surge
        self.history.append((t, self.smoothed))

        final = min(3, self.level + (1 if surge else 0))
        if final != self.final_level and not events:
            events.append(f"Mức cảnh báo: {LEVELS[final]['name']}")
        self.final_level = final
        return {"t": t, "count": count, "smoothed": self.smoothed, "density": rho,
                "level": final, "surge": surge, "events": events}
