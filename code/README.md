# Đếm người & cảnh báo đám đông bất thường (CSRNet + luật mật độ)

```
code/
├── train_csrnet_shanghaitech.ipynb  # B1–B3: tải dữ liệu, tiền xử lý, huấn luyện, đánh giá, xuất trọng số
├── crowd_core.py                    # mô hình CSRNet, suy luận, luật cảnh báo (CrowdMonitor)
├── server.py                        # B4: web app Flask (ảnh, video, webcam/RTSP) + API
├── templates/index.html             # giao diện
├── static/css, static/js            # style.css, app.js (không dùng framework)
├── static/assets/                   # nền giấy, hạt nhiễu, ảnh minh họa
├── app.py                           # bản demo Streamlit cũ (không bắt buộc)
├── requirements.txt
└── weights/                         # chép csrnet_partB.pth vào đây
```

## B1–B3. Huấn luyện (Google Colab, GPU T4)
1. Mở `train_csrnet_shanghaitech.ipynb` trên Colab → Runtime → Change runtime type → **T4 GPU**.
2. Run all. Notebook tự tải ShanghaiTech qua `kagglehub`; nếu lỗi, tải thủ công rồi gán `DATA_ROOT`.
3. Mặc định 150 epoch (~1 giờ trên T4). Muốn chạy nhanh để thử: sửa `epochs` trong `CFG`.
4. Cuối notebook tải về `outputs.zip`, gồm:
   - `csrnet_partB.pth`: trọng số mô hình tốt nhất
   - `results.json`: MAE, RMSE, tốc độ (điền vào Bảng 4.3 của báo cáo)
   - `hist_counts.png`, `density_gt.png`, `training_curves.png`, `scatter_test.png`, `test_samples.png`, `alert_levels.png`: hình cho báo cáo

## B4. Chạy web app (máy cá nhân, CPU cũng chạy được)
```bash
pip install -r requirements.txt
# chép outputs/csrnet_partB.pth vào code/weights/
python server.py          # mở http://127.0.0.1:5000
```
- **Ảnh:** thả/chọn ảnh → số người, mật độ người/m², bản đồ nhiệt, mức cảnh báo.
- **Video / Camera:** phân tích video tải lên hoặc webcam/RTSP theo thời gian thực (Server-Sent Events):
  khung hình, biểu đồ số người kèm ngưỡng cảnh báo, nhật ký sự kiện, tải CSV.
- **Cấu hình** (nút trên thanh điều hướng nổi): diện tích vùng quan sát, loại đám đông, chế độ riêng tư,
  độ phân giải, ngưỡng & cửa sổ tăng đột biến, số lần đo xác nhận. Được lưu trong trình duyệt.
- Biến môi trường: `CROWD_WEIGHTS`, `CROWD_HOST` (mặc định 127.0.0.1), `CROWD_PORT` (5000),
  `CONTACT_EMAIL` (hiện nút Email trên thanh liên hệ, bấm để sao chép).
