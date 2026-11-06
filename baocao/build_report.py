"""Dựng lại báo cáo theo một pipeline duy nhất (CSRNet + luật cảnh báo mật độ + app demo).

Cách dùng:
    python build_report.py                  # chưa có kết quả huấn luyện: để trống (…) các ô số liệu/hình kết quả
    python build_report.py ../output        # sau khi chạy notebook: tự điền results.json và chèn hình .png
                                            # (mặc định tự tìm ../output rồi ../code/outputs)

Giữ nguyên trang bìa, mục lục (trường TOC) và định dạng của file gốc; phần thân được viết lại.
"""
import copy
import json
import os
import re
import sys
from xml.sax.saxutils import escape

import docx
from docx.oxml import parse_xml
from docx.shared import Cm

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "BaoCao_He_thong_dem_nguoi_phat_hien_dam_dong_bat_thuong.docx")
DST = os.path.join(HERE, "BaoCao_He_thong_dem_nguoi_phat_hien_dam_dong_bat_thuong_final.docx")
IMG = os.path.join(HERE, "hinh_anh")
OUTPUTS = sys.argv[1] if len(sys.argv) > 1 else next(
    (d for d in (os.path.join(HERE, "..", "output"), os.path.join(HERE, "..", "code", "outputs")) if os.path.isdir(d)),
    os.path.join(HERE, "..", "output"))
RES, HIST = {}, []
if os.path.exists(os.path.join(OUTPUTS, "history.json")):
    HIST = [h for h in json.load(open(os.path.join(OUTPUTS, "history.json"), encoding="utf-8")) if "val_mae" in h]
if os.path.exists(os.path.join(OUTPUTS, "results.json")):
    RES = json.load(open(os.path.join(OUTPUTS, "results.json"), encoding="utf-8"))

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"

doc = docx.Document(SRC)
body = doc.element.body
orig = list(body)

# ---------------------------------------------------------------- ảnh có sẵn trong file gốc
EXISTING = {}
for e in orig:
    for dp in e.iter(WP + "docPr"):
        EXISTING[dp.get("name")] = e

# ---------------------------------------------------------------- trang bìa
for p in doc.paragraphs[:12]:
    for r in p.runs:
        r.text = r.text.replace("BÁO CÁO TIẾN ĐỘ ĐỀ TÀI", "BÁO CÁO ĐỀ TÀI").replace("Tháng 10/2026", "Tháng 11/2026")

# giữ bìa (0–11), tiêu đề MỤC LỤC (12) và trường TOC (13); xóa phần còn lại trừ sectPr cuối
for e in orig[14:]:
    if not e.tag.endswith("sectPr"):
        body.remove(e)
SECT = body[-1]

# danh sách đánh số mới (bắt đầu lại từ 1) dựa trên abstractNum dạng số "%1." của file gốc
numbering = doc.part.numbering_part.element
_next_num = [100]


def new_numbered_list():
    _next_num[0] += 1
    n = _next_num[0]
    numbering.append(parse_xml(
        f'<w:num {W_NS} w:numId="{n}"><w:abstractNumId w:val="2"/>'
        f'<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>'))
    return n


# ---------------------------------------------------------------- tạo phần tử
CONTENT = []        # các phần tử thân báo cáo theo thứ tự
FIGS, TABS = [], []  # cho danh mục hình / bảng
chap = [0]
fig_no, tab_no = [0], [0]


def runs(text, rpr=""):
    """**đậm** và _nghiêng_ trong chuỗi -> các run."""
    out = []
    for part in re.split(r"(\*\*.+?\*\*|__.+?__)", text):
        if not part:
            continue
        extra = ""
        if part.startswith("**"):
            part, extra = part[2:-2], "<w:b/><w:bCs/>"
        elif part.startswith("__"):
            part, extra = part[2:-2], "<w:i/><w:iCs/>"
        out.append(f'<w:r><w:rPr>{extra}{rpr}</w:rPr><w:t xml:space="preserve">{escape(part)}</w:t></w:r>')
    return "".join(out)


def add(xml):
    el = parse_xml(xml)
    CONTENT.append(el)
    return el


def H1(text, numbered=True):
    if numbered:
        chap[0] += 1
        fig_no[0] = tab_no[0] = 0
    add(f'<w:p {W_NS}><w:pPr><w:pStyle w:val="Heading1"/><w:pageBreakBefore/></w:pPr>{runs(text)}</w:p>')


def H2(text):
    add(f'<w:p {W_NS}><w:pPr><w:pStyle w:val="Heading2"/></w:pPr>{runs(text)}</w:p>')


def H3(text):
    add(f'<w:p {W_NS}><w:pPr><w:pStyle w:val="Heading3"/></w:pPr>{runs(text)}</w:p>')


def P(text):
    add(f'<w:p {W_NS}><w:pPr><w:spacing w:after="120" w:line="312" w:lineRule="auto"/>'
        f'<w:ind w:firstLine="567"/><w:jc w:val="both"/></w:pPr>{runs(text)}</w:p>')


def BUL(items, num_id=2):
    for t in items:
        add(f'<w:p {W_NS}><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/>'
            f'<w:numId w:val="{num_id}"/></w:numPr><w:spacing w:after="80" w:line="312" w:lineRule="auto"/>'
            f'<w:jc w:val="both"/></w:pPr>{runs(t)}</w:p>')


def NUM(items):
    BUL(items, new_numbered_list())


def FORMULA(text):
    add(f'<w:p {W_NS}><w:pPr><w:spacing w:before="60" w:after="120" w:line="300" w:lineRule="auto"/>'
        f'<w:jc w:val="center"/></w:pPr>{runs(text, "<w:i/><w:iCs/>")}</w:p>')


def CODE(lines):
    for i, line in enumerate(lines):
        after = "120" if i == len(lines) - 1 else "0"
        add(f'<w:p {W_NS}><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>'
            f'<w:spacing w:after="{after}" w:line="240" w:lineRule="auto"/><w:ind w:left="567"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="20"/>'
            f'<w:szCs w:val="20"/></w:rPr><w:t xml:space="preserve">{escape(line)}</w:t></w:r></w:p>')


def _caption(kind, text):
    if kind == "Hình":
        fig_no[0] += 1
        label = f"Hình {chap[0]}.{fig_no[0]}."
        FIGS.append(f"{label} {text}")
        spacing = '<w:keepNext/><w:spacing w:after="20" w:line="276" w:lineRule="auto"/>'
    else:
        tab_no[0] += 1
        label = f"Bảng {chap[0]}.{tab_no[0]}."
        TABS.append(f"{label} {text}")
        spacing = '<w:keepNext/><w:spacing w:before="160" w:after="80"/>'
    add(f'<w:p {W_NS}><w:pPr>{spacing}<w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">{label} </w:t></w:r>'
        f'<w:r><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>')
    return label


def NOTE(text, center=True, after="200"):
    jc = '<w:jc w:val="center"/>' if center else ""
    add(f'<w:p {W_NS}><w:pPr><w:spacing w:before="0" w:after="{after}" w:line="276" w:lineRule="auto"/>{jc}</w:pPr>'
        f'<w:r><w:rPr><w:i/><w:iCs/><w:color w:val="444444"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
        f'<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>')


def _picture_paragraph(path, width_cm):
    p = doc.add_paragraph()
    p.add_run().add_picture(path, width=Cm(width_cm))
    el = p._p
    body.remove(el)
    el.insert(0, parse_xml(f'<w:pPr {W_NS}><w:keepNext/><w:spacing w:before="120" w:after="60"/><w:jc w:val="center"/></w:pPr>'))
    return el


def FIG(caption, source=None, existing=None, path=None, width=14.0, placeholder=None):
    """Hình: lấy lại ảnh có sẵn trong file gốc, chèn ảnh mới, hoặc để khung chờ nếu chưa có file."""
    if existing:
        CONTENT.append(copy.deepcopy(EXISTING[existing]))
    elif path and os.path.exists(path):
        CONTENT.append(_picture_paragraph(path, width))
    else:
        add(f'<w:p {W_NS}><w:pPr><w:keepNext/><w:pBdr>'
            + "".join(f'<w:{s} w:val="dashed" w:sz="6" w:space="8" w:color="999999"/>' for s in ("top", "left", "bottom", "right"))
            + '</w:pBdr><w:spacing w:before="200" w:after="200"/><w:ind w:left="1134" w:right="1134"/><w:jc w:val="center"/></w:pPr>'
            f'<w:r><w:rPr><w:color w:val="808080"/></w:rPr><w:t xml:space="preserve">{escape(placeholder)}</w:t></w:r></w:p>')
    label = _caption("Hình", caption)
    if source:
        NOTE(source)
    else:
        NOTE("", after="160")
    return label


def _cell(text, width, header=False, center=False):
    fill = '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/>' if header else ""
    jc = '<w:jc w:val="center"/>' if (header or center) else ""
    rpr = ("<w:b/><w:bCs/>" if header else "") + '<w:sz w:val="23"/><w:szCs w:val="23"/>'
    borders = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="808080"/>' for s in ("top", "left", "bottom", "right"))
    paras = "".join(
        f'<w:p><w:pPr><w:spacing w:after="0" w:line="264" w:lineRule="auto"/>{jc}</w:pPr>'
        f'{runs(line, rpr)}</w:p>' for line in str(text).split("\n"))
    return (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/><w:tcBorders>{borders}</w:tcBorders>{fill}'
            f'<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/>'
            f'<w:right w:w="100" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/></w:tcPr>{paras}</w:tc>')


def TABLE(caption, header, rows, widths, note=None, center_cols=()):
    total = 9071
    ws = [int(total * w / sum(widths)) for w in widths]
    if caption:
        _caption("Bảng", caption)
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in ws)
    trs = '<w:tr><w:trPr><w:tblHeader/></w:trPr>' + "".join(_cell(h, w, header=True) for h, w in zip(header, ws)) + "</w:tr>"
    for r in rows:
        trs += "<w:tr><w:trPr><w:cantSplit/></w:trPr>" + "".join(
            _cell(c, w, center=i in center_cols) for i, (c, w) in enumerate(zip(r, ws))) + "</w:tr>"
    add(f'<w:tbl {W_NS}><w:tblPr><w:tblW w:w="{total}" w:type="dxa"/><w:jc w:val="center"/>'
        f'<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        f'</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{trs}</w:tbl>')
    if note:
        NOTE(note, center=False)
    else:
        NOTE("", after="120")


def R(key, fmt="{:.2f}", default="…"):
    v = RES
    for k in key.split("."):
        if isinstance(v, list) and k.isdigit() and int(k) < len(v):
            v = v[int(k)]
        elif isinstance(v, dict) and k in v:
            v = v[k]
        else:
            return default
    return vn(fmt.format(v)) if isinstance(v, (int, float)) else str(v)


def vn(s):
    """Định dạng số kiểu Việt Nam: 1,234.5 -> 1.234,5"""
    return s.replace(",", "\0").replace(".", ",").replace("\0", ".")


def OUT(name):
    return os.path.join(OUTPUTS, name)


def PENDING(name):
    return f"[Hình kết quả – tự chèn từ code/outputs/{name} sau khi chạy notebook rồi chạy lại build_report.py]"


# ---------------------------------------------------------------- hình sơ đồ pipeline
def draw_pipeline(path):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch

    fig, ax = plt.subplots(figsize=(13, 5.6))
    ax.set_xlim(0, 13); ax.set_ylim(0, 5.6); ax.axis("off")

    def box(x, y, w, h, title, sub, color):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.12",
                                    fc=color, ec="#37474f", lw=1.2))
        ax.text(x + w / 2, y + h * 0.66, title, ha="center", va="center", fontsize=10.5, weight="bold")
        ax.text(x + w / 2, y + h * 0.3, sub, ha="center", va="center", fontsize=8.6, color="#263238")

    def arrow(x1, y1, x2, y2):
        ax.annotate("", (x2, y2), (x1, y1), arrowprops=dict(arrowstyle="-|>", lw=1.5, color="#37474f"))

    ax.text(0.1, 5.35, "A. HUẤN LUYỆN (offline – notebook Colab)", fontsize=11, weight="bold", color="#1565c0")
    ax.text(0.1, 2.55, "B. VẬN HÀNH (online – ứng dụng demo)", fontsize=11, weight="bold", color="#2e7d32")
    blue, green, orange = "#e3f2fd", "#e8f5e9", "#fff3e0"
    w, h = 2.25, 1.35
    xs = [0.1, 2.75, 5.4, 8.05, 10.7]
    top = [("1. Dữ liệu", "ShanghaiTech Part B\n716 ảnh, nhãn điểm"),
           ("2. Tiền xử lý", "Bản đồ mật độ GT\nGauss σ=15, 1/8"),
           ("3. Huấn luyện", "CSRNet (VGG-16 +\ntích chập giãn)"),
           ("4. Đánh giá", "MAE, RMSE, FPS\ntrên 316 ảnh test"),
           ("Trọng số", "csrnet_partB.pth")]
    bot = [("Camera / video", "khung hình RTSP,\nvideo, ảnh"),
           ("CSRNet", "bản đồ mật độ D\nC = Σ D(x)"),
           ("Quy đổi mật độ", "ρ = C / A\n(người/m²)"),
           ("5. Luật cảnh báo", "quá tải + tăng\nđột biến, 4 mức"),
           ("6. Giao diện", "bản đồ nhiệt, mức,\nnhật ký, CSV")]
    for i, (t, s) in enumerate(top):
        box(xs[i], 3.65, w, h, t, s, orange if i == 4 else blue)
        if i:
            arrow(xs[i - 1] + w, 3.65 + h / 2, xs[i], 3.65 + h / 2)
    for i, (t, s) in enumerate(bot):
        box(xs[i], 0.85, w, h, t, s, green)
        if i:
            arrow(xs[i - 1] + w, 0.85 + h / 2, xs[i], 0.85 + h / 2)
    arrow(xs[4] + w / 2, 3.65, xs[1] + w / 2, 0.85 + h)
    ax.text(7.2, 3.05, "nạp trọng số", fontsize=9, style="italic", color="#e65100")
    ax.text(6.5, 0.25, "An toàn thông tin xuyên suốt: xử lý trong bộ nhớ, làm mờ ảnh hiển thị, chỉ lưu số liệu, "
            "nạp trọng số an toàn (6.)", ha="center", fontsize=9.5, color="#b71c1c")
    plt.tight_layout()
    plt.savefig(path, dpi=200)
    plt.close(fig)


PIPE = os.path.join(IMG, "pipeline.png")
draw_pipeline(PIPE)

# ================================================================ NỘI DUNG
H1("LỜI MỞ ĐẦU", numbered=False)
P("Những nơi công cộng đông người như lễ hội, sân vận động, nhà ga, phố đi bộ hay các sự kiện đón năm mới luôn tiềm ẩn nguy cơ mất an toàn. Khi mật độ người vượt quá sức chứa của không gian, một sự xô đẩy nhỏ có thể biến thành thảm họa chèn ép đám đông (crowd crush), như các vụ tại Itaewon (Hàn Quốc, 2022), Love Parade (Đức, 2010) hay Mina (Ả Rập Xê Út, 2015). Điểm chung của các thảm họa này là dấu hiệu nguy hiểm đã xuất hiện từ trước — mật độ tăng dần, dòng người dồn ứ — nhưng lực lượng quản lý không nhận ra hoặc không phản ứng kịp.")
P("Song song đó, hệ thống camera giám sát (CCTV) đã phủ rộng khắp các đô thị, nhưng phần lớn vẫn được theo dõi thủ công. Một người vận hành không thể quan sát liên tục hàng chục màn hình. Thị giác máy tính và học sâu cho phép **tự động đếm người và ước lượng mật độ đám đông** từ chính các camera này, từ đó **cảnh báo sớm** khi đám đông trở nên bất thường.")
P("Báo cáo trình bày toàn bộ quá trình thực hiện đề tài theo **một pipeline thống nhất**: (1) phân tích bài toán và chọn hướng tiếp cận ước lượng bản đồ mật độ; (2) chuẩn bị bộ dữ liệu ShanghaiTech Part B và sinh bản đồ mật độ; (3) huấn luyện và đánh giá mạng CSRNet; (4) xây dựng mô-đun cảnh báo đám đông bất thường dựa trên mật độ người/m² và tốc độ tăng số người; (5) tích hợp thành ứng dụng demo; (6) phân tích và áp dụng các biện pháp an toàn thông tin, bảo vệ quyền riêng tư.")
P("Mã nguồn gồm notebook huấn luyện (train_csrnet_shanghaitech.ipynb), mô-đun lõi (crowd_core.py) và ứng dụng demo (app.py). Hình minh họa lý thuyết được trích nguyên gốc từ các bài báo khoa học đã công bố hoặc từ Wikimedia Commons theo giấy phép Creative Commons, nguồn ghi dưới mỗi hình; hình kết quả thực nghiệm do nhóm tạo ra từ notebook. Báo cáo không sử dụng hình ảnh do AI tạo ra.")

# ---------------------------------------------------------------- CHƯƠNG 1
H1("CHƯƠNG 1. TỔNG QUAN ĐỀ TÀI")
H2("1.1. Bối cảnh và lý do chọn đề tài")
P("Thảm họa đám đông không phải là sự kiện hiếm. Một số vụ tiêu biểu trong những năm gần đây:")
BUL(["**Itaewon, Seoul (29/10/2022):** khoảng 22 giờ 20 trong dịp lễ Halloween, dòng người dồn vào một con hẻm dốc, hẹp gây chèn ép, làm **159 người tử vong và 196 người bị thương**. Cuộc điều tra của cảnh sát kết luận nguyên nhân chính là sự chuẩn bị không đầy đủ của cảnh sát và chính quyền dù đã có nhiều cảnh báo từ trước [1].",
     "**Love Parade, Duisburg (24/7/2010):** 21 người chết do ngạt thở khi cố thoát khỏi một đoạn dốc dẫn vào khu lễ hội, 652 người bị thương [2].",
     "**Mina, Mecca (24/9/2015):** vụ chèn ép trong mùa hành hương Hajj; số liệu chính thức của Ả Rập Xê Út là 769 người chết, trong khi các hãng tin AP và AFP ước tính trên 2.200 người [3]."])
FIG("Đám đông dày đặc trên một tuyến phố ở Itaewon (Seoul) tối 29/10/2022 — đêm xảy ra thảm họa chèn ép.",
    "Nguồn: Watchers Club, “Itaewon (이태원) 2022 Halloween.jpg”, Wikimedia Commons, giấy phép CC BY 3.0.", existing="itaewon_crowd")
FIG("Con hẻm nơi xảy ra vụ chèn ép đám đông tại Itaewon, ảnh chụp ngày 01/11/2022.",
    "Nguồn: *Youngjin, “Itaewon Halloween crowd crush incident place on November 1, 2022.jpg”, Wikimedia Commons, giấy phép CC BY-SA 3.0.", existing="itaewon_alley")
P("Các nghiên cứu về khoa học đám đông (crowd science) chỉ ra rằng phần lớn nạn nhân tử vong do ngạt vì bị ép (compressive asphyxia) chứ không phải do “hoảng loạn giẫm đạp” như cách hiểu phổ biến [4]. Nghĩa là thảm họa có thể phòng tránh **nếu đo được mật độ và phát hiện sớm tình trạng quá tải** để điều tiết dòng người. Đây chính là động lực và cũng là định hướng kỹ thuật của đề tài.")
FIG("Đám đông hành hương quanh Kaaba (Mecca) trong mùa Hajj năm 2018 — ví dụ về đám đông có mật độ cực cao mà bài toán đếm người phải xử lý.",
    "Nguồn: Adli Wahid (chỉnh sửa nhỏ: Basile Morin), “The Kaaba during Hajj - edited.jpg”, Wikimedia Commons, giấy phép CC BY-SA 4.0.", existing="hajj")

H2("1.2. Phát biểu bài toán")
P("Đề tài giải quyết một bài toán thống nhất: **từ khung hình camera giám sát, ước lượng số người và mật độ đám đông, rồi phát hiện tình trạng đám đông bất thường có nguy cơ gây mất an toàn.** Cụ thể:")
BUL(["**Đầu vào:** ảnh hoặc khung hình video từ camera giám sát cố định, cùng diện tích thực tế A (m²) của vùng camera quan sát.",
     "**Bước 1 – Đếm người bằng bản đồ mật độ:** mô hình học sâu sinh bản đồ mật độ D(x); số người C bằng tổng các giá trị trên bản đồ.",
     "**Bước 2 – Quy đổi mật độ:** ρ = C / A (người/m²).",
     "**Bước 3 – Phát hiện bất thường:** đám đông được xem là bất thường khi (i) mật độ vượt ngưỡng an toàn — **quá tải**, hoặc (ii) số người tăng nhanh bất thường trong thời gian ngắn — **tăng đột biến/dồn ứ**.",
     "**Đầu ra:** số người, bản đồ nhiệt mật độ, mức cảnh báo bốn cấp (Xanh – Vàng – Cam – Đỏ) kèm khuyến nghị hành động và nhật ký sự kiện."])
P("Định nghĩa “bất thường” theo mật độ bám sát nguyên nhân trực tiếp của các thảm họa chèn ép ở mục 1.1, có thể giải thích rõ ràng cho người vận hành và dùng trực tiếp đầu ra của mô hình đếm người. Các dạng bất thường về hành vi (ẩu đả, tháo chạy, xe lao vào đám đông) cần mô hình phân tích video riêng và được xếp vào hướng phát triển (Chương 7).")

H2("1.3. Mục tiêu và phạm vi")
BUL(["Nắm được cơ sở lý thuyết về mật độ đám đông, ngưỡng an toàn và phương pháp đếm người bằng bản đồ mật độ.",
     "Chuẩn bị dữ liệu: tải ShanghaiTech Part B, sinh bản đồ mật độ ground truth, chia tập huấn luyện/kiểm định/kiểm thử.",
     "Cài đặt, huấn luyện và đánh giá mạng CSRNet bằng PyTorch theo độ đo MAE, RMSE và tốc độ xử lý.",
     "Xây dựng mô-đun cảnh báo đám đông bất thường dựa trên luật mật độ và tăng đột biến, có cơ chế chống cảnh báo chập chờn.",
     "Xây dựng ứng dụng demo xử lý ảnh, video và luồng camera.",
     "Phân tích rủi ro an toàn thông tin của hệ thống và áp dụng các biện pháp bảo vệ quyền riêng tư ngay trong ứng dụng demo."])
P("**Phạm vi:** hệ thống chỉ phân tích ở mức đám đông, **không nhận dạng danh tính** (không dùng nhận dạng khuôn mặt) — lựa chọn có chủ đích để giảm rủi ro quyền riêng tư (Chương 6). Thực nghiệm dùng bộ dữ liệu công khai; việc lắp đặt và hiệu chỉnh tại hiện trường thực tế nằm ngoài phạm vi đề tài.")

H2("1.4. Pipeline thực hiện đề tài")
P("Toàn bộ đề tài được tổ chức thành một pipeline gồm hai nhánh dùng chung một mô hình (Hình 1.4): nhánh **huấn luyện** chạy một lần trên Google Colab để tạo trọng số CSRNet; nhánh **vận hành** dùng trọng số đó để xử lý khung hình, quy đổi mật độ, áp luật cảnh báo và hiển thị cho người vận hành.")
FIG("Pipeline của đề tài: nhánh huấn luyện (A) tạo trọng số CSRNet, nhánh vận hành (B) dùng trọng số để đếm người và cảnh báo.",
    "Nguồn: nhóm thực hiện.", path=PIPE, width=16.0)
TABLE("Các bước của pipeline và sản phẩm tương ứng",
      ["Bước", "Nội dung", "Sản phẩm", "Chương"],
      [["1", "Phân tích bài toán, chọn phương pháp", "Hướng tiếp cận: bản đồ mật độ + luật mật độ", "2"],
       ["2", "Dữ liệu và tiền xử lý", "Bản đồ mật độ .npy; notebook mục 2–5", "3"],
       ["3", "Huấn luyện và đánh giá CSRNet", "csrnet_partB.pth, results.json; notebook mục 6–8", "4"],
       ["4", "Mô-đun cảnh báo bất thường", "Lớp CrowdMonitor trong crowd_core.py", "5"],
       ["5", "Ứng dụng demo", "app.py (Streamlit)", "5"],
       ["6", "An toàn thông tin, quyền riêng tư", "Biện pháp trong demo và đề xuất khi triển khai", "6"]],
      [1, 4, 5, 1.2], center_cols=(0, 3))

H2("1.5. Ý nghĩa thực tiễn")
P("Hệ thống có thể áp dụng cho quản lý sự kiện lớn (lễ hội, bắn pháo hoa, đón năm mới, ca nhạc ngoài trời), nhà ga – bến xe – sân bay, sân vận động, trung tâm thương mại và khu du lịch. Lợi ích chính: (i) cung cấp số liệu khách quan về sức chứa theo thời gian thực; (ii) giảm tải cho người giám sát camera; (iii) cảnh báo sớm để điều tiết, phân luồng trước khi mật độ đạt mức nguy hiểm; (iv) lưu lại số liệu thống kê phục vụ quy hoạch và rút kinh nghiệm sau sự kiện.")

# ---------------------------------------------------------------- CHƯƠNG 2
H1("CHƯƠNG 2. CƠ SỞ LÝ THUYẾT")
H2("2.1. Mật độ đám đông và ngưỡng an toàn")
P("Mật độ đám đông được đo bằng số người trên một mét vuông (người/m²). Fruin (1971) đưa ra khái niệm “mức phục vụ” (Level of Service) cho không gian đi bộ [5]; G. Keith Still phát triển thêm trong lĩnh vực khoa học đám đông [4]. Theo các tài liệu này, đám đông di chuyển bắt đầu bị cản trở khi mật độ khoảng 2–3 người/m² và nguy cơ tăng mạnh khi vượt khoảng 4 người/m²; đối với đám đông đứng yên, khoảng 5 người/m² thường được xem là giới hạn trên, từ 6 người/m² trở lên nguy cơ chèn ép nguy hiểm tăng rõ rệt. Các con số mang tính tham khảo, phụ thuộc địa hình (dốc, cầu thang, lối thoát) và cần hiệu chỉnh cho từng địa điểm.")
P("Từ đó, đề tài xây dựng bốn mức cảnh báo (Bảng 2.1). Đây là cơ sở của luật phát hiện bất thường ở Chương 5 và là lý do hệ thống phải **chuyển số đếm thành mật độ trên diện tích thực** chứ không chỉ đếm.")
TABLE("Mức cảnh báo theo mật độ (tham khảo, cần hiệu chỉnh theo địa điểm)",
      ["Mức", "Đám đông di chuyển", "Đám đông đứng", "Hành động đề xuất"],
      [["Xanh – Bình thường", "dưới 2 người/m²", "dưới 3 người/m²", "Giám sát thông thường"],
       ["Vàng – Chú ý", "2 – 3 người/m²", "3 – 4 người/m²", "Theo dõi sát, chuẩn bị phương án phân luồng"],
       ["Cam – Cảnh báo", "3 – 4 người/m²", "4 – 5 người/m²", "Hạn chế người vào, mở thêm lối thoát, điều lực lượng"],
       ["Đỏ – Nguy hiểm", "trên 4 người/m²", "trên 5 người/m²", "Dừng dòng người vào, giải tỏa khẩn cấp, thông báo loa"]],
      [2.2, 2, 2, 3.6], note="Ngưỡng do nhóm tổng hợp từ Fruin [5] và Still [4]; được cài đặt trong crowd_core.py (biến THRESHOLDS).")

H2("2.2. Các hướng tiếp cận bài toán đếm người và lựa chọn của đề tài")
P("Theo khảo sát của Gao và cộng sự [8], các phương pháp đếm người phát triển qua ba nhóm: dựa trên phát hiện (detection-based), dựa trên hồi quy (regression-based) và dựa trên ước lượng bản đồ mật độ (density estimation-based); từ năm 2015–2016, mạng nơ-ron tích chập (CNN) trở thành hướng chủ đạo (Hình 2.1).")
FIG("Lịch sử phát triển của bài toán đếm người: từ phương pháp dựa trên phát hiện, hồi quy đến ước lượng mật độ bằng mạng học sâu.",
    "Nguồn: G. Gao, J. Gao, Q. Liu, Q. Wang, Y. Wang, “CNN-based Density Estimation and Crowd Counting: A Survey”, arXiv:2003.12783, 2020, Fig. 1 [8].", existing="survey_fig1")
BUL(["**Dựa trên phát hiện:** dùng bộ phát hiện đối tượng (ví dụ họ YOLO) khoanh từng người rồi đếm. Chính xác khi đám đông thưa nhưng bỏ sót nhiều khi người chỉ còn vài pixel và che khuất lẫn nhau.",
     "**Dựa trên hồi quy:** ánh xạ đặc trưng toàn ảnh thành một con số. Đơn giản nhưng không cho biết người tập trung ở đâu.",
     "**Dựa trên bản đồ mật độ:** mạng dự đoán một bản đồ mà tổng giá trị là số người. Vừa đếm được, vừa cho biết phân bố không gian; hoạt động tốt với đám đông dày."])
P("Đề tài chọn **hướng bản đồ mật độ** vì: (i) camera giám sát thường đặt cao, người ở xa nhỏ và bị che khuất — đúng điểm yếu của phương pháp phát hiện; (ii) bản đồ mật độ cho biết các “điểm nóng” cục bộ, cần thiết cho cảnh báo; (iii) chỉ cần nhãn điểm đầu người, có sẵn trong các bộ dữ liệu chuẩn.")

H2("2.3. Bản đồ mật độ")
P("Nhãn huấn luyện được tạo từ chú thích điểm đầu người: mỗi đầu người tại vị trí xᵢ được “làm mờ” bằng một nhân Gauss chuẩn hóa G_σ (tổng bằng 1):")
FORMULA("F(x) = Σᵢ δ(x − xᵢ) ∗ G_σᵢ(x)")
P("Vì mỗi nhân có tổng bằng 1 nên tổng toàn bản đồ đúng bằng số người: C = Σₓ F(x). Độ rộng σ có hai cách chọn: **cố định** (σ = 15 px cho ShanghaiTech Part B, như CSRNet [7]) khi kích thước đầu người tương đối đồng đều; hoặc **thích ứng hình học** σᵢ = β·d̄ᵢ với d̄ᵢ là khoảng cách trung bình tới 3 người gần nhất, β = 0,3 [6], dùng cho đám đông rất dày có phối cảnh mạnh (Part A).")
FIG("Ảnh gốc và bản đồ mật độ tương ứng, tạo bằng nhân Gauss thích ứng hình học (geometry-adaptive).",
    "Nguồn: Y. Zhang, D. Zhou, S. Chen, S. Gao, Y. Ma, “Single-Image Crowd Counting via Multi-Column Convolutional Neural Network”, CVPR 2016, Fig. 2 [6].", existing="mcnn_fig2")
P("Hình 2.3 cho thấy giá trị của bản đồ mật độ so với một con số đếm: ba ảnh cùng có 95 người nhưng phân bố hoàn toàn khác nhau, nên mức độ nguy hiểm cục bộ cũng khác nhau.")
FIG("Ba ảnh trong ShanghaiTech Part B cùng chứa 95 người nhưng phân bố không gian khác nhau (hàng trên) và bản đồ mật độ tương ứng (hàng dưới).",
    "Nguồn: Y. Li, X. Zhang, D. Chen, “CSRNet: Dilated Convolutional Neural Networks for Understanding the Highly Congested Scenes”, CVPR 2018 (arXiv:1802.10062), Fig. 1 [7].", existing="csrnet_fig1")

H2("2.4. Mạng CSRNet")
P("MCNN [6] dùng ba nhánh CNN song song với kích thước bộ lọc khác nhau để bắt người ở nhiều tỉ lệ. Li và cộng sự [7] chỉ ra cấu trúc nhiều cột này dư thừa, khó huấn luyện, và đề xuất **CSRNet** — một mạng một cột, sâu hơn, gồm ba phần:")
BUL(["**Frontend:** 10 lớp tích chập 3×3 đầu tiên của VGG-16 [9] (đã huấn luyện trước trên ImageNet), xen 3 lớp max-pooling nên đầu ra có độ phân giải bằng 1/8 ảnh vào.",
     "**Backend:** 6 lớp **tích chập giãn** (dilated convolution) [10] với hệ số giãn r = 2, số kênh 512–512–512–256–128–64.",
     "**Đầu ra:** một lớp tích chập 1×1 sinh bản đồ mật độ một kênh."])
P("Tích chập giãn chèn khoảng trống giữa các phần tử của bộ lọc:")
FORMULA("y(m, n) = Σᵢ Σⱼ x(m + r·i, n + r·j) · w(i, j)")
P("Với r = 2, bộ lọc 3×3 có vùng tiếp nhận 5×5 mà không tăng số tham số và không phải giảm độ phân giải bằng pooling. Nhờ đó backend nắm được ngữ cảnh rộng (cần để phân biệt đầu người với nền phức tạp) trong khi vẫn giữ chi tiết không gian của bản đồ mật độ. Mạng được huấn luyện bằng hàm mất mát bình phương sai số giữa bản đồ dự đoán và bản đồ ground truth:")
FORMULA("L(Θ) = 1/(2N) · Σᵢ ‖ D(Xᵢ; Θ) − Dᵢᴳᵀ ‖²")
FIG("Kết quả của CSRNet trên ShanghaiTech Part A: cột trái là ảnh gốc, cột giữa là bản đồ mật độ thực (ground truth), cột phải là bản đồ do CSRNet sinh ra.",
    "Nguồn: Y. Li et al., CVPR 2018, Fig. 7 [7].", existing="csrnet_fig7")

H2("2.5. Phát hiện đám đông bất thường dựa trên mật độ")
P("“Bất thường” là các sự kiện hiếm gặp, khác biệt với trạng thái thường ngày của cảnh. Nhiều nghiên cứu phát hiện bất thường về hành vi bằng mô hình học sâu trên video, như dự đoán khung hình tương lai [11] hay học đa thể hiện trên video giám sát thực tế [12]; tuy nhiên các mô hình này cần dữ liệu video riêng cho từng cảnh, chi phí tính toán lớn và có tỉ lệ cảnh báo giả cao trên dữ liệu thực. Với mục tiêu phòng chống chèn ép đám đông, đề tài chọn hướng **phát hiện bất thường dựa trên mật độ** — dùng trực tiếp đầu ra của mô hình đếm, dễ giải thích, dễ hiệu chỉnh — gồm hai luật:")
BUL(["**Quá tải:** mức cảnh báo được xác định bằng cách so mật độ ρₜ = Cₜ / A với các ngưỡng ở Bảng 2.1.",
     "**Tăng đột biến:** số người tăng nhanh trong cửa sổ W giây — dấu hiệu dòng người đang dồn vào khu vực, xảy ra trước khi mật độ đạt mức nguy hiểm. Khi có tăng đột biến, mức cảnh báo được nâng thêm một bậc:"])
FORMULA("Rₜ = (C̃ₜ − C̃ₜ₋W) / C̃ₜ₋W ≥ θ   và   C̃ₜ₋W ≥ C_min")
P("Để tránh cảnh báo chập chờn do nhiễu của mô hình, số đếm được làm trơn bằng trung bình trượt hàm mũ C̃ₜ = α·Cₜ + (1 − α)·C̃ₜ₋₁, và mức cảnh báo chỉ thay đổi khi mức mới được xác nhận ở K lần đo liên tiếp. Điều kiện C̃ₜ₋W ≥ C_min loại bỏ các “tăng đột biến” vô nghĩa khi khu vực gần như vắng người.")

H2("2.6. Độ đo đánh giá")
P("Mô hình đếm người được đánh giá bằng sai số tuyệt đối trung bình (MAE) và căn bậc hai sai số bình phương trung bình (RMSE; nhiều bài báo vẫn ghi là MSE):")
FORMULA("MAE = (1/N)·Σᵢ |Cᵢ − Ĉᵢ|        RMSE = √( (1/N)·Σᵢ (Cᵢ − Ĉᵢ)² )")
P("trong đó Cᵢ là số người thực, Ĉᵢ là số người ước lượng trên ảnh thứ i. MAE phản ánh độ chính xác, RMSE phạt nặng các ảnh sai lớn nên phản ánh độ ổn định. Ngoài ra đề tài đo tốc độ xử lý (ảnh/giây) và tỉ lệ ảnh có mức cảnh báo từ số đếm dự đoán trùng với mức từ số đếm thực.")

H2("2.7. Kết quả công bố của các công trình liên quan")
P("Bảng 2.2 tổng hợp kết quả do các tác giả công bố trên ShanghaiTech. Đây là mốc tham chiếu để so sánh với mô hình do nhóm huấn luyện ở Chương 4. Sai số giảm mạnh qua các thế hệ mô hình: CSRNet giảm khoảng 60% MAE trên Part B so với MCNN nhờ mạng sâu hơn và tích chập giãn; P2PNet [23] (dự đoán trực tiếp tập điểm đầu người thay vì bản đồ mật độ) giảm thêm khoảng 40% so với CSRNet.")
TABLE("Sai số đếm người trên ShanghaiTech theo công bố của các tác giả (càng thấp càng tốt)",
      ["Phương pháp", "Part A – MAE", "Part A – RMSE", "Part B – MAE", "Part B – RMSE"],
      [["MCNN (CVPR 2016) [6]", "110,2", "173,2", "26,4", "41,3"],
       ["CSRNet (CVPR 2018) [7]", "68,2", "115,0", "10,6", "16,0"],
       ["P2PNet (ICCV 2021) [23]", "52,74", "85,06", "6,25", "9,9"]],
      [3.2, 1.6, 1.6, 1.6, 1.6], center_cols=(1, 2, 3, 4),
      note="Số liệu do các tác giả công bố trong bài báo gốc.")

# ---------------------------------------------------------------- CHƯƠNG 3
H1("CHƯƠNG 3. DỮ LIỆU VÀ TIỀN XỬ LÝ")
H2("3.1. Khảo sát và lựa chọn bộ dữ liệu")
TABLE("Một số bộ dữ liệu chuẩn cho bài toán đếm người",
      ["Bộ dữ liệu", "Quy mô", "Đặc điểm"],
      [["ShanghaiTech Part A (2016) [6]", "482 ảnh; 241.677 người", "Ảnh Internet, đám đông rất dày"],
       ["ShanghaiTech Part B (2016) [6]", "716 ảnh; 88.488 người", "Ảnh chụp đường phố Thượng Hải từ trên cao, kích thước cố định 768×1024"],
       ["UCF-QNRF (2018) [17]", "1.535 ảnh; khoảng 1,25 triệu người", "Đám đông cực dày, độ phân giải cao"],
       ["JHU-Crowd++ (2020) [18]", "4.372 ảnh; khoảng 1,51 triệu người", "Đa dạng thời tiết, ánh sáng"],
       ["NWPU-Crowd (2020) [19]", "5.109 ảnh; 2.133.375 người", "Lớn nhất, có cả ảnh không có người"]],
      [3, 3, 4], note="Số liệu lấy từ các bài báo gốc tương ứng.")
P("Đề tài chọn **ShanghaiTech Part B** vì: (i) ảnh chụp đường phố, khu mua sắm từ góc cao — gần với góc nhìn camera giám sát thực tế; (ii) số người mỗi ảnh từ 9 đến 578, đủ trải các mức cảnh báo; (iii) quy mô vừa phải, ảnh cùng kích thước nên huấn luyện được trên GPU miễn phí của Google Colab trong khoảng một giờ; (iv) là bộ chuẩn được dùng rộng rãi nên kết quả so sánh được với các công trình đã công bố (MCNN, CSRNet).")
FIG("Ảnh mẫu trong bộ dữ liệu ShanghaiTech: (a) Part A — đám đông rất dày; (b) Part B — đường phố Thượng Hải.",
    "Nguồn: Y. Zhang et al., CVPR 2016, Fig. 1 [6].", existing="mcnn_fig1")

H2("3.2. Cấu trúc và thống kê dữ liệu")
P("Notebook tải bộ dữ liệu bằng thư viện kagglehub và tự dò thư mục part_B. Cấu trúc thư mục:")
CODE(["part_B/",
      "├── train_data/  images/IMG_1.jpg …    ground-truth/GT_IMG_1.mat …   (400 ảnh)",
      "└── test_data/   images/IMG_1.jpg …    ground-truth/GT_IMG_1.mat …   (316 ảnh)"])
P("Mỗi file .mat chứa trường image_info gồm ma trận N×2 tọa độ (x, y) của N đầu người. Bảng 3.2 tóm tắt thống kê do notebook tính (mục 3 của notebook).")
TABLE("Thống kê số người trên ShanghaiTech Part B",
      ["Tập", "Số ảnh", "Tổng số người", "Ít nhất", "Nhiều nhất", "Trung bình"],
      [["Train", R("data_stats.train.n", "{:d}", "400"), R("data_stats.train.total", "{:,d}"), R("data_stats.train.min", "{:d}"), R("data_stats.train.max", "{:d}"), R("data_stats.train.mean", "{:.1f}")],
       ["Test", R("data_stats.test.n", "{:d}", "316"), R("data_stats.test.total", "{:,d}"), R("data_stats.test.min", "{:d}"), R("data_stats.test.max", "{:d}"), R("data_stats.test.mean", "{:.1f}")],
       ["Toàn bộ (theo [6])", "716", "88.488", "9", "578", "123,6"]],
      [2.2, 1.3, 1.8, 1.3, 1.5, 1.5], center_cols=(1, 2, 3, 4, 5),
      note="Bản dữ liệu tải từ Kaggle có tổng " + R("data_stats.train.total", "{:,d}") + " + " + R("data_stats.test.total", "{:,d}")
           + " người, lệch nhẹ so với con số 88.488 công bố trong [6]; báo cáo dùng số liệu đo trực tiếp trên bản dữ liệu đã huấn luyện.")

H2("3.3. Khảo sát dữ liệu (EDA)")
P("Trước khi huấn luyện, nhóm khảo sát dữ liệu để hiểu đặc điểm của bài toán và lấy căn cứ chọn tham số tiền xử lý. Toàn bộ hình trong mục này do notebook (mục 3) sinh ra từ " + R("eda.so_anh", "{:d}", "716") + " ảnh, cùng kích thước 768×1024.")
FIG("Ảnh train sắp xếp theo số người, từ thưa (trái trên) đến dày (phải dưới).", "Nguồn: nhóm thực hiện (notebook, mục 3.2).",
    path=OUT("gallery.png"), width=16, placeholder=PENDING("gallery.png"))
H3("3.3.1. Phân bố số người")
P("Số người mỗi ảnh lệch phải mạnh: trung vị chỉ khoảng " + R("data_stats.train.median", "{:.0f}") + " người nhưng có ảnh tới " + R("data_stats.train.max", "{:d}") + " người (Hình 3.3). Vì vậy MAE bị chi phối bởi số ít ảnh rất đông, nên Chương 4 phân tích thêm sai số theo nhóm mật độ. Kiểm định Kolmogorov–Smirnov giữa train và test cho p-value = " + R("eda.ks_pvalue_train_test", "{:.3f}") + " > 0,05, tức hai tập có phân bố số người không khác biệt có ý nghĩa thống kê — kết quả trên test phản ánh đúng năng lực trên dữ liệu cùng loại.")
FIG("Phân bố số người mỗi ảnh: histogram, boxplot và hàm phân phối tích luỹ của tập train và test.", "Nguồn: nhóm thực hiện (notebook, mục 3.3).",
    path=OUT("hist_counts.png"), width=16, placeholder=PENDING("hist_counts.png"))
FIG("Nhãn điểm đầu người trên ba ảnh thưa – vừa – dày; hàng dưới phóng to ô 160×160 px đông nhất (khung vàng).", "Nguồn: nhóm thực hiện (notebook, mục 3.4).",
    path=OUT("sample_points.png"), width=15.5, placeholder=PENDING("sample_points.png"))
H3("3.3.2. Phân bố không gian và hiệu ứng phối cảnh")
P("Gộp vị trí của toàn bộ " + R("data_stats.train.total", "{:,d}") + " đầu người trên tập train (Hình 3.5) cho thấy người tập trung dày ở khoảng 1/5 phía trên ảnh — vùng xa camera — và thưa dần về phía dưới. Nguyên nhân là camera đặt cao nhìn chéo xuống: vùng xa phủ diện tích mặt đất lớn hơn và người trông nhỏ hơn.")
FIG("Vị trí của mọi đầu người trên tập train (trái) và phân bố theo chiều dọc ảnh (phải).", "Nguồn: nhóm thực hiện (notebook, mục 3.5).",
    path=OUT("spatial_heatmap.png"), width=15.5, placeholder=PENDING("spatial_heatmap.png"))
P("Khoảng cách từ mỗi đầu người tới đầu người gần nhất là thước đo gián tiếp kích thước đầu người. Trung vị của khoảng cách này là **" + R("eda.nn_median_px", "{:.1f}") + " px**, gần như trùng với σ = 15 px được chọn cho nhân Gauss — đây là căn cứ thực nghiệm cho lựa chọn σ của đề tài. Khoảng cách tăng mạnh theo chiều dọc: trung vị khoảng " + R("eda.nn_median_top_vs_bottom.0", "{:.0f}") + " px ở dải trên cùng nhưng tới khoảng " + R("eda.nn_median_top_vs_bottom.1", "{:.0f}") + " px ở dải dưới cùng (Hình 3.6), cho thấy hiệu ứng phối cảnh rất rõ; σ cố định vì vậy chỉ là một xấp xỉ.")
FIG("Phân bố khoảng cách tới đầu người gần nhất (trái) và sự thay đổi của nó theo vị trí dọc trong ảnh (phải).", "Nguồn: nhóm thực hiện (notebook, mục 3.6).",
    path=OUT("perspective_nn.png"), width=15.5, placeholder=PENDING("perspective_nn.png"))
H3("3.3.3. Mật độ cục bộ")
P("Chia mỗi ảnh thành các ô 128×128 px: **" + R("eda.ti_le_o_trong_%", "{:.0f}") + "% số ô không có người nào**, trong khi ô đông nhất chứa tới " + R("eda.o_dong_nhat", "{:d}") + " người (Hình 3.7). Bản đồ mật độ vì vậy rất thưa — phần lớn giá trị bằng 0 — nên hàm mất mát MSE dễ bị vùng nền chi phối; đây là lý do nhóm nhân bản đồ mật độ với hệ số 100 khi huấn luyện. Số người ở ô đông nhất tăng gần tuyến tính theo tổng số người của ảnh: ảnh càng đông thì càng có điểm nóng cục bộ dày, đúng loại vùng cần cảnh báo.")
FIG("Phân bố số người trong mỗi ô 128×128 px (trái, thang log) và quan hệ giữa tổng số người với ô đông nhất của từng ảnh (phải).", "Nguồn: nhóm thực hiện (notebook, mục 3.7).",
    path=OUT("local_density.png"), width=15.5, placeholder=PENDING("local_density.png"))
H3("3.3.4. Đặc trưng ảnh và mức cảnh báo")
P("Độ sáng trung bình của ảnh nằm trong khoảng " + R("eda.do_sang_min_max.0", "{:.0f}") + "–" + R("eda.do_sang_min_max.1", "{:.0f}") + " (thang 0–255), phân bố giống nhau giữa train và test, và không có tương quan rõ với số người (Hình 3.8). Bộ dữ liệu hầu như không có ảnh ban đêm, nên mô hình chưa được kiểm chứng trong điều kiện thiếu sáng — một hạn chế cần lưu ý khi triển khai.")
FIG("Độ sáng, độ tương phản của ảnh và quan hệ giữa độ sáng với số người.", "Nguồn: nhóm thực hiện (notebook, mục 3.8).",
    path=OUT("image_props.png"), width=16, placeholder=PENDING("image_props.png"))
lv = RES.get("eda", {}).get("so_anh_theo_muc", {})
P("Giả định mỗi ảnh phủ 100 m² và áp ngưỡng đám đông di chuyển, dữ liệu trải đủ bốn mức cảnh báo: "
  + (", ".join(f"{k} {v.get('train', 0)} ảnh train / {v.get('test', 0)} ảnh test" for k, v in lv.items()) if lv else "…")
  + " (Hình 3.9). Mức Cam và Đỏ chiếm tỉ lệ nhỏ — đúng với thực tế là tình huống nguy hiểm hiếm — nên khi đánh giá luật cảnh báo cần xem riêng các ảnh đông thay vì chỉ nhìn tỉ lệ đúng chung.")
FIG("Số ảnh thuộc từng mức cảnh báo (giả định A = 100 m²).", "Nguồn: nhóm thực hiện (notebook, mục 3.9).",
    path=OUT("level_distribution.png"), width=11, placeholder=PENDING("level_distribution.png"))

H2("3.4. Sinh bản đồ mật độ ground truth")
P("Bản đồ mật độ được sinh một lần và lưu thành file .npy theo các bước:")
NUM(["Đọc tọa độ đầu người từ file .mat; điểm nằm ngoài ảnh được kẹp vào biên.",
     "Đặt tại mỗi điểm một nhân Gauss σ = 15 px, cắt ở bán kính 3σ. Phần nhân bị cắt ở mép ảnh được **chuẩn hóa lại** để tổng vẫn bằng 1 — đảm bảo tổng bản đồ đúng bằng số người.",
     "**Cộng gộp theo khối 8×8** để đưa bản đồ về độ phân giải 1/8 (96×128 với ảnh 768×1024) — đúng kích thước đầu ra của CSRNet. Cộng gộp (thay vì nội suy) giữ nguyên tổng, tức giữ nguyên số người.",
     "Lưu bản đồ thành file .npy để không phải tính lại ở mỗi epoch."])
FIG("Ảnh huấn luyện và bản đồ mật độ ground truth 1/8 do notebook sinh ra.", "Nguồn: nhóm thực hiện (notebook, mục 4).",
    path=OUT("density_gt.png"), width=15, placeholder=PENDING("density_gt.png"))
P("Hình 3.11 so sánh các cách chọn σ trên cùng vùng đông nhất của một ảnh. σ = 4 px cho các “gai” nhọn tách rời, khó học; σ = 30 px làm các đầu người sát nhau hoà thành một khối, mất thông tin vị trí; σ = 15 px — xấp xỉ khoảng cách trung vị giữa hai đầu người gần nhau (mục 3.3.2) — cân bằng giữa hai thái cực. Tổng của mọi bản đồ đều đúng bằng số người nhờ bước chuẩn hoá nhân Gauss.")
FIG("So sánh bản đồ mật độ với σ = 4, 15, 30 px và σ thích ứng trên cùng một vùng ảnh đông.", "Nguồn: nhóm thực hiện (notebook, mục 4).",
    path=OUT("sigma_compare.png"), width=16, placeholder=PENDING("sigma_compare.png"))

H2("3.5. Chia tập và tăng cường dữ liệu")
BUL(["**Chia tập:** 400 ảnh train gốc được chia ngẫu nhiên (seed 42) thành **360 ảnh huấn luyện và 40 ảnh kiểm định (val)**; val dùng để chọn epoch tốt nhất. **316 ảnh test** chỉ dùng một lần ở cuối để báo cáo kết quả, tránh “nhìn trộm” tập test.",
     "**Cắt ngẫu nhiên** ô 512×512 với tọa độ là bội số của 8 để ô cắt trên bản đồ 1/8 khớp chính xác với ô cắt trên ảnh.",
     "**Lật ngang** ngẫu nhiên (xác suất 0,5) đồng thời cho ảnh và bản đồ mật độ.",
     "**Thay đổi độ sáng** ngẫu nhiên trong khoảng ±20% để mô phỏng điều kiện ánh sáng khác nhau.",
     "**Chuẩn hóa** theo trung bình và độ lệch chuẩn của ImageNet vì frontend là VGG-16 huấn luyện trước trên ImageNet."])
FIG("Hàm phân phối tích luỹ số người của ba tập train / val / test.", "Nguồn: nhóm thực hiện (notebook, mục 5).",
    path=OUT("split_cdf.png"), width=11, placeholder=PENDING("split_cdf.png"))
FIG("Mẫu huấn luyện sau tăng cường (mỗi ảnh lấy hai lần cho hai ô cắt khác nhau), bản đồ mật độ phủ lên ảnh.", "Nguồn: nhóm thực hiện (notebook, mục 5).",
    path=OUT("augment_preview.png"), width=16, placeholder=PENDING("augment_preview.png"))

# ---------------------------------------------------------------- CHƯƠNG 4
H1("CHƯƠNG 4. HUẤN LUYỆN VÀ ĐÁNH GIÁ MÔ HÌNH")
H2("4.1. Môi trường thực nghiệm")
P(f"Mô hình được cài đặt bằng PyTorch [20], xử lý ảnh bằng OpenCV, chạy trên Google Colab với GPU {R('gpu', '{}', 'NVIDIA T4')}. Toàn bộ quy trình nằm trong notebook train_csrnet_shanghaitech.ipynb và có thể chạy lại từ đầu bằng lệnh Run all; các hình và số liệu trong chương này được notebook xuất ra thư mục outputs/.")

H2("4.2. Kiến trúc mô hình cài đặt")
TABLE("Kiến trúc CSRNet cài đặt (ảnh vào 3×768×1024)",
      ["Khối", "Cấu hình lớp", "Kích thước đầu ra"],
      [["Frontend – khối 1", "2 × Conv3×3-64, ReLU; MaxPool 2×2", "64 × 384 × 512"],
       ["Frontend – khối 2", "2 × Conv3×3-128, ReLU; MaxPool 2×2", "128 × 192 × 256"],
       ["Frontend – khối 3", "3 × Conv3×3-256, ReLU; MaxPool 2×2", "256 × 96 × 128"],
       ["Frontend – khối 4", "3 × Conv3×3-512, ReLU", "512 × 96 × 128"],
       ["Backend", "3 × Conv3×3-512, 1 × Conv3×3-256, 1 × Conv3×3-128, 1 × Conv3×3-64; dilation = 2, ReLU", "64 × 96 × 128"],
       ["Đầu ra", "Conv1×1-1", "1 × 96 × 128"]],
      [2.2, 5, 2.2], center_cols=(2,),
      note="Tổng số tham số: 16,26 triệu. Trọng số frontend khởi tạo từ VGG-16 ImageNet; backend khởi tạo Gauss với độ lệch chuẩn 0,01.")

H2("4.3. Cấu hình huấn luyện")
TABLE("Siêu tham số huấn luyện",
      ["Tham số", "Giá trị", "Ghi chú"],
      [["Hàm mất mát", "MSE (tổng theo điểm ảnh / batch)", "Giữa bản đồ dự đoán và ground truth"],
       ["Thuật toán tối ưu", "Adam [13]", "weight decay 1e-4"],
       ["Tốc độ học", "1e-5, giảm theo cosine", "Giảm dần tới 5% giá trị ban đầu"],
       ["Số epoch", R("cfg.epochs", "{:d}", "150"), "Đánh giá trên val sau mỗi 2 epoch"],
       ["Batch size", "4", "Ô cắt 512×512"],
       ["Hệ số nhân mật độ", "100", "Nhân bản đồ GT để gradient không quá nhỏ; chia lại khi đếm"],
       ["Độ chính xác hỗn hợp (AMP)", "Có", "Tăng tốc trên GPU"],
       ["Chọn mô hình", "MAE nhỏ nhất trên val", "Lưu checkpoint tốt nhất"]],
      [2.5, 3, 4])

H2("4.4. Quá trình huấn luyện")
FIG("Hàm mất mát trên tập train và sai số MAE/RMSE trên tập val theo epoch.", "Nguồn: nhóm thực hiện (notebook, mục 7).",
    path=OUT("training_curves.png"), width=15, placeholder=PENDING("training_curves.png"))
if RES:
    last_val = vn("%.2f" % HIST[-1]["val_mae"]) if HIST else "…"
    P("Hàm mất mát trên tập train giảm nhanh trong khoảng 20 epoch đầu rồi giảm chậm dần. Sai số trên tập val dao động mạnh trong khoảng 100 epoch đầu (có những đỉnh MAE vượt 40 khi tốc độ học còn lớn), sau đó ổn định quanh 11–12 khi tốc độ học giảm dần theo lịch cosine. Mô hình tốt nhất đạt ở **epoch " + R("best_epoch", "{:d}") + "** với MAE val = " + R("val_mae") + "; ở epoch cuối MAE val nhích lên " + last_val + ", cho thấy việc lưu checkpoint theo MAE val tốt nhất (thay vì lấy epoch cuối) là cần thiết. Mô hình được huấn luyện " + R("cfg.epochs", "{:d}") + " epoch trên GPU " + R("gpu", "{}") + " của Google Colab.")
else:
    P(f"Mô hình tốt nhất đạt được ở epoch {R('best_epoch', '{:d}')} với MAE trên tập val bằng {R('val_mae')}.")

H2("4.5. Kết quả trên tập test và so sánh với các công trình đã công bố")
PAPER_B = [("MCNN (CVPR 2016) [6]", 26.4, 41.3), ("CSRNet – bài báo gốc (CVPR 2018) [7]", 10.6, 16.0), ("P2PNet (ICCV 2021) [23]", 6.25, 9.9)]
BASE_MAE, BASE_RMSE = 10.6, 16.0


def delta(v, base):
    d = v - base
    return vn("%+.2f (%+.1f%%)" % (d, d / base * 100))


rows = []
for name, m_, r_ in PAPER_B:
    is_base = m_ == BASE_MAE
    rows.append([name, vn(str(m_)), vn(str(r_)),
                 "mốc so sánh" if is_base else delta(m_, BASE_MAE), "mốc so sánh" if is_base else delta(r_, BASE_RMSE)])
if RES:
    rows.append(["**CSRNet – nhóm huấn luyện lại**", "**" + R("test_mae") + "**", "**" + R("test_rmse") + "**",
                 "**" + delta(RES["test_mae"], BASE_MAE) + "**", "**" + delta(RES["test_rmse"], BASE_RMSE) + "**"])
else:
    rows.append(["**CSRNet – nhóm huấn luyện lại**", "…", "…", "…", "…"])
TABLE("So sánh sai số đếm người trên tập test ShanghaiTech Part B (càng thấp càng tốt)",
      ["Phương pháp", "MAE", "RMSE", "Δ MAE so với CSRNet gốc", "Δ RMSE so với CSRNet gốc"],
      rows, [3.6, 1.1, 1.1, 2, 2], center_cols=(1, 2, 3, 4),
      note="Ba dòng đầu là số liệu các tác giả công bố trong bài báo gốc; dòng cuối do notebook đo trên 316 ảnh test. Δ âm nghĩa là sai số thấp hơn (tốt hơn) CSRNet gốc.")
if RES:
    mae, rmse = RES["test_mae"], RES["test_rmse"]
    if mae < BASE_MAE and rmse < BASE_RMSE:
        verdict = ("**thấp hơn kết quả công bố của bài báo gốc " + vn("%.1f" % ((BASE_MAE - mae) / BASE_MAE * 100))
                   + "% về MAE và " + vn("%.1f" % ((BASE_RMSE - rmse) / BASE_RMSE * 100)) + "% về RMSE**")
    else:
        verdict = "so với kết quả công bố của bài báo gốc là 10,6 / 16,0"
    P("Với cùng kiến trúc CSRNet, mô hình do nhóm huấn luyện đạt **MAE = " + vn("%.2f" % mae) + " và RMSE = " + vn("%.2f" % rmse) + "**, "
      + verdict + ", và thấp hơn MCNN " + vn("%.1f" % ((26.4 - mae) / 26.4 * 100))
      + "% về MAE. RMSE giảm cho thấy mô hình không chỉ chính xác hơn về trung bình mà còn ít mắc các lỗi lớn hơn.")
    P("Cần đánh giá kết quả này một cách thận trọng: (i) đây là kết quả của một lần huấn luyện (seed 42), chưa lặp lại nhiều lần để ước lượng độ dao động; (ii) nhóm chưa làm thí nghiệm loại bỏ từng thay đổi (ablation) nên chưa tách được đóng góp riêng của từng yếu tố trong Bảng 4.4; (iii) P2PNet — kiến trúc mới hơn, dự đoán trực tiếp vị trí từng người — vẫn tốt hơn rõ rệt (MAE 6,25). Như vậy cải tiến của nhóm nằm ở **quy trình huấn luyện**, không phải ở kiến trúc mạng.")

H3("Những điểm nhóm thay đổi so với CSRNet gốc")
TABLE("So sánh cấu hình huấn luyện giữa CSRNet gốc và nhóm thực hiện",
      ["Thành phần", "CSRNet gốc [7]", "Nhóm thực hiện", "Tác dụng dự kiến"],
      [["Thuật toán tối ưu", "SGD, tốc độ học cố định 1e-6", "Adam, tốc độ học 1e-5 giảm theo cosine", "Hội tụ nhanh hơn; tốc độ học nhỏ dần giúp mô hình ổn định ở cuối (Hình 4.1)"],
       ["Bản đồ mật độ 1/8", "Nội suy bicubic bản đồ gốc rồi nhân 64 (mã nguồn công khai)", "Cộng gộp khối 8×8; chuẩn hoá nhân Gauss bị cắt ở biên", "Tổng bản đồ đúng bằng số người — nhãn huấn luyện chính xác hơn"],
       ["Tăng cường dữ liệu", "9 ô cắt bằng 1/4 ảnh + lật ngang", "Cắt ngẫu nhiên 512×512 mỗi lần lấy mẫu + lật ngang + đổi độ sáng ±20%", "Mỗi epoch thấy các ô cắt khác nhau, dữ liệu đa dạng hơn"],
       ["Thang giá trị mật độ", "Giữ nguyên", "Nhân 100 khi huấn luyện", "Gradient đủ lớn khi 54% vùng ảnh trống (mục 3.3.3)"],
       ["Chọn mô hình", "—", "Tách 40 ảnh val, lưu checkpoint có MAE val tốt nhất; test chỉ dùng một lần", "Tránh quá khớp, kết quả test khách quan"],
       ["Tốc độ huấn luyện", "—", "Độ chính xác hỗn hợp (AMP), batch 4", "Huấn luyện xong trong khoảng 1 giờ trên GPU T4 miễn phí"]],
      [2, 2.7, 3, 3.3],
      note="Cấu hình CSRNet gốc theo bài báo [7] và mã nguồn công khai của tác giả; “—” là thông tin không được nêu.")

H2("4.6. Phân tích sai số")
bins = RES.get("by_bin", [])
TABLE("Sai số theo nhóm mật độ trên tập test",
      ["Số người trong ảnh", "Số ảnh", "MAE", "Sai số tương đối"],
      [[b["nhom"], str(b["so_anh"]), vn("%.2f" % b["mae"]), vn("%.1f%%" % b["sai_so_tuong_doi_%"])] for b in bins]
      or [[g, "…", "…", "…"] for g in ("0–50", "50–100", "100–200", "200–∞")],
      [3, 2, 2, 2.5], center_cols=(0, 1, 2, 3))
if bins:
    rel = [b["sai_so_tuong_doi_%"] for b in bins]
    P("MAE tăng theo số người (từ " + vn("%.1f" % bins[0]["mae"]) + " ở nhóm thưa nhất lên " + vn("%.1f" % bins[-1]["mae"])
      + " ở nhóm đông nhất), nhưng **sai số tương đối gần như không đổi, chỉ trong khoảng " + vn("%.1f" % min(rel)) + "–" + vn("%.1f" % max(rel))
      + "%** ở mọi nhóm. Nghĩa là sai số tỉ lệ với quy mô đám đông, mô hình không bị “sập” ở ảnh rất đông. Với mục tiêu cảnh báo, sai số tương đối khoảng 8% chỉ làm lệch mức cảnh báo khi mật độ thực nằm sát ngưỡng (mục 5.2).")
FIG("Số người dự đoán so với số người thực trên 316 ảnh test (đường đỏ: dự đoán đúng tuyệt đối).", "Nguồn: nhóm thực hiện (notebook, mục 8).",
    path=OUT("scatter_test.png"), width=10, placeholder=PENDING("scatter_test.png"))
FIG("Một số ảnh test: ảnh gốc, bản đồ mật độ ground truth và bản đồ do CSRNet dự đoán.", "Nguồn: nhóm thực hiện (notebook, mục 8).",
    path=OUT("test_samples.png"), width=15.5, placeholder=PENDING("test_samples.png"))
if RES:
    P("Tốc độ suy luận đạt **" + vn("%.1f" % RES["fps_768x1024"]) + " ảnh/giây** với ảnh 768×1024 trên GPU " + RES["gpu"]
      + "; trên máy tính cá nhân chỉ có CPU khoảng 2 giây/ảnh. Trong vận hành chỉ cần phân tích khoảng 1 khung hình/giây cho mỗi camera, nên một GPU tầm trung có thể phục vụ hơn chục camera.")

H2("4.7. Thử nghiệm trên ảnh ngoài bộ dữ liệu")
P("Để đánh giá khả năng tổng quát hoá, nhóm chạy mô hình trên bốn ảnh không thuộc ShanghaiTech (các ảnh Wikimedia Commons ở Chương 1 và 6). Các ảnh này không có nhãn nên chỉ đánh giá định tính (Hình 4.4):")
BUL(["**Đám đông rất dày (Kaaba mùa Hajj):** bản đồ mật độ phủ đúng vùng sân có người và để trống khối Kaaba ở giữa, dù loại cảnh này dày hơn nhiều so với dữ liệu huấn luyện.",
     "**Itaewon 29/10/2022:** mô hình bắt tốt đám đông nhỏ ở phía xa, nhưng **bỏ sót những người ở gần camera** (kích thước lớn) ở tiền cảnh — vì trong ShanghaiTech Part B người chủ yếu có kích thước nhỏ (mục 3.3.2).",
     "**Ảnh camera CCTV không có người:** mô hình vẫn cho khoảng 3–4 “người” giả ở ống khói và cột camera — các vật thể có hình dạng tương tự đầu người khi nhìn từ xa."])
P("Hai hiện tượng sau là minh chứng cụ thể cho **khác biệt miền dữ liệu** (domain gap): khi triển khai thực tế cần tinh chỉnh mô hình bằng ảnh từ chính camera tại hiện trường, bổ sung ảnh người ở cận cảnh và ảnh nền không có người, và đặt vùng quan tâm (ROI) để loại bỏ các vùng không phải mặt đất.")
FIG("Kết quả của mô hình trên bốn ảnh ngoài bộ dữ liệu: ảnh gốc (hàng trên) và bản đồ mật độ dự đoán (hàng dưới).",
    "Nguồn: nhóm thực hiện; ảnh gốc từ Wikimedia Commons (xem nguồn ở Hình 1.1, 1.2, 1.3, 6.1).",
    path=OUT("wild_test.png"), width=16, placeholder=PENDING("wild_test.png"))

# ---------------------------------------------------------------- CHƯƠNG 5
H1("CHƯƠNG 5. MÔ-ĐUN CẢNH BÁO VÀ ỨNG DỤNG DEMO")
H2("5.1. Quy đổi số đếm thành mật độ")
P("Mật độ được tính bằng ρ = C / A, trong đó A là diện tích thực (m²) của vùng mặt đất mà camera quan sát. Khi lắp đặt, A được xác định bằng cách đo trực tiếp hoặc hiệu chỉnh camera (ma trận đồng nhất – homography – giữa mặt phẳng ảnh và mặt đất). Trong ứng dụng demo, người vận hành nhập A cho từng camera. Cách tính này giả định mật độ phân bố tương đối đều trong khung hình; với camera có phối cảnh mạnh, hướng phát triển là dùng bản đồ diện tích theo từng điểm ảnh từ homography.")

H2("5.2. Thuật toán cảnh báo (CrowdMonitor)")
P("Lớp CrowdMonitor trong crowd_core.py nhận lần lượt các cặp (thời điểm t, số đếm Cₜ) và thực hiện:")
NUM(["Làm trơn số đếm: C̃ₜ = α·Cₜ + (1 − α)·C̃ₜ₋₁.",
     "Tính mật độ ρₜ = C̃ₜ / A và mức thô theo ngưỡng của loại đám đông (đứng hoặc di chuyển) ở Bảng 2.1.",
     "Chỉ đổi mức khi mức thô mới được xác nhận ở K lần đo liên tiếp (chống chập chờn); sinh sự kiện “mật độ tăng/giảm mức”.",
     "Giữ hàng đợi số đếm trong cửa sổ W giây, so C̃ₜ với giá trị đầu cửa sổ để phát hiện tăng đột biến; sinh sự kiện khi tăng đột biến bắt đầu.",
     "Mức cuối cùng = mức mật độ, cộng thêm một bậc nếu đang tăng đột biến (tối đa mức Đỏ)."])
TABLE("Tham số mặc định của mô-đun cảnh báo (chỉnh được trên giao diện)",
      ["Tham số", "Mặc định", "Ý nghĩa"],
      [["A", "100 m²", "Diện tích vùng quan sát"],
       ["Loại đám đông", "di chuyển", "Chọn bộ ngưỡng ở Bảng 2.1"],
       ["α", "0,5", "Hệ số làm trơn số đếm"],
       ["K", "3", "Số lần đo liên tiếp để đổi mức"],
       ["W", "10 giây", "Cửa sổ thời gian xét tăng đột biến"],
       ["θ", "50%", "Ngưỡng tăng đột biến"],
       ["C_min", "10 người", "Số người tối thiểu để xét tăng đột biến"],
       ["Chu kỳ phân tích", "1 giây", "Khoảng thời gian giữa hai khung hình được phân tích"]],
      [2, 2, 5], center_cols=(0, 1))
P(f"Để kiểm tra luật mật độ, notebook (mục 9) giả định mỗi ảnh test phủ A = 100 m² và so mức cảnh báo tính từ số đếm dự đoán với mức tính từ số đếm thực: {R('alert_agreement', '{:.1f}%')} số ảnh cho cùng mức (Hình 5.1). Các trường hợp lệch mức đều là ảnh có mật độ thực nằm sát ranh giới giữa hai mức (quanh 2, 3 hoặc 4 người/m²), phù hợp với sai số tương đối khoảng 8% của mô hình; cơ chế làm trơn và xác nhận K lần liên tiếp giúp giảm cảnh báo chập chờn trong những trường hợp này.")
FIG("Mật độ dự đoán và mật độ thực của các ảnh test (giả định A = 100 m²) so với các vùng mức cảnh báo.", "Nguồn: nhóm thực hiện (notebook, mục 9).",
    path=OUT("alert_levels.png"), width=15.5, placeholder=PENDING("alert_levels.png"))

H2("5.3. Kiến trúc ứng dụng demo")
P("Ứng dụng được viết bằng Python với Streamlit [21] cho giao diện web, OpenCV đọc ảnh/video/luồng camera, PyTorch chạy mô hình; chạy được trên máy tính cá nhân không có GPU. Mã nguồn chia hai tệp: crowd_core.py chứa mô hình, hàm suy luận, hàm vẽ bản đồ nhiệt và lớp CrowdMonitor (dùng chung với notebook); app.py chứa giao diện. Luồng xử lý mỗi khung hình: đọc khung → thu nhỏ về cạnh dài 1024 px và bội số của 8 → CSRNet → số người, bản đồ nhiệt → CrowdMonitor → cập nhật giao diện và nhật ký.")
TABLE("Chức năng của ứng dụng demo",
      ["Chức năng", "Mô tả"],
      [["Phân tích ảnh", "Tải ảnh lên; hiển thị số người, mật độ người/m², thời gian xử lý, bản đồ nhiệt và mức cảnh báo kèm hành động đề xuất"],
       ["Phân tích video", "Tải video lên; phân tích theo chu kỳ, hiển thị khung hình kèm bản đồ nhiệt, biểu đồ số người theo thời gian, mức cảnh báo và cảnh báo tăng đột biến"],
       ["Camera trực tiếp", "Nhận webcam hoặc URL RTSP của camera IP, theo dõi trong khoảng thời gian chọn trước"],
       ["Cấu hình", "Diện tích vùng quan sát, loại đám đông, độ phân giải xử lý, các tham số tăng đột biến và chống chập chờn"],
       ["Nhật ký", "Bảng sự kiện cảnh báo; xuất số liệu theo thời gian ra file CSV"],
       ["Chế độ riêng tư", "Làm mờ ảnh hiển thị (bật mặc định); chỉ bản đồ nhiệt mật độ được thể hiện rõ"]],
      [2.2, 7])

H2("5.4. Cài đặt và chạy")
CODE(["pip install -r requirements.txt",
      "# chép csrnet_partB.pth (tải từ notebook) vào code/weights/",
      "streamlit run app.py"])
P("Trình duyệt mở tại địa chỉ http://localhost:8501. Nếu chưa có file trọng số, ứng dụng báo lỗi và hướng dẫn chạy notebook trước.")

H2("5.5. Kết quả chạy demo")
P("**Phân tích ảnh (Hình 5.2):** với ảnh đám đông Itaewon (Hình 1.1) và diện tích giả định 100 m², ứng dụng đếm được khoảng 52 người, mật độ 0,52 người/m² — mức Xanh; thời gian xử lý khoảng 2 giây trên CPU. Ảnh hiển thị được làm mờ theo chế độ riêng tư, chỉ bản đồ nhiệt thể hiện rõ. Như đã phân tích ở mục 4.7, con số này chưa tính những người ở tiền cảnh gần camera.")
FIG("Giao diện phân tích ảnh: số người, mật độ, thời gian xử lý, bản đồ nhiệt và mức cảnh báo.", "Nguồn: nhóm thực hiện (ảnh chụp màn hình ứng dụng).",
    path=OUT("demo_image.png"), width=15.5, placeholder=PENDING("demo_image.png"))
P("**Phân tích video (Hình 5.3):** do không có video giám sát thực tế có nhãn, nhóm tạo một **video mô phỏng** dài 30 giây bằng cách thu phóng dần từ một vùng nhỏ ra toàn bộ ảnh Hajj (Hình 1.3), khiến số người trong khung hình tăng từ khoảng 80 lên khoảng 7.400 — mô phỏng tình huống dòng người dồn vào khu vực giám sát. Với A = 1.800 m², nhật ký ghi nhận đúng trình tự: giây thứ 5 phát hiện **tăng đột biến** (80 → 146 người trong 10 giây) và nâng mức lên Vàng dù mật độ còn thấp; sau đó mật độ lần lượt vượt các ngưỡng Vàng (2,41 người/m², giây 16), Cam (3,44, giây 21) và Đỏ (4,08, giây 28). Điều này cho thấy luật tăng đột biến giúp **cảnh báo sớm hơn khoảng 11 giây** so với chỉ dựa vào ngưỡng mật độ.")
FIG("Giao diện phân tích video mô phỏng: bản đồ nhiệt, biểu đồ số người theo thời gian, mức cảnh báo và nhật ký sự kiện.", "Nguồn: nhóm thực hiện (ảnh chụp màn hình ứng dụng; video mô phỏng tạo từ ảnh ở Hình 1.3).",
    path=OUT("demo_video.png"), width=15.5, placeholder=PENDING("demo_video.png"))

# ---------------------------------------------------------------- CHƯƠNG 6
H1("CHƯƠNG 6. AN TOÀN THÔNG TIN VÀ QUYỀN RIÊNG TƯ")
P("Một hệ thống giám sát phục vụ an ninh lại có thể trở thành mục tiêu tấn công hoặc công cụ xâm phạm quyền riêng tư nếu không được bảo vệ đúng mức. Chương này phân tích các mối đe dọa đối với pipeline ở Hình 1.4, đề xuất biện pháp khi triển khai thực tế và trình bày các biện pháp đã cài đặt trong ứng dụng demo.")
FIG("Camera giám sát lắp đặt trên đường phố — nguồn dữ liệu đầu vào chính của hệ thống.",
    "Nguồn: Hans Wolff, “CCTV camera on Round Church Street, Cambridge.jpg”, Wikimedia Commons, phạm vi công cộng (Public domain).", existing="cctv")
H2("6.1. Tài sản cần bảo vệ")
BUL(["Luồng video trực tiếp và video lưu trữ (chứa hình ảnh cá nhân).",
     "Thiết bị: camera IP, máy tính xử lý, máy chủ.",
     "Mô hình AI (file trọng số) và dữ liệu huấn luyện.",
     "Số liệu đếm, mật độ, nhật ký cảnh báo — dữ liệu ra quyết định của lực lượng an ninh.",
     "Tài khoản và quyền truy cập của người vận hành."])
H2("6.2. Mô hình mối đe dọa")
TABLE("Phân tích mối đe dọa theo mô hình STRIDE",
      ["Loại (STRIDE)", "Kịch bản tấn công trên hệ thống", "Hậu quả"],
      [["Giả mạo (Spoofing)", "Giả mạo camera/thiết bị gửi khung hình hoặc số liệu giả; đăng nhập bằng mật khẩu mặc định của camera", "Số liệu sai, che giấu tình trạng quá tải"],
       ["Sửa đổi (Tampering)", "Phát lại (replay) video cũ vào luồng RTSP không mã hóa; che, xoay camera; tấn công đối kháng làm mô hình đếm thiếu; thay file trọng số bằng file độc hại", "Hệ thống “mù”, không cảnh báo khi có nguy hiểm; thực thi mã độc"],
       ["Chối bỏ (Repudiation)", "Người vận hành bỏ qua/xóa cảnh báo mà không để lại dấu vết", "Không truy được trách nhiệm sau sự cố"],
       ["Lộ thông tin (Information disclosure)", "Nghe lén luồng video, truy cập trái phép kho lưu trữ, nhân viên nội bộ lạm dụng hình ảnh", "Vi phạm quyền riêng tư, vi phạm pháp luật về dữ liệu cá nhân"],
       ["Từ chối dịch vụ (DoS)", "Làm nghẽn mạng camera hoặc máy chủ xử lý; camera bị chiếm làm botnet", "Mất giám sát đúng lúc cao điểm"],
       ["Leo thang đặc quyền (Elevation of privilege)", "Khai thác lỗ hổng firmware camera hoặc giao diện web để chiếm quyền quản trị", "Kiểm soát toàn bộ hệ thống"]],
      [2.3, 4.5, 3])
P("Nguy cơ chiếm quyền camera là có thật: mạng botnet Mirai (2016) đã lây nhiễm hàng trăm nghìn thiết bị IoT — trong đó có nhiều camera IP và đầu ghi hình — chủ yếu bằng cách dò mật khẩu mặc định, đạt đỉnh khoảng 600.000 thiết bị và được dùng cho các cuộc tấn công từ chối dịch vụ quy mô rất lớn [14].")
P("Với thành phần AI, tấn công đối kháng (adversarial attack) là mối đe dọa đặc thù. Thys và cộng sự [15] đã tạo ra một miếng dán in được, khi cầm trước người thì bộ phát hiện người YOLOv2 không còn phát hiện ra người đó (Hình 6.2). Kỹ thuật tương tự có thể được dùng để làm mô hình đếm người đếm thiếu, khiến hệ thống không cảnh báo quá tải.")
FIG("Tấn công đối kháng vào bộ phát hiện người: người bên trái được phát hiện (khung “person”), người bên phải cầm miếng dán đối kháng thì không bị phát hiện.",
    "Nguồn: S. Thys, W. Van Ranst, T. Goedemé, “Fooling automated surveillance cameras: adversarial patches to attack person detection”, CVPR Workshops 2019 (arXiv:1904.08653), Fig. 1 [15].", existing="thys_fig1")
H2("6.3. Biện pháp bảo vệ khi triển khai thực tế")
BUL(["**Thiết bị và mạng:** đổi mật khẩu mặc định, tắt dịch vụ không dùng (Telnet, UPnP), cập nhật firmware; tách mạng camera thành VLAN riêng, không mở trực tiếp ra Internet, dùng VPN cho truy cập từ xa; giám sát nhật ký và lưu lượng mạng (IDS/SIEM).",
     "**Dữ liệu khi truyền và lưu trữ:** mã hóa luồng video (RTSP qua TLS/SRTP) và kênh gửi số liệu (MQTT/HTTPS qua TLS); xác thực hai chiều (mTLS) giữa thiết bị và máy chủ; gắn dấu thời gian và mã xác thực (HMAC/chữ ký số) cho số liệu, clip sự kiện để phát hiện sửa đổi, phát lại; mã hóa dữ liệu lưu trữ và tự động xóa khi hết hạn.",
     "**Mô hình AI:** phát hiện camera bị che, rung, đổi góc hoặc hình ảnh “đứng hình”; đối chiếu chéo số đếm giữa các camera chồng lấn hoặc với dữ liệu cổng soát vé; huấn luyện đối kháng; kiểm soát nguồn gốc dữ liệu huấn luyện và kiểm tra mã băm (SHA-256) của file trọng số trước khi nạp.",
     "**Kiểm soát truy cập và truy vết:** phân quyền theo vai trò (RBAC) — người xem số liệu không được xem video gốc; xác thực đa yếu tố cho tài khoản quản trị; ghi nhật ký mọi thao tác xem, xuất, xóa dữ liệu và xử lý cảnh báo (chống chối bỏ).",
     "**Con người trong vòng quyết định (human-in-the-loop):** AI chỉ cảnh báo và khuyến nghị, người vận hành xác nhận trước khi hành động."])
H2("6.4. Quyền riêng tư và tuân thủ pháp luật")
P("Hình ảnh của một người là dữ liệu cá nhân. Tại Việt Nam, Luật Bảo vệ dữ liệu cá nhân số 91/2025/QH15 được Quốc hội thông qua ngày 26/6/2025 và có hiệu lực từ ngày 01/01/2026 [16], được hướng dẫn bởi Nghị định 356/2025/NĐ-CP [22]. Luật ghi nhận các quyền của chủ thể dữ liệu (được biết, đồng ý, truy cập, chỉnh sửa, yêu cầu xóa…) và nghiêm cấm mua bán dữ liệu cá nhân. Hệ thống cần tuân thủ ngay từ thiết kế:")
BUL(["**Tối thiểu hóa dữ liệu:** hệ thống chỉ cần số lượng và mật độ, không cần danh tính → không dùng nhận dạng khuôn mặt; thiết bị xử lý chỉ gửi số liệu ẩn danh về trung tâm.",
     "**Ẩn danh hóa:** làm mờ hình ảnh khi hiển thị cho người không có thẩm quyền.",
     "**Minh bạch:** đặt biển thông báo khu vực có camera giám sát và mục đích sử dụng; công bố chính sách lưu trữ.",
     "**Giới hạn mục đích và thời hạn:** dữ liệu chỉ dùng cho an toàn đám đông; xóa theo thời hạn đã công bố.",
     "**Đánh giá tác động** xử lý dữ liệu cá nhân trước khi triển khai và định kỳ, theo quy định của pháp luật."])
H2("6.5. Các biện pháp đã cài đặt trong ứng dụng demo")
TABLE("Biện pháp an toàn thông tin và quyền riêng tư trong mã nguồn",
      ["Biện pháp", "Cách cài đặt", "Mối đe dọa giảm thiểu"],
      [["Không nhận dạng danh tính", "Mô hình chỉ xuất bản đồ mật độ; không có mô-đun khuôn mặt, không lưu vị trí từng người", "Lộ thông tin, lạm dụng giám sát"],
       ["Xử lý trong bộ nhớ, không lưu ảnh", "Ảnh được giải mã trực tiếp trong bộ nhớ; video tải lên ghi ra file tạm và bị xóa ngay trong khối finally sau khi xử lý", "Lộ thông tin"],
       ["Chỉ lưu số liệu", "Nhật ký và file CSV chỉ chứa thời điểm, số người, mật độ, mức cảnh báo", "Lộ thông tin; tối thiểu hóa dữ liệu"],
       ["Chế độ riêng tư", "Làm mờ Gauss ảnh hiển thị (bật mặc định), chỉ bản đồ nhiệt thể hiện rõ", "Lộ thông tin khi trình chiếu"],
       ["Nạp trọng số an toàn", "torch.load(…, weights_only=True): chỉ giải nén tensor và kiểu cơ bản, không thực thi mã pickle tùy ý", "Sửa đổi: file trọng số độc hại"],
       ["Chống cảnh báo chập chờn", "Làm trơn EMA và xác nhận K lần liên tiếp trước khi đổi mức", "Nhiễu mô hình, làm người vận hành bỏ qua cảnh báo"],
       ["Con người quyết định", "Ứng dụng chỉ hiển thị mức và hành động đề xuất, không tự động điều khiển thiết bị", "Sai lệch do tấn công đối kháng"]],
      [2.3, 4.6, 2.6])

# ---------------------------------------------------------------- CHƯƠNG 7
H1("CHƯƠNG 7. KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN")
H2("7.1. Kết quả đạt được")
P("Đề tài đã xây dựng hoàn chỉnh một pipeline từ dữ liệu đến ứng dụng cho bài toán đếm người và phát hiện đám đông bất thường: chuẩn bị bộ dữ liệu ShanghaiTech Part B và sinh bản đồ mật độ bảo toàn số người; huấn luyện mạng CSRNet " + ("đạt MAE " + vn("%.2f" % RES["test_mae"]) + ", RMSE " + vn("%.2f" % RES["test_rmse"]) + " trên tập test — thấp hơn kết quả công bố của bài báo gốc (10,6 / 16,0) nhờ cải tiến quy trình huấn luyện" if RES else "và đánh giá trên tập test") + "; xây dựng mô-đun cảnh báo bốn mức dựa trên mật độ người/m² và tăng đột biến, có cơ chế chống cảnh báo chập chờn; tích hợp thành ứng dụng demo xử lý ảnh, video và luồng camera. Về an toàn thông tin, báo cáo phân tích mối đe dọa theo STRIDE, đề xuất biện pháp khi triển khai và cài đặt các biện pháp bảo vệ quyền riêng tư ngay trong ứng dụng, phù hợp với Luật Bảo vệ dữ liệu cá nhân 2025.")
H2("7.2. Hạn chế")
BUL(["Mô hình chỉ huấn luyện trên một bộ dữ liệu (đường phố Thượng Hải); chưa đánh giá trên dữ liệu camera tại Việt Nam, ban đêm hay thời tiết xấu.",
     "Thử nghiệm ngoài bộ dữ liệu (mục 4.7) cho thấy mô hình bỏ sót người ở cận cảnh và đếm nhầm một số vật thể nền khi không có người.",
     "Kết quả vượt bài báo gốc mới từ một lần huấn luyện, chưa có thí nghiệm lặp lại và ablation.",
     "Diện tích vùng quan sát nhập thủ công và giả định mật độ đều trong khung hình; chưa dùng hiệu chỉnh phối cảnh.",
     "Chỉ phát hiện bất thường về mật độ và tốc độ tăng số người; chưa phát hiện bất thường về hành vi (tháo chạy, ẩu đả, dòng ngược chiều).",
     "Chưa thử nghiệm độ bền của mô hình trước tấn công đối kháng."])
H2("7.3. Hướng phát triển")
BUL(["Tinh chỉnh mô hình trên các bộ dữ liệu lớn hơn (NWPU-Crowd, JHU-Crowd++) và trên ảnh gán nhãn từ camera thực tế.",
     "Dùng homography để tính diện tích theo từng điểm ảnh, cho mật độ cục bộ chính xác và phát hiện điểm nóng trong khung hình.",
     "Bổ sung nhánh phân tích chuyển động bằng luồng quang học và các mô hình phát hiện bất thường hành vi [11], [12].",
     "Triển khai trên thiết bị biên (xuất ONNX/TensorRT), gửi số liệu về trung tâm qua kênh mã hóa TLS, quản lý nhiều camera.",
     "Thử nghiệm tấn công đối kháng lên mô hình đếm và áp dụng huấn luyện đối kháng."])

# ---------------------------------------------------------------- TÀI LIỆU THAM KHẢO
H1("TÀI LIỆU THAM KHẢO", numbered=False)
REFS = [
    "“Seoul Halloween crowd crush,” Wikipedia. [Trực tuyến]. https://en.wikipedia.org/wiki/Seoul_Halloween_crowd_crush (truy cập 12/10/2026).",
    "“Love Parade disaster,” Wikipedia. [Trực tuyến]. https://en.wikipedia.org/wiki/Love_Parade_disaster (truy cập 12/10/2026).",
    "“2015 Mina crowd crush,” Wikipedia. [Trực tuyến]. https://en.wikipedia.org/wiki/2015_Mina_crowd_crush (truy cập 12/10/2026).",
    "G. K. Still, Introduction to Crowd Science. Boca Raton, FL, USA: CRC Press, 2014.",
    "J. J. Fruin, Pedestrian Planning and Design. New York, NY, USA: Metropolitan Association of Urban Designers and Environmental Planners, 1971.",
    "Y. Zhang, D. Zhou, S. Chen, S. Gao, and Y. Ma, “Single-image crowd counting via multi-column convolutional neural network,” in Proc. IEEE CVPR, 2016, pp. 589–597.",
    "Y. Li, X. Zhang, and D. Chen, “CSRNet: Dilated convolutional neural networks for understanding the highly congested scenes,” in Proc. IEEE CVPR, 2018, pp. 1091–1100. arXiv:1802.10062.",
    "G. Gao, J. Gao, Q. Liu, Q. Wang, and Y. Wang, “CNN-based density estimation and crowd counting: A survey,” arXiv:2003.12783, 2020.",
    "K. Simonyan and A. Zisserman, “Very deep convolutional networks for large-scale image recognition,” in Proc. ICLR, 2015. arXiv:1409.1556.",
    "F. Yu and V. Koltun, “Multi-scale context aggregation by dilated convolutions,” in Proc. ICLR, 2016. arXiv:1511.07122.",
    "W. Liu, W. Luo, D. Lian, and S. Gao, “Future frame prediction for anomaly detection – A new baseline,” in Proc. IEEE CVPR, 2018. arXiv:1712.09867.",
    "W. Sultani, C. Chen, and M. Shah, “Real-world anomaly detection in surveillance videos,” in Proc. IEEE CVPR, 2018. arXiv:1801.04264.",
    "D. P. Kingma and J. Ba, “Adam: A method for stochastic optimization,” in Proc. ICLR, 2015. arXiv:1412.6980.",
    "M. Antonakakis et al., “Understanding the Mirai botnet,” in Proc. 26th USENIX Security Symposium, 2017, pp. 1093–1110.",
    "S. Thys, W. Van Ranst, and T. Goedemé, “Fooling automated surveillance cameras: Adversarial patches to attack person detection,” in Proc. IEEE CVPR Workshops, 2019. arXiv:1904.08653.",
    "Quốc hội nước CHXHCN Việt Nam, Luật Bảo vệ dữ liệu cá nhân, Luật số 91/2025/QH15, ngày 26/6/2025 (hiệu lực từ 01/01/2026).",
    "H. Idrees et al., “Composition loss for counting, density map estimation and localization in dense crowds,” in Proc. ECCV, 2018.",
    "V. A. Sindagi, R. Yasarla, and V. M. Patel, “JHU-CROWD++: Large-scale crowd counting dataset and a benchmark method,” IEEE Trans. Pattern Anal. Mach. Intell., 2020.",
    "Q. Wang, J. Gao, W. Lin, and X. Li, “NWPU-Crowd: A large-scale benchmark for crowd counting and localization,” IEEE Trans. Pattern Anal. Mach. Intell., vol. 43, no. 6, pp. 2141–2149, 2021. arXiv:2001.03360.",
    "A. Paszke et al., “PyTorch: An imperative style, high-performance deep learning library,” in Proc. NeurIPS, 2019.",
    "Streamlit, “Streamlit documentation.” [Trực tuyến]. https://docs.streamlit.io (truy cập 02/11/2026).",
    "Chính phủ nước CHXHCN Việt Nam, Nghị định 356/2025/NĐ-CP quy định chi tiết và hướng dẫn thi hành Luật Bảo vệ dữ liệu cá nhân, ngày 31/12/2025.",
    "Q. Song et al., “Rethinking counting and localization in crowds: A purely point-based framework,” in Proc. IEEE ICCV, 2021. arXiv:2107.12746.",
]
for i, r in enumerate(REFS, 1):
    add(f'<w:p {W_NS}><w:pPr><w:spacing w:after="80" w:line="276" w:lineRule="auto"/><w:ind w:left="567" w:hanging="567"/>'
        f'<w:jc w:val="both"/></w:pPr>{runs(f"[{i}]  {r}")}</w:p>')
add(f'<w:p {W_NS}><w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr>{runs("Nguồn ảnh Wikimedia Commons (truy cập 12/10/2026):")}</w:p>')
BUL(["https://commons.wikimedia.org/wiki/File:Itaewon_(%EC%9D%B4%ED%83%9C%EC%9B%90)_2022_Halloween.jpg — CC BY 3.0",
     "https://commons.wikimedia.org/wiki/File:Itaewon_Halloween_crowd_crush_incident_place_on_November_1,_2022.jpg — CC BY-SA 3.0",
     "https://commons.wikimedia.org/wiki/File:The_Kaaba_during_Hajj_-_edited.jpg — CC BY-SA 4.0",
     "https://commons.wikimedia.org/wiki/File:CCTV_camera_on_Round_Church_Street,_Cambridge.jpg — Public domain"])

# ================================================================ danh mục hình, bảng, từ viết tắt
FRONT = []


def front(xml):
    FRONT.append(parse_xml(xml))


def list_heading(text):
    front(f'<w:p {W_NS}><w:pPr><w:pStyle w:val="Heading1"/><w:pageBreakBefore/></w:pPr>{runs(text)}</w:p>')


def list_entry(text):
    front(f'<w:p {W_NS}><w:pPr><w:spacing w:after="60" w:line="276" w:lineRule="auto"/><w:ind w:left="1134" w:hanging="1134"/></w:pPr>'
          f'<w:r><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>')


list_heading("DANH MỤC HÌNH ẢNH")
for f in FIGS:
    list_entry(f)
front(f'<w:p {W_NS}><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>{runs("DANH MỤC BẢNG")}</w:p>')
for t in TABS:
    list_entry(t)
front(f'<w:p {W_NS}><w:pPr><w:pStyle w:val="Heading1"/><w:pageBreakBefore/></w:pPr>{runs("DANH MỤC TỪ VIẾT TẮT")}</w:p>')
CONTENT_SAVE = CONTENT
CONTENT = []
TABLE(None, ["Từ viết tắt", "Ý nghĩa"],
      [["AMP", "Automatic Mixed Precision – huấn luyện độ chính xác hỗn hợp"],
       ["CCTV", "Closed-Circuit Television – hệ thống camera giám sát"],
       ["CNN", "Convolutional Neural Network – mạng nơ-ron tích chập"],
       ["CSRNet", "Congested Scene Recognition Network"],
       ["EMA", "Exponential Moving Average – trung bình trượt hàm mũ"],
       ["GT", "Ground Truth – nhãn thực"],
       ["MAE", "Mean Absolute Error – sai số tuyệt đối trung bình"],
       ["MCNN", "Multi-Column Convolutional Neural Network"],
       ["RBAC", "Role-Based Access Control – kiểm soát truy cập theo vai trò"],
       ["RMSE", "Root Mean Squared Error – căn bậc hai sai số bình phương trung bình"],
       ["RTSP", "Real Time Streaming Protocol – giao thức truyền luồng video thời gian thực"],
       ["STRIDE", "Spoofing, Tampering, Repudiation, Information disclosure, Denial of service, Elevation of privilege"],
       ["TLS", "Transport Layer Security – giao thức bảo mật tầng giao vận"],
       ["VGG", "Mạng CNN của Visual Geometry Group (Đại học Oxford)"]],
      [2, 7])
FRONT += CONTENT
CONTENT = CONTENT_SAVE

for el in FRONT + CONTENT:
    SECT.addprevious(el)

# yêu cầu Word cập nhật mục lục khi mở file
settings = doc.settings.element
if not settings.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}updateFields"):
    settings.append(parse_xml(f'<w:updateFields {W_NS} w:val="true"/>'))

doc.save(DST)
print("Đã lưu:", DST)
print(f"{len(FIGS)} hình, {len(TABS)} bảng; kết quả huấn luyện: {'CÓ (' + OUTPUTS + ')' if RES else 'chưa có – các ô số liệu để …'}")
