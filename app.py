# V5.24.21 — HEADER FRAME + FINAL ANSWER MATRIX + STABLE PDF→MATHTYPE
# ============================================================
# - DOCX preserve: thêm lại khung Sở/Trường/Kỳ thi/Môn/Mã đề bằng XML trực tiếp, không qua python-docx save.
# - Cuối file Word có bảng đáp án tổng hợp tất cả mã đề, theo từng Phần I/II/III.
# - PDF→Word: gộp mảnh LaTeX liên tiếp và Toggle MathType theo batch nhỏ, mỗi batch mở/đóng Word riêng.
# - Word/PDF sau khi tạo được snapshot vào BytesIO rồi mới gửi về trình duyệt.
# - Có retry chờ Word/MathType/XeLaTeX/antivirus nhả handle trên Windows.
# - PDF dùng UUID trong thư mục TEMP, không ghi đè exports/De_... cố định.
# - File vật lý được xóa best-effort sau khi đã đọc vào RAM.
#
# V5.24.6 — ĐỔI TÊN THƯ MỤC + CHỈ SỐ CỰC TRỊ PDF + TOPBAR ĐỒNG BỘ
# ============================================================
# - Đổi tên thư mục Tài liệu/Ngân hàng mà giữ nguyên toàn bộ nội dung bên trong.
# - Nhận dạng CĐ/CT in nhỏ, hạ thấp trong PDF thành y_{CĐ}, y_{CT} chuẩn.
# - Ngày giờ, họ tên và công cụ topbar có ngay trong HTML lần tải đầu tiên.
# - Giữ cấu hình Docker/Render và dữ liệu bền vững để triển khai Internet.
#
# V5.24.9 — SỬA NHẬN PHẦN 1/2/3/4 + PDF LATEX MIKTEX/WINDOWS
# ============================================================
# V5.24.16 — WORD OOXML FALLBACK + WINDOWS TEMP EXPORT
# ============================================================
# - LibreOffice DOCX→HTML chạy bằng profile tạm riêng để tránh profile lock trên Windows.
# - Nếu LibreOffice vẫn thất bại, upload DOCX tự động chuyển sang parser OOXML trực tiếp; không dừng chức năng đảo đề.
# - File export Word trên Windows dùng thư mục TEMP hệ điều hành, tránh Permission denied trong thư mục dự án D:\\...\\exports.
# - DOCX nguồn vẫn được giữ nguyên để khi tải Word sau đảo, MathType OLE/hình/lời giải được clone trực tiếp từ OOXML gốc.
#
# V5.24.15 — WINDOWS EXPORT LOCK FIX
# ============================================================
# - Mỗi lần xuất Word dùng tên file vật lý duy nhất trong exports, không ghi đè file cũ đang bị Word/MathType/browser khóa.
# - Áp dụng cho cả Tất cả mã đề và từng mã đề.
# - Tên file người dùng tải về vẫn giữ nguyên, không lộ suffix tạm.
#
# V5.24.14 — WORD PREVIEW 1:1 + PDF→WORD MATHTYPE ONE-PASS
# ============================================================
# - Hình DOCX trên web dùng đúng ảnh LibreOffice đã render từ Word, không thay bằng EMF/WMF→SVG.
# - Geometry wp:extent/VML vẫn lấy từ OOXML gốc nên Câu 1/3/4 giữ đúng kích thước.
# - PDF→Word: Toggle TeX chỉ chạy MỘT lần; không còn lỗi toggle lần hai quay MathType về TeX.
# - VBScript chuyển từng vùng TeX từ cuối tài liệu về đầu và kiểm tra Equation.DSMT4 sau khi lưu.
#
# V5.24.13 — WORD FIGURES EXACT PREVIEW + MATHTYPE VIA WINDOWS COM WITHOUT PYWIN32
# - Khi đảo DOCX, giữ nguyên package Word, MathType OLE và quan hệ ảnh gốc.
# - Đồng bộ lại ShapeID/ObjectID khi clone hình/MathType để không lệch hình.
# - Thêm xuất PDF riêng từ dữ liệu LaTeX/MathJax thay vì convert Word MathType.
# - Thêm file Excel danh sách/tổng hợp học sinh luyện đề cho giáo viên.
#
# V5.24.7 — SỬA TRỘN WORD MATHTYPE + AI THEO YÊU CẦU
# ============================================================
# - XPath khi đảo/xuất DOCX dùng local-name(), không phụ thuộc prefix w:.
# - Trang xem tài liệu giữ cùng tab và hiển thị lại menu dọc bên trái.
# - Tạo đề AI gửi đúng yêu cầu/ma trận tới Responses API và kiểm tra cấu trúc.
# - Bổ sung cookie, proxy, header bảo mật và mật khẩu khởi tạo khi chạy production.
#
# V5.24.4 — PDF NHIỀU LỚP + XEM TÀI LIỆU CÙNG TAB + LỚP LUYỆN ĐỀ
# ============================================================
# - Tài liệu mở ngay trong tab hiện tại và bắt đầu từ đầu trang xem.
# - Theo dõi luyện đề hiển thị lớp suy ra từ tên thư mục Ngân hàng.
# - PDF Word nhiều lớp được hợp nhất: biến/số mũ/căn/ngoặc/α/Δ/ℝ thành LaTeX.
# - Loại ảnh mặt nạ công thức bị trùng, chỉ giữ ảnh minh họa thật của đề.
# - Giữ cấu hình Docker/Render và vùng dữ liệu bền vững để triển khai Internet.
#
# V5.24.3 — TOÁN TỬ CHUẨN + THEO DÕI LUYỆN ĐỀ THỜI GIAN THỰC
# ============================================================
# - Lim/min/max giữ kiểu toán tử và điều kiện nằm dưới như MathType nguồn.
# - Cực đại/cực tiểu y_CĐ, y_CT được dựng thành chỉ số dưới ngay ngắn.
# - Ngân hàng có Chọn tất cả và bảng theo dõi học sinh luyện đề trực tiếp.
# - Thống kê điểm, thời gian làm bài và số lần vi phạm theo từng lượt luyện.
# - Trang xem tài liệu chỉ còn nội dung tài liệu và thanh menu ngang.
#
# V5.24.2 — CHUẨN HÓA MIN/MAX/LIM VÀ BỐ CỤC CÔNG THỨC PDF
# ============================================================
# - Chuẩn hóa khóa phương án A/B/C/D trước khi ghi cơ sở dữ liệu.
# - Gộp an toàn các mảnh phương án bị parser Word/PDF nhận lặp khóa.
# - Không làm mất chữ, công thức hoặc hình ảnh trong mảnh bị lặp.
# - Giữ nguyên luồng thư mục, MathType/LaTeX/MathJax và chấm điểm.
#
# V5.24.0 — THƯ MỤC HAI TẦNG CHO TÀI LIỆU/NGÂN HÀNG/LUYỆN ĐỀ
# ============================================================
# - Thư mục hiển thị dạng ô chức năng và nằm giữa vùng nội dung.
# - Chỉ hiển thị danh sách tài liệu/đề sau khi người dùng mở thư mục.
# - Giáo viên upload ngay trong thư mục; học sinh duyệt cùng cấu trúc.
# - Giữ nguyên toàn bộ luồng Word/PDF, MathType, làm bài và chấm điểm.
#
# V5.23.9 — PDF NHIỀU TẦNG + THƯ MỤC TÀI LIỆU/NGÂN HÀNG
# ============================================================
# - Tách chính xác chữ khỏi hình vectơ ở câu có bảng/đồ thị đặt sát nội dung.
# - Khôi phục căn, phân số, tích phân nhiều tầng trong đề, đáp án và lời giải PDF.
# - Thêm thư mục riêng cho Tài liệu/Ngân hàng; Tài liệu đồng bộ sang học sinh.
# - Giữ nguyên toàn bộ luồng Word/MathType và giao diện thi đã ổn định.
#
# V5.23.8 — ĐÁP ÁN PDF ĐÚNG BỐ CỤC VÀ MENU Ở TÂM VÙNG NỘI DUNG
# ============================================================
# - Nhận diện dấu vectơ, dấu góc, tích phân và ngoặc co giãn theo font TeX.
# - Giữ các đồ thị/hình học/bảng biến thiên dạng vectơ dưới dạng ảnh sắc nét.
# - Sửa hệ phương trình nhiều cột và không để nhãn hình rơi vào phương án.
# - Giữ nguyên toàn bộ luồng Word/MathType và lời giải chi tiết đã ổn định.
#
# V5.23.5 — LỜI GIẢI CHI TIẾT THEO TỪNG CÂU WORD/PDF
# - Tách nội dung sau nhãn "Lời giải" và lưu riêng theo từng câu.
# - Giáo viên bấm để xem lời giải; học sinh không nhận dữ liệu lời giải.
# - Đọc đáp án Phần I/II/III trực tiếp từ phần lời giải của đề có đáp án.
# - Sửa bộ đếm \left/\right để không loại nhầm công thức có \rightarrow.
#
# V5.23.4 — SỬA DỨT ĐIỂM CÔNG THỨC VECTƠ CÂU 11
# - Nhận cả hai thứ tự glyph vectơ của PDF: dấu mũi tên trước hoặc sau chữ.
# - Không dùng cặp \left...\right cho dấu độ dài/góc vectơ lấy từ PDF.
# - Tự sửa cả đề đã lưu, kể cả biểu thức góc và 120° từng bị tách thành 2 khối.
# - Giáo viên/học sinh dùng cùng một bộ lọc TeX an toàn nên hiển thị giống nhau.
# - Giữ nguyên toàn bộ luồng Word/MathType, hình và bố cục đã ổn định.
#
# V5.23.3 — HOÀN THIỆN BỐ CỤC VÀ NỘI DUNG ĐỀ PDF
# - Sửa công thức độ dài/góc giữa hai vectơ ở Câu 11.
# - Đưa hình Câu 12 về cuối đề bài, trước các phương án trả lời.
# - Khôi phục khoảng trắng bị mất theo tọa độ ký tự của PDF.
# - Phần II giữ đúng bốn ý Đúng/Sai như luồng Word.
#
# V5.23.2 — PDF VECTOR -> LATEX NGẦM -> MATHJAX
# - Giữ nguyên hoàn toàn luồng Word/MathType đã ổn định.
# - PDF số được dựng lại công thức từ font, glyph, tọa độ và đường phân số.
# - Phân số, căn, số mũ/chỉ số, tích phân, vectơ và hệ phương trình thành LaTeX.
# - Chỉ hình minh họa thật được giữ nguyên bytes và tỷ lệ từ PDF.
# - Giáo viên/học sinh chỉ thấy MathJax dàn công thức, không thấy mã LaTeX.
#
# V5.23.1 — CỔNG TIỆN ÍCH + PDF WORD CHỈNH SỬA + TRỢ LÝ TOÁN
# - Tiện ích mở thành trang cổng giống khu vực Giáo viên/Học sinh.
# - PDF số được tách thành văn bản Word, công thức LaTeX và ảnh gốc.
# - Giáo viên AI trở thành trợ lý giải Toán có thể nhận chữ và ảnh.
# - Bổ sung ba trò chơi trí tuệ an toàn cho học sinh phổ thông.
# - Giữ nguyên toàn bộ luồng đề thi/MathType của V5.23.0.
#
# V5.23.0 — CHẾ ĐỘ LÀM BÀI TẬP TRUNG + TIỆN ÍCH + TRIỂN KHAI INTERNET
# - Ẩn menu dọc chỉ trên màn hình học sinh đang làm bài.
# - Thêm PDF sang Word, Giáo viên AI và Trò chơi học sinh.
# - Dữ liệu có thể chuyển sang ổ đĩa bền vững bằng APP_DATA_DIR.
# - Không thay đổi luồng MathType -> LaTeX -> MathJax của V5.22.14.
#
# V5.22.14 — GIỮ ẢNH CÂU HỎI QUA MỌI LẦN CẬP NHẬT
# - Ảnh câu hỏi mới được lưu trong uploads/question_media để đi cùng dữ liệu.
# - Vẫn đọc ảnh cũ trong static/question_media, không làm mất đề đã tạo.
# - Nếu URL ảnh cũ bị gãy, tự trích lại hình từ DOCX/PDF gốc trong uploads.
# - Không thay đổi luồng MathType -> LaTeX -> MathJax của V5.22.13.
#
# V5.22.13 — KHÔI PHỤC BỘ NẠP MATHJAX ĐÃ HOẠT ĐỘNG Ở V5.22.0
# - CDN jsDelivr là nguồn chính như V5.22.0; file cục bộ là nguồn dự phòng.
# - Bỏ pageReady tùy biến vì lỗi khởi động khiến trình duyệt lộ nguyên \(...\).
# - Chỉ gắn mathjax-ready sau khi startup.promise hoàn tất.
#
# V5.22.12 — CHỈ HIỆN SAU KHI LATEX ĐÃ ĐƯỢC DÀN THÀNH CÔNG THỨC
# - MathJax pageReady dàn xong toàn trang rồi mới mở vùng công thức.
# - Không lộ \dfrac, \sqrt hoặc dấu phân cách trong lúc trang đang tải.
# - Có timeout an toàn: MathJax lỗi thì không để vùng công thức trắng vĩnh viễn.
#
# V5.22.11 — MATHTYPE -> LATEX NGẦM -> CÔNG THỨC WEB ĐỒNG NHẤT
# - Công thức Word không còn được lưu/hiển thị dưới dạng ảnh PNG/SVG.
# - MTEF-py/OMML tạo LaTeX chuẩn; MathJax cục bộ dàn lại trên mọi màn hình.
# - Đề cũ tự bỏ fallback_url để không trộn ảnh lớn/nhỏ với công thức MathJax.
# - Giữ bộ sửa riêng Câu 10B và chuẩn hóa \dfrac, \sqrt.
# - Hình minh họa thật vẫn được khôi phục nguyên chất lượng và kích thước Word.
#
# V5.22.7 — SỬA CÔNG THỨC BỊ ẨN TRÊN MÀN HÌNH GIÁO VIÊN/HỌC SINH
# - Không còn phụ thuộc visibility:hidden và thời điểm gắn class math-rendered.
# - Chờ MathJax startup.promise trước khi dựng lại công thức chưa xử lý.
# - Đổi phiên bản URL CSS để trình duyệt không dùng bộ nhớ đệm của bản cũ.
#
# V5.22.6 — MATHTYPE -> LATEX NGẦM -> CÔNG THỨC WEB THUẦN
# - Khôi phục cơ chế MathJax tự dựng ổn định của V5.22.2.
# - Công thức không đi qua luồng ảnh và không lưu ảnh MathType dự phòng.
# - Giữ bộ sửa Câu 10B, \dfrac, \sqrt và chuẩn hóa ký hiệu MTEF.
#
# V5.22.5 — HIỂN THỊ MATHTYPE ỔN ĐỊNH TRÊN MỌI MÀN HÌNH
# - Chuẩn hóa \frac/\tfrac thành \dfrac và chuẩn hóa nhóm \sqrt{...}.
# - MathJax chủ động dựng lại công thức cho giáo viên/học sinh khi trang sẵn sàng.
# - Giữ ảnh preview MathType ẩn làm dự phòng; không còn vùng công thức trống.
#
# V5.22.4 — MTEF -> LATEX NGẦM -> CÔNG THỨC WEB
# - Mọi MathType/OMML trong DOCX được chuẩn hóa thành LaTeX không có $...$.
# - Giáo viên/học sinh chỉ thấy công thức MathJax đã dàn kiểu như MathType.
# - Sửa bộ giải mã prime Euclid Symbol và phân thức thiếu mẫu ở Câu 10B.
# - Chỉ hình minh họa thật được sao chép/xuất SVG và giữ kích thước Word.
#
# V5.22.3 — SỬA MATH INPUT ERROR + TOPBAR TẢI ĐỒNG THỜI
# - Loại ký tự Private Use do MathType chèn vào công thức A, B.
# - Không hiển thị MathJax merror; dùng ảnh MathType dự phòng.
# - Biểu tượng topbar được đóng gói cục bộ và đồng hồ khởi tạo tức thì.
#
# V5.22.2 — KHÔNG HIỂN THỊ MÃ LATEX RA MÀN HÌNH
# - LaTeX chỉ là dữ liệu ngầm; giao diện chỉ hiện công thức đã dựng.
# - Nếu MathJax gặp lỗi, tự dùng ảnh MathType gốc làm dự phòng.
#
# V5.22.1 — MENU MÀU + RENDER CÔNG THỨC CỤC BỘ
# - Thanh ngang có chọn màu, ngày giờ và họ tên người đăng nhập.
# - Chuẩn hóa dữ liệu cũ có $...$ trước khi MathJax dựng công thức.
# - Đóng gói MathJax trong dự án để không phụ thuộc CDN/Internet.
#
# V5.22.0 — MATHTYPE/OMML -> LATEX KHI UPLOAD WORD
# - Phần II trên màn hình làm bài: a), b), c), d).
# - option_key trong DB vẫn giữ A/B/C/D để tương thích chấm điểm.
# - Đây chỉ là thay đổi nhãn hiển thị, không thay đáp án hay scoring.
# ============================================================

import os, re, io, json, uuid, random, sqlite3, zipfile, shutil, subprocess, tempfile, copy, base64, time, threading
import urllib.request as urllib_request
import urllib.error as urllib_error
import urllib.parse as urllib_parse
from datetime import datetime, timedelta, timezone
from pathlib import Path
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, abort, send_from_directory, send_file
from werkzeug.utils import secure_filename
from werkzeug.middleware.proxy_fix import ProxyFix

# ============================================================
# V5.21.0 — Mammoth chỉ là lớp HTML preview cho DOCX.
# MathType OLE không được Mammoth chuyển trực tiếp thành LaTeX.
# Vì vậy hệ thống vẫn giữ parser Word/PDF cũ để thi online và chấm điểm.
# ============================================================
try:
    import mammoth
    MAMMOTH_AVAILABLE = True
except Exception:
    mammoth = None
    MAMMOTH_AVAILABLE = False

from werkzeug.security import generate_password_hash, check_password_hash
from lxml import etree
from PIL import Image, ImageChops
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from math_conversion import (
    extract_docx_equations,
    normalise_latex,
    replace_formula_images_with_latex,
)

BASE = Path(__file__).resolve().parent
DATA_ROOT = Path(os.environ.get("APP_DATA_DIR", str(BASE))).expanduser()
INSTANCE_DIR = DATA_ROOT / "instance"
UPLOAD_DIR = DATA_ROOT / "uploads"
LEGACY_MEDIA_DIR = BASE / "static" / "question_media"
# Ảnh câu hỏi là dữ liệu do người dùng tải lên, không phải tài nguyên mã nguồn.
# Lưu cùng uploads để khi nâng cấp ứng dụng, ảnh không bị mất theo thư mục static.
MEDIA_DIR = UPLOAD_DIR / "question_media"
DB = INSTANCE_DIR / "database.db"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
MEDIA_DIR.mkdir(parents=True, exist_ok=True)
LEGACY_MEDIA_DIR.mkdir(parents=True, exist_ok=True)
INSTANCE_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
APP_ENV = os.environ.get("APP_ENV", "development").strip().lower()
IS_PRODUCTION = APP_ENV == "production" or bool(os.environ.get("RENDER"))
SECRET_KEY = os.environ.get("SECRET_KEY", "").strip()
if IS_PRODUCTION and (
    len(SECRET_KEY) < 32 or SECRET_KEY == "change-this-secret-key"
):
    raise RuntimeError(
        "SECRET_KEY phải là chuỗi ngẫu nhiên tối thiểu 32 ký tự khi chạy production."
    )
app.secret_key = SECRET_KEY or "change-this-secret-key"
app.config.update(
    MAX_CONTENT_LENGTH=50 * 1024 * 1024,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=IS_PRODUCTION,
    PREFERRED_URL_SCHEME="https" if IS_PRODUCTION else "http",
)
if os.environ.get("TRUST_PROXY_HEADERS", "1" if IS_PRODUCTION else "0") == "1":
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
ALLOWED = {"docx", "pdf"}


@app.context_processor
def _initial_topbar_clock():
    """Render the clock with the first HTML response; JavaScript keeps it live."""
    now = datetime.now(timezone(timedelta(hours=7)))
    weekdays = (
        "Thứ Hai", "Thứ Ba", "Thứ Tư", "Thứ Năm",
        "Thứ Sáu", "Thứ Bảy", "Chủ Nhật",
    )
    return {
        "topbar_time": now.strftime("%H:%M"),
        "topbar_date": (
            f"{weekdays[now.weekday()]}, "
            f"{now.day:02d}/{now.month:02d}/{now.year:04d}"
        ),
    }


@app.before_request
def _verify_same_origin_for_writes():
    """Lightweight CSRF protection without changing the existing forms."""
    if request.method not in {"POST", "PUT", "PATCH", "DELETE"}:
        return None
    if app.testing or not IS_PRODUCTION:
        return None
    supplied = request.headers.get("Origin") or request.headers.get("Referer")
    if not supplied:
        abort(403)
    parsed = urllib_parse.urlsplit(supplied)
    if parsed.netloc.lower() != request.host.lower():
        abort(403)
    return None


@app.after_request
def _add_security_headers(response):
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    response.headers.setdefault(
        "Permissions-Policy", "camera=(), microphone=(), geolocation=()"
    )
    response.headers.setdefault(
        "Content-Security-Policy",
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net; "
        "style-src 'self' 'unsafe-inline'; "
        "img-src 'self' data: blob:; "
        "font-src 'self' data: https://cdn.jsdelivr.net; "
        "connect-src 'self'; frame-src 'self' blob:; "
        "frame-ancestors 'self'; base-uri 'self'; form-action 'self'",
    )
    if session.get("user_id"):
        response.headers.setdefault("Cache-Control", "no-store, private")
    if IS_PRODUCTION:
        response.headers.setdefault(
            "Strict-Transport-Security", "max-age=31536000; includeSubDomains"
        )
    return response


def _question_media_filename(url):
    """Return a safe stored filename from either the new or legacy media URL."""
    raw = str(url or "").split("?", 1)[0].replace("\\", "/")
    marker = "/question-media/"
    legacy_marker = "/static/question_media/"
    if marker in raw:
        name = raw.rsplit(marker, 1)[-1]
    elif legacy_marker in raw:
        name = raw.rsplit(legacy_marker, 1)[-1]
    else:
        return None
    name = Path(name).name
    return name if name and secure_filename(name) == name else None


def _question_media_path(url):
    """Resolve new persistent media first, then media created by older versions."""
    name = _question_media_filename(url)
    if not name:
        return None
    for folder in (MEDIA_DIR, LEGACY_MEDIA_DIR):
        candidate = folder / name
        if candidate.is_file():
            return candidate
    return None


def _question_media_url(filename):
    return url_for("question_media_file", filename=Path(filename).name)


def _may_read_question_media(filename):
    # Avoid exposing one teacher's question image to another logged-in
    # tenant merely because its unguessable file name was copied.
    uid=session.get('user_id'); role=session.get('role')
    if not uid or role not in ('teacher','student'): return False
    needle='%'+filename+'%'
    with db() as c:
        owners=[]
        for sql in (
          """SELECT DISTINCT e.teacher_id AS owner FROM questions q
             JOIN exams e ON e.id=q.exam_id WHERE q.blocks_json LIKE ? OR q.solution_json LIKE ?""",
          """SELECT DISTINCT e.teacher_id AS owner FROM options o
             JOIN questions q ON q.id=o.question_id JOIN exams e ON e.id=q.exam_id
             WHERE o.blocks_json LIKE ?""",
          """SELECT DISTINCT b.uploaded_by AS owner FROM bank_questions q
             JOIN bank_sources b ON b.id=q.source_id WHERE q.blocks_json LIKE ? OR q.solution_json LIKE ?""",
          """SELECT DISTINCT b.uploaded_by AS owner FROM bank_options o
             JOIN bank_questions q ON q.id=o.bank_question_id JOIN bank_sources b ON b.id=q.source_id
             WHERE o.blocks_json LIKE ?""",
          """SELECT DISTINCT b.uploaded_by AS owner FROM shuffle_questions q
             JOIN shuffle_sources b ON b.id=q.source_id WHERE q.blocks_json LIKE ? OR q.solution_json LIKE ?""",
          """SELECT DISTINCT b.uploaded_by AS owner FROM shuffle_options o
             JOIN shuffle_questions q ON q.id=o.shuffle_question_id JOIN shuffle_sources b ON b.id=q.source_id
             WHERE o.blocks_json LIKE ?""",
        ):
            params=(needle,needle) if sql.count('?')==2 else (needle,)
            try: owners.extend(r['owner'] for r in c.execute(sql,params).fetchall())
            except sqlite3.OperationalError: continue
        if role=='teacher': return int(uid) in owners
        return any(_is_teacher_student(c,owner,int(uid)) for owner in owners)


@app.get("/question-media/<path:filename>")
def question_media_file(filename):
    """Serve persistent images and transparently support old static images."""
    safe = Path(filename).name
    if safe != filename or not safe:
        abort(404)
    if not _may_read_question_media(safe): abort(404)
    if (MEDIA_DIR / safe).is_file():
        return send_from_directory(MEDIA_DIR, safe)
    if (LEGACY_MEDIA_DIR / safe).is_file():
        return send_from_directory(LEGACY_MEDIA_DIR, safe)
    abort(404)


@app.get("/static/question_media/<path:filename>")
def legacy_question_media_file(filename):
    """Keep every URL saved by V5.22.13 and earlier working."""
    return question_media_file(filename)


@app.template_filter("math_tex")
def math_tex(value):
    """Return canonical, delimiter-free TeX for MathJax only."""
    return _browser_safe_math_tex(value)


def _canonicalize_pdf_vector_marks(value):
    """Turn either PDF combining-arrow order into one canonical ``\\vec``."""
    latex=str(value or "")
    latex=re.sub(
        r"⃗\s*([A-Za-zΑ-ω])",
        lambda match: r"\vec{" + match.group(1) + "}",
        latex,
    )
    latex=re.sub(
        r"([A-Za-zΑ-ω])\s*⃗",
        lambda match: r"\vec{" + match.group(1) + "}",
        latex,
    )
    return latex


def _contains_vector_tex(value):
    """Nhận mọi dạng dấu véc-tơ mà hệ thống đang hỗ trợ."""
    return bool(re.search(r"\\(?:vec|overrightarrow)\b|⃗", str(value or "")))


def _browser_safe_math_tex(value):
    """Canonicalize vector delimiters before MathJax sees stored TeX.

    Some already-imported PDF questions contain a ``\\left`` half in one
    block and the matching ``\\right`` half in another block. MathJax then
    creates an merror even though the same source looks reasonable in the
    database. Vector length and vector-angle expressions do not need dynamic
    delimiters, so use the TeX-core ``|``/parentheses accepted by every bundled
    MathJax path. This is deliberately limited to formulas containing
    ``\\vec`` and therefore does not alter established Word/MathType formulas.
    """
    latex = _canonicalize_pdf_vector_marks(value)
    if _contains_vector_tex(latex):
        for old, new in (
            (r"\left\lvert", "|"),
            (r"\right\rvert", "|"),
            (r"\left|", "|"),
            (r"\right|", "|"),
            (r"\lvert", "|"),
            (r"\rvert", "|"),
            (r"\left(", "("),
            (r"\right)", ")"),
        ):
            latex = latex.replace(old, new)
    return normalise_latex(latex) or ""

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
V_NS = "urn:schemas-microsoft-com:vml"
O_NS = "urn:schemas-microsoft-com:office:office"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def db():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    c.execute("PRAGMA busy_timeout=15000")
    return c


def init_db():
    with db() as c:
        c.execute("PRAGMA journal_mode=WAL")
        c.executescript("""
        CREATE TABLE IF NOT EXISTS users(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('teacher','student')),
            full_name TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS exams(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            source_file TEXT,
            source_stored_name TEXT DEFAULT '',
            question_count INTEGER DEFAULT 0,
            published INTEGER DEFAULT 0,
            is_practice INTEGER NOT NULL DEFAULT 0,
            practice_source_id INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS questions(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_id INTEGER NOT NULL REFERENCES exams(id) ON DELETE CASCADE,
            original_order INTEGER NOT NULL,
            section_order INTEGER NOT NULL DEFAULT 1,
            blocks_json TEXT NOT NULL DEFAULT '[]',
            solution_json TEXT NOT NULL DEFAULT '[]',
            image_url TEXT,
            correct_option TEXT,
            UNIQUE(exam_id, original_order)
        );
        CREATE TABLE IF NOT EXISTS options(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            question_id INTEGER NOT NULL REFERENCES questions(id) ON DELETE CASCADE,
            option_key TEXT NOT NULL,
            blocks_json TEXT NOT NULL DEFAULT '[]',
            image_url TEXT,
            UNIQUE(question_id, option_key)
        );
        CREATE TABLE IF NOT EXISTS student_exams(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_id INTEGER NOT NULL REFERENCES exams(id) ON DELETE CASCADE,
            student_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            question_order_json TEXT NOT NULL,
            option_map_json TEXT NOT NULL,
            started_at TEXT DEFAULT CURRENT_TIMESTAMP,
            last_activity_at TEXT DEFAULT CURRENT_TIMESTAMP,
            submitted_at TEXT,
            score REAL
        );
        CREATE TABLE IF NOT EXISTS student_answers(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_exam_id INTEGER NOT NULL REFERENCES student_exams(id) ON DELETE CASCADE,
            question_id INTEGER NOT NULL REFERENCES questions(id) ON DELETE CASCADE,
            selected_original_key TEXT,
            is_correct INTEGER,
            UNIQUE(student_exam_id, question_id)
        );
        CREATE TABLE IF NOT EXISTS anti_cheat_logs(id INTEGER PRIMARY KEY AUTOINCREMENT, student_exam_id INTEGER NOT NULL, event_type TEXT, details TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS exam_settings(id INTEGER PRIMARY KEY AUTOINCREMENT, exam_id INTEGER UNIQUE NOT NULL, answer_key_json TEXT DEFAULT '{}');
        """)
        # Robust migration for databases created by V1/V2/V3.
        # Older databases may contain NOT NULL stem_json/content_json columns.
        # Newer code uses blocks_json, so we keep both columns synchronized.
        cols = {r[1] for r in c.execute("PRAGMA table_info(questions)").fetchall()}
        if "stem_json" not in cols:
            c.execute("ALTER TABLE questions ADD COLUMN stem_json TEXT NOT NULL DEFAULT '[]'")
            cols.add("stem_json")
        if "blocks_json" not in cols:
            c.execute("ALTER TABLE questions ADD COLUMN blocks_json TEXT NOT NULL DEFAULT '[]'")
            cols.add("blocks_json")
        if "stem_json" in cols and "blocks_json" in cols:
            c.execute("UPDATE questions SET blocks_json=stem_json WHERE (blocks_json IS NULL OR blocks_json='' OR blocks_json='[]') AND stem_json IS NOT NULL AND stem_json!='[]'")
            c.execute("UPDATE questions SET stem_json=blocks_json WHERE (stem_json IS NULL OR stem_json='') AND blocks_json IS NOT NULL")

        ocols = {r[1] for r in c.execute("PRAGMA table_info(options)").fetchall()}
        if "content_json" not in ocols:
            c.execute("ALTER TABLE options ADD COLUMN content_json TEXT NOT NULL DEFAULT '[]'")
            ocols.add("content_json")
        if "blocks_json" not in ocols:
            c.execute("ALTER TABLE options ADD COLUMN blocks_json TEXT NOT NULL DEFAULT '[]'")
            ocols.add("blocks_json")
        if "content_json" in ocols and "blocks_json" in ocols:
            c.execute("UPDATE options SET blocks_json=content_json WHERE (blocks_json IS NULL OR blocks_json='' OR blocks_json='[]') AND content_json IS NOT NULL AND content_json!='[]'")
            c.execute("UPDATE options SET content_json=blocks_json WHERE (content_json IS NULL OR content_json='') AND blocks_json IS NOT NULL")
        # Section-local numbering: every section starts at Câu 1.
        qcols = {r[1] for r in c.execute("PRAGMA table_info(questions)").fetchall()}
        if "section_order" not in qcols:
            c.execute("ALTER TABLE questions ADD COLUMN section_order INTEGER NOT NULL DEFAULT 1")
            # Backfill legacy rows by their old source order within each section.
            for sec in (1,2,3,4):
                rows = c.execute("SELECT id FROM questions WHERE section=? ORDER BY original_order,id", (sec,)).fetchall()
                for idx, row in enumerate(rows, 1):
                    c.execute("UPDATE questions SET section_order=? WHERE id=?", (idx, row[0]))

        for table, col, spec in [
            ('users','student_code','TEXT'),('users','class_name','TEXT'),('users','excel_row','INTEGER'),
            ('exams','source_stored_name',"TEXT DEFAULT ''"),('exams','is_practice','INTEGER NOT NULL DEFAULT 0'),('exams','practice_source_id','INTEGER'),('exams','part2_scoring_mode','INTEGER DEFAULT 1'),('exams','duration_minutes','INTEGER DEFAULT 90'),('exams','part1_points','REAL DEFAULT 2.5'),('exams','part2_points','REAL DEFAULT 2.5'),('exams','part3_points','REAL DEFAULT 2.5'),('exams','part4_points','REAL DEFAULT 2.5'),('exams','p2_item_a','REAL DEFAULT 0.625'),('exams','p2_item_b','REAL DEFAULT 0.625'),('exams','p2_item_c','REAL DEFAULT 0.625'),('exams','p2_item_d','REAL DEFAULT 0.625'),('exams','anti_cheat','INTEGER DEFAULT 1'),('exams','max_violations','INTEGER DEFAULT 3'),
            ('questions','section','INTEGER DEFAULT 1'),('questions','qtype',"TEXT DEFAULT 'mcq'"),('questions','short_answer','TEXT'),('questions','tf_answers_json',"TEXT DEFAULT '[]'"),('questions','solution_json',"TEXT NOT NULL DEFAULT '[]'"),('questions','manual_score','REAL'),('questions','points','REAL DEFAULT 0'),
            ('options','correct_tf','INTEGER'),('student_exams','violations','INTEGER DEFAULT 0'),('student_exams','last_activity_at','TEXT'),('student_answers','answer_json','TEXT'),('student_answers','score','REAL DEFAULT 0')]:
            cols={r[1] for r in c.execute(f'PRAGMA table_info({table})').fetchall()}
            if col not in cols: c.execute(f'ALTER TABLE {table} ADD COLUMN {col} {spec}')

        c.execute(
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_exams_practice_source "
            "ON exams(practice_source_id) "
            "WHERE is_practice=1 AND practice_source_id IS NOT NULL"
        )


        # ============================================================

        # ============================================================
        # V5.21.2 — TRẠNG THÁI KIỂM TRA BẢN NHÁP NGÂN HÀNG ĐỀ
        # Lưu trong DB để không phụ thuộc session/cookie.
        # ============================================================
        try:
            c.execute(
                "ALTER TABLE exams ADD COLUMN draft_checked INTEGER DEFAULT 0"
            )
        except sqlite3.OperationalError:
            pass

        # V5.13.0 — KHỐI TÀI KHOẢN ĐĂNG KÝ + QUẢN TRỊ
        # Chỉ BỔ SUNG cột/bảng mới, không thay đổi các chức năng V5.12.6.
        # ============================================================
        user_cols = {r[1] for r in c.execute("PRAGMA table_info(users)").fetchall()}
        for col, spec in [
            ("email", "TEXT"),
            ("phone", "TEXT"),
            ("is_admin", "INTEGER NOT NULL DEFAULT 0")
        ]:
            if col not in user_cols:
                c.execute(f"ALTER TABLE users ADD COLUMN {col} {spec}")
                user_cols.add(col)

        c.execute("""
            CREATE TABLE IF NOT EXISTS registration_requests(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                requested_role TEXT NOT NULL CHECK(requested_role IN ('teacher','student')),
                username TEXT NOT NULL,
                password_hash TEXT NOT NULL,
                full_name TEXT NOT NULL,
                email TEXT,
                phone TEXT,
                student_code TEXT,
                status TEXT NOT NULL DEFAULT 'pending'
                    CHECK(status IN ('pending','approved','rejected')),
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                decided_at TEXT,
                approved_user_id INTEGER
            )
        """)

        # ============================================================
        # V5.14.0 — GIỚI THIỆU WEBSITE + TÀI LIỆU GIÁO VIÊN
        # ============================================================
        c.execute("""
            CREATE TABLE IF NOT EXISTS site_settings(
                setting_key TEXT PRIMARY KEY,
                setting_value TEXT NOT NULL DEFAULT ''
            )
        """)
        # V5.15.1 — SQLite không nối chuỗi bằng cách đặt nhiều
        # literal cạnh nhau như Python. Dùng tham số để tránh syntax error.
        c.execute(
            """
            INSERT OR IGNORE INTO site_settings(setting_key,setting_value)
            VALUES(?,?)
            """,
            (
                "homepage_intro",
                "Chào mừng đến với Hệ thống Thi Trực Tuyến. "
                "Giáo viên có thể quản lý đề thi, tài liệu, kết quả; "
                "học sinh có thể làm bài và xem tài liệu học tập."
            )
        )

        c.execute("""
            CREATE TABLE IF NOT EXISTS document_folders(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL COLLATE NOCASE,
                created_by INTEGER NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(created_by,name)
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS documents(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                description TEXT,
                stored_name TEXT NOT NULL,
                original_name TEXT NOT NULL,
                uploaded_by INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        document_cols = {
            row[1] for row in c.execute("PRAGMA table_info(documents)").fetchall()
        }
        if "folder_id" not in document_cols:
            c.execute("ALTER TABLE documents ADD COLUMN folder_id INTEGER")
        c.execute(
            "CREATE INDEX IF NOT EXISTS idx_documents_folder_id "
            "ON documents(folder_id)"
        )


        # ============================================================
        # V5.17.0 — TIN TỨC TRANG CHỦ DO ADMIN QUẢN LÝ
        # ============================================================
        c.execute("""
            CREATE TABLE IF NOT EXISTS news_posts(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                content TEXT NOT NULL,
                image_name TEXT,
                video_name TEXT,
                created_by INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # ============================================================
        # V5.17.0 — ĐẢO / TRỘN ĐỀ ĐỘC LẬP
        # Không sử dụng/chỉnh sửa các bảng exams/questions/options cũ.
        # ============================================================
        c.execute("""
            CREATE TABLE IF NOT EXISTS shuffle_sources(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                source_file TEXT NOT NULL,
                uploaded_by INTEGER NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS shuffle_questions(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_id INTEGER NOT NULL,
                original_order INTEGER NOT NULL,
                section_order INTEGER NOT NULL,
                section INTEGER NOT NULL,
                qtype TEXT NOT NULL,
                blocks_json TEXT NOT NULL DEFAULT '[]',
                solution_json TEXT NOT NULL DEFAULT '[]',
                correct_option TEXT,
                short_answer TEXT,
                tf_answers_json TEXT NOT NULL DEFAULT '[]'
            )
        """)
        shuffle_qcols = {r[1] for r in c.execute("PRAGMA table_info(shuffle_questions)").fetchall()}
        if "solution_json" not in shuffle_qcols:
            c.execute("ALTER TABLE shuffle_questions ADD COLUMN solution_json TEXT NOT NULL DEFAULT '[]'")

        c.execute("""
            CREATE TABLE IF NOT EXISTS shuffle_options(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                shuffle_question_id INTEGER NOT NULL,
                option_key TEXT NOT NULL,
                blocks_json TEXT NOT NULL DEFAULT '[]',
                correct_tf INTEGER,
                UNIQUE(shuffle_question_id, option_key)
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS shuffle_versions(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_id INTEGER NOT NULL,
                version_json TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # ============================================================
        # V5.15.0 — NGÂN HÀNG ĐỀ/CÂU HỎI
        # Chỉ BỔ SUNG bảng mới; không sửa schema thi của V5.12.6.
        # ============================================================
        c.execute("""
            CREATE TABLE IF NOT EXISTS bank_folders(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL COLLATE NOCASE,
                created_by INTEGER NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(created_by,name)
            )
        """)

        c.execute("""
            CREATE TABLE IF NOT EXISTS bank_sources(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                source_file TEXT,
                stored_name TEXT DEFAULT '',
                uploaded_by INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        bank_source_cols = {
            row[1] for row in c.execute("PRAGMA table_info(bank_sources)").fetchall()
        }
        if "folder_id" not in bank_source_cols:
            c.execute("ALTER TABLE bank_sources ADD COLUMN folder_id INTEGER")
        c.execute(
            "CREATE INDEX IF NOT EXISTS idx_bank_sources_folder_id "
            "ON bank_sources(folder_id)"
        )

        c.execute("""
            CREATE TABLE IF NOT EXISTS bank_questions(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_id INTEGER NOT NULL,
                original_order INTEGER NOT NULL,
                section_order INTEGER NOT NULL,
                section INTEGER NOT NULL,
                qtype TEXT NOT NULL,
                blocks_json TEXT NOT NULL DEFAULT '[]',
                solution_json TEXT NOT NULL DEFAULT '[]',
                correct_option TEXT,
                short_answer TEXT,
                tf_answers_json TEXT NOT NULL DEFAULT '[]',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        """)
        bank_qcols = {r[1] for r in c.execute("PRAGMA table_info(bank_questions)").fetchall()}
        if "solution_json" not in bank_qcols:
            c.execute("ALTER TABLE bank_questions ADD COLUMN solution_json TEXT NOT NULL DEFAULT '[]'")

        c.execute("""
            CREATE TABLE IF NOT EXISTS bank_options(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bank_question_id INTEGER NOT NULL,
                option_key TEXT NOT NULL,
                blocks_json TEXT NOT NULL DEFAULT '[]',
                correct_tf INTEGER,
                UNIQUE(bank_question_id, option_key)
            )
        """)



        # Tạo tài khoản quản trị mặc định nếu hệ thống chưa có admin.
        # Admin dùng role='teacher' để không phá CHECK(role) cũ,
        # quyền admin được xác định riêng bằng is_admin=1.
        # Trên production bắt buộc lấy mật khẩu từ biến môi trường, tuyệt đối
        # không tạo tài khoản có mật khẩu công khai "admin123".
        admin = c.execute(
            "SELECT id FROM users WHERE is_admin=1 LIMIT 1"
        ).fetchone()
        if not admin:
            existing = c.execute(
                "SELECT id FROM users WHERE username=?",
                ("admin",)
            ).fetchone()
            if existing:
                c.execute(
                    "UPDATE users SET is_admin=1 WHERE id=?",
                    (existing["id"],)
                )
            else:
                bootstrap_password = os.environ.get(
                    "ADMIN_INITIAL_PASSWORD", ""
                ).strip()
                # A publicly known initial admin password is unsafe on a
                # multi-teacher site. Set ADMIN_INITIAL_PASSWORD explicitly.
                if not bootstrap_password:
                    app.logger.warning(
                        "Chưa tạo admin: cần ADMIN_INITIAL_PASSWORD trên production."
                    )
                else:
                    c.execute("""
                        INSERT INTO users(
                            username,password_hash,role,full_name,is_admin
                        ) VALUES(?,?,?,?,1)
                    """,(
                        "admin",
                        generate_password_hash(bootstrap_password),
                        "teacher",
                        "Quản trị viên"
                    ))

        # Tài khoản mẫu chỉ dành cho máy phát triển. Máy chủ thật mặc định tắt.
        enable_demo_accounts = os.environ.get(
            "ENABLE_DEMO_ACCOUNTS", "0"
        ) == "1"

        # KHỐI MIGRATION V5.11.1 — giữ dữ liệu cũ, mặc định dùng Phương án 1.
        # p2_item_a..d không còn quyết định điểm Phần II; điểm được tính theo
        # phương án của TỪNG CÂU Phần II. Các cột cũ vẫn giữ để tương thích DB.

        if enable_demo_accounts:
            for u, p, role, name in [
                ("teacher", "teacher123", "teacher", "Giáo viên Demo"),
                ("student", "student123", "student", "Học sinh Demo")
            ]:
                if not c.execute("SELECT 1 FROM users WHERE username=?", (u,)).fetchone():
                    c.execute(
                        "INSERT INTO users(username,password_hash,role,full_name) "
                        "VALUES(?,?,?,?)",
                        (u, generate_password_hash(p), role, name)
                    )

            # Học sinh demo dùng mã học sinh làm tài khoản/mật khẩu.
            demo = c.execute(
                "SELECT id FROM users WHERE student_code=?", ("2305089483",)
            ).fetchone()
            if demo:
                c.execute("""
                    UPDATE users
                    SET full_name=?,username=?,password_hash=?,role='student'
                    WHERE id=?
                """,(
                    "Học sinh Demo", "hs_2305089483",
                    generate_password_hash("2305089483"), demo[0]
                ))
            elif not c.execute(
                "SELECT 1 FROM users WHERE username=?", ("hs_2305089483",)
            ).fetchone():
                c.execute("""
                    INSERT INTO users(
                        username,password_hash,role,full_name,student_code
                    ) VALUES(?,?,?,?,?)
                """,(
                    "hs_2305089483", generate_password_hash("2305089483"),
                    "student", "Học sinh Demo", "2305089483"
                ))


        # V5.25.0: isolate every teacher's stored resources. Preserve old DB.
        init_teacher_space_schema(c)


def login_required(role=None):
    def deco(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            if "user_id" not in session:
                return redirect(url_for("login", next=request.path))
            if role and session.get("role") != role:
                abort(403)
            return fn(*args, **kwargs)
        return wrapper
    return deco


def init_teacher_space_schema(c):
    cols = {r[1] for r in c.execute("PRAGMA table_info(exams)")}
    if "teacher_id" not in cols:
        c.execute("ALTER TABLE exams ADD COLUMN teacher_id INTEGER REFERENCES users(id)")
    # Existing exams lack author information; attach bank/practice exams to
    # their bank owner. For otherwise unattributed legacy exams, use sole
    # non-admin teacher if unambiguous, else the admin for manual review.
    c.execute("""UPDATE exams SET teacher_id=(
        SELECT b.uploaded_by FROM bank_sources b
        WHERE b.id=exams.practice_source_id
    ) WHERE teacher_id IS NULL AND practice_source_id IS NOT NULL""")
    c.execute("""UPDATE exams SET teacher_id=(
        SELECT b.uploaded_by FROM bank_sources b
        WHERE exams.source_file='question_bank:' || b.id
    ) WHERE teacher_id IS NULL AND source_file LIKE 'question_bank:%'""")
    owners=c.execute("SELECT id FROM users WHERE role='teacher' AND COALESCE(is_admin,0)=0").fetchall()
    legacy_owner=owners[0][0] if len(owners)==1 else None
    if legacy_owner is None:
        admin=c.execute("SELECT id FROM users WHERE is_admin=1 ORDER BY id LIMIT 1").fetchone()
        legacy_owner=admin[0] if admin else None
    if legacy_owner is not None:
        c.execute("UPDATE exams SET teacher_id=? WHERE teacher_id IS NULL",(legacy_owner,))
    c.executescript("""
        CREATE TABLE IF NOT EXISTS teacher_students(
            teacher_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            student_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            class_name TEXT NOT NULL DEFAULT '',
            joined_at TEXT DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY(teacher_id,student_id)
        );
        CREATE TABLE IF NOT EXISTS teacher_invites(
            token TEXT PRIMARY KEY, teacher_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            class_name TEXT NOT NULL DEFAULT '', active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS exam_invites(
            token TEXT PRIMARY KEY, exam_id INTEGER NOT NULL REFERENCES exams(id) ON DELETE CASCADE,
            teacher_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            active INTEGER NOT NULL DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS exam_students(
            exam_id INTEGER NOT NULL REFERENCES exams(id) ON DELETE CASCADE,
            student_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            PRIMARY KEY(exam_id,student_id)
        );
        CREATE INDEX IF NOT EXISTS idx_exams_teacher ON exams(teacher_id);
        CREATE INDEX IF NOT EXISTS idx_students_teacher ON teacher_students(teacher_id);
        CREATE INDEX IF NOT EXISTS idx_students_student ON teacher_students(student_id);
    """)
    # Historical associations are imported ONCE only. Repeating this on every
    # restart would re-enroll a pupil after a teacher removed them.
    c.execute("CREATE TABLE IF NOT EXISTS app_migration_flags(name TEXT PRIMARY KEY)")
    done=c.execute("SELECT 1 FROM app_migration_flags WHERE name='teacher_membership_backfill_v525'").fetchone()
    if not done:
        c.execute("""INSERT OR IGNORE INTO teacher_students(teacher_id,student_id)
           SELECT DISTINCT e.teacher_id,se.student_id FROM student_exams se
           JOIN exams e ON e.id=se.exam_id
           WHERE e.teacher_id IS NOT NULL""")
        c.execute("INSERT INTO app_migration_flags(name) VALUES('teacher_membership_backfill_v525')")


def _is_teacher_student(c, teacher_id, student_id):
    return bool(c.execute("SELECT 1 FROM teacher_students WHERE teacher_id=? AND student_id=?",
                          (teacher_id,student_id)).fetchone())


def _may_read_document(c, doc_id, user_id, role):
    row=c.execute("SELECT uploaded_by FROM documents WHERE id=?",(doc_id,)).fetchone()
    if not row: return False
    if role=='teacher': return row['uploaded_by']==user_id
    return role=='student' and _is_teacher_student(c,row['uploaded_by'],user_id)


@app.before_request
def _teacher_space_gate():
    # Check ID-bearing routes before reaching old legacy handlers. Every
    # export, preview, update, grade and delete remains covered here.
    uid=session.get('user_id'); role=session.get('role')
    # Built-in Flask static routing may serve legacy images before the
    # custom /static/question_media handler, including anonymous visits.
    if request.path.startswith('/static/question_media/'):
        if not _may_read_question_media(Path(request.path).name): abort(404)
    # V5.24.x forgot an authentication check on /document/<id>/download.
    if request.path.startswith('/document/') and uid is None:
        abort(403)
    if uid is None: return None
    endpoint=request.endpoint or ''
    kw=request.view_args or {}
    path=request.path
    try: uid=int(uid)
    except (ValueError,TypeError): abort(403)
    with db() as c:
        if 'doc_id' in kw and endpoint in {
            'view_document','document_inline_file','document_preview_pdf','download_document',
            'teacher_delete_document'}:
            if not _may_read_document(c,kw['doc_id'],uid,role): abort(404)
        if path.startswith('/teacher/'):
            if role!='teacher': abort(403)
            if 'exam_id' in kw:
                row=c.execute('SELECT teacher_id FROM exams WHERE id=?',(kw['exam_id'],)).fetchone()
                if not row or row['teacher_id']!=uid: abort(404)
            if 'source_id' in kw:
                table='shuffle_sources' if path.startswith('/teacher/shuffle-exams/') else 'bank_sources'
                row=c.execute(f'SELECT uploaded_by FROM {table} WHERE id=?',(kw['source_id'],)).fetchone()
                if not row or row['uploaded_by']!=uid: abort(404)
            if 'folder_id' in kw:
                table='document_folders' if path.startswith('/teacher/documents/') else 'bank_folders'
                row=c.execute(f'SELECT created_by FROM {table} WHERE id=?',(kw['folder_id'],)).fetchone()
                if not row or row['created_by']!=uid: abort(404)
            if 'attempt_id' in kw:
                row=c.execute("""SELECT e.teacher_id FROM student_exams se
                    JOIN exams e ON e.id=se.exam_id WHERE se.id=?""",(kw['attempt_id'],)).fetchone()
                if not row or row['teacher_id']!=uid: abort(404)
            if 'student_exam_id' in kw:
                row=c.execute("""SELECT e.teacher_id FROM student_exams se
                    JOIN exams e ON e.id=se.exam_id WHERE se.id=?""",(kw['student_exam_id'],)).fetchone()
                if not row or row['teacher_id']!=uid: abort(404)
        elif path.startswith('/student/') and role=='student':
            if 'exam_id' in kw:
                row=c.execute('SELECT teacher_id,published FROM exams WHERE id=?',(kw['exam_id'],)).fetchone()
                if not row or not row['published'] or not _is_teacher_student(c,row['teacher_id'],uid): abort(404)
                # Explicit link distribution: if links were generated for this exam,
                # only pupils who accepted an exam link can start a new attempt.
                if c.execute('SELECT 1 FROM exam_invites WHERE exam_id=?',(kw['exam_id'],)).fetchone():
                    if not c.execute('SELECT 1 FROM exam_students WHERE exam_id=? AND student_id=?',
                                     (kw['exam_id'],uid)).fetchone(): abort(404)
            if 'source_id' in kw:
                row=c.execute('SELECT uploaded_by FROM bank_sources WHERE id=?',(kw['source_id'],)).fetchone()
                if not row or not _is_teacher_student(c,row['uploaded_by'],uid): abort(404)
            if 'folder_id' in kw:
                table='document_folders' if path.startswith('/student/documents/') else 'bank_folders'
                row=c.execute(f'SELECT created_by FROM {table} WHERE id=?',(kw['folder_id'],)).fetchone()
                if not row or not _is_teacher_student(c,row['created_by'],uid): abort(404)
            if 'student_exam_id' in kw:
                row=c.execute('''SELECT se.student_id,e.teacher_id FROM student_exams se
                     JOIN exams e ON e.id=se.exam_id WHERE se.id=?''',(kw['student_exam_id'],)).fetchone()
                if not row or row['student_id']!=uid: abort(404)
                if endpoint not in ('result',) and not _is_teacher_student(c,row['teacher_id'],uid): abort(404)
    return None


@app.get('/teacher/my-students')
@login_required('teacher')
def teacher_my_students():
    with db() as c:
        students=c.execute("""SELECT u.id,u.full_name,u.username,u.student_code,t.class_name,t.joined_at
          FROM teacher_students t JOIN users u ON u.id=t.student_id
          WHERE t.teacher_id=? ORDER BY t.class_name,u.full_name""",(session['user_id'],)).fetchall()
        invites=c.execute("""SELECT token,class_name,active FROM teacher_invites
          WHERE teacher_id=? ORDER BY created_at DESC""",(session['user_id'],)).fetchall()
    return render_template('teacher_my_students.html',students=students,invites=invites)


@app.post('/teacher/my-students/invite')
@login_required('teacher')
def teacher_create_invite():
    import secrets
    class_name=normalize_text(request.form.get('class_name'))[:80]
    with db() as c:
        token=secrets.token_urlsafe(24)
        c.execute('INSERT INTO teacher_invites(token,teacher_id,class_name) VALUES(?,?,?)',
                  (token,session['user_id'],class_name))
    flash('Đã tạo liên kết mời học sinh. Chỉ học sinh đăng nhập mới có thể tham gia.')
    return redirect(url_for('teacher_my_students'))


@app.post('/teacher/my-students/invite/<token>/revoke')
@login_required('teacher')
def teacher_revoke_invite(token):
    with db() as c:
        c.execute('UPDATE teacher_invites SET active=0 WHERE token=? AND teacher_id=?',
                  (token,session['user_id']))
    return redirect(url_for('teacher_my_students'))


@app.post('/teacher/my-students/<int:student_id>/remove')
@login_required('teacher')
def teacher_remove_student(student_id):
    with db() as c:
        c.execute('DELETE FROM teacher_students WHERE teacher_id=? AND student_id=?',
                  (session['user_id'],student_id))
        c.execute("""DELETE FROM exam_students WHERE student_id=? AND exam_id IN
          (SELECT id FROM exams WHERE teacher_id=?)""",(student_id,session['user_id']))
    return redirect(url_for('teacher_my_students'))


@app.route('/join/teacher/<token>',methods=['GET','POST'])
def join_teacher(token):
    with db() as c:
        invite=c.execute("""SELECT i.*,u.full_name AS teacher_name FROM teacher_invites i
           JOIN users u ON u.id=i.teacher_id WHERE i.token=? AND i.active=1""",(token,)).fetchone()
        if not invite: abort(404)
        if session.get('role')!='student':
            flash('Hãy đăng nhập bằng tài khoản học sinh để tham gia lớp.')
            return redirect(url_for('login',next=request.path))
        if request.method=='POST':
            c.execute("""INSERT INTO teacher_students(teacher_id,student_id,class_name)
              VALUES(?,?,?) ON CONFLICT(teacher_id,student_id) DO UPDATE SET
              class_name=excluded.class_name""",
              (invite['teacher_id'],session['user_id'],invite['class_name']))
            flash('Đã tham gia lớp của giáo viên.')
            return redirect(url_for('student_portal'))
    return render_template('join_teacher.html',invite=invite)


@app.post('/teacher/exam/<int:exam_id>/invite')
@login_required('teacher')
def create_exam_invite(exam_id):
    import secrets
    with db() as c:
        exam=c.execute('SELECT * FROM exams WHERE id=? AND teacher_id=?',
                       (exam_id,session['user_id'])).fetchone()
        if not exam: abort(404)
        if not exam['published']:
            flash('Hãy xuất bản đề trước khi tạo liên kết chia sẻ.')
            return redirect(url_for('teacher_dashboard'))
        token=secrets.token_urlsafe(24)
        c.execute('INSERT INTO exam_invites(token,exam_id,teacher_id) VALUES(?,?,?)',
                  (token,exam_id,session['user_id']))
    flash('Đã tạo liên kết bài kiểm tra, xem ở ô bên dưới đề.')
    return redirect(url_for('teacher_dashboard'))


@app.post('/teacher/exam/<int:exam_id>/invite/<token>/revoke')
@login_required('teacher')
def revoke_exam_invite(exam_id,token):
    with db() as c:
        c.execute('UPDATE exam_invites SET active=0 WHERE exam_id=? AND teacher_id=? AND token=?',
                  (exam_id,session['user_id'],token))
    return redirect(url_for('teacher_dashboard'))


@app.route('/join/exam/<token>',methods=['GET','POST'])
def join_exam(token):
    with db() as c:
        invite=c.execute("""SELECT i.*,e.title,e.published,u.full_name AS teacher_name
           FROM exam_invites i JOIN exams e ON e.id=i.exam_id
           JOIN users u ON u.id=i.teacher_id
           WHERE i.token=? AND i.active=1 AND e.teacher_id=i.teacher_id""",(token,)).fetchone()
        if not invite or not invite['published']: abort(404)
        if session.get('role')!='student':
            flash('Hãy đăng nhập bằng tài khoản học sinh để mở đề.')
            return redirect(url_for('login',next=request.path))
        if request.method=='POST':
            c.execute("""INSERT OR IGNORE INTO teacher_students(teacher_id,student_id)
                  VALUES(?,?)""",(invite['teacher_id'],session['user_id']))
            c.execute('INSERT OR IGNORE INTO exam_students(exam_id,student_id) VALUES(?,?)',
                      (invite['exam_id'],session['user_id']))
            return redirect(url_for('take_exam',exam_id=invite['exam_id']))
    return render_template('join_exam.html',invite=invite)


@app.context_processor
def _teacher_space_links():
    return {'my_teacher_id':session.get('user_id') if session.get('role')=='teacher' else None}



# ============================================================
# V5.21.0 — DOCX -> HTML BẰNG MAMMOTH (PREVIEW)
# ------------------------------------------------------------
# MỤC ĐÍCH:
# - Giữ cấu trúc đoạn/bảng/ảnh của Word tốt hơn text thuần.
# - Dùng cho giáo viên kiểm tra file Word đã upload.
#
# KHÔNG dùng hàm này để thay thế parser thi:
# - MathType OLE không tự biến thành LaTeX.
# - Hệ thống thi vẫn dùng extract_word_blocks()/build_questions()
#   để tách câu, trộn đề, chấm điểm và gắn nút chọn đáp án.
# ============================================================
def docx_to_mammoth_html(path):
    if not MAMMOTH_AVAILABLE:
        return None, (
            "Máy chưa cài thư viện mammoth. "
            "Chạy: pip install mammoth"
        )

    try:
        with open(path, "rb") as f:
            result = mammoth.convert_to_html(
                f,
                style_map="""
                    p[style-name='Normal'] => p.exam-word-paragraph:fresh
                    table => table.exam-word-table:fresh
                """
            )

        html = result.value or ""
        messages = [str(m.message) for m in result.messages]
        return html, messages
    except Exception as e:
        return None, str(e)


def normalize_text(s):
    return re.sub(r"[ \t\r\n]+", " ", (s or "").replace("\xa0", " ")).strip()


def ext(name):
    return name.rsplit(".", 1)[-1].lower() if "." in name else ""


def jdump(v):
    return json.dumps(v, ensure_ascii=False)


def find_soffice():
    """Find LibreOffice on Windows/Linux/macOS.

    Priority:
    1) SOFFICE_PATH environment variable
    2) PATH (soffice/libreoffice)
    3) common Windows installation folders
    4) Windows registry (if available)
    """
    candidates = []

    env_path = os.environ.get("SOFFICE_PATH", "").strip().strip('"')
    if env_path:
        candidates.append(env_path)

    for cmd in ("soffice", "libreoffice"):
        found = shutil.which(cmd)
        if found:
            candidates.append(found)

    # IMPORTANT: normal Python strings are used here intentionally.
    # The previous V3 used raw strings containing doubled backslashes,
    # which made Windows paths fail Path.exists().
    if os.name == "nt":
        candidates.extend([
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\LibreOffice\program\soffice.exe"),
            os.path.expandvars(r"%ProgramFiles%\LibreOffice\program\soffice.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\LibreOffice\program\soffice.exe"),
        ])

        # Registry lookup catches custom installation locations.
        try:
            import winreg
            for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                for subkey in (
                    r"SOFTWARE\LibreOffice\UNO\InstallPath",
                    r"SOFTWARE\WOW6432Node\LibreOffice\UNO\InstallPath",
                ):
                    try:
                        with winreg.OpenKey(hive, subkey) as key:
                            value, _ = winreg.QueryValueEx(key, None)
                            if value:
                                candidates.append(str(Path(value) / "soffice.exe"))
                    except OSError:
                        pass
        except Exception:
            pass

    seen = set()
    for candidate in candidates:
        if not candidate:
            continue
        try:
            candidate = str(Path(candidate).expanduser())
        except Exception:
            continue
        if candidate in seen:
            continue
        seen.add(candidate)
        if Path(candidate).is_file():
            return candidate

    return None


def convert_wmf_to_png(wmf_path, out_dir):
    """LibreOffice renders old WMF previews reliably; then crop white margins."""
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / (Path(wmf_path).stem + "_render.png")
    if target.exists():
        return target
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "Chưa tìm thấy LibreOffice/soffice. "
            "Hãy cài LibreOffice hoặc đặt biến môi trường SOFFICE_PATH trỏ tới "
            "soffice.exe (ví dụ C:\\Program Files\\LibreOffice\\program\\soffice.exe)."
        )
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        src = td / Path(wmf_path).name
        shutil.copy2(wmf_path, src)
        cmd = [soffice, "--headless", "--convert-to", "png", "--outdir", str(td), str(src)]
        p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60, text=True)
        rendered = td / (src.stem + ".png")
        if p.returncode != 0 or not rendered.exists():
            raise RuntimeError("LibreOffice không thể render công thức WMF.")
        im = Image.open(rendered).convert("RGB")
        bg = Image.new("RGB", im.size, "white")
        bbox = ImageChops.difference(im, bg).getbbox()
        if bbox:
            pad = 5
            bbox = (max(0,bbox[0]-pad), max(0,bbox[1]-pad), min(im.width,bbox[2]+pad), min(im.height,bbox[3]+pad))
            im = im.crop(bbox)
        # Keep enough pixels for mobile screens while avoiding huge files.
        if im.width < 120:
            scale = 120 / im.width
            im = im.resize((int(im.width*scale), int(im.height*scale)), Image.Resampling.LANCZOS)
        im.save(target, "PNG", optimize=True)
    return target


def copy_media(src, original_name, kind):
    src = Path(src)
    safe = secure_filename(src.name) or (uuid.uuid4().hex + ".png")
    out_name = f"{uuid.uuid4().hex}_{safe}"
    out = MEDIA_DIR / out_name
    if src.suffix.lower() == ".wmf":
        out = MEDIA_DIR / (Path(out_name).stem + ".png")
        rendered = convert_wmf_to_png(src, MEDIA_DIR / "_wmf_cache")
        shutil.copy2(rendered, out)
    else:
        shutil.copy2(src, out)
    return _question_media_url(out.name)


def _length_to_px(value):
    """KHỐI HÌNH ẢNH — đổi kích thước Word (pt/cm/in/px) sang pixel."""
    if not value: return None
    m=re.search(r'([-+]?\d+(?:\.\d+)?)\s*(pt|px|cm|mm|in)?', str(value).lower())
    if not m: return None
    n=float(m.group(1)); u=m.group(2) or 'px'
    return int(round(n * {'px':1,'pt':96/72,'in':96,'cm':96/2.54,'mm':96/25.4}[u]))

def _node_image_crop(node):
    """Return Word's source-picture crop (a:srcRect) as 1/100000 percentages.

    Word often stores an EMF/WMF with a large transparent canvas and uses
    ``a:srcRect`` to crop the canvas before drawing it.  LibreOffice exports
    the whole canvas to HTML, so applying only wp:extent makes the visible
    graph look much smaller.  Preserve the crop as well as the extent.
    """
    try:
        # DrawingML: a:blip is inside pic:blipFill; srcRect is its sibling.
        parent=node.getparent()
        for _ in range(5):
            if parent is None:
                break
            rect=parent.find(f'{{{A_NS}}}srcRect')
            if rect is not None:
                vals={}
                for key in ('l','t','r','b'):
                    try:
                        vals[key]=max(0,min(99999,int(rect.get(key,'0') or 0)))
                    except Exception:
                        vals[key]=0
                if any(vals.values()):
                    return vals
                return None
            parent=parent.getparent()
    except Exception:
        pass
    return None


def _word_preview_crop_matches_extent(url, crop, width_px, height_px):
    """Only crop an unmodified LO GIF when Word's crop and extent agree.

    This avoids accidentally cropping a preview that was already transformed
    by another renderer, while handling LO's full-canvas Picture GIFs.
    """
    path=_question_media_path(url)
    if not path or not path.is_file() or path.suffix.lower() != '.gif':
        return False
    if not crop or not width_px or not height_px:
        return False
    try:
        with Image.open(path) as source:
            w,h=source.size
        width_fraction=1.0-(crop.get('l',0)+crop.get('r',0))/100000.0
        height_fraction=1.0-(crop.get('t',0)+crop.get('b',0))/100000.0
        if width_fraction<=0 or height_fraction<=0:
            return False
        aspect_after=(w*width_fraction)/(h*height_fraction)
        aspect_word=width_px/float(height_px)
        return abs(aspect_after/aspect_word-1) < .055
    except (ValueError,TypeError,OSError):
        return False


def _crop_question_preview_to_word_src_rect(url, crop):
    """Crop a LibreOffice preview exactly like Word's ``a:srcRect``.

    The cropped raster is only for web preview / PDF-derived fallback.  The
    original DOCX package is never rewritten, so native Word/MathType media
    remain byte-for-byte intact when a DOCX source is shuffled.
    """
    if not url or not crop:
        return url
    src=_question_media_path(url)
    if not src or not src.is_file():
        return url
    try:
        with Image.open(src) as im0:
            rgba=im0.convert('RGBA')
            # LibreOffice GIF previews use a transparent magenta palette.
            # Flatten onto Word's white page before saving PNG so browsers or
            # downstream converters never expose the palette colour.
            white=Image.new('RGBA',rgba.size,(255,255,255,255))
            white.alpha_composite(rgba)
            im=white.convert('RGB')
            w,h=im.size
            l=int(round(w*float(crop.get('l',0))/100000.0))
            t=int(round(h*float(crop.get('t',0))/100000.0))
            r=int(round(w*(1.0-float(crop.get('r',0))/100000.0)))
            b=int(round(h*(1.0-float(crop.get('b',0))/100000.0)))
            l=max(0,min(w-1,l)); t=max(0,min(h-1,t))
            r=max(l+1,min(w,r)); b=max(t+1,min(h,b))
            # Ignore malformed/no-op crops.
            if l==0 and t==0 and r==w and b==h:
                return url
            cropped=im.crop((l,t,r,b))
            out=MEDIA_DIR/(uuid.uuid4().hex+'_wordcrop.png')
            cropped.save(out,'PNG',optimize=True)
        return _question_media_url(out.name)
    except Exception:
        app.logger.exception('Không thể áp dụng a:srcRect cho ảnh Word preview')
        return url




def _normalize_lo_figure_preview(url, word_crop=None):
    """Trim *only* LO's transparent margins on legacy Word figure GIF previews.

    A LibreOffice HTML preview has already had the document picture transform
    applied.  Applying Word's ``a:srcRect`` to that GIF a second time clips the
    graph.  The LO figure itself often occupies just half the transparent GIF
    canvas (see OTC-1-L12 Part I Q1/Q3/Q4).  Trim the actual alpha bounds and
    let OOXML's ``wp:extent`` determine its displayed size in the browser.

    Never touch MathType/OLE previews, original DOCX bytes, or non-GIF images.
    """
    if not url or not word_crop:
        return url, False
    src = _question_media_path(url)
    if not src or not src.is_file() or src.suffix.lower() != '.gif':
        return url, False
    try:
        with Image.open(src) as source:
            im = source.convert('RGBA')
        w, h = im.size
        if w < 30 or h < 30:
            return url, False
        alpha = im.getchannel('A')
        bounds = alpha.getbbox()
        if not bounds:
            return url, False
        x0, y0, x1, y1 = bounds
        # Source of the bug: substantial transparent space around the visible
        # drawing, even though the HTML width/height matches wp:extent.
        fraction = (x1-x0) * (y1-y0) / float(w*h)
        if fraction >= 0.88:
            return url, False
        pad = max(2, round(min(w, h)*0.008))
        bounds = (max(0,x0-pad), max(0,y0-pad),
                  min(w,x1+pad), min(h,y1+pad))
        cropped = im.crop(bounds)
        # The source GIF uses transparent magenta; flatten for consistent
        # Word-white appearance in Chrome/Firefox and in PDF exports.
        base = Image.new('RGBA', cropped.size, (255,255,255,255))
        base.alpha_composite(cropped)
        output = MEDIA_DIR / (uuid.uuid4().hex + '_word_figure.png')
        base.convert('RGB').save(output, 'PNG', optimize=True)
        return _question_media_url(output.name), True
    except Exception:
        app.logger.exception('Không thể chuẩn hóa lề trong suốt của hình Word; dùng ảnh gốc')
        return url, False


def _node_image_size(node):
    """KHỐI HÌNH ẢNH — lấy kích thước hiển thị trong Word từ VML/DrawingML."""
    # DrawingML: wp:extent cx/cy are EMU.
    p=node
    for _ in range(8):
        if p is None: break
        ext=p.find(f'{{{"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}}}extent')
        if ext is not None:
            try:
                w=int(ext.get('cx','0'))/914400*96; h=int(ext.get('cy','0'))/914400*96
                if w>0 and h>0: return int(round(w)),int(round(h))
            except Exception: pass
        p=p.getparent()
    # VML: style="width:...;height:...".
    p=node
    for _ in range(8):
        if p is None: break
        style=p.get('style','')
        wm=re.search(r'width\s*:\s*([^;]+)',style,re.I); hm=re.search(r'height\s*:\s*([^;]+)',style,re.I)
        if wm and hm:
            w=_length_to_px(wm.group(1)); h=_length_to_px(hm.group(1))
            if w and h: return w,h
        p=p.getparent()
    return None,None


def _docx_rel_targets(archive):
    """Map relationship IDs to normalized members inside a DOCX archive."""
    try:
        root=etree.fromstring(archive.read('word/_rels/document.xml.rels'))
    except Exception:
        return {}
    result={}
    for rel in root.findall(f'{{{REL_NS}}}Relationship'):
        rid=rel.get('Id'); target=(rel.get('Target') or '').replace('\\','/')
        if rid and target:
            while target.startswith('../'): target=target[3:]
            if not target.startswith('word/'): target='word/'+target.lstrip('/')
            result[rid]=target
    return result


def _export_docx_media(archive,target,role):
    """Export the original DOCX media; vector drawings become crisp SVG."""
    if not target or target not in archive.namelist(): return None
    suffix=Path(target).suffix.lower()
    browser_types={'.png','.jpg','.jpeg','.gif','.webp','.svg'}
    if suffix in browser_types:
        out=MEDIA_DIR/(uuid.uuid4().hex+suffix)
        out.write_bytes(archive.read(target))  # byte-for-byte copy of source image
        return _question_media_url(out.name)

    # Word commonly stores diagrams and MathType previews as EMF/WMF. Convert
    # the original vector file straight to SVG, never via LibreOffice HTML.
    soffice=find_soffice()
    if soffice and suffix in ('.emf','.wmf'):
        with tempfile.TemporaryDirectory() as td:
            td=Path(td); src=td/Path(target).name; src.write_bytes(archive.read(target))
            proc=subprocess.run(
                [soffice,'--headless','--convert-to','svg','--outdir',str(td),str(src)],
                stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,timeout=60
            )
            rendered=td/(src.stem+'.svg')
            if proc.returncode==0 and rendered.exists():
                out=MEDIA_DIR/(uuid.uuid4().hex+'.svg')
                shutil.copy2(rendered,out)
                _crop_libreoffice_svg(out)
                return _question_media_url(out.name)
    return None


def _export_docx_media_batch(archive,targets):
    """Export many original DOCX assets in one LibreOffice process.

    A large exam can contain hundreds of MathType WMF previews. Starting
    LibreOffice once per formula is slow and unreliable, so all vector assets
    are converted together, then cropped individually.
    """
    targets=list(dict.fromkeys(t for t in targets if t in archive.namelist()))
    result={}; browser_types={'.png','.jpg','.jpeg','.gif','.webp','.svg'}
    vector_targets=[t for t in targets if Path(t).suffix.lower() in ('.emf','.wmf')]

    for target in targets:
        suffix=Path(target).suffix.lower()
        if suffix in browser_types:
            out=MEDIA_DIR/(uuid.uuid4().hex+suffix)
            out.write_bytes(archive.read(target))
            result[target]=_question_media_url(out.name)

    soffice=find_soffice()
    if vector_targets and soffice:
        with tempfile.TemporaryDirectory() as td:
            td=Path(td); sources=[]; source_for={}
            for index,target in enumerate(vector_targets):
                # Prefix prevents collisions if two DOCX folders use one name.
                src=td/f'{index:04d}_{Path(target).name}'
                src.write_bytes(archive.read(target)); sources.append(src); source_for[target]=src
            proc=_run_soffice_isolated(
                ['--convert-to','svg','--outdir',str(td),*[str(x) for x in sources]],
                timeout=240
            )
            if proc.returncode==0:
                for target,src in source_for.items():
                    rendered=td/(src.stem+'.svg')
                    if not rendered.exists(): continue
                    out=MEDIA_DIR/(uuid.uuid4().hex+'.svg')
                    shutil.copy2(rendered,out); _crop_libreoffice_svg(out)
                    result[target]=_question_media_url(out.name)
    return result


def _crop_libreoffice_svg(path):
    """Crop LibreOffice's page-sized SVG to its real vector drawing bounds."""
    try:
        parser=etree.XMLParser(resolve_entities=False,no_network=True)
        tree=etree.parse(str(path),parser); root=tree.getroot()
        boxes=root.xpath(".//*[local-name()='rect' and contains(concat(' ',normalize-space(@class),' '),' BoundingBox ')]")
        coords=[]
        for box in boxes:
            x=float(box.get('x','0')); y=float(box.get('y','0'))
            width=float(box.get('width','0')); height=float(box.get('height','0'))
            if width>0 and height>0: coords.append((x,y,x+width,y+height))
        if not coords: return False
        left=min(x[0] for x in coords); top=min(x[1] for x in coords)
        right=max(x[2] for x in coords); bottom=max(x[3] for x in coords)
        pad=max(8.0,min(right-left,bottom-top)*0.015)
        left-=pad; top-=pad; width=(right-left)+pad; height=(bottom-top)+pad
        root.set('viewBox',f'{left:g} {top:g} {width:g} {height:g}')
        root.set('width',f'{width:g}px'); root.set('height',f'{height:g}px')
        root.set('preserveAspectRatio','xMidYMid meet')
        tree.write(str(path),encoding='UTF-8',xml_declaration=True)
        return True
    except Exception:
        app.logger.exception('Không thể crop SVG; giữ SVG đầy đủ')
        return False


def _restore_original_docx_assets(path,equations=None):
    """Restore only real figures from the original DOCX.

    V5.22.10 deliberately does not export MathType WMF/EMF previews to SVG.
    Those SVGs can lose the legacy font's glyph coordinates and produce the
    overlapping characters shown by the user. Formula previews come from the
    same LibreOffice HTML path that displayed correctly in V5.22.0.
    """
    figures=[]
    with zipfile.ZipFile(path) as archive:
        root=etree.fromstring(archive.read('word/document.xml'))
        rels=_docx_rel_targets(archive)
        ns={'w':W_NS,'r':R_NS,'a':A_NS,'v':V_NS}

        # Real figures only: exclude VML previews nested inside MathType OLE.
        nodes=root.xpath(
            './/a:blip[not(ancestor::w:object)] | '
            './/v:imagedata[not(ancestor::w:object)]',namespaces=ns
        )
        figure_specs=[]
        for node in nodes:
            rid=node.get(f'{{{R_NS}}}embed') or node.get(f'{{{R_NS}}}id')
            target=rels.get(rid,''); width,height=_node_image_size(node)
            if target: figure_specs.append((target,width,height))

        all_targets=[x[0] for x in figure_specs]
        exported=_export_docx_media_batch(archive,all_targets)

        for target,width,height in figure_specs:
            media_url=exported.get(target)
            if media_url:
                figures.append({
                    'type':'image','url':media_url,'role':'figure',
                    'display_width_px':width,'display_height_px':height,
                    'original_media':True,
                })

    return figures


def _replace_figure_images_with_originals(lines,figures):
    """Replace LibreOffice HTML figure copies by original DOCX media in order."""
    index=0
    for line in lines:
        for pos,block in enumerate(line):
            if block.get('type')=='image' and block.get('role')=='figure':
                if index<len(figures): line[pos]=figures[index]
                index+=1
    return lines

def _html_image_role(img, src_path):
    """Phân loại ảnh HTML do LibreOffice tạo từ DOCX.

    LibreOffice đặt MathType/OLE thật theo tên ``Object1``, ``Object2``...
    nhưng hình vẽ Word có thể cũng được xuất thành GIF nhỏ với tên ``Picture``.
    Bản cũ chỉ nhìn đuôi GIF/kích thước nên các Picture nhỏ bị đếm nhầm thành
    MathType (ví dụ 293 vị trí trong khi DOCX chỉ có 283 OLE).
    """
    extn = Path(src_path).suffix.lower()
    alt = (img.get('alt') or '').strip().lower()
    name = (img.get('name') or '').strip().lower()
    marker = f"{alt} {name}"
    try:
        w=int(float(img.get('width','0') or 0)); h=int(float(img.get('height','0') or 0))
    except Exception:
        w=h=0

    # Tên ObjectN là tín hiệu ổn định nhất của MathType/OLE sau khi Writer xuất HTML.
    if re.fullmatch(r'object\s*\d+', name, re.I):
        return 'formula'

    # Picture/Shape/Image là hình minh hoạ, kể cả khi LibreOffice xuất thành GIF nhỏ.
    if any(token in marker for token in ('picture', 'shape', 'image', 'graphic')):
        return 'figure'
    if extn in ('.png','.jpg','.jpeg','.bmp','.webp'):
        return 'figure'
    if w>=160 or h>=100:
        return 'figure'

    # GIF nhỏ không có tên Object được coi là hình để không làm lệch chuỗi MTEF.
    # MathType chuẩn đã được bắt bởi ObjectN ở trên.
    return 'figure'


# ============================================================
# KHỐI 1B — CHUẨN HÓA KÍCH THƯỚC CÔNG THỨC MATHTYPE
# Công thức được LibreOffice render thành ảnh. Ảnh thường có
# khoảng trắng dư quanh nét toán học, làm công thức nhìn nhỏ
# hơn chữ dù CSS đã đặt cùng chiều cao.
# Hàm này cắt phần trắng dư, giữ nguyên tỷ lệ, sau đó CSS
# hiển thị ở chiều cao chuẩn tương ứng Times New Roman 13.5pt.
# ============================================================
def normalize_formula_image(src, dst):
    try:
        from PIL import Image, ImageChops
        with Image.open(src) as source_image:
            im = source_image.convert("RGBA")
        alpha = im.getchannel("A")

        # Một số preview MathType/Equation Editor dùng #FF00FF làm chroma-key
        # trong suốt. Khi LibreOffice xuất GIF mà mất cờ transparency, màu này
        # trở thành nền hồng và nét công thức bị ám đỏ (đúng lỗi Câu 10B).
        # Lấy màu nền ở bốn góc rồi khôi phục thành alpha; nét toán giữ màu đen
        # để không còn viền tím do anti-alias trộn với nền chroma-key.
        if alpha.getextrema() == (255,255) and im.width and im.height:
            corners=[
                im.getpixel((0,0))[:3],
                im.getpixel((im.width-1,0))[:3],
                im.getpixel((0,im.height-1))[:3],
                im.getpixel((im.width-1,im.height-1))[:3],
            ]
            bg=max(set(corners),key=corners.count)
            is_magenta=bg[0]>=200 and bg[1]<=90 and bg[2]>=200
            is_neutral_light=max(bg)-min(bg)<=18 and min(bg)>=225
            if is_magenta or is_neutral_light:
                max_distance=max(1,(bg[0]**2+bg[1]**2+bg[2]**2)**0.5)
                pixels=[]
                source_pixels=(
                    im.get_flattened_data()
                    if hasattr(im,"get_flattened_data")
                    else im.getdata()
                )
                for red,green,blue,_ in source_pixels:
                    distance=(
                        (red-bg[0])**2+
                        (green-bg[1])**2+
                        (blue-bg[2])**2
                    )**0.5
                    ink_alpha=max(0,min(255,round(255*distance/max_distance)))
                    pixels.append((0,0,0,ink_alpha))
                im.putdata(pixels)
                alpha=im.getchannel("A")

        bbox = alpha.getbbox()

        # Nếu ảnh có nền trắng nhưng alpha toàn phần, tìm vùng có nét mực.
        if not bbox:
            rgb = im.convert("RGB")
            bg = Image.new("RGB", rgb.size, (255,255,255))
            diff = ImageChops.difference(rgb, bg)
            bbox = diff.point(lambda p: 255 if p > 10 else 0).getbbox()

        if bbox:
            l,t,r,b=bbox
            pad=2
            bbox=(max(0,l-pad), max(0,t-pad), min(im.width,r+pad), min(im.height,b+pad))
            im=im.crop(bbox)

        im.save(dst, format="PNG", optimize=True)
        return True
    except Exception:
        shutil.copy2(src, dst)
        return False


# ============================================================
# KHỐI WORD RIÊNG — V5.12.0
# DOCX được LibreOffice chuyển sang HTML để bảo toàn MathType và
# hình. Kích thước công thức Word dùng kích thước HTML gốc với
# hiệu chỉnh nhẹ 1.05x; KHÔNG dùng scale PDF.
# ============================================================

# ============================================================
# KHỐI WORD RIÊNG — V5.12.3
# ------------------------------------------------------------
# Mục tiêu:
# - Chữ đề: Times New Roman 13.5pt.
# - MathType trong Word đang hơi lớn hơn chữ thường.
# - Vì vậy Word dùng hệ số RIÊNG:
#       WORD_FORMULA_SCALE ≈ 0.98
#       WORD_FORMULA_MIN_PX = 18
# - KHÔNG dùng hệ số PDF cho Word.
# ============================================================

def _sync_docx_question_figures_from_ooxml(questions, path):
    """V5.24.13: gắn lại hình của từng câu theo đúng geometry trong DOCX gốc.

    LibreOffice HTML chỉ dùng để lấy thứ tự text/MathType. Kích thước hình minh họa
    phải lấy từ wp:extent/VML của document.xml; nếu không một số đồ thị (đặc biệt
    Câu 1/3/4) có thể bị thu nhỏ trong bản xem và khi dựng Word fallback.
    """
    try:
        with zipfile.ZipFile(path) as archive:
            root = etree.fromstring(archive.read('word/document.xml'))
            rels = _docx_rel_targets(archive)
            body = root.xpath(".//*[local-name()='body']")[0]

            # Gom top-level XML theo Phần/Câu, giữ đúng thứ tự Word gốc.
            grouped = {}
            sec = None
            qno = None
            for el in list(body):
                if etree.QName(el).localname == 'sectPr':
                    continue
                text = _xml_para_text(el)
                detected = _word_section_no(text)
                if detected:
                    sec = detected
                    qno = None
                    continue
                m = re.match(r"^\s*Câu\s+(\d+)\s*[:.]?", text or '', re.I)
                if m:
                    qno = int(m.group(1))
                if sec and qno:
                    grouped.setdefault((sec,qno), []).append(el)

            # Export đúng media nằm trong từng câu, không ghép theo thứ tự toàn tài liệu.
            for q in questions or []:
                key=(int(q.get('section') or 0), int(q.get('section_order') or 0))
                els=grouped.get(key) or []
                specs=[]
                for el in els:
                    nodes=el.xpath(
                        ".//*[local-name()='blip' and not(ancestor::*[local-name()='object'])] | "
                        ".//*[local-name()='imagedata' and not(ancestor::*[local-name()='object'])]"
                    )
                    for node in nodes:
                        rid=(node.get(f'{{{R_NS}}}embed') or node.get(f'{{{R_NS}}}id'))
                        target=rels.get(rid,'')
                        if not target:
                            continue
                        w,h=_node_image_size(node)
                        crop=_node_image_crop(node)
                        specs.append((target,w,h,crop))
                if not specs:
                    continue
                # V5.24.13: KHÔNG thay ảnh preview bằng SVG chuyển từ EMF/WMF.
                # LibreOffice HTML đã render đúng phần nét/canvas như Word gốc;
                # việc EMF -> SVG + crop có thể làm đồ thị Câu 1/3/4 nhìn nhỏ
                # dù wp:extent vẫn đúng. Ta chỉ đồng bộ geometry OOXML chính xác
                # lên block ảnh hiện có và giữ nguyên URL preview đã render.
                exact=[(w,h,crop) for _target,w,h,crop in specs]
                if not exact:
                    continue

                # Các block của một câu được duyệt theo stem -> options -> solution.
                containers=[q.get('stem') or []]
                containers += [op.get('content') or [] for op in q.get('options') or []]
                containers += [q.get('solution') or []]
                idx=0
                for blocks in containers:
                    for b in blocks:
                        if b.get('type')=='image' and b.get('role')=='figure':
                            if idx < len(exact):
                                w,h,crop=exact[idx]
                                if crop and not b.get('word_src_rect_applied'):
                                    # V5.25.2: LO's Picture*.gif is the FULL
                                    # EMF canvas; the srcRect crop Word applies
                                    # at display time is NOT baked into it.
                                    # Alpha trimming cannot remove long axes
                                    # and tails outside Word's crop rectangle.
                                    # Apply a:srcRect ONCE to this unmodified
                                    # GIF, after checking that the resulting
                                    # aspect ratio matches wp:extent. Keep the
                                    # original DOCX and MathType bytes intact.
                                    url=b.get('url')
                                    if _word_preview_crop_matches_extent(url,crop,w,h):
                                        cropped=_crop_question_preview_to_word_src_rect(url,crop)
                                        if cropped != url:
                                            b['url']=cropped
                                            b['word_preview_cropped_exactly']=True
                                    b['word_src_rect_applied']=True
                                    b['word_src_rect']=crop
                                if w and h:
                                    b['display_width_px']=w
                                    b['display_height_px']=h
                                    b['word_width_cm']=round(w*2.54/96.0,4)
                                    b['word_height_cm']=round(h*2.54/96.0,4)
                                b['original_geometry']=True
                            idx+=1
        return questions
    except Exception:
        app.logger.exception('Không thể đồng bộ geometry hình theo từng câu Word gốc')
        return questions



def _run_soffice_isolated(args, timeout=120):
    """Run LibreOffice with a unique user profile.

    On Windows, a normal ``soffice --headless`` can silently attach to an
    already-running GUI LibreOffice instance or fail because the profile is
    locked.  A private UserInstallation makes each conversion independent.
    """
    soffice=find_soffice()
    if not soffice:
        return None
    with tempfile.TemporaryDirectory(prefix='lo_profile_') as profdir:
        profile_uri=Path(profdir).resolve().as_uri()
        cmd=[soffice, f'-env:UserInstallation={profile_uri}', '--headless', '--nologo', '--nodefault', '--nolockcheck', '--nofirststartwizard'] + list(args)
        return subprocess.run(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, timeout=timeout
        )


def _extract_word_blocks_ooxml_fallback(path):
    """Direct DOCX parser used when LibreOffice HTML conversion fails.

    This path never rewrites the source DOCX. It reads document.xml, MathType
    OLE/OMML and media directly from the ZIP package, so upload/shuffle remains
    available even when LibreOffice is busy, missing, or cannot export HTML.
    """
    equations=extract_docx_equations(path, BASE / 'vendor' / 'mtef_py')
    missing=[str(i+1) for i,e in enumerate(equations) if not e.get('latex')]
    if missing:
        raise RuntimeError('Không thể chuẩn hóa LaTeX cho công thức số: '+', '.join(missing))

    with zipfile.ZipFile(path) as archive:
        root=etree.fromstring(archive.read('word/document.xml'))
        rels=_docx_rel_targets(archive)
        ns={'w':W_NS,'r':R_NS,'a':A_NS,'v':V_NS,'m':'http://schemas.openxmlformats.org/officeDocument/2006/math','o':O_NS}
        figure_nodes=root.xpath(
            './/a:blip[not(ancestor::w:object)] | .//v:imagedata[not(ancestor::w:object)]',
            namespaces=ns
        )
        targets=[]
        for n in figure_nodes:
            rid=n.get(f'{{{R_NS}}}embed') or n.get(f'{{{R_NS}}}id')
            t=rels.get(rid,'')
            if t: targets.append(t)
        exported=_export_docx_media_batch(archive, targets)

        eq_index=0
        lines=[]

        def image_block(node):
            rid=node.get(f'{{{R_NS}}}embed') or node.get(f'{{{R_NS}}}id')
            target=rels.get(rid,'')
            url=exported.get(target)
            if not url:
                return None
            w,h=_node_image_size(node)
            crop=_node_image_crop(node)
            if crop:
                url=_crop_question_preview_to_word_src_rect(url,crop)
            b={'type':'image','url':url,'role':'figure','original_media':True}
            if crop:
                b['word_src_rect']=crop; b['word_src_rect_applied']=True
            if w and h:
                b['display_width_px']=w; b['display_height_px']=h
                b['word_width_cm']=round(w*2.54/96.0,4); b['word_height_cm']=round(h*2.54/96.0,4)
            return b

        def walk(node, blocks):
            nonlocal eq_index
            local=etree.QName(node).localname
            # One MathType OLE object = one canonical LaTeX block. Do not walk
            # its VML preview, otherwise it would be misclassified as a figure.
            if local=='object' and node.xpath('.//*[local-name()="OLEObject"]'):
                if eq_index < len(equations):
                    e=equations[eq_index]; eq_index+=1
                    blocks.append({'type':'latex','value':e.get('latex',''),'display':e.get('display','inline')})
                return
            if local=='oMath':
                # Only top-level oMath, nested math nodes are represented once.
                if any(etree.QName(a).localname=='oMath' for a in node.iterancestors()):
                    return
                if eq_index < len(equations):
                    e=equations[eq_index]; eq_index+=1
                    blocks.append({'type':'latex','value':e.get('latex',''),'display':e.get('display','inline')})
                return
            if local=='t' and node.text:
                if blocks and blocks[-1].get('type')=='text': blocks[-1]['value']+=node.text
                else: blocks.append({'type':'text','value':node.text})
                return
            if local=='tab':
                blocks.append({'type':'text','value':'\t'}); return
            if local in ('br','cr'):
                blocks.append({'type':'text','value':'\n'}); return
            if local=='blip' and not node.xpath('ancestor::*[local-name()="object"]'):
                b=image_block(node)
                if b: blocks.append(b)
                return
            if local=='imagedata' and not node.xpath('ancestor::*[local-name()="object"]'):
                b=image_block(node)
                if b: blocks.append(b)
                return
            for child in node:
                walk(child, blocks)

        body=root.xpath('.//*[local-name()="body"]')[0]
        for el in body:
            if etree.QName(el).localname=='sectPr':
                continue
            if etree.QName(el).localname=='tbl':
                # Flatten table rows in document order; keep each row as one
                # logical line so question boundaries remain stable.
                for tr in el.xpath('./*[local-name()="tr"]'):
                    blocks=[]
                    cells=tr.xpath('./*[local-name()="tc"]')
                    for ci,tc in enumerate(cells):
                        if ci: blocks.append({'type':'text','value':'\t'})
                        walk(tc, blocks)
                    if blocks: lines.append(blocks)
            else:
                blocks=[]; walk(el, blocks)
                if blocks: lines.append(blocks)

    if eq_index != len(equations):
        raise RuntimeError(
            f'OOXML đọc được {eq_index}/{len(equations)} công thức MathType/OMML. '
            'Hệ thống dừng để tránh làm mất công thức.'
        )
    return lines

def extract_word_blocks(path):
    """KHỐI 1 — ĐỌC WORD BẰNG LIBREOFFICE HTML.
    Cách này giữ được cả MathType lẫn VML group/đồ thị của Phần II mà XML thuần thường bỏ sót.
    Thứ tự <p> và <img> được giữ nguyên theo Word.
    """
    # V5.24.16: LibreOffice is an optimisation, not a hard dependency.
    # Use an isolated profile first to avoid Windows profile/GUI lock errors.
    with tempfile.TemporaryDirectory() as td:
        td=Path(td); outdir=td/'html'; outdir.mkdir()
        proc=None
        # Copy to a short ASCII temp name. This avoids edge cases where
        # LibreOffice on Windows cannot reopen a DOCX from an accented/long
        # project path such as D:\WED\HỌC-WED\... .
        lo_source=td/'source.docx'
        shutil.copy2(path,lo_source)
        try:
            proc=_run_soffice_isolated(
                ['--convert-to','html','--outdir',str(outdir),str(lo_source)], timeout=120
            )
        except Exception:
            app.logger.exception('LibreOffice HTML isolated conversion failed; using OOXML fallback')
        htmls=list(outdir.glob('*.html'))
        if proc is None or proc.returncode!=0 or not htmls:
            app.logger.warning(
                'LibreOffice không chuyển được DOCX sang HTML; dùng parser OOXML trực tiếp. stdout=%s stderr=%s',
                (getattr(proc,'stdout','') or '')[-500:], (getattr(proc,'stderr','') or '')[-500:]
            )
            return _extract_word_blocks_ooxml_fallback(path)
        html=htmls[0]
        root=etree.HTML(html.read_bytes())
        result=[]
        for p in root.xpath('//p'):
            blocks=[]
            def walk(el):
                if el.text: blocks.append({'type':'text','value':el.text})
                for child in el:
                    if child.tag=='img':
                        src=child.get('src','')
                        m=re.search(r'(_html_[A-Za-z0-9]+\.(?:gif|png|jpg|jpeg|bmp|webp))',src,re.I)
                        if m:
                            candidate=next((x for x in outdir.iterdir() if x.name.lower().endswith(m.group(1).lower())),None)
                            if candidate and candidate.is_file():
                                role=_html_image_role(child,candidate)
                                out=MEDIA_DIR/(uuid.uuid4().hex+candidate.suffix.lower())
                                if role=='formula':
                                    # Chỉ giữ một vị trí tạm để ghép đúng thứ
                                    # tự MTEF/OMML. Không ghi ảnh công thức ra
                                    # static; vị trí này sẽ được thay hoàn toàn
                                    # bằng block LaTeX bên dưới.
                                    item={'type':'image','role':'formula'}
                                else:
                                    out=MEDIA_DIR/(uuid.uuid4().hex+candidate.suffix.lower())
                                    shutil.copy2(candidate,out)
                                    item={'type':'image','url':_question_media_url(out.name),'role':role}
                                try:
                                    w=int(float(child.get('width','0') or 0)); h=int(float(child.get('height','0') or 0))
                                    if w>0 and h>0:
                                        # V5.11.6: với MathType, dùng chiều cao gốc của Word
                                        # nhưng tối thiểu 18px (13.5pt CSS ≈ 18px).
                                        # Phân số/căn được phép cao hơn 18px để chữ/số bên
                                        # trong không bị thu nhỏ hơn chữ thường.
                                        if role=='formula':
                                            h=max(18,min(round(h*0.98),54))
                                        item.update(display_width_px=w,display_height_px=h)
                                except Exception:
                                    if role=='formula':
                                        item.update(display_height_px=18)
                                blocks.append(item)
                    else:
                        walk(child)
                    if child.tail: blocks.append({'type':'text','value':child.tail})
            walk(p)
            clean=[]
            for b in blocks:
                if b['type']=='image' and clean and b==clean[-1]: continue
                if b['type']=='text' and not b['value'].strip() and not any(x['type']=='image' for x in blocks): continue
                clean.append(b)
            if clean: result.append(clean)
        # V5.22.11: toàn bộ MathType/OMML thành LaTeX chuẩn. Chỉ hình minh họa
        # thật còn là ảnh; công thức được MathJax dàn đồng nhất trên web.
        equations = extract_docx_equations(path, BASE / "vendor" / "mtef_py")
        missing=[str(i+1) for i,e in enumerate(equations) if not e.get('latex')]
        if missing:
            raise RuntimeError(
                'Không thể chuẩn hóa LaTeX cho công thức số: '+', '.join(missing)
            )
        # V5.24.14: GIỮ nguyên ảnh figure do LibreOffice HTML render từ chính Word.
        # Không thay bằng EMF/WMF -> SVG nữa. Với các đồ thị Câu 1/3/4,
        # SVG chuyển từ vector gốc có canvas/viewBox khác Word nên phần nét nhìn nhỏ.
        # Geometry chính xác sẽ được _sync_docx_question_figures_from_ooxml()
        # gắn lại theo wp:extent/VML ngay sau build_questions().
        replace_formula_images_with_latex(result, equations)
        return result



# ============================================================
# KHỐI PDF RIÊNG — V5.12.3
# ------------------------------------------------------------
# Mục tiêu:
# - PDF raster có khoảng trắng/padding nội tại nên glyph thật
#   (chữ, số, dấu + - =, căn, phân số...) nhìn nhỏ hơn chữ đề.
# - Vì vậy PDF dùng hệ số RIÊNG lớn hơn Word:
#       PDF_FORMULA_SCALE ≈ 1.34
#       PDF_FORMULA_MIN_PX = 28
# - Phân số vẫn crop nguyên cụm tử + mẫu, không cắt mẫu số.
# - KHÔNG dùng hệ số Word cho PDF.
# ============================================================
def extract_pdf_blocks(path):
    """Đọc PDF số thành text + LaTeX + hình minh họa gốc.

    PDF không chứa đối tượng MathType/OMML như DOCX.  Với PDF số, bộ đọc dùng
    font và tọa độ từng glyph để dựng lại số mũ, chỉ số, phân số, căn và tích
    phân thành LaTeX chuẩn.  Chỉ ảnh minh họa thật còn được lưu dưới dạng ảnh;
    công thức không còn bị crop thành các bitmap lớn nhỏ khác nhau.
    """
    import pymupdf as fitz

    def save_figure(page, atom):
        block = atom["block"]
        raw = block.get("image")
        suffix = "." + str(block.get("ext") or "png").lower().lstrip(".")
        if suffix not in {".png", ".jpg", ".jpeg", ".gif", ".webp"}:
            suffix = ".png"
            raw = None
        out = MEDIA_DIR / (uuid.uuid4().hex + suffix)
        if raw:
            out.write_bytes(raw)
        else:
            rect = fitz.Rect(atom["bbox"])
            pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=rect, alpha=False)
            pix.save(str(out))
        width = max(1, round((atom["bbox"][2] - atom["bbox"][0]) * 96 / 72))
        height = max(1, round((atom["bbox"][3] - atom["bbox"][1]) * 96 / 72))
        return {
            "type": "image",
            "url": _question_media_url(out.name),
            "role": "figure",
            "display_width_px": width,
            "display_height_px": height,
            "original_media": bool(raw),
        }

    result = []
    document = fitz.open(path)
    try:
        for page in document:
            rows = _pdf_editable_rows(page)
            fraction_rules = _pdf_fraction_rules(page)

            for row in rows:
                atoms = _pdf_visual_order_atoms(row["atoms"])
                normal_sizes = [
                    float(atom.get("size") or 12) for atom in atoms
                    if atom.get("kind") == "text" and not atom.get("math")
                ]
                row_text_size = max(normal_sizes, default=12.0)
                row_text = "".join(
                    str(atom.get("text") or "")
                    for atom in atoms
                    if atom.get("kind") == "text" and not atom.get("math")
                )
                # Bỏ dòng đầu/cuối trang lặp lại; nếu giữ, nó có thể bị nối
                # nhầm vào ý cuối của câu trước sau khi PDF sang trang mới.
                if re.search(r"ÔN\s+TẬP\s+THI\s+TN\s+THPT", row_text, re.I):
                    continue
                if row.get("figure") and len(atoms) == 1 and atoms[0]["kind"] == "image":
                    result.append([save_figure(page, atoms[0])])
                    continue

                blocks = []
                math_run = []
                previous_x1 = None

                def append_text(value, atom=None):
                    nonlocal previous_x1
                    value = str(value or "")
                    if not value:
                        return
                    gap_before = (
                        float(atom["bbox"][0]) - previous_x1
                        if atom is not None and previous_x1 is not None else 0.0
                    )
                    if atom is not None:
                        size = float(atom.get("size") or 12)
                        if previous_x1 is not None and gap_before > max(1.5, size * .18):
                            value = " " + value
                        previous_x1 = float(atom["bbox"][2])
                    if (blocks and blocks[-1].get("type") == "latex"
                            and value and not value[0].isspace()
                            and gap_before > 1.0 and value[0].isalpha()):
                        value = " " + value
                    if blocks and blocks[-1].get("type") == "text":
                        blocks[-1]["value"] += value
                    else:
                        blocks.append({
                            "type": "text", "value": value,
                            "source": "pdf-vector",
                        })

                def flush_math():
                    nonlocal math_run, previous_x1
                    if not math_run:
                        return
                    x0 = min(atom["bbox"][0] for atom in math_run)
                    x1 = max(atom["bbox"][2] for atom in math_run)
                    if previous_x1 is not None and x0 - previous_x1 > 2:
                        append_text(" ")
                    # A row containing only display math is initially created
                    # from the first (often superscript) glyph encountered.
                    # Use the dominant-size glyph baseline so x^3/x^2 stay as
                    # exponents instead of becoming plain trailing digits.
                    formula_baseline = row["cy"]
                    if not any(
                        atom.get("kind") == "text" and not atom.get("math")
                        for atom in atoms
                    ):
                        dominant_size = max(
                            float(atom.get("size") or 12) for atom in math_run
                        )
                        dominant_centers = sorted(
                            float(atom["cy"]) for atom in math_run
                            if float(atom.get("size") or 12) >= dominant_size * .9
                        )
                        if dominant_centers:
                            formula_baseline = dominant_centers[
                                len(dominant_centers) // 2
                            ]
                    latex = _pdf_formula_atoms_to_latex(
                        math_run,
                        formula_baseline,
                        fraction_rules=row.get("fraction_rules", fraction_rules),
                    )
                    latex = normalise_latex(latex or "")
                    if latex:
                        # PDF TeX có thể dùng font văn bản cho đơn vị ``m`` và
                        # font toán nhỏ cho số mũ ². Gộp lại để web hiển thị m²
                        # đúng vị trí thay vì chữ ``2`` đứng riêng trên baseline.
                        unit_match = None
                        if blocks and blocks[-1].get("type") == "text":
                            unit_match = re.search(
                                r"(?<![A-Za-z])(mm|cm|dm|km|m)$",
                                str(blocks[-1].get("value") or ""),
                            )
                        is_raised_unit_power = (
                            unit_match is not None
                            and math_run
                            and max(float(atom.get("size") or 12) for atom in math_run)
                                <= row_text_size * .82
                            and max(float(atom["cy"]) for atom in math_run)
                                < float(row["cy"]) - 1.5
                        )
                        if is_raised_unit_power:
                            unit = unit_match.group(1)
                            blocks[-1]["value"] = blocks[-1]["value"][:unit_match.start()]
                            if not blocks[-1]["value"]:
                                blocks.pop()
                            latex = rf"\mathrm{{{unit}}}^{{{latex}}}"
                        # Vector PDF overlays sometimes spell a single
                        # reaction arrow as the three glyphs "⎯⎯→". MathJax
                        # should receive one command, not raw box lines.
                        latex = re.sub(
                            r"[⎯━─―]{1,6}\\to\b", r"\\longrightarrow ", latex
                        )
                        blocks.append({
                            "type": "latex",
                            "value": latex,
                            "display": "inline",
                            "source": "pdf-vector",
                        })
                    previous_x1 = x1
                    math_run = []

                def attach_chemical_subscript(atom):
                    """Preserve PDF chemistry glyph geometry without converting prose to TeX."""
                    nonlocal previous_x1
                    if (
                        math_run or not blocks or blocks[-1].get("type") != "text"
                        or atom.get("kind") != "text"
                        or not re.fullmatch(r"[0-9]{1,2}", str(atom.get("text") or ""))
                        or float(atom.get("size") or 12) > min(10, row_text_size * .80)
                        or previous_x1 is None
                        or not (-1.5 <= float(atom["bbox"][0]) - previous_x1 <= 2.5)
                        or float(atom.get("cy") or 0) <= float(row["cy"]) + .8
                    ):
                        return False
                    text_block = blocks[-1]
                    previous_text = str(text_block.get("value") or "")
                    # Ca(OH)2 / Fe(NO3)3: PDF exports the closing ')' and
                    # its lowered digit in different spans. Preserve the
                    # entire parenthesised group as one editable math block.
                    if (previous_text == ")" and len(blocks) >= 3
                            and blocks[-2].get("type") == "latex"
                            and blocks[-3].get("type") == "text"):
                        group_prefix = re.search(
                            r"(?<![A-Za-zÀ-ỹ])([A-Z][a-z]?)\($",
                            str(blocks[-3].get("value") or ""),
                        )
                        inside = str(blocks[-2].get("value") or "")
                        if group_prefix and inside.startswith(r"\mathrm{"):
                            preceding = blocks[-3]
                            element = group_prefix.group(1)
                            preceding["value"] = preceding["value"][:group_prefix.start()]
                            blocks[-2:] = [{
                                "type": "latex",
                                "value": (r"\mathrm{" + element + "}(" + inside
                                          + ")_{" + str(atom["text"]) + "}"),
                                "display": "inline", "source": "pdf-vector",
                            }]
                            if not preceding["value"]:
                                blocks.pop(-2)
                            previous_x1 = float(atom["bbox"][2])
                            return True
                    # Match only an element/group suffix; never turn a normal
                    # Vietnamese word ending with 'O' or 'H' into chemistry.
                    match = re.search(
                        r"(?<![A-Za-zÀ-ỹ])(?:[0-9]+)?(?:[A-Z][a-z]?|\([A-Z][A-Za-z]*\)){1,9}$",
                        previous_text,
                    )
                    if match is None:
                        return False
                    prefix, chemical = previous_text[:match.start()], match.group()
                    if prefix:
                        text_block["value"] = prefix
                    else:
                        blocks.pop()
                    coefficient = re.match(r"^([0-9]+)([A-Z].*)$", chemical)
                    chemistry_tex = (
                        coefficient.group(1) + r"\mathrm{" + coefficient.group(2) + "}"
                        if coefficient else r"\mathrm{" + chemical + "}"
                    )
                    blocks.append({
                        "type": "latex",
                        "value": chemistry_tex + "_{" + str(atom["text"]) + "}",
                        "display": "inline", "source": "pdf-vector",
                    })
                    previous_x1 = float(atom["bbox"][2])
                    return True

                for atom in atoms:
                    if atom.get("math"):
                        if attach_chemical_subscript(atom):
                            continue
                        math_run.append(atom)
                        continue
                    flush_math()
                    if atom["kind"] == "image":
                        # PDF không có lớp chữ cho bitmap. Đây là nội dung ảnh
                        # thật, được giữ nguyên bytes và đúng tỷ lệ của PDF.
                        blocks.append(save_figure(page, atom))
                        previous_x1 = atom["bbox"][2]
                        continue
                    text = atom.get("text") or ""
                    # PDF TeX thường in A/B/C/D bằng font đậm nhưng không có
                    # dấu chấm trong lớp text. Bổ sung dấu phân cách nội bộ để
                    # bộ tách phương án nhận đủ bốn đáp án; giao diện vẫn chỉ
                    # sinh nhãn một lần như luồng Word.
                    flags = int(atom.get("flags") or 0)
                    if re.fullmatch(r"\s*[ABCD]\s*", text) and flags & 16:
                        text = text.strip() + ". "
                    append_text(text, atom)

                flush_math()
                if blocks:
                    result.append(merge_text(blocks))
    finally:
        document.close()

    return result


def extract_pdf_blocks_fallback(path):
    """
    ============================================================
    KHỐI PDF DỰ PHÒNG — V5.12.2
    ============================================================
    Chỉ chạy khi parser PDF chính không tạo được câu hỏi.

    Thuật toán:
    - gom span theo tâm-Y với tolerance nhỏ;
    - giữ hình lớn thành dòng riêng;
    - crop math-run theo bbox union.
    Đây là lớp bảo vệ để không còn lỗi "Không nhận diện được câu hỏi"
    với các PDF có cấu trúc font/baseline khác.
    """
    import pymupdf as fitz

    MATH_FONT_RE=re.compile(
        r'(symbol|timesnewroman|cambria|stix|math|mt extra|euclid|cmmi|cmsy)',
        re.I
    )

    def is_math_span(sp):
        font=sp.get('font') or ''
        size=float(sp.get('size') or 0)
        return bool(MATH_FONT_RE.search(font)) or size<9

    def save_clip(page,rect,scale=4.0):
        clip=fitz.Rect(
            max(0,rect.x0-1),max(0,rect.y0-2.5),
            min(page.rect.width,rect.x1+1),
            min(page.rect.height,rect.y1+2.5)
        )
        pix=page.get_pixmap(
            matrix=fitz.Matrix(scale,scale),
            clip=clip,alpha=False
        )
        out=MEDIA_DIR/(uuid.uuid4().hex+'.png')
        pix.save(str(out))
        return _question_media_url(out.name)

    doc=fitz.open(path)
    result=[]

    for page in doc:
        atoms=[]
        d=page.get_text('dict')

        for b in d.get('blocks',[]):
            if b.get('type')!=0:
                continue
            for line in b.get('lines',[]):
                for sp in line.get('spans',[]):
                    text=normalize_text(sp.get('text',''))
                    if not text:
                        continue
                    bb=tuple(sp.get('bbox',(0,0,0,0)))
                    atoms.append({
                        'kind':'text',
                        'text':text,
                        'math':is_math_span(sp),
                        'bbox':bb,
                        'y':(bb[1]+bb[3])/2,
                        'x':bb[0]
                    })

        for b in d.get('blocks',[]):
            if b.get('type')!=1:
                continue
            x0,y0,x1,y1=map(float,b.get('bbox',(0,0,0,0)))
            w=x1-x0
            h=y1-y0
            if w<8 or h<8:
                continue
            role='figure' if (w>=50 and h>=25) else 'formula'
            atoms.append({
                'kind':'image',
                'role':role,
                'url':save_clip(page,fitz.Rect(x0,y0,x1,y1)),
                'bbox':(x0,y0,x1,y1),
                'display_width_px':max(1,round(w*96/72)),
                'display_height_px':max(1,round(h*96/72)),
                'y':(y0+y1)/2,
                'x':x0
            })

        atoms.sort(key=lambda z:(z['y'],z['x']))
        visual=[]

        for atom in atoms:
            if atom.get('role')=='figure':
                visual.append([atom])
                continue

            target=None
            h=max(1,atom['bbox'][3]-atom['bbox'][1])
            threshold=max(4.5,min(9,h*.7))

            for line in visual:
                if any(x.get('role')=='figure' for x in line):
                    continue
                line_y=sum(x['y'] for x in line)/len(line)
                if abs(atom['y']-line_y)<=threshold:
                    target=line
                    break

            if target is None:
                visual.append([atom])
            else:
                target.append(atom)

        for line in visual:
            if len(line)==1 and line[0].get('role')=='figure':
                x=line[0]
                result.append([{
                    'type':'image',
                    'url':x['url'],
                    'role':'figure',
                    'display_width_px':x['display_width_px'],
                    'display_height_px':x['display_height_px']
                }])
                continue

            line.sort(key=lambda z:z['x'])
            blocks=[]
            math_run=[]

            def flush_math():
                nonlocal math_run
                if not math_run:
                    return
                rect=fitz.Rect(
                    min(x['bbox'][0] for x in math_run),
                    min(x['bbox'][1] for x in math_run),
                    max(x['bbox'][2] for x in math_run),
                    max(x['bbox'][3] for x in math_run)
                )
                raw=''.join(x.get('text','') for x in math_run)
                blocks.append({
                    'type':'image',
                    'url':save_clip(page,rect),
                    'role':'formula',
                    'formula_text':raw,
                    'display_width_px':max(1,round(rect.width*96/72)),
                    'display_height_px':max(
                        22,
                        min(64,round((rect.height+5)*96/72*1.10))
                    )
                })
                math_run=[]

            i=0
            while i<len(line):
                atom=line[i]

                if atom['kind']=='image':
                    flush_math()
                    blocks.append({
                        'type':'image',
                        'url':atom['url'],
                        'role':atom['role'],
                        'display_width_px':atom['display_width_px'],
                        'display_height_px':atom['display_height_px']
                    })
                elif atom.get('math'):
                    math_run.append(atom)
                else:
                    txt=atom.get('text','')
                    punct_only=bool(re.fullmatch(
                        r'[\s.,;:!?%\)\(\]\[\-]*',
                        txt or ''
                    ))
                    if (
                        punct_only and math_run and i+1<len(line)
                        and line[i+1].get('math')
                    ):
                        math_run.append(atom)
                    else:
                        flush_math()
                        blocks.append({
                            'type':'text',
                            'value':txt
                        })
                i+=1

            flush_math()
            if blocks:
                result.append(blocks)

    return result


# ============================================================
# V5.15.0 — KHỐI LATEX CHO TẠO ĐỀ THỦ CÔNG
# ------------------------------------------------------------
# Không dùng cho Word/PDF. Word/PDF vẫn chạy đúng parser V5.12.6+.
#
# Quy ước nhập:
#   Văn bản thường: f(x) là hàm số...
#   Công thức inline: $x^2+1$
#   Công thức riêng dòng: $$\frac{x+1}{x-1}$$
#
# Kết quả lưu dưới dạng block type='latex'.
# Frontend dùng MathJax để hiển thị.
# ============================================================
def manual_latex_blocks(text):
    text = str(text or "")
    if not text:
        return []

    pattern = re.compile(r"(\$\$.*?\$\$|\$.*?\$)", re.S)
    parts = pattern.split(text)
    blocks = []

    for part in parts:
        if not part:
            continue

        if part.startswith("$$") and part.endswith("$$"):
            formula = normalise_latex(part[2:-2])
            if formula:
                blocks.append({
                    "type": "latex",
                    "value": formula,
                    "display": "block"
                })
        elif part.startswith("$") and part.endswith("$"):
            formula = normalise_latex(part[1:-1])
            if formula:
                blocks.append({
                    "type": "latex",
                    "value": formula,
                    "display": "inline"
                })
        else:
            blocks.append({
                "type": "text",
                "value": part
            })

    return blocks


def parse_tf_key(value):
    """
    Nhận các dạng:
      ĐSĐS
      DSDS
      true,false,true,false
      1,0,1,0
    Trả về list 4 bool.
    """
    raw = normalize_text(str(value or "")).upper()
    raw = raw.replace("TRUE", "Đ").replace("FALSE", "S")
    raw = raw.replace("1", "Đ").replace("0", "S")
    raw = raw.replace(",", "").replace(";", "").replace(" ", "")
    raw = raw.replace("D", "Đ")
    vals = []
    for ch in raw:
        if ch == "Đ":
            vals.append(True)
        elif ch == "S":
            vals.append(False)
    return vals[:4]

def blocks_text(blocks):
    """Text cho parser; formula_text giúp công thức ảnh vẫn được nhận diện."""
    vals=[]
    for b in blocks:
        if b.get('type')=='text': vals.append(b.get('value',''))
        elif b.get('type')=='image' and b.get('formula_text'): vals.append(b.get('formula_text',''))
    return normalize_text(' '.join(vals))

def merge_text(blocks):
    out=[]
    for b in blocks:
        if b["type"]=="text":
            if out and out[-1]["type"]=="text": out[-1]["value"] += b["value"]
            else: out.append(dict(b))
        elif (
            b.get("type")=="latex" and out and out[-1].get("type")=="latex"
            and b.get("source")=="pdf-vector"
            and out[-1].get("source")=="pdf-vector"
            and b.get("display","inline")=="inline"
            and out[-1].get("display","inline")=="inline"
        ):
            # Một công thức PDF có thể bị ngắt ở mép dòng (Câu 11: biểu thức
            # góc vectơ và 120°). Dữ liệu ngầm phải là một biểu thức TeX hoàn
            # chỉnh trước khi MathJax dựng ra màn hình.
            out[-1]["value"] += b.get("value","")
        else: out.append(dict(b))
    for b in out:
        if b["type"]=="text": b["value"]=re.sub(r"\s+"," ",b["value"])
    return out


def _extend_parsed_blocks(target, incoming):
    """Nối các dòng PDF nhưng không làm thay đổi luồng khối của Word."""
    incoming=merge_text(incoming)
    if target and incoming:
        left=target[-1]
        right=incoming[0]
        if (
            left.get("type")=="text" and right.get("type")=="text"
            and left.get("source")=="pdf-vector"
            and right.get("source")=="pdf-vector"
        ):
            left_value=str(left.get("value") or "")
            right_value=str(right.get("value") or "")
            if (
                left_value and right_value
                and not left_value[-1].isspace()
                and not right_value[0].isspace()
                and right_value[0] not in ".,;:!?%)]}"
                and left_value[-1] not in "([{" 
            ):
                left["value"] += " "
        elif (
            left.get("source")=="pdf-vector"
            and right.get("source")=="pdf-vector"
            and left.get("type") in {"text", "latex"}
            and right.get("type")=="text"
        ):
            # Một dòng PDF có thể kết thúc bằng công thức và dòng sau bắt đầu
            # bằng văn bản (ví dụ ``t=6`` rồi ``giờ``). Giữ một khoảng cách
            # hiển thị thay vì để hai nội dung dính liền nhau.
            left_value=str(left.get("value") or "")
            right_value=str(right.get("value") or "")
            if (
                left_value and right_value
                and not right_value[0].isspace()
                and right_value[0] not in ".,;:!?%)]}"
            ):
                right["value"]=" "+right_value
    target.extend(incoming)
    target[:]=merge_text(target)
    return target


def split_option_atoms(line, case_sensitive=False):
    """Split one paragraph containing A...B...C...D with formulas/images between labels."""
    # Find option labels in text atoms while preserving image atoms.
    parts=[]; current=None; buffer=[]
    for atom in line:
        if atom["type"]=="text":
            text=atom["value"]
            pos=[]
            flags=0 if case_sensitive else re.I
            for m in re.finditer(r"(?:^|\s|\.|\n)\s*([ABCD])\s*[.)]\s*", text, flags):
                pos.append(m)
            if pos:
                last=0
                for m in pos:
                    prefix=text[last:m.start()]
                    if prefix: buffer.append({"type":"text","value":prefix})
                    if current is not None:
                        parts.append((current, merge_text(buffer))); buffer=[]
                    current=m.group(1).upper(); last=m.end()
                tail=text[last:]
                if tail: buffer.append({"type":"text","value":tail})
            else:
                buffer.append(atom)
        else:
            buffer.append(atom)
    if current is not None:
        parts.append((current, merge_text(buffer)))
    return parts


def detect_section(text):
    """Nhận tiêu đề phần bằng cả số La Mã và chữ số 1–4.

    Nhiều file Word dùng ``Phần 1.``, ``Phần 2.`` trong khi PDF/đề khác
    dùng ``PHẦN I``, ``PHẦN II``. Bản cũ chỉ nhận số La Mã nên Câu 1
    của Phần 2 bị nối tiếp thành Câu 13 của Phần 1. Chỉ nhận ở đầu dòng
    để chữ "Phần II" xuất hiện trong lời giải không làm ngắt câu.
    """
    m=re.match(
        r"^\s*(?:[•Ê☞\uf0dc]\s*)?Phần\s*(I{1,3}|IV|[1-4])\b",
        text or "", re.I
    )
    if not m:
        return None
    token=m.group(1).upper()
    return {'I':1,'II':2,'III':3,'IV':4,'1':1,'2':2,'3':3,'4':4}.get(token)

def split_solution_marker(line):
    """Tách dòng tại nhãn ``Lời giải`` nhưng giữ nguyên thứ tự mọi block.

    Một số PDF dùng glyph bullet được trích thành chữ ``Ê`` ngay trước nhãn,
    nên biểu thức chấp nhận cả ``Ê Lời giải.`` và ``Lời giải:``.
    """
    before=[]; after=[]; found=False
    marker_re=re.compile(r"(?:Ê\s*)?Lời\s*giải\s*[:：.]?",re.I)
    for block in line:
        if found:
            after.append(dict(block)); continue
        if block.get('type')!='text':
            before.append(dict(block)); continue
        value=str(block.get('value') or '')
        marker=marker_re.search(value)
        if not marker:
            before.append(dict(block)); continue
        left=value[:marker.start()]
        right=value[marker.end():]
        if left.strip():
            copied=dict(block); copied['value']=left
            before.append(copied)
        if right.strip():
            copied=dict(block); copied['value']=right
            after.append(copied)
        found=True
    return merge_text(before),merge_text(after),found


def clean_solution_line(line):
    """Bỏ leader/chữ kết thúc do PDF sinh ra, không đụng tới công thức/hình."""
    cleaned=[]
    for block in line:
        copied=dict(block)
        if copied.get('type')=='text':
            value=str(copied.get('value') or '')
            value=re.sub(r"(?:\s*[.·…]){5,}\s*(?:□)?\s*$","",value)
            value=re.sub(r"\s*□\s*$","",value)
            copied['value']=value
            if not value.strip():
                continue
        cleaned.append(copied)
    return merge_text(cleaned)


def append_solution_line(target,line):
    """Giữ ranh giới đoạn để lời giải dài không bị dính liền chữ."""
    line=clean_solution_line(line)
    if not line:
        return
    if target and target[-1].get('type')!='break':
        target.append({'type':'break'})
    target.extend(line)


def update_answer_from_solution(q,line):
    """Đọc đáp án nằm trong lời giải của ba phần, nhưng vẫn giữ lời giải để xem."""
    values=[]
    for block in line:
        if block.get('type') in ('text','latex'):
            values.append(str(block.get('value') or ''))
        elif block.get('type')=='image' and block.get('formula_text'):
            values.append(str(block.get('formula_text') or ''))
    text_value=normalize_text(' '.join(values))
    if not text_value:
        return

    if q.get('qtype')=='mcq':
        match=re.search(
            r"(?:Chọn(?:\s+đáp\s*án)?|Đáp\s*án)\s*[:：.\-]?\s*([ABCD])\b",
            text_value,re.I
        )
        if match:
            q['correct_option']=match.group(1).upper()
        return

    if q.get('qtype')=='tf':
        answer_map=q.setdefault('_solution_tf_answers',{})
        for match in re.finditer(r"\b([abcd])\s*[).:：\-]?\s*(đúng|sai)\b",text_value,re.I):
            answer_map[match.group(1).upper()]=match.group(2).lower()=='đúng'
        if all(key in answer_map for key in 'ABCD'):
            q['tf_answers']=[answer_map[key] for key in 'ABCD']
        return

    if q.get('qtype')=='short':
        match=re.search(
            r"(?:Trả\s*lời|Đáp\s*án)\s*[:：.\-]?\s*"
            r"([+\-]?\d+(?:[.,]\d+)?(?:\s*/\s*\d+(?:[.,]\d+)?)?)",
            text_value,re.I
        )
        if match:
            q['short_answer']=re.sub(r"\s+","",match.group(1))


def canonicalize_parsed_options(q):
    """Bảo đảm mỗi câu chỉ có một phương án cho từng khóa A/B/C/D.

    DOCX/PDF đôi khi ngắt một phương án thành nhiều dòng hoặc nhận nhầm một
    cụm ``A.``–``D.`` trong nội dung thành nhãn mới. Trước V5.24.1, các mảnh
    đó tạo hai bản ghi có cùng ``option_key`` và SQLite dừng toàn bộ lần tải
    bằng lỗi UNIQUE. Hàm này gộp các mảnh trùng mà vẫn giữ chữ, công thức và
    hình ảnh; câu đã tách đúng A/B/C/D được giữ nguyên.
    """
    normalized=[]
    by_key={}
    for raw_option in q.get('options',[]) or []:
        key=str(raw_option.get('key') or '').strip().upper()
        content=merge_text(raw_option.get('content') or [])

        # Một mảnh không có khóa hợp lệ vẫn phải được giữ lại. Nó thuộc phần
        # nội dung gần nhất, hoặc thuộc đề bài nếu chưa có phương án nào.
        if key not in 'ABCD' or len(key)!=1:
            target=(
                normalized[-1]['content'] if normalized
                else q.setdefault('stem',[])
            )
            _extend_parsed_blocks(
                target,
                [{'type':'text','value':' '}] + content if target else content,
            )
            continue

        if key not in by_key:
            option=dict(raw_option)
            option['key']=key
            option['content']=content
            normalized.append(option)
            by_key[key]=option
            continue

        # Bỏ đúng một lớp PDF/Word bị nhân đôi hoàn toàn.
        existing=by_key[key]
        if content == existing.get('content',[]):
            continue

        # Nếu A/B/C/D đã đủ rồi mà lại gặp nhãn lùi về trước, đây thường là
        # chữ giống nhãn nằm trong nội dung phương án cuối. Trả nhãn về dạng
        # văn bản và nối đúng tại vị trí đang đọc. Các trường hợp còn lại là
        # một phương án bị tách thành nhiều mảnh, nên gộp về đúng khóa đó.
        if len(by_key)==4 and normalized[-1]['key']!=key:
            destination=normalized[-1]
            label=(key.lower()+') ') if q.get('qtype')=='tf' else (key+'. ')
            incoming=[{'type':'text','value':' '+label}]+content
        else:
            destination=existing
            incoming=[{'type':'text','value':' '}]+content
        _extend_parsed_blocks(destination['content'],incoming)

    q['options']=normalized
    return q


def finalize_parsed_question(q):
    canonicalize_parsed_options(q)
    q.pop('_solution_tf_answers',None)
    return q


QUESTION_PREFIX_RE=re.compile(
    r"^\s*C\s*[âa]\s*u\s+(\d+)\s*[:.]?\s*",re.I
)


def match_question_prefix(line):
    """Nhận cả ``Câu 5`` và ``C âu 5`` do hình neo trong DOCX tách chữ."""
    text_only=''.join(
        str(block.get('value') or '')
        for block in line if block.get('type')=='text'
    )
    return QUESTION_PREFIX_RE.match(text_only)


def strip_question_prefix(line,match):
    """Bỏ nhãn Câu N xuyên qua nhiều text block, vẫn giữ hình/công thức."""
    remaining=match.end()
    result=[]
    for block in line:
        if block.get('type')!='text':
            result.append(dict(block)); continue
        value=str(block.get('value') or '')
        if remaining >= len(value):
            remaining-=len(value)
            continue
        if remaining:
            value=value[remaining:]
            remaining=0
        if value:
            copied=dict(block); copied['value']=value
            result.append(copied)
    return merge_text(result)

def parse_answer_marker(text):
    """Return a hidden answer marker found at the end of a question."""
    if not text: return None
    patterns=[
        (r"(?:Đáp\s*án|ĐA|Answer)\s*[:：-]\s*([ABCD])\b", "mcq"),
        (r"(?:Đáp\s*án|ĐA|Answer)\s*[:：-]\s*((?:[ĐđSsTt](?:\s*[,;/]\s*)?){4})\s*$", "tf"),
        (r"(?:Đáp\s*án|ĐA|Answer)\s*[:：-]\s*([-0-9,]+)\s*$", "short"),
    ]
    for pat, kind in patterns:
        m=re.search(pat,text,re.I)
        if m: return kind,m.group(1).strip(),m.start()
    return None

def trim_answer_marker(line):
    """Split blocks before the text 'Đáp án: ...' so the answer is never shown to students."""
    out=[]; marker=None
    for b in line:
        if b.get('type')!='text' or marker is not None:
            out.append(b); continue
        val=b.get('value','')
        # marker may be a separate paragraph or after option D.
        m=re.search(r"(?:Đáp\s*án|ĐA|Answer)\s*[:：-]",val,re.I)
        if m:
            before=val[:m.start()]
            if before: out.append({'type':'text','value':before})
            tail=val[m.end():]
            # Capture a compact answer value from the remainder.
            marker_text=tail.strip()
            marker=parse_answer_marker('Đáp án: '+marker_text)
            if marker is None:
                # For markers that continue until line end, infer a raw value.
                marker_text=re.split(r'\s{2,}',marker_text)[0].strip()
                marker=('raw',marker_text,0)
            else:
                marker=(marker[0],marker[1],0)
            continue
        out.append(b)
    return out, marker

def apply_answer(q, marker):
    if not marker: return
    kind,val,_=marker
    if q['qtype']=='mcq' and kind=='mcq': q['correct_option']=val.upper()
    elif q['qtype']=='tf' and kind=='tf':
        q['tf_answers']=[x.lower() in ('đ','d','t') for x in re.findall(r'[ĐđSsTt]', val)]
    elif q['qtype']=='short' and kind in ('short','raw'):
        q['short_answer']=val.strip()

def build_questions(lines):
    qs=[]; current=None; current_opt=None; section=1; in_solution=False
    section_counters={1:0,2:0,3:0,4:0}
    global_order=0
    for raw_line in lines:
        line=merge_text(raw_line); s=blocks_text(line)
        # Dòng chỉ chứa công thức LaTeX (thường là phương trình hiển thị trong
        # lời giải PDF) vẫn là nội dung hợp lệ. Bản cũ chỉ xét text/ảnh nên làm
        # mất trọn các dòng phân số, căn và tích phân dù LaTeX đã dựng đúng.
        if not s and not any(b.get('type') in ('image','latex') for b in line):
            continue
        sec=detect_section(s)
        if sec:
            if current: qs.append(finalize_parsed_question(current)); current=None; current_opt=None
            in_solution=False
            section=sec; continue

        # Nhận câu mới trước khi xử lý trạng thái lời giải của câu trước.
        m=match_question_prefix(line)
        if m:
            if current: qs.append(finalize_parsed_question(current))
            in_solution=False
            clean_line, marker = trim_answer_marker(line)
            qnum=int(m.group(1))
            new=strip_question_prefix(clean_line,m)
            # The source may restart numbering at Câu 1 for every section.
            # Keep a unique global order for the database and a separate section_order for display.
            section_counters[section] += 1
            global_order += 1
            section_no = section_counters[section]
            qtype={1:'mcq',2:'tf',3:'short',4:'essay'}[section]
            current={'original_order':global_order,'section_order':section_no,'section':section,'qtype':qtype,'stem':[], 'options':[],
                     'correct_option':None,'short_answer':None,'tf_answers':[],'solution':[]}
            # If the question stem and A/B/C/D are in the same Word paragraph,
            # split them immediately instead of treating the options as stem text.
            if qtype in ('mcq','tf'):
                first_label=None
                for bi,b in enumerate(new):
                    if b.get('type')=='text':
                        option_flags=0 if qtype=='mcq' else re.I
                        mm=re.search(r'(?:^|\s)([ABCD])\s*[.)]\s*', b.get('value',''), option_flags)
                        if mm:
                            first_label=(bi,mm); break
                if first_label:
                    bi,mm=first_label
                    prefix=new[:bi]
                    if prefix: current['stem'].extend(prefix)
                    val=new[bi].get('value',''); before=val[:mm.start()].strip()
                    if before: current['stem'].append({'type':'text','value':before})
                    suffix=val[mm.start():]
                    rest=[{'type':'text','value':suffix}]+new[bi+1:]
                    parts0=split_option_atoms(rest,case_sensitive=qtype=='mcq')
                    for key,content in parts0:
                        current['options'].append({'key':key,'content':content})
                    current_opt=current['options'][-1] if current['options'] else None
                else:
                    current['stem']=merge_text(new); current_opt=None
            else:
                current['stem']=merge_text(new); current_opt=None
            apply_answer(current,marker); continue
        if not current: continue

        # Từ nhãn "Lời giải" trở đi, nội dung thuộc riêng câu hiện tại và
        # tuyệt đối không được rơi vào đề/đáp án hiển thị cho học sinh.
        before_solution, after_solution, has_solution = split_solution_marker(line)
        if has_solution:
            if before_solution and not in_solution:
                clean_before, marker = trim_answer_marker(before_solution)
                parts=split_option_atoms(clean_before,case_sensitive=current['qtype']=='mcq') if current['qtype'] in ('mcq','tf') else []
                if parts:
                    for key,content in parts:
                        current['options'].append({'key':key,'content':content}); current_opt=current['options'][-1]
                elif clean_before:
                    _extend_parsed_blocks(
                        current_opt['content'] if current_opt else current['stem'],
                        clean_before,
                    )
                apply_answer(current,marker)
            in_solution=True
            if after_solution:
                append_solution_line(current['solution'],after_solution)
                update_answer_from_solution(current,after_solution)
            continue

        if in_solution:
            append_solution_line(current['solution'],line)
            update_answer_from_solution(current,line)
            continue

        # A hidden answer marker may be at the end of a question/option paragraph.
        clean_line, marker = trim_answer_marker(line)
        parts=split_option_atoms(clean_line,case_sensitive=current['qtype']=='mcq') if current['qtype'] in ('mcq','tf') else []
        if parts:
            for key,content in parts:
                current['options'].append({'key':key,'content':content}); current_opt=current['options'][-1]
        elif clean_line:
            _extend_parsed_blocks(
                current_opt['content'] if current_opt else current['stem'],
                clean_line,
            )
        apply_answer(current,marker)
    if current: qs.append(finalize_parsed_question(current))
    return qs

def parse_answer_key_text(raw):
    mcq={}; tf={}; short={}
    raw=raw or ''
    for m in re.finditer(r'(?:Câu\s*)?(\d+)\s*[:.\-]\s*([ABCD])\b',raw,re.I): mcq[int(m.group(1))]=m.group(2).upper()
    for m in re.finditer(r'(?:TF|ĐS|DS)\s*(\d+)\s*[:.\-]\s*([ĐđTtSsFf](?:\s*[,;/]\s*[ĐđTtSsFf]){1,3})',raw,re.I):
        vals=[]
        for x in re.split(r'[,;/\s]+',m.group(2).strip()):
            if x: vals.append(x.lower() in ('đ','d','t','true'))
        tf[int(m.group(1))]=vals
    for m in re.finditer(r'(?:TLN|SHORT)\s*(\d+)\s*[:.\-]\s*([^;\n]+)',raw,re.I): short[int(m.group(1))]=m.group(2).strip()
    return {'mcq':mcq,'tf':tf,'short':short}

def first_image(blocks):
    for b in blocks:
        if b.get("type")=="image": return b.get("url")
    return None


def get_exam_questions(exam_id):
    with db() as c:
        qs=c.execute("SELECT * FROM questions WHERE exam_id=? ORDER BY section, section_order, original_order",(exam_id,)).fetchall()
        out=[]
        for q in qs:
            ops=c.execute("SELECT * FROM options WHERE question_id=? ORDER BY option_key",(q["id"],)).fetchall()
            qblocks = q["blocks_json"] or (q["stem_json"] if "stem_json" in q.keys() else "[]") or "[]"
            solution_json = q["solution_json"] if "solution_json" in q.keys() else "[]"
            out.append({"id":q["id"],"original_order":q["original_order"],"section_order":q["section_order"],"section":q["section"],"qtype":q["qtype"],"blocks":json.loads(qblocks),"solution":json.loads(solution_json or "[]"),"correct_option":q["correct_option"],"short_answer":q["short_answer"],"tf_answers":json.loads(q["tf_answers_json"] or "[]") if "tf_answers_json" in q.keys() else [],"answer_text":((q["correct_option"] or "") if q["qtype"]=="mcq" else ("".join("Đ" if x else "S" for x in (json.loads(q["tf_answers_json"] or "[]") if q["qtype"]=="tf" else []))) if q["qtype"]=="tf" else (q["short_answer"] or "")),"options":[{"key":o["option_key"],"blocks":json.loads((o["blocks_json"] or (o["content_json"] if "content_json" in o.keys() else "[]") or "[]"))} for o in ops]})
        return out


def create_student_exam(exam_id, student_id):
    with db() as c:
        ensure_practice_exam_schema(c)
        old=c.execute("SELECT * FROM student_exams WHERE exam_id=? AND student_id=? AND submitted_at IS NULL ORDER BY id DESC LIMIT 1",(exam_id,student_id)).fetchone()
        if old:
            c.execute(
                "UPDATE student_exams SET last_activity_at=CURRENT_TIMESTAMP WHERE id=?",
                (old["id"],),
            )
            return c.execute(
                "SELECT * FROM student_exams WHERE id=?", (old["id"],)
            ).fetchone()
        exam=c.execute("SELECT * FROM exams WHERE id=? AND published=1",(exam_id,)).fetchone()
        if not exam: return None
        qs=c.execute("SELECT id,section,section_order FROM questions WHERE exam_id=? ORDER BY section,section_order,original_order",(exam_id,)).fetchall()
        # Randomize questions independently inside each part, while keeping Parts I→IV in order.
        by_section={1:[],2:[],3:[],4:[]}
        for r in qs: by_section.setdefault(r[1],[]).append(r[0])
        order=[]
        for sec in (1,2,3,4):
            part=list(by_section.get(sec,[])); random.shuffle(part); order.extend(part)
        option_map={}
        for qid in order:
            qrow=c.execute("SELECT section FROM questions WHERE id=?",(qid,)).fetchone()
            keys=[r[0] for r in c.execute("SELECT option_key FROM options WHERE question_id=? ORDER BY option_key",(qid,)).fetchall()]
            # KHỐI 4 — RANDOM PHƯƠNG ÁN: chỉ Phần I đảo A-D; Phần II tuyệt đối giữ nguyên A-D.
            if qrow and qrow[0]==1: random.shuffle(keys)
            option_map[str(qid)]=keys
        cur=c.execute("INSERT INTO student_exams(exam_id,student_id,question_order_json,option_map_json,last_activity_at) VALUES(?,?,?,?,CURRENT_TIMESTAMP)",(exam_id,student_id,jdump(order),jdump(option_map)))
        return c.execute("SELECT * FROM student_exams WHERE id=?",(cur.lastrowid,)).fetchone()



# ============================================================
# V5.13.0 — KHỐI QUYỀN QUẢN TRỊ
# ============================================================
def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login", next=request.path))
        if not session.get("is_admin"):
            abort(403)
        return fn(*args, **kwargs)
    return wrapper


# ============================================================
# V5.13.0 — ĐĂNG KÝ TÀI KHOẢN GIÁO VIÊN / HỌC SINH
# Tài khoản mới chỉ được dùng sau khi Admin bấm Đồng ý.
# ============================================================
@app.route("/register/<role>", methods=["GET","POST"])
def register_account(role):
    if role not in ("teacher","student"):
        abort(404)

    if request.method == "POST":
        full_name = normalize_text(request.form.get("full_name"))
        username = normalize_text(request.form.get("username"))
        password = request.form.get("password","")
        password2 = request.form.get("password2","")
        email = normalize_text(request.form.get("email"))
        phone = normalize_text(request.form.get("phone"))
        student_code = normalize_text(request.form.get("student_code"))

        if not full_name or not username or not password:
            flash("Vui lòng nhập đầy đủ họ tên, tài khoản và mật khẩu.")
            return redirect(request.url)

        if len(username) < 4:
            flash("Tên đăng nhập cần ít nhất 4 ký tự.")
            return redirect(request.url)

        if len(password) < 6:
            flash("Mật khẩu cần ít nhất 6 ký tự.")
            return redirect(request.url)

        if password != password2:
            flash("Mật khẩu nhập lại chưa khớp.")
            return redirect(request.url)

        if role == "student" and student_code:
            # Nếu mã đã được giáo viên import từ Excel, học sinh không cần
            # chờ đăng ký để làm bài: ưu tiên tài khoản/mật khẩu là mã HS.
            with db() as c:
                excel_student = c.execute(
                    "SELECT id FROM users WHERE student_code=? AND role='student'",
                    (student_code,)
                ).fetchone()
            if excel_student:
                flash(
                    "Mã học sinh đã có trong danh sách lớp. "
                    "Bạn có thể đăng nhập ngay bằng Mã học sinh cho cả "
                    "tài khoản và mật khẩu."
                )
                return redirect(url_for("login", role="student"))

        with db() as c:
            if c.execute(
                "SELECT 1 FROM users WHERE username=?",
                (username,)
            ).fetchone():
                flash("Tên đăng nhập đã được sử dụng.")
                return redirect(request.url)

            if c.execute("""
                SELECT 1 FROM registration_requests
                WHERE username=? AND status='pending'
            """,(username,)).fetchone():
                flash("Tài khoản này đang chờ Admin duyệt.")
                return redirect(request.url)

            if role == "student" and student_code:
                if c.execute("""
                    SELECT 1 FROM registration_requests
                    WHERE student_code=? AND status='pending'
                """,(student_code,)).fetchone():
                    flash("Mã học sinh này đã có yêu cầu đang chờ duyệt.")
                    return redirect(request.url)

            c.execute("""
                INSERT INTO registration_requests(
                    requested_role,username,password_hash,full_name,
                    email,phone,student_code,status
                ) VALUES(?,?,?,?,?,?,?,'pending')
            """,(
                role,username,generate_password_hash(password),full_name,
                email or None,phone or None,student_code or None
            ))

        flash("Đăng ký thành công. Vui lòng chờ Admin phê duyệt.")
        return redirect(url_for("login", role=role))

    return render_template("register.html", role=role)


# ============================================================
# V5.13.0 — ADMIN DASHBOARD
# ============================================================
@app.get("/admin")
@admin_required
def admin_dashboard():
    with db() as c:
        requests = c.execute("""
            SELECT * FROM registration_requests
            ORDER BY CASE status
                WHEN 'pending' THEN 0
                WHEN 'approved' THEN 1
                ELSE 2 END,
                id DESC
        """).fetchall()

        members = c.execute("""
            SELECT id,username,role,full_name,student_code,email,phone,is_admin
            FROM users
            ORDER BY is_admin DESC, role, full_name
        """).fetchall()
        all_exams=c.execute("""SELECT e.id,e.title,e.teacher_id,u.full_name AS owner_name
            FROM exams e LEFT JOIN users u ON u.id=e.teacher_id
            WHERE COALESCE(e.is_practice,0)=0 ORDER BY e.id DESC LIMIT 300""").fetchall()
        teachers=c.execute("SELECT id,full_name,username FROM users WHERE role='teacher' AND COALESCE(is_admin,0)=0 ORDER BY full_name").fetchall()

        intro_row = c.execute(
            "SELECT setting_value FROM site_settings WHERE setting_key='homepage_intro'"
        ).fetchone()

        news = c.execute("""
            SELECT * FROM news_posts ORDER BY id DESC
        """).fetchall()

    return render_template(
        "admin_dashboard.html",
        requests=requests,
        members=members,
        all_exams=all_exams,
        teachers=teachers,
        homepage_intro=(intro_row["setting_value"] if intro_row else ""),
        news=news
    )


@app.post('/admin/exam/<int:exam_id>/transfer')
@admin_required
def admin_transfer_exam(exam_id):
    try: teacher_id=int(request.form.get('teacher_id',''))
    except (ValueError,TypeError): abort(400)
    with db() as c:
        teacher=c.execute("SELECT id FROM users WHERE id=? AND role='teacher' AND COALESCE(is_admin,0)=0",(teacher_id,)).fetchone()
        exam=c.execute("SELECT id FROM exams WHERE id=?",(exam_id,)).fetchone()
        if not teacher or not exam: abort(404)
        if c.execute('SELECT 1 FROM student_exams WHERE exam_id=? LIMIT 1',(exam_id,)).fetchone():
            flash('Đề đã có bài làm học sinh; không thể tự ý giao lại vì kết quả thuộc lớp cũ.')
            return redirect(url_for('admin_dashboard'))
        c.execute('UPDATE exams SET teacher_id=? WHERE id=?',(teacher_id,exam_id))
        c.execute('DELETE FROM exam_students WHERE exam_id=?',(exam_id,))
        c.execute('UPDATE exam_invites SET active=0 WHERE exam_id=?',(exam_id,))
    flash('Đã giao đề cho giáo viên mới. Các link chia sẻ cũ đã bị vô hiệu hóa.')
    return redirect(url_for('admin_dashboard'))


@app.post("/admin/request/<int:req_id>/approve")
@admin_required
def admin_approve_request(req_id):
    with db() as c:
        req = c.execute(
            "SELECT * FROM registration_requests WHERE id=?",
            (req_id,)
        ).fetchone()

        if not req:
            abort(404)

        if req["status"] != "pending":
            flash("Yêu cầu này đã được xử lý.")
            return redirect(url_for("admin_dashboard"))

        role = req["requested_role"]
        approved_user_id = None

        # HỌC SINH: nếu Mã HS đã có từ Excel thì ưu tiên tài khoản đó.
        if role == "student" and req["student_code"]:
            old = c.execute("""
                SELECT * FROM users
                WHERE student_code=? AND role='student'
            """,(req["student_code"],)).fetchone()

            if old:
                approved_user_id = old["id"]
                # Không đổi password ở đây nếu account đến từ Excel:
                # Excel luôn ưu tiên password = mã học sinh.
                c.execute("""
                    UPDATE users
                    SET full_name=?,
                        email=COALESCE(NULLIF(?,''),email),
                        phone=COALESCE(NULLIF(?,''),phone)
                    WHERE id=?
                """,(
                    req["full_name"],req["email"] or "",
                    req["phone"] or "",approved_user_id
                ))

        if approved_user_id is None:
            if c.execute(
                "SELECT 1 FROM users WHERE username=?",
                (req["username"],)
            ).fetchone():
                flash("Không thể duyệt: tên đăng nhập đã tồn tại.")
                return redirect(url_for("admin_dashboard"))

            cur = c.execute("""
                INSERT INTO users(
                    username,password_hash,role,full_name,
                    student_code,email,phone,is_admin
                ) VALUES(?,?,?,?,?,?,?,0)
            """,(
                req["username"],req["password_hash"],role,req["full_name"],
                req["student_code"] or None,req["email"] or None,
                req["phone"] or None
            ))
            approved_user_id = cur.lastrowid

        c.execute("""
            UPDATE registration_requests
            SET status='approved',
                decided_at=CURRENT_TIMESTAMP,
                approved_user_id=?
            WHERE id=?
        """,(approved_user_id,req_id))

    flash("Đã chấp nhận thành viên.")
    return redirect(url_for("admin_dashboard"))


@app.post("/admin/request/<int:req_id>/reject")
@admin_required
def admin_reject_request(req_id):
    with db() as c:
        c.execute("""
            UPDATE registration_requests
            SET status='rejected', decided_at=CURRENT_TIMESTAMP
            WHERE id=? AND status='pending'
        """,(req_id,))
    flash("Đã từ chối yêu cầu đăng ký.")
    return redirect(url_for("admin_dashboard"))


@app.post("/admin/request/<int:req_id>/delete")
@admin_required
def admin_delete_request(req_id):
    with db() as c:
        c.execute(
            "DELETE FROM registration_requests WHERE id=?",
            (req_id,)
        )
    flash("Đã xóa yêu cầu đăng ký.")
    return redirect(url_for("admin_dashboard"))


# ============================================================
# V5.13.0 — ADMIN XÓA THÀNH VIÊN SAU KHI ĐÃ CHẤP NHẬN
# ============================================================
@app.post("/admin/member/<int:user_id>/delete")
@admin_required
def admin_delete_member(user_id):
    if user_id == session.get("user_id"):
        flash("Admin không thể tự xóa tài khoản đang đăng nhập.")
        return redirect(url_for("admin_dashboard"))

    with db() as c:
        member = c.execute(
            "SELECT * FROM users WHERE id=?",
            (user_id,)
        ).fetchone()

        if not member:
            abort(404)

        if int(member["is_admin"] or 0):
            flash("Không xóa tài khoản quản trị bằng chức năng này.")
            return redirect(url_for("admin_dashboard"))

        # Never orphan teacher-owned exams or silently expose data after
        # deleting a teacher. Transfer/remove teacher resources explicitly.
        if member['role']=='teacher':
            owned=sum(c.execute(f'SELECT COUNT(*) FROM {table} WHERE {col}=?',
                                (user_id,)).fetchone()[0] for table,col in (
                ('exams','teacher_id'),('documents','uploaded_by'),
                ('shuffle_sources','uploaded_by'),('bank_sources','uploaded_by'),
                ('document_folders','created_by'),('bank_folders','created_by')))
            if owned:
                flash('Giáo viên còn đề hoặc học liệu. Hãy chuyển giao/xóa dữ liệu trước khi xóa tài khoản.')
                return redirect(url_for('admin_dashboard'))
        # student_exams/student_answers có ON DELETE CASCADE theo users.
        c.execute("DELETE FROM users WHERE id=?", (user_id,))

        # Giữ lịch sử yêu cầu nhưng bỏ liên kết user đã bị xóa.
        c.execute("""
            UPDATE registration_requests
            SET approved_user_id=NULL
            WHERE approved_user_id=?
        """,(user_id,))

    flash("Đã xóa thành viên.")
    return redirect(url_for("admin_dashboard"))



# ============================================================
# V5.14.0 — ADMIN VIẾT NỘI DUNG GIỚI THIỆU WEBSITE
# ============================================================
@app.post("/admin/site-intro")
@admin_required
def admin_update_site_intro():
    intro = (request.form.get("homepage_intro") or "").strip()
    with db() as c:
        c.execute("""
            INSERT INTO site_settings(setting_key,setting_value)
            VALUES('homepage_intro',?)
            ON CONFLICT(setting_key) DO UPDATE SET setting_value=excluded.setting_value
        """,(intro,))
    flash("Đã cập nhật phần giới thiệu Trang chủ.")
    return redirect(url_for("admin_dashboard"))


# ============================================================

# ============================================================
# V5.17.0 — ADMIN ĐĂNG TIN TỨC / ẢNH / VIDEO LÊN TRANG CHỦ
# ============================================================
LEGACY_NEWS_MEDIA_DIR = BASE / "static" / "news_media"
NEWS_MEDIA_DIR = DATA_ROOT / "news_media"
NEWS_MEDIA_DIR.mkdir(parents=True, exist_ok=True)
LEGACY_NEWS_MEDIA_DIR.mkdir(parents=True, exist_ok=True)
NEWS_IMAGE_EXT = {"jpg","jpeg","png","webp","gif"}
NEWS_VIDEO_EXT = {"mp4","webm","ogg","mov"}


def _news_media_path(filename):
    """Resolve persistent news media and remain compatible with older builds."""
    safe = Path(str(filename or "")).name
    if not safe or safe != str(filename or ""):
        return None
    for folder in (NEWS_MEDIA_DIR, LEGACY_NEWS_MEDIA_DIR):
        candidate = folder / safe
        if candidate.is_file():
            return candidate
    return None


@app.get("/uploaded/news/<path:filename>")
def news_media_file(filename):
    path = _news_media_path(filename)
    if not path:
        abort(404)
    return send_from_directory(path.parent, path.name)

@app.post("/admin/news/create")
@admin_required
def admin_create_news():
    title = normalize_text(request.form.get("title"))
    content = (request.form.get("content") or "").strip()
    image = request.files.get("image")
    video = request.files.get("video")

    if not title or not content:
        flash("Tin tức cần có tiêu đề và nội dung.")
        return redirect(url_for("admin_dashboard"))

    image_name = None
    video_name = None

    if image and image.filename:
        image_ext = ext(image.filename).lower()
        if image_ext not in NEWS_IMAGE_EXT:
            flash("Ảnh chỉ nhận JPG/JPEG/PNG/WEBP/GIF.")
            return redirect(url_for("admin_dashboard"))
        image_name = uuid.uuid4().hex + "_" + secure_filename(image.filename)
        image.save(NEWS_MEDIA_DIR / image_name)

    if video and video.filename:
        video_ext = ext(video.filename).lower()
        if video_ext not in NEWS_VIDEO_EXT:
            flash("Video chỉ nhận MP4/WEBM/OGG/MOV.")
            return redirect(url_for("admin_dashboard"))
        video_name = uuid.uuid4().hex + "_" + secure_filename(video.filename)
        video.save(NEWS_MEDIA_DIR / video_name)

    with db() as c:
        c.execute("""
            INSERT INTO news_posts(
                title,content,image_name,video_name,created_by
            ) VALUES(?,?,?,?,?)
        """,(
            title,content,image_name,video_name,session["user_id"]
        ))

    flash("Đã đăng tin tức lên Trang chủ.")
    return redirect(url_for("admin_dashboard"))


@app.post("/admin/news/<int:news_id>/delete")
@admin_required
def admin_delete_news(news_id):
    with db() as c:
        post = c.execute(
            "SELECT * FROM news_posts WHERE id=?",
            (news_id,)
        ).fetchone()
        if not post:
            abort(404)
        c.execute("DELETE FROM news_posts WHERE id=?", (news_id,))

    for name in (post["image_name"], post["video_name"]):
        if name:
            try:
                p = _news_media_path(name)
                if p and p.is_file():
                    p.unlink()
            except Exception:
                pass

    flash("Đã xóa tin tức.")
    return redirect(url_for("admin_dashboard"))


# V5.13.0 — TRANG LIÊN HỆ
# ============================================================
@app.get("/contact")
def contact():
    return render_template("contact.html")


# ============================================================
# V5.23.0 — NHÓM TIỆN ÍCH
# ============================================================
PDF_MATH_FONT_RE = re.compile(
    r"(cambria\s*math|symbol|stix|math|mt\s*extra|euclid|"
    r"cmmi|cmsy|cmr\d*|cmex|msbm|timesnewromanps-italicmt)",
    re.I,
)
PDF_MATH_SYMBOL_RE = re.compile(
    r"[=±×÷√∫∑∏∞≤≥≠≈∈∉∪∩⊂⊆∠πθαβγδλμσω∆∇→↔^⁰¹²³⁴⁵⁶⁷⁸⁹₀₁₂₃₄₅₆₇₈₉\uf000-\uf8ff]"
)
PDF_SUPERSCRIPTS = str.maketrans(
    "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾", "0123456789+-=()"
)
PDF_SUBSCRIPTS = str.maketrans(
    "₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎", "0123456789+-=()"
)


def _pdf_span_is_math(span):
    raw_text = str(span.get("text") or "")
    font = str(span.get("font") or "")
    font_lower = font.lower()
    if "yhcmex" in font.lower() and raw_text == "\xa0":
        return True
    text = raw_text.strip()
    if not text:
        return False
    # Times New Roman *italic* is also used for exam instructions and
    # Vietnamese prose; its PDF font name alone is not a math signal.
    if "timesnewromanps-italicmt" in font_lower:
        words = re.findall(r"[A-Za-zÀ-ỹ]{2,}", text)
        if len(words) >= 2 or re.search(r"[À-ỹ]", text):
            return False
    # Units in chemistry equations need to remain prose ("= -285,66 kJ");
    # turning the whole span into TeX produces merged letters or invalid units.
    if ("timesnewromanps" in font_lower
            and re.search(r"\b(?:kJ|kcal|mol|gam)\b", text, re.I)
            and re.search(r"\s", text)):
        return False
    if PDF_MATH_FONT_RE.search(font):
        return True
    # MS Word chemistry PDF: VnTime is used for the H, r, 298, 0
    # of Δ_r H^0_298. The prose is Times New Roman, not VnTime.
    if "vntime" in font_lower and re.fullmatch(r"(?:H|r|[0-9]{1,4})", text):
        return True
    # In chemistry PDFs a 2+ charge, stoichiometric index or exponent is
    # frequently typeset in *regular* Times New Roman at 8.5pt, while the
    # surrounding sentence is 13pt. Keep those compact, reduced-size glyphs
    # in the math stream without classifying the entire sentence as math.
    if (
        float(span.get("size") or 12) < 9.5
        and re.fullmatch(r"[0-9+\-]{1,4}", text)
        and "timesnewromanps" in font_lower
    ):
        return True
    if PDF_MATH_SYMBOL_RE.search(text):
        # V5.25.3: độ dài 28 ký tự KHÔNG phải tín hiệu toán học.
        # "Giá trị của Δ", "Cho phản ứng →" là văn xuôi kèm ký hiệu,
        # không được chuyển cả cụm thành một lệnh TeX (sẽ làm mất dấu cách).
        words = re.findall(r"[A-Za-zÀ-ỹ]{2,}", text)
        if len(words) >= 2 and re.search(r"[À-ỹ]|\s", text):
            return False
        return len(words) <= 1
    # Word may export the variables of one equation with the surrounding
    # Palatino italic font while operators, brackets and digits use SymbolMT.
    # Treat only compact ASCII italic runs as math; explanatory italic prose
    # still contains spaces/diacritics and remains ordinary text.  This keeps
    # a, b, S.ABC, x and script-sized exponents on the equation baseline.
    flags = int(span.get("flags") or 0)
    if (
        ("italic" in font_lower or flags & 2)
        and len(text) <= 28
        and re.fullmatch(r"[A-Za-z0-9\s.,;:+\-*/]+", text)
        and (len(re.findall(r"[A-Za-z]{2,}", text)) <= 1
             or text.lower().strip() in {"sin x", "cos x", "log x", "ln x"})
    ):
        return True
    # Upright function names are often emitted with the prose font even when
    # they are part of the same visual formula.
    if text.lower() in {
        "sin", "cos", "tan", "cot", "log", "ln", "lim", "min", "max"
    }:
        return True
    return bool(
        re.search(r"[A-Za-z0-9)]\s*[=<>]\s*[-+(A-Za-z0-9]", text)
        and not re.search(r"[A-Za-zÀ-ỹ]{3,}\s+[A-Za-zÀ-ỹ]{3,}", text)
    )


def _pdf_recover_span_text(span):
    """Khôi phục dấu cách bị mất trong lớp chữ của PDF.

    Một số bộ tạo PDF bỏ ký tự space nhưng vẫn chừa đúng một khoảng trống
    theo tọa độ.  Chỉ áp dụng cho span văn bản thường; công thức vẫn được xử
    lý từng glyph bởi bộ dựng LaTeX riêng nên không bị chèn khoảng trắng sai.
    """
    chars = list(span.get("chars") or [])
    if not chars:
        return str(span.get("text") or "")

    font_size = max(1.0, float(span.get("size") or 12))
    gap_threshold = max(1.5, font_size * .18)
    pieces = []
    previous_x1 = None
    previous_char = ""

    for char in chars:
        value = str(char.get("c") or "")
        bbox = char.get("bbox") or span.get("bbox") or (0, 0, 0, 0)
        x0 = float(bbox[0])
        if (
            previous_x1 is not None
            and value
            and not value.isspace()
            and previous_char
            and not previous_char.isspace()
            and x0 - previous_x1 > gap_threshold
            and value not in ".,;:!?%)]}"
            and previous_char not in "([{" 
        ):
            pieces.append(" ")
        pieces.append(value)
        previous_x1 = float(bbox[2])
        previous_char = value

    return "".join(pieces)


def pdf_math_text_to_latex(value):
    """Convert extractable PDF math characters into canonical LaTeX source."""
    text = str(value or "").strip()
    if not text:
        return ""

    text = text.replace("−", "-").replace("–", "-").replace("·", r"\cdot ")
    # Một số PDF TeX xuất dấu gạch phủ của ``\ne`` thành ký tự kết hợp riêng.
    text = text.replace("̸=", r"\ne ").replace("≠", r"\ne ")
    text = re.sub(r"(?<!\\)%", r"\\%", text)
    text = re.sub(r"(?<!\\)#", r"\\#", text)
    text = re.sub(r"⃗\s*([A-Za-zΑ-ω])", lambda match: r"\vec{" + match.group(1) + "}", text)
    text = re.sub(
        r"\(([^()]+)\)\s*[/⁄]\s*\(([^()]+)\)",
        lambda match: rf"\dfrac{{{match.group(1)}}}{{{match.group(2)}}}",
        text,
    )
    text = re.sub(
        r"(?<![\\\w}])([A-Za-z0-9]+)\s*[/⁄]\s*([A-Za-z0-9]+)",
        lambda match: rf"\dfrac{{{match.group(1)}}}{{{match.group(2)}}}",
        text,
    )
    text = re.sub(
        r"√\s*\(([^()]*)\)",
        lambda match: rf"\sqrt{{{match.group(1)}}}",
        text,
    )
    text = re.sub(
        r"√\s*([A-Za-z0-9]+)",
        lambda match: rf"\sqrt{{{match.group(1)}}}",
        text,
    )

    def superscript(match):
        return "^{" + match.group(0).translate(PDF_SUPERSCRIPTS) + "}"

    def subscript(match):
        return "_{" + match.group(0).translate(PDF_SUBSCRIPTS) + "}"

    text = re.sub(r"[⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾]+", superscript, text)
    text = re.sub(r"[₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎]+", subscript, text)

    replacements = {
        "≤": r"\le ", "≥": r"\ge ", "≠": r"\ne ", "≈": r"\approx ",
        "±": r"\pm ", "×": r"\times ", "÷": r"\div ", "∞": r"\infty ",
        "∈": r"\in ", "∉": r"\notin ", "∪": r"\cup ", "∩": r"\cap ",
        "⊂": r"\subset ", "⊆": r"\subseteq ", "∠": r"\angle ",
        "∫": r"\int ", "∑": r"\sum ", "∏": r"\prod ", "→": r"\to ",
        "↔": r"\leftrightarrow ", "π": r"\pi ", "θ": r"\theta ",
        "∼": r"\sim ", "△": r"\triangle ",
        "α": r"\alpha ", "β": r"\beta ", "γ": r"\gamma ",
        "δ": r"\delta ", "λ": r"\lambda ", "μ": r"\mu ",
        "σ": r"\sigma ", "ω": r"\omega ", "∆": r"\Delta ",
        "∇": r"\nabla ", "°": r"^{\circ}", "◦": r"^{\circ}",
        "′": "'", "⊥": r"\perp ",
        "\uf03e": ">", "\uf044": r"\Delta ", "\uf061": r"\alpha ",
        "\uf0a1": r"\mathbb{R}", "\uf0a2": "'", "\uf0a5": r"\infty ",
        "\uf0b0": r"^{\circ}", "\uf0c7": r"\cap ", "\uf0c8": r"\cup ",
        "½": r"\dfrac{1}{2}", "⅓": r"\dfrac{1}{3}",
        "⅔": r"\dfrac{2}{3}", "¼": r"\dfrac{1}{4}",
        "¾": r"\dfrac{3}{4}",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)

    text = re.sub(r"\\sqrt\{\}\s*([A-Za-z0-9]+)", r"\\sqrt{\1}", text)
    return normalise_latex(text) or text.strip()


def _set_word_run_font(run, name="Times New Roman", size=13.5):
    run.font.name = name
    run.font.size = Pt(size)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)


def _add_pdf_text_line(word, line, page_width):
    spans = sorted(
        line.get("spans", []),
        key=lambda item: float((item.get("bbox") or (0, 0, 0, 0))[0]),
    )
    if not spans:
        return False

    paragraph = word.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(1.5)
    line_bbox = line.get("bbox") or (0, 0, page_width, 0)
    center = (float(line_bbox[0]) + float(line_bbox[2])) / 2
    joined = "".join(str(span.get("text") or "") for span in spans).strip()
    if joined and len(joined) < 90 and abs(center - page_width / 2) < 18:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    previous_x1 = None
    for span in spans:
        raw_text = str(span.get("text") or "")
        if not raw_text:
            continue
        bbox = span.get("bbox") or (0, 0, 0, 0)
        font_size = max(8.0, min(20.0, float(span.get("size") or 13.5)))
        if (
            previous_x1 is not None
            and float(bbox[0]) - previous_x1 > max(1.5, font_size * 0.18)
            and not raw_text.startswith(" ")
        ):
            spacer = paragraph.add_run(" ")
            _set_word_run_font(spacer, size=font_size)

        if _pdf_span_is_math(span):
            latex = pdf_math_text_to_latex(raw_text)
            run = paragraph.add_run("$" + latex + "$" if latex else raw_text)
            _set_word_run_font(run, "Cambria Math", max(10.5, font_size))
        else:
            run = paragraph.add_run(raw_text)
            _set_word_run_font(run, size=font_size)
            flags = int(span.get("flags") or 0)
            run.bold = bool(flags & 16)
            run.italic = bool(flags & 2)
        previous_x1 = float(bbox[2])

    if not paragraph.text:
        paragraph._element.getparent().remove(paragraph._element)
        return False
    return True


def _pdf_fraction_rules(page):
    """Return short horizontal rules that visually separate PDF fractions."""
    rules = []
    try:
        for drawing in page.get_drawings():
            rect = drawing.get("rect")
            if rect is None:
                continue
            width = float(rect.x1 - rect.x0)
            height = abs(float(rect.y1 - rect.y0))
            if 4 <= width <= 120 and height <= 1.2:
                rules.append((float(rect.x0), (float(rect.y0) + float(rect.y1)) / 2, float(rect.x1)))
    except Exception:
        return []
    return rules


def _math_atom_source(atom):
    """Translate a PDF glyph with font-aware TeX corrections."""
    if atom.get("latex_override"):
        return str(atom["latex_override"])
    raw = str(atom.get("text") or "")
    font = str(atom.get("font") or "").lower()
    # Một số bản YHCMEX ánh xạ dấu căn thành NBSP. Không được ``strip`` trước
    # khi xét font, nếu không toàn bộ dấu căn biến mất khỏi công thức.
    if "yhcmex" in font and raw == "\xa0":
        return "__PDF_SQRT__"
    text = raw.strip()
    if not text:
        return ""

    # Legacy SymbolMT/Calibri symbol slots kept by Microsoft Word's PDF
    # exporter. Their Unicode values are private and must never reach MathJax.
    private_symbols = {
        "\uf03e": ">", "\uf044": r"\Delta ", "\uf061": r"\alpha ",
        "\uf0a1": r"\mathbb{R}", "\uf0a2": "'", "\uf0a5": r"\infty ",
        "\uf0b0": r"\circ ", "\uf0c7": r"\cap ", "\uf0c8": r"\cup ",
    }
    if text in private_symbols:
        return private_symbols[text]

    # Computer Modern extension fonts expose legacy glyph positions rather
    # than their Unicode meaning through some PDF ToUnicode maps.
    if text in {"°", "◦"}:
        return r"\circ "
    if text == "√":
        return "__PDF_SQRT__"
    if "cmex" in font:
        if text in {"R", "Z"}:
            return r"\int "
        # CMEX10 uses glyph slots ``h``/``i`` for the tall square brackets
        # around limits such as ``[-3/2;0]``.  Their PDF ToUnicode map exposes
        # the slot names instead of the visible delimiters.
        if text == "h":
            return "["
        if text == "i":
            return "]"
        if raw == "\x00":
            return "("
        if raw == "\x01":
            return ")"
        if text in {"p", "√"}:
            return "__PDF_SQRT__"
        if any("\uf8e0" <= char <= "\uf8ff" for char in text):
            return ""
    # YHCMEX is another Computer Modern extension encoding used by TeX.
    # Its ToUnicode table exposes scalable delimiters/accents as unrelated
    # Latin characters. Mapping by font avoids showing Å/ã/’ on the web.
    if "yhcmex" in font:
        if text == "Å":
            return "("
        if text == "ã":
            return ")"
        if text == "î":
            return "["
        if text == "ó":
            return "]"
        if text == "ß":
            return r"\{"
        if text == "™":
            return r"\}"
        if text == "»":
            return "__PDF_SQRT__"
        if text == "’":
            return ""
    if "msbm" in font and re.fullmatch(r"[RNCQZ]", text):
        return r"\mathbb{" + text + "}"
    if "cmsy" in font:
        if text == "\\":
            return r"\setminus"
        if text == "{":
            return r"\{"
        if text == "}":
            return r"\}"
    token = pdf_math_text_to_latex(raw)
    # Prevent adjacent glyphs from turning ``\le`` + ``x`` into the unknown
    # command ``\lex`` (the same applies to \in, \Delta, ...).
    if re.fullmatch(r"\\[A-Za-z]+", token or ""):
        token += " "
    return token


def _join_pdf_math_atoms(atoms):
    return "".join(
        _math_atom_source(atom)
        for atom in sorted(atoms, key=lambda item: (item["bbox"][0], item["cy"]))
    )


def _pdf_union_bbox(atoms):
    """Return the smallest PDF rectangle containing every supplied atom."""
    return (
        min(atom["bbox"][0] for atom in atoms),
        min(atom["bbox"][1] for atom in atoms),
        max(atom["bbox"][2] for atom in atoms),
        max(atom["bbox"][3] for atom in atoms),
    )


def _pdf_mask_symbol_to_latex(block):
    """Recognise Word's tiny transparent symbol pictures.

    Recent Microsoft Word PDFs sometimes paint ``R``, alpha and Delta as a
    2x2 colour image plus a high-resolution monochrome soft mask.  PyMuPDF
    exposes the visible mask and a private-use text placeholder separately.
    Keeping both creates a blank/oversized image next to duplicated text.
    The three shapes have stable normalized ink profiles, so convert only
    high-confidence matches and leave every real illustration untouched.
    """
    mask = block.get("mask")
    if not isinstance(mask, (bytes, bytearray)):
        return None
    if int(block.get("width") or 0) > 4 or int(block.get("height") or 0) > 4:
        return None
    try:
        image = Image.open(io.BytesIO(mask)).convert("L")
        bbox = image.getbbox()
        if not bbox:
            return None
        crop = image.crop(bbox)
        width, height = crop.size
        if width < 12 or height < 12:
            return None

        def ink_width(fraction):
            y = min(height - 1, max(0, round((height - 1) * fraction)))
            xs = [x for x in range(width) if crop.getpixel((x, y)) > 20]
            return (max(xs) - min(xs) + 1) if xs else 0

        top_ratio = ink_width(.10) / max(1, width)
        aspect = width / max(1, height)
        # Lower-case alpha is visibly wider than tall after transparent
        # padding is removed; R and Delta are close to square.
        if 1.12 <= aspect <= 1.55 and .25 <= top_ratio <= .70:
            return r"\alpha"
        if .82 <= aspect <= 1.20 and top_ratio <= .32:
            return r"\Delta"
        if .82 <= aspect <= 1.20 and top_ratio >= .48:
            return r"\mathbb{R}"
    except Exception:
        return None
    return None


def _pdf_drawing_radical_atoms(page):
    """Recover square-root signs exported only as vector drawing paths."""
    result = []
    try:
        # A root used in a small subscript (for example ``log_{sqrt(3)}``)
        # can have a vinculum shorter than the 4pt lower bound used for
        # fractions.  Inspect tiny horizontal drawing segments here as well;
        # the radical-shape checks below keep ordinary underlines out.
        rules = list(_pdf_fraction_rules(page))
        for drawing in page.get_drawings():
            rect = drawing.get("rect")
            if rect is None:
                continue
            width = float(rect.x1 - rect.x0)
            height = abs(float(rect.y1 - rect.y0))
            if 2.2 <= width <= 80 and height <= .9:
                candidate = (
                    float(rect.x0),
                    (float(rect.y0) + float(rect.y1)) / 2,
                    float(rect.x1),
                )
                if candidate not in rules:
                    rules.append(candidate)
        clusters = page.cluster_drawings()
    except Exception:
        return result
    for cluster in clusters:
        width = float(cluster.x1 - cluster.x0)
        height = float(cluster.y1 - cluster.y0)
        if not (7 <= width <= 80 and 5 <= height <= 55):
            continue
        candidates = [
            rule for rule in rules
            if float(cluster.x0) + width * .30 <= rule[0] <= float(cluster.x1) + 1
            and rule[2] <= float(cluster.x1) + 2
            and abs(rule[1] - float(cluster.y0)) <= max(2.5, height * .18)
            and float(cluster.x0) <= rule[0] - 2
            and float(cluster.y1) >= rule[1] + 4
        ]
        if not candidates:
            continue
        top_rule = max(candidates, key=lambda rule: rule[2] - rule[0])
        bbox = (
            float(cluster.x0), float(cluster.y0),
            float(top_rule[0]), float(cluster.y1),
        )
        result.append({
            "kind": "text", "math": True, "text": "√",
            "font": "PDF-drawing-radical", "size": min(18.0, height),
            "flags": 0, "bbox": bbox, "x": bbox[0],
            "cy": (bbox[1] + bbox[3]) / 2,
            "radical_rule": top_rule,
        })
    return result


def _pdf_radical_rule_keys(atoms, fraction_rules):
    """Return drawing rules used as radical vincula, not fraction bars.

    TeX draws both the top of ``sqrt`` and a fraction separator as a short
    horizontal rule. Treating a radical vinculum as a fraction bar can pull a
    numerator or denominator from the adjacent solution line. The radical
    glyph immediately to the left makes these rules unambiguous.
    """
    result = set()
    for root in atoms:
        if _math_atom_source(root) != "__PDF_SQRT__":
            continue
        rx1 = float(root["bbox"][2])
        root_cy = float(root["cy"])
        matches = [
            rule for rule in fraction_rules
            if rx1 - 1 <= rule[0] <= rx1 + 10
            and abs(root_cy - rule[1]) <= 14
        ]
        if matches:
            result.add(min(
                matches,
                key=lambda rule: (abs(root_cy - rule[1]), rule[0]),
            ))
    return result


def _pdf_prepare_math_atoms(atoms, fraction_rules=None):
    """Collapse TeX overlay glyphs before baseline grouping.

    TeX draws vectors and angle hats by placing an accent glyph over letters.
    Integral limits are also positioned far from the prose baseline. Reading
    those pieces independently creates ``-\to AB``, ``A’BC`` and stray limit
    characters. This pass combines each visual object into canonical LaTeX.
    """
    atoms = list(atoms or [])
    fraction_rules = list(fraction_rules or [])
    radical_rule_keys = _pdf_radical_rule_keys(atoms, fraction_rules)
    consumed = set()
    synthetic = []

    def make_atom(members, latex, cy=None, font="PDF-overlay"):
        bbox = _pdf_union_bbox(members)
        base = max(members, key=lambda atom: float(atom.get("size") or 0))
        return {
            "kind": "text", "math": True,
            "text": latex, "latex_override": latex,
            "font": font, "size": base.get("size") or 12,
            "flags": 0, "bbox": bbox, "x": bbox[0],
            "cy": float(cy if cy is not None else (bbox[1] + bbox[3]) / 2),
        }

    # SymbolMT encodes tall parentheses/brackets as separate top, middle and
    # bottom glyphs. Collapse each vertical stack before assigning rows;
    # otherwise a quotient such as (u/v)' is split across three paragraphs.
    delimiter_specs = (
        ({"\uf0e6", "\uf0e7", "\uf0e8"}, "("),
        ({"\uf0f6", "\uf0f7", "\uf0f8"}, ")"),
        ({"\uf0e9", "\uf0ea", "\uf0eb"}, "["),
        ({"\uf0f9", "\uf0fa", "\uf0fb"}, "]"),
    )
    for slots, latex in delimiter_specs:
        candidates = [
            atom for atom in atoms
            if id(atom) not in consumed
            and str(atom.get("text") or "").strip() in slots
            and "symbol" in str(atom.get("font") or "").lower()
        ]
        groups = []
        for atom in sorted(candidates, key=lambda item: (item["x"], item["cy"])):
            group = next((
                values for values in groups
                if abs(float(values[0]["x"]) - float(atom["x"])) <= 1.2
            ), None)
            if group is None:
                groups.append([atom])
            else:
                group.append(atom)
        for group in groups:
            distinct = {str(atom.get("text") or "").strip() for atom in group}
            if len(group) < 2 or len(distinct) < 2:
                continue
            consumed.update(id(atom) for atom in group)
            synthetic.append(make_atom(group, latex, font="PDF-delimiter"))

    # Wide angle hats are exported by YHCMEX as an overlaying apostrophe or,
    # in some TeX encodings, as a division-looking glyph.
    for accent in atoms:
        if id(accent) in consumed:
            continue
        font = str(accent.get("font") or "").lower()
        if (
            "yhcmex" not in font
            or str(accent.get("text") or "").strip() not in {"’", "÷"}
        ):
            continue
        ax0, _, ax1, _ = accent["bbox"]
        bases = [
            atom for atom in atoms
            if id(atom) not in consumed and atom is not accent
            and "cmmi" in str(atom.get("font") or "").lower()
            and re.fullmatch(r"[A-Za-zΑ-ω]", str(atom.get("text") or "").strip())
            and ax0 - 1 <= (atom["bbox"][0] + atom["bbox"][2]) / 2 <= ax1 + 1
            and abs(float(atom["cy"]) - float(accent["cy"])) <= 8
        ]
        bases.sort(key=lambda atom: atom["bbox"][0])
        letters = "".join(str(atom.get("text") or "").strip() for atom in bases)
        if len(letters) < 2:
            continue
        members = [accent] + bases
        consumed.update(id(atom) for atom in members)
        base_cy = sum(float(atom["cy"]) for atom in bases) / len(bases)
        synthetic.append(make_atom(
            members, r"\widehat{" + letters + "}", base_cy, "PDF-angle"
        ))

    # Vector arrows consist of one or more CMSY dash/arrow glyphs above the
    # complete segment name. Group neighbouring accent pieces first.
    arrow_atoms = [
        atom for atom in atoms
        if id(atom) not in consumed
        and "cmsy" in str(atom.get("font") or "").lower()
        and str(atom.get("text") or "").strip() in {"−", "→"}
    ]
    arrow_groups = []
    for atom in sorted(arrow_atoms, key=lambda item: (item["cy"], item["bbox"][0])):
        group = next((
            candidate for candidate in arrow_groups
            if abs(float(candidate[0]["cy"]) - float(atom["cy"])) <= 1
            and float(atom["bbox"][0]) - max(item["bbox"][2] for item in candidate) <= 1.5
            and float(atom["bbox"][0]) >= min(item["bbox"][0] for item in candidate) - 1
        ), None)
        if group is None:
            arrow_groups.append([atom])
        else:
            group.append(atom)

    for group in arrow_groups:
        if not any(str(atom.get("text") or "").strip() == "→" for atom in group):
            continue
        gx0 = min(atom["bbox"][0] for atom in group)
        gx1 = max(atom["bbox"][2] for atom in group)
        arrow_cy = sum(float(atom["cy"]) for atom in group) / len(group)
        bases = [
            atom for atom in atoms
            if id(atom) not in consumed and atom not in group
            and "cmmi" in str(atom.get("font") or "").lower()
            and re.fullmatch(r"[A-Za-zΑ-ω]", str(atom.get("text") or "").strip())
            and gx0 - 1 <= (atom["bbox"][0] + atom["bbox"][2]) / 2 <= gx1 + 1
            and 3 <= float(atom["cy"]) - arrow_cy <= 12
        ]
        bases.sort(key=lambda atom: atom["bbox"][0])
        letters = "".join(str(atom.get("text") or "").strip() for atom in bases)
        if not letters:
            continue
        members = group + bases
        consumed.update(id(atom) for atom in members)
        base_cy = sum(float(atom["cy"]) for atom in bases) / len(bases)
        command = r"\overrightarrow" if len(letters) > 1 else r"\vec"
        synthetic.append(make_atom(
            members, command + "{" + letters + "}", base_cy, "PDF-vector"
        ))

    # Bind scalable set braces to the expression they enclose. Without this,
    # a brace centred between numerator and denominator can be assigned to a
    # neighbouring heading while the fraction remains on the domain line.
    left_set_braces = [
        atom for atom in atoms
        if id(atom) not in consumed
        and "yhcmex" in str(atom.get("font") or "").lower()
        and str(atom.get("text") or "").strip() == "ß"
    ]
    right_set_braces = [
        atom for atom in atoms
        if id(atom) not in consumed
        and "yhcmex" in str(atom.get("font") or "").lower()
        and str(atom.get("text") or "").strip() == "™"
    ]
    for left in sorted(left_set_braces, key=lambda atom: atom["bbox"][0]):
        if id(left) in consumed:
            continue
        right = min((
            atom for atom in right_set_braces
            if id(atom) not in consumed
            and float(atom["bbox"][0]) > float(left["bbox"][2])
            and abs(float(atom["cy"]) - float(left["cy"])) <= 2.5
        ), key=lambda atom: float(atom["bbox"][0]), default=None)
        if right is None:
            continue
        inside = [
            atom for atom in atoms
            if id(atom) not in consumed and atom is not left and atom is not right
            and float(left["bbox"][2]) - 1 <= (
                float(atom["bbox"][0]) + float(atom["bbox"][2])
            ) / 2 <= float(right["bbox"][0]) + 1
            and abs(float(atom["cy"]) - float(left["cy"])) <= 20
        ]
        if not inside:
            continue
        local_rules = [
            rule for rule in fraction_rules
            if float(left["bbox"][2]) - 2 <= rule[0]
            and rule[2] <= float(right["bbox"][0]) + 2
        ]
        local_baseline = (
            min(local_rules, key=lambda rule: abs(rule[1] - float(left["cy"])))[1]
            if local_rules else float(left["cy"])
        )
        value = _pdf_formula_atoms_to_latex(
            inside, local_baseline, fraction_rules=local_rules
        )
        if not value:
            continue
        members = [left] + inside + [right]
        consumed.update(id(atom) for atom in members)
        synthetic.append(make_atom(
            members, r"\{" + value + r"\}", local_baseline, "PDF-set"
        ))

    # TeX places limits underneath textual operators, so sorting every glyph
    # by X produces strings such as ``m_{[0;1]}in`` and ``l_{x\to0}im``.
    # Recognise the roman operator word on its own baseline first, then bind
    # all of its smaller upper/lower glyphs into one canonical LaTeX atom.
    # This applies everywhere: question text, options and detailed solutions.
    operator_commands = {
        "lim": r"\lim", "min": r"\min", "max": r"\max",
    }
    roman_letters = [
        atom for atom in atoms
        if id(atom) not in consumed
        and "cmr" in str(atom.get("font") or "").lower()
        and re.fullmatch(r"[A-Za-z]", str(atom.get("text") or "").strip())
    ]
    operator_groups = []
    for start in sorted(roman_letters, key=lambda item: (item["cy"], item["bbox"][0])):
        if id(start) in consumed:
            continue
        for word in operator_commands:
            if str(start.get("text") or "").strip() != word[0]:
                continue
            chain = [start]
            for expected in word[1:]:
                previous = chain[-1]
                matches = [
                    atom for atom in roman_letters
                    if atom not in chain and id(atom) not in consumed
                    and str(atom.get("text") or "").strip() == expected
                    and abs(float(atom["cy"]) - float(start["cy"])) <= 1.2
                    and abs(float(atom.get("size") or 0) - float(start.get("size") or 0)) <= 1
                    and -.8 <= float(atom["bbox"][0]) - float(previous["bbox"][2]) <= 2.2
                ]
                if not matches:
                    chain = []
                    break
                chain.append(min(
                    matches,
                    key=lambda atom: abs(float(atom["bbox"][0]) - float(previous["bbox"][2])),
                ))
            if chain:
                operator_groups.append((word, chain))
                consumed.update(id(atom) for atom in chain)
                break

    for word, chain in operator_groups:
        base_cy = sum(float(atom["cy"]) for atom in chain) / len(chain)
        base_size = max(float(atom.get("size") or 12) for atom in chain)
        x0 = min(float(atom["bbox"][0]) for atom in chain)
        x1 = max(float(atom["bbox"][2]) for atom in chain)
        positioned = [
            atom for atom in atoms
            if id(atom) not in consumed
            and float(atom.get("size") or 12) <= base_size * .84
            and x0 - 10 <= (float(atom["bbox"][0]) + float(atom["bbox"][2])) / 2 <= x1 + 10
            and 2.5 <= abs(float(atom["cy"]) - base_cy) <= 28
        ]
        lower_candidates = [
            atom for atom in positioned if float(atom["cy"]) > base_cy
        ]
        # Choose the nearest compact baseline. This prevents an exponent in a
        # following large fraction (for example the 2 in x^2) from becoming a
        # bogus part of ``lim_{x\to...}`` merely because their X ranges overlap.
        if lower_candidates:
            nearest_delta = min(float(atom["cy"]) - base_cy for atom in lower_candidates)
            lower = [
                atom for atom in lower_candidates
                if float(atom["cy"]) - base_cy <= nearest_delta + 5.2
            ]
            # A fractional bound such as ``[-3/2;0]`` spans two small
            # baselines. Once its numerator touches a local fraction rule,
            # include the matching denominator as part of the same limit.
            for rule_x0, rule_y, rule_x1 in fraction_rules:
                if not any(
                    rule_x0 - 2 <= (float(atom["bbox"][0]) + float(atom["bbox"][2])) / 2 <= rule_x1 + 2
                    and abs(float(atom["cy"]) - rule_y) <= 6
                    for atom in lower
                ):
                    continue
                for atom in lower_candidates:
                    center = (float(atom["bbox"][0]) + float(atom["bbox"][2])) / 2
                    if rule_x0 - 2 <= center <= rule_x1 + 2 and abs(float(atom["cy"]) - rule_y) <= 12:
                        if atom not in lower:
                            lower.append(atom)
        else:
            lower = []
        latex = operator_commands[word]
        members = list(chain)

        def positioned_latex(group):
            if not group:
                return ""
            largest = max(float(atom.get("size") or 12) for atom in group)
            centers = sorted(
                float(atom["cy"]) for atom in group
                if float(atom.get("size") or 12) >= largest * .9
            )
            local_baseline = centers[len(centers) // 2]
            return _pdf_formula_atoms_to_latex(
                group, local_baseline, fraction_rules=fraction_rules
            )

        lower_tex = positioned_latex(lower)
        if lower_tex:
            latex += "_{" + lower_tex + "}"
            members.extend(lower)
        consumed.update(id(atom) for atom in members)
        synthetic.append(make_atom(members, latex, base_cy, "PDF-operator"))

    # CMEX exposes integral signs as Z/R. Bind their upper/lower limit atoms
    # before row assignment, including the relatively distant lower limit.
    for integral in atoms:
        if id(integral) in consumed:
            continue
        font = str(integral.get("font") or "").lower()
        raw = str(integral.get("text") or "").strip()
        if "cmex" not in font or raw not in {"R", "Z"}:
            continue
        ix0, _, ix1, _ = integral["bbox"]
        icy = float(integral["cy"])
        candidates = [
            atom for atom in atoms
            if id(atom) not in consumed and atom is not integral
            and float(atom.get("size") or 12) <= 9
            and ix0 - 2 <= (atom["bbox"][0] + atom["bbox"][2]) / 2 <= ix1 + 3
        ]
        upper = [atom for atom in candidates if 3 <= icy - float(atom["cy"]) <= 18]
        lower = [atom for atom in candidates if 8 <= float(atom["cy"]) - icy <= 32]
        upper_atom = min(upper, key=lambda atom: abs(float(atom["cy"]) - icy), default=None)
        lower_atom = min(lower, key=lambda atom: abs(float(atom["cy"]) - icy), default=None)
        latex = r"\int"
        members = [integral]
        if lower_atom is not None:
            lower_tex = _join_pdf_math_atoms([lower_atom])
            if lower_tex:
                latex += "_{" + lower_tex + "}"
                members.append(lower_atom)
        if upper_atom is not None:
            upper_tex = _join_pdf_math_atoms([upper_atom])
            if upper_tex:
                latex += "^{" + upper_tex + "}"
                members.append(upper_atom)
        if len(members) == 1:
            # Keep the following identifier separate from the control word:
            # ``\int{}N`` is valid TeX, while ``\intN`` is not.
            latex += "{}"
        consumed.update(id(atom) for atom in members)
        synthetic.append(make_atom(members, latex, icy, "PDF-integral"))

    # A radical sign and its vinculum are drawing/glyph overlays. Collapse the
    # complete visual object here, before prose baselines are chosen. This is
    # important when a radical sits in the numerator of a fraction on a line
    # immediately above it: treating the sign and radicand separately can send
    # one half to the previous option and produce ``sqrt{3AB}``.
    for root in atoms:
        if id(root) in consumed or _math_atom_source(root) != "__PDF_SQRT__":
            continue
        rx1 = float(root["bbox"][2])
        root_cy = float(root["cy"])
        rules = [
            rule for rule in fraction_rules
            if rx1 - 1 <= rule[0] <= rx1 + 10
            and abs(root_cy - rule[1]) <= 14
        ]
        if not rules and root.get("radical_rule"):
            rules = [tuple(root["radical_rule"])]
        if not rules:
            continue
        root_rule = min(
            rules, key=lambda rule: (abs(root_cy - rule[1]), rule[0])
        )
        rule_x0, rule_y, rule_x1 = root_rule
        radicand = [
            atom for atom in atoms
            if id(atom) not in consumed and atom is not root
            and rule_x0 - 1 <= (atom["bbox"][0] + atom["bbox"][2]) / 2 <= rule_x1 + 2
            and rule_y - 1 <= float(atom["cy"]) <= rule_y + 32
            and not (
                str(atom.get("text") or "").strip() in {".", ",", ";"}
                and float(atom["bbox"][0]) >= rule_x1 - .2
            )
        ]
        if not radicand:
            continue
        radicand.sort(key=lambda atom: (atom["bbox"][0], atom["cy"]))
        largest = max(float(atom.get("size") or 12) for atom in radicand)
        baseline_candidates = sorted(
            float(atom["cy"]) for atom in radicand
            if float(atom.get("size") or 12) >= largest * .9
        )
        radicand_baseline = baseline_candidates[len(baseline_candidates) // 2]
        inside = _pdf_formula_atoms_to_latex(
            radicand,
            radicand_baseline,
            fraction_rules=[
                rule for rule in fraction_rules
                if rule not in radical_rule_keys
            ],
        )
        if not inside:
            continue
        members = [root] + radicand
        consumed.update(id(atom) for atom in members)
        base_cy = sum(float(atom["cy"]) for atom in radicand) / len(radicand)
        synthetic.append(make_atom(
            members, r"\sqrt{" + inside + "}", base_cy, "PDF-radical"
        ))

    return [atom for atom in atoms if id(atom) not in consumed] + synthetic


def _pdf_collapse_fraction_atoms(atoms, fraction_rules):
    """Collapse fractions inside one already-bounded formula or cases block.

    Limiting the operation to a known formula prevents a nearby fraction bar
    from borrowing glyphs from the next prose line. This is especially useful
    for systems containing several ratios on the same row.
    """
    prepared = list(atoms or [])

    def local_baseline(group):
        biggest = max(float(atom.get("size") or 12) for atom in group)
        centers = sorted(
            float(atom["cy"]) for atom in group
            if float(atom.get("size") or 12) >= biggest * .9
        )
        return centers[len(centers) // 2]

    for rule_x0, rule_y, rule_x1 in sorted(
        fraction_rules or [], key=lambda rule: (rule[1], rule[0])
    ):
        numerator = [
            atom for atom in prepared
            if rule_x0 - 2 <= (atom["bbox"][0] + atom["bbox"][2]) / 2 <= rule_x1 + 2
            and rule_y - 18 <= float(atom["cy"]) < rule_y - .3
        ]
        denominator = [
            atom for atom in prepared
            if rule_x0 - 2 <= (atom["bbox"][0] + atom["bbox"][2]) / 2 <= rule_x1 + 2
            and rule_y + .3 < float(atom["cy"]) <= rule_y + 18
        ]
        if not numerator or not denominator:
            continue
        top = _pdf_formula_atoms_to_latex(
            numerator, local_baseline(numerator), fraction_rules=[]
        )
        bottom = _pdf_formula_atoms_to_latex(
            denominator, local_baseline(denominator), fraction_rules=[]
        )
        if not top or not bottom:
            continue
        members = numerator + denominator
        member_ids = {id(atom) for atom in members}
        prepared = [atom for atom in prepared if id(atom) not in member_ids]
        bbox = _pdf_union_bbox(members)
        base = max(members, key=lambda atom: float(atom.get("size") or 0))
        prepared.append({
            "kind": "text", "math": True,
            "text": rf"\dfrac{{{top}}}{{{bottom}}}",
            "latex_override": rf"\dfrac{{{top}}}{{{bottom}}}",
            "font": "PDF-fraction", "size": base.get("size") or 12,
            "flags": 0, "bbox": bbox, "x": bbox[0], "cy": float(rule_y),
        })
    return prepared


def _pdf_formula_atoms_to_latex(atoms, baseline, fraction_rules=None):
    """Rebuild a visual PDF formula from glyph coordinates.

    Horizontal rules supplied by the PDF drawing layer make fraction recovery
    deterministic. Font size and Y position recover superscripts, subscripts
    and integral limits. The result never contains display delimiters; the
    templates add them only while MathJax renders the formula.
    """
    atoms = [atom for atom in atoms if str(atom.get("text") or "").strip()]
    if not atoms:
        return ""

    fraction_rules = list(fraction_rules or [])

    # A radical extension glyph precedes its complete radicand. Handle it
    # before fractions so sqrt(m/n) becomes one nested expression.
    roots = [atom for atom in atoms if _math_atom_source(atom) == "__PDF_SQRT__"]
    if roots:
        root = min(roots, key=lambda atom: atom["bbox"][0])
        prefix = [atom for atom in atoms if atom["bbox"][2] <= root["bbox"][0] + .2]
        radicand = [atom for atom in atoms if atom["bbox"][0] >= root["bbox"][2] - 1]
        if radicand:
            left = _pdf_formula_atoms_to_latex(prefix, baseline, fraction_rules)
            inside = _pdf_formula_atoms_to_latex(radicand, baseline, fraction_rules)
            return left + r"\sqrt{" + inside + "}"

    # The PDF drawing layer contains the actual fraction bar. It avoids
    # confusing integral limits or a lone exponent with a numerator/denominator.
    for rule_x0, rule_y, rule_x1 in sorted(fraction_rules, key=lambda rule: rule[0]):
        formula_x0 = min(atom["bbox"][0] for atom in atoms)
        formula_x1 = max(atom["bbox"][2] for atom in atoms)
        formula_y0 = min(atom["bbox"][1] for atom in atoms)
        formula_y1 = max(atom["bbox"][3] for atom in atoms)
        if rule_x1 < formula_x0 - 2 or rule_x0 > formula_x1 + 2:
            continue
        if rule_y < formula_y0 - 2 or rule_y > formula_y1 + 2:
            continue
        numerator = [
            atom for atom in atoms
            if rule_x0 - 2 <= (atom["bbox"][0] + atom["bbox"][2]) / 2 <= rule_x1 + 2
            and atom["cy"] < rule_y - .3
        ]
        denominator = [
            atom for atom in atoms
            if rule_x0 - 2 <= (atom["bbox"][0] + atom["bbox"][2]) / 2 <= rule_x1 + 2
            and atom["cy"] > rule_y + .3
        ]
        if not numerator or not denominator:
            continue
        used = {id(atom) for atom in numerator + denominator}
        remaining = [atom for atom in atoms if id(atom) not in used]
        prefix = [atom for atom in remaining if atom["bbox"][2] <= rule_x0 + .5]
        suffix = [atom for atom in remaining if atom not in prefix]
        nested_rules = [
            rule for rule in fraction_rules
            if rule != (rule_x0, rule_y, rule_x1)
        ]
        left = _pdf_formula_atoms_to_latex(prefix, baseline, nested_rules) if prefix else ""
        def local_baseline(group):
            biggest = max(float(atom.get("size") or 12) for atom in group)
            bases = [
                atom["cy"] for atom in group
                if float(atom.get("size") or 12) >= biggest * .9
            ]
            bases.sort()
            return bases[len(bases) // 2]

        top = _pdf_formula_atoms_to_latex(
            numerator, local_baseline(numerator), nested_rules
        )
        bottom = _pdf_formula_atoms_to_latex(
            denominator, local_baseline(denominator), nested_rules
        )
        right = _pdf_formula_atoms_to_latex(suffix, baseline, nested_rules) if suffix else ""
        if top and bottom:
            return left + rf"\dfrac{{{top}}}{{{bottom}}}" + right

    ordered = sorted(atoms, key=lambda item: (item["bbox"][0], item["cy"]))
    base_size = max(float(atom.get("size") or 12) for atom in ordered)
    pieces = []
    raised = []
    lowered = []

    def flush_positioned():
        nonlocal raised, lowered
        if raised:
            pieces.append("^{" + _join_pdf_math_atoms(raised) + "}")
            raised = []
        if lowered:
            pieces.append("_{" + _join_pdf_math_atoms(lowered) + "}")
            lowered = []

    for atom in ordered:
        atom_size = float(atom.get("size") or base_size)
        if atom_size < base_size * 0.82 and atom["cy"] < baseline - 1.5:
            raised.append(atom)
            continue
        if atom_size < base_size * 0.82 and atom["cy"] > baseline + 1.5:
            lowered.append(atom)
            continue
        flush_positioned()
        token = _math_atom_source(atom)
        if token != "__PDF_SQRT__":
            pieces.append(token)
    flush_positioned()
    combined = "".join(pieces)
    combined = _canonicalize_pdf_vector_marks(combined)
    # Dùng dấu TeX cơ bản cho độ dài và góc vectơ. Chúng không phụ thuộc gói
    # mở rộng và vẫn an toàn nếu PDF từng ngắt công thức tại mép dòng.
    combined = re.sub(
        r"\|(\s*\\vec\{[^{}]+\}\s*)\|",
        lambda match: "|" + match.group(1).strip() + "|",
        combined,
    )
    combined = re.sub(
        r"\((\s*\\vec\{[^{}]+\}\s*,\s*\\vec\{[^{}]+\}\s*)\)",
        lambda match: "(" + match.group(1).strip() + ")",
        combined,
    )
    combined = re.sub(r"\\sqrt\{\}\s*([A-Za-z0-9]+)", r"\\sqrt{\1}", combined)
    combined = pdf_math_text_to_latex(combined)
    return normalise_latex(combined) or combined.strip()


def _pdf_promote_extrema_subscripts(normal_atoms, math_atoms):
    """Move small/lowered CĐ, CT markers next to ``y`` into the math layer.

    Vietnamese Word PDFs frequently export the variable with a TeX font but
    the extrema abbreviation with the prose font.  Geometry is authoritative:
    only an exact, smaller, lowered marker immediately following a math ``y``
    is promoted.  This avoids treating ordinary occurrences of CĐ/CT as math.
    """
    retained = []
    promoted = []
    candidates = list(math_atoms)
    for atom in normal_atoms:
        marker = str(atom.get("text") or "").strip()
        if marker.upper() not in {"CĐ", "CT", "CD"}:
            retained.append(atom)
            continue

        ax0 = float(atom["bbox"][0])
        marker_size = float(atom.get("size") or 12)
        previous_y = []
        for candidate in candidates:
            source = str(_math_atom_source(candidate) or "").strip()
            if source.lower() != "y":
                continue
            gap = ax0 - float(candidate["bbox"][2])
            base_size = float(candidate.get("size") or 12)
            baseline_drop = float(atom["cy"]) - float(candidate["cy"])
            if (
                -0.8 <= gap <= 4.5
                and marker_size <= base_size * 0.82
                and 1.2 <= baseline_drop <= 8.0
            ):
                previous_y.append((abs(gap), candidate))

        if not previous_y:
            retained.append(atom)
            continue

        promoted_atom = dict(atom)
        promoted_atom["math"] = True
        promoted_atom["text"] = marker.upper()
        promoted_atom["source"] = "pdf-extrema-subscript"
        promoted.append(promoted_atom)
        candidates.append(promoted_atom)

    return retained, list(math_atoms) + promoted


def _pdf_editable_rows(page):
    """Group PDF character spans by visual baseline before writing Word."""
    page_dict = page.get_text("rawdict", sort=True)
    normal_atoms = []
    math_atoms = []
    image_atoms = []
    symbol_image_atoms = []

    for block in page_dict.get("blocks", []):
        if block.get("type") == 0:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    chars = span.get("chars", [])
                    raw_text = "".join(str(char.get("c") or "") for char in chars)
                    probe = dict(span)
                    probe["text"] = raw_text
                    is_math = _pdf_span_is_math(probe)
                    if is_math:
                        for char in chars:
                            char_text = str(char.get("c") or "")
                            preserve_encoded_radical = (
                                "yhcmex" in str(span.get("font") or "").lower()
                                and char_text == "\xa0"
                            )
                            if not char_text.strip() and not preserve_encoded_radical:
                                continue
                            bbox = tuple(map(float, char.get("bbox", span.get("bbox"))))
                            math_atoms.append({
                                "kind": "text", "math": True, "text": char_text,
                                "font": span.get("font"), "size": span.get("size"),
                                "flags": span.get("flags"), "bbox": bbox,
                                "x": bbox[0], "cy": (bbox[1] + bbox[3]) / 2,
                            })
                    else:
                        text = _pdf_recover_span_text(span)
                    if not is_math and text.strip():
                        bbox = tuple(map(float, span.get("bbox", (0, 0, 0, 0))))
                        normal_atoms.append({
                            "kind": "text", "math": False, "text": text,
                            "font": span.get("font"), "size": span.get("size"),
                            "flags": span.get("flags"), "bbox": bbox,
                            "x": bbox[0], "cy": (bbox[1] + bbox[3]) / 2,
                            "source": "pdf-vector",
                        })
        elif block.get("type") == 1:
            bbox = tuple(map(float, block.get("bbox", (0, 0, 0, 0))))
            width = bbox[2] - bbox[0]
            height = bbox[3] - bbox[1]
            symbol_latex = _pdf_mask_symbol_to_latex(block)
            if symbol_latex:
                symbol_image_atoms.append({
                    "kind": "text", "math": True,
                    "text": symbol_latex, "latex_override": symbol_latex,
                    "font": "PDF-mask-symbol", "size": min(16.0, height),
                    "flags": 0, "bbox": bbox, "x": bbox[0],
                    "cy": (bbox[1] + bbox[3]) / 2,
                })
                continue
            if width >= 8 and height >= 8:
                image_atoms.append({
                    "kind": "image", "math": False, "bbox": bbox,
                    "x": bbox[0], "cy": (bbox[1] + bbox[3]) / 2,
                    "figure": width >= 50 and height >= 25,
                    "block": block,
                })

    if symbol_image_atoms:
        # Remove the private placeholder printed on top of the soft-mask
        # symbol.  The bitmap's PDF box is intentionally much wider than its
        # visible glyph, whereas the placeholder has the exact glyph box.
        # Copy that box first so alpha/Delta/R cannot overlap the following
        # equals sign or prose while rows are sorted.
        for symbol in symbol_image_atoms:
            sx0, sy0, sx1, sy1 = map(float, symbol["bbox"])
            placeholders = []
            for atom in math_atoms:
                raw = str(atom.get("text") or "")
                if not raw or not all(
                    0xE000 <= ord(char) <= 0xF8FF for char in raw
                ):
                    continue
                cx = (float(atom["bbox"][0]) + float(atom["bbox"][2])) / 2
                cy = (float(atom["bbox"][1]) + float(atom["bbox"][3])) / 2
                if sx0 <= cx <= sx1 and sy0 <= cy <= sy1:
                    placeholders.append(atom)
            if placeholders:
                placeholder = min(placeholders, key=lambda atom: (
                    abs(float(atom["bbox"][0]) - sx0),
                    abs(float(atom["cy"]) - float(symbol["cy"])),
                ))
                symbol["bbox"] = tuple(map(float, placeholder["bbox"]))
                symbol["x"] = symbol["bbox"][0]
                symbol["cy"] = (symbol["bbox"][1] + symbol["bbox"][3]) / 2
                symbol["size"] = float(placeholder.get("size") or 12)

        # The mask-derived LaTeX atom is now the single source of truth.
        def covered_private_placeholder(atom):
            raw = str(atom.get("text") or "")
            if not raw or not all(0xE000 <= ord(char) <= 0xF8FF for char in raw):
                return False
            cx = (float(atom["bbox"][0]) + float(atom["bbox"][2])) / 2
            cy = (float(atom["bbox"][1]) + float(atom["bbox"][3])) / 2
            return any(
                float(symbol["bbox"][0]) <= cx <= float(symbol["bbox"][2])
                and float(symbol["bbox"][1]) <= cy <= float(symbol["bbox"][3])
                for symbol in symbol_image_atoms
            )

        math_atoms = [
            atom for atom in math_atoms
            if not covered_private_placeholder(atom)
        ]
        math_atoms.extend(symbol_image_atoms)

    # Dots/semicolons inside formulas are sometimes exported with the prose
    # font (multiplication dots, S.ABC, interval separators).  Leaving them in
    # ``normal_atoms`` creates a false baseline that steals derivative primes
    # and splits exponents from their bases. Reclassify only compact
    # punctuation that is horizontally adjacent to math on this visual line.
    punctuation_math = []
    retained_normal = []
    for atom in normal_atoms:
        token = str(atom.get("text") or "").strip()
        if token not in {".", ",", ";", ":"}:
            retained_normal.append(atom)
            continue
        ax0, _, ax1, _ = map(float, atom["bbox"])
        neighbours = [
            candidate for candidate in math_atoms
            if abs(float(candidate["cy"]) - float(atom["cy"])) <= 13
            and (
                -1.5 <= float(candidate["bbox"][0]) - ax1 <= 6
                or -1.5 <= ax0 - float(candidate["bbox"][2]) <= 6
                or not (
                    float(candidate["bbox"][2]) < ax0
                    or float(candidate["bbox"][0]) > ax1
                )
            )
        ]
        if neighbours:
            promoted = dict(atom)
            promoted["math"] = True
            # A terminal dot can share its left edge with the exponent or a
            # closing delimiter. Preserve visual reading order: a^3. and
            # (ABCD)., while multiplication dots remain between two bases.
            trailing = [
                candidate for candidate in neighbours
                if ax0 - .5 <= float(candidate["bbox"][0]) <= ax1 + .5
                and (
                    float(candidate.get("size") or 12)
                        <= float(atom.get("size") or 12) * .82
                    or _math_atom_source(candidate) in {")", "]", r"\}", "'"}
                )
            ]
            if trailing:
                new_x = max(float(candidate["bbox"][2]) for candidate in trailing) + .1
                promoted["bbox"] = (
                    new_x, float(atom["bbox"][1]),
                    max(new_x + .2, float(atom["bbox"][2])),
                    float(atom["bbox"][3]),
                )
                promoted["x"] = new_x
            punctuation_math.append(promoted)
        else:
            retained_normal.append(atom)
    normal_atoms = retained_normal
    math_atoms.extend(punctuation_math)

    # In Vietnamese math documents, Word may encode y in a TeX font while
    # CĐ/CT is a smaller prose-font span. Recover the shared formula before
    # row grouping so it becomes y_{\mathrm{CĐ}} / y_{\mathrm{CT}} in LaTeX.
    normal_atoms, math_atoms = _pdf_promote_extrema_subscripts(
        normal_atoms, math_atoms
    )

    # Some Word PDFs have no radical character at all: the complete root sign
    # is a vector path. Recreate that missing math atom before fraction/radical
    # grouping so the radicand and nested fraction remain together.
    math_atoms.extend(_pdf_drawing_radical_atoms(page))

    all_fraction_rules = _pdf_fraction_rules(page)
    radical_rule_keys = _pdf_radical_rule_keys(math_atoms, all_fraction_rules)
    fraction_rules = [
        rule for rule in all_fraction_rules if rule not in radical_rule_keys
    ]

    # Collapse visual TeX overlays before rows are selected. In particular,
    # this keeps integral limits, radicals and vector/angle accents with their
    # formula instead of attaching pieces to adjacent answer lines.
    math_atoms = _pdf_prepare_math_atoms(math_atoms, all_fraction_rules)

    # Graphs, geometry figures and variation/statistics tables in a TeX PDF
    # are often pure drawing commands, not embedded images. Detect sizeable
    # connected drawing clusters, retain them as sharp rendered figures and
    # remove their internal labels from the text stream. Otherwise axis labels
    # can incorrectly become part of option D or of a nearby solution line.
    vector_figure_rects = []
    try:
        drawing_rects = []
        # Ignore page-spanning decorative paths before connected-component
        # clustering. One such path can otherwise connect independent tables,
        # graphs and borders into a page-sized "figure" that hides whole
        # questions (observed with Part III, questions 2 and 3).
        clusterable_drawings = []
        for drawing in page.get_drawings():
            rect = drawing.get("rect")
            if rect is None:
                continue
            width = float(rect.x1 - rect.x0)
            height = float(rect.y1 - rect.y0)
            if height > float(page.rect.height) * .55:
                continue
            if width > float(page.rect.width) * .94:
                continue
            clusterable_drawings.append(drawing)
        for cluster in page.cluster_drawings(drawings=clusterable_drawings):
            width = float(cluster.x1 - cluster.x0)
            height = float(cluster.y1 - cluster.y0)
            if width < 80 or height < 20:
                continue
            drawing_rects.append((
                float(cluster.x0), float(cluster.y0),
                float(cluster.x1), float(cluster.y1),
            ))

        # Chỉ nới khung tới đúng nhãn ngắn, tách biệt (A/B/C, x/y, số trên
        # trục...). Biên cố định 22–26 pt trước đây đã kéo cả câu văn phía trên
        # hoặc lời giải phía dưới vào ảnh ở Câu 3/Câu 9.
        label_atoms = normal_atoms + math_atoms

        def isolated_short_label(atom):
            token = str(atom.get("text") or "").strip()
            if not re.fullmatch(r"[A-Za-z0-9]", token):
                return False
            peers = [
                item for item in label_atoms
                if item is not atom and abs(float(item["cy"]) - float(atom["cy"])) <= 3.5
            ]
            if len(peers) > 2:
                return False
            return not any(
                item.get("kind") == "text" and not item.get("math")
                and len(str(item.get("text") or "").strip()) > 3
                for item in peers
            )

        assigned_labels = {index: [] for index in range(len(drawing_rects))}
        for atom in label_atoms:
            ax0, ay0, ax1, ay1 = map(float, atom["bbox"])
            cx = (ax0 + ax1) / 2
            cy = (ay0 + ay1) / 2
            candidates = []
            for index, (x0, y0, x1, y1) in enumerate(drawing_rects):
                dx = max(x0 - cx, 0.0, cx - x1)
                dy = max(y0 - cy, 0.0, cy - y1)
                if dx <= 18 and dy <= 18:
                    inside = dx == 0 and dy == 0
                    token = str(atom.get("text") or "").strip()
                    tight_axis_label = (
                        re.fullmatch(r"[A-Za-z0-9]", token) is not None
                        and min(dx, dy) == 0 and max(dx, dy) <= 12
                    )
                    if inside or isolated_short_label(atom) or tight_axis_label:
                        candidates.append((dx * dx + dy * dy, index))
            if candidates:
                assigned_labels[min(candidates)[1]].append(atom)

        for index, bbox in enumerate(drawing_rects):
            members = assigned_labels[index]
            x0, y0, x1, y1 = bbox
            # Chemistry PDFs often underline *an entire question+options row*
            # using a long decorative vector rule. The resulting cluster is
            # not a drawing: converting it to an image removes "Câu 1" and
            # all four A/B/C/D choices from the parser. Check the original
            # visible text before consuming any glyphs inside the cluster.
            covered_text = " ".join(
                str(atom.get("text") or "")
                for atom in normal_atoms
                if x0 - 2 <= (float(atom["bbox"][0]) + float(atom["bbox"][2])) / 2 <= x1 + 2
                and y0 - 2 <= float(atom["cy"]) <= y1 + 2
            )
            if (x1 - x0 > float(page.rect.width) * .75
                    and re.search(r"\bCâu\s*\d+\s*[.:]", covered_text, re.I)):
                continue
            if members:
                x0 = min(x0, min(float(atom["bbox"][0]) for atom in members))
                y0 = min(y0, min(float(atom["bbox"][1]) for atom in members))
                x1 = max(x1, max(float(atom["bbox"][2]) for atom in members))
                y1 = max(y1, max(float(atom["bbox"][3]) for atom in members))
            vector_figure_rects.append((
                max(0.0, x0 - 3), max(0.0, y0 - 3),
                min(float(page.rect.width), x1 + 3),
                min(float(page.rect.height), y1 + 3),
            ))
    except Exception:
        vector_figure_rects = []

    def atom_in_vector_figure(atom):
        cx = (float(atom["bbox"][0]) + float(atom["bbox"][2])) / 2
        cy = (float(atom["bbox"][1]) + float(atom["bbox"][3])) / 2
        return any(
            x0 <= cx <= x1 and y0 <= cy <= y1
            for x0, y0, x1, y1 in vector_figure_rects
        )

    if vector_figure_rects:
        # Các nét ngang nằm trong bảng/đồ thị không phải là gạch phân số.
        fraction_rules = [
            rule for rule in fraction_rules
            if not any(
                x0 <= (rule[0] + rule[2]) / 2 <= x1 and y0 <= rule[1] <= y1
                for x0, y0, x1, y1 in vector_figure_rects
            )
        ]
        normal_atoms = [atom for atom in normal_atoms if not atom_in_vector_figure(atom)]
        math_atoms = [atom for atom in math_atoms if not atom_in_vector_figure(atom)]
        image_atoms = [atom for atom in image_atoms if not atom_in_vector_figure(atom)]
        for bbox in vector_figure_rects:
            image_atoms.append({
                "kind": "image", "math": False, "bbox": bbox,
                "x": bbox[0], "cy": (bbox[1] + bbox[3]) / 2,
                "figure": True,
                "block": {
                    "bbox": bbox, "ext": "png", "vector_figure": True,
                },
            })

    # Large TeX braces are exported as several private-use CMEX glyphs. Merge
    # each brace and its right-hand rows into a single valid ``cases`` block so
    # the three equations cannot drift into neighbouring prose lines.
    brace_parts = []
    for atom in math_atoms:
        font = str(atom.get("font") or "").lower()
        text = str(atom.get("text") or "").strip()
        if (
            ("cmex" in font and any("\uf8e0" <= char <= "\uf8ff" for char in text))
            or ("yhcmex" in font and text in {"®", "ñ"})
            or ("cmex" in font and text == "(")
        ):
            brace_parts.append(atom)
    brace_groups = []
    for atom in sorted(brace_parts, key=lambda item: (item["x"], item["cy"])):
        group = next(
            (candidate for candidate in brace_groups if abs(candidate[0]["x"] - atom["x"]) <= 1.5),
            None,
        )
        if group is None:
            brace_groups.append([atom])
        else:
            group.append(atom)

    consumed = set()
    synthetic_cases = []
    sorted_brace_groups = sorted(brace_groups, key=lambda items: items[0]["x"])
    for group_index, group in enumerate(sorted_brace_groups):
        compact_token = str(group[0].get("text") or "").strip()
        compact_font = str(group[0].get("font") or "").lower()
        compact_brace = len(group) == 1 and (
            compact_token in {"®", "ñ", "("}
            or (
                "cmex" in compact_font
                and any("\uf8e0" <= char <= "\uf8ff" for char in compact_token)
            )
        )
        if len(group) < 2 and not compact_brace:
            continue
        brace_x1 = max(atom["bbox"][2] for atom in group)
        y0 = min(atom["bbox"][1] for atom in group) - 2
        bottom_pad = 16 if compact_brace else (12 if len(group) == 2 else 8)
        y1 = max(atom["bbox"][3] for atom in group) + bottom_pad
        group_y0 = min(atom["bbox"][1] for atom in group)
        group_y1 = max(atom["bbox"][3] for atom in group)
        following_overlapping_groups = [
            candidate for candidate in sorted_brace_groups[group_index + 1:]
            if min(atom["bbox"][1] for atom in candidate) <= group_y1 + 2
            and max(atom["bbox"][3] for atom in candidate) >= group_y0 - 2
        ]
        next_brace_x = (
            following_overlapping_groups[0][0]["x"]
            if following_overlapping_groups else None
        )
        right_edge = min(
            brace_x1 + 150,
            next_brace_x - 2 if next_brace_x is not None else brace_x1 + 150,
        )
        transition_atoms = [
            atom for atom in math_atoms
            if float(atom["bbox"][0]) > brace_x1
            and y0 <= float(atom["cy"]) <= y1
            and str(atom.get("text") or "").strip() in {"⇔", "⇒"}
        ]
        if transition_atoms:
            right_edge = min(
                right_edge,
                min(float(atom["bbox"][0]) for atom in transition_atoms) - 2,
            )
        normal_stops = [
            atom for atom in normal_atoms
            if float(atom["bbox"][0]) > brace_x1
            and y0 <= float(atom["cy"]) <= y1
            and str(atom.get("text") or "").strip().lower() in {"và", "hay"}
        ]
        if normal_stops:
            right_edge = min(
                right_edge,
                min(float(atom["bbox"][0]) for atom in normal_stops) - 2,
            )
        original_members = [
            atom for atom in math_atoms
            if atom not in group
            and brace_x1 - 1 <= atom["bbox"][0] <= right_edge
            and y0 <= atom["cy"] <= y1
        ]
        members = _pdf_collapse_fraction_atoms(original_members, fraction_rules)
        equation_rows = []
        for atom in sorted(members, key=lambda item: (item["cy"], item["x"])):
            target = next(
                (line for line in equation_rows if abs(line[0]["cy"] - atom["cy"]) <= 5.0),
                None,
            )
            if target is None:
                equation_rows.append([atom])
            else:
                target.append(atom)
        if len(equation_rows) < 2:
            continue
        equations = []
        for line in equation_rows:
            line_baseline = sorted(atom["cy"] for atom in line)[len(line) // 2]
            value = _pdf_formula_atoms_to_latex(
                line, line_baseline, fraction_rules=fraction_rules
            )
            if value:
                equations.append(value)
        if len(equations) < 2:
            continue
        # ``_pdf_collapse_fraction_atoms`` replaces numerator/denominator
        # glyphs with synthetic atoms. Consume the ORIGINAL glyphs here;
        # otherwise those source digits are assigned again to a nearby prose
        # row (for example a stray ``4`` before AB or ``OKOM1`` after a
        # system). The synthetic atoms only exist while building this cases
        # expression and are not members of ``math_atoms``.
        consumed.update(id(atom) for atom in group + original_members)
        all_members = group + original_members
        bbox = (
            min(atom["bbox"][0] for atom in all_members),
            min(atom["bbox"][1] for atom in all_members),
            max(atom["bbox"][2] for atom in all_members),
            max(atom["bbox"][3] for atom in all_members),
        )
        synthetic_cases.append({
            "kind": "text", "math": True,
            "text": "cases", "latex_override": (
                r"\begin{cases}" + r" \\ ".join(equations) + r"\end{cases}"
            ),
            "font": "PDF-cases", "size": 12, "flags": 0,
            "bbox": bbox, "x": bbox[0], "cy": (y0 + y1) / 2,
        })
    if consumed:
        math_atoms = [atom for atom in math_atoms if id(atom) not in consumed]
        math_atoms.extend(synthetic_cases)

    rows = []
    for atom in sorted(normal_atoms, key=lambda item: (item["cy"], item["x"])):
        target = next(
            (row for row in rows if not row.get("figure") and abs(atom["cy"] - row["cy"]) <= 3.5),
            None,
        )
        if target is None:
            rows.append({"cy": atom["cy"], "atoms": [atom], "figure": False})
        else:
            target["atoms"].append(atom)
            anchors = [item for item in target["atoms"] if not item.get("math")]
            target["cy"] = sum(item["cy"] for item in anchors) / len(anchors)

    for atom in sorted(math_atoms, key=lambda item: (item["cy"], item["x"])):
        candidates = [row for row in rows if not row.get("figure")]
        atom_center_x = (atom["bbox"][0] + atom["bbox"][2]) / 2
        nearest = min(candidates, key=lambda row: abs(atom["cy"] - row["cy"])) if candidates else None
        overlay_target = None
        if _math_atom_source(atom) == "'":
            # A derivative prime may be vertically equidistant from the
            # question line above and its answer line below. Its X overlap
            # with the closing parenthesis/base is the reliable association.
            lower_rows = [
                candidate_row for candidate_row in candidates
                if -1 <= float(candidate_row["cy"]) - float(atom["cy"]) <= 18
            ]
            if lower_rows:
                overlay_target = min(
                    lower_rows,
                    key=lambda candidate_row: float(candidate_row["cy"]) - float(atom["cy"]),
                )
            else:
                overlay_rows = []
                for candidate_row in candidates:
                    distance = min((
                        abs(atom_center_x - (
                            float(base["bbox"][0]) + float(base["bbox"][2])
                        ) / 2)
                        for base in candidate_row.get("atoms", [])
                        if base.get("math")
                        and float(base["bbox"][0]) - 2
                            <= atom_center_x <= float(base["bbox"][2]) + 5
                    ), default=None)
                    if distance is not None:
                        overlay_rows.append((distance, candidate_row))
                if overlay_rows:
                    overlay_target = min(overlay_rows, key=lambda item: item[0])[1]
        touching_rules = [] if atom.get("latex_override") else [
            rule for rule in fraction_rules
            if rule[0] - 2 <= atom_center_x <= rule[2] + 2
            and abs(atom["cy"] - rule[1]) <= 13
        ]
        root_rules = []
        if _math_atom_source(atom) == "__PDF_SQRT__":
            root_rules = [
                rule for rule in fraction_rules
                if atom["bbox"][2] - 1 <= rule[0] <= atom["bbox"][2] + 8
                and abs(atom["cy"] - rule[1]) <= 13
            ]
        fraction_target = None
        selected_fraction_rule = None
        if touching_rules:
            selected_fraction_rule = min(
                touching_rules, key=lambda item: abs(atom["cy"] - item[1])
            )
            fraction_target = next((
                row for row in candidates
                if row.get("fraction_line")
                and abs(float(row["cy"]) - selected_fraction_rule[1]) <= 1.0
            ), None)
            if fraction_target is None:
                # Denominators of fractions nested inside a larger fraction
                # may sit 20+ points from the prose baseline. Keep them in the
                # already-created parent fraction row instead of emitting a
                # stray line such as ``xx_2`` below the formula.
                sx0, sy, sx1 = selected_fraction_rule
                parent_rules = [
                    rule for rule in fraction_rules
                    if rule != selected_fraction_rule
                    and rule[0] - 1 <= sx0 and sx1 <= rule[2] + 1
                    and abs(rule[1] - sy) <= 24
                ]
                for parent in sorted(parent_rules, key=lambda rule: abs(rule[1] - sy)):
                    parent_row = next((
                        row for row in candidates
                        if row.get("fraction_line")
                        and abs(float(row["cy"]) - parent[1]) <= 2.5
                    ), None)
                    if parent_row is not None:
                        fraction_target = parent_row
                        break
            if fraction_target is None:
                nearby = min(
                    candidates,
                    key=lambda row: abs(selected_fraction_rule[1] - row["cy"]),
                    default=None,
                )
                if nearby is not None and abs(selected_fraction_rule[1] - nearby["cy"]) <= 12:
                    fraction_target = nearby
                    fraction_target["fraction_line"] = True
                else:
                    fraction_target = {
                        "cy": selected_fraction_rule[1], "atoms": [],
                        "figure": False, "fraction_line": True,
                    }
                    rows.append(fraction_target)
                    candidates.append(fraction_target)

        if overlay_target is not None:
            target = overlay_target
        elif fraction_target is not None:
            target = fraction_target
        elif candidates and root_rules:
            rule = min(root_rules, key=lambda item: abs(atom["cy"] - item[1]))
            below = [row for row in candidates if row["cy"] >= rule[1]]
            target = min(below or candidates, key=lambda row: abs(rule[1] - row["cy"]))
        elif candidates and touching_rules and (nearest is None or abs(atom["cy"] - nearest["cy"]) > 3.5):
            # Numerator/denominator atoms belong to the baseline crossing the
            # fraction bar, even when another prose line is marginally nearer.
            rule = min(touching_rules, key=lambda item: abs(atom["cy"] - item[1]))
            target = min(candidates, key=lambda row: abs(rule[1] - row["cy"]))
        else:
            target = nearest
        attach_distance = 32 if selected_fraction_rule is not None else 18
        if target is not None and abs(atom["cy"] - target["cy"]) <= attach_distance:
            target["atoms"].append(atom)
        else:
            orphan = next(
                (row for row in rows if not row.get("figure") and abs(atom["cy"] - row["cy"]) <= 18),
                None,
            )
            if orphan:
                orphan["atoms"].append(atom)
            else:
                rows.append({"cy": atom["cy"], "atoms": [atom], "figure": False})

    def row_has_answer_label(row):
        for item in row.get("atoms", []):
            if item.get("kind") != "text" or item.get("math"):
                continue
            value = str(item.get("text") or "").strip()
            if re.fullmatch(r"[ABCDabcd][.)]?", value):
                return True
        return False

    def row_plain_text(row):
        return "".join(
            str(item.get("text") or "")
            for item in sorted(row.get("atoms", []), key=lambda value: value.get("x", 0))
            if item.get("kind") == "text" and not item.get("math")
        )

    for atom in image_atoms:
        if atom["figure"]:
            # Hình bên phải có thể cao hơn nhiều dòng. Nếu đỉnh hình nằm trong
            # đề bài nhưng tâm hình nằm dưới dòng A–D, sắp nó ngay trước dòng
            # phương án để build_questions() gắn hình vào stem thay vì đáp án D.
            overlapping_options = [
                row for row in rows
                if float(atom["bbox"][1]) - 2 <= row["cy"] <= float(atom["bbox"][3]) + 2
                and row_has_answer_label(row)
            ]
            if overlapping_options:
                option_y = min(row["cy"] for row in overlapping_options)
                atom["pdf_placement"] = "prompt"
                rows.append({
                    "cy": atom["cy"], "sort_y": option_y - .1,
                    "atoms": [atom], "figure": True,
                })
            elif float(atom["bbox"][0]) >= float(page.rect.width) * .5:
                # Hình đặt bên phải một câu trả lời ngắn phải đứng sau toàn bộ
                # đề bài, thay vì chen giữa hai dòng văn bản theo tâm-Y của ảnh.
                question_rows = [
                    row for row in rows
                    if float(atom["bbox"][1]) - 20 <= row["cy"] <= float(atom["bbox"][3]) + 2
                    and re.search(r"\bCâu\s*\d+\b", row_plain_text(row), re.I)
                ]
                if question_rows:
                    question_y = min(row["cy"] for row in question_rows)
                    solution_y = min((
                        row["cy"] for row in rows
                        if row["cy"] > question_y
                        and re.search(r"Lời\s*giải", row_plain_text(row), re.I)
                    ), default=float("inf"))
                    prompt_rows = [
                        row for row in rows
                        if question_y <= row["cy"] < solution_y
                    ]
                    last_prompt_y = max(
                        (row["cy"] for row in prompt_rows), default=atom["cy"]
                    )
                    sort_y = min(last_prompt_y + .1, solution_y - .1)
                    rows.append({
                        "cy": atom["cy"], "sort_y": sort_y,
                        "atoms": [atom], "figure": True,
                    })
                else:
                    rows.append({"cy": atom["cy"], "atoms": [atom], "figure": True})
            else:
                rows.append({"cy": atom["cy"], "atoms": [atom], "figure": True})
        else:
            candidates = [row for row in rows if not row.get("figure")]
            target = min(candidates, key=lambda row: abs(atom["cy"] - row["cy"])) if candidates else None
            if target is not None and abs(atom["cy"] - target["cy"]) <= 18:
                target["atoms"].append(atom)
            else:
                rows.append({"cy": atom["cy"], "atoms": [atom], "figure": True})

    for row in rows:
        row["fraction_rules"] = fraction_rules
    return sorted(rows, key=lambda row: row.get("sort_y", row["cy"]))


def _pdf_visual_order_atoms(atoms):
    """Order PDF atoms by their painted position, correcting tiny overlaps."""
    values = list(atoms or [])

    def visual_order(item):
        x = float(item["x"])
        if item.get("kind") == "text" and not item.get("math"):
            # Word may begin the next prose span slightly before a closing
            # delimiter/superscript that is nevertheless painted first.
            raw_text = str(item.get("text") or "")
            boundary_span = (
                bool(raw_text[:1].isspace())
                or raw_text.lstrip().startswith((".", ",", ";", ":"))
            )
            overlapping_math = [
                atom for atom in values
                if atom.get("math")
                and boundary_span
                and x - 1.5 <= float(atom["bbox"][0]) <= x + 5.5
                and float(atom["bbox"][2]) <= x + 10
            ]
            if overlapping_math:
                x = max(float(atom["bbox"][0]) for atom in overlapping_math) + .2
        return (x, float(item["cy"]))

    return sorted(values, key=visual_order)


def _add_pdf_editable_row(word, page, row, available_width_pt):
    if row.get("figure") and len(row["atoms"]) == 1:
        _add_pdf_image_block(word, page, row["atoms"][0]["block"], available_width_pt)
        return True

    paragraph = word.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(1.5)
    ordered = _pdf_visual_order_atoms(row["atoms"])
    formula_atoms = []
    previous_x1 = None
    normal_anchors = [atom for atom in ordered if atom["kind"] == "text" and not atom.get("math")]
    if normal_anchors:
        formula_baseline = row["cy"]
    else:
        math_centers = [atom["cy"] for atom in ordered if atom.get("math")]
        formula_baseline = (
            (min(math_centers) + max(math_centers)) / 2
            if math_centers else row["cy"]
        )

    def add_gap(x0, size=12):
        nonlocal previous_x1
        if previous_x1 is not None and x0 - previous_x1 > max(1.5, size * 0.18):
            spacer = paragraph.add_run(" ")
            _set_word_run_font(spacer, size=size)

    def flush_formula():
        nonlocal formula_atoms, previous_x1
        if not formula_atoms:
            return
        x0 = min(atom["bbox"][0] for atom in formula_atoms)
        x1 = max(atom["bbox"][2] for atom in formula_atoms)
        size = max(10.5, max(float(atom.get("size") or 12) for atom in formula_atoms))
        add_gap(x0, size)
        latex = _pdf_formula_atoms_to_latex(
            formula_atoms,
            formula_baseline,
            fraction_rules=row.get("fraction_rules", []),
        )
        run = paragraph.add_run("$" + latex + "$" if latex else "")
        _set_word_run_font(run, "Cambria Math", size)
        previous_x1 = x1
        formula_atoms = []

    for atom in ordered:
        if atom.get("math"):
            formula_atoms.append(atom)
            continue
        flush_formula()
        if atom["kind"] == "image":
            # Bitmap công thức không có lớp ký tự: giữ hình thay vì làm mất nó.
            _add_pdf_image_block(word, page, atom["block"], available_width_pt)
            continue
        size = max(8.0, min(20.0, float(atom.get("size") or 13.5)))
        add_gap(atom["bbox"][0], size)
        run = paragraph.add_run(atom.get("text") or "")
        _set_word_run_font(run, size=size)
        flags = int(atom.get("flags") or 0)
        run.bold = bool(flags & 16)
        run.italic = bool(flags & 2)
        previous_x1 = atom["bbox"][2]
    flush_formula()

    if not paragraph.text:
        paragraph._element.getparent().remove(paragraph._element)
        return False
    return True


def _add_pdf_image_block(word, page, block, available_width_pt):
    import pymupdf as fitz

    bbox = block.get("bbox") or (0, 0, page.rect.width, page.rect.height)
    width_pt = max(12.0, float(bbox[2]) - float(bbox[0]))
    height_pt = max(12.0, float(bbox[3]) - float(bbox[1]))
    if width_pt > available_width_pt:
        ratio = available_width_pt / width_pt
        width_pt = available_width_pt
        height_pt *= ratio

    image_bytes = block.get("image")
    if not image_bytes:
        pixmap = page.get_pixmap(
            matrix=fitz.Matrix(3, 3), clip=fitz.Rect(bbox), alpha=False
        )
        image_bytes = pixmap.tobytes("png")

    paragraph = word.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_after = Pt(2)
    try:
        paragraph.add_run().add_picture(
            io.BytesIO(image_bytes), width=Pt(width_pt), height=Pt(height_pt)
        )
    except Exception:
        # Chỉ định dạng Word không nhận trực tiếp mới được render lại thành PNG.
        pixmap = page.get_pixmap(
            matrix=fitz.Matrix(3, 3), clip=fitz.Rect(bbox), alpha=False
        )
        paragraph.add_run().add_picture(
            io.BytesIO(pixmap.tobytes("png")),
            width=Pt(width_pt),
            height=Pt(height_pt),
        )


def pdf_to_editable_word(pdf_bytes):
    """Convert PDF into editable text, canonical LaTeX and original images."""
    import pymupdf as fitz

    source = fitz.open(stream=pdf_bytes, filetype="pdf")
    if source.page_count < 1:
        source.close()
        raise ValueError("Tệp PDF không có trang nội dung.")

    output = io.BytesIO()
    word = Document()
    section = word.sections[0]
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    available_width = section.page_width - section.left_margin - section.right_margin
    available_width_pt = float(available_width) / 12700

    normal_style = word.styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style.font.size = Pt(13.5)
    normal_style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    word.core_properties.title = "Tài liệu chuyển từ PDF"
    word.core_properties.subject = "Văn bản chỉnh sửa được, LaTeX và ảnh gốc"

    try:
        for page_index, page in enumerate(source):
            if page_index > 0:
                word.add_page_break()

            rows = _pdf_editable_rows(page)
            content_count = 0
            for row in rows:
                if _add_pdf_editable_row(word, page, row, available_width_pt):
                    content_count += 1

            # Với PDF scan không có lớp chữ/ảnh tách được, giữ nguyên toàn trang.
            if content_count == 0:
                pixmap = page.get_pixmap(
                    matrix=fitz.Matrix(2.5, 2.5), alpha=False
                )
                paragraph = word.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                paragraph.add_run().add_picture(
                    io.BytesIO(pixmap.tobytes("png")), width=available_width
                )
    finally:
        source.close()

    word.save(output)
    output.seek(0)
    return output


def pdf_pages_to_word(pdf_bytes):
    """Backward-compatible name retained for V5.23.0 integrations."""
    return pdf_to_editable_word(pdf_bytes)


@app.route("/utilities/pdf-to-word", methods=["GET", "POST"])
@login_required()
def pdf_to_word():
    if request.method == "POST":
        uploaded = request.files.get("file")
        if not uploaded or not uploaded.filename:
            flash("Vui lòng chọn tệp PDF.")
            return redirect(request.url)
        if ext(uploaded.filename) != "pdf":
            flash("Tiện ích này chỉ nhận tệp PDF.")
            return redirect(request.url)

        try:
            raw = uploaded.read()
            if not raw.startswith(b"%PDF-"):
                raise ValueError("Tệp được chọn không phải PDF hợp lệ.")
            result = pdf_to_editable_word(raw)
        except Exception as exc:
            flash("Không thể chuyển tệp PDF: " + str(exc))
            return redirect(request.url)

        base_name = secure_filename(Path(uploaded.filename).stem) or "tai_lieu"
        return send_file(
            result,
            as_attachment=True,
            download_name=base_name + ".docx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
        )

    return render_template("pdf_to_word.html")


@app.get("/utilities")
def utilities_portal():
    return render_template("utilities_portal.html")


def _response_output_text(payload):
    direct = str(payload.get("output_text") or "").strip()
    if direct:
        return direct
    parts = []
    for item in payload.get("output", []):
        for content in item.get("content", []):
            if content.get("type") == "output_text" and content.get("text"):
                parts.append(str(content["text"]))
    return "\n".join(parts).strip()


def ask_math_teacher_ai(problem, grade="THPT", approach="step_by_step", image=None):
    """Ask the configured multimodal model to solve a mathematics problem."""
    api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("Chưa cấu hình OPENAI_API_KEY trên máy chủ.")

    instructions = (
        "Bạn là trợ lý chuyên môn Toán dành cho giáo viên phổ thông Việt Nam. "
        "Hãy kiểm tra dữ kiện, giải chính xác từng bước, nêu điều kiện và kết luận. "
        "Ưu tiên phương pháp phù hợp chương trình " + grade + ". "
        "Mọi công thức phải dùng LaTeX với \\( ... \\) hoặc \\[ ... \\]. "
        "Không dùng khối mã. Nếu đề thiếu hoặc ảnh không rõ, nói rõ phần cần bổ sung. "
        "Không bịa dữ kiện và luôn khuyến khích người học tự đối chiếu các bước."
    )
    if approach == "hint":
        instructions += " Chỉ đưa gợi ý theo từng bước, chưa nêu đáp số ngay."
    elif approach == "short":
        instructions += " Trình bày ngắn gọn nhưng không bỏ bước lập luận chính."

    content = [{
        "type": "input_text",
        "text": problem or "Hãy đọc và giải bài toán trong ảnh đính kèm.",
    }]
    if image:
        image_bytes, image_mime = image
        content.append({
            "type": "input_image",
            "image_url": (
                "data:" + image_mime + ";base64,"
                + base64.b64encode(image_bytes).decode("ascii")
            ),
            "detail": "high",
        })

    request_body = json.dumps({
        "model": os.environ.get("OPENAI_MODEL", "gpt-5.6-luna"),
        "instructions": instructions,
        "input": [{"role": "user", "content": content}],
        "max_output_tokens": 5000,
    }).encode("utf-8")
    api_request = urllib_request.Request(
        "https://api.openai.com/v1/responses",
        data=request_body,
        headers={
            "Authorization": "Bearer " + api_key,
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib_request.urlopen(api_request, timeout=150) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib_error.HTTPError as exc:
        app.logger.exception("OpenAI API trả về HTTP %s", exc.code)
        raise RuntimeError(
            "Dịch vụ AI từ chối yêu cầu. Hãy kiểm tra API key, model và hạn mức."
        ) from exc
    except urllib_error.URLError as exc:
        raise RuntimeError("Máy chủ chưa kết nối được với dịch vụ AI.") from exc

    answer = _response_output_text(payload)
    if not answer:
        raise RuntimeError("Dịch vụ AI chưa trả về nội dung lời giải.")
    return answer


@app.route("/utilities/teacher-ai", methods=["GET", "POST"])
def teacher_ai_utility():
    teacher_access = (
        session.get("role") == "teacher" or bool(session.get("is_admin"))
    )
    solution = None
    error = None
    problem = ""
    grade = "THPT"
    approach = "step_by_step"

    if request.method == "POST":
        if not teacher_access:
            return redirect(url_for("login", role="teacher", next=request.path))

        problem = (request.form.get("problem") or "").strip()[:12000]
        grade = request.form.get("grade") or "THPT"
        if grade not in {"THCS", "THPT"}:
            grade = "THPT"
        approach = request.form.get("approach") or "step_by_step"
        if approach not in {"step_by_step", "hint", "short"}:
            approach = "step_by_step"

        image_data = None
        uploaded_image = request.files.get("problem_image")
        if uploaded_image and uploaded_image.filename:
            raw_image = uploaded_image.read(8 * 1024 * 1024 + 1)
            if len(raw_image) > 8 * 1024 * 1024:
                error = "Ảnh bài toán không được lớn hơn 8 MB."
            else:
                try:
                    probe = Image.open(io.BytesIO(raw_image))
                    image_format = str(probe.format or "").upper()
                    probe.verify()
                    mime_map = {
                        "JPEG": "image/jpeg", "PNG": "image/png",
                        "WEBP": "image/webp", "GIF": "image/gif",
                    }
                    if image_format not in mime_map:
                        raise ValueError("Chỉ nhận JPG, PNG, WEBP hoặc GIF.")
                    image_data = (raw_image, mime_map[image_format])
                except Exception as exc:
                    error = "Ảnh bài toán không hợp lệ: " + str(exc)

        if not error and not problem and not image_data:
            error = "Hãy nhập đề bài hoặc tải ảnh bài toán."
        if not error:
            try:
                solution = ask_math_teacher_ai(
                    problem, grade=grade, approach=approach, image=image_data
                )
            except Exception as exc:
                error = str(exc)

    return render_template(
        "teacher_ai_info.html",
        teacher_access=teacher_access,
        ai_configured=bool(os.environ.get("OPENAI_API_KEY", "").strip()),
        ai_model=os.environ.get("OPENAI_MODEL", "gpt-5.6-luna"),
        solution=solution,
        error=error,
        problem=problem,
        grade=grade,
        approach=approach,
    )


@app.get("/utilities/student-games")
def student_games():
    return render_template("student_games.html")


@app.get("/health")
def health():
    return jsonify(ok=True, version="5.25.0")


# ============================================================
# V5.15.1 — TRANG CỔNG GIÁO VIÊN / HỌC SINH
# ------------------------------------------------------------
# Các chức năng không còn nằm trong menu xổ xuống.
# Menu chỉ dẫn đến trang Giáo viên/Học sinh.
# Các hộp chức năng nằm TRÊN TRANG WEB tương ứng.
# ============================================================
@app.get("/teacher-portal")
def teacher_portal():
    return render_template("teacher_portal.html")


@app.get("/student-portal")
def student_portal():
    return render_template("student_portal.html")


@app.route("/")
def index():
    """
    V5.17.0 — Trang chủ:
    - phần giới thiệu do Admin viết;
    - tin tức mới nhất do Admin đăng, gồm text / ảnh / video.
    """
    with db() as c:
        row = c.execute(
            "SELECT setting_value FROM site_settings WHERE setting_key='homepage_intro'"
        ).fetchone()
        news = c.execute("""
            SELECT * FROM news_posts
            ORDER BY id DESC
        """).fetchall()

    intro = row["setting_value"] if row else ""
    return render_template(
        "index.html",
        homepage_intro=intro,
        news=news
    )

@app.route("/login",methods=["GET","POST"])
def login():
    if request.method=="POST":
        u=normalize_text(request.form.get("username")); p=request.form.get("password","")
        with db() as c:
            row=c.execute("SELECT * FROM users WHERE username=?",(u,)).fetchone()
            # KHỐI ĐĂNG NHẬP HỌC SINH — cho phép dùng Mã học sinh làm tài khoản.
            if not row:
                row=c.execute("SELECT * FROM users WHERE student_code=? AND role='student'",(u,)).fetchone()
        if row and check_password_hash(row["password_hash"],p):
            is_admin = int(row["is_admin"] or 0) if "is_admin" in row.keys() else 0
            # Chống session fixation: cấp phiên mới sau khi xác thực thành công.
            session.clear()
            session.update(
                user_id=row["id"],
                username=row["username"],
                role=row["role"],
                full_name=row["full_name"],
                is_admin=is_admin
            )
            next_url = request.args.get("next") or request.form.get("next")
            if next_url and next_url.startswith("/") and not next_url.startswith("//"):
                return redirect(next_url)
            if is_admin:
                return redirect(url_for("admin_dashboard"))
            # V5.16.0 — đăng nhập xong vào thẳng khu vực chức năng.
            if row["role"] == "teacher":
                return redirect(url_for("teacher_portal"))
            if row["role"] == "student":
                return redirect(url_for("student_portal"))
            return redirect(url_for("index"))
        flash("Sai tài khoản hoặc mật khẩu.")
    return render_template("login.html")

@app.route('/account/change-password',methods=['GET','POST'])
@login_required()
def change_own_password():
    if request.method=='POST':
        current=request.form.get('current_password','')
        new=request.form.get('new_password','')
        repeat=request.form.get('repeat_password','')
        with db() as c:
            user=c.execute('SELECT password_hash FROM users WHERE id=?',(session['user_id'],)).fetchone()
            if not user or not check_password_hash(user['password_hash'],current):
                flash('Mật khẩu cũ không chính xác.')
            elif len(new)<10 or new!=repeat:
                flash('Mật khẩu mới cần ít nhất 10 ký tự và nhập lại chính xác.')
            else:
                c.execute('UPDATE users SET password_hash=? WHERE id=?',(generate_password_hash(new),session['user_id']))
                flash('Đã đổi mật khẩu.')
                return redirect(url_for('index'))
    return render_template('change_password.html')


@app.get("/logout")
def logout(): session.clear(); return redirect(url_for("login"))

@app.get("/teacher")
@login_required("teacher")
def teacher_dashboard():
    with db() as c:
        ensure_practice_exam_schema(c)
        exams=c.execute(
            "SELECT * FROM exams WHERE teacher_id=? AND COALESCE(is_practice,0)=0 ORDER BY id DESC",(session["user_id"],)
        ).fetchall()
        invite_rows=c.execute("SELECT * FROM exam_invites WHERE teacher_id=? AND active=1 ORDER BY created_at DESC",(session["user_id"],)).fetchall()
    links={}
    for inv in invite_rows: links.setdefault(inv['exam_id'],[]).append(inv['token'])
    return render_template("teacher_dashboard.html",exams=exams,exam_links=links,roster_exists=_teacher_roster_path(session["user_id"]).exists(),
      initial_accounts_exists=_teacher_roster_path(session["user_id"]).with_name(f"teacher_{int(session['user_id'])}_initial_accounts.csv").exists())

@app.route("/teacher/upload",methods=["GET","POST"])
@login_required("teacher")
def upload_exam():
    if request.method=="POST":
        f=request.files.get("file"); title=normalize_text(request.form.get("title")) or "Đề thi"
        if not f or not f.filename or ext(f.filename) not in ALLOWED:
            flash("Chỉ nhận DOCX hoặc PDF."); return redirect(request.url)
        saved=UPLOAD_DIR/(uuid.uuid4().hex+"_"+secure_filename(f.filename)); f.save(saved)
        try:
            # ============================================================
            # KHỐI ĐIỀU PHỐI UPLOAD — WORD/PDF XỬ LÝ RIÊNG
            # ============================================================
            source_ext=ext(f.filename)

            if source_ext=="docx":
                # WORD: LibreOffice HTML + MathType/image parser riêng.
                lines=extract_word_blocks(saved)
                questions=build_questions(lines)
                _sync_docx_question_figures_from_ooxml(questions, saved)
            else:
                # PDF: parser baseline-anchor chính.
                lines=extract_pdf_blocks(saved)
                questions=build_questions(lines)

                # Nếu PDF lạ không nhận diện được câu, tự động thử parser
                # dự phòng thay vì báo lỗi ngay cho giáo viên.
                if not questions:
                    lines=extract_pdf_blocks_fallback(saved)
                    questions=build_questions(lines)

            if not questions:
                raise RuntimeError(
                    "Không nhận diện được câu hỏi sau cả hai bộ xử lý PDF/Word."
                )
            pts=[float(str(request.form.get(k,'2.5')).replace(',','.')) for k in ('p1','p2','p3','p4')]
            if abs(sum(pts)-10)>0.001: raise RuntimeError("Tổng điểm 4 phần phải bằng 10.")
            # V5.11.1 — chọn phương án chấm cho TỪNG CÂU Phần II.
            p2_mode=int(request.form.get('p2_mode') or 1)
            if p2_mode not in (1,2): p2_mode=1
            # Mỗi câu Phần II tối đa 1 điểm; vì vậy tổng điểm Phần II
            # phải khớp số câu Phần II để tổng điểm không bị sai.
            n_p2=sum(1 for q in questions if q.get('section')==2)
            if abs(pts[1]-n_p2)>0.001:
                raise RuntimeError(f"Phần II có {n_p2} câu nên điểm Phần II phải bằng {n_p2} điểm (mỗi câu tối đa 1 điểm).")
            p2_items=[0.25,0.25,0.25,0.25]
            with db() as c:
                ensure_source_media_columns(c)
                cur=c.execute("INSERT INTO exams(title,source_file,source_stored_name,question_count,published,duration_minutes,part1_points,part2_points,part3_points,part4_points,p2_item_a,p2_item_b,p2_item_c,p2_item_d,part2_scoring_mode,anti_cheat,max_violations) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(title,f.filename,saved.name,len(questions),0,int(request.form.get('duration') or 90),*pts,*p2_items,p2_mode,1 if request.form.get('anti_cheat') else 0,int(request.form.get('max_violations') or 3)))
                eid=cur.lastrowid
                c.execute("UPDATE exams SET teacher_id=? WHERE id=?",(session["user_id"],eid))
                for q in questions:
                    correct=q.get('correct_option') if q['section']==1 else None
                    short_answer=q.get('short_answer') if q['section']==3 else None
                    tf_answers=q.get('tf_answers',[]) if q['section']==2 else []
                    stem=jdump(q['stem'])
                    iq=c.execute("INSERT INTO questions(exam_id,original_order,section_order,section,qtype,blocks_json,stem_json,solution_json,image_url,correct_option,short_answer,tf_answers_json) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",(eid,q['original_order'],q['section_order'],q['section'],q['qtype'],stem,stem,jdump(q.get('solution',[])),first_image(q['stem']),correct,short_answer,jdump(tf_answers)))
                    qid=iq.lastrowid
                    for op in q['options']:
                        oj=jdump(op['content']); tf_val=None
                        if q['section']==2:
                            vals=q.get('tf_answers',[])
                            idx=ord(op['key'].upper())-ord('A') if op['key'] else -1
                            if 0<=idx<len(vals): tf_val=1 if vals[idx] else 0
                        c.execute("INSERT INTO options(question_id,option_key,blocks_json,content_json,image_url,correct_tf) VALUES(?,?,?,?,?,?)",(qid,op['key'],oj,oj,first_image(op['content']),tf_val))
                flash(f"Đã phân tích {len(questions)} câu. Thầy chỉ cần bấm Xuất bản khi sẵn sàng.")
            return redirect(url_for("edit_exam",exam_id=eid))
        except Exception as e: flash("Lỗi xử lý file: "+str(e))
    return render_template("upload_exam.html")


# ============================================================
# V5.21.1 — KIỂM TRA BẢN NHÁP TRƯỚC KHI XUẤT BẢN
# ------------------------------------------------------------
# Hàm này chỉ kiểm tra, KHÔNG xuất bản.
# Dùng chung cho:
#   - Bước 3: Kiểm tra bản nháp
#   - Bước 4: Xuất bản
# ============================================================
def validate_exam_before_publish(c, exam_id):
    exam = c.execute(
        "SELECT * FROM exams WHERE id=?",
        (exam_id,)
    ).fetchone()

    if not exam:
        return None, ["Không tìm thấy đề."]

    n = c.execute(
        "SELECT COUNT(*) FROM questions WHERE exam_id=?",
        (exam_id,)
    ).fetchone()[0]

    if not n:
        return exam, ["Đề không có câu hỏi."]

    errors = []

    n2 = c.execute(
        "SELECT COUNT(*) FROM questions WHERE exam_id=? AND section=2",
        (exam_id,)
    ).fetchone()[0]

    if abs(float(exam["part2_points"] or 0) - n2) > 0.001:
        errors.append(
            f"Phần II có {n2} câu nên điểm Phần II phải bằng "
            f"{n2} điểm (mỗi câu tối đa 1 điểm)."
        )

    if int(exam["part2_scoring_mode"] or 1) not in (1, 2):
        errors.append("Phương án chấm Phần II không hợp lệ.")

    rows = c.execute("""
        SELECT id,original_order,section_order,section,qtype,
               correct_option,short_answer,tf_answers_json
        FROM questions
        WHERE exam_id=?
        ORDER BY section,section_order,original_order
    """,(exam_id,)).fetchall()

    missing = []

    for q in rows:
        opt_count = c.execute(
            "SELECT COUNT(*) FROM options WHERE question_id=?",
            (q["id"],)
        ).fetchone()[0]

        sec_label = ["","I","II","III","IV"][q["section"]]

        if q["section"] == 1:
            if opt_count != 4 or not q["correct_option"]:
                missing.append(
                    f"Phần {sec_label} - Câu {q['section_order']}"
                )

        elif q["section"] == 2:
            try:
                tf = json.loads(q["tf_answers_json"] or "[]")
            except Exception:
                tf = []

            if opt_count != 4 or len(tf) < 4:
                missing.append(
                    f"Phần {sec_label} - Câu {q['section_order']}"
                )

        elif q["section"] == 3:
            if not re.fullmatch(
                r"[-0-9,]+",
                (q["short_answer"] or "").strip()
            ):
                missing.append(
                    f"Phần {sec_label} - Câu {q['section_order']}"
                )

    if missing:
        errors.append(
            "Chưa nhận được đáp án đúng trong file ở: "
            + ", ".join(missing)
            + ". Hãy kiểm tra lại đáp án của các câu này."
        )

    return exam, errors


def ensure_draft_checked_column(c):
    """
    V5.21.4 — đảm bảo DB cũ cũng có cột draft_checked ngay lúc dùng.
    Không phụ thuộc việc init_db() đã migrate thành công hay chưa.
    """
    cols = {
        row[1]
        for row in c.execute("PRAGMA table_info(exams)").fetchall()
    }
    if "draft_checked" not in cols:
        c.execute(
            "ALTER TABLE exams ADD COLUMN draft_checked INTEGER DEFAULT 0"
        )


def is_draft_checked_db(exam):
    """
    V5.21.2 — trạng thái kiểm tra được lưu trong DB.
    Có fallback an toàn nếu DB cũ chưa migrate xong.
    """
    try:
        return bool(int(exam["draft_checked"] or 0))
    except Exception:
        return False


def mark_draft_checked(exam_id):
    with db() as c:
        ensure_draft_checked_column(c)
        c.execute(
            "UPDATE exams SET draft_checked=1 WHERE id=?",
            (exam_id,)
        )


def invalidate_draft_check(exam_id):
    with db() as c:
        ensure_draft_checked_column(c)
        c.execute(
            "UPDATE exams SET draft_checked=0 WHERE id=?",
            (exam_id,)
        )


_MATH_PREVIEW_KEYS=('fallback_url','fallback_width_px','fallback_height_px')


def _strip_math_preview_fields(blocks):
    """Remove display-image metadata without changing LaTeX or other blocks."""
    changed=False
    for block in blocks:
        if block.get('type')!='latex':
            continue
        for key in _MATH_PREVIEW_KEYS:
            if key in block:
                block.pop(key,None)
                changed=True
    return changed


def _repair_pdf_vector_math_blocks(blocks):
    """Repair legacy PDF vector TeX in place without touching other content."""
    repaired=[]
    changed=False
    for original in blocks or []:
        block=dict(original)
        raw_value=str(block.get("value") or "")
        if (
            block.get("type")=="latex"
            and _contains_vector_tex(raw_value)
        ):
            safe=_browser_safe_math_tex(block.get("value"))
            if safe and safe!=block.get("value"):
                block["value"]=safe
                changed=True

        # Older PDF data can hold ``(u,v)=`` and ``120°`` in two adjacent
        # inline blocks. Merge only this precise vector-expression boundary.
        if (
            block.get("type")=="latex" and repaired
            and repaired[-1].get("type")=="latex"
            and _contains_vector_tex(repaired[-1].get("value"))
            and str(repaired[-1].get("value") or "").rstrip().endswith("=")
            and re.match(r"^\s*[+-]?\d", str(block.get("value") or ""))
        ):
            joined=(str(repaired[-1].get("value") or "")
                    + str(block.get("value") or ""))
            safe=_browser_safe_math_tex(joined)
            if safe:
                repaired[-1]["value"]=safe
                changed=True
                continue
        repaired.append(block)

    if repaired != list(blocks or []):
        blocks[:]=repaired
        changed=True
    return changed


def ensure_exam_latex_only(exam_id):
    """Migrate old exams to the single safe LaTeX -> MathJax display path."""
    updates=[]
    with db() as c:
        question_rows=c.execute(
            'SELECT id,blocks_json FROM questions WHERE exam_id=?',(exam_id,)
        ).fetchall()
        option_rows=c.execute(
            'SELECT o.id,o.blocks_json FROM options o '
            'JOIN questions q ON q.id=o.question_id WHERE q.exam_id=?',(exam_id,)
        ).fetchall()
    for table,rows in (('questions',question_rows),('options',option_rows)):
        for row in rows:
            blocks=json.loads(row['blocks_json'] or '[]')
            changed=_strip_math_preview_fields(blocks)
            changed=_repair_pdf_vector_math_blocks(blocks) or changed
            if changed:
                updates.append((table,jdump(blocks),row['id']))
    if updates:
        with db() as c:
            for table,payload,row_id in updates:
                c.execute(f'UPDATE {table} SET blocks_json=? WHERE id=?',(payload,row_id))
    return True


# ============================================================
# V5.22.14 — TỰ KHÔI PHỤC HÌNH CÂU HỎI BỊ MẤT
# ============================================================
def ensure_source_media_columns(c):
    """Add source-file pointers without changing any existing exam data."""
    exam_cols = {
        row[1] for row in c.execute("PRAGMA table_info(exams)").fetchall()
    }
    if "source_stored_name" not in exam_cols:
        c.execute("ALTER TABLE exams ADD COLUMN source_stored_name TEXT DEFAULT ''")


def _find_uploaded_source(stored_name=None, original_name=None):
    """Find the retained original upload, including files saved by older builds."""
    stored = Path(str(stored_name or "")).name
    if stored:
        candidate = UPLOAD_DIR / stored
        if candidate.is_file():
            return candidate

    safe_original = secure_filename(str(original_name or ""))
    if not safe_original:
        return None
    matches = [
        path for path in UPLOAD_DIR.iterdir()
        if path.is_file() and path.name.endswith(safe_original)
    ]
    return max(matches, key=lambda path: path.stat().st_mtime) if matches else None


def _source_questions(path):
    """Parse a retained source through the same verified upload pipeline."""
    if not path or not path.is_file():
        return []
    if path.suffix.lower() == ".docx":
        questions = build_questions(extract_word_blocks(path))
        return _sync_docx_question_figures_from_ooxml(questions, path)
    if path.suffix.lower() == ".pdf":
        lines = extract_pdf_blocks(path)
        questions = build_questions(lines)
        if not questions:
            questions = build_questions(extract_pdf_blocks_fallback(path))
        return questions
    return []


def _missing_figure(block):
    return (
        block.get("type") == "image"
        and block.get("role") == "figure"
        and _question_media_path(block.get("url")) is None
    )


def _blocks_have_missing_figures(blocks):
    return any(_missing_figure(block) for block in blocks or [])


def _restore_figure_urls(existing, fresh):
    """Replace only broken real-figure blocks; text and LaTeX stay untouched."""
    current = [dict(block) for block in (existing or [])]
    positions = [
        index for index, block in enumerate(current)
        if block.get("type") == "image" and block.get("role") == "figure"
    ]
    replacements = [
        dict(block) for block in (fresh or [])
        if block.get("type") == "image"
        and block.get("role") == "figure"
        and _question_media_path(block.get("url")) is not None
    ]
    if len(positions) != len(replacements):
        return current, False

    changed = False
    for index, replacement in zip(positions, replacements):
        if _missing_figure(current[index]):
            current[index] = replacement
            changed = True
    return current, changed


def _question_content_signature(blocks):
    """Match a selected/cloned question without depending on its new number."""
    parts = []
    for block in blocks or []:
        if block.get("type") in ("text", "latex"):
            parts.append(str(block.get("value") or ""))
    return re.sub(r"\s+", "", "".join(parts)).lower()


def _match_fresh_question(row, blocks, fresh_questions, fresh_by_key):
    """Prefer content identity; numbering is only a safe fallback."""
    section = int(row["section"] or 1)
    signature = _question_content_signature(blocks)
    if signature:
        matches = [
            question for question in fresh_questions
            if int(question.get("section") or 1) == section
            and _question_content_signature(question.get("stem", [])) == signature
        ]
        if len(matches) == 1:
            return matches[0]
    return fresh_by_key.get((section, int(row["section_order"] or 1)))


def _exam_source_row(c, exam):
    """Resolve direct uploads and exams cloned from a question bank."""
    source_file = str(exam["source_file"] or "")
    bank_match = re.fullmatch(r"(?:question_bank|practice_bank):(\d+)", source_file)
    if bank_match:
        ensure_question_bank_schema(c)
        source = c.execute(
            "SELECT source_file,stored_name FROM bank_sources WHERE id=?",
            (int(bank_match.group(1)),)
        ).fetchone()
        if source:
            return source["stored_name"], source["source_file"]
        return "", ""
    stored = (
        exam["source_stored_name"]
        if "source_stored_name" in exam.keys() else ""
    )
    return stored, source_file


def ensure_exam_question_media(exam_id):
    """Repair broken figure URLs in a normal or bank-created exam on demand."""
    try:
        with db() as c:
            ensure_source_media_columns(c)
            exam = c.execute("SELECT * FROM exams WHERE id=?", (exam_id,)).fetchone()
            if not exam:
                return False
            question_rows = c.execute(
                "SELECT id,section,section_order,blocks_json FROM questions "
                "WHERE exam_id=? ORDER BY section,section_order,id",
                (exam_id,)
            ).fetchall()
            option_rows = c.execute(
                "SELECT o.id,o.question_id,o.option_key,o.blocks_json "
                "FROM options o JOIN questions q ON q.id=o.question_id "
                "WHERE q.exam_id=? ORDER BY o.question_id,o.option_key",
                (exam_id,)
            ).fetchall()

            def load_blocks(row):
                try:
                    return json.loads(row["blocks_json"] or "[]")
                except Exception:
                    return []

            if not any(
                _blocks_have_missing_figures(load_blocks(row))
                for row in [*question_rows, *option_rows]
            ):
                return False
            stored_name, original_name = _exam_source_row(c, exam)

        source_path = _find_uploaded_source(stored_name, original_name)
        fresh_questions = _source_questions(source_path)
        if not fresh_questions:
            return False
        fresh_by_key = {
            (int(q.get("section") or 1), int(q.get("section_order") or 1)): q
            for q in fresh_questions
        }

        changed = False
        with db() as c:
            for row in question_rows:
                blocks = load_blocks(row)
                fresh = _match_fresh_question(
                    row, blocks, fresh_questions, fresh_by_key
                )
                if not fresh:
                    continue
                repaired, row_changed = _restore_figure_urls(blocks, fresh.get("stem", []))
                if row_changed:
                    payload = jdump(repaired)
                    c.execute(
                        "UPDATE questions SET blocks_json=?,stem_json=?,image_url=? WHERE id=?",
                        (payload, payload, first_image(repaired), row["id"])
                    )
                    changed = True

            fresh_options = {}
            for row in question_rows:
                fresh = _match_fresh_question(
                    row, load_blocks(row), fresh_questions, fresh_by_key
                )
                if fresh:
                    fresh_options[row["id"]] = {
                        str(op.get("key") or "").upper(): op.get("content", [])
                        for op in fresh.get("options", [])
                    }
            for row in option_rows:
                fresh = fresh_options.get(row["question_id"], {}).get(
                    str(row["option_key"] or "").upper()
                )
                if fresh is None:
                    continue
                blocks = load_blocks(row)
                repaired, row_changed = _restore_figure_urls(blocks, fresh)
                if row_changed:
                    payload = jdump(repaired)
                    c.execute(
                        "UPDATE options SET blocks_json=?,content_json=?,image_url=? WHERE id=?",
                        (payload, payload, first_image(repaired), row["id"])
                    )
                    changed = True
        return changed
    except Exception:
        app.logger.exception("Không thể tự khôi phục hình câu hỏi của đề %s", exam_id)
        return False


def ensure_bank_question_media(source_id):
    """Repair real figures in the question-bank preview from its retained file."""
    try:
        with db() as c:
            ensure_question_bank_schema(c)
            source = c.execute(
                "SELECT source_file,stored_name FROM bank_sources WHERE id=?",
                (source_id,)
            ).fetchone()
            if not source:
                return False
            question_rows = c.execute(
                "SELECT id,section,section_order,blocks_json FROM bank_questions "
                "WHERE source_id=? ORDER BY section,section_order,id",
                (source_id,)
            ).fetchall()
            option_rows = c.execute(
                "SELECT o.id,o.bank_question_id,o.option_key,o.blocks_json "
                "FROM bank_options o JOIN bank_questions q "
                "ON q.id=o.bank_question_id WHERE q.source_id=? "
                "ORDER BY o.bank_question_id,o.option_key",
                (source_id,)
            ).fetchall()

        def load_blocks(row):
            try:
                return json.loads(row["blocks_json"] or "[]")
            except Exception:
                return []

        if not any(
            _blocks_have_missing_figures(load_blocks(row))
            for row in [*question_rows, *option_rows]
        ):
            return False

        source_path = _find_uploaded_source(
            source["stored_name"], source["source_file"]
        )
        fresh_questions = _source_questions(source_path)
        fresh_by_key = {
            (int(q.get("section") or 1), int(q.get("section_order") or 1)): q
            for q in fresh_questions
        }
        if not fresh_by_key:
            return False

        changed = False
        with db() as c:
            fresh_options = {}
            for row in question_rows:
                blocks = load_blocks(row)
                fresh = _match_fresh_question(
                    row, blocks, fresh_questions, fresh_by_key
                )
                if not fresh:
                    continue
                repaired, row_changed = _restore_figure_urls(
                    blocks, fresh.get("stem", [])
                )
                if row_changed:
                    c.execute(
                        "UPDATE bank_questions SET blocks_json=? WHERE id=?",
                        (jdump(repaired), row["id"])
                    )
                    changed = True
                fresh_options[row["id"]] = {
                    str(op.get("key") or "").upper(): op.get("content", [])
                    for op in fresh.get("options", [])
                }

            for row in option_rows:
                fresh = fresh_options.get(row["bank_question_id"], {}).get(
                    str(row["option_key"] or "").upper()
                )
                if fresh is None:
                    continue
                repaired, row_changed = _restore_figure_urls(load_blocks(row), fresh)
                if row_changed:
                    c.execute(
                        "UPDATE bank_options SET blocks_json=? WHERE id=?",
                        (jdump(repaired), row["id"])
                    )
                    changed = True
        return changed
    except Exception:
        app.logger.exception(
            "Không thể tự khôi phục hình của ngân hàng đề %s", source_id
        )
        return False


def ensure_shuffle_question_media(source_id):
    """Repair real figures in shuffled-exam previews from the original upload."""
    try:
        with db() as c:
            ensure_shuffle_schema(c)
            source = c.execute(
                "SELECT source_file,stored_name FROM shuffle_sources WHERE id=?",
                (source_id,)
            ).fetchone()
            if not source:
                return False
            question_rows = c.execute(
                "SELECT id,section,section_order,blocks_json FROM shuffle_questions "
                "WHERE source_id=? ORDER BY section,section_order,id",
                (source_id,)
            ).fetchall()
            option_rows = c.execute(
                "SELECT o.id,o.shuffle_question_id,o.option_key,o.blocks_json "
                "FROM shuffle_options o JOIN shuffle_questions q "
                "ON q.id=o.shuffle_question_id WHERE q.source_id=? "
                "ORDER BY o.shuffle_question_id,o.option_key",
                (source_id,)
            ).fetchall()

        def load_blocks(row):
            try:
                return json.loads(row["blocks_json"] or "[]")
            except Exception:
                return []

        if not any(
            _blocks_have_missing_figures(load_blocks(row))
            for row in [*question_rows, *option_rows]
        ):
            return False

        source_path = _find_uploaded_source(
            source["stored_name"], source["source_file"]
        )
        fresh_questions = _source_questions(source_path)
        fresh_by_key = {
            (int(q.get("section") or 1), int(q.get("section_order") or 1)): q
            for q in fresh_questions
        }
        if not fresh_by_key:
            return False

        changed = False
        with db() as c:
            fresh_options = {}
            for row in question_rows:
                blocks = load_blocks(row)
                fresh = _match_fresh_question(
                    row, blocks, fresh_questions, fresh_by_key
                )
                if not fresh:
                    continue
                repaired, row_changed = _restore_figure_urls(
                    blocks, fresh.get("stem", [])
                )
                if row_changed:
                    c.execute(
                        "UPDATE shuffle_questions SET blocks_json=? WHERE id=?",
                        (jdump(repaired), row["id"])
                    )
                    changed = True
                fresh_options[row["id"]] = {
                    str(op.get("key") or "").upper(): op.get("content", [])
                    for op in fresh.get("options", [])
                }

            for row in option_rows:
                fresh = fresh_options.get(row["shuffle_question_id"], {}).get(
                    str(row["option_key"] or "").upper()
                )
                if fresh is None:
                    continue
                repaired, row_changed = _restore_figure_urls(load_blocks(row), fresh)
                if row_changed:
                    c.execute(
                        "UPDATE shuffle_options SET blocks_json=? WHERE id=?",
                        (jdump(repaired), row["id"])
                    )
                    changed = True
        return changed
    except Exception:
        app.logger.exception(
            "Không thể tự khôi phục hình của đề đảo %s", source_id
        )
        return False


@app.get("/teacher/exam/<int:exam_id>/edit")
@login_required("teacher")
def edit_exam(exam_id):
    ensure_exam_question_media(exam_id)
    ensure_exam_latex_only(exam_id)
    with db() as c:
        ensure_draft_checked_column(c)
        exam = c.execute(
            "SELECT * FROM exams WHERE id=?",
            (exam_id,)
        ).fetchone()

    if not exam or int(exam["teacher_id"] or -1) != int(session["user_id"]):
        abort(404)

    is_bank_exam = str(exam["source_file"] or "").startswith(
        "question_bank:"
    )
    draft_checked = is_draft_checked_db(exam)

    return render_template(
        "add_question.html",
        exam=exam,
        questions=get_exam_questions(exam_id),
        is_bank_exam=is_bank_exam,
        draft_checked=draft_checked
    )


@app.route(
    "/teacher/exam/<int:exam_id>/check-draft",
    methods=["GET","POST"]
)
@login_required("teacher")
def check_exam_draft(exam_id):
    """
    ============================================================
    BƯỚC 3 — KIỂM TRA BẢN NHÁP (V5.21.4)
    ============================================================
    Đây là bước xác nhận của giáo viên.

    Để nút luôn hoạt động ổn định:
    - route nhận cả GET và POST;
    - kiểm tra DB/cột draft_checked ngay trong route;
    - nếu đề có ít nhất 1 câu -> đánh dấu draft_checked=1;
    - các cảnh báo đáp án/điểm chỉ là cảnh báo, không khóa Bước 4.
    ============================================================
    """
    with db() as c:
        ensure_draft_checked_column(c)

        exam = c.execute(
            "SELECT * FROM exams WHERE id=?",
            (exam_id,)
        ).fetchone()
        if not exam:
            abort(404)

        n = c.execute(
            "SELECT COUNT(*) FROM questions WHERE exam_id=?",
            (exam_id,)
        ).fetchone()[0]

        if n <= 0:
            flash("Bước 3 chưa thực hiện được vì bản nháp chưa có câu hỏi.")
            return redirect(url_for("edit_exam", exam_id=exam_id))

        # Validator chỉ dùng để đưa cảnh báo cho giáo viên.
        _, warnings = validate_exam_before_publish(c, exam_id)

        # Đánh dấu trực tiếp trong cùng transaction.
        c.execute(
            "UPDATE exams SET draft_checked=1 WHERE id=?",
            (exam_id,)
        )

    if warnings:
        flash("Bước 3 đã hoàn tất. Có cảnh báo để bạn xem lại:")
        for msg in warnings:
            flash("Cảnh báo: " + msg)
    else:
        flash("Bước 3 đã hoàn tất. Bây giờ có thể bấm Bước 4 — Xuất bản.")

    return redirect(
        url_for("edit_exam", exam_id=exam_id) + "#bankWorkflow"
    )


@app.post("/teacher/exam/<int:exam_id>/save")
@login_required("teacher")
def save_exam(exam_id):
    # Nếu thay đổi dữ liệu sau khi kiểm tra thì phải kiểm tra lại.
    invalidate_draft_check(exam_id)
    data=request.get_json(silent=True) or {}
    with db() as c:
        for q in data.get("questions",[]):
            key = q.get("correct_option") or None
            c.execute("UPDATE questions SET correct_option=? WHERE id=? AND exam_id=?",(key,q["id"],exam_id))
        n = c.execute("SELECT COUNT(*) FROM questions WHERE exam_id=?",(exam_id,)).fetchone()[0]
        bad = c.execute("SELECT COUNT(*) FROM questions WHERE exam_id=? AND (correct_option IS NULL OR correct_option='')",(exam_id,)).fetchone()[0]
        optbad = c.execute("SELECT COUNT(*) FROM questions q WHERE q.exam_id=? AND (SELECT COUNT(*) FROM options o WHERE o.question_id=q.id)<2",(exam_id,)).fetchone()[0]
        # KHỐI XUẤT BẢN — lưu đáp án không tự xuất bản; giáo viên phải bấm nút Xuất bản.
        published=False
    return jsonify(ok=True,published=published)

@app.post("/teacher/exam/<int:exam_id>/scores")
@login_required("teacher")
def save_score_settings(exam_id):
    # V5.25.3: check ownership BEFORE changing the draft/check state.
    with db() as c:
        exam=c.execute("SELECT * FROM exams WHERE id=?",(exam_id,)).fetchone()
        if not exam or int(exam['teacher_id'] or -1) != int(session['user_id']):
            abort(404)
        if exam['published']:
            flash('Không thể đổi phương án chấm sau khi đề đã xuất bản. Hãy tạo bản nháp mới.')
            return redirect(url_for('edit_exam',exam_id=exam_id))
        try: mode=int(request.form.get('p2_mode','1'))
        except ValueError: mode=1
        if mode not in (1,2): mode=1
        n2=c.execute('SELECT COUNT(*) FROM questions WHERE exam_id=? AND section=2',(exam_id,)).fetchone()[0]
        if abs(float(exam['part2_points'] or 0)-n2)>0.001:
            flash(f'Phần II có {n2} câu nên điểm Phần II phải bằng {n2} điểm (mỗi câu tối đa 1 điểm).')
            return redirect(url_for('edit_exam',exam_id=exam_id))
        # Only an ACTUAL scoring change invalidates the prior draft check;
        # saving the already-checked mode and clicking Xuất bản must work.
        ensure_draft_checked_column(c)
        c.execute("""UPDATE exams SET part2_scoring_mode=?,
                  draft_checked=CASE WHEN COALESCE(part2_scoring_mode,1)=?
                  THEN COALESCE(draft_checked,0) ELSE 0 END
                  WHERE id=? AND teacher_id=?""",
                  (mode,mode,exam_id,session['user_id']))
    if request.form.get('score_action') == 'publish':
        # Save the selected scoring mode first, then invoke the same
        # publisher used by the teacher dashboard (including its validation).
        return publish_exam(exam_id)
    flash('Đã lưu phương án chấm từng câu Phần II.')
    return redirect(url_for('edit_exam',exam_id=exam_id))

@app.route(
    "/teacher/exam/<int:exam_id>/publish",
    methods=["GET","POST"]
)
@login_required("teacher")
def publish_exam(exam_id):
    """
    ============================================================
    BƯỚC 4 — XUẤT BẢN CHO HỌC SINH (V5.21.4)
    ============================================================
    Route nhận cả GET và POST để nút/đường dẫn đều hoạt động.

    Với đề Ngân hàng:
    - yêu cầu draft_checked=1;
    - nếu đạt -> UPDATE published=1;
    - quay lại cùng màn hình để giáo viên thấy trạng thái Đã xuất bản.
    ============================================================
    """
    with db() as c:
        ensure_draft_checked_column(c)

        exam = c.execute(
            "SELECT * FROM exams WHERE id=?",
            (exam_id,)
        ).fetchone()
        if not exam or int(exam["teacher_id"] or -1) != int(session["user_id"]):
            abort(404)

        is_bank_exam = str(exam["source_file"] or "").startswith(
            "question_bank:"
        )

        if is_bank_exam and not bool(int(exam["draft_checked"] or 0)):
            flash(
                "Chưa thể xuất bản: hãy bấm Bước 3 — Kiểm tra trước."
            )
            return redirect(
                url_for("edit_exam", exam_id=exam_id) + "#bankWorkflow"
            )

        n = c.execute(
            "SELECT COUNT(*) FROM questions WHERE exam_id=?",
            (exam_id,)
        ).fetchone()[0]

        if n <= 0:
            flash("Không thể xuất bản vì đề chưa có câu hỏi.")
            return redirect(url_for("edit_exam", exam_id=exam_id))

        c.execute(
            "UPDATE exams SET published=1 WHERE id=?",
            (exam_id,)
        )

    flash("Bước 4 hoàn tất: Đề đã xuất bản cho học sinh.")
    return redirect(
        url_for("edit_exam", exam_id=exam_id) + "#bankWorkflow"
    )


@app.post("/teacher/exam/<int:exam_id>/unpublish")
@login_required("teacher")
def unpublish_exam(exam_id):
    with db() as c: c.execute("UPDATE exams SET published=0 WHERE id=?",(exam_id,))
    return redirect(url_for("edit_exam",exam_id=exam_id))

def _teacher_roster_path(teacher_id):
    p=INSTANCE_DIR/'rosters'/f'teacher_{int(teacher_id)}.xlsx'
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def import_roster(path, teacher_id):
    from openpyxl import load_workbook
    p=Path(path); xlsx=p
    if p.suffix.lower()=='.xls':
        soffice=find_soffice()
        if not soffice: raise RuntimeError('Cần LibreOffice để đọc XLS.')
        td=Path(tempfile.mkdtemp()); subprocess.run([soffice,'--headless','--convert-to','xlsx','--outdir',str(td),str(p)],check=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,timeout=60); xlsx=td/(p.stem+'.xlsx')
    wb=load_workbook(xlsx); ws=wb[wb.sheetnames[0]]; header=None
    for row in ws.iter_rows():
        vals=[str(c.value or '').strip() for c in row]
        if any('Mã học sinh' in v for v in vals) and any('Họ và tên' in v for v in vals): header=row[0].row; break
    if not header: raise RuntimeError('Không tìm thấy tiêu đề Mã học sinh/Họ và tên.')
    headers={str(c.value).strip():c.column for c in ws[header] if c.value is not None}; code_col=headers['Mã học sinh']; name_col=headers['Họ và tên']
    students=[]
    credentials=[]
    enrolled_count=0
    for r in range(header+2,ws.max_row+1):
        code=ws.cell(r,code_col).value; a=ws.cell(r,name_col).value; b=ws.cell(r,name_col+1).value if name_col<ws.max_column else None; full=normalize_text(f'{a or ""} {b or ""}')
        if code and full: students.append((str(code).strip(),full,r))
    with db() as c:
        for code,full,row in students:
            # ========================================================
            # If student already exists from another class, they must
            # accept teacher invite before this teacher can update roster.
            # ========================================================
            old=c.execute(
                "SELECT id,username FROM users WHERE student_code=? AND role='student'",
                (code,)
            ).fetchone()

            if old:
                # Existing pupil account: another teacher must not silently
                # attach it via Excel/student code. Pupil accepts invite.
                linked=c.execute("SELECT 1 FROM teacher_students WHERE teacher_id=? AND student_id=?",(teacher_id,old['id'])).fetchone()
                if not linked:
                    continue
                owner=c.execute(
                    "SELECT id FROM users WHERE username=? AND id<>?",
                    (code,old["id"])
                ).fetchone()
                username = old["username"] if owner else code
                c.execute(
                    "UPDATE users SET full_name=COALESCE(NULLIF(full_name,''),?),excel_row=? WHERE id=?",
                    (full,row,old["id"])
                )
                user_id = old["id"]
            else:
                owner=c.execute(
                    "SELECT id FROM users WHERE username=?",
                    (code,)
                ).fetchone()
                username = code if not owner else 'hs_'+re.sub(r'\W+','_',code)
                import secrets
                temporary_password=secrets.token_urlsafe(16)
                cur=c.execute(
                    "INSERT INTO users(username,password_hash,role,full_name,student_code,excel_row) VALUES(?,?,?,?,?,?)",
                    (username,generate_password_hash(temporary_password),'student',full,code,row)
                )
                credentials.append((username,code,temporary_password))
                user_id = cur.lastrowid

            c.execute("INSERT OR IGNORE INTO teacher_students(teacher_id,student_id) VALUES(?,?)",(teacher_id,user_id))
            enrolled_count+=1
            # Nếu trước đây có request pending trùng mã HS, Excel được ưu tiên
            # và request được ghi nhận là đã cấp tài khoản tự động.
            c.execute(
                "UPDATE registration_requests SET status='approved',decided_at=CURRENT_TIMESTAMP,approved_user_id=? WHERE requested_role='student' AND student_code=? AND status='pending'",
                (user_id,code)
            )

    shutil.copy2(xlsx,_teacher_roster_path(teacher_id))
    # Kept outside public/static: only the importing teacher may fetch and
    # delete this temporary credentials sheet. Existing passwords untouched.
    creds_path=_teacher_roster_path(teacher_id).with_name(f'teacher_{int(teacher_id)}_initial_accounts.csv')
    if credentials:
        import csv
        with creds_path.open('w',newline='',encoding='utf-8-sig') as handle:
            writer=csv.writer(handle)
            writer.writerow(['Tên tài khoản','Mã học sinh','Mật khẩu tạm'])
            writer.writerows(credentials)
        try: creds_path.chmod(0o600)
        except OSError: pass
    return enrolled_count

def _exam_teacher_id(exam_id):
    with db() as c:
        row=c.execute('SELECT teacher_id FROM exams WHERE id=?',(exam_id,)).fetchone()
        return row['teacher_id'] if row else -1


def update_roster_scores(exam_id):
    """V5.11.5 — cập nhật Excel lớp theo Mã học sinh, gồm điểm và gian lận."""
    from openpyxl import load_workbook
    p=_teacher_roster_path(_exam_teacher_id(exam_id))
    if not p.exists(): return
    wb=load_workbook(p); ws=wb[wb.sheetnames[0]]; header=None
    for row in ws.iter_rows():
        vals=[str(c.value or '').strip() for c in row]
        if any('Mã học sinh' in v for v in vals): header=row[0].row; break
    if not header: return
    headers={str(c.value).strip():c.column for c in ws[header] if c.value is not None}; code_col=headers.get('Mã học sinh')
    if not code_col: return
    for name in ('TN','DS','TLN','TL','Tổng','Lỗi gian lận','Chi tiết gian lận'):
        if name not in headers: headers[name]=ws.max_column+1; ws.cell(header,headers[name]).value=name
    with db() as c:
        rows=c.execute("SELECT se.id,se.score,se.violations,u.student_code FROM student_exams se LEFT JOIN users u ON u.id=se.student_id WHERE se.exam_id=? AND se.submitted_at IS NOT NULL",(exam_id,)).fetchall()
        for r in rows:
            parts={}
            for sec in (1,2,3):
                v=c.execute("SELECT COALESCE(SUM(a.score),0) FROM student_answers a JOIN questions q ON q.id=a.question_id WHERE a.student_exam_id=? AND q.section=?",(r['id'],sec)).fetchone()[0]
                parts[sec]=round(float(v or 0),2)
            total=round(float(r['score'] or 0),2); parts[4]=round(max(0,total-sum(parts.values())),2)
            logs=c.execute("SELECT event_type,details,created_at FROM anti_cheat_logs WHERE student_exam_id=? ORDER BY id",(r['id'],)).fetchall()
            detail='; '.join(f"{x['created_at']} — {x['event_type']}: {x['details'] or ''}" for x in logs) or 'Không có'
            for rr in range(header+1,ws.max_row+1):
                if str(ws.cell(rr,code_col).value or '').strip()==str(r['student_code'] or '').strip():
                    for name,val in [('TN',parts[1]),('DS',parts[2]),('TLN',parts[3]),('TL',parts[4]),('Tổng',total),('Lỗi gian lận',int(r['violations'] or 0)),('Chi tiết gian lận',detail)]: ws.cell(rr,headers[name]).value=val
                    break
    ws.freeze_panes=f'A{header+1}'; wb.save(p)

@app.post("/teacher/roster")
@login_required("teacher")
def roster_upload():
    f=request.files.get('roster')
    if not f or ext(f.filename) not in {'xls','xlsx'}: flash('Hãy tải XLS/XLSX.'); return redirect(url_for('teacher_dashboard'))
    p=UPLOAD_DIR/(uuid.uuid4().hex+'_'+secure_filename(f.filename)); f.save(p)
    try: flash(f'Đã đồng bộ {import_roster(p,session["user_id"])} tài khoản vào lớp. Học sinh đã có tài khoản nhưng chưa tham gia lớp cần dùng link mời. Nếu có tài khoản mới, hãy tải file mật khẩu tạm ở bảng điều khiển.')
    except Exception as e: flash('Lỗi danh sách Excel: '+str(e))
    return redirect(url_for('teacher_dashboard'))

@app.get('/teacher/roster/initial-accounts.csv')
@login_required('teacher')
def roster_initial_accounts():
    from flask import send_file
    path=_teacher_roster_path(session['user_id']).with_name(f"teacher_{int(session['user_id'])}_initial_accounts.csv")
    if not path.is_file(): abort(404)
    return send_file(path,as_attachment=True,download_name='Tai_khoan_hoc_sinh_moi.csv',mimetype='text/csv')


@app.post('/teacher/roster/initial-accounts/delete')
@login_required('teacher')
def roster_initial_accounts_delete():
    path=_teacher_roster_path(session['user_id']).with_name(f"teacher_{int(session['user_id'])}_initial_accounts.csv")
    try: path.unlink(missing_ok=True)
    except OSError: flash('Không xóa được file. Vui lòng kiểm tra quyền Windows.')
    return redirect(url_for('teacher_dashboard'))


@app.get("/teacher/roster/download")
@login_required("teacher")
def roster_download():
    p=_teacher_roster_path(session['user_id'])
    if not p.exists(): abort(404)
    return send_file(p,as_attachment=True,download_name='Danh_sach_lop_cap_nhat.xlsx')


# ============================================================
# V5.14.0 — TÀI LIỆU GIÁO VIÊN UPLOAD
# Giáo viên upload, học sinh chỉ xem/tải.
# ============================================================
LEGACY_DOCUMENT_DIR = BASE / "static" / "documents"
DOCUMENT_DIR = DATA_ROOT / "documents"
DOCUMENT_DIR.mkdir(parents=True, exist_ok=True)
LEGACY_DOCUMENT_DIR.mkdir(parents=True, exist_ok=True)
DOCUMENT_ALLOWED = {"pdf","doc","docx","ppt","pptx","xls","xlsx","jpg","jpeg","png","txt"}


def ensure_document_folder_schema(c):
    """Create the additive folder schema without changing legacy documents."""
    c.execute("""
        CREATE TABLE IF NOT EXISTS document_folders(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL COLLATE NOCASE,
            created_by INTEGER NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(created_by,name)
        )
    """)
    columns = {
        row[1] for row in c.execute("PRAGMA table_info(documents)").fetchall()
    }
    if "folder_id" not in columns:
        c.execute("ALTER TABLE documents ADD COLUMN folder_id INTEGER")
    c.execute(
        "CREATE INDEX IF NOT EXISTS idx_documents_folder_id "
        "ON documents(folder_id)"
    )


def _owned_folder_id(c, table_name, raw_value, owner_id):
    """Return a validated folder id owned by the signed-in teacher."""
    value = str(raw_value or "").strip()
    if not value:
        return None
    if not value.isdigit() or int(value) <= 0:
        raise ValueError("Thư mục đã chọn không hợp lệ.")
    folder_id = int(value)
    row = c.execute(
        f"SELECT id FROM {table_name} WHERE id=? AND created_by=?",
        (folder_id, owner_id),
    ).fetchone()
    if not row:
        raise ValueError("Thư mục không tồn tại hoặc không thuộc tài khoản này.")
    return folder_id


def _rename_owned_folder(c, table_name, folder_id, owner_id, raw_name):
    """Rename one owned folder without changing its id or child records."""
    if table_name not in {"document_folders", "bank_folders"}:
        raise ValueError("Loại thư mục không hợp lệ.")
    name = normalize_text(raw_name)
    if not name:
        raise ValueError("Vui lòng nhập tên thư mục mới.")
    if len(name) > 120:
        raise ValueError("Tên thư mục tối đa 120 ký tự.")
    row = c.execute(
        f"SELECT id FROM {table_name} WHERE id=? AND created_by=?",
        (folder_id, owner_id),
    ).fetchone()
    if not row:
        raise LookupError("Thư mục không tồn tại hoặc không thuộc tài khoản này.")
    c.execute(
        f"UPDATE {table_name} SET name=? WHERE id=? AND created_by=?",
        (name, folder_id, owner_id),
    )
    return name


def _document_path(stored_name):
    """Resolve new persistent documents first and support legacy documents."""
    safe = Path(str(stored_name or "")).name
    if not safe or safe != str(stored_name or ""):
        return None
    for folder in (DOCUMENT_DIR, LEGACY_DOCUMENT_DIR):
        candidate = folder / safe
        if candidate.is_file():
            return candidate
    return DOCUMENT_DIR / safe


@app.post("/teacher/documents/folders")
@login_required("teacher")
def teacher_create_document_folder():
    name = normalize_text(request.form.get("name"))
    if not name:
        flash("Vui lòng nhập tên thư mục tài liệu.")
        return redirect(url_for("teacher_documents"))
    if len(name) > 120:
        flash("Tên thư mục tối đa 120 ký tự.")
        return redirect(url_for("teacher_documents"))
    try:
        with db() as c:
            ensure_document_folder_schema(c)
            c.execute(
                "INSERT INTO document_folders(name,created_by) VALUES(?,?)",
                (name, session["user_id"]),
            )
    except sqlite3.IntegrityError:
        flash("Bạn đã có một thư mục tài liệu cùng tên.")
        return redirect(url_for("teacher_documents"))
    flash("Đã tạo thư mục tài liệu “" + name + "”.")
    return redirect(url_for("teacher_documents"))


@app.post("/teacher/documents/folders/<int:folder_id>/rename")
@login_required("teacher")
def teacher_rename_document_folder(folder_id):
    try:
        with db() as c:
            ensure_document_folder_schema(c)
            name = _rename_owned_folder(
                c, "document_folders", folder_id,
                session["user_id"], request.form.get("name"),
            )
    except LookupError:
        abort(404)
    except ValueError as error:
        flash(str(error))
        return redirect(url_for("teacher_document_folder", folder_id=folder_id))
    except sqlite3.IntegrityError:
        flash("Bạn đã có một thư mục tài liệu cùng tên.")
        return redirect(url_for("teacher_document_folder", folder_id=folder_id))
    flash("Đã đổi tên thư mục tài liệu thành “" + name + "”.")
    return redirect(url_for("teacher_document_folder", folder_id=folder_id))


@app.route("/teacher/documents", methods=["GET","POST"])
@login_required("teacher")
def teacher_documents():
    if request.method == "POST":
        raw_folder_id = request.form.get("folder_id")
        return_url = (
            url_for("teacher_document_folder", folder_id=int(raw_folder_id))
            if str(raw_folder_id or "").isdigit() and int(raw_folder_id) > 0
            else url_for("teacher_documents")
        )
        f = request.files.get("document")
        title = normalize_text(request.form.get("title")) or (
            secure_filename(f.filename) if f and f.filename else "Tài liệu"
        )
        description = normalize_text(request.form.get("description"))

        if not f or not f.filename:
            flash("Vui lòng chọn file tài liệu.")
            return redirect(return_url)

        extension = ext(f.filename)
        if extension not in DOCUMENT_ALLOWED:
            flash("Định dạng tài liệu chưa được hỗ trợ.")
            return redirect(return_url)

        stored = uuid.uuid4().hex + "_" + secure_filename(f.filename)
        path = DOCUMENT_DIR / stored
        try:
            with db() as c:
                ensure_document_folder_schema(c)
                folder_id = _owned_folder_id(
                    c, "document_folders", request.form.get("folder_id"),
                    session["user_id"],
                )
                f.save(path)
                c.execute("""
                    INSERT INTO documents(
                        title,description,stored_name,original_name,
                        uploaded_by,folder_id
                    ) VALUES(?,?,?,?,?,?)
                """,(
                    title,description,stored,f.filename,
                    session["user_id"],folder_id
                ))
        except ValueError as error:
            if path.exists():
                path.unlink()
            flash(str(error))
            return redirect(return_url)
        flash("Đã tải tài liệu lên.")
        return redirect(return_url)

    with db() as c:
        ensure_document_folder_schema(c)
        folders = c.execute("""
            SELECT f.*,
                   (SELECT COUNT(*) FROM documents d
                    WHERE d.folder_id=f.id AND d.uploaded_by=?) AS item_count
            FROM document_folders f
            WHERE f.created_by=?
            ORDER BY f.name COLLATE NOCASE,f.id
        """,(session["user_id"],session["user_id"])).fetchall()
        unfiled_count = c.execute("""
            SELECT COUNT(*) FROM documents
            WHERE uploaded_by=? AND folder_id IS NULL
        """,(session["user_id"],)).fetchone()[0]
    return render_template(
        "teacher_documents.html",
        folders=folders,
        unfiled_count=unfiled_count,
    )


@app.get("/teacher/documents/folder/<int:folder_id>")
@login_required("teacher")
def teacher_document_folder(folder_id):
    with db() as c:
        ensure_document_folder_schema(c)
        folder = c.execute("""
            SELECT * FROM document_folders
            WHERE id=? AND created_by=?
        """,(folder_id,session["user_id"])).fetchone()
        if not folder:
            abort(404)
        documents = c.execute("""
            SELECT * FROM documents
            WHERE folder_id=? AND uploaded_by=?
            ORDER BY id DESC
        """,(folder_id,session["user_id"])).fetchall()
    return render_template(
        "teacher_document_folder.html",
        folder=folder,
        documents=documents,
        is_unfiled=False,
    )


@app.get("/teacher/documents/unfiled")
@login_required("teacher")
def teacher_unfiled_documents():
    with db() as c:
        ensure_document_folder_schema(c)
        documents = c.execute("""
            SELECT * FROM documents
            WHERE folder_id IS NULL AND uploaded_by=?
            ORDER BY id DESC
        """,(session["user_id"],)).fetchall()
    return render_template(
        "teacher_document_folder.html",
        folder={"id":None,"name":"Tài liệu chưa phân loại"},
        documents=documents,
        is_unfiled=True,
    )


@app.post("/teacher/document/<int:doc_id>/delete")
@login_required("teacher")
def teacher_delete_document(doc_id):
    with db() as c:
        doc = c.execute(
            "SELECT * FROM documents WHERE id=? AND uploaded_by=?",
            (doc_id,session["user_id"])
        ).fetchone()
        if not doc:
            abort(404)
        c.execute("DELETE FROM documents WHERE id=?", (doc_id,))

    try:
        for folder in (DOCUMENT_DIR, LEGACY_DOCUMENT_DIR):
            p = folder / Path(doc["stored_name"]).name
            if p.exists() and p.is_file():
                p.unlink()
    except Exception:
        pass

    flash("Đã xóa tài liệu.")
    if doc["folder_id"] is not None:
        return redirect(url_for(
            "teacher_document_folder", folder_id=doc["folder_id"]
        ))
    return redirect(url_for("teacher_unfiled_documents"))


@app.get("/student/documents")
@login_required("student")
def student_documents():
    with db() as c:
        ensure_document_folder_schema(c)
        folders = c.execute("""
            SELECT f.*,COALESCE(u.full_name,'Giáo viên') AS teacher_name,
                   (SELECT COUNT(*) FROM documents d
                    WHERE d.folder_id=f.id) AS item_count
            FROM document_folders f
            JOIN teacher_students ts ON ts.teacher_id=f.created_by AND ts.student_id=?
            LEFT JOIN users u ON u.id=f.created_by
            ORDER BY f.name COLLATE NOCASE,f.id
        """,(session["user_id"],)).fetchall()
        unfiled_count = c.execute(
            "SELECT COUNT(*) FROM documents d JOIN teacher_students ts ON ts.teacher_id=d.uploaded_by WHERE d.folder_id IS NULL AND ts.student_id=?",(session["user_id"],)
        ).fetchone()[0]
    return render_template(
        "student_documents.html",
        folders=folders,
        unfiled_count=unfiled_count,
    )


@app.get("/student/documents/folder/<int:folder_id>")
@login_required("student")
def student_document_folder(folder_id):
    with db() as c:
        ensure_document_folder_schema(c)
        folder = c.execute("""
            SELECT f.*,COALESCE(u.full_name,'Giáo viên') AS teacher_name
            FROM document_folders f
            LEFT JOIN users u ON u.id=f.created_by
            WHERE f.id=?
        """,(folder_id,)).fetchone()
        if not folder:
            abort(404)
        documents = c.execute("""
            SELECT * FROM documents
            WHERE folder_id=?
            ORDER BY id DESC
        """,(folder_id,)).fetchall()
    return render_template(
        "student_document_folder.html",
        folder=folder,
        documents=documents,
        is_unfiled=False,
    )


@app.get("/student/documents/unfiled")
@login_required("student")
def student_unfiled_documents():
    with db() as c:
        ensure_document_folder_schema(c)
        documents = c.execute("""
            SELECT d.*,COALESCE(u.full_name,'Giáo viên') AS teacher_name
            FROM documents d
            LEFT JOIN users u ON u.id=d.uploaded_by
            JOIN teacher_students ts ON ts.teacher_id=d.uploaded_by AND ts.student_id=?
            WHERE d.folder_id IS NULL
            ORDER BY d.id DESC
        """,(session["user_id"],)).fetchall()
    return render_template(
        "student_document_folder.html",
        folder={"id":None,"name":"Tài liệu chưa phân loại","teacher_name":""},
        documents=documents,
        is_unfiled=True,
    )



@app.get("/document/<int:doc_id>/view")
def view_document(doc_id):
    """
    V5.16.4 — Trang xem tài liệu dùng chung cho Giáo viên/Học sinh.

    - PDF: iframe trực tiếp.
    - Ảnh: hiển thị ảnh.
    - TXT: hiển thị text.
    - DOC/DOCX/XLS/XLSX/PPT/PPTX:
      thử chuyển sang PDF bằng LibreOffice rồi hiển thị.
    """
    if "user_id" not in session:
        return redirect(url_for("login"))

    with db() as c:
        doc = c.execute(
            "SELECT * FROM documents WHERE id=?",
            (doc_id,)
        ).fetchone()

    if not doc:
        abort(404)

    path = _document_path(doc["stored_name"])
    if not path or not path.exists():
        abort(404)

    extension = ext(doc["original_name"]).lower()
    preview_kind = None
    preview_url = None
    text_content = None
    preview_error = None

    if extension == "pdf":
        preview_kind = "pdf"
        preview_url = url_for("document_inline_file", doc_id=doc_id)

    elif extension in ("jpg","jpeg","png","gif","webp"):
        preview_kind = "image"
        preview_url = url_for("document_inline_file", doc_id=doc_id)

    elif extension in ("txt","csv"):
        preview_kind = "text"
        try:
            text_content = path.read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            preview_error = str(e)

    elif extension in ("doc","docx","xls","xlsx","ppt","pptx"):
        try:
            preview_dir = INSTANCE_DIR / "document_previews"
            preview_dir.mkdir(parents=True, exist_ok=True)

            expected = preview_dir / (path.stem + ".pdf")
            if not expected.exists():
                soffice = find_soffice()
                cmd = [
                    soffice,
                    "--headless",
                    "--convert-to", "pdf",
                    "--outdir", str(preview_dir),
                    str(path)
                ]
                subprocess.run(
                    cmd,
                    check=True,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=90
                )

            if expected.exists():
                preview_kind = "pdf"
                preview_url = url_for("document_preview_pdf", doc_id=doc_id)
            else:
                preview_error = "Không tạo được bản xem trước PDF."
        except Exception as e:
            preview_error = (
                "Không thể tạo bản xem trước tài liệu Office. "
                "Hãy kiểm tra LibreOffice/SOFFICE_PATH. Chi tiết: " + str(e)
            )
    else:
        preview_error = "Định dạng này chưa hỗ trợ xem trực tiếp."

    return render_template(
        "document_view.html",
        doc=doc,
        preview_kind=preview_kind,
        preview_url=preview_url,
        text_content=text_content,
        preview_error=preview_error
    )


@app.get("/document/<int:doc_id>/inline")
def document_inline_file(doc_id):
    if "user_id" not in session:
        abort(403)

    with db() as c:
        doc = c.execute(
            "SELECT * FROM documents WHERE id=?",
            (doc_id,)
        ).fetchone()

    if not doc:
        abort(404)

    path = _document_path(doc["stored_name"])
    if not path or not path.exists():
        abort(404)

    return send_file(
        path,
        as_attachment=False,
        download_name=doc["original_name"]
    )


@app.get("/document/<int:doc_id>/preview.pdf")
def document_preview_pdf(doc_id):
    if "user_id" not in session:
        abort(403)

    with db() as c:
        doc = c.execute(
            "SELECT * FROM documents WHERE id=?",
            (doc_id,)
        ).fetchone()

    if not doc:
        abort(404)

    original = _document_path(doc["stored_name"])
    if not original:
        abort(404)
    preview = INSTANCE_DIR / "document_previews" / (original.stem + ".pdf")

    if not preview.exists():
        abort(404)

    return send_file(
        preview,
        mimetype="application/pdf",
        as_attachment=False,
        download_name=Path(doc["original_name"]).stem + ".pdf"
    )


@app.get("/document/<int:doc_id>/download")
def download_document(doc_id):
    with db() as c:
        doc = c.execute(
            "SELECT * FROM documents WHERE id=?",
            (doc_id,)
        ).fetchone()
    if not doc:
        abort(404)

    path = _document_path(doc["stored_name"])
    if not path or not path.exists():
        abort(404)

    return send_file(
        path,
        as_attachment=True,
        download_name=doc["original_name"]
    )



# ============================================================
# V5.14.0 — TẠO ĐỀ: THỦ CÔNG / AI
# ============================================================
@app.get("/teacher/create-exam")
@login_required("teacher")
def create_exam_hub():
    return render_template("create_exam_hub.html")


@app.route("/teacher/create-exam/manual", methods=["GET","POST"])
@login_required("teacher")
def create_exam_manual():
    """
    V5.15.0 — TẠO ĐỀ THỦ CÔNG BẰNG LATEX

    Không sửa upload Word/PDF cũ.
    Giáo viên gõ câu hỏi theo LaTeX, hệ thống tạo một exam NHÁP
    theo schema cũ, redirect sang màn hình kiểm tra hiện có,
    sau đó giáo viên bấm Xuất bản như V5.12.6.
    """
    if request.method == "GET":
        return render_template("create_exam_manual.html")

    title = normalize_text(request.form.get("title")) or "Đề thủ công"
    duration = int(request.form.get("duration") or 90)

    try:
        raw_questions = json.loads(request.form.get("questions_json") or "[]")
    except Exception:
        flash("Dữ liệu câu hỏi không hợp lệ.")
        return redirect(request.url)

    if not raw_questions:
        flash("Hãy tạo ít nhất một câu hỏi.")
        return redirect(request.url)

    try:
        pts = [
            float(str(request.form.get(k,"0")).replace(",","."))
            for k in ("p1","p2","p3","p4")
        ]
    except ValueError:
        flash("Điểm các phần không hợp lệ.")
        return redirect(request.url)

    if abs(sum(pts)-10) > 0.001:
        flash("Tổng điểm 4 phần phải bằng 10.")
        return redirect(request.url)

    try:
        p2_mode = int(request.form.get("p2_mode") or 1)
    except Exception:
        p2_mode = 1
    if p2_mode not in (1,2):
        p2_mode = 1

    clean = []
    sec_counter = {1:0,2:0,3:0,4:0}

    for item in raw_questions:
        try:
            sec = int(item.get("section"))
        except Exception:
            continue

        if sec not in (1,2,3,4):
            continue

        stem_text = str(item.get("stem") or "").strip()
        if not stem_text:
            continue

        sec_counter[sec] += 1

        if sec == 1:
            qtype = "mcq"
        elif sec == 2:
            qtype = "tf"
        elif sec == 3:
            qtype = "short"
        else:
            qtype = "essay"

        options = []
        for key in ("A","B","C","D"):
            val = str(item.get("options",{}).get(key,"") or "").strip()
            if val:
                options.append({
                    "key": key,
                    "content": manual_latex_blocks(val)
                })

        correct_option = None
        short_answer = None
        tf_answers = []

        answer = str(item.get("answer") or "").strip()

        if sec == 1:
            answer = answer.upper()
            if answer not in ("A","B","C","D"):
                flash(
                    f"Câu {sec_counter[sec]} Phần I cần đáp án A/B/C/D."
                )
                return redirect(request.url)
            if len(options) < 2:
                flash(
                    f"Câu {sec_counter[sec]} Phần I cần ít nhất 2 phương án."
                )
                return redirect(request.url)
            correct_option = answer

        elif sec == 2:
            if len(options) != 4:
                flash(
                    f"Câu {sec_counter[sec]} Phần II cần đủ 4 ý A/B/C/D."
                )
                return redirect(request.url)
            tf_answers = parse_tf_key(answer)
            if len(tf_answers) != 4:
                flash(
                    f"Câu {sec_counter[sec]} Phần II cần đáp án dạng ĐSĐS."
                )
                return redirect(request.url)

        elif sec == 3:
            if not answer:
                flash(
                    f"Câu {sec_counter[sec]} Phần III cần đáp án ngắn."
                )
                return redirect(request.url)
            short_answer = answer

        clean.append({
            "section": sec,
            "section_order": sec_counter[sec],
            "qtype": qtype,
            "stem": manual_latex_blocks(stem_text),
            "options": options,
            "correct_option": correct_option,
            "short_answer": short_answer,
            "tf_answers": tf_answers
        })

    if not clean:
        flash("Không có câu hợp lệ để tạo đề.")
        return redirect(request.url)

    n2 = sum(1 for q in clean if q["section"]==2)
    if abs(pts[1]-n2) > 0.001:
        flash(
            f"Phần II có {n2} câu nên điểm Phần II phải bằng {n2} điểm."
        )
        return redirect(request.url)

    with db() as c:
        cur = c.execute("""
            INSERT INTO exams(
                title,source_file,question_count,published,
                duration_minutes,
                part1_points,part2_points,part3_points,part4_points,
                p2_item_a,p2_item_b,p2_item_c,p2_item_d,
                part2_scoring_mode,anti_cheat,max_violations
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """,(
            title,
            "manual_latex",
            len(clean),
            0,
            duration,
            *pts,
            0.25,0.25,0.25,0.25,
            p2_mode,
            1 if request.form.get("anti_cheat") else 0,
            int(request.form.get("max_violations") or 3)
        ))
        exam_id = cur.lastrowid
        c.execute("UPDATE exams SET teacher_id=? WHERE id=?",(session["user_id"],exam_id))

        global_order = 0
        for q in clean:
            global_order += 1
            stem = jdump(q["stem"])

            iq = c.execute("""
                INSERT INTO questions(
                    exam_id,original_order,section_order,
                    section,qtype,blocks_json,stem_json,
                    image_url,correct_option,short_answer,tf_answers_json
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?)
            """,(
                exam_id,
                global_order,
                q["section_order"],
                q["section"],
                q["qtype"],
                stem,
                stem,
                None,
                q["correct_option"],
                q["short_answer"],
                jdump(q["tf_answers"])
            ))
            qid = iq.lastrowid

            for op in q["options"]:
                oj = jdump(op["content"])
                tf_val = None
                if q["section"] == 2:
                    oi = ord(op["key"])-ord("A")
                    if 0 <= oi < len(q["tf_answers"]):
                        tf_val = 1 if q["tf_answers"][oi] else 0

                c.execute("""
                    INSERT INTO options(
                        question_id,option_key,
                        blocks_json,content_json,image_url,correct_tf
                    ) VALUES(?,?,?,?,?,?)
                """,(
                    qid,
                    op["key"],
                    oj,
                    oj,
                    None,
                    tf_val
                ))

    flash(
        "Đã tạo bản nháp LaTeX. Hãy kiểm tra đề, "
        "sau đó bấm Xuất bản cho học sinh."
    )
    return redirect(url_for("edit_exam",exam_id=exam_id))



# ============================================================
# V5.16.0 — SINH CÂU HỎI BẢN NHÁP NỘI BỘ
# Không gọi API ngoài; mục tiêu là nhập yêu cầu xong có câu hiển thị.
# ============================================================
def parse_ai_section_counts(prompt):
    """
    V5.16.4 — đọc chính xác số câu theo từng Phần I/II/III/IV.

    Ví dụ:
      "Phần I 3 câu; Phần II 2 câu; Phần III 1 câu; Phần IV 2 câu"
      -> {1:3, 2:2, 3:1, 4:2}

    Cách làm:
    - tìm vị trí từng cụm "Phần I/II/III/IV";
    - cắt riêng đoạn của mỗi phần đến trước phần kế tiếp;
    - chỉ tìm số câu bên trong đoạn đó.
    Nhờ vậy số câu của Phần II không bị lấy nhầm sang Phần III.
    """
    text = normalize_text(prompt).lower()
    counts = {1:0,2:0,3:0,4:0}

    roman_to_sec = {"i":1, "ii":2, "iii":3, "iv":4}

    matches = list(re.finditer(r"phần\s*(iv|iii|ii|i)\b", text))

    if matches:
        for idx, m in enumerate(matches):
            sec = roman_to_sec[m.group(1)]
            seg_start = m.end()
            seg_end = matches[idx+1].start() if idx+1 < len(matches) else len(text)
            segment = text[seg_start:seg_end]

            # ưu tiên số ngay sau tên phần: "Phần II: 4 câu"
            num = re.search(r"[:\-\s]*(\d{1,2})\s*câu", segment)
            if num:
                counts[sec] = max(1, min(50, int(num.group(1))))
            else:
                # nếu đoạn chỉ ghi tên phần mà không có số, mặc định 1 câu.
                counts[sec] = 1

        # Trường hợp "Tạo 5 câu Phần I" — số nằm trước tên phần.
        for idx, m in enumerate(matches):
            sec = roman_to_sec[m.group(1)]
            left_start = matches[idx-1].end() if idx > 0 else 0
            left = text[left_start:m.start()]
            prior = re.search(r"(\d{1,2})\s*câu\s*$", left.strip())
            if prior:
                counts[sec] = max(1, min(50, int(prior.group(1))))

        return counts

    # Không ghi "Phần", nhận diện theo loại câu.
    type_patterns = [
        (2, r"(\d{1,2})\s*câu[^;,.]*?(?:đúng\s*/?\s*sai|đúng sai)"),
        (3, r"(\d{1,2})\s*câu[^;,.]*?trả lời ngắn"),
        (4, r"(\d{1,2})\s*câu[^;,.]*?tự luận"),
        (1, r"(\d{1,2})\s*câu[^;,.]*?(?:trắc nghiệm nhiều lựa chọn|nhiều lựa chọn)")
    ]
    for sec, pat in type_patterns:
        m = re.search(pat, text)
        if m:
            counts[sec] = max(1, min(50, int(m.group(1))))

    if not any(counts.values()):
        m = re.search(r"(\d{1,2})\s*câu", text)
        counts[1] = max(1, min(50, int(m.group(1)))) if m else 1

    return counts

def ai_topic_from_prompt(prompt):
    text = normalize_text(prompt).lower()
    if "tích phân" in text:
        return "tích phân"
    if "xác suất" in text:
        return "xác suất"
    if "đạo hàm" in text:
        return "đạo hàm"
    if "cực trị" in text:
        return "cực trị"
    if "hàm số" in text:
        return "hàm số"
    if "logarit" in text or "mũ" in text:
        return "mũ logarit"
    return normalize_text(prompt) or "Toán học"


def local_ai_question_draft(prompt, count=1):
    """
    Sinh đúng cấu trúc 4 phần của hệ thống.
    Công thức dùng \\(...\\) để MathJax hiển thị trực tiếp thành toán học.
    """
    counts = parse_ai_section_counts(prompt)
    topic = ai_topic_from_prompt(prompt)
    questions = []

    # PHẦN I: đúng 4 phương án, 1 đáp án đúng.
    for i in range(counts[1]):
        n=i+1
        if topic == "đạo hàm":
            stem = rf"Cho hàm số \(y=x^{n+2}-{n+1}x\). Đạo hàm của hàm số là"
            options = [
                rf"\({n+2}x^{n+1}-{n+1}\)",
                rf"\({n+2}x^{n+1}+{n+1}\)",
                rf"\(x^{n+1}-{n+1}\)",
                rf"\({n+1}x^{n+1}-{n+1}\)"
            ]
            answer="A"
        elif topic == "tích phân":
            stem = rf"Tính \(\int_0^1 {n+1}x^{n}\,dx\)."
            options = [r"\(0\)",r"\(1\)",rf"\({n+1}\)",rf"\(\frac{{1}}{{{n+1}}}\)"]
            answer="B"
        elif topic == "xác suất":
            stem = "Gieo một đồng xu cân đối một lần. Xác suất xuất hiện mặt ngửa bằng"
            options = [r"\(0\)",r"\(\frac{1}{2}\)",r"\(\frac{3}{4}\)",r"\(1\)"]
            answer="B"
        else:
            stem = f"Câu trắc nghiệm {n} theo yêu cầu: {topic}"
            options = ["Phương án A","Phương án B","Phương án C","Phương án D"]
            answer="A"
        questions.append({"section":1,"stem":stem,"options":options,"answer":answer})

    # PHẦN II: đúng 4 ý a,b,c,d và 4 đáp án Đ/S.
    for i in range(counts[2]):
        n=i+1
        if topic in ("đạo hàm","hàm số","cực trị"):
            stem = rf"Cho hàm số \(f(x)=x^2-{2*n}x+1\). Xét tính đúng sai của các khẳng định sau:"
            options = [
                rf"\(f'(x)=2x-{2*n}\).",
                rf"Phương trình \(f'(x)=0\) có nghiệm \(x={n}\).",
                rf"Hàm số nghịch biến trên \((-\infty;{n})\).",
                r"Giá trị nhỏ nhất của hàm số bằng \(1\)."
            ]
            answer="ĐĐĐS"
        elif topic == "tích phân":
            stem = "Xét tính đúng sai của các khẳng định sau:"
            options = [
                r"\(\int_0^1 2x\,dx=1\).",
                r"\(\int_0^2 x\,dx=2\).",
                r"\(\int_0^1 x^2\,dx=\frac13\).",
                r"\(\int_0^1 1\,dx=0\)."
            ]
            answer="ĐĐĐS"
        else:
            stem=f"Câu đúng/sai {n} theo yêu cầu: {topic}"
            options=["Khẳng định a)","Khẳng định b)","Khẳng định c)","Khẳng định d)"]
            answer="ĐSĐS"
        questions.append({"section":2,"stem":stem,"options":options,"answer":answer})

    # PHẦN III: trả lời ngắn, không có phương án.
    for i in range(counts[3]):
        n=i+1
        if topic in ("đạo hàm","hàm số","cực trị"):
            stem = rf"Cho \(f(x)=x^2-{2*n}x+1\). Hoành độ điểm cực tiểu của đồ thị là bao nhiêu?"
            answer=str(n)
        elif topic == "tích phân":
            stem = rf"Tính \(\int_0^1 {n+1}x^{n}\,dx\)."
            answer="1"
        elif topic == "xác suất":
            stem="Gieo một con xúc xắc cân đối. Xác suất xuất hiện số chẵn bằng bao nhiêu?"
            answer="0,5"
        else:
            stem=f"Câu trả lời ngắn {n} theo yêu cầu: {topic}"
            answer="0"
        questions.append({"section":3,"stem":stem,"options":[],"answer":answer})

    # PHẦN IV: tự luận, giáo viên chấm.
    for i in range(counts[4]):
        n=i+1
        if topic == "đạo hàm":
            stem = rf"Cho hàm số \(y=x^3-3x^2+{n}\). Hãy khảo sát sự biến thiên và tìm các điểm cực trị của hàm số."
        elif topic == "tích phân":
            stem = rf"Tính \(I=\int_0^1 (x^{n+1}+x)\,dx\) và trình bày đầy đủ lời giải."
        else:
            stem=f"Câu tự luận {n} theo yêu cầu: {topic}"
        questions.append({"section":4,"stem":stem,"options":[],"answer":""})

    for i,q in enumerate(questions,start=1):
        q["number"]=i
    return questions


def ai_default_points(questions):
    counts={1:0,2:0,3:0,4:0}
    for q in questions:
        sec=int(q.get("section") or 1)
        if sec in counts:
            counts[sec]+=1

    pts={1:0.0,2:float(counts[2]),3:0.0,4:0.0}
    remain=max(0.0,10.0-pts[2])
    active=[sec for sec in (1,3,4) if counts[sec]>0]

    if active:
        share=remain/len(active)
        for sec in active:
            pts[sec]=round(share,2)
        pts[active[-1]]=round(pts[active[-1]]+(10-sum(pts.values())),2)

    return pts


def _ai_matrix_content(filename, raw):
    """Extract a compact text/image representation from an uploaded matrix."""
    suffix = Path(filename or "").suffix.lower()
    if len(raw) > app.config["MAX_CONTENT_LENGTH"]:
        raise ValueError("Tệp ma trận vượt quá giới hạn 50 MB.")

    if suffix == ".doc":
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "matrix.doc"
            source.write_bytes(raw)
            subprocess.run(
                [
                    find_soffice(), "--headless", "--convert-to", "docx",
                    "--outdir", tmp, str(source),
                ],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=90,
            )
            converted = Path(tmp) / "matrix.docx"
            if not converted.exists():
                raise ValueError("Không chuyển được ma trận .doc sang .docx.")
            raw = converted.read_bytes()
            suffix = ".docx"

    if suffix == ".docx":
        document = Document(io.BytesIO(raw))
        rows = [p.text.strip() for p in document.paragraphs if p.text.strip()]
        for table in document.tables:
            for row in table.rows:
                values = [cell.text.strip() for cell in row.cells]
                if any(values):
                    rows.append(" | ".join(values))
        return "\n".join(rows)[:40000], None

    if suffix == ".pdf":
        import pymupdf as fitz
        pdf = fitz.open(stream=raw, filetype="pdf")
        try:
            content = "\n".join(page.get_text("text") for page in pdf)
        finally:
            pdf.close()
        return content[:40000], None

    if suffix in {".xls", ".xlsx"}:
        if suffix == ".xls":
            raise ValueError("Ma trận Excel cũ .xls cần lưu lại thành .xlsx.")
        from openpyxl import load_workbook
        workbook = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        lines = []
        try:
            for sheet in workbook.worksheets:
                lines.append("Trang: " + sheet.title)
                for row in sheet.iter_rows(values_only=True):
                    values = [str(value).strip() for value in row if value is not None]
                    if values:
                        lines.append(" | ".join(values))
        finally:
            workbook.close()
        return "\n".join(lines)[:40000], None

    image_mimes = {
        ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg"
    }
    if suffix in image_mimes:
        # Pillow kiểm tra dữ liệu thật, tránh gửi một tệp giả ảnh lên API.
        with Image.open(io.BytesIO(raw)) as image:
            image.verify()
        return "", (raw, image_mimes[suffix])

    raise ValueError("Định dạng ma trận chưa được hỗ trợ.")


def _normalize_ai_math_delimiters(value):
    """Store AI formulas through the same silent LaTeX pipeline as manual input."""
    text = str(value or "").strip()
    text = re.sub(r"\\\[(.*?)\\\]", lambda m: "$$" + m.group(1) + "$$", text, flags=re.S)
    text = re.sub(r"\\\((.*?)\\\)", lambda m: "$" + m.group(1) + "$", text, flags=re.S)
    return text


def _validate_ai_questions(payload, expected_counts):
    questions = payload.get("questions") if isinstance(payload, dict) else None
    if not isinstance(questions, list) or not questions:
        raise ValueError("AI chưa trả về danh sách câu hỏi.")

    clean = []
    counts = {1: 0, 2: 0, 3: 0, 4: 0}
    for index, item in enumerate(questions, 1):
        if not isinstance(item, dict):
            raise ValueError(f"Câu {index} của AI không đúng cấu trúc.")
        try:
            section = int(item.get("section"))
        except (TypeError, ValueError):
            section = 0
        if section not in counts:
            raise ValueError(f"Câu {index} chưa xác định đúng Phần I/II/III/IV.")

        stem = _normalize_ai_math_delimiters(item.get("stem"))
        options = [
            _normalize_ai_math_delimiters(value)
            for value in (item.get("options") or [])
            if str(value or "").strip()
        ]
        answer = normalize_text(item.get("answer")).strip()
        if not stem:
            raise ValueError(f"Câu {index} không có nội dung.")

        if section == 1:
            if len(options) != 4 or answer.upper() not in {"A", "B", "C", "D"}:
                raise ValueError(
                    f"Câu {index} Phần I cần 4 phương án và một đáp án A/B/C/D."
                )
            answer = answer.upper()
        elif section == 2:
            tf_values = parse_tf_key(answer)
            if len(options) != 4 or len(tf_values) != 4:
                raise ValueError(
                    f"Câu {index} Phần II cần 4 ý và đáp án gồm đúng 4 giá trị Đ/S."
                )
            answer = "".join("Đ" if value else "S" for value in tf_values)
        elif section == 3:
            options = []
            answer = answer.replace(".", ",")
            if not re.fullmatch(r"-?\d+(?:,\d+)?", answer):
                raise ValueError(
                    f"Đáp án Câu {index} Phần III phải là một số, dùng dấu phẩy thập phân."
                )
        else:
            options = []
            answer = ""

        counts[section] += 1
        clean.append({
            "section": section,
            "stem": stem,
            "options": options,
            "answer": answer,
        })

    expected = {sec: int(expected_counts.get(sec, 0)) for sec in counts}
    if counts != expected:
        detail = ", ".join(
            f"Phần {['', 'I', 'II', 'III', 'IV'][sec]}: {counts[sec]}/{expected[sec]}"
            for sec in counts
        )
        raise ValueError("AI trả về sai số lượng câu (" + detail + ").")
    return clean


def generate_ai_question_draft(
    prompt, matrix_text="", matrix_image=None, expected_counts=None
):
    """Generate an exact four-section exam draft with Responses structured output."""
    api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError(
            "Chưa cấu hình OPENAI_API_KEY; không thể tạo đề đúng yêu cầu bằng AI."
        )

    counts = (
        {
            sec: max(0, min(50, int(expected_counts.get(sec, 0))))
            for sec in (1, 2, 3, 4)
        }
        if expected_counts
        else parse_ai_section_counts(
            prompt + ("\n" + matrix_text if matrix_text else "")
        )
    )
    if not any(counts.values()):
        raise ValueError("Cần ít nhất một câu hỏi.")
    count_text = ", ".join(
        f"Phần {['', 'I', 'II', 'III', 'IV'][sec]}: {counts[sec]} câu"
        for sec in (1, 2, 3, 4)
    )
    instructions = (
        "Bạn là giáo viên Toán phổ thông Việt Nam, nhiệm vụ là tạo đề kiểm tra "
        "đúng tuyệt đối theo yêu cầu và ma trận. Không tự thay đổi số lượng câu. "
        "Phần I: mỗi câu có 4 phương án, answer là A/B/C/D. "
        "Phần II: mỗi câu có đúng 4 ý, answer là chuỗi 4 ký tự Đ/S. "
        "Phần III: không có phương án, answer là số dùng dấu phẩy thập phân. "
        "Phần IV: không có phương án và answer để trống. "
        "Không viết tiền tố Câu 1 hay A./B./C./D. trong stem/options vì giao diện tự thêm. "
        "Công thức phải là LaTeX chuẩn đặt trong $...$ hoặc $$...$$; dùng \\dfrac, "
        "\\sqrt, \\lim, \\min, \\max, \\overrightarrow khi phù hợp. "
        "Câu hỏi phải có dữ kiện đủ, đáp án chính xác và phù hợp cấp học được yêu cầu. "
        "Nếu giáo viên nêu lớp, chủ đề/chương, mức độ nhận biết-thông hiểu-vận dụng, "
        "dạng toán, phạm vi kiến thức, ngữ cảnh thực tế hoặc yêu cầu tránh một phương pháp, "
        "phải bám đúng các ràng buộc đó. Không sinh câu mẫu chung chung khác chủ đề. "
        "Mỗi câu phải độc lập, không lặp lại dữ kiện/cấu trúc một cách máy móc."
    )
    user_text = (
        "YÊU CẦU CỦA GIÁO VIÊN:\n" + prompt +
        "\n\nSỐ CÂU BẮT BUỘC:\n" + count_text
    )
    if matrix_text:
        user_text += "\n\nNỘI DUNG MA TRẬN/TỆP THAM CHIẾU:\n" + matrix_text

    content = [{"type": "input_text", "text": user_text}]
    if matrix_image:
        raw, mime = matrix_image
        content.append({
            "type": "input_image",
            "image_url": "data:" + mime + ";base64," + base64.b64encode(raw).decode("ascii"),
            "detail": "high",
        })

    question_schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "section": {"type": "integer", "enum": [1, 2, 3, 4]},
            "stem": {"type": "string"},
            "options": {"type": "array", "items": {"type": "string"}},
            "answer": {"type": "string"},
        },
        "required": ["section", "stem", "options", "answer"],
    }
    body = json.dumps({
        "model": os.environ.get("OPENAI_MODEL", "gpt-5.6-luna"),
        "instructions": instructions,
        "input": [{"role": "user", "content": content}],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "vietnamese_math_exam",
                "strict": True,
                "schema": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "questions": {
                            "type": "array",
                            "items": question_schema,
                            "minItems": sum(counts.values()),
                            "maxItems": sum(counts.values()),
                        }
                    },
                    "required": ["questions"],
                },
            }
        },
        "max_output_tokens": min(50000, max(12000, 1800 * sum(counts.values()))),
        "store": False,
    }, ensure_ascii=False).encode("utf-8")
    api_request = urllib_request.Request(
        "https://api.openai.com/v1/responses",
        data=body,
        headers={
            "Authorization": "Bearer " + api_key,
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib_request.urlopen(api_request, timeout=240) as response:
            response_payload = json.loads(response.read().decode("utf-8"))
    except urllib_error.HTTPError as exc:
        app.logger.exception("Tạo đề AI nhận HTTP %s", exc.code)
        raise RuntimeError(
            "Dịch vụ AI từ chối yêu cầu. Hãy kiểm tra API key, model và hạn mức."
        ) from exc
    except urllib_error.URLError as exc:
        raise RuntimeError("Máy chủ chưa kết nối được với dịch vụ AI.") from exc

    output_text = _response_output_text(response_payload)
    if not output_text:
        raise RuntimeError("Dịch vụ AI chưa trả về nội dung đề.")
    try:
        generated = json.loads(output_text)
    except json.JSONDecodeError as exc:
        raise RuntimeError("Dịch vụ AI trả về dữ liệu đề không hợp lệ.") from exc
    return _validate_ai_questions(generated, counts)


@app.route("/teacher/create-exam/ai", methods=["GET","POST"])
@login_required("teacher")
def create_exam_ai():
    """
    Giao diện nhập yêu cầu hoặc upload ma trận và tạo bản nháp bằng AI thật.
    Đầu ra được ép schema, kiểm tra số câu và kiểm tra đáp án trước khi hiển thị.
    """
    draft = None
    matrix_name = None

    if request.method == "POST":
        prompt = normalize_text(request.form.get("prompt"))
        matrix = request.files.get("matrix")
        explicit_counts = {}
        for sec in (1, 2, 3, 4):
            try:
                explicit_counts[sec] = max(
                    0, min(50, int(request.form.get(f"count_p{sec}") or 0))
                )
            except ValueError:
                explicit_counts[sec] = 0
        use_explicit_counts = any(explicit_counts.values())

        matrix_text = ""
        matrix_image = None
        if matrix and matrix.filename:
            matrix_name = secure_filename(matrix.filename)
            raw_matrix = matrix.read()
            try:
                matrix_text, matrix_image = _ai_matrix_content(
                    matrix_name, raw_matrix
                )
            except Exception as exc:
                flash("Không đọc được ma trận: " + str(exc))
                return redirect(request.url)
            matrix_dir = INSTANCE_DIR / "ai_matrices"
            matrix_dir.mkdir(parents=True, exist_ok=True)
            (matrix_dir / (uuid.uuid4().hex + "_" + matrix_name)).write_bytes(
                raw_matrix
            )

        if not prompt and not matrix_name:
            flash("Hãy nhập yêu cầu tạo đề hoặc tải ma trận lên.")
            return redirect(request.url)

        request_text = prompt or "Tạo đề đúng theo ma trận đã tải lên."
        try:
            questions = generate_ai_question_draft(
                request_text,
                matrix_text=matrix_text,
                matrix_image=matrix_image,
                expected_counts=explicit_counts if use_explicit_counts else None,
            )
        except Exception as exc:
            flash("Không thể tạo đề bằng AI: " + str(exc))
            return render_template(
                "create_exam_ai.html",
                draft=None,
                matrix_name=matrix_name,
                default_points={1: 10, 2: 0, 3: 0, 4: 0},
                ai_configured=bool(os.environ.get("OPENAI_API_KEY", "").strip()),
                ai_model=os.environ.get("OPENAI_MODEL", "gpt-5.6-luna"),
                prompt_value=prompt,
                count_values=explicit_counts,
            )
        draft = {
            "title": "Bản nháp đề kiểm tra",
            "request": request_text,
            "matrix": matrix_name,
            "sections": [
                "Phần I. Trắc nghiệm nhiều lựa chọn",
                "Phần II. Trắc nghiệm đúng/sai",
                "Phần III. Trả lời ngắn",
                "Phần IV. Tự luận"
            ],
            "questions": questions
        }

    default_points = ai_default_points(draft["questions"]) if draft else {1:10,2:0,3:0,4:0}
    return render_template(
        "create_exam_ai.html",
        draft=draft,
        matrix_name=matrix_name,
        default_points=default_points,
        ai_configured=bool(os.environ.get("OPENAI_API_KEY", "").strip()),
        ai_model=os.environ.get("OPENAI_MODEL", "gpt-5.6-luna"),
        prompt_value="",
        count_values={1: 0, 2: 0, 3: 0, 4: 0},
    )



# ============================================================

# ============================================================

# ============================================================
# V5.17.0 — MENU ĐẢO ĐỀ / TRỘN ĐỀ
# ------------------------------------------------------------
# QUY TẮC:
# - Phần I: đảo câu + đảo A/B/C/D.
# - Phần II: chỉ đảo câu, KHÔNG đảo 4 ý a/b/c/d.
# - Phần III: chỉ đảo câu.
# - Phần IV: đảo câu.
# - Đáp án luôn đi theo câu.
# - Hình/công thức nằm trong blocks_json nên luôn đi cùng câu/đáp án.
# ============================================================
def ensure_shuffle_schema(c):
    c.execute("""
        CREATE TABLE IF NOT EXISTS shuffle_sources(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            source_file TEXT NOT NULL,
            uploaded_by INTEGER NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS shuffle_questions(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_id INTEGER NOT NULL,
            original_order INTEGER NOT NULL,
            section_order INTEGER NOT NULL,
            section INTEGER NOT NULL,
            qtype TEXT NOT NULL,
            blocks_json TEXT NOT NULL DEFAULT '[]',
            solution_json TEXT NOT NULL DEFAULT '[]',
            correct_option TEXT,
            short_answer TEXT,
            tf_answers_json TEXT NOT NULL DEFAULT '[]'
        )
    """)
    shuffle_qcols = {
        row[1] for row in c.execute("PRAGMA table_info(shuffle_questions)").fetchall()
    }
    if "solution_json" not in shuffle_qcols:
        c.execute(
            "ALTER TABLE shuffle_questions "
            "ADD COLUMN solution_json TEXT NOT NULL DEFAULT '[]'"
        )
    c.execute("""
        CREATE TABLE IF NOT EXISTS shuffle_options(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shuffle_question_id INTEGER NOT NULL,
            option_key TEXT NOT NULL,
            blocks_json TEXT NOT NULL DEFAULT '[]',
            correct_tf INTEGER,
            UNIQUE(shuffle_question_id, option_key)
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS shuffle_versions(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_id INTEGER NOT NULL,
            version_json TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # V5.18.0 — thông tin đầu đề + số mã đề.
    for col, ddl in [
        ("department","TEXT DEFAULT ''"),
        ("school","TEXT DEFAULT ''"),
        ("exam_name","TEXT DEFAULT ''"),
        ("subject","TEXT DEFAULT ''"),
        ("school_year","TEXT DEFAULT ''"),
        ("duration_text","TEXT DEFAULT '90 phút'"),
        ("version_count","INTEGER DEFAULT 1"),
        ("start_code","INTEGER DEFAULT 689"),
        ("codes_json","TEXT DEFAULT '[]'"),
        ("stored_name","TEXT DEFAULT ''")
    ]:
        try:
            c.execute(f"ALTER TABLE shuffle_sources ADD COLUMN {col} {ddl}")
        except sqlite3.OperationalError:
            pass


def load_shuffle_questions(c, source_id):
    rows = c.execute("""
        SELECT * FROM shuffle_questions
        WHERE source_id=?
        ORDER BY section,section_order,original_order,id
    """,(source_id,)).fetchall()

    out = []
    for q in rows:
        ops = c.execute("""
            SELECT * FROM shuffle_options
            WHERE shuffle_question_id=?
            ORDER BY option_key
        """,(q["id"],)).fetchall()

        out.append({
            "id": q["id"],
            "section": q["section"],
            "section_order": q["section_order"],
            "original_order": q["original_order"],
            "qtype": q["qtype"],
            "blocks": json.loads(q["blocks_json"] or "[]"),
            "solution": json.loads(q["solution_json"] or "[]"),
            "correct_option": q["correct_option"],
            "short_answer": q["short_answer"],
            "tf_answers": json.loads(q["tf_answers_json"] or "[]"),
            "options": [
                {
                    "key": op["option_key"],
                    "blocks": json.loads(op["blocks_json"] or "[]"),
                    "correct_tf": op["correct_tf"]
                }
                for op in ops
            ]
        })

    return out


def make_shuffle_version(questions):
    """
    Sinh một mã trộn và remap đáp án Phần I chính xác.

    JSON chỉ lưu IDs + thứ tự option, không copy nội dung nên:
    hình/công thức/câu hỏi không thể bị tách rời.
    """
    by_section = {1:[],2:[],3:[],4:[]}
    qmap = {q["id"]: q for q in questions}

    for q in questions:
        by_section[q["section"]].append(q["id"])

    for sec in (1,2,3,4):
        random.shuffle(by_section[sec])

    entries = []

    for sec in (1,2,3,4):
        for new_no, qid in enumerate(by_section[sec], start=1):
            q = qmap[qid]
            option_order = [o["key"] for o in q["options"]]
            answer_display = ""

            if sec == 1:
                # Chỉ Phần I đảo phương án.
                random.shuffle(option_order)

                # Đáp án mới = vị trí của key đúng cũ trong thứ tự option mới.
                old_correct = (q["correct_option"] or "").upper()
                if old_correct in option_order:
                    new_idx = option_order.index(old_correct)
                    answer_display = "ABCD"[new_idx]

            elif sec == 2:
                # Giữ nguyên A/B/C/D và chuỗi Đ/S tương ứng.
                vals = q["tf_answers"]
                answer_display = "".join("Đ" if x else "S" for x in vals)

            elif sec == 3:
                answer_display = q["short_answer"] or ""

            else:
                answer_display = "Giáo viên chấm"

            entries.append({
                "section": sec,
                "number": new_no,
                "question_id": qid,
                "option_order": option_order,
                "answer": answer_display
            })

    return entries


@app.route("/teacher/shuffle-exams", methods=["GET","POST"])
@login_required("teacher")
def shuffle_exam_manager():
    if request.method == "POST":
        f = request.files.get("file")
        exam_name = normalize_text(request.form.get("exam_name"))
        subject = normalize_text(request.form.get("subject"))
        title = exam_name or subject or Path(f.filename).stem if f and f.filename else "Đề cần trộn"

        if not f or not f.filename or ext(f.filename) not in ALLOWED:
            flash("Đảo đề chỉ nhận file DOCX hoặc PDF.")
            return redirect(request.url)

        try:
            version_count = max(1, min(100, int(request.form.get("version_count") or 1)))
        except Exception:
            version_count = 1

        raw_codes = (request.form.get("exam_codes") or "").strip()
        exam_codes = [
            x.strip()
            for x in re.split(r"[,;\n\r]+", raw_codes)
            if x.strip()
        ]

        if len(exam_codes) != version_count:
            flash(
                f"Bạn chọn {version_count} đề nên cần nhập đúng {version_count} mã đề. "
                f"Hiện đã nhập {len(exam_codes)} mã."
            )
            return redirect(request.url)

        if len(set(exam_codes)) != len(exam_codes):
            flash("Các mã đề không được trùng nhau.")
            return redirect(request.url)

        if any(not re.fullmatch(r"[0-9]{1,12}", code) for code in exam_codes):
            flash("Mỗi mã đề chỉ gồm chữ số, tối đa 12 chữ số.")
            return redirect(request.url)

        saved = UPLOAD_DIR / (
            uuid.uuid4().hex + "_shuffle_" + secure_filename(f.filename)
        )

        # ============================================================
        # LƯU FILE NGUỒN NHỊ PHÂN NGUYÊN BẢN
        # Không đọc text rồi save lại bằng python-docx.
        # File gốc tải lên được lưu trực tiếp bằng FileStorage.save().
        # ============================================================
        f.save(saved)

        try:
            source_ext = ext(f.filename)

            # Dùng nguyên parser Word/PDF hiện có.
            if source_ext == "docx":
                lines = extract_word_blocks(saved)
                questions = build_questions(lines)
                # V5.24.13: kích thước hình theo đúng từng câu trong OOXML nguồn, giữ preview Word.
                _sync_docx_question_figures_from_ooxml(questions, saved)
            else:
                lines = extract_pdf_blocks(saved)
                questions = build_questions(lines)
                if not questions:
                    lines = extract_pdf_blocks_fallback(saved)
                    questions = build_questions(lines)

            if not questions:
                raise RuntimeError("Không nhận diện được câu hỏi trong đề.")

            with db() as c:
                ensure_shuffle_schema(c)

                source_id = c.execute("""
                    INSERT INTO shuffle_sources(
                        title,source_file,uploaded_by,
                        department,school,exam_name,subject,school_year,
                        duration_text,version_count,start_code,codes_json,stored_name
                    ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,(
                    title,f.filename,session["user_id"],
                    (request.form.get("department") or "").strip(),
                    (request.form.get("school") or "").strip(),
                    exam_name,
                    subject,
                    (request.form.get("school_year") or "").strip(),
                    (request.form.get("duration_text") or "90 phút").strip(),
                    version_count,689,jdump(exam_codes),saved.name
                )).lastrowid

                for q in questions:
                    qid = c.execute("""
                        INSERT INTO shuffle_questions(
                            source_id,original_order,section_order,section,qtype,
                            blocks_json,solution_json,correct_option,short_answer,tf_answers_json
                        ) VALUES(?,?,?,?,?,?,?,?,?,?)
                    """,(
                        source_id,
                        q["original_order"],
                        q["section_order"],
                        q["section"],
                        q["qtype"],
                        jdump(q["stem"]),
                        jdump(q.get("solution",[])),
                        q.get("correct_option"),
                        q.get("short_answer"),
                        jdump(q.get("tf_answers",[]))
                    )).lastrowid

                    for op in q.get("options",[]):
                        tf_val = None
                        if q["section"] == 2:
                            vals = q.get("tf_answers",[])
                            oi = ord(op["key"].upper()) - ord("A")
                            if 0 <= oi < len(vals):
                                tf_val = 1 if vals[oi] else 0

                        c.execute("""
                            INSERT INTO shuffle_options(
                                shuffle_question_id,option_key,
                                blocks_json,correct_tf
                            ) VALUES(?,?,?,?)
                        """,(
                            qid,
                            op["key"],
                            jdump(op["content"]),
                            tf_val
                        ))

            flash(
                f"Đã tải {len(questions)} câu. "
                "Bấm Kiểm tra để xem đề hoặc Đảo để tạo mã đề."
            )
            return redirect(url_for("shuffle_exam_manager"))

        except Exception as e:
            flash("Lỗi tải đề để trộn: " + str(e))
            return redirect(request.url)

    with db() as c:
        ensure_shuffle_schema(c)
        sources = c.execute("""
            SELECT s.*,
                   (SELECT COUNT(*) FROM shuffle_questions q
                    WHERE q.source_id=s.id) AS question_count
            FROM shuffle_sources s
            WHERE s.uploaded_by=?
            ORDER BY s.id DESC
        """,(session["user_id"],)).fetchall()

    return render_template(
        "shuffle_exam_manager.html",
        sources=sources
    )




@app.get("/teacher/shuffle-exams/<int:source_id>/word-preview")
@login_required("teacher")
def shuffle_exam_word_preview(source_id):
    """
    V5.21.0 — Xem file DOCX nguồn qua Mammoth HTML.

    Đây là màn hình THAM CHIẾU định dạng Word.
    Luồng thi online/trộn/chấm không thay đổi.
    """
    with db() as c:
        ensure_shuffle_schema(c)
        source = c.execute("""
            SELECT * FROM shuffle_sources
            WHERE id=? AND uploaded_by=?
        """,(source_id,session["user_id"])).fetchone()

    if not source:
        abort(404)

    if not (source["source_file"] or "").lower().endswith(".docx"):
        flash("Mammoth preview chỉ áp dụng cho file DOCX.")
        return redirect(url_for("shuffle_exam_manager"))

    stored = source["stored_name"] if "stored_name" in source.keys() else ""
    path = UPLOAD_DIR / stored if stored else None
    if not path or not path.exists():
        flash("Không tìm thấy file Word nguồn.")
        return redirect(url_for("shuffle_exam_manager"))

    # V5.24.11: Mammoth/LibreOffice chỉ có thể hiển thị MathType OLE dưới
    # dạng preview image. Để màn hình "Xem đề gốc" không làm giáo viên
    # hiểu nhầm MathType đã bị mất, dùng chính dữ liệu đã parse sang LaTeX
    # và MathJax ở trang Kiểm tra. File nguồn tải xuống vẫn là DOCX nguyên bản.
    return redirect(url_for("shuffle_exam_check", source_id=source_id))


@app.get("/teacher/shuffle-exams/<int:source_id>/source")
@login_required("teacher")
def shuffle_exam_source_download(source_id):
    """
    V5.20.4 — tải lại đúng file Word/PDF giáo viên đã upload.

    Màn hình Kiểm tra đề chưa có version_id nên không được gọi
    shuffle_exam_docx(source_id, version_id).
    """
    with db() as c:
        ensure_shuffle_schema(c)
        source = c.execute("""
            SELECT * FROM shuffle_sources
            WHERE id=? AND uploaded_by=?
        """,(source_id,session["user_id"])).fetchone()

    if not source:
        abort(404)

    stored = source["stored_name"] if "stored_name" in source.keys() else ""
    if not stored:
        flash(
            "Đề này được tải từ phiên bản cũ nên chưa lưu đường dẫn file nguồn. "
            "Hãy xóa đề và tải lại một lần."
        )
        return redirect(url_for("shuffle_exam_manager"))

    path = UPLOAD_DIR / stored
    if not path.exists():
        flash("Không tìm thấy file nguồn trên máy chủ.")
        return redirect(url_for("shuffle_exam_manager"))

    return send_file(
        path,
        as_attachment=True,
        download_name=source["source_file"] or path.name
    )


@app.get("/teacher/shuffle-exams/<int:source_id>/check")
@login_required("teacher")
def shuffle_exam_check(source_id):
    ensure_shuffle_question_media(source_id)
    with db() as c:
        ensure_shuffle_schema(c)
        source = c.execute("""
            SELECT * FROM shuffle_sources
            WHERE id=? AND uploaded_by=?
        """,(source_id,session["user_id"])).fetchone()

        if not source:
            abort(404)

        questions = load_shuffle_questions(c, source_id)

    # Tạo cấu trúc preview nguyên bản.
    entries = []
    for q in questions:
        if q["section"] == 1:
            answer = q["correct_option"] or ""
        elif q["section"] == 2:
            answer = "".join("Đ" if x else "S" for x in q["tf_answers"])
        elif q["section"] == 3:
            answer = q["short_answer"] or ""
        else:
            answer = "Giáo viên chấm"

        entries.append({
            "section": q["section"],
            "number": q["section_order"],
            "question": q,
            "option_order": [o["key"] for o in q["options"]],
            "answer": answer
        })

    return render_template(
        "shuffle_exam_view.html",
        source=source,
        entries=entries,
        shuffled=False,
        version_id=None
    )


@app.post("/teacher/shuffle-exams/<int:source_id>/mix")
@login_required("teacher")
def shuffle_exam_mix(source_id):
    """
    V5.20.2 — ĐẢO ĐỀ AN TOÀN

    - Validate đủ câu trước khi random.
    - Phần I cần đúng 4 option và đáp án A/B/C/D.
    - Phần II tuyệt đối không đảo option.
    - Không gọi LibreOffice khi bấm Đảo; chỉ tạo dữ liệu trộn và mở
      trang xem HTML, tránh lỗi preview làm người dùng tưởng nút Đảo lỗi.
    - Nếu có lỗi, rollback và báo rõ ngay trên trang Giáo viên.
    """
    try:
        with db() as c:
            ensure_shuffle_schema(c)

            source=c.execute(
                "SELECT * FROM shuffle_sources WHERE id=? AND uploaded_by=?",
                (source_id,session["user_id"])
            ).fetchone()
            if not source:
                abort(404)

            questions=load_shuffle_questions(c,source_id)
            if not questions:
                flash("Đề này chưa có câu hỏi để đảo. Hãy xóa và tải lại file Word/PDF.")
                return redirect(url_for("shuffle_exam_manager"))

            try:
                codes=json.loads(source["codes_json"] or "[]")
            except Exception:
                codes=[]

            count=max(1,min(100,int(source["version_count"] or 1)))

            # Legacy source from older build: show useful message instead of crashing.
            if len(codes)!=count:
                flash(
                    "Đề này được tải ở phiên bản cũ hoặc danh sách mã đề chưa đủ. "
                    "Hãy xóa đề này, tải lại file và nhập đúng số mã đề tương ứng."
                )
                return redirect(url_for("shuffle_exam_manager"))

            # Strict validation before random.
            for q in questions:
                sec=int(q.get("section") or 1)
                ops=q.get("options") or []

                if sec==1:
                    keys=[str(o.get("key","")).upper() for o in ops]
                    if keys!=["A","B","C","D"]:
                        flash(
                            f"Phần I - Câu {q.get('section_order')}: "
                            "không nhận đủ 4 phương án A/B/C/D. "
                            "Hãy bấm Kiểm tra để xem lại file nguồn."
                        )
                        return redirect(url_for("shuffle_exam_manager"))
                    if (q.get("correct_option") or "").upper() not in keys:
                        flash(
                            f"Phần I - Câu {q.get('section_order')}: "
                            "chưa nhận đúng đáp án A/B/C/D."
                        )
                        return redirect(url_for("shuffle_exam_manager"))

                elif sec==2:
                    # Phần II giữ nguyên option order.
                    keys=[str(o.get("key","")).upper() for o in ops]
                    if keys and keys!=["A","B","C","D"]:
                        flash(
                            f"Phần II - Câu {q.get('section_order')}: "
                            "không nhận đủ 4 ý a/b/c/d."
                        )
                        return redirect(url_for("shuffle_exam_manager"))

            # Recreate exactly requested number of versions.
            c.execute("DELETE FROM shuffle_versions WHERE source_id=?",(source_id,))

            for exam_code in codes:
                entries=make_shuffle_version(questions)

                # Guardrail: Part II order can never change.
                for item in entries:
                    if int(item.get("section") or 0)==2:
                        q=next(
                            (x for x in questions if x["id"]==item["question_id"]),
                            None
                        )
                        if q:
                            original=[o["key"] for o in q.get("options",[])]
                            item["option_order"]=original

                payload={
                    "exam_code":str(exam_code),
                    "entries":entries
                }
                c.execute(
                    "INSERT INTO shuffle_versions(source_id,version_json) VALUES(?,?)",
                    (source_id,jdump(payload))
                )

        flash(f"Đã đảo thành công {count} mã đề.")
        return redirect(url_for("shuffle_exam_versions",source_id=source_id))

    except Exception as e:
        # Keep app alive and expose actionable error instead of Flask traceback.
        flash("Lỗi khi đảo đề: "+str(e))
        return redirect(url_for("shuffle_exam_manager"))


@app.get("/teacher/shuffle-exams/<int:source_id>/versions")
@login_required("teacher")
def shuffle_exam_versions(source_id):
    ensure_shuffle_question_media(source_id)
    with db() as c:
        ensure_shuffle_schema(c)
        source=c.execute(
            "SELECT * FROM shuffle_sources WHERE id=? AND uploaded_by=?",
            (source_id,session["user_id"])
        ).fetchone()
        if not source:
            abort(404)

        rows=c.execute(
            "SELECT * FROM shuffle_versions WHERE source_id=? ORDER BY id",
            (source_id,)
        ).fetchall()

        questions=load_shuffle_questions(c,source_id)
        qmap={q["id"]:q for q in questions}

        versions=[]
        for r in rows:
            data=json.loads(r["version_json"])
            entries=[]
            for item in data.get("entries",[]):
                q=qmap.get(item["question_id"])
                if not q:
                    continue
                entries.append({
                    "section":item["section"],
                    "number":item["number"],
                    "question":q,
                    "option_order":item["option_order"],
                    "answer":item["answer"]
                })
            versions.append({
                "id":r["id"],
                "exam_code":str(data.get("exam_code","")),
                "entries":entries
            })

    return render_template(
        "shuffle_exam_versions.html",
        source=source,
        versions=versions
    )


@app.get("/teacher/shuffle-exams/<int:source_id>/version/<int:version_id>")
@login_required("teacher")
def shuffle_exam_version(source_id, version_id):
    ensure_shuffle_question_media(source_id)
    with db() as c:
        ensure_shuffle_schema(c)

        source = c.execute("""
            SELECT * FROM shuffle_sources
            WHERE id=? AND uploaded_by=?
        """,(source_id,session["user_id"])).fetchone()

        if not source:
            abort(404)

        row = c.execute("""
            SELECT * FROM shuffle_versions
            WHERE id=? AND source_id=?
        """,(version_id,source_id)).fetchone()

        if not row:
            abort(404)

        payload = json.loads(row["version_json"])
        version = payload.get("entries", payload if isinstance(payload,list) else [])
        exam_code = payload.get("exam_code", int(source["start_code"] or 689))
        questions = load_shuffle_questions(c, source_id)
        qmap = {q["id"]: q for q in questions}

    entries = []
    for item in version:
        q = qmap.get(item["question_id"])
        if not q:
            continue
        entries.append({
            "section": item["section"],
            "number": item["number"],
            "question": q,
            "option_order": item["option_order"],
            "answer": item["answer"]
        })

    return render_template(
        "shuffle_exam_view.html",
        source=source,
        entries=entries,
        shuffled=True,
        version_id=version_id,
        exam_code=exam_code
    )



def _set_tnr(run,size=13.5,bold=False,italic=False):
    run.font.name="Times New Roman"; run.font.size=Pt(size)
    run.bold=bold; run.italic=italic
    rPr=run._element.get_or_add_rPr()
    rFonts=rPr.rFonts
    if rFonts is None:
        rFonts=OxmlElement("w:rFonts"); rPr.insert(0,rFonts)
    for a in ("ascii","hAnsi","eastAsia","cs"):
        rFonts.set(qn("w:"+a),"Times New Roman")

def _addp(doc,text="",bold=False,align=None,size=13.5,italic=False):
    p=doc.add_paragraph()
    if align is not None: p.alignment=align
    p.paragraph_format.space_after=Pt(2)
    r=p.add_run(text); _set_tnr(r,size,bold,italic)
    return p

def _blocks_plain(blocks):
    return " ".join((b.get("value","") for b in blocks or [] if b.get("type") in ("text","latex"))).strip()

_MATHTYPE_EXPORT_LOCK = threading.Lock()


def _mathtype_tex_compat(latex):
    """Normalize LaTeX to a MathType 7 Toggle-TeX-friendly subset."""
    value=str(latex or '').strip()
    if not value:
        return value
    value=value.replace(r'\dfrac',r'\frac').replace(r'\tfrac',r'\frac')
    value=value.replace(r'\limits','')
    value=re.sub(r'\\operatorname\{([^{}]+)\}',r'\\mathrm{\1}',value)
    value=re.sub(r'\\(?:,|;|:|!)(?![A-Za-z])',' ',value)
    return re.sub(r'\s+',' ',value).strip()


def _coalesce_mathtype_blocks(blocks):
    """Merge consecutive PDF LaTeX fragments into one editable formula."""
    out=[]; pending=[]; pending_display=False
    def flush():
        nonlocal pending,pending_display
        if pending:
            out.append({'type':'latex','value':' '.join(pending).strip(),
                        'display':'block' if pending_display else 'inline'})
        pending=[]; pending_display=False
    for b in (blocks or []):
        if b.get('type')=='latex':
            val=_mathtype_tex_compat(b.get('value',''))
            if val:
                pending.append(val)
                pending_display = pending_display or str(b.get('display') or '').lower()=='block'
            continue
        flush(); out.append(b)
    flush(); return out


def _coalesce_plain_latex_blocks(blocks):
    """Merge adjacent PDF LaTeX fragments without MathType-specific rewriting."""
    out=[]; pending=[]; pending_display=False
    def flush():
        nonlocal pending,pending_display
        if pending:
            out.append({
                'type':'latex',
                'value':' '.join(pending).strip(),
                'display':'block' if pending_display else 'inline'
            })
        pending=[]; pending_display=False
    for b in (blocks or []):
        if b.get('type')=='latex':
            val=normalise_latex(b.get('value','')) or str(b.get('value') or '').strip()
            if val:
                pending.append(val)
                pending_display = pending_display or str(b.get('display') or '').lower()=='block'
            continue
        flush(); out.append(b)
    flush()
    return out


def _mathtype_toggle_payload(latex):
    '''Make PDF TeX robust for MathType 7 Toggle TeX without visible changes.

    MathType 7 documents a Toggle TeX limitation for strings containing no
    operator (for example abc or 2x). PDF extraction naturally creates many
    such atomic math runs. Add a zero-width vertical phantom containing '+'
    only to those runs; Texvc then sees an operator while the rendered formula
    remains visually unchanged.
    '''
    value=_mathtype_tex_compat(latex)
    if not value:
        return value
    operator_re=(
        r'[+\-=<>]|'
        r'\\(?:pm|mp|times|cdot|div|ast|star|circ|bullet|oplus|otimes|cup|cap|setminus|'
        r'wedge|vee|leq?|geq?|ne|neq|approx|sim|simeq|equiv|propto|in|notin|ni|subset|'
        r'supset|subseteq|supseteq|to|mapsto|rightarrow|leftarrow|Rightarrow|Leftarrow|'
        r'leftrightarrow|parallel|perp|mid)\b'
    )
    if re.search(operator_re,value):
        return value
    if r'\vphantom{+}' not in value:
        value += r'\vphantom{+}'
    return value


def _add_blocks_docx(doc,blocks,prefix="", latex_for_mathtype=False, plain_latex=False):
    """
    V5.19.0 — tái sử dụng metadata kích thước công thức do parser V5.12.x tạo.
    Không ép mọi MathType về cùng width; dùng display_height_px cho formula
    và display_width_px cho figure để giữ đúng tỷ lệ/glyph.
    """
    p=doc.add_paragraph()
    p.paragraph_format.space_after=Pt(2)
    p.paragraph_format.line_spacing=1.0

    if prefix:
        _set_tnr(p.add_run(prefix),13.5,True)

    render_blocks=(
        _coalesce_mathtype_blocks(blocks) if latex_for_mathtype
        else _coalesce_plain_latex_blocks(blocks) if plain_latex
        else (blocks or [])
    )
    for b in render_blocks:
        btype=b.get("type")

        if btype in ("text","latex"):
            val=b.get("value","")
            if val:
                if btype=="latex" and latex_for_mathtype:
                    # Legacy path kept only for compatibility. PDF shuffle export
                    # no longer invokes MathType automatically from V5.24.21.
                    val=_mathtype_toggle_payload(val)
                    if str(b.get("display") or "").lower()=="block":
                        val="\\["+val+"\\]"
                    else:
                        val="$"+val+"$"
                elif btype=="latex" and plain_latex:
                    # V5.24.21: PDF -> shuffled Word exports canonical LaTeX,
                    # exactly like the PDF-to-Word utility. No MathType COM call.
                    val=normalise_latex(val) or str(val).strip()
                    if str(b.get("display") or "").lower()=="block":
                        val="\\["+val+"\\]"
                    else:
                        val="$"+val+"$"
                run=p.add_run(val)
                if btype=="latex" and plain_latex:
                    _set_word_run_font(run,"Cambria Math",13.5)
                else:
                    _set_tnr(run,13.5)

        elif btype=="image" and b.get("url"):
            try:
                path=_question_media_path(b["url"])
                if not path:
                    continue

                rr=p.add_run()

                if b.get("role")=="figure":
                    word_cm=float(b.get("word_width_cm") or 0)
                    wpx=float(b.get("display_width_px") or 0)
                    if word_cm>0:
                        rr.add_picture(str(path),width=Cm(min(16.0,max(1.0,word_cm))))
                    elif wpx>0:
                        width_cm=min(16.0,max(1.0,wpx*2.54/96.0))
                        rr.add_picture(str(path),width=Cm(width_cm))
                    else:
                        rr.add_picture(str(path),width=Cm(6.5))
                else:
                    # 13.5pt ≈ 18 CSS px. Parser V5.12.x đã hiệu chỉnh riêng
                    # Word/PDF bằng display_height_px để glyph nhìn cùng cỡ.
                    hpx=float(b.get("display_height_px") or 18)
                    height_cm=max(0.42,min(1.45,hpx*2.54/96.0))
                    rr.add_picture(str(path),height=Cm(height_cm))
            except Exception:
                pass


def _apply_table_grid(table):
    """Create table borders even when source DOCX has no 'Table Grid' style."""
    tblPr = table._tbl.tblPr
    borders = tblPr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)

    for edge in ("top","left","bottom","right","insideH","insideV"):
        tag="w:"+edge
        node=borders.find(qn(tag))
        if node is None:
            node=OxmlElement(tag)
            borders.append(node)
        node.set(qn("w:val"),"single")
        node.set(qn("w:sz"),"8")
        node.set(qn("w:space"),"0")
        node.set(qn("w:color"),"000000")


def _cell(cell,text,bold=False,size=10):
    cell.text=""; p=cell.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    _set_tnr(p.add_run(str(text)),size,bold)
    cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER


def _write_solution_docx(doc, q, latex_for_mathtype=False, plain_latex=False):
    sol=q.get('solution') or []
    if not sol:
        return
    _addp(doc,'Lời giải',True,WD_ALIGN_PARAGRAPH.CENTER,13.5)
    _add_blocks_docx(doc,sol,'',latex_for_mathtype=latex_for_mathtype,plain_latex=plain_latex)


def _count_mathtype_ole_docx(path):
    try:
        with zipfile.ZipFile(path) as z:
            root=etree.fromstring(z.read('word/document.xml'))
        return len(root.xpath(".//*[local-name()='OLEObject' and @ProgID='Equation.DSMT4']"))
    except Exception:
        return 0


def _count_tex_markers_docx(path):
    """Đếm công thức TeX do exporter PDF->Word tự chèn trước khi Toggle TeX."""
    try:
        with zipfile.ZipFile(path) as z:
            root=etree.fromstring(z.read('word/document.xml'))
        count=0
        for node in root.xpath(".//*[local-name()='t']"):
            text=node.text or ''
            if (text.startswith('$') and text.endswith('$') and len(text)>=2) or (
                text.startswith('\\[') and text.endswith('\\]') and len(text)>=4
            ):
                count+=1
        return count
    except Exception:
        return 0


def _sample_tex_markers_docx(path, limit=6):
    try:
        with zipfile.ZipFile(path) as z:
            root=etree.fromstring(z.read('word/document.xml'))
        out=[]
        for node in root.xpath(".//*[local-name()='t']"):
            text=node.text or ''
            if ((text.startswith('$') and text.endswith('$') and len(text)>=2) or
                (text.startswith('\\[') and text.endswith('\\]') and len(text)>=4)):
                out.append(text[:180])
                if len(out)>=limit:
                    break
        return out
    except Exception:
        return []


def _run_mathtype_toggle_batch(path, batch_limit=8, timeout_s=300):
    """Convert a bounded number of TeX markers in a fresh Word/MathType session."""
    full=str(Path(path).resolve())
    vb_path=full.replace('"','""')
    script = r'''On Error Resume Next
Dim word, doc, txt, pEndDollar, pStartDollar, pEndDisp, pStartDisp
Dim useDisp, rs, re, rr, ok, beforeText, doneCount, failCount

Function ToggleCurrentRange(rng)
  Dim localOk, localErr, i, probe
  localOk = False
  localErr = ""
  rng.Select
  beforeText = rng.Text
  Err.Clear
  word.Run "MTCommand_TeXToggle"
  If Err.Number = 0 Then
    localOk = True
  Else
    localErr = Err.Description
    Err.Clear
    word.Run "MathTypeCommands.UILib.MTCommand_TeXToggle"
    If Err.Number = 0 Then
      localOk = True
    Else
      localErr = localErr & " | " & Err.Description
    End If
  End If
  If Not localOk Then
    WScript.Echo "ERROR: MathType Toggle TeX khong chay: " & localErr
    ToggleCurrentRange = False
    Exit Function
  End If

  For i = 1 To 100
    WScript.Sleep 100
    probe = rng.Text
    If probe <> beforeText Then
      WScript.Sleep 180
      ToggleCurrentRange = True
      Exit Function
    End If
  Next
  ToggleCurrentRange = False
End Function

Err.Clear
Set word = CreateObject("Word.Application")
If Err.Number <> 0 Then
  WScript.Echo "ERROR: Khong mo duoc Microsoft Word COM: " & Err.Description
  WScript.Quit 10
End If
word.Visible = False
word.DisplayAlerts = 0
word.ScreenUpdating = False
Err.Clear
Set doc = word.Documents.Open("__DOCX_PATH__", False, False, False)
If Err.Number <> 0 Then
  WScript.Echo "ERROR: Khong mo duoc DOCX: " & Err.Description
  word.Quit 0
  WScript.Quit 11
End If

doneCount = 0
failCount = 0
Do While doneCount < __BATCH_LIMIT__
  txt = doc.Content.Text

  pEndDisp = InStrRev(txt, "\]")
  pStartDisp = 0
  If pEndDisp > 0 Then pStartDisp = InStrRev(txt, "\[", pEndDisp - 1)

  pEndDollar = InStrRev(txt, "$")
  pStartDollar = 0
  If pEndDollar > 0 Then pStartDollar = InStrRev(txt, "$", pEndDollar - 1)

  If pStartDisp = 0 And pStartDollar = 0 Then Exit Do
  useDisp = False
  If pStartDisp > 0 And pStartDollar > 0 Then
    If pEndDisp > pEndDollar Then useDisp = True
  ElseIf pStartDisp > 0 Then
    useDisp = True
  End If

  If useDisp Then
    rs = pStartDisp - 1
    re = pEndDisp + 1
  Else
    rs = pStartDollar - 1
    re = pEndDollar
  End If

  Set rr = doc.Range(rs, re)
  ok = ToggleCurrentRange(rr)
  If Not ok Then
    failCount = failCount + 1
    Exit Do
  End If
  doneCount = doneCount + 1
Loop

Err.Clear
doc.Save
If Err.Number <> 0 Then
  WScript.Echo "ERROR: Khong luu duoc Word sau batch: " & Err.Description
  doc.Close False
  word.Quit 0
  WScript.Quit 13
End If
WScript.Sleep 350
doc.Close False
word.NormalTemplate.Saved = True
word.Quit 0
Set doc = Nothing
Set word = Nothing
WScript.Sleep 650
WScript.Echo "OK batch-toggle=" & doneCount & " failed=" & failCount
WScript.Quit 0
'''.replace('__DOCX_PATH__',vb_path).replace('__BATCH_LIMIT__',str(max(1,int(batch_limit))))

    tmp=None
    try:
        fd,tmp_name=tempfile.mkstemp(prefix='mathtype_batch_',suffix='.vbs')
        os.close(fd)
        tmp=Path(tmp_name)
        tmp.write_text(script,encoding='utf-16')
        cscript=os.environ.get('CSCRIPT_PATH') or 'cscript.exe'
        proc=subprocess.run(
            [cscript,'//nologo',str(tmp)],
            stdout=subprocess.PIPE,stderr=subprocess.PIPE,
            text=True,encoding='utf-8',errors='replace',timeout=timeout_s,
        )
        output=((proc.stdout or '')+'\n'+(proc.stderr or '')).strip()
        if proc.returncode!=0:
            raise RuntimeError(output[-1800:] if output else f'Mã lỗi {proc.returncode}.')
        return output
    finally:
        try:
            if tmp and tmp.exists(): tmp.unlink()
        except Exception:
            pass


def _convert_tex_docx_to_real_mathtype_windows(path, require=True):
    """Convert PDF-derived TeX to real Equation.DSMT4 without window buildup."""
    expected=_count_tex_markers_docx(path)
    initial_ole=_count_mathtype_ole_docx(path)
    if expected<=0:
        if initial_ole>0:
            return True
        if require:
            raise RuntimeError('Word chưa có công thức TeX để chuyển sang MathType.')
        return False

    if os.name!='nt':
        if require:
            raise RuntimeError('Nguồn PDF muốn xuất Word MathType thật cần Windows có Microsoft Word + MathType 7.')
        return False

    last_output=''
    stagnant=0
    # One conversion at a time per web process. A double click must not launch
    # two Word/MathType COM jobs concurrently.
    with _MATHTYPE_EXPORT_LOCK:
        remaining=_count_tex_markers_docx(path)
        batches=0
        max_batches=max(8,(remaining+7)//8+8)
        while remaining>0 and batches<max_batches:
            batches+=1
            before=remaining
            try:
                last_output=_run_mathtype_toggle_batch(path,batch_limit=8,timeout_s=300)
            except FileNotFoundError as exc:
                raise RuntimeError(
                    'Không tìm thấy cscript.exe. Không cần pywin32; chỉ cần Windows Script Host + Word + MathType 7.'
                ) from exc
            except subprocess.TimeoutExpired as exc:
                raise RuntimeError(
                    'Word/MathType xử lý một batch quá lâu. Hãy đóng hộp thoại Word/MathType đang mở rồi thử lại.'
                ) from exc
            time.sleep(0.45)
            remaining=_count_tex_markers_docx(path)
            if remaining>=before:
                stagnant+=1
            else:
                stagnant=0
            if stagnant>=2:
                sample=' | '.join(_sample_tex_markers_docx(path,4))
                raise RuntimeError(
                    f'MathType không tiến thêm sau {batches} batch; còn {remaining} công thức TeX. '
                    f'Mẫu còn lại: {sample or "(không đọc được)"}. Log: {last_output[-600:]}'
                )

    ole=_count_mathtype_ole_docx(path)
    remaining=_count_tex_markers_docx(path)
    required_ole=initial_ole+expected
    if remaining or ole<required_ole:
        sample=' | '.join(_sample_tex_markers_docx(path,4))
        raise RuntimeError(
            f'MathType chưa chuyển đủ: TeX ban đầu {expected}, Equation.DSMT4 cần ít nhất {required_ole}, '
            f'hiện có {ole}, TeX còn lại {remaining}. Mẫu còn lại: {sample or "(không có)"}. '
            f'Log batch cuối: {last_output[-500:]}'
        )
    return True

def build_shuffle_docx(source,version,questions,path,plain_latex=False):
    doc=Document(); sec=doc.sections[0]
    sec.top_margin=Cm(1.5); sec.bottom_margin=Cm(1.5); sec.left_margin=Cm(1.6); sec.right_margin=Cm(1.6)
    doc.styles["Normal"].font.name="Times New Roman"; doc.styles["Normal"].font.size=Pt(13.5)
    t=doc.add_table(rows=1,cols=2); t.style="Table Grid"; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    left=((source["department"] or "")+"\n"+(source["school"] or "")).strip()
    right=(source["exam_name"] or source["title"] or "ĐỀ KIỂM TRA")
    if source["school_year"]: right+="\nNĂM HỌC "+source["school_year"]
    if source["subject"]: right+="\nMÔN: "+source["subject"]
    right+="\nThời gian: "+(source["duration_text"] or "90 phút")
    _cell(t.cell(0,0),left,True,11); _cell(t.cell(0,1),right,True,11)
    _addp(doc,f"MÃ ĐỀ: {version['exam_code']}",True,WD_ALIGN_PARAGRAPH.RIGHT,12)

    labels={1:"PHẦN I. TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN.",
            2:"PHẦN II. TRẮC NGHIỆM ĐÚNG SAI.",
            3:"PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN.",
            4:"PHẦN IV. TỰ LUẬN."}
    qmap={q["id"]:q for q in questions}; last=0
    for item in version["entries"]:
        q=qmap[item["question_id"]]
        if item["section"]!=last:
            last=item["section"]; _addp(doc,labels[last],True,size=13.5)
        _add_blocks_docx(doc,q["blocks"],f"Câu {item['number']}: ",latex_for_mathtype=not plain_latex,plain_latex=plain_latex)
        omap={o["key"]:o for o in q["options"]}
        for oi,key in enumerate(item["option_order"]):
            if key in omap:
                lab=("abcd"[oi] if item["section"]==2 else "ABCD"[oi])+". "
                _add_blocks_docx(doc,omap[key]["blocks"],lab,latex_for_mathtype=not plain_latex,plain_latex=plain_latex)
        _write_solution_docx(doc,q,latex_for_mathtype=not plain_latex,plain_latex=plain_latex)

    doc.add_page_break()
    _addp(doc,"ĐÁP ÁN",True,WD_ALIGN_PARAGRAPH.CENTER,14)
    _addp(doc,f"ĐỀ {version['exam_code']}.",True,size=12)
    for sn in (1,2,3):
        items=[x for x in version["entries"] if x["section"]==sn]
        if not items: continue
        _addp(doc,labels[sn],True,size=11)
        if sn in (1,3):
            tt=doc.add_table(rows=2,cols=len(items)+1); tt.style="Table Grid"
            _cell(tt.cell(0,0),"CÂU",True); _cell(tt.cell(1,0),"ĐÁP ÁN",True)
            for j,x in enumerate(items,1):
                _cell(tt.cell(0,j),x["number"],True); _cell(tt.cell(1,j),x["answer"])
        else:
            tt=doc.add_table(rows=5,cols=len(items)+1); tt.style="Table Grid"
            _cell(tt.cell(0,0),"CÂU",True)
            for j,x in enumerate(items,1): _cell(tt.cell(0,j),x["number"],True)
            for ri,lab in enumerate("abcd",1):
                _cell(tt.cell(ri,0),lab,True)
                for j,x in enumerate(items,1):
                    ans=x["answer"]; _cell(tt.cell(ri,j),ans[ri-1] if ri-1<len(ans) else "")
    # Chỉ dựng DOCX có TeX markers. Việc TeX -> MathType phải do caller
    # thực hiện đúng MỘT lần; Toggle TeX hai lần sẽ quay Equation.DSMT4 về TeX.
    doc.save(path)



# ============================================================
# V5.20.0 — XUẤT WORD GIỮ NGUYÊN MATHTYPE/OLE VÀ VỊ TRÍ HÌNH
# ------------------------------------------------------------
# QUAN TRỌNG:
# - Với nguồn DOCX: KHÔNG chuyển MathType thành ảnh khi xuất Word.
# - Mỗi câu được clone trực tiếp từ document.xml của file Word gốc.
# - w:object / o:OLEObject / w:drawing được giữ nguyên.
# - Vì cả câu được clone như một khối XML, hình vẽ giữ đúng vị trí
#   tương đối trong câu như file giáo viên upload.
# - Phần I vẫn đảo A/B/C/D bằng cách đổi thứ tự các XML segment
#   của chính paragraph phương án, không raster hóa MathType.
# - Với nguồn PDF: PDF vốn không chứa MathType editable, vì vậy
#   chỉ có thể giữ hình thức hiển thị bằng parser PDF cũ.
# ============================================================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
O_NS = "urn:schemas-microsoft-com:office:office"
V_NS = "urn:schemas-microsoft-com:vml"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

def _xml_para_text(el):
    # Không dùng prefix ``w:`` vì một số phần tử XML clone từ DOCX cũ
    # không còn namespace map gắn trực tiếp trên node. local-name() hoạt
    # động ổn định với cả python-docx BaseOxmlElement và lxml Element.
    return "".join(el.xpath(".//*[local-name()='t']/text()")).strip()

def _word_section_no(text):
    """Nhận tiêu đề phần Word ở cả hai kiểu: Phần I/II/III/IV và Phần 1/2/3/4.

    File Word thực tế thường có ký hiệu bullet/☞ đứng trước chữ ``Phần``.
    Hàm xuất Word bảo toàn MathType phải dùng cùng quy tắc với parser upload;
    nếu không, DOCX sẽ rơi sang exporter ảnh và làm mất OLE MathType.
    """
    t = normalize_text(text or "")
    # Bỏ ký hiệu trang trí ở đầu nhưng không làm thay đổi nội dung phía sau.
    t = re.sub(r"^[^A-Za-zÀ-ỹ0-9]*", "", t).strip()
    m = re.match(r"^Phần\s*(IV|III|II|I|[1-4])\b", t, re.I)
    if not m:
        return None
    token = m.group(1).upper()
    return {"I":1,"II":2,"III":3,"IV":4,"1":1,"2":2,"3":3,"4":4}[token]

def _is_answer_para(text):
    return bool(re.match(r"^\s*(Đáp\s*án|Đáp\s*số|Trả\s*lời)\s*:", text or "", re.I))

def _split_original_word_questions(doc):
    """
    Trả về:
      headings[section] = XML heading
      chunks[section] = list câu, mỗi câu là list XML element theo đúng thứ tự gốc.
    """
    body = doc._element.body
    children = [copy.deepcopy(x) for x in list(body) if x.tag != qn("w:sectPr")]

    headings = {}
    chunks = {1:[],2:[],3:[],4:[]}
    sec = None
    current = None

    for el in children:
        text = _xml_para_text(el)
        detected = _word_section_no(text)

        if detected:
            if current is not None and sec:
                chunks[sec].append(current)
                current = None
            sec = detected
            headings[sec] = copy.deepcopy(el)
            continue

        if sec is None:
            continue

        if _is_answer_para(text):
            # Đáp án được đưa xuống trang cuối, không chèn trong thân đề.
            continue

        if re.match(r"^\s*Câu\s+\d+\s*[:.]?", text, re.I):
            if current is not None:
                chunks[sec].append(current)
            current = [copy.deepcopy(el)]
        else:
            if current is None:
                # Phần IV trong một số file không ghi "Câu 1":
                # vẫn giữ nội dung thành một khối câu.
                current = [copy.deepcopy(el)]
            else:
                current.append(copy.deepcopy(el))

    if current is not None and sec:
        chunks[sec].append(current)

    # Loại khối rỗng thuần túy.
    for sn in (1,2,3,4):
        chunks[sn] = [
            ch for ch in chunks[sn]
            if any(_xml_para_text(x) or len(x) for x in ch)
        ]

    return headings, chunks

def _paragraph_option_labels(p):
    """
    Tìm vị trí child bắt đầu A/B/C/D trong một paragraph.
    Hỗ trợ dạng Word phổ biến:
      A. <MathType> . B. <MathType> . C. ... D. ...
    """
    if etree.QName(p).localname != "p":
        return []

    labels = []
    for idx, child in enumerate(list(p)):
        if etree.QName(child).localname == "pPr":
            continue
        txt = "".join(child.xpath(".//*[local-name()='t']/text()"))
        m = re.match(r"^\s*([ABCD])\s*[.)]", txt, re.I)
        if m:
            labels.append((idx, m.group(1).upper()))
    return labels

def _replace_first_label(segment, new_label):
    for child in segment:
        ts = child.xpath(".//*[local-name()='t']")
        for t in ts:
            old = t.text or ""
            if re.match(r"^\s*[ABCD]\s*[.)]", old, re.I):
                t.text = re.sub(
                    r"^(\s*)[ABCD](\s*[.)])",
                    lambda m: m.group(1) + new_label + m.group(2),
                    old,
                    count=1,
                    flags=re.I
                )
                return

def _reorder_options_inside_paragraph(p, option_order):
    labels = _paragraph_option_labels(p)
    if len(labels) != 4:
        return False

    children = list(p)
    ppr = children[0] if children and etree.QName(children[0]).localname=="pPr" else None

    prefix_end = labels[0][0]
    prefix = [copy.deepcopy(x) for x in children[:prefix_end] if x is not ppr]

    segs = {}
    for j,(start,key) in enumerate(labels):
        stop = labels[j+1][0] if j+1 < len(labels) else len(children)
        segs[key] = [copy.deepcopy(x) for x in children[start:stop]]

    for child in list(p):
        if child is not ppr:
            p.remove(child)

    for x in prefix:
        p.append(x)

    for display_idx, old_key in enumerate(option_order):
        segment = [copy.deepcopy(x) for x in segs.get(old_key,[])]
        if not segment:
            continue
        _replace_first_label(segment, "ABCD"[display_idx])
        for x in segment:
            p.append(x)

    return True

def _reorder_separate_option_paragraphs(chunk, option_order):
    """
    Nếu A/B/C/D là 4 paragraph riêng, đổi thứ tự 4 paragraph đó.
    """
    found = {}
    idxs = []

    for idx, el in enumerate(chunk):
        text = _xml_para_text(el)
        m = re.match(r"^\s*([ABCD])\s*[.)]", text, re.I)
        if m:
            key = m.group(1).upper()
            if key not in found:
                found[key] = el
                idxs.append(idx)

    if len(found) != 4:
        return chunk

    first = min(idxs)
    last = max(idxs)
    # Chỉ xử lý khi 4 paragraph option nằm trong một vùng liên tiếp/không xa.
    if last-first > 7:
        return chunk

    rebuilt = []
    for idx,el in enumerate(chunk):
        if idx == first:
            for display_idx,old_key in enumerate(option_order):
                cp = copy.deepcopy(found[old_key])
                # đổi label hiển thị
                for t in cp.xpath(".//*[local-name()='t']"):
                    old=t.text or ""
                    if re.match(r"^\s*[ABCD]\s*[.)]",old,re.I):
                        t.text=re.sub(
                            r"^(\s*)[ABCD](\s*[.)])",
                            lambda m:m.group(1)+"ABCD"[display_idx]+m.group(2),
                            old,count=1,flags=re.I
                        )
                        break
                rebuilt.append(cp)
        if idx not in idxs:
            rebuilt.append(copy.deepcopy(el))
    return rebuilt

def _apply_section1_option_shuffle_to_chunk(chunk, option_order):
    chunk = [copy.deepcopy(x) for x in chunk]

    # Trường hợp cả 4 phương án trong cùng paragraph.
    for el in chunk:
        if _reorder_options_inside_paragraph(el, option_order):
            return chunk

    # Trường hợp mỗi phương án là paragraph riêng.
    return _reorder_separate_option_paragraphs(chunk, option_order)



def _preserve_original_mathtype_and_text(elements):
    """
    =================================================================
    KHỐI BẢO TOÀN ĐỊNH DẠNG WORD / MATHTYPE — V5.20.7
    =================================================================

    Dựa trên nguyên tắc:
    - Không đọc text rồi dựng lại câu bằng python-docx.
    - Không ép font/kích thước lại trên các run của câu.
    - Không scale VML shape của MathType.
    - Không raster hóa OLE MathType.
    - Không thay đổi ảnh công thức nhỏ / hình vẽ.

    Các element truyền vào đây là bản clone trực tiếp từ OOXML của file
    Word giáo viên upload. Hàm này CỐ Ý không thay đổi chúng.

    Vì vậy:
    - chữ trong câu giữ đúng font/cỡ như file Word gốc;
    - MathType giữ đúng Define Sizes của file gốc
      (ví dụ Full 12pt, Script 7pt, Symbol 18pt...);
    - hình vẽ giữ đúng kích thước/vị trí tương đối của câu.

    Lưu ý:
    File đã đảo không thể byte-for-byte giống file gốc vì thứ tự câu thay đổi.
    Nhưng nội dung từng câu/OLE/drawing được giữ nguyên, chỉ thay số câu
    và nhãn phương án khi cần.
    =================================================================
    """
    return elements


def _normalize_text_runs_135(elements):
    """
    Chữ thường -> Times New Roman 13.5pt.
    MathType/OLE, ảnh công thức và hình vẽ giữ nguyên đối tượng gốc.
    """
    for el in elements:
        for run in el.xpath(".//*[local-name()='r']"):
            if (
                run.xpath(".//*[local-name()='object']")
                or run.xpath(".//*[local-name()='drawing']")
                or run.xpath(".//*[local-name()='pict']")
            ):
                continue

            rPr=run.find(qn("w:rPr"))
            if rPr is None:
                rPr=OxmlElement("w:rPr")
                run.insert(0,rPr)

            rFonts=rPr.find(qn("w:rFonts"))
            if rFonts is None:
                rFonts=OxmlElement("w:rFonts")
                rPr.insert(0,rFonts)
            for a in ("ascii","hAnsi","eastAsia","cs"):
                rFonts.set(qn("w:"+a),"Times New Roman")

            sz=rPr.find(qn("w:sz"))
            if sz is None:
                sz=OxmlElement("w:sz")
                rPr.append(sz)
            sz.set(qn("w:val"),"27")

            szCs=rPr.find(qn("w:szCs"))
            if szCs is None:
                szCs=OxmlElement("w:szCs")
                rPr.append(szCs)
            szCs.set(qn("w:val"),"27")


def _uniquify_drawing_ids(elements, seed):
    """
    Đổi ID hiển thị khi clone câu nhưng PHẢI giữ liên kết VML/OLE đồng bộ.

    Word/MathType thường có cặp:
        <v:shape id="_x0000_s1026"> ...
        <o:OLEObject ShapeID="_x0000_s1026" .../>
    V5.24.7 đổi ``v:shape/@id`` mà chưa đổi ``OLEObject/@ShapeID``. Một số
    MathType/hình vì vậy có thể dùng preview sai hoặc lệch sau khi đảo.

    V5.24.8 tạo bảng ánh xạ old_shape_id -> new_shape_id rồi cập nhật mọi
    thuộc tính ShapeID tương ứng. Relationship rId, kích thước, crop, anchor,
    wrap và bytes ảnh/OLE vẫn giữ nguyên từ DOCX giáo viên upload.
    """
    counter = seed * 10000
    shape_map = {}

    for el in elements:
        for node in el.xpath(".//*[local-name()='docPr']"):
            counter += 1
            node.set("id", str(counter))

        for node in el.xpath(".//*[local-name()='shape']"):
            old_id = node.get("id")
            counter += 1
            new_id = f"_x0000_s{counter}"
            node.set("id", new_id)
            if old_id:
                shape_map[old_id] = new_id

    for el in elements:
        for node in el.xpath(".//*[local-name()='OLEObject']"):
            old_shape = node.get("ShapeID")
            if old_shape in shape_map:
                node.set("ShapeID", shape_map[old_shape])
            counter += 1
            if node.get("ObjectID") is not None:
                node.set("ObjectID", f"_{counter}")

        # Một số Word cũ ghi tham chiếu shape bằng thuộc tính shapeid/ShapeID
        # trên node Office/VML khác. Chỉ thay đúng giá trị đã ánh xạ.
        for node in el.xpath(".//*"):
            for attr_name in list(node.attrib):
                if etree.QName(attr_name).localname.lower() == "shapeid":
                    old_value = node.attrib.get(attr_name)
                    if old_value in shape_map:
                        node.attrib[attr_name] = shape_map[old_value]

def _verify_preserved_word_assets(original_path, output_path):
    """Fail closed if Word export changed/dropped source media or MathType OLE."""
    import hashlib

    def package_hashes(path, prefixes):
        result={}
        with zipfile.ZipFile(path) as archive:
            for name in archive.namelist():
                if any(name.startswith(prefix) for prefix in prefixes) and not name.endswith("/"):
                    result[name]=hashlib.sha256(archive.read(name)).hexdigest()
        return result

    source_assets=package_hashes(original_path,("word/media/","word/embeddings/"))
    output_assets=package_hashes(output_path,("word/media/","word/embeddings/"))
    missing=[name for name,digest in source_assets.items() if output_assets.get(name)!=digest]
    if missing:
        raise RuntimeError(
            "Word sau khi đảo không bảo toàn nguyên bytes hình/MathType: "
            + ", ".join(missing[:8])
        )

    with zipfile.ZipFile(original_path) as archive:
        source_root=etree.fromstring(archive.read("word/document.xml"))
    with zipfile.ZipFile(output_path) as archive:
        root=etree.fromstring(archive.read("word/document.xml"))

    # Nguồn Word MathType thật dùng OLE ProgID Equation.DSMT4. Với file xuất
    # một mã đề, số OLE phải bằng nguồn; với nhiều mã đề phải là bội số tương
    # ứng. Quan trọng nhất: mọi OLE trong file xuất vẫn phải là OLE MathType,
    # không được thay bằng w:drawing/image.
    source_oles=source_root.xpath(".//*[local-name()='OLEObject']")
    output_oles=root.xpath(".//*[local-name()='OLEObject']")
    if source_oles and not output_oles:
        raise RuntimeError("Word sau khi đảo đã mất toàn bộ OLE MathType.")
    bad_prog=[n.get("ProgID") for n in output_oles if n.get("ProgID") not in (None,"Equation.DSMT4")]
    if bad_prog:
        raise RuntimeError("Word sau khi đảo có OLE không còn là MathType Equation.DSMT4.")

    shape_ids={
        node.get("id") for node in root.xpath(".//*[local-name()='shape']")
        if node.get("id")
    }
    broken=[
        node.get("ShapeID") for node in root.xpath(".//*[local-name()='OLEObject']")
        if node.get("ShapeID") and node.get("ShapeID") not in shape_ids
    ]
    if broken:
        raise RuntimeError(
            "Word sau khi đảo có OLE/ShapeID không đồng bộ: "
            + ", ".join(broken[:8])
        )
    return True


def _append_xml_before_sectpr(doc, elements):
    body = doc._element.body
    sectPr = body.sectPr
    for el in elements:
        if sectPr is not None:
            sectPr.addprevious(el)
        else:
            body.append(el)

def _clear_doc_body_keep_sectpr(doc):
    body = doc._element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)

def _source_saved_path(source):
    stored = source["stored_name"] if "stored_name" in source.keys() else ""
    if stored:
        p = UPLOAD_DIR / stored
        if p.exists():
            return p
    return None


# =====================================================================
# CHIẾN LƯỢC HIỂN THỊ WEB — V5.21.0
# ---------------------------------------------------------------------
# A. FILE WORD TẢI VỀ:
#    - giữ nguyên OOXML/OLE MathType như V5.20.7;
#    - không chuyển MathType thành ảnh/LaTeX.
#
# B. HIỂN THỊ TRÊN TRÌNH DUYỆT:
#    - block type="latex" -> template bọc \( ... \) hoặc \[ ... \];
#    - MathJax trong base.html render công thức SVG sắc nét;
#    - CSS `mjx-container { font-size:105% }` cân với TNR 13.5pt;
#    - block type="image" (MathType/PDF chưa chuyển được sang LaTeX)
#      vẫn dùng ảnh gốc + display_height_px riêng từng công thức.
#
# C. HÌNH VẼ:
#    - role="figure" luôn tách hàng;
#    - giữ aspect ratio và đi cùng block của câu khi trộn.
#
# LƯU Ý:
# MathJax không thể tự đọc trực tiếp OLE MathType nhúng trong DOCX.
# Vì vậy chỉ dùng MathJax sau khi công thức đã ở dạng LaTeX/MathML.
# =====================================================================

# =====================================================================
# SƠ ĐỒ CÁC KHỐI CODE XỬ LÝ WORD / MATHTYPE — V5.20.7
# ---------------------------------------------------------------------
# 1. f.save(saved)
#    -> lưu file upload nhị phân nguyên bản vào uploads/.
#
# 2. shuffle_exam_source_download()
#    -> tải lại đúng file gốc đã upload, không dựng lại.
#
# 3. _split_original_word_questions()
#    -> chia document.xml thành từng khối câu nhưng KHÔNG dựng lại text.
#
# 4. _apply_section1_option_shuffle_to_chunk()
#    -> chỉ Phần I: đổi thứ tự XML của A/B/C/D.
#    -> MathType/hình trong từng phương án đi nguyên khối.
#
# 5. _preserve_original_mathtype_and_text()
#    -> CỐ Ý không sửa font, cỡ chữ, VML shape, OLE MathType, drawing.
#    -> Define Sizes của MathType trong Word gốc được giữ nguyên.
#
# 6. _uniquify_drawing_ids()
#    -> chỉ đổi ID hiển thị để tránh trùng ID khi clone nhiều mã đề.
#    -> không thay dữ liệu MathType/image.
#
# 7. _append_xml_before_sectpr()
#    -> chèn nguyên khối OOXML của câu sang vị trí mới.
#
# 8. build_all_shuffle_docx_preserve_word()
#    -> tạo các mã đề + gom đáp án ở cuối file.
#    -> thân câu lấy trực tiếp từ OOXML Word gốc.
#
# 9. _set_tnr(), _addp(), _cell()
#    -> chỉ áp dụng Times New Roman 13.5pt cho phần MỚI do hệ thống tạo
#       như tiêu đề, mã đề, header, bảng đáp án.
# =====================================================================



# =====================================================================
# V5.24.11 — XUẤT WORD TRỰC TIẾP TRÊN PACKAGE OOXML GỐC
# ---------------------------------------------------------------------
# Không dùng python-docx để save lại file nguồn. Chỉ thay word/document.xml
# bằng các XML block clone từ chính document.xml gốc. Nhờ đó:
# - word/embeddings/*.bin (MathType OLE) giữ nguyên byte;
# - word/media/* giữ nguyên byte;
# - drawing/VML style, crop, extent, anchor giữ nguyên;
# - với 1 mã đề, ShapeID/ObjectID giữ nguyên hoàn toàn.
# =====================================================================

def _split_original_word_questions_xml(root):
    body = root.xpath(".//*[local-name()='body']")[0]
    preamble=[]
    headings={}
    chunks={1:[],2:[],3:[],4:[]}
    sec=None
    current=None
    started=False
    sectpr=None

    for raw in list(body):
        if etree.QName(raw).localname == 'sectPr':
            sectpr=copy.deepcopy(raw)
            continue
        el=copy.deepcopy(raw)
        text=_xml_para_text(el)
        detected=_word_section_no(text)
        if detected:
            started=True
            if current is not None and sec:
                chunks[sec].append(current)
                current=None
            sec=detected
            headings[sec]=copy.deepcopy(el)
            continue
        if not started:
            preamble.append(copy.deepcopy(el))
            continue
        if sec is None:
            continue
        if _is_answer_para(text):
            # Giữ đáp án/lời giải nằm trong câu; chỉ bỏ dòng đáp án tổng hợp
            # có dạng rõ ràng "Đáp án:" ở cấp body ngoài câu.
            if current is None:
                continue
        if re.match(r"^\s*Câu\s+\d+\s*[:.]?", text, re.I):
            if current is not None:
                chunks[sec].append(current)
            current=[copy.deepcopy(el)]
        else:
            if current is None:
                current=[copy.deepcopy(el)]
            else:
                current.append(copy.deepcopy(el))
    if current is not None and sec:
        chunks[sec].append(current)
    for sn in (1,2,3,4):
        chunks[sn]=[ch for ch in chunks[sn] if any(_xml_para_text(x) or len(x) for x in ch)]
    return preamble,headings,chunks,sectpr

def _new_page_break_xml():
    p=OxmlElement('w:p')
    r=OxmlElement('w:r')
    br=OxmlElement('w:br')
    br.set(qn('w:type'),'page')
    r.append(br); p.append(r)
    return p

def _exam_code_paragraph_xml(code):
    p=OxmlElement('w:p')
    pPr=OxmlElement('w:pPr')
    jc=OxmlElement('w:jc'); jc.set(qn('w:val'),'right'); pPr.append(jc)
    p.append(pPr)
    r=OxmlElement('w:r')
    rPr=OxmlElement('w:rPr')
    b=OxmlElement('w:b'); rPr.append(b)
    fonts=OxmlElement('w:rFonts')
    for a in ('ascii','hAnsi','eastAsia','cs'):
        fonts.set(qn('w:'+a),'Times New Roman')
    rPr.append(fonts)
    sz=OxmlElement('w:sz'); sz.set(qn('w:val'),'24'); rPr.append(sz)
    r.append(rPr)
    t=OxmlElement('w:t'); t.text=f'MÃ ĐỀ: {code}'; r.append(t); p.append(r)
    return p

def _docx_generated_body_elements(doc):
    """Return body elements (excluding sectPr) from a temporary Document."""
    return [copy.deepcopy(x) for x in list(doc._element.body) if etree.QName(x).localname!='sectPr']


def _metadata_header_xml(source, exam_code):
    """Create the complete teacher-entered header frame without touching source DOCX."""
    d=Document()
    for el in list(d._element.body):
        if etree.QName(el).localname!='sectPr':
            d._element.body.remove(el)
    t=d.add_table(rows=1,cols=2)
    _apply_table_grid(t)
    t.alignment=WD_TABLE_ALIGNMENT.CENTER
    left=((source['department'] or '')+'\n'+(source['school'] or '')).strip()
    right=(source['exam_name'] or source['title'] or 'ĐỀ KIỂM TRA')
    if source['school_year']:
        right+='\nNĂM HỌC '+str(source['school_year'])
    if source['subject']:
        right+='\nMÔN: '+str(source['subject'])
    right+='\nThời gian: '+str(source['duration_text'] or '90 phút')
    _cell(t.cell(0,0),left,True,11)
    _cell(t.cell(0,1),right,True,11)
    _addp(d,f'MÃ ĐỀ: {exam_code}',True,WD_ALIGN_PARAGRAPH.RIGHT,12)
    return _docx_generated_body_elements(d)


def _write_consolidated_answer_matrix_docx(doc, versions):
    """Write one compact answer matrix for all codes at the end of a DOCX."""
    _addp(doc,'BẢNG ĐÁP ÁN CÁC MÃ ĐỀ',True,WD_ALIGN_PARAGRAPH.CENTER,14)
    labels={1:'PHẦN I',2:'PHẦN II',3:'PHẦN III'}
    for sn in (1,2,3):
        by_version=[]; max_no=0
        for v in versions:
            mp={int(x['number']):str(x.get('answer') or '') for x in v.get('entries',[]) if int(x.get('section') or 0)==sn}
            if mp:
                max_no=max(max_no,max(mp))
            by_version.append((v,mp))
        if max_no<=0:
            continue
        _addp(doc,labels[sn],True,size=11)
        tt=doc.add_table(rows=1+len(by_version),cols=1+max_no)
        _apply_table_grid(tt); tt.alignment=WD_TABLE_ALIGNMENT.CENTER
        _cell(tt.cell(0,0),'MÃ ĐỀ',True,8.5)
        for n in range(1,max_no+1):
            _cell(tt.cell(0,n),n,True,8.5)
        for ri,(v,mp) in enumerate(by_version,1):
            _cell(tt.cell(ri,0),v.get('exam_code',''),True,8.5)
            for n in range(1,max_no+1):
                _cell(tt.cell(ri,n),mp.get(n,''),False,8.5)


def _answer_summary_xml(versions):
    """Build one final-page answer matrix containing every generated exam code."""
    d=Document()
    for el in list(d._element.body):
        if etree.QName(el).localname!='sectPr':
            d._element.body.remove(el)
    d._element.body.insert(0,_new_page_break_xml())
    _write_consolidated_answer_matrix_docx(d,versions)
    return _docx_generated_body_elements(d)


def _strip_inline_answer_paragraphs(qchunk):
    """Remove only explicit answer-label lines from question body; keep solutions."""
    out=[]
    for el in qchunk:
        txt=_xml_para_text(el).strip()
        if re.match(r'^(?:Đáp\s*án|Đáp\s*số|Trả\s*lời)\s*:',txt,re.I):
            continue
        out.append(el)
    return out


def _renumber_question_chunk(qchunk, new_number):
    for el in qchunk:
        text=_xml_para_text(el)
        if re.match(r"^\s*Câu\s+\d+\s*[:.]?", text, re.I):
            for t in el.xpath(".//*[local-name()='t']"):
                if t.text and re.search(r"Câu\s+\d+",t.text,re.I):
                    t.text=re.sub(r"Câu\s+\d+",f'Câu {new_number}',t.text,count=1,flags=re.I)
                    return

def _uniquify_clone_ids_preserve_mathtype(elements, serial):
    """Only for duplicate versions in ALL.docx; preserve Word/MathType ID style.

    First version never calls this. Subsequent clones need unique IDs, but V5.24.10
    changed ``_x0000_i...`` to ``_x0000_s...``. Here we preserve the original
    ``i`` prefix and keep OLE ShapeID paired with the cloned VML shape.
    """
    shape_map={}
    base=200000 + serial*10000
    n=0
    for el in elements:
        for node in el.xpath(".//*[local-name()='docPr']"):
            n+=1; node.set('id',str(base+n))
        for node in el.xpath(".//*[local-name()='shape']"):
            old=node.get('id')
            if not old:
                continue
            n+=1
            m=re.match(r'^(.*?)(\d+)$',old)
            prefix=m.group(1) if m else '_x0000_i'
            # MathType source commonly uses _x0000_iNNNN; keep that convention.
            if old.startswith('_x0000_i'):
                prefix='_x0000_i'
            new=f'{prefix}{base+n}'
            node.set('id',new); shape_map[old]=new
    for el in elements:
        for ole in el.xpath(".//*[local-name()='OLEObject']"):
            old=ole.get('ShapeID')
            if old in shape_map:
                ole.set('ShapeID',shape_map[old])
            if ole.get('ObjectID') is not None:
                n+=1
                ole.set('ObjectID',f'_{1900000000 + base + n}')
        for node in el.xpath('.//*'):
            for attr in list(node.attrib):
                if etree.QName(attr).localname.lower()=='shapeid':
                    old=node.attrib.get(attr)
                    if old in shape_map:
                        node.attrib[attr]=shape_map[old]

def _verify_direct_ooxml_export(original_path,output_path,version_count=1):
    import hashlib,collections
    def hashes(path,prefix):
        out={}
        with zipfile.ZipFile(path) as z:
            for name in z.namelist():
                if name.startswith(prefix) and not name.endswith('/'):
                    out[name]=hashlib.sha256(z.read(name)).hexdigest()
        return out
    for prefix in ('word/media/','word/embeddings/'):
        if hashes(original_path,prefix)!=hashes(output_path,prefix):
            raise RuntimeError('DOCX xuất ra đã thay đổi bytes '+prefix+'; hệ thống dừng để bảo vệ MathType/hình gốc.')
    with zipfile.ZipFile(original_path) as z:
        sr=etree.fromstring(z.read('word/document.xml'))
    with zipfile.ZipFile(output_path) as z:
        rr=etree.fromstring(z.read('word/document.xml'))
    soles=sr.xpath(".//*[local-name()='OLEObject' and @ProgID='Equation.DSMT4']")
    roles=rr.xpath(".//*[local-name()='OLEObject' and @ProgID='Equation.DSMT4']")
    if len(roles) != len(soles)*max(1,int(version_count)):
        raise RuntimeError(f'Số MathType OLE sau xuất không đúng: nguồn {len(soles)}, xuất {len(roles)}.')
    # Every OLE r:id must still resolve in document.xml.rels.
    with zipfile.ZipFile(output_path) as z:
        rel=etree.fromstring(z.read('word/_rels/document.xml.rels'))
    rel_ids={x.get('Id') for x in rel}
    rid_attr='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    missing=[o.get(rid_attr) for o in roles if o.get(rid_attr) not in rel_ids]
    if missing:
        raise RuntimeError('MathType OLE bị mất relationship: '+', '.join(map(str,missing[:8])))
    # Visual size metadata must be cloned, not normalized. Compare multisets.
    def shape_styles(root):
        return collections.Counter(n.get('style') or '' for n in root.xpath(".//*[local-name()='shape']"))
    src_styles=shape_styles(sr); out_styles=shape_styles(rr)
    for style,count in src_styles.items():
        if out_styles[style] < count*max(1,int(version_count)):
            raise RuntimeError('Kích thước VML/hình MathType không được bảo toàn: '+style)

    # Hình minh họa Câu 1/3/4 của file mẫu là DrawingML inline, kích thước
    # nằm ở wp:extent chứ không nằm trong VML style. So sánh multiset extent
    # để file Word tải về không thể âm thầm co hình dù bytes media vẫn đúng.
    def drawing_extents(root):
        return collections.Counter(
            (n.get('cx') or '', n.get('cy') or '')
            for n in root.xpath(".//*[local-name()='inline' or local-name()='anchor']/*[local-name()='extent']")
        )
    src_ext=drawing_extents(sr); out_ext=drawing_extents(rr)
    for extent,count in src_ext.items():
        if out_ext[extent] < count*max(1,int(version_count)):
            raise RuntimeError(
                'Kích thước DrawingML không được bảo toàn: '+str(extent)
            )
    return True

def build_all_shuffle_docx_direct_ooxml(source,versions,questions,output_path):
    original=_source_saved_path(source)
    if not original or original.suffix.lower()!='.docx':
        return False
    with zipfile.ZipFile(original,'r') as zin:
        original_xml=zin.read('word/document.xml')
    root=etree.fromstring(original_xml)
    preamble,headings,chunks,sectpr=_split_original_word_questions_xml(root)
    expected={1:0,2:0,3:0,4:0}
    for q in questions:
        expected[int(q['section'])]+=1
    for sn in (1,2,3,4):
        if expected[sn] and len(chunks.get(sn,[])) < expected[sn]:
            raise RuntimeError(f'Word gốc chỉ tách được {len(chunks.get(sn,[]))}/{expected[sn]} câu ở Phần {sn}.')
    qmap={q['id']:q for q in questions}
    body=root.xpath(".//*[local-name()='body']")[0]
    for c in list(body):
        body.remove(c)
    has_teacher_header=any(str(source[k] or '').strip() for k in ('department','school','exam_name','subject','school_year','duration_text'))
    for vi,version in enumerate(versions):
        if vi:
            body.append(_new_page_break_xml())
        if has_teacher_header:
            # Khung Sở/Trường/Kỳ thi/Môn/Mã đề được dựng từ metadata giáo viên
            # rồi chèn trực tiếp vào document.xml; source DOCX không bị save lại.
            for el in _metadata_header_xml(source,version['exam_code']):
                body.append(el)
        else:
            # Nếu giáo viên không nhập metadata, giữ nguyên phần đầu Word gốc.
            for el in preamble:
                cp=copy.deepcopy(el)
                if vi:
                    _uniquify_clone_ids_preserve_mathtype([cp],vi*1000)
                body.append(cp)
            body.append(_exam_code_paragraph_xml(version['exam_code']))
        last=0
        for item in version['entries']:
            sn=int(item['section']); q=qmap[item['question_id']]
            if sn!=last:
                last=sn
                if headings.get(sn) is not None:
                    body.append(copy.deepcopy(headings[sn]))
            idx=max(0,int(q['section_order'])-1)
            if idx>=len(chunks[sn]):
                raise RuntimeError(f'Không ánh xạ được Phần {sn} - Câu {q["section_order"]} về Word gốc.')
            qchunk=[copy.deepcopy(x) for x in chunks[sn][idx]]
            _renumber_question_chunk(qchunk,item['number'])
            if sn==1:
                qchunk=_apply_section1_option_shuffle_to_chunk(qchunk,item.get('option_order') or ['A','B','C','D'])
            # First version: preserve every original ShapeID/ObjectID verbatim.
            # Later versions: only uniqueness changes, not sizing/anchor/media/OLE bytes.
            if vi:
                _uniquify_clone_ids_preserve_mathtype(qchunk,vi*1000+int(item['number']))
            for el in qchunk:
                body.append(el)
    # Trang cuối cùng: bảng đáp án tổng hợp của toàn bộ mã đề.
    for el in _answer_summary_xml(versions):
        body.append(el)
    if sectpr is not None:
        body.append(copy.deepcopy(sectpr))
    new_xml=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone='yes')
    output_path=Path(output_path); output_path.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(original,'r') as zin, zipfile.ZipFile(output_path,'w') as zout:
        for info in zin.infolist():
            data=new_xml if info.filename=='word/document.xml' else zin.read(info.filename)
            zout.writestr(info,data)
    _verify_direct_ooxml_export(original,output_path,len(versions))
    return True

def build_all_shuffle_docx_preserve_word(source, versions, questions, output_path):
    """
    Dùng file DOCX upload làm package nền để giữ nguyên:
      - MathType OLE object,
      - ảnh công thức nhỏ,
      - hình vẽ,
      - relationship media/embeddings.
    """
    original = _source_saved_path(source)
    if not original or original.suffix.lower() != ".docx":
        return False

    doc = Document(str(original))
    headings, chunks = _split_original_word_questions(doc)

    # Nếu không nhận đủ chunk theo parser, không mạo hiểm làm sai:
    # fallback về exporter cũ.
    counts_expected = {1:0,2:0,3:0,4:0}
    for q in questions:
        counts_expected[q["section"]] += 1

    for sec in (1,2,3):
        if counts_expected[sec] and len(chunks.get(sec,[])) < counts_expected[sec]:
            return False

    _clear_doc_body_keep_sectpr(doc)

    sec = doc.sections[0]
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(1.6)
    sec.right_margin = Cm(1.6)

    # KHÔNG đổi style Normal của file nguồn:
    # nếu đổi ở đây, các run kế thừa style có thể bị thay font/cỡ.
    # Header/đáp án mới dùng _set_tnr() riêng.
    qmap = {q["id"]:q for q in questions}

    for vi,version in enumerate(versions):
        if vi > 0:
            doc.add_page_break()

        # Header thông tin đề.
        table = doc.add_table(rows=1, cols=2)
        _apply_table_grid(table)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        left = ((source["department"] or "") + "\n" + (source["school"] or "")).strip()
        right = source["exam_name"] or source["title"] or "ĐỀ KIỂM TRA"
        if source["school_year"]:
            right += "\nNĂM HỌC " + source["school_year"]
        if source["subject"]:
            right += "\nMÔN: " + source["subject"]
        right += "\nThời gian: " + (source["duration_text"] or "90 phút")

        _cell(table.cell(0,0), left, True, 11)
        _cell(table.cell(0,1), right, True, 11)
        _addp(doc, f"MÃ ĐỀ: {version['exam_code']}", True, WD_ALIGN_PARAGRAPH.RIGHT, 12)

        last_sec = 0

        for item in version["entries"]:
            secno = item["section"]
            q = qmap[item["question_id"]]

            if secno != last_sec:
                last_sec = secno
                heading = headings.get(secno)
                if heading is not None:
                    _append_xml_before_sectpr(doc, [copy.deepcopy(heading)])
                else:
                    labels={
                        1:"PHẦN I. TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN.",
                        2:"PHẦN II. TRẮC NGHIỆM ĐÚNG SAI.",
                        3:"PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN.",
                        4:"PHẦN IV. TỰ LUẬN."
                    }
                    _addp(doc, labels[secno], True, size=13.5)

            source_idx = max(0, int(q["section_order"]) - 1)
            sec_chunks = chunks.get(secno,[])
            if source_idx >= len(sec_chunks):
                raise RuntimeError(
                    f"Không ánh xạ được Phần {secno} - Câu {q['section_order']} "
                    "về Word gốc. Hệ thống dừng xuất để tránh làm mất MathType/hình."
                )

            qchunk = [copy.deepcopy(x) for x in sec_chunks[source_idx]]

            # Câu mới phải đánh lại số sau khi đảo.
            for el in qchunk:
                text = _xml_para_text(el)
                if re.match(r"^\s*Câu\s+\d+\s*[:.]?", text, re.I):
                    for t in el.xpath(".//*[local-name()='t']"):
                        if t.text and re.search(r"Câu\s+\d+", t.text, re.I):
                            t.text = re.sub(
                                r"Câu\s+\d+",
                                f"Câu {item['number']}",
                                t.text,
                                count=1,
                                flags=re.I
                            )
                            break
                    break

            if secno == 1:
                qchunk = _apply_section1_option_shuffle_to_chunk(
                    qchunk,
                    item["option_order"]
                )

            # =========================================================
            # GIỮ NGUYÊN FONT + MATHTYPE + HÌNH CỦA CÂU WORD GỐC
            # Không ép 13.5pt vào thân câu, không scale MathType.
            # =========================================================
            _preserve_original_mathtype_and_text(qchunk)

            # Chỉ đổi ID hiển thị để tránh trùng khi clone nhiều mã đề.
            # Relationship tới OLE/image gốc vẫn được giữ.
            _uniquify_drawing_ids(qchunk, (vi+1)*100 + item["number"])
            _append_xml_before_sectpr(doc, qchunk)

    # ĐÁP ÁN CUỐI FILE — gồm tất cả mã đề.
    doc.add_page_break()
    _addp(doc, "ĐÁP ÁN TẤT CẢ MÃ ĐỀ", True, WD_ALIGN_PARAGRAPH.CENTER, 14)
    for vi,version in enumerate(versions):
        if vi > 0:
            doc.add_page_break()
        _write_answer_key_docx(doc, source, version)

    doc.save(str(output_path))
    _verify_preserved_word_assets(original, output_path)
    return True


def _write_exam_body_docx(doc, source, version, questions, latex_for_mathtype=False, plain_latex=False):
    labels={
        1:"PHẦN I. TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN.",
        2:"PHẦN II. TRẮC NGHIỆM ĐÚNG SAI.",
        3:"PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN.",
        4:"PHẦN IV. TỰ LUẬN."
    }
    qmap={q["id"]:q for q in questions}

    t=doc.add_table(rows=1,cols=2)
    _apply_table_grid(t)
    t.alignment=WD_TABLE_ALIGNMENT.CENTER

    left=((source["department"] or "")+"\n"+(source["school"] or "")).strip()
    right=(source["exam_name"] or source["title"] or "ĐỀ KIỂM TRA")
    if source["school_year"]:
        right+="\nNĂM HỌC "+source["school_year"]
    if source["subject"]:
        right+="\nMÔN: "+source["subject"]
    right+="\nThời gian: "+(source["duration_text"] or "90 phút")

    _cell(t.cell(0,0),left,True,11)
    _cell(t.cell(0,1),right,True,11)
    _addp(doc,f"MÃ ĐỀ: {version['exam_code']}",True,WD_ALIGN_PARAGRAPH.RIGHT,12)

    last=0
    for item in version["entries"]:
        q=qmap[item["question_id"]]
        if item["section"]!=last:
            last=item["section"]
            _addp(doc,labels[last],True,size=13.5)

        _add_blocks_docx(doc,q["blocks"],f"Câu {item['number']}: ",latex_for_mathtype=latex_for_mathtype,plain_latex=plain_latex)

        omap={o["key"]:o for o in q["options"]}
        for oi,key in enumerate(item["option_order"]):
            if key in omap:
                lab=("abcd"[oi] if item["section"]==2 else "ABCD"[oi])+". "
                _add_blocks_docx(doc,omap[key]["blocks"],lab,latex_for_mathtype=latex_for_mathtype,plain_latex=plain_latex)
        _write_solution_docx(doc,q,latex_for_mathtype=latex_for_mathtype,plain_latex=plain_latex)


def _write_answer_key_docx(doc, source, version):
    labels={
        1:"PHẦN I. TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN.",
        2:"PHẦN II. TRẮC NGHIỆM ĐÚNG SAI.",
        3:"PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN."
    }

    _addp(doc,"ĐÁP ÁN",True,WD_ALIGN_PARAGRAPH.CENTER,14)
    _addp(doc,f"ĐỀ {version['exam_code']}.",True,size=12)

    for sn in (1,2,3):
        items=[x for x in version["entries"] if x["section"]==sn]
        if not items:
            continue

        _addp(doc,labels[sn],True,size=11)

        if sn in (1,3):
            tt=doc.add_table(rows=2,cols=len(items)+1)
            _apply_table_grid(tt)
            tt.alignment=WD_TABLE_ALIGNMENT.CENTER
            _cell(tt.cell(0,0),"CÂU",True)
            _cell(tt.cell(1,0),"ĐÁP ÁN",True)
            for j,x in enumerate(items,1):
                _cell(tt.cell(0,j),x["number"],True)
                _cell(tt.cell(1,j),x["answer"])
        else:
            tt=doc.add_table(rows=5,cols=len(items)+1)
            _apply_table_grid(tt)
            tt.alignment=WD_TABLE_ALIGNMENT.CENTER
            _cell(tt.cell(0,0),"CÂU",True)
            for j,x in enumerate(items,1):
                _cell(tt.cell(0,j),x["number"],True)
            for ri,lab in enumerate("abcd",1):
                _cell(tt.cell(ri,0),lab,True)
                for j,x in enumerate(items,1):
                    ans=x["answer"]
                    _cell(tt.cell(ri,j),ans[ri-1] if ri-1<len(ans) else "")



def _load_all_shuffle_export_data(source_id, user_id):
    with db() as c:
        ensure_shuffle_schema(c)
        source=c.execute(
            "SELECT * FROM shuffle_sources WHERE id=? AND uploaded_by=?",
            (source_id,user_id)
        ).fetchone()
        if not source:
            return None,None,None

        rows=c.execute(
            "SELECT * FROM shuffle_versions WHERE source_id=? ORDER BY id",
            (source_id,)
        ).fetchall()
        questions=load_shuffle_questions(c,source_id)

    versions=[]
    for row in rows:
        payload=json.loads(row["version_json"])
        versions.append({
            "exam_code":str(payload.get("exam_code","")),
            "entries":payload.get("entries",[])
        })
    return source,versions,questions

def _shuffle_export_dir():
    """Writable, lock-resistant work directory for generated shuffle files."""
    configured=(os.environ.get("SHUFFLE_EXPORT_DIR") or "").strip()
    if configured:
        export_dir=Path(configured).expanduser()
    elif os.name=="nt":
        # Do not work in D:\...\project\exports on Windows. Word/MathType,
        # antivirus or Explorer preview can keep that path locked.
        export_dir=Path(tempfile.gettempdir())/"QuanLyThiTrucTuyen_exports"
    else:
        export_dir=BASE/"exports"
    export_dir.mkdir(parents=True,exist_ok=True)
    return export_dir


def _new_shuffle_export_path(prefix, suffix=".docx"):
    """Tạo đường dẫn export duy nhất mà KHÔNG pre-create file."""
    export_dir=_shuffle_export_dir()
    safe=re.sub(r"[^A-Za-z0-9_.-]+","_",str(prefix or "export")).strip("._") or "export"
    return export_dir/f"{safe}_{uuid.uuid4().hex}{suffix}"


def _cleanup_old_shuffle_exports(max_age_hours=24):
    """Dọn best-effort file export cũ; bỏ qua file đang bị Windows khóa."""
    try:
        export_dir=_shuffle_export_dir()
        if not export_dir.exists():
            return
        cutoff=time.time()-(max_age_hours*3600)
        for fp in export_dir.iterdir():
            try:
                if fp.is_file() and fp.stat().st_mtime < cutoff:
                    fp.unlink()
            except (PermissionError,OSError):
                pass
    except Exception:
        pass




def _read_generated_file_bytes(path, attempts=40, delay=0.15):
    """Read a generated export only after external Windows apps release it.

    Word + MathType and antivirus scanners can keep a short-lived handle after
    the producer process exits.  Sending the path directly via Flask makes the
    HTTP response depend on that handle.  V5.24.17 instead snapshots the bytes
    first; once this function returns the browser never touches the physical
    file.
    """
    path=Path(path)
    last=None
    for i in range(max(1,int(attempts))):
        try:
            data=path.read_bytes()
            if not data:
                raise RuntimeError(f'File xuất rỗng: {path.name}')
            return data
        except (PermissionError,OSError) as exc:
            last=exc
            time.sleep(delay*(1.0 + min(i,10)/10.0))
    if last:
        raise RuntimeError(
            f'Windows vẫn đang khóa file {path.name} sau khi tạo xong: {last}'
        ) from last
    raise RuntimeError(f'Không đọc được file xuất: {path.name}')


def _send_generated_file_from_memory(path, download_name, mimetype=None):
    """Snapshot generated file to RAM, then best-effort remove physical file."""
    path=Path(path)
    data=_read_generated_file_bytes(path)
    try:
        path.unlink()
    except (PermissionError,OSError):
        # Safe to leave for the 24h cleanup; response already owns the bytes.
        pass
    bio=io.BytesIO(data)
    bio.seek(0)
    return send_file(
        bio,
        as_attachment=True,
        download_name=download_name,
        mimetype=mimetype,
        max_age=0,
    )



def _save_docx_with_retry(doc, preferred_path, attempts=4):
    """Save a DOCX without reusing a Windows-locked pathname.

    If Word/MathType/antivirus still owns a handle, switch to a fresh UUID path
    instead of surfacing Errno 13 to the teacher.
    """
    path=Path(preferred_path)
    last=None
    for i in range(max(1,attempts)):
        try:
            path.parent.mkdir(parents=True,exist_ok=True)
            doc.save(str(path))
            return path
        except PermissionError as exc:
            last=exc
            path=_new_shuffle_export_path(path.stem, path.suffix or '.docx')
            time.sleep(0.15*(i+1))
    if last:
        raise last
    return path

def _build_all_shuffle_docx_file(source, versions, questions):
    _cleanup_old_shuffle_exports()
    path=_new_shuffle_export_path(f"Tat_ca_ma_de_{source['id']}")

    # Ưu tiên DOCX-preserve để giữ MathType OLE thật và vị trí hình.
    preserved = build_all_shuffle_docx_direct_ooxml(
        source,versions,questions,path
    )

    if preserved:
        return path, True

    # NGUỒN DOCX: tuyệt đối không fallback sang bộ dựng block/image.
    # Fallback đó biến MathType OLE thành ảnh và làm thay đổi kích thước hình.
    # Nếu không thể clone OOXML gốc, phải dừng để người dùng biết thay vì
    # tạo ra một file Word nhìn có vẻ hợp lệ nhưng đã mất MathType editable.
    original = _source_saved_path(source)
    source_name = str(source["source_file"] or "").lower()
    if source_name.endswith(".docx") or (original and original.suffix.lower()==".docx"):
        raise RuntimeError(
            "Không thể xuất Word MathType bằng OOXML gốc. "
            "Hệ thống đã dừng để tránh chuyển MathType thành ảnh. "
            "Hãy tải lại file DOCX nguồn bằng phiên bản này rồi đảo lại."
        )

    # Chỉ nguồn PDF mới được phép dùng bộ dựng Word block/image vì PDF vốn
    # không chứa OLE MathType editable để bảo toàn.
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Cm(1.5)
    sec.bottom_margin=Cm(1.5)
    sec.left_margin=Cm(1.6)
    sec.right_margin=Cm(1.6)

    doc.styles["Normal"].font.name="Times New Roman"
    doc.styles["Normal"].font.size=Pt(13.5)

    for vi,version in enumerate(versions):
        if vi>0:
            doc.add_page_break()
        _write_exam_body_docx(doc,source,version,questions,latex_for_mathtype=False,plain_latex=True)

    doc.add_page_break()
    _write_consolidated_answer_matrix_docx(doc,versions)

    path=_save_docx_with_retry(doc,path)
    # V5.24.21: PDF source stays as editable canonical LaTeX in Word.
    # Teacher may convert it to MathType manually using the MathType tab.
    return path, False

@app.get("/teacher/shuffle-exams/<int:source_id>/all.docx")
@login_required("teacher")
def shuffle_exam_all_docx(source_id):
    source,versions,questions = _load_all_shuffle_export_data(
        source_id,session["user_id"]
    )

    if not source:
        abort(404)
    if not versions:
        flash("Chưa có mã đề đã đảo.")
        return redirect(url_for("shuffle_exam_manager"))

    try:
        path,preserved = _build_all_shuffle_docx_file(
            source,versions,questions
        )
    except Exception as e:
        flash("Lỗi tạo file Word: " + str(e) + " | Hãy thử xóa đề cũ và upload lại nếu đề được tạo từ phiên bản trước.")
        return redirect(url_for("shuffle_exam_versions", source_id=source_id))

    if not path.exists():
        flash("Không tạo được file Word.")
        return redirect(url_for("shuffle_exam_versions", source_id=source_id))

    try:
        original = _source_saved_path(source)
        source_name = str(source["source_file"] or "").lower()
        is_pdf = source_name.endswith(".pdf") or (original and original.suffix.lower()==".pdf")
        download_name = "Tat_ca_ma_de_LaTeX.docx" if is_pdf else "Tat_ca_ma_de_MathType.docx"
        return _send_generated_file_from_memory(
            path,
            download_name,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        flash("Lỗi tải file Word sau khi đã tạo: " + str(e))
        return redirect(url_for("shuffle_exam_versions", source_id=source_id))

def _tex_escape_text(value):
    """Escape ordinary Vietnamese text without touching mathematical LaTeX."""
    text=str(value or "")
    replacements={
        "\\":r"\textbackslash{}",
        "&":r"\&","%":r"\%","$":r"\$","#":r"\#","_":r"\_",
        "{":r"\{","}":r"\}","~":r"\textasciitilde{}","^":r"\textasciicircum{}",
    }
    return "".join(replacements.get(ch,ch) for ch in text)


def _shuffle_blocks_export_tex(blocks):
    """Convert stored blocks to XeLaTeX; only true figures remain images."""
    out=[]
    for block in blocks or []:
        btype=str(block.get("type") or "")
        if btype=="latex":
            tex=normalise_latex(block.get("value") or "") or ""
            # XeLaTeX math fonts thường không có chữ Đ. Các chỉ số cực trị
            # tiếng Việt phải dùng font chữ văn bản bên trong công thức.
            tex=re.sub(r"\\mathrm\{(CĐ|CT|CD)\}", r"\\text{\1}", tex)
            tex=re.sub(r"\{(CĐ|CT|CD)\}", r"{\\text{\1}}", tex)
            if not tex:
                continue
            if block.get("display")=="block":
                out.append("\\["+tex+"\\]")
            else:
                out.append("\\("+tex+"\\)")
        elif btype=="image" and block.get("role")=="figure":
            media=_question_media_path(block.get("url") or "")
            if not media or not media.exists():
                continue
            width=float(block.get("display_width_px") or 0)
            width_cm=min(16.0,max(2.0,width*2.54/96.0)) if width>0 else 10.0
            path=str(media.resolve()).replace("\\","/")
            out.append(
                "\\begin{center}\\includegraphics[width="+f"{width_cm:.2f}"+
                "cm]{\\detokenize{"+path+"}}\\end{center}"
            )
        elif btype=="image":
            continue
        else:
            value=str(block.get("value") or "")
            if value:
                out.append(_tex_escape_text(value).replace("\n","\\\\\n"))
    return "".join(out)


def _shuffle_pdf_tex(source, versions, questions):
    """Build a complete XeLaTeX document from the shuffled LaTeX data."""
    qmap={q["id"]:q for q in questions}
    labels={
        1:"PHẦN I. TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN.",
        2:"PHẦN II. TRẮC NGHIỆM ĐÚNG SAI.",
        3:"PHẦN III. TRẮC NGHIỆM TRẢ LỜI NGẮN.",
        4:"PHẦN IV. TỰ LUẬN.",
    }
    parts=["""\\documentclass[12pt,a4paper]{article}
\\usepackage[margin=1.5cm]{geometry}
\\usepackage{fontspec}
\\usepackage{amsmath,amssymb}
\\usepackage{graphicx}
\\usepackage{array,longtable}
\\usepackage{enumitem}
\\defaultfontfeatures{Ligatures=TeX}
\\IfFontExistsTF{Times New Roman}{%
  \\setmainfont{Times New Roman}%
}{%
  \\IfFontExistsTF{DejaVu Serif}{\\setmainfont{DejaVu Serif}}{\\setmainfont{Latin Modern Roman}}%
}
\\setlength{\\parindent}{0pt}
\\setlength{\\parskip}{3pt}
\\newcommand{\\examsection}[1]{\\vspace{4pt}\\textbf{#1}\\par}
\\begin{document}
"""]
    for vi,version in enumerate(versions):
        if vi:
            parts.append("\\newpage\n")
        left="\\\\\n".join(
            _tex_escape_text(x) for x in [source["department"] or "",source["school"] or ""] if x
        )
        right=[_tex_escape_text(source["exam_name"] or source["title"] or "ĐỀ KIỂM TRA")]
        if source["school_year"]:
            right.append(_tex_escape_text("NĂM HỌC "+source["school_year"]))
        if source["subject"]:
            right.append(_tex_escape_text("MÔN: "+source["subject"]))
        right.append(_tex_escape_text("Thời gian: "+(source["duration_text"] or "90 phút")))
        parts.append(
            "\\begin{tabular}{|p{0.42\\textwidth}|p{0.50\\textwidth}|}\\hline\n"
            "\\centering\\textbf{"+left+"} & \\centering\\textbf{"+"\\\\\n".join(right)+"}\\tabularnewline\\hline\n"
            "\\end{tabular}\n"
        )
        parts.append("\\begin{flushright}\\textbf{MÃ ĐỀ: "+_tex_escape_text(version["exam_code"])+"}\\end{flushright}\n")
        last=0
        for item in version["entries"]:
            q=qmap.get(item["question_id"])
            if not q:
                continue
            sec=int(item["section"])
            if sec!=last:
                last=sec
                parts.append("\\examsection{"+_tex_escape_text(labels[sec])+"}\n")
            parts.append("\\textbf{Câu "+str(item["number"])+":} "+_shuffle_blocks_export_tex(q.get("blocks") or [])+"\n\n")
            omap={o["key"]:o for o in q.get("options",[])}
            for oi,key in enumerate(item.get("option_order") or []):
                op=omap.get(key)
                if not op:
                    continue
                label=("abcd"[oi] if sec==2 else "ABCD"[oi])+"."
                parts.append("\\hspace*{0.6cm}\\textbf{"+label+"} "+_shuffle_blocks_export_tex(op.get("blocks") or [])+"\\par\n")

        parts.append("\\newpage\\begin{center}\\Large\\textbf{ĐÁP ÁN — MÃ ĐỀ "+_tex_escape_text(version["exam_code"])+"}\\end{center}\n")
        for sec in (1,2,3):
            rows=[x for x in version["entries"] if int(x["section"])==sec]
            if not rows:
                continue
            parts.append("\\textbf{"+_tex_escape_text(labels[sec])+"}\\par\n")
            answer_line="; ".join("Câu "+str(x["number"])+": "+str(x.get("answer") or "") for x in rows)
            parts.append(_tex_escape_text(answer_line)+"\\par\n")
    parts.append("\\end{document}\n")
    return "".join(parts)


def _find_xelatex_executable():
    configured=(os.environ.get("XELATEX_PATH") or "").strip()
    if configured and Path(configured).exists():
        return configured
    found=shutil.which("xelatex")
    if found:
        return found
    if os.name=="nt":
        candidates=[]
        for root in (os.environ.get("ProgramFiles"), os.environ.get("LOCALAPPDATA")):
            if not root:
                continue
            rp=Path(root)
            candidates.extend([
                rp/"MiKTeX"/"miktex"/"bin"/"x64"/"xelatex.exe",
                rp/"Programs"/"MiKTeX"/"miktex"/"bin"/"x64"/"xelatex.exe",
                rp/"MiKTeX"/"miktex"/"bin"/"xelatex.exe",
            ])
        for c in candidates:
            if c.exists():
                return str(c)
    return None


def _build_shuffle_pdf_file(source, versions, questions, output_path):
    """Compile PDF directly from LaTeX with XeLaTeX; never convert Word MathType."""
    xelatex=_find_xelatex_executable()
    if not xelatex:
        raise RuntimeError(
            "Không tìm thấy XeLaTeX. Nếu dùng MiKTeX trên Windows, hãy đặt XELATEX_PATH "
            "trỏ tới xelatex.exe hoặc thêm thư mục MiKTeX vào PATH."
        )
    output_path=Path(output_path)
    output_path.parent.mkdir(parents=True,exist_ok=True)
    workdir=Path(tempfile.mkdtemp(prefix="shuffle_tex_",dir=str(output_path.parent)))
    tex_path=workdir/"exam.tex"
    tex_path.write_text(_shuffle_pdf_tex(source,versions,questions),encoding="utf-8")
    try:
        proc=subprocess.run(
            [xelatex,"-interaction=nonstopmode","-halt-on-error","-no-shell-escape",tex_path.name],
            cwd=str(workdir),stdout=subprocess.PIPE,stderr=subprocess.STDOUT,timeout=240
        )
        built=workdir/"exam.pdf"
        if proc.returncode!=0 or not built.exists() or built.stat().st_size<1000:
            detail=proc.stdout.decode("utf-8","ignore")[-2500:]
            raise RuntimeError(detail or "XeLaTeX không tạo được PDF.")
        shutil.copy2(built,output_path)
    finally:
        shutil.rmtree(workdir,ignore_errors=True)
    return output_path


@app.get("/teacher/shuffle-exams/<int:source_id>/all-preview.pdf")
@login_required("teacher")
def shuffle_exam_all_preview_pdf(source_id):
    # PDF tất cả mã đề: công thức dựng trực tiếp từ LaTeX/MathJax.
    source,versions,questions=_load_all_shuffle_export_data(source_id,session["user_id"])
    if not source:
        abort(404)
    if not versions:
        abort(404)
    _cleanup_old_shuffle_exports()
    path=_new_shuffle_export_path(f"Tat_ca_ma_de_{source_id}_latex", suffix=".pdf")
    try:
        _build_shuffle_pdf_file(source,versions,questions,path)
        return _send_generated_file_from_memory(
            path,"Tat_ca_ma_de_LaTeX.pdf","application/pdf"
        )
    except Exception as e:
        flash("Không tạo/tải được PDF LaTeX: "+str(e))
        return redirect(url_for("shuffle_exam_versions",source_id=source_id))


@app.get("/teacher/shuffle-exams/<int:source_id>/version/<int:version_id>/pdf")
@login_required("teacher")
def shuffle_exam_pdf(source_id,version_id):
    with db() as c:
        ensure_shuffle_schema(c)
        source=c.execute(
            "SELECT * FROM shuffle_sources WHERE id=? AND uploaded_by=?",
            (source_id,session["user_id"])
        ).fetchone()
        row=c.execute(
            "SELECT * FROM shuffle_versions WHERE id=? AND source_id=?",
            (version_id,source_id)
        ).fetchone()
        if not source or not row:
            abort(404)
        payload=json.loads(row["version_json"])
        version={"exam_code":payload.get("exam_code",689),"entries":payload.get("entries",[])}
        questions=load_shuffle_questions(c,source_id)
    _cleanup_old_shuffle_exports()
    path=_new_shuffle_export_path(f"De_{version['exam_code']}_latex", suffix=".pdf")
    try:
        _build_shuffle_pdf_file(source,[version],questions,path)
        return _send_generated_file_from_memory(
            path,f"De_{version['exam_code']}_LaTeX.pdf","application/pdf"
        )
    except Exception as e:
        flash("Không tạo/tải được PDF LaTeX: "+str(e))
        return redirect(url_for("shuffle_exam_version",source_id=source_id,version_id=version_id))


@app.get("/teacher/shuffle-exams/<int:source_id>/version/<int:version_id>/docx")
@login_required("teacher")
def shuffle_exam_docx(source_id,version_id):
    with db() as c:
        ensure_shuffle_schema(c)
        source=c.execute("SELECT * FROM shuffle_sources WHERE id=? AND uploaded_by=?",(source_id,session["user_id"])).fetchone()
        row=c.execute("SELECT * FROM shuffle_versions WHERE id=? AND source_id=?",(version_id,source_id)).fetchone()
        if not source or not row: abort(404)
        payload=json.loads(row["version_json"])
        version={"exam_code":payload.get("exam_code",689),"entries":payload.get("entries",[])}
        questions=load_shuffle_questions(c,source_id)

    _cleanup_old_shuffle_exports()
    path=_new_shuffle_export_path(f"De_{version['exam_code']}")

    # DOCX nguồn: clone trực tiếp XML của câu gốc để giữ w:object/o:OLEObject.
    # Nhờ vậy MathType sau khi trộn vẫn là OLE MathType có thể mở/chỉnh sửa,
    # không bị chuyển thành ảnh hoặc công thức dựng lại.
    try:
        preserved = build_all_shuffle_docx_direct_ooxml(
            source, [version], questions, path
        )
        if not preserved:
            original = _source_saved_path(source)
            source_name = str(source["source_file"] or "").lower()
            if source_name.endswith(".docx") or (original and original.suffix.lower()==".docx"):
                raise RuntimeError(
                    "Không thể xuất Word MathType bằng OOXML gốc; "
                    "đã dừng để tránh biến công thức thành ảnh."
                )
            # V5.24.21: PDF không còn tự động gọi MathType. Xuất Word
            # với LaTeX chuẩn ($...$/\[...\]) giống tiện ích PDF sang Word;
            # giáo viên tự Toggle TeX -> MathType sau khi tải nếu muốn.
            build_shuffle_docx(source,version,questions,path,plain_latex=True)
    except Exception as e:
        flash("Lỗi tạo file Word: " + str(e))
        return redirect(url_for("shuffle_exam_version",source_id=source_id,version_id=version_id))

    try:
        original = _source_saved_path(source)
        source_name = str(source["source_file"] or "").lower()
        is_pdf = source_name.endswith(".pdf") or (original and original.suffix.lower()==".pdf")
        download_name = (
            f"De_{version['exam_code']}_LaTeX.docx" if is_pdf
            else f"De_{version['exam_code']}_MathType.docx"
        )
        return _send_generated_file_from_memory(
            path,
            download_name,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        flash("Lỗi tải file Word sau khi đã tạo: " + str(e))
        return redirect(url_for("shuffle_exam_version",source_id=source_id,version_id=version_id))

@app.post("/teacher/shuffle-exams/<int:source_id>/delete")
@login_required("teacher")
def shuffle_exam_delete(source_id):
    with db() as c:
        ensure_shuffle_schema(c)

        source = c.execute("""
            SELECT * FROM shuffle_sources
            WHERE id=? AND uploaded_by=?
        """,(source_id,session["user_id"])).fetchone()

        if not source:
            abort(404)

        qids = [
            r["id"] for r in c.execute(
                "SELECT id FROM shuffle_questions WHERE source_id=?",
                (source_id,)
            ).fetchall()
        ]

        if qids:
            marks = ",".join("?" for _ in qids)
            c.execute(
                f"DELETE FROM shuffle_options "
                f"WHERE shuffle_question_id IN ({marks})",
                qids
            )

        c.execute("DELETE FROM shuffle_versions WHERE source_id=?",(source_id,))
        c.execute("DELETE FROM shuffle_questions WHERE source_id=?",(source_id,))
        c.execute("DELETE FROM shuffle_sources WHERE id=?",(source_id,))

    flash("Đã xóa đề khỏi mục Đảo đề.")
    return redirect(url_for("shuffle_exam_manager"))


# V5.16.0 — ĐẢM BẢO SCHEMA NGÂN HÀNG ĐỀ
# Chỉ tạo bank_* nếu thiếu; không sửa exams/questions/options cũ.
# ============================================================
def ensure_question_bank_schema(c):
    c.execute("""
        CREATE TABLE IF NOT EXISTS bank_folders(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL COLLATE NOCASE,
            created_by INTEGER NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(created_by,name)
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS bank_sources(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            source_file TEXT,
            stored_name TEXT DEFAULT '',
            uploaded_by INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    source_cols = {
        row[1] for row in c.execute("PRAGMA table_info(bank_sources)").fetchall()
    }
    if "stored_name" not in source_cols:
        c.execute("ALTER TABLE bank_sources ADD COLUMN stored_name TEXT DEFAULT ''")
        source_cols.add("stored_name")
    if "folder_id" not in source_cols:
        c.execute("ALTER TABLE bank_sources ADD COLUMN folder_id INTEGER")
    c.execute(
        "CREATE INDEX IF NOT EXISTS idx_bank_sources_folder_id "
        "ON bank_sources(folder_id)"
    )
    c.execute("""
        CREATE TABLE IF NOT EXISTS bank_questions(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_id INTEGER NOT NULL,
            original_order INTEGER NOT NULL,
            section_order INTEGER NOT NULL,
            section INTEGER NOT NULL,
            qtype TEXT NOT NULL,
            blocks_json TEXT NOT NULL DEFAULT '[]',
            solution_json TEXT NOT NULL DEFAULT '[]',
            correct_option TEXT,
            short_answer TEXT,
            tf_answers_json TEXT NOT NULL DEFAULT '[]',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    bank_qcols = {
        row[1] for row in c.execute("PRAGMA table_info(bank_questions)").fetchall()
    }
    if "solution_json" not in bank_qcols:
        c.execute(
            "ALTER TABLE bank_questions "
            "ADD COLUMN solution_json TEXT NOT NULL DEFAULT '[]'"
        )
    c.execute("""
        CREATE TABLE IF NOT EXISTS bank_options(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bank_question_id INTEGER NOT NULL,
            option_key TEXT NOT NULL,
            blocks_json TEXT NOT NULL DEFAULT '[]',
            correct_tf INTEGER,
            UNIQUE(bank_question_id, option_key)
        )
    """)


def ensure_practice_exam_schema(c):
    """Add the practice linkage without changing any regular exam data."""
    columns = {
        row[1] for row in c.execute("PRAGMA table_info(exams)").fetchall()
    }
    if "is_practice" not in columns:
        c.execute(
            "ALTER TABLE exams "
            "ADD COLUMN is_practice INTEGER NOT NULL DEFAULT 0"
        )
    if "practice_source_id" not in columns:
        c.execute("ALTER TABLE exams ADD COLUMN practice_source_id INTEGER")
    attempt_columns = {
        row[1] for row in c.execute("PRAGMA table_info(student_exams)").fetchall()
    }
    if "last_activity_at" not in attempt_columns:
        c.execute("ALTER TABLE student_exams ADD COLUMN last_activity_at TEXT")
        c.execute(
            "UPDATE student_exams SET last_activity_at=started_at "
            "WHERE last_activity_at IS NULL"
        )
    c.execute(
        "CREATE UNIQUE INDEX IF NOT EXISTS idx_exams_practice_source "
        "ON exams(practice_source_id) "
        "WHERE is_practice=1 AND practice_source_id IS NOT NULL"
    )


# V5.15.0 — NGÂN HÀNG ĐỀ: UPLOAD WORD/PDF
# Dùng lại extract_word_blocks/extract_pdf_blocks hiện có,
# nhưng lưu câu vào bảng bank_* riêng, KHÔNG tạo exam ngay.
# ============================================================
# ============================================================
# V5.16.2 — TẠO ĐỀ TỪ BẢN NHÁP AI
# Tạo exam nháp theo schema cũ, sau đó chuyển sang màn hình kiểm tra
# hiện tại để giáo viên có thể kiểm tra lần cuối và bấm Xuất bản.
# ============================================================
@app.post("/teacher/create-exam/ai/publish")
@login_required("teacher")
# ============================================================
# V5.21.10 — Giao diện học sinh Phần II hiển thị a), b), c), d).
# DB vẫn giữ option_key A/B/C/D để tương thích logic chấm điểm.
# ============================================================
def publish_ai_draft():
    """
    V5.16.3 — trang AI là bước xem/xóa/thêm câu.
    Khi bấm Xuất bản, đề vào hệ thống cũ với published=1.
    """
    try:
        questions=json.loads(request.form.get("questions_json") or "[]")
    except Exception:
        questions=[]

    if not questions:
        flash("Không có câu hỏi AI để xuất bản.")
        return redirect(url_for("create_exam_ai"))

    title=normalize_text(request.form.get("title")) or "Đề tạo bằng AI"
    duration=int(request.form.get("duration") or 90)

    try:
        pts=[float(str(request.form.get(k,"0")).replace(",",".")) for k in ("p1","p2","p3","p4")]
    except ValueError:
        flash("Điểm các phần không hợp lệ.")
        return redirect(url_for("create_exam_ai"))

    if abs(sum(pts)-10)>0.001:
        flash("Tổng điểm 4 phần phải bằng 10.")
        return redirect(url_for("create_exam_ai"))

    n2=sum(1 for q in questions if int(q.get("section") or 1)==2)
    if abs(pts[1]-n2)>0.001:
        flash(f"Phần II có {n2} câu nên điểm Phần II phải bằng {n2} điểm.")
        return redirect(url_for("create_exam_ai"))

    sec_counter={1:0,2:0,3:0,4:0}

    with db() as c:
        cur=c.execute("""
            INSERT INTO exams(
                title,source_file,question_count,published,duration_minutes,
                part1_points,part2_points,part3_points,part4_points,
                p2_item_a,p2_item_b,p2_item_c,p2_item_d,
                part2_scoring_mode,anti_cheat,max_violations
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """,(
            title,"ai_generated",len(questions),1,duration,*pts,
            0.25,0.25,0.25,0.25,
            int(request.form.get("p2_mode") or 1),
            1 if request.form.get("anti_cheat") else 0,
            int(request.form.get("max_violations") or 3)
        ))
        exam_id=cur.lastrowid
        c.execute("UPDATE exams SET teacher_id=? WHERE id=?",(session["user_id"],exam_id))

        for global_order,item in enumerate(questions,start=1):
            sec=int(item.get("section") or 1)
            if sec not in (1,2,3,4):
                sec=1
            sec_counter[sec]+=1

            stem=manual_latex_blocks(item.get("stem") or "")
            options_raw=item.get("options") or []
            answer=str(item.get("answer") or "").strip()

            qtype="mcq" if sec==1 else "tf" if sec==2 else "short" if sec==3 else "essay"
            correct_option=None
            short_answer=None
            tf_answers=[]

            if sec==1:
                if len(options_raw)!=4:
                    flash("Mỗi câu Phần I phải có đúng 4 phương án.")
                    return redirect(url_for("create_exam_ai"))
                correct_option=answer.upper()
                if correct_option not in ("A","B","C","D"):
                    flash("Đáp án Phần I phải là A/B/C/D.")
                    return redirect(url_for("create_exam_ai"))

            elif sec==2:
                if len(options_raw)!=4:
                    flash("Mỗi câu Phần II phải có đúng 4 ý a/b/c/d.")
                    return redirect(url_for("create_exam_ai"))
                tf_answers=parse_tf_key(answer)
                if len(tf_answers)!=4:
                    flash("Đáp án Phần II phải có đúng 4 giá trị Đ/S.")
                    return redirect(url_for("create_exam_ai"))

            elif sec==3:
                if not re.fullmatch(r"-?\d+(?:,\d+)?",answer or ""):
                    flash("Đáp án Phần III chỉ dùng chữ số, dấu - và dấu phẩy.")
                    return redirect(url_for("create_exam_ai"))
                short_answer=answer

            stem_json=jdump(stem)
            qid=c.execute("""
                INSERT INTO questions(
                    exam_id,original_order,section_order,section,qtype,
                    blocks_json,stem_json,image_url,correct_option,
                    short_answer,tf_answers_json
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?)
            """,(
                exam_id,global_order,sec_counter[sec],sec,qtype,
                stem_json,stem_json,None,correct_option,short_answer,jdump(tf_answers)
            )).lastrowid

            for oi,op in enumerate(options_raw[:4]):
                # V5.21.9:
                # DB vẫn lưu option_key A/B/C/D để tương thích toàn hệ thống.
                # Riêng giao diện AI Phần II hiển thị a), b), c), d).
                # Không đổi key DB thành chữ thường để tránh ảnh hưởng chấm điểm.
                key="ABCD"[oi]
                oj=jdump(manual_latex_blocks(op))
                tf_val=None
                if sec==2 and oi<len(tf_answers):
                    tf_val=1 if tf_answers[oi] else 0
                c.execute("""
                    INSERT INTO options(
                        question_id,option_key,blocks_json,content_json,
                        image_url,correct_tf
                    ) VALUES(?,?,?,?,?,?)
                """,(qid,key,oj,oj,None,tf_val))

    flash("Đã xuất bản đề AI. Học sinh có thể vào làm bài.")
    return redirect(url_for("teacher_dashboard"))



@app.post("/teacher/question-bank/folders")
@login_required("teacher")
def teacher_create_bank_folder():
    name = normalize_text(request.form.get("name"))
    if not name:
        flash("Vui lòng nhập tên thư mục ngân hàng.")
        return redirect(url_for("question_bank"))
    if len(name) > 120:
        flash("Tên thư mục tối đa 120 ký tự.")
        return redirect(url_for("question_bank"))
    try:
        with db() as c:
            ensure_question_bank_schema(c)
            c.execute(
                "INSERT INTO bank_folders(name,created_by) VALUES(?,?)",
                (name, session["user_id"]),
            )
    except sqlite3.IntegrityError:
        flash("Bạn đã có một thư mục ngân hàng cùng tên.")
        return redirect(url_for("question_bank"))
    flash("Đã tạo thư mục ngân hàng “" + name + "”.")
    return redirect(url_for("question_bank"))


@app.post("/teacher/question-bank/folders/<int:folder_id>/rename")
@login_required("teacher")
def teacher_rename_bank_folder(folder_id):
    try:
        with db() as c:
            ensure_question_bank_schema(c)
            name = _rename_owned_folder(
                c, "bank_folders", folder_id,
                session["user_id"], request.form.get("name"),
            )
    except LookupError:
        abort(404)
    except ValueError as error:
        flash(str(error))
        return redirect(url_for("teacher_bank_folder", folder_id=folder_id))
    except sqlite3.IntegrityError:
        flash("Bạn đã có một thư mục ngân hàng cùng tên.")
        return redirect(url_for("teacher_bank_folder", folder_id=folder_id))
    flash("Đã đổi tên thư mục ngân hàng thành “" + name + "”.")
    return redirect(url_for("teacher_bank_folder", folder_id=folder_id))


@app.route("/teacher/question-bank", methods=["GET","POST"])
@login_required("teacher")
def question_bank():
    if request.method == "POST":
        raw_folder_id = request.form.get("folder_id")
        return_url = (
            url_for("teacher_bank_folder", folder_id=int(raw_folder_id))
            if str(raw_folder_id or "").isdigit() and int(raw_folder_id) > 0
            else url_for("question_bank")
        )
        f = request.files.get("file")
        title = normalize_text(request.form.get("title")) or "Ngân hàng đề"

        if not f or not f.filename or ext(f.filename) not in ALLOWED:
            flash("Ngân hàng đề chỉ nhận DOCX hoặc PDF.")
            return redirect(return_url)

        try:
            with db() as c:
                ensure_question_bank_schema(c)
                folder_id = _owned_folder_id(
                    c, "bank_folders", request.form.get("folder_id"),
                    session["user_id"],
                )
        except ValueError as error:
            flash(str(error))
            return redirect(return_url)

        saved = UPLOAD_DIR / (
            uuid.uuid4().hex + "_bank_" + secure_filename(f.filename)
        )
        f.save(saved)

        try:
            source_ext = ext(f.filename)

            # --------------------------------------------------------
            # WORD/PDF VẪN XỬ LÝ RIÊNG, giữ nguyên code parser cũ.
            # --------------------------------------------------------
            if source_ext == "docx":
                lines = extract_word_blocks(saved)
                questions = build_questions(lines)
                _sync_docx_question_figures_from_ooxml(questions, saved)
            else:
                lines = extract_pdf_blocks(saved)
                questions = build_questions(lines)
                if not questions:
                    lines = extract_pdf_blocks_fallback(saved)
                    questions = build_questions(lines)

            if not questions:
                raise RuntimeError("Không nhận diện được câu hỏi trong file ngân hàng.")

            with db() as c:
                ensure_question_bank_schema(c)
                cur = c.execute("""
                    INSERT INTO bank_sources(
                        title,source_file,stored_name,uploaded_by,folder_id
                    ) VALUES(?,?,?,?,?)
                """,(
                    title,f.filename,saved.name,session["user_id"],folder_id
                ))
                source_id = cur.lastrowid

                for q in questions:
                    curq = c.execute("""
                        INSERT INTO bank_questions(
                            source_id,original_order,section_order,
                            section,qtype,blocks_json,solution_json,
                            correct_option,short_answer,tf_answers_json
                        ) VALUES(?,?,?,?,?,?,?,?,?,?)
                    """,(
                        source_id,
                        q["original_order"],
                        q["section_order"],
                        q["section"],
                        q["qtype"],
                        jdump(q["stem"]),
                        jdump(q.get("solution",[])),
                        q.get("correct_option"),
                        q.get("short_answer"),
                        jdump(q.get("tf_answers",[]))
                    ))
                    bank_qid = curq.lastrowid

                    for op in q.get("options",[]):
                        tf_val = None
                        if q["section"] == 2:
                            vals = q.get("tf_answers",[])
                            oi = ord(op["key"].upper())-ord("A")
                            if 0 <= oi < len(vals):
                                tf_val = 1 if vals[oi] else 0

                        c.execute("""
                            INSERT INTO bank_options(
                                bank_question_id,option_key,
                                blocks_json,correct_tf
                            ) VALUES(?,?,?,?)
                        """,(
                            bank_qid,
                            op["key"],
                            jdump(op["content"]),
                            tf_val
                        ))

            flash(
                f"Đã đưa {len(questions)} câu vào Ngân hàng đề. "
                "Giáo viên có thể chọn câu để tạo bài kiểm tra."
            )
            return redirect(return_url)

        except Exception as e:
            flash("Lỗi tải Ngân hàng đề: " + str(e))
            return redirect(return_url)

    with db() as c:
        ensure_question_bank_schema(c)
        folders = c.execute("""
            SELECT f.*,
                   (SELECT COUNT(*) FROM bank_sources s
                    WHERE s.folder_id=f.id AND s.uploaded_by=?) AS item_count
            FROM bank_folders f
            WHERE f.created_by=?
            ORDER BY f.name COLLATE NOCASE,f.id
        """,(session["user_id"],session["user_id"])).fetchall()
        unfiled_count = c.execute("""
            SELECT COUNT(*) FROM bank_sources
            WHERE uploaded_by=? AND folder_id IS NULL
        """,(session["user_id"],)).fetchone()[0]

    return render_template(
        "question_bank.html",
        folders=folders,
        unfiled_count=unfiled_count,
    )


@app.get("/teacher/question-bank/folder/<int:folder_id>")
@login_required("teacher")
def teacher_bank_folder(folder_id):
    with db() as c:
        ensure_question_bank_schema(c)
        folder = c.execute("""
            SELECT * FROM bank_folders
            WHERE id=? AND created_by=?
        """,(folder_id,session["user_id"])).fetchone()
        if not folder:
            abort(404)
        sources = c.execute("""
            SELECT s.*,
                   (SELECT COUNT(*) FROM bank_questions q
                    WHERE q.source_id=s.id) AS question_count
            FROM bank_sources s
            WHERE s.folder_id=? AND s.uploaded_by=?
            ORDER BY s.id DESC
        """,(folder_id,session["user_id"])).fetchall()
    return render_template(
        "question_bank_folder.html",
        folder=folder,
        sources=sources,
        is_unfiled=False,
    )


@app.get("/teacher/question-bank/unfiled")
@login_required("teacher")
def teacher_unfiled_bank():
    with db() as c:
        ensure_question_bank_schema(c)
        sources = c.execute("""
            SELECT s.*,
                   (SELECT COUNT(*) FROM bank_questions q
                    WHERE q.source_id=s.id) AS question_count
            FROM bank_sources s
            WHERE s.folder_id IS NULL AND s.uploaded_by=?
            ORDER BY s.id DESC
        """,(session["user_id"],)).fetchall()
    return render_template(
        "question_bank_folder.html",
        folder={"id":None,"name":"Ngân hàng chưa phân loại"},
        sources=sources,
        is_unfiled=True,
    )


# ============================================================
# V5.15.0 — XEM ĐẦY ĐỦ CÂU TRONG MỘT NGÂN HÀNG
# ============================================================
@app.get("/teacher/question-bank/<int:source_id>")
@login_required("teacher")
def question_bank_detail(source_id):
    ensure_bank_question_media(source_id)
    with db() as c:
        ensure_question_bank_schema(c)
        source = c.execute("""
            SELECT s.*,f.name AS folder_name
            FROM bank_sources s
            LEFT JOIN bank_folders f ON f.id=s.folder_id
            WHERE s.id=? AND s.uploaded_by=?
        """,(source_id,session["user_id"])).fetchone()

        if not source:
            abort(404)

        rows = c.execute("""
            SELECT * FROM bank_questions
            WHERE source_id=?
            ORDER BY section,section_order,original_order,id
        """,(source_id,)).fetchall()

        questions = []
        for q in rows:
            options = c.execute("""
                SELECT * FROM bank_options
                WHERE bank_question_id=?
                ORDER BY option_key
            """,(q["id"],)).fetchall()

            tf_vals = json.loads(q["tf_answers_json"] or "[]")

            if q["section"] == 1:
                answer_text = q["correct_option"] or ""
            elif q["section"] == 2:
                answer_text = "".join("Đ" if x else "S" for x in tf_vals)
            elif q["section"] == 3:
                answer_text = q["short_answer"] or ""
            else:
                answer_text = "Giáo viên chấm"

            questions.append({
                "id": q["id"],
                "section": q["section"],
                "section_order": q["section_order"],
                "qtype": q["qtype"],
                "blocks": json.loads(q["blocks_json"] or "[]"),
                "solution": json.loads(q["solution_json"] or "[]"),
                "options": [
                    {
                        "key": o["option_key"],
                        "blocks": json.loads(o["blocks_json"] or "[]")
                    }
                    for o in options
                ],
                "answer_text": answer_text
            })

    return render_template(
        "question_bank_detail.html",
        source=source,
        questions=questions
    )


# ============================================================

# ============================================================
# V5.21.5 — QUY TRÌNH NGÂN HÀNG ĐỀ RIÊNG BIỆT
# ============================================================
def get_owned_bank_exam(c, exam_id, teacher_id):
    """
    Lấy đề được tạo từ Ngân hàng và xác nhận đúng giáo viên sở hữu.

    Bảng exams cũ không có teacher_id, vì vậy ta:
    1. đọc exams.source_file = "question_bank:<source_id>";
    2. lấy source_id;
    3. đối chiếu bank_sources.uploaded_by với giáo viên đăng nhập.
    """
    exam = c.execute(
        "SELECT * FROM exams WHERE id=?",
        (exam_id,)
    ).fetchone()

    if not exam:
        return None, None

    source_file = str(exam["source_file"] or "")
    m = re.fullmatch(r"question_bank:(\d+)", source_file)
    if not m:
        return None, None

    source_id = int(m.group(1))
    source = c.execute(
        "SELECT * FROM bank_sources WHERE id=? AND uploaded_by=?",
        (source_id, teacher_id)
    ).fetchone()

    if not source:
        return None, None

    return exam, source


def ensure_bank_workflow_column(c):
    """
    Đảm bảo database cũ có cột draft_checked.
    Hàm được gọi ngay trong từng route nên không phụ thuộc migration cũ.
    """
    cols = {
        row[1]
        for row in c.execute("PRAGMA table_info(exams)").fetchall()
    }
    if "draft_checked" not in cols:
        c.execute(
            "ALTER TABLE exams ADD COLUMN draft_checked INTEGER DEFAULT 0"
        )


# V5.21.7 — QUY TRÌNH TẠO ĐỀ TỪ NGÂN HÀNG CHỈ 3 BƯỚC
# ------------------------------------------------------------
# Bước 1: Chọn câu.
# Bước 2: Tạo đề và hiển thị ngay các câu đã chọn.
#         Có nút Hủy để quay lại chọn bộ câu mới.
# Bước 3: Xuất bản cho học sinh.
#
# Không còn bước trung gian khác trong luồng Ngân hàng đề.
# ============================================================

def _safe_float(value, default=0.0):
    """Đọc số thực từ form; hỗ trợ dấu phẩy thập phân."""
    try:
        return float(str(value).strip().replace(",", "."))
    except Exception:
        return float(default)


def _safe_int(value, default=0, minimum=None, maximum=None):
    """Đọc số nguyên an toàn, không để form lỗi làm hỏng Bước 2."""
    try:
        v = int(str(value).strip())
    except Exception:
        v = int(default)

    if minimum is not None:
        v = max(minimum, v)
    if maximum is not None:
        v = min(maximum, v)
    return v


def normalize_bank_exam_points(requested_points, selected_section_counts):
    """
    ============================================================
    CÂN ĐIỂM TỰ ĐỘNG CHO ĐỀ TẠO TỪ NGÂN HÀNG
    ============================================================
    Quy tắc cũ của hệ thống:
    - Phần II: mỗi câu tối đa 1 điểm.
    => tổng điểm Phần II phải bằng số câu Phần II được chọn.

    Để Bước 2 KHÔNG BỊ CHẶN:
    - p2 tự động = số câu Phần II.
    - p1, p3, p4 giữ tỷ lệ giáo viên nhập trong phần điểm còn lại.
    - nếu một phần không có câu thì điểm phần đó = 0.
    - tổng cuối cùng luôn = 10 (trừ trường hợp chọn >10 câu Phần II).

    Hàm trả về:
        [p1,p2,p3,p4], note
    """
    counts = {
        sec: int(selected_section_counts.get(sec, 0) or 0)
        for sec in (1,2,3,4)
    }

    n2 = counts[2]
    if n2 > 10:
        raise ValueError(
            "Phần II có hơn 10 câu được chọn. "
            "Mỗi câu tối đa 1 điểm nên tổng điểm sẽ vượt 10."
        )

    requested = list(requested_points or [2,2,3,3])
    while len(requested) < 4:
        requested.append(0.0)

    # Phần không có câu không nhận điểm.
    for idx,sec in enumerate((1,2,3,4)):
        if counts[sec] <= 0:
            requested[idx] = 0.0

    p2 = float(n2)
    remaining = max(0.0, 10.0 - p2)

    active_other = [
        idx for idx,sec in enumerate((1,3,4))
        if counts[sec] > 0
    ]

    result = [0.0, p2, 0.0, 0.0]

    # Map index trong requested: P1=0, P3=2, P4=3
    idx_map = [0,2,3]
    active_idx = [
        idx_map[i]
        for i,sec in enumerate((1,3,4))
        if counts[sec] > 0
    ]

    if active_idx:
        weights = [
            max(0.0, float(requested[i] or 0))
            for i in active_idx
        ]
        total_weight = sum(weights)

        if total_weight <= 0:
            weights = [1.0] * len(active_idx)
            total_weight = float(len(active_idx))

        allocated = 0.0
        for k,i in enumerate(active_idx):
            if k == len(active_idx)-1:
                value = remaining - allocated
            else:
                value = round(remaining * weights[k] / total_weight, 4)
                allocated += value
            result[i] = max(0.0, value)

    # Nếu chỉ có Phần II thì p2 phải là 10 mới đủ tổng;
    # trường hợp n2<10 vẫn giữ đúng điểm thật, không giả tăng điểm mỗi câu.
    total = sum(result)

    note = None
    original = [
        round(float(x or 0),4)
        for x in requested[:4]
    ]

    if abs(original[1] - p2) > 0.001 or abs(total - 10.0) < 0.001:
        note = (
            f"Điểm đã được cân tự động theo các câu được chọn: "
            f"Phần I={result[0]:g}, Phần II={result[1]:g}, "
            f"Phần III={result[2]:g}, Phần IV={result[3]:g}."
        )

    return result, note


@app.post("/teacher/question-bank/<int:source_id>/build-exam")
@login_required("teacher")
def build_exam_from_bank(source_id):
    """
    ============================================================
    BƯỚC 2 — TẠO ĐỀ TỪ CÁC CÂU ĐÃ CHỌN (V5.21.8)
    ============================================================

    Yêu cầu:
    - chỉ cần chọn ít nhất 1 câu;
    - không còn bước Kiểm tra;
    - không chặn vì điểm Phần II nhập không khớp số câu;
    - tạo xong chuyển thẳng sang trang HIỂN THỊ ĐỀ;
    - tại trang đó: Hủy hoặc Xuất bản.
    ============================================================
    """
    selected = []

    # --------------------------------------------------------
    # KHỐI 1 — đọc danh sách câu được chọn.
    # --------------------------------------------------------
    for value in request.form.getlist("question_ids"):
        try:
            qid = int(value)
            if qid > 0 and qid not in selected:
                selected.append(qid)
        except Exception:
            continue

    if not selected:
        flash(
            "Bước 1 chưa có câu nào được chọn. "
            "Hãy tích ít nhất một câu rồi bấm Bước 2 — Tạo đề."
        )
        return redirect(
            url_for("question_bank_detail", source_id=source_id)
        )

    # --------------------------------------------------------
    # KHỐI 2 — đọc thông tin đề an toàn.
    # Không để dữ liệu form sai kiểu gây lỗi 500.
    # --------------------------------------------------------
    title = normalize_text(request.form.get("title")) or "Đề từ ngân hàng"
    duration = _safe_int(
        request.form.get("duration"),
        default=90,
        minimum=1,
        maximum=600
    )

    requested_pts = [
        _safe_float(request.form.get("p1"), 2),
        _safe_float(request.form.get("p2"), 2),
        _safe_float(request.form.get("p3"), 3),
        _safe_float(request.form.get("p4"), 3),
    ]

    p2_mode = _safe_int(
        request.form.get("p2_mode"),
        default=1,
        minimum=1,
        maximum=2
    )
    if p2_mode not in (1,2):
        p2_mode = 1

    max_violations = _safe_int(
        request.form.get("max_violations"),
        default=3,
        minimum=1,
        maximum=100
    )

    marks = ",".join("?" for _ in selected)

    try:
        with db() as c:
            # ------------------------------------------------
            # KHỐI 3 — đảm bảo bảng Ngân hàng tồn tại.
            # ------------------------------------------------
            ensure_question_bank_schema(c)

            source = c.execute("""
                SELECT * FROM bank_sources
                WHERE id=? AND uploaded_by=?
            """,(source_id,session["user_id"])).fetchone()

            if not source:
                abort(404)

            # ------------------------------------------------
            # KHỐI 4 — chỉ lấy đúng câu thuộc ngân hàng này.
            # ------------------------------------------------
            bank_rows = c.execute(f"""
                SELECT * FROM bank_questions
                WHERE source_id=? AND id IN ({marks})
                ORDER BY section,section_order,original_order,id
            """,(source_id,*selected)).fetchall()

            if not bank_rows:
                flash(
                    "Không tìm thấy các câu đã chọn trong Ngân hàng. "
                    "Hãy tải lại trang và chọn lại."
                )
                return redirect(
                    url_for("question_bank_detail", source_id=source_id)
                )

            if len(bank_rows) != len(selected):
                flash(
                    "Một số câu đã chọn không còn tồn tại. "
                    "Hệ thống chỉ tạo đề khi tất cả câu đều hợp lệ."
                )
                return redirect(
                    url_for("question_bank_detail", source_id=source_id)
                )

            # ------------------------------------------------
            # KHỐI 5 — đếm câu từng phần và tự cân điểm.
            #
            # Đây là điểm sửa quan trọng:
            # KHÔNG redirect về Bước 1 chỉ vì p2 != số câu PII.
            # ------------------------------------------------
            counts = {1:0,2:0,3:0,4:0}
            for row in bank_rows:
                sec = int(row["section"] or 1)
                if sec in counts:
                    counts[sec] += 1

            try:
                pts, point_note = normalize_bank_exam_points(
                    requested_pts,
                    counts
                )
            except ValueError as e:
                flash(str(e))
                return redirect(
                    url_for("question_bank_detail", source_id=source_id)
                )

            # ------------------------------------------------
            # KHỐI 6 — tạo exam ở trạng thái nháp published=0.
            # ------------------------------------------------
            cur = c.execute("""
                INSERT INTO exams(
                    title,source_file,question_count,published,
                    duration_minutes,
                    part1_points,part2_points,part3_points,part4_points,
                    p2_item_a,p2_item_b,p2_item_c,p2_item_d,
                    part2_scoring_mode,anti_cheat,max_violations
                ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,(
                title,
                "question_bank:" + str(source_id),
                len(bank_rows),
                0,
                duration,
                *pts,
                0.25,0.25,0.25,0.25,
                p2_mode,
                1 if request.form.get("anti_cheat") else 0,
                max_violations
            ))
            exam_id = cur.lastrowid
            c.execute("UPDATE exams SET teacher_id=? WHERE id=?",(session["user_id"],exam_id))

            # ------------------------------------------------
            # KHỐI 7 — clone nguyên câu + đáp án + options
            # từ Ngân hàng sang đề nháp.
            # ------------------------------------------------
            sec_counter = {1:0,2:0,3:0,4:0}
            global_order = 0

            for bq in bank_rows:
                global_order += 1
                sec = int(bq["section"] or 1)
                if sec not in sec_counter:
                    sec = 1
                sec_counter[sec] += 1

                stem = bq["blocks_json"] or "[]"

                # JSON lỗi ở một câu không được làm hỏng cả đề.
                try:
                    stem_blocks = json.loads(stem)
                except Exception:
                    stem_blocks = []
                    stem = "[]"

                qid = c.execute("""
                    INSERT INTO questions(
                        exam_id,original_order,section_order,
                        section,qtype,blocks_json,stem_json,solution_json,
                        image_url,correct_option,short_answer,tf_answers_json
                    ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
                """,(
                    exam_id,
                    global_order,
                    sec_counter[sec],
                    sec,
                    bq["qtype"],
                    stem,
                    stem,
                    bq["solution_json"] or "[]",
                    first_image(stem_blocks),
                    bq["correct_option"],
                    bq["short_answer"],
                    bq["tf_answers_json"] or "[]"
                )).lastrowid

                ops = c.execute("""
                    SELECT * FROM bank_options
                    WHERE bank_question_id=?
                    ORDER BY option_key
                """,(bq["id"],)).fetchall()

                for op in ops:
                    oj = op["blocks_json"] or "[]"

                    try:
                        option_blocks = json.loads(oj)
                    except Exception:
                        option_blocks = []
                        oj = "[]"

                    c.execute("""
                        INSERT INTO options(
                            question_id,option_key,
                            blocks_json,content_json,
                            image_url,correct_tf
                        ) VALUES(?,?,?,?,?,?)
                    """,(
                        qid,
                        op["option_key"],
                        oj,
                        oj,
                        first_image(option_blocks),
                        op["correct_tf"]
                    ))

        # ----------------------------------------------------
        # KHỐI 8 — Bước 2 thành công:
        # chuyển THẲNG sang trang hiển thị đề vừa tạo.
        # ----------------------------------------------------
        flash(
            f"Bước 2 hoàn tất: đã tạo đề gồm {len(bank_rows)} câu."
        )
        if point_note:
            flash(point_note)

        return redirect(
            url_for("bank_exam_preview", exam_id=exam_id)
        )

    except Exception as e:
        # Không để lỗi DB/JSON biến thành trang trắng.
        flash(
            "Không tạo được đề từ Ngân hàng: " + str(e)
        )
        return redirect(
            url_for("question_bank_detail", source_id=source_id)
        )



@app.get("/teacher/question-bank/exam/<int:exam_id>/preview")
@login_required("teacher")
def bank_exam_preview(exam_id):
    """
    BƯỚC 2 — HIỂN THỊ ĐỀ ĐÃ TẠO TỪ CÂU ĐƯỢC CHỌN.

    Trang này:
    - hiển thị toàn bộ câu hỏi đã chọn;
    - hiển thị đáp án cho giáo viên kiểm tra;
    - có nút HỦY để xóa đề nháp và quay lại chọn câu;
    - có nút XUẤT BẢN để chuyển published=1.
    """
    ensure_exam_question_media(exam_id)
    with db() as c:
        exam, source = get_owned_bank_exam(
            c,
            exam_id,
            session["user_id"]
        )
        if not exam:
            abort(404)

        question_count = c.execute(
            "SELECT COUNT(*) FROM questions WHERE exam_id=?",
            (exam_id,)
        ).fetchone()[0]

    return render_template(
        "bank_exam_preview.html",
        exam=exam,
        source=source,
        questions=get_exam_questions(exam_id),
        question_count=question_count
    )


@app.post("/teacher/question-bank/exam/<int:exam_id>/cancel")
@login_required("teacher")
def bank_exam_cancel(exam_id):
    """
    HỦY BẢN NHÁP:
    - xóa đề vừa tạo khỏi exams/questions/options;
    - quay lại đúng ngân hàng để giáo viên chọn câu mới.
    """
    with db() as c:
        exam, source = get_owned_bank_exam(
            c,
            exam_id,
            session["user_id"]
        )
        if not exam:
            abort(404)

        source_id = source["id"]

        qids = [
            r["id"] for r in c.execute(
                "SELECT id FROM questions WHERE exam_id=?",
                (exam_id,)
            ).fetchall()
        ]

        if qids:
            marks = ",".join("?" for _ in qids)

            # Xóa options của các câu
            c.execute(
                f"DELETE FROM options WHERE question_id IN ({marks})",
                qids
            )

            # Các bảng phụ có thể không tồn tại ở DB cũ.
            for table, col in [
                ("answers", "question_id"),
                ("student_answers", "question_id"),
            ]:
                try:
                    c.execute(
                        f"DELETE FROM {table} WHERE {col} IN ({marks})",
                        qids
                    )
                except sqlite3.OperationalError:
                    pass

        # Xóa các dữ liệu làm bài nếu có nhưng đề nháp thường chưa có.
        for sql, params in [
            ("DELETE FROM attempts WHERE exam_id=?", (exam_id,)),
            ("DELETE FROM questions WHERE exam_id=?", (exam_id,)),
            ("DELETE FROM exams WHERE id=?", (exam_id,))
        ]:
            try:
                c.execute(sql, params)
            except sqlite3.OperationalError:
                pass

    flash("Đã hủy đề nháp. Bạn có thể chọn lại các câu khác.")
    return redirect(
        url_for("question_bank_detail", source_id=source_id)
    )


@app.route(
    "/teacher/question-bank/exam/<int:exam_id>/publish",
    methods=["GET","POST"]
)
@login_required("teacher")
def bank_exam_publish(exam_id):
    """
    ============================================================
    BƯỚC 3 — XUẤT BẢN (V5.21.8)
    ============================================================
    Không còn draft_checked / bước kiểm tra trung gian.

    Chỉ cần:
    - đề thuộc đúng giáo viên;
    - đề có ít nhất 1 câu.
    Sau đó:
        published = 1
    ============================================================
    """
    try:
        with db() as c:
            exam, source = get_owned_bank_exam(
                c,
                exam_id,
                session["user_id"]
            )

            if not exam:
                abort(404)

            n = c.execute(
                "SELECT COUNT(*) FROM questions WHERE exam_id=?",
                (exam_id,)
            ).fetchone()[0]

            if n <= 0:
                flash("Không thể xuất bản vì đề chưa có câu hỏi.")
                return redirect(
                    url_for("bank_exam_preview", exam_id=exam_id)
                )

            c.execute(
                "UPDATE exams SET published=1 WHERE id=?",
                (exam_id,)
            )

        flash(
            "Bước 3 hoàn tất: đề đã xuất bản. "
            "Học sinh có thể vào mục Làm bài."
        )
        return redirect(
            url_for("bank_exam_preview", exam_id=exam_id)
        )

    except Exception as e:
        flash("Không xuất bản được đề: " + str(e))
        return redirect(
            url_for("bank_exam_preview", exam_id=exam_id)
        )


def _materialize_practice_exam(c, source_id):
    """Create one hidden exam mirror for a bank source and reuse it later."""
    ensure_question_bank_schema(c)
    ensure_practice_exam_schema(c)

    source = c.execute("""
        SELECT s.*,COALESCE(u.full_name,'Giáo viên') AS teacher_name
        FROM bank_sources s
        LEFT JOIN users u ON u.id=s.uploaded_by
        WHERE s.id=?
    """, (source_id,)).fetchone()
    if not source:
        raise LookupError("Ngân hàng đề không còn tồn tại.")

    existing = c.execute("""
        SELECT * FROM exams
        WHERE is_practice=1 AND practice_source_id=?
        LIMIT 1
    """, (source_id,)).fetchone()
    if existing:
        return existing["id"]

    bank_rows = c.execute("""
        SELECT * FROM bank_questions
        WHERE source_id=?
        ORDER BY section,section_order,original_order,id
    """, (source_id,)).fetchall()
    if not bank_rows:
        raise ValueError("Ngân hàng này chưa có câu hỏi để luyện tập.")

    counts = {1:0,2:0,3:0,4:0}
    for row in bank_rows:
        section = int(row["section"] or 1)
        counts[section if section in counts else 1] += 1

    points, _ = normalize_bank_exam_points(
        [counts[1], counts[2], counts[3], counts[4]],
        counts,
    )
    cur = c.execute("""
        INSERT INTO exams(
            title,source_file,source_stored_name,question_count,published,
            is_practice,practice_source_id,duration_minutes,
            part1_points,part2_points,part3_points,part4_points,
            p2_item_a,p2_item_b,p2_item_c,p2_item_d,
            part2_scoring_mode,anti_cheat,max_violations
        ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (
        "Luyện đề: " + source["title"],
        "practice_bank:" + str(source_id),
        source["stored_name"] or "",
        len(bank_rows), 1, 1, source_id, 90,
        *points,
        0.25,0.25,0.25,0.25,
        1, 1, 3,
    ))
    exam_id = cur.lastrowid
    c.execute("UPDATE exams SET teacher_id=? WHERE id=?",(source["uploaded_by"],exam_id))

    section_counters = {1:0,2:0,3:0,4:0}
    for global_order, bank_question in enumerate(bank_rows, 1):
        section = int(bank_question["section"] or 1)
        if section not in section_counters:
            section = 1
        section_counters[section] += 1
        stem_json = bank_question["blocks_json"] or "[]"
        try:
            stem_blocks = json.loads(stem_json)
        except Exception:
            stem_blocks = []
            stem_json = "[]"

        question_id = c.execute("""
            INSERT INTO questions(
                exam_id,original_order,section_order,section,qtype,
                blocks_json,stem_json,solution_json,image_url,
                correct_option,short_answer,tf_answers_json
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            exam_id,global_order,section_counters[section],section,
            bank_question["qtype"],stem_json,stem_json,
            bank_question["solution_json"] or "[]",
            first_image(stem_blocks),bank_question["correct_option"],
            bank_question["short_answer"],
            bank_question["tf_answers_json"] or "[]",
        )).lastrowid

        bank_options = c.execute("""
            SELECT * FROM bank_options
            WHERE bank_question_id=?
            ORDER BY option_key
        """, (bank_question["id"],)).fetchall()
        for option in bank_options:
            option_json = option["blocks_json"] or "[]"
            try:
                option_blocks = json.loads(option_json)
            except Exception:
                option_blocks = []
                option_json = "[]"
            c.execute("""
                INSERT INTO options(
                    question_id,option_key,blocks_json,content_json,
                    image_url,correct_tf
                ) VALUES(?,?,?,?,?,?)
            """, (
                question_id,option["option_key"],option_json,option_json,
                first_image(option_blocks),option["correct_tf"],
            ))

    return exam_id


@app.get("/student/practice")
@login_required("student")
def student_practice():
    """Show bank folders first; sources are listed only inside a folder."""
    with db() as c:
        ensure_question_bank_schema(c)
        ensure_practice_exam_schema(c)
        folders = c.execute("""
            SELECT f.*,COALESCE(u.full_name,'Giáo viên') AS teacher_name,
                   (SELECT COUNT(*) FROM bank_sources s
                    WHERE s.folder_id=f.id) AS item_count
            FROM bank_folders f
            JOIN teacher_students ts ON ts.teacher_id=f.created_by AND ts.student_id=?
            LEFT JOIN users u ON u.id=f.created_by
            ORDER BY f.name COLLATE NOCASE,f.id
        """,(session["user_id"],)).fetchall()
        unfiled_count = c.execute(
            "SELECT COUNT(*) FROM bank_sources b JOIN teacher_students ts ON ts.teacher_id=b.uploaded_by WHERE b.folder_id IS NULL AND ts.student_id=?",(session["user_id"],)
        ).fetchone()[0]
    return render_template(
        "student_practice.html",
        folders=folders,
        unfiled_count=unfiled_count,
    )


def _practice_sources_in_folder(c, student_id, folder_id=None, unfiled=False):
    condition = "s.folder_id IS NULL" if unfiled else "s.folder_id=?"
    params = [student_id,student_id,student_id]
    if not unfiled:
        params.append(folder_id)
    return c.execute(f"""
        SELECT s.*,COALESCE(u.full_name,'Giáo viên') AS teacher_name,
               (SELECT COUNT(*) FROM bank_questions q
                WHERE q.source_id=s.id) AS question_count,
               e.id AS practice_exam_id,
               (SELECT COUNT(*) FROM student_exams se
                WHERE se.exam_id=e.id AND se.student_id=?
                  AND se.submitted_at IS NOT NULL) AS attempt_count,
               (SELECT se.score FROM student_exams se
                WHERE se.exam_id=e.id AND se.student_id=?
                  AND se.submitted_at IS NOT NULL
                ORDER BY se.id DESC LIMIT 1) AS last_score
        FROM bank_sources s
        LEFT JOIN users u ON u.id=s.uploaded_by
        LEFT JOIN exams e
          ON e.is_practice=1 AND e.practice_source_id=s.id
        JOIN teacher_students ts ON ts.teacher_id=s.uploaded_by AND ts.student_id=?
        WHERE {condition}
        ORDER BY s.id DESC
    """,params).fetchall()


def _format_attempt_duration(total_seconds):
    """Format a practice duration without relying on browser locale."""
    seconds = max(0, int(total_seconds or 0))
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    if hours:
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    return f"{minutes:02d}:{seconds:02d}"


def _practice_grade_from_folder(folder_name):
    """Infer the practice grade while retaining the teacher's folder name."""
    name = normalize_text(folder_name)
    if not name:
        return "Chưa phân loại"
    match = re.search(
        r"\b(?:lớp|lop|khối|khoi)\s*[-_:]?\s*(1[0-2]|[6-9])\b",
        name,
        re.I,
    )
    if not match:
        match = re.search(r"\b(1[0-2]|[6-9])\b", name)
    return f"Lớp {match.group(1)}" if match else name


def _practice_monitor_payload(c, teacher_id):
    """Return only practice attempts belonging to this teacher's bank."""
    ensure_question_bank_schema(c)
    ensure_practice_exam_schema(c)
    rows = c.execute("""
        SELECT se.id AS attempt_id,se.started_at,se.last_activity_at,
               se.submitted_at,se.score,COALESCE(se.violations,0) AS violations,
               s.id AS source_id,s.title AS source_title,
               COALESCE(f.name,'Chưa phân loại') AS folder_name,
               COALESCE(u.full_name,'Học sinh') AS student_name,
               COALESCE(u.student_code,'') AS student_code,
               COALESCE(ts.class_name,u.class_name,'') AS class_name,
               COALESCE(e.question_count,0) AS question_count,
               COUNT(sa.id) AS answered_count,
               CAST(MAX(0,ROUND(
                   (julianday(COALESCE(se.submitted_at,CURRENT_TIMESTAMP))-
                    julianday(se.started_at))*86400
               )) AS INTEGER) AS elapsed_seconds,
               CASE
                   WHEN se.submitted_at IS NULL
                    AND COALESCE(se.last_activity_at,se.started_at)
                        >= datetime('now','-35 seconds')
                   THEN 1 ELSE 0
               END AS is_live,
               strftime('%H:%M:%S %d/%m/%Y',se.started_at,'localtime')
                   AS started_display,
               strftime('%H:%M:%S %d/%m/%Y',se.submitted_at,'localtime')
                   AS submitted_display
        FROM student_exams se
        JOIN exams e ON e.id=se.exam_id AND e.is_practice=1
        JOIN bank_sources s ON s.id=e.practice_source_id
        LEFT JOIN bank_folders f ON f.id=s.folder_id
        LEFT JOIN users u ON u.id=se.student_id
        LEFT JOIN teacher_students ts ON ts.teacher_id=s.uploaded_by AND ts.student_id=se.student_id
        LEFT JOIN student_answers sa ON sa.student_exam_id=se.id
        WHERE s.uploaded_by=?
        GROUP BY se.id
        ORDER BY CASE WHEN se.submitted_at IS NULL THEN 0 ELSE 1 END,
                 COALESCE(se.last_activity_at,se.submitted_at,se.started_at) DESC,
                 se.id DESC
        LIMIT 300
    """, (teacher_id,)).fetchall()

    attempts = []
    for row in rows:
        submitted = bool(row["submitted_at"])
        live = bool(row["is_live"]) and not submitted
        attempts.append({
            "attempt_id": row["attempt_id"],
            "student_name": row["student_name"],
            "student_code": row["student_code"],
            "class_name": row["class_name"],
            "source_title": row["source_title"],
            "folder_name": row["folder_name"],
            "practice_grade": _practice_grade_from_folder(row["folder_name"]),
            "answered_count": int(row["answered_count"] or 0),
            "question_count": int(row["question_count"] or 0),
            "score": round(float(row["score"] or 0), 2) if submitted else None,
            "violations": int(row["violations"] or 0),
            "duration": _format_attempt_duration(row["elapsed_seconds"]),
            "started_at": row["started_display"] or "",
            "submitted_at": row["submitted_display"] or "",
            "status": "Đang luyện" if live else (
                "Đã nộp" if submitted else "Tạm ngắt"
            ),
            "is_live": live,
        })

    completed = [item for item in attempts if item["score"] is not None]
    scores = [item["score"] for item in completed]
    payload = {
        "active": [item for item in attempts if item["is_live"]],
        "attempts": attempts,
        "summary": {
            "active_count": sum(1 for item in attempts if item["is_live"]),
            "student_count": len({
                (item["student_code"] or item["student_name"])
                for item in attempts
            }),
            "completed_count": len(completed),
            "average_score": round(sum(scores) / len(scores), 2) if scores else 0,
            "total_violations": sum(item["violations"] for item in attempts),
        },
        "updated_at": c.execute(
            "SELECT strftime('%H:%M:%S %d/%m/%Y','now','localtime')"
        ).fetchone()[0],
    }
    return payload


@app.get("/teacher/question-bank/practice-monitor")
@login_required("teacher")
def practice_monitor():
    with db() as c:
        payload = _practice_monitor_payload(c, session["user_id"])
    return render_template(
        "practice_monitor.html",
        monitor_data=payload,
        monitor_api=url_for("practice_monitor_data"),
    )


@app.get("/teacher/question-bank/practice-monitor/data")
@login_required("teacher")
def practice_monitor_data():
    with db() as c:
        payload = _practice_monitor_payload(c, session["user_id"])
    return jsonify(payload)


@app.get("/teacher/question-bank/practice-monitor.xlsx")
@login_required("teacher")
def practice_monitor_export_xlsx():
    """Xuất đầy đủ danh sách và thống kê học sinh luyện đề của giáo viên."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    with db() as c:
        ensure_question_bank_schema(c)
        ensure_practice_exam_schema(c)
        rows=c.execute("""
            SELECT se.id AS attempt_id,se.started_at,se.submitted_at,se.score,
                   COALESCE(se.violations,0) AS violations,
                   s.title AS source_title,
                   COALESCE(f.name,'Chưa phân loại') AS folder_name,
                   COALESCE(u.full_name,'Học sinh') AS student_name,
                   COALESCE(u.student_code,'') AS student_code,
                   COALESCE(u.class_name,'') AS class_name,
                   COALESCE(e.question_count,0) AS question_count,
                   COUNT(sa.id) AS answered_count,
                   CAST(MAX(0,ROUND(
                       (julianday(COALESCE(se.submitted_at,CURRENT_TIMESTAMP))-
                        julianday(se.started_at))*86400
                   )) AS INTEGER) AS elapsed_seconds,
                   CASE
                       WHEN se.submitted_at IS NULL
                        AND COALESCE(se.last_activity_at,se.started_at)
                            >= datetime('now','-35 seconds')
                       THEN 1 ELSE 0
                   END AS is_live,
                   strftime('%H:%M:%S %d/%m/%Y',se.started_at,'localtime') AS started_display,
                   strftime('%H:%M:%S %d/%m/%Y',se.submitted_at,'localtime') AS submitted_display
            FROM student_exams se
            JOIN exams e ON e.id=se.exam_id AND e.is_practice=1
            JOIN bank_sources s ON s.id=e.practice_source_id
            LEFT JOIN bank_folders f ON f.id=s.folder_id
            LEFT JOIN users u ON u.id=se.student_id
            LEFT JOIN student_answers sa ON sa.student_exam_id=se.id
            WHERE s.uploaded_by=?
            GROUP BY se.id
            ORDER BY u.class_name COLLATE NOCASE,u.full_name COLLATE NOCASE,se.id
        """,(session["user_id"],)).fetchall()

    wb=Workbook()
    ws=wb.active
    ws.title="Danh sách luyện đề"
    headers=[
        "STT","Mã học sinh","Họ và tên","Lớp","Đề luyện","Thư mục",
        "Bắt đầu","Nộp bài","Thời gian làm","Tiến độ","Điểm",
        "Vi phạm","Trạng thái"
    ]
    ws.append(headers)

    student_summary={}
    for idx,row in enumerate(rows,1):
        submitted=bool(row["submitted_at"])
        live=bool(row["is_live"]) and not submitted
        status="Đang luyện" if live else ("Đã nộp" if submitted else "Tạm ngắt")
        score=round(float(row["score"] or 0),2) if submitted else None
        key=(row["student_code"] or "",row["student_name"],row["class_name"] or "")
        stat=student_summary.setdefault(key,{"attempts":0,"completed":0,"scores":[],"violations":0,"last_score":None})
        stat["attempts"]+=1
        stat["violations"]+=int(row["violations"] or 0)
        if submitted:
            stat["completed"]+=1
            stat["scores"].append(score)
            stat["last_score"]=score
        ws.append([
            idx,row["student_code"],row["student_name"],row["class_name"],
            row["source_title"],row["folder_name"],row["started_display"],
            row["submitted_display"] or "",_format_attempt_duration(row["elapsed_seconds"]),
            f"{int(row['answered_count'] or 0)}/{int(row['question_count'] or 0)}",
            score if score is not None else "",int(row["violations"] or 0),status
        ])

    ws2=wb.create_sheet("Tổng hợp học sinh")
    ws2.append([
        "STT","Mã học sinh","Họ và tên","Lớp","Số lượt luyện","Đã nộp",
        "Điểm gần nhất","Điểm cao nhất","Điểm trung bình","Tổng vi phạm"
    ])
    for idx,(key,stat) in enumerate(sorted(student_summary.items(),key=lambda x:(x[0][2],x[0][1])),1):
        code,name,class_name=key
        scores=stat["scores"]
        ws2.append([
            idx,code,name,class_name,stat["attempts"],stat["completed"],
            stat["last_score"] if stat["last_score"] is not None else "",
            max(scores) if scores else "",
            round(sum(scores)/len(scores),2) if scores else "",
            stat["violations"]
        ])

    thin=Side(style="thin",color="808080")
    for sheet in (ws,ws2):
        sheet.freeze_panes="A2"
        sheet.auto_filter.ref=sheet.dimensions
        for cell in sheet[1]:
            cell.font=Font(bold=True)
            cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        for row_cells in sheet.iter_rows():
            for cell in row_cells:
                cell.alignment=Alignment(vertical="center",wrap_text=True)
                cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
        widths=[]
        for col in range(1,sheet.max_column+1):
            max_len=max(len(str(sheet.cell(r,col).value or "")) for r in range(1,min(sheet.max_row,200)+1))
            widths.append(min(max(max_len+2,10),34))
            sheet.column_dimensions[get_column_letter(col)].width=widths[-1]

    export_dir=BASE/"exports"
    export_dir.mkdir(exist_ok=True)
    output=export_dir/f"Danh_sach_hoc_sinh_luyen_de_{session['user_id']}.xlsx"
    wb.save(output)
    return send_file(
        output,as_attachment=True,
        download_name="Danh_sach_hoc_sinh_luyen_de.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.post("/teacher/question-bank/practice-monitor/<int:attempt_id>/delete")
@login_required("teacher")
def delete_practice_attempt(attempt_id):
    """Delete one owned practice attempt, never the student's account."""
    with db() as c:
        ensure_question_bank_schema(c)
        ensure_practice_exam_schema(c)
        owned = c.execute("""
            SELECT se.id,se.submitted_at,se.last_activity_at,se.started_at
            FROM student_exams se
            JOIN exams e ON e.id=se.exam_id AND e.is_practice=1
            JOIN bank_sources s ON s.id=e.practice_source_id
            WHERE se.id=? AND s.uploaded_by=?
        """, (attempt_id, session["user_id"])).fetchone()
        if not owned:
            abort(404)

        # Không xóa một lượt đang hoạt động để tránh làm mất bài học sinh.
        is_live = c.execute("""
            SELECT CASE
                WHEN ? IS NULL AND COALESCE(?,?) >= datetime('now','-35 seconds')
                THEN 1 ELSE 0 END
        """, (
            owned["submitted_at"], owned["last_activity_at"], owned["started_at"]
        )).fetchone()[0]
        if is_live:
            return jsonify({
                "ok": False,
                "error": "Không thể xóa khi học sinh đang luyện đề."
            }), 409

        c.execute(
            "DELETE FROM anti_cheat_logs WHERE student_exam_id=?", (attempt_id,)
        )
        c.execute(
            "DELETE FROM student_answers WHERE student_exam_id=?", (attempt_id,)
        )
        c.execute("DELETE FROM student_exams WHERE id=?", (attempt_id,))
        c.commit()
        payload = _practice_monitor_payload(c, session["user_id"])

    if request.accept_mimetypes.best == "application/json":
        return jsonify({"ok": True, "monitor": payload})
    flash("Đã xóa lượt luyện; tài khoản học sinh vẫn được giữ nguyên.")
    return redirect(url_for("practice_monitor"))


@app.get("/student/practice/folder/<int:folder_id>")
@login_required("student")
def student_practice_folder(folder_id):
    with db() as c:
        ensure_question_bank_schema(c)
        ensure_practice_exam_schema(c)
        folder = c.execute("""
            SELECT f.*,COALESCE(u.full_name,'Giáo viên') AS teacher_name
            FROM bank_folders f
            LEFT JOIN users u ON u.id=f.created_by
            WHERE f.id=?
        """,(folder_id,)).fetchone()
        if not folder:
            abort(404)
        sources = _practice_sources_in_folder(
            c,session["user_id"],folder_id=folder_id
        )
    return render_template(
        "student_practice_folder.html",
        folder=folder,
        sources=sources,
        is_unfiled=False,
    )


@app.get("/student/practice/unfiled")
@login_required("student")
def student_unfiled_practice():
    with db() as c:
        ensure_question_bank_schema(c)
        ensure_practice_exam_schema(c)
        sources = _practice_sources_in_folder(
            c,session["user_id"],unfiled=True
        )
    return render_template(
        "student_practice_folder.html",
        folder={"id":None,"name":"Đề luyện chưa phân loại","teacher_name":""},
        sources=sources,
        is_unfiled=True,
    )


@app.post("/student/practice/<int:source_id>/start")
@login_required("student")
def start_practice(source_id):
    try:
        ensure_bank_question_media(source_id)
        with db() as c:
            exam_id = _materialize_practice_exam(c, source_id)
    except (LookupError, ValueError) as error:
        flash(str(error))
        return redirect(url_for("student_practice"))
    except Exception as error:
        flash("Không thể mở đề luyện tập: " + str(error))
        return redirect(url_for("student_practice"))
    return redirect(url_for("take_exam", exam_id=exam_id))



@app.post("/teacher/question-bank/<int:source_id>/delete")
@login_required("teacher")
def delete_question_bank(source_id):
    with db() as c:
        ensure_question_bank_schema(c)
        ensure_practice_exam_schema(c)
        source = c.execute("""
            SELECT * FROM bank_sources
            WHERE id=? AND uploaded_by=?
        """,(source_id,session["user_id"])).fetchone()
        if not source:
            abort(404)
        folder_id = source["folder_id"]

        # The hidden mirror and all its practice attempts are no longer valid
        # after the teacher removes the source bank.
        c.execute(
            "DELETE FROM exams "
            "WHERE is_practice=1 AND practice_source_id=?",
            (source_id,),
        )

        qids = [
            r["id"] for r in c.execute(
                "SELECT id FROM bank_questions WHERE source_id=?",
                (source_id,)
            ).fetchall()
        ]

        if qids:
            marks = ",".join("?" for _ in qids)
            c.execute(
                f"DELETE FROM bank_options "
                f"WHERE bank_question_id IN ({marks})",
                qids
            )

        c.execute("DELETE FROM bank_questions WHERE source_id=?",(source_id,))
        c.execute("DELETE FROM bank_sources WHERE id=?",(source_id,))

    flash("Đã xóa ngân hàng đề.")
    if folder_id is not None:
        return redirect(url_for("teacher_bank_folder",folder_id=folder_id))
    return redirect(url_for("teacher_unfiled_bank"))


@app.get("/student")
@login_required("student")
def student_dashboard():
    """
    V5.12.5 — HỌC SINH CHỈ THẤY ĐỀ HỢP LỆ

    Điều kiện hiển thị:
    - đề vẫn còn tồn tại trong bảng exams;
    - giáo viên đã xuất bản: published = 1.

    Khi giáo viên xóa đề, bản ghi exams bị xóa vật lý nên đề tự
    biến mất khỏi danh sách học sinh ngay lần tải trang tiếp theo.
    """
    with db() as c:
        ensure_practice_exam_schema(c)
        exams=c.execute(
            "SELECT e.* FROM exams e "
            "JOIN teacher_students ts ON ts.teacher_id=e.teacher_id AND ts.student_id=? "
            "WHERE e.published=1 AND COALESCE(e.is_practice,0)=0 "
            "AND (NOT EXISTS(SELECT 1 FROM exam_invites ei WHERE ei.exam_id=e.id) "
            " OR EXISTS(SELECT 1 FROM exam_students es WHERE es.exam_id=e.id AND es.student_id=?)) "
            "ORDER BY e.id DESC",(session["user_id"],session["user_id"])
        ).fetchall()

    return render_template(
        "student_dashboard.html",
        exams=exams
    )

@app.get("/student/exam/<int:exam_id>")
@login_required("student")
def take_exam(exam_id):
    # KHỐI 5 — HỌC SINH: chỉ mở đề đã xuất bản; phiên làm bài được tạo tự động.
    with db() as c:
        # V5.12.5: kể cả học sinh còn URL cũ, đề đã xóa hoặc
        # chưa xuất bản đều không thể mở.
        pub=c.execute(
            "SELECT id FROM exams WHERE id=? AND published=1",
            (exam_id,)
        ).fetchone()
    if not pub:
        flash("Đề chưa được giáo viên xuất bản hoặc không tồn tại.")
        return redirect(url_for("student_dashboard"))
    ensure_exam_question_media(exam_id)
    ensure_exam_latex_only(exam_id)
    se=create_student_exam(exam_id,session["user_id"])
    if not se:
        flash("Không thể tạo phiên làm bài. Hãy báo giáo viên kiểm tra đề đã xuất bản.")
        return redirect(url_for("student_dashboard"))
    order=json.loads(se["question_order_json"]); omap=json.loads(se["option_map_json"])
    with db() as c:
        exam=c.execute("SELECT * FROM exams WHERE id=?",(exam_id,)).fetchone(); data=[]
        display_counters={1:0,2:0,3:0,4:0}
        last_section=None
        for qid in order:
            q=c.execute("SELECT * FROM questions WHERE id=?",(qid,)).fetchone()
            if not q: continue
            opts={o["option_key"]:o for o in c.execute("SELECT * FROM options WHERE question_id=?",(qid,)).fetchall()}
            display=[]
            for key in omap.get(str(qid),[]):
                o=opts[key]; display.append({"key":key,"blocks":json.loads((o["blocks_json"] or (o["content_json"] if "content_json" in o.keys() else "[]") or "[]"))})
            ans=c.execute("SELECT selected_original_key,answer_json FROM student_answers WHERE student_exam_id=? AND question_id=?",(se["id"],qid)).fetchone()
            qblocks = q["blocks_json"] or (q["stem_json"] if "stem_json" in q.keys() else "[]") or "[]"
            selected = None
            if ans:
                if q["qtype"]=="mcq": selected=ans["selected_original_key"]
                elif ans["answer_json"]:
                    try: selected=json.loads(ans["answer_json"])
                    except: selected=None
            display_counters[q["section"]]=display_counters.get(q["section"],0)+1
            data.append({"id":q["id"],"order":q["original_order"],"section_order":display_counters[q["section"]],"section":q["section"],"qtype":q["qtype"],"blocks":json.loads(qblocks),"options":display,"selected":selected})
    return render_template("ai_exam.html",exam=exam,student_exam=se,questions=data)

# ============================================================
# V5.11.1 — CHẤM ĐIỂM TỪNG CÂU PHẦN II
# Hai phương án áp dụng ĐỘC LẬP cho từng câu Phần II.
# Mỗi câu có 4 ý và tối đa 1,00 điểm.
# PA1: 1/2/3/4 ý đúng = 0,25/0,50/0,75/1,00.
# PA2: 1/2/3/4 ý đúng = 0,10/0,25/0,50/1,00.
# Nếu Phần II có nhiều câu, điểm Phần II = tổng điểm của từng câu.
# ============================================================
def score_part2_question(statement_correct, mode=1):
    correct_count = sum(bool(x) for x in list(statement_correct)[:4])
    if mode == 2:
        return {0:0.00,1:0.10,2:0.25,3:0.50,4:1.00}[correct_count]
    return correct_count * 0.25

# ============================================================
# KHỐI ĐIỂM — TÍNH ĐIỂM TỪNG CÂU
# Phần I và III: điểm mỗi câu bằng nhau.
# Phần II: MỖI CÂU tối đa 1 điểm. Chọn một trong 2 phương án
# chấm cho từng câu; không chia 1 điểm của cả Phần II cho toàn bộ câu.
# ============================================================
def question_score(c, exam, q, answer):
    sec=q['section']; qtype=q['qtype']
    if sec==1:
        n=c.execute('SELECT COUNT(*) FROM questions WHERE exam_id=? AND section=1',(exam['id'],)).fetchone()[0]
        correct=int(str(answer)==str(q['correct_option'])) if q['correct_option'] else 0
        return round(float(exam['part1_points'])/n,4) if n and correct else 0
    if sec==3:
        n=c.execute('SELECT COUNT(*) FROM questions WHERE exam_id=? AND section=3',(exam['id'],)).fetchone()[0]
        got=normalize_text(str(answer or '')).strip().lower()
        target=normalize_text(q['short_answer'] or '').strip().lower()
        correct=bool(target) and got==target
        return round(float(exam['part3_points'])/n,4) if n and correct else 0
    if sec==2:
        # V5.11.1 — chấm độc lập TỪNG CÂU theo phương án của đề.
        vals={str(k):bool(v) for k,v in (answer or {}).items()}
        rows=c.execute('SELECT option_key,correct_tf FROM options WHERE question_id=? ORDER BY option_key',(q['id'],)).fetchall()
        statement_correct=[]
        for r in rows[:4]:
            statement_correct.append(bool(r['correct_tf'])==vals.get(r['option_key']))
        return round(score_part2_question(statement_correct, int(exam['part2_scoring_mode'] or 1)),4)
    return 0

@app.post("/student/exam/<int:student_exam_id>/answer")
@login_required("student")
def save_answer(student_exam_id):
    data=request.get_json(silent=True) or {}; qid=int(data.get("question_id")); answer=data.get("answer")
    with db() as c:
        ensure_practice_exam_schema(c)
        se=c.execute("SELECT * FROM student_exams WHERE id=? AND student_id=? AND submitted_at IS NULL",(student_exam_id,session["user_id"])).fetchone()
        if not se: abort(403)
        q=c.execute("SELECT * FROM questions WHERE id=? AND exam_id=?",(qid,se["exam_id"])).fetchone()
        if not q: return jsonify(ok=False),400
        exam=c.execute("SELECT * FROM exams WHERE id=?",(se["exam_id"],)).fetchone()
        correct=None
        if q["qtype"]=="mcq":
            allowed=json.loads(se["option_map_json"]).get(str(qid),[])
            if str(answer) not in allowed: return jsonify(ok=False,error="Phương án không hợp lệ"),400
            correct=int(str(answer)==str(q["correct_option"])) if q["correct_option"] else None
        elif q["qtype"]=="tf":
            vals={str(k):bool(v) for k,v in (answer or {}).items()}; rows=c.execute("SELECT option_key,correct_tf FROM options WHERE question_id=?",(qid,)).fetchall()
            correct=int(bool(rows) and all(bool(r["correct_tf"])==vals.get(r["option_key"]) for r in rows))
        elif q["qtype"]=="short":
            got=normalize_text(str(answer or "")).strip()
            if not re.fullmatch(r"[-0-9,]+", got):
                return jsonify(ok=False,error="Câu trả lời ngắn chỉ cho phép chữ số, dấu - và dấu ,"),400
            target=normalize_text(q["short_answer"] or "").lower(); correct=int(bool(target) and got.lower()==target)
        score=question_score(c,exam,q,answer)
        c.execute("INSERT INTO student_answers(student_exam_id,question_id,answer_json,is_correct,score,selected_original_key) VALUES(?,?,?,?,?,?) ON CONFLICT(student_exam_id,question_id) DO UPDATE SET answer_json=excluded.answer_json,is_correct=excluded.is_correct,score=excluded.score,selected_original_key=excluded.selected_original_key",(student_exam_id,qid,jdump(answer),correct,score,str(answer) if q["qtype"]=="mcq" else None))
        c.execute(
            "UPDATE student_exams SET last_activity_at=CURRENT_TIMESTAMP WHERE id=?",
            (student_exam_id,),
        )
    return jsonify(ok=True,score=score)


@app.post("/student/exam/<int:student_exam_id>/heartbeat")
@login_required("student")
def practice_heartbeat(student_exam_id):
    """Refresh live status only for an open practice attempt."""
    with db() as c:
        ensure_practice_exam_schema(c)
        attempt = c.execute("""
            SELECT se.id
            FROM student_exams se
            JOIN exams e ON e.id=se.exam_id
            WHERE se.id=? AND se.student_id=?
              AND se.submitted_at IS NULL AND e.is_practice=1
        """, (student_exam_id, session["user_id"])).fetchone()
        if not attempt:
            return jsonify(ok=True, ignored=True)
        c.execute(
            "UPDATE student_exams SET last_activity_at=CURRENT_TIMESTAMP WHERE id=?",
            (student_exam_id,),
        )
    return jsonify(ok=True)

@app.post("/student/exam/<int:student_exam_id>/event")
@login_required("student")
def anti_event(student_exam_id):
    """
    V5.11.6 — CHỐNG GIAN LẬN THEO YÊU CẦU MỚI

    Chỉ tính 1 vi phạm khi học sinh đã bắt đầu làm bài, chưa nộp,
    và rời khỏi màn hình/trang làm bài.

    Các event khác không làm tăng violations.
    """
    data=request.get_json(silent=True) or {}
    event=normalize_text(data.get("event"))[:80]
    details=normalize_text(data.get("details"))[:300]

    with db() as c:
        ensure_practice_exam_schema(c)
        se=c.execute(
            "SELECT * FROM student_exams WHERE id=? AND student_id=? AND submitted_at IS NULL",
            (student_exam_id,session["user_id"])
        ).fetchone()
        if not se:
            # Nếu bài đã nộp thì không ghi thêm vi phạm.
            return jsonify(ok=True,ignored=True)

        c.execute(
            "UPDATE student_exams SET last_activity_at=CURRENT_TIMESTAMP WHERE id=?",
            (student_exam_id,),
        )

        exam=c.execute("SELECT * FROM exams WHERE id=?",(se["exam_id"],)).fetchone()

        # Chỉ duy nhất event này được tính là vi phạm.
        if event != "leave_exam_screen":
            return jsonify(ok=True,violations=int(se["violations"] or 0),ignored=True)

        violations=int(se["violations"] or 0)+1
        c.execute(
            "INSERT INTO anti_cheat_logs(student_exam_id,event_type,details) VALUES(?,?,?)",
            (student_exam_id,event,details or "Rời màn hình/trang khi chưa nộp bài")
        )
        c.execute(
            "UPDATE student_exams "
            "SET violations=?,last_activity_at=CURRENT_TIMESTAMP WHERE id=?",
            (violations,student_exam_id)
        )

        force=bool(
            exam and exam["anti_cheat"] and violations>=int(exam["max_violations"] or 0)
        )

    return jsonify(ok=True,violations=violations,force_submit=force)

@app.post("/student/exam/<int:student_exam_id>/submit")
@login_required("student")
def submit_exam(student_exam_id):
    with db() as c:
        ensure_practice_exam_schema(c)
        se=c.execute("SELECT * FROM student_exams WHERE id=? AND student_id=? AND submitted_at IS NULL",(student_exam_id,session["user_id"])).fetchone()
        if not se: abort(403)
        e=c.execute("SELECT * FROM exams WHERE id=?",(se["exam_id"],)).fetchone(); total=0
        # Điểm I–III đã được tính và lưu theo từng câu ngay khi học sinh trả lời.
        total=float(c.execute("SELECT COALESCE(SUM(score),0) FROM student_answers WHERE student_exam_id=?",(student_exam_id,)).fetchone()[0])
        # V5.11.6: bấm Nộp bài không tạo thêm event/log vi phạm.
        # Part IV is teacher graded later; current score is automatically scored I-III.
        total=round(total,2)
        c.execute("UPDATE student_exams SET submitted_at=CURRENT_TIMESTAMP,last_activity_at=CURRENT_TIMESTAMP,score=? WHERE id=?",(total,student_exam_id))
        # Luyện đề chấm giống bài kiểm tra nhưng không ghi vào bảng điểm lớp.
        if not int(e["is_practice"] or 0):
            update_roster_scores(e["id"])
    return jsonify(ok=True,score=total)

@app.route("/teacher/exam/<int:exam_id>/grade", methods=["GET","POST"])
@login_required("teacher")
def grade_exam(exam_id):
    with db() as c:
        exam=c.execute("SELECT * FROM exams WHERE id=?",(exam_id,)).fetchone()
        if not exam: abort(404)
        if request.method=="POST":
            rows=c.execute("SELECT se.id,se.score FROM student_exams se WHERE se.exam_id=? AND se.submitted_at IS NOT NULL",(exam_id,)).fetchall()
            for r in rows:
                raw=request.form.get(f"part4_{r['id']}","0").replace(',','.')
                try: p4=max(0,min(float(raw),float(exam['part4_points'])))
                except: p4=0
                total=round(float(r['score'] or 0)+p4,2)
                c.execute("UPDATE student_exams SET score=? WHERE id=?",(total,r['id']))
            update_roster_scores(exam_id)
            flash("Đã lưu điểm Phần IV và cập nhật Excel.")
            return redirect(url_for("grade_exam",exam_id=exam_id))
        rows=c.execute("SELECT se.id,se.score,u.full_name,u.student_code FROM student_exams se JOIN users u ON u.id=se.student_id WHERE se.exam_id=? AND se.submitted_at IS NOT NULL ORDER BY u.full_name",(exam_id,)).fetchall()
    return render_template("grade.html",exam=exam,rows=rows)


# ============================================================
# KHỐI 8 — KẾT QUẢ THEO TỪNG ĐỀ + EXCEL
# Giáo viên bấm "Xem kết quả & Excel" để xem điểm từng học sinh,
# số lỗi gian lận, chi tiết lỗi và tải Excel riêng cho đề.
# ============================================================
def _result_rows(exam_id):
    """V5.11.5 — kết quả giáo viên, an toàn với dữ liệu cũ/NULL."""
    with db() as c:
        exam=c.execute("SELECT * FROM exams WHERE id=?",(exam_id,)).fetchone()
        if not exam: return None,[]
        students=c.execute("""SELECT se.id AS student_exam_id,COALESCE(se.score,0) AS score,
            COALESCE(se.violations,0) AS violations,se.submitted_at,
            COALESCE(u.full_name,'Không xác định') AS full_name,COALESCE(u.student_code,'') AS student_code
            FROM student_exams se LEFT JOIN users u ON u.id=se.student_id
            WHERE se.exam_id=? ORDER BY COALESCE(u.full_name,''),se.id""",(exam_id,)).fetchall()
        rows=[]
        for r in students:
            parts={}
            for sec in (1,2,3):
                v=c.execute("""SELECT COALESCE(SUM(COALESCE(a.score,0)),0) FROM student_answers a
                    JOIN questions q ON q.id=a.question_id WHERE a.student_exam_id=? AND q.section=?""",(r['student_exam_id'],sec)).fetchone()[0]
                try: parts[sec]=round(float(v or 0),2)
                except: parts[sec]=0.0
            try: total=round(float(r['score'] or 0),2)
            except: total=0.0
            p4=round(max(0,total-parts[1]-parts[2]-parts[3]),2)
            logs=c.execute("SELECT event_type,details,created_at FROM anti_cheat_logs WHERE student_exam_id=? ORDER BY id",(r['student_exam_id'],)).fetchall()
            detail='; '.join(f"{x['created_at']} — {x['event_type']}: {x['details'] or ''}" for x in logs) if logs else 'Không có'
            rows.append({'student_exam_id':r['student_exam_id'],'full_name':r['full_name'] or '','student_code':r['student_code'] or '',
                         'p1':parts[1],'p2':parts[2],'p3':parts[3],'p4':p4,'score':total,'violations':int(r['violations'] or 0),
                         'anti_detail':detail,'submitted_at':r['submitted_at'] or ''})
        return exam,rows

@app.get('/teacher/exam/<int:exam_id>/results')
@login_required('teacher')
def exam_results(exam_id):
    exam,rows=_result_rows(exam_id)
    if not exam: abort(404)
    return render_template('exam_results.html',exam=exam,rows=rows)

@app.get('/teacher/exam/<int:exam_id>/results.xlsx')
@login_required('teacher')
def exam_results_excel(exam_id):
    """
    V5.11.9 — TẢI EXCEL KẾT QUẢ THEO TOÀN BỘ DANH SÁCH LỚP

    File danh sách thật của hệ thống được import_roster() lưu tại:
        instance/roster.xlsx

    Quy tắc:
    - Giữ toàn bộ học sinh trong danh sách.
    - Không làm bài: mọi điểm = 0, Lỗi = 0,
      Nhận xét = "Không làm bài".
    - Có mở bài nhưng chưa nộp:
      Nhận xét = "Chưa nộp bài".
    - Đã nộp:
      cập nhật điểm, lỗi và Nhận xét = "Đã làm bài".
    """
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font
    from flask import send_file

    with db() as c:
        exam = c.execute("SELECT * FROM exams WHERE id=?", (exam_id,)).fetchone()
        if not exam:
            abort(404)

        # V5.11.9: dùng đúng file roster mà import_roster() đang quản lý.
        roster = _teacher_roster_path(session['user_id'])
        if not roster.exists():
            flash("Chưa có file danh sách học sinh. Hãy tải danh sách Excel lên trước.")
            return redirect(url_for("exam_results", exam_id=exam_id))

        wb = load_workbook(roster)
        ws = wb[wb.sheetnames[0]]

        # Tìm hàng tiêu đề, không giả định tiêu đề ở dòng 1.
        header_row = None
        for rr in range(1, min(ws.max_row, 20) + 1):
            vals = [normalize_text(str(ws.cell(rr, cc).value or "")) for cc in range(1, ws.max_column + 1)]
            if any(v.lower() in ("mã học sinh", "mã hs") for v in vals):
                header_row = rr
                break

        if not header_row:
            flash("Không tìm thấy cột 'Mã học sinh' trong file danh sách.")
            return redirect(url_for("exam_results", exam_id=exam_id))

        def norm(v):
            return normalize_text(str(v or "")).strip().lower()

        headers = {
            norm(ws.cell(header_row, col).value): col
            for col in range(1, ws.max_column + 1)
            if ws.cell(header_row, col).value is not None
        }

        def find_col(*names):
            for name in names:
                col = headers.get(norm(name))
                if col:
                    return col
            return None

        code_col = find_col("Mã học sinh", "Mã HS")
        name_col = find_col("Họ và tên", "Họ tên", "Tên học sinh")

        if not code_col:
            flash("File danh sách thiếu cột Mã học sinh.")
            return redirect(url_for("exam_results", exam_id=exam_id))

        def ensure_col(title, *aliases):
            col = find_col(title, *aliases)
            if col:
                return col
            col = ws.max_column + 1
            ws.cell(header_row, col).value = title
            ws.cell(header_row, col).font = Font(bold=True)
            headers[norm(title)] = col
            return col

        tn_col = ensure_col("TN", "Phần I")
        ds_col = ensure_col("DS", "ĐS", "Phần II")
        tln_col = ensure_col("TLN", "Phần III")
        tl_col = ensure_col("TL", "Phần IV")
        err_col = ensure_col("Lỗi", "Vi phạm", "Số lỗi", "Lỗi gian lận")
        total_col = ensure_col("Tổng", "Tổng điểm", "Điểm")
        note_col = ensure_col("Nhận xét", "Ghi chú")

        # Khởi tạo toàn bộ danh sách: 0 điểm + Không làm bài.
        row_by_code = {}
        for rr in range(header_row + 1, ws.max_row + 1):
            code = normalize_text(str(ws.cell(rr, code_col).value or "")).strip()
            name = normalize_text(str(ws.cell(rr, name_col).value or "")).strip() if name_col else ""

            if not code and not name:
                continue

            if code:
                row_by_code[code] = rr

            for col in (tn_col, ds_col, tln_col, tl_col, err_col, total_col):
                ws.cell(rr, col).value = 0
            ws.cell(rr, note_col).value = "Không làm bài"

        # Lấy phiên mới nhất của từng học sinh trong đúng đề.
        attempts = c.execute("""
            SELECT se.*, u.student_code
            FROM student_exams se
            LEFT JOIN users u ON u.id=se.student_id
            JOIN (
                SELECT student_id, MAX(id) AS max_id
                FROM student_exams
                WHERE exam_id=?
                GROUP BY student_id
            ) latest ON latest.max_id=se.id
            WHERE se.exam_id=?
        """, (exam_id, exam_id)).fetchall()

        for se in attempts:
            code = normalize_text(str(se["student_code"] or "")).strip()
            rr = row_by_code.get(code)
            if not rr:
                continue

            parts = {}
            for sec in (1, 2, 3):
                val = c.execute("""
                    SELECT COALESCE(SUM(COALESCE(a.score,0)),0)
                    FROM student_answers a
                    JOIN questions q ON q.id=a.question_id
                    WHERE a.student_exam_id=? AND q.section=?
                """, (se["id"], sec)).fetchone()[0]
                parts[sec] = round(float(val or 0), 2)

            total = round(float(se["score"] or 0), 2)
            p4 = round(max(0.0, total - parts[1] - parts[2] - parts[3]), 2)

            ws.cell(rr, tn_col).value = parts[1]
            ws.cell(rr, ds_col).value = parts[2]
            ws.cell(rr, tln_col).value = parts[3]
            ws.cell(rr, tl_col).value = p4
            ws.cell(rr, err_col).value = int(se["violations"] or 0)
            ws.cell(rr, total_col).value = total
            ws.cell(rr, note_col).value = "Đã làm bài" if se["submitted_at"] else "Chưa nộp bài"

        # Định dạng cơ bản.
        ws.freeze_panes = f"A{header_row + 1}"
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="center")

        export_dir = BASE / "exports"
        export_dir.mkdir(exist_ok=True)
        output = export_dir / f"ket_qua_de_{exam_id}.xlsx"
        wb.save(output)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"Ket_qua_de_{exam_id}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )



# ============================================================
# V5.12.4 — XÓA ĐỀ KIỂM TRA
# ============================================================
# Chức năng dành cho giáo viên:
# - Chỉ xóa đề thuộc tài khoản giáo viên đang đăng nhập.
# - Xóa dữ liệu bài làm/đáp án/câu hỏi liên quan trước.
# - Xóa file đề gốc trên ổ đĩa nếu file vẫn tồn tại.
# - Dùng POST để tránh xóa nhầm chỉ vì mở một đường dẫn.
# ============================================================
@app.post("/teacher/exam/<int:exam_id>/delete")
@login_required("teacher")
def delete_exam(exam_id):
    """
    V5.12.6 — XÓA ĐỀ CHỈ DÀNH CHO GIÁO VIÊN

    Sau khi xóa:
    - route được bảo vệ bởi @login_required('teacher');
    - đề biến mất khỏi bảng exams;
    - học sinh không còn thấy đề trên dashboard;
    - URL cũ cũng không mở được vì take_exam yêu cầu đề còn tồn tại
      và published=1;
    - dọn dữ liệu bài làm, log gian lận, đáp án, phương án và câu hỏi.
    """
    with db() as c:
        # V5.12.6:
        # Bảng exams của dự án hiện tại không có cột teacher_id.
        # Route này đã có @login_required("teacher"), vì vậy chỉ
        # tài khoản giáo viên mới có quyền thực hiện thao tác xóa.
        exam = c.execute(
            "SELECT * FROM exams WHERE id=?",
            (exam_id,)
        ).fetchone()

        if not exam:
            abort(404)

        # Ghi nhớ file vật lý của đề trước khi xóa DB.
        physical_files = []
        for key in ("file_path", "source_path", "original_path"):
            if key in exam.keys() and exam[key]:
                try:
                    physical_files.append(Path(exam[key]))
                except Exception:
                    pass

        attempt_ids = [
            r["id"]
            for r in c.execute(
                "SELECT id FROM student_exams WHERE exam_id=?",
                (exam_id,)
            ).fetchall()
        ]

        question_ids = [
            r["id"]
            for r in c.execute(
                "SELECT id FROM questions WHERE exam_id=?",
                (exam_id,)
            ).fetchall()
        ]

        # 1. Log chống gian lận.
        if attempt_ids:
            marks = ",".join("?" for _ in attempt_ids)
            c.execute(
                f"DELETE FROM anti_cheat_logs "
                f"WHERE student_exam_id IN ({marks})",
                attempt_ids
            )

            # 2. Đáp án học sinh.
            c.execute(
                f"DELETE FROM student_answers "
                f"WHERE student_exam_id IN ({marks})",
                attempt_ids
            )

        # 3. Phiên làm bài.
        c.execute(
            "DELETE FROM student_exams WHERE exam_id=?",
            (exam_id,)
        )

        # 4. Phương án câu hỏi — xóa tường minh để hỗ trợ DB cũ.
        if question_ids:
            marks = ",".join("?" for _ in question_ids)
            c.execute(
                f"DELETE FROM options "
                f"WHERE question_id IN ({marks})",
                question_ids
            )

        # 5. Câu hỏi và cấu hình đề.
        c.execute(
            "DELETE FROM questions WHERE exam_id=?",
            (exam_id,)
        )

        table_names = {
            r["name"]
            for r in c.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        }

        if "exam_settings" in table_names:
            c.execute(
                "DELETE FROM exam_settings WHERE exam_id=?",
                (exam_id,)
            )

        # 6. Xóa bản ghi đề cuối cùng.
        c.execute(
            "DELETE FROM exams WHERE id=?",
            (exam_id,)
        )

        c.commit()

    # 7. Xóa file vật lý nếu có.
    for p in physical_files:
        try:
            if p.exists() and p.is_file():
                p.unlink()
        except Exception:
            pass

    flash(
        "Đã xóa đề. Học sinh sẽ không còn nhìn thấy hoặc làm đề này."
    )
    return redirect(url_for("teacher_dashboard"))

@app.get('/teacher/exam/<int:exam_id>/student/<int:student_exam_id>')
@login_required('teacher')
def student_result_detail(exam_id,student_exam_id):
    with db() as c:
        se=c.execute("SELECT se.*,COALESCE(u.full_name,'Không xác định') AS full_name,COALESCE(u.student_code,'') AS student_code FROM student_exams se LEFT JOIN users u ON u.id=se.student_id WHERE se.id=? AND se.exam_id=?",(student_exam_id,exam_id)).fetchone()
        if not se: abort(404)
        exam=c.execute('SELECT * FROM exams WHERE id=?',(exam_id,)).fetchone()
        answers=c.execute('SELECT a.*,q.section,q.section_order,q.qtype,q.blocks_json,q.correct_option,q.short_answer FROM student_answers a JOIN questions q ON q.id=a.question_id WHERE a.student_exam_id=? ORDER BY q.section,q.section_order',(student_exam_id,)).fetchall()
        logs=c.execute('SELECT event_type,details,created_at FROM anti_cheat_logs WHERE student_exam_id=? ORDER BY id',(student_exam_id,)).fetchall()
    return render_template('student_result_detail.html',exam=exam,se=se,answers=answers,logs=logs)

@app.get("/student/result/<int:student_exam_id>")
@login_required("student")
def result(student_exam_id):
    with db() as c:
        ensure_practice_exam_schema(c)
        se=c.execute(
            "SELECT * FROM student_exams WHERE id=? AND student_id=?",
            (student_exam_id,session["user_id"]),
        ).fetchone()
        exam=(
            c.execute("SELECT * FROM exams WHERE id=?",(se["exam_id"],)).fetchone()
            if se else None
        )
    if not se: abort(404)
    return render_template("result.html",se=se,exam=exam)

if __name__=="__main__":
    init_db()
    app.run(
        host=os.environ.get("APP_HOST", "127.0.0.1"),
        port=int(os.environ.get("PORT", "5000")),
        debug=os.environ.get("APP_DEBUG", "0") == "1",
    )
