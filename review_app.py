"""
Phase 3 — Factory Copilot: Duyệt Taxonomy + Truy xuất dữ liệu
2 tab trong cùng 1 trang:
  1. Duyệt Taxonomy — nhóm reviewer xem hàng đợi, Duyệt/Từ chối
  2. Truy xuất dữ liệu — 6 câu hỏi mục tiêu Phase 0, ai cũng xem được (không cần chọn reviewer)

Cài đặt trước khi chạy:
    pip install streamlit psycopg2-binary pandas

Chạy thử trên máy (mở trình duyệt tự động):
    streamlit run review_app.py
"""

import streamlit as st
import streamlit.components.v1 as components
import base64
import psycopg2
import pandas as pd
import anthropic
import json
import io
import requests
import uuid
import re
import difflib
import smtplib
from email.mime.text import MIMEText
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from datetime import datetime

# ============================================================
# THÔNG TIN KẾT NỐI — đọc từ Streamlit Secrets, không viết thẳng vào code (giống supplier_portal.py).
# Khi chạy thử local: tạo file .streamlit/secrets.toml cạnh review_app.py, KHÔNG đẩy lên GitHub.
# Khi deploy lên Streamlit Cloud: điền ở Settings → Secrets của app.
# ============================================================
DB_HOST = st.secrets.get("DB_HOST", "aws-0-ap-northeast-2.pooler.supabase.com")
DB_PORT = st.secrets.get("DB_PORT", 5432)
DB_NAME = st.secrets.get("DB_NAME", "postgres")
DB_USER = st.secrets.get("DB_USER", "postgres.sdlkfcwjfvtvpjwcmdxr")
DB_PASSWORD = st.secrets.get("DB_PASSWORD", "")
ANTHROPIC_API_KEY = st.secrets.get("ANTHROPIC_API_KEY", "")

if not DB_PASSWORD:
    st.error(
        "Missing Secrets configuration (DB_PASSWORD) — go to App settings → Secrets."
    )
    st.stop()

PRODUCT_GROUPS = [
    "HANGTAG", "WOVEN", "BAG GARM",
    "Ribbon", "Carton", "All Products",
    "BADGE", "BAG OTHER", "BOX RIGID", "BUTTON", "HEADER", "HEAT LABEL", "HOOK",
    "OTHER", "PRINT", "RAW INK", "RAW OTHER", "RAW RIBBON", "RAW STK", "RAW THERMO",
    "RFID", "RFID HTG", "RFID STK", "RIDER", "RIS CARE", "RIS HTG", "RIS STK",
    "STICKER", "STRING", "WATERFALL",
]
RC_CATEGORIES = [
    "Material", "Machine", "Method", "Man", "Environment", "Measurement",
    "Supplier", "Artwork", "Data", "Logistics", "Customer", "Unknown",
]
SEVERITIES = ["Critical", "Major", "Minor"]
CAPA_TYPES = ["Corrective", "Preventive"]
BRANDS = [
    "GYMSHARK", "HESTRA", "VAUDE", "PASSENGER", "NINEPINE", "NOSO",
    "SANDQVIST", "MARTIN MAGNUSSON", "STREET ONE & CECIL", "FRED PERRY",
    "HELLY HANSEN", "JACK WOLFSKIN", "LACOSTE",
]
# Đánh giá ban đầu của CS ngay lúc nhập complaint — phân biệt complaint THẬT (lỗi từ Nilorn) với
# complaint do chính khách hàng gây ra nhưng vẫn phát sinh khiếu nại. Chốt ngay lúc nhập liệu (CS
# thường đã biết qua trao đổi với khách hàng), có thể sửa lại sau nếu điều tra ra khác. Mặc định
# luôn là "Lỗi thật" vì đa số complaint là thật, CS chỉ cần đổi khi biết chắc là lỗi khách hàng.
COMPLAINT_VALIDITY_OPTIONS = ["Genuine Defect (Nilorn)", "Customer-Caused Issue"]

# SUPPLIER_PORTAL_BASE_URL — đã deploy thật lên Streamlit Community Cloud (public URL cố định).
SUPPLIER_PORTAL_BASE_URL = "https://nilorn-supplier-app-zccldhcz5ghgnrk5mfkisg.streamlit.app"

# ============================================================
# THÔNG BÁO QUA EMAIL — dùng Gmail SMTP riêng, KHÔNG qua Microsoft 365 công ty.
# Cách lấy GMAIL_NOTIFY_APP_PASSWORD: bật "2-Step Verification" cho tài khoản Gmail này tại
# myaccount.google.com/security, sau đó vào myaccount.google.com/apppasswords để tạo App Password
# 16 ký tự (KHÔNG dùng mật khẩu Gmail thường — Google sẽ từ chối).
# ============================================================
GMAIL_NOTIFY_ADDRESS = st.secrets.get("GMAIL_NOTIFY_ADDRESS", "")
GMAIL_NOTIFY_APP_PASSWORD = st.secrets.get("GMAIL_NOTIFY_APP_PASSWORD", "")
APPROVER_EMAILS = [e.strip() for e in st.secrets.get("APPROVER_EMAILS", "").split(",") if e.strip()]

# Hộp thư CS chung — LUÔN nhận thông báo mỗi khi có complaint mới (nhập tay hoặc NCC tự khai báo),
# không phân biệt "Ghi nhận bởi" là ai. Cố định (không qua Secrets) vì đây là địa chỉ dùng chung
# cố định của phòng CS, không đổi theo từng lần deploy.
CS_GENERAL_EMAIL = "CS.NVN@vn.nilorn.com"

# Link tới chính app này — chèn vào email báo approver để bấm mở app ngay. Đọc từ Secrets để dễ
# đổi khi URL đổi (ví dụ sau khi deploy lên Streamlit Cloud), không cần sửa code.
REVIEW_APP_URL = st.secrets.get("REVIEW_APP_URL", "https://nilorn-review-app.streamlit.app")

# Supabase Storage — dùng chung bucket "defect-media" với supplier_portal.py, để CS có thể thay
# ảnh/video sau khi nhà cung cấp đã nộp.
SUPABASE_PROJECT_REF = st.secrets.get("SUPABASE_PROJECT_REF", "sdlkfcwjfvtvpjwcmdxr")
SUPABASE_URL = f"https://{SUPABASE_PROJECT_REF}.supabase.co"
SUPABASE_SERVICE_KEY = st.secrets.get("SUPABASE_SERVICE_KEY", "")
STORAGE_BUCKET = "defect-media"


def upload_to_storage(file_bytes, filename, content_type):
    """Upload 1 file lên Supabase Storage bucket, trả về URL công khai — giống hệt hàm cùng tên
    trong supplier_portal.py, dùng ở đây để CS thay ảnh/video sau khi nhà cung cấp đã nộp."""
    if not SUPABASE_SERVICE_KEY:
        raise Exception("SUPABASE_SERVICE_KEY not set in review_app.py.")
    safe_name = re.sub(r"[^a-zA-Z0-9._-]", "_", filename)
    path = f"{datetime.now().strftime('%Y%m%d')}/{uuid.uuid4().hex}_{safe_name}"
    url = f"{SUPABASE_URL}/storage/v1/object/{STORAGE_BUCKET}/{path}"
    headers = {
        "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
        "apikey": SUPABASE_SERVICE_KEY,
        "Content-Type": content_type or "application/octet-stream",
    }
    resp = requests.post(url, headers=headers, data=file_bytes, timeout=60)
    if resp.status_code not in (200, 201):
        raise Exception(f"Upload failed ({resp.status_code}): {resp.text[:200]}")
    return f"{SUPABASE_URL}/storage/v1/object/public/{STORAGE_BUCKET}/{path}"



def send_notification_email(to_addrs, subject, body):
    """Gửi email thông báo qua Gmail SMTP — miễn phí, tách biệt hoàn toàn khỏi Microsoft 365 công ty,
    nên không phụ thuộc IT duyệt gì cả. Trả về (True, None) nếu gửi thành công, (False, lý_do_lỗi)
    nếu thất bại — KHÔNG raise exception (không làm crash luồng chính của app), nhưng nơi gọi hàm
    này cần tự kiểm tra kết quả trả về để hiện đúng trạng thái cho người dùng, thay vì mặc định coi
    là đã gửi thành công."""
    if isinstance(to_addrs, str):
        to_addrs = [to_addrs]
    to_addrs = [a for a in to_addrs if a]
    if not to_addrs:
        return False, "No recipient email address"
    try:
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = subject
        msg["From"] = GMAIL_NOTIFY_ADDRESS
        msg["To"] = ", ".join(to_addrs)
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=10) as server:
            server.starttls()
            server.login(GMAIL_NOTIFY_ADDRESS, GMAIL_NOTIFY_APP_PASSWORD)
            server.sendmail(GMAIL_NOTIFY_ADDRESS, to_addrs, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)


def _log_notify_attempt(conn, ref_id, emails, mail_ok, mail_err):
    """Ghi lại kết quả gửi email vào notify_debug_log — dùng để debug (xem trực tiếp qua Supabase
    Table Editor) thay vì mò log Streamlit Cloud, giống hệt cơ chế đã dùng ở supplier_portal.py."""
    try:
        with conn.cursor() as cur:
            cur.execute(
                "insert into notify_debug_log (complaint_id, cs_email, mail_ok, mail_err) "
                "values (%s, %s, %s, %s);",
                (ref_id, ", ".join(emails) if emails else None, mail_ok, mail_err),
            )
        conn.commit()
    except Exception:
        pass


def notify_recorded_by_assignment(conn, complaint_id, staff_id, so_po):
    """Gửi email báo khi 1 complaint được gán/đổi người ghi nhận (CS phụ trách) — gửi tới CẢ hộp
    thư CS chung (CS_GENERAL_EMAIL) LẪN email riêng của người vừa được gán, để cả 2 đều biết,
    không phân biệt complaint đến từ luồng nào (nhập tay hay NCC tự khai báo). Dùng khi CS
    SỬA/GÁN LẠI người phụ trách sau này qua trang chi tiết — KHÔNG dùng cho lúc tạo complaint mới
    (lúc đó dùng notify_new_complaint_recorded, vì luôn phải báo CS chung dù có chọn người hay không)."""
    if not staff_id:
        return
    try:
        with conn.cursor() as cur:
            cur.execute("select name, email from cs_staff where staff_id = %s;", (staff_id,))
            staff_row = cur.fetchone()
    except Exception:
        return
    if not staff_row:
        return
    staff_name, staff_email = staff_row
    recipients = [CS_GENERAL_EMAIL]
    if staff_email:
        recipients.append(staff_email)
    body = (
        f"Complaint (SO/PO: {so_po or '(none)'}) was just recorded by {staff_name}, now in charge.\n\n"
        f"Open the app, 'Data Lookup' tab to view details: {REVIEW_APP_URL}"
    )
    mail_ok, mail_err = send_notification_email(
        recipients, f"[Nilorn Internal AI] Complaint {so_po or ''} — Recorded by {staff_name}", body,
    )
    _log_notify_attempt(conn, complaint_id, recipients, mail_ok, mail_err)


def notify_new_complaint_recorded(conn, complaint_id, staff_id, so_po, source_label="New Complaint (Manual Entry)"):
    """Gửi email báo có 1 complaint MỚI vừa được tạo — LUÔN gửi tới hộp thư CS chung
    (CS_GENERAL_EMAIL) bất kể có chọn 'Ghi nhận bởi' hay không, và gửi thêm tới đúng người đó
    (email riêng) nếu đã chọn. Dùng đúng lúc TẠO complaint mới (cả nhập tay lẫn NCC tự khai báo)."""
    recipients = [CS_GENERAL_EMAIL]
    staff_name = None
    if staff_id:
        try:
            with conn.cursor() as cur:
                cur.execute("select name, email from cs_staff where staff_id = %s;", (staff_id,))
                staff_row = cur.fetchone()
            if staff_row:
                staff_name, staff_email = staff_row
                if staff_email and staff_email not in recipients:
                    recipients.append(staff_email)
        except Exception:
            pass
    body = (
        f"A new complaint has just been recorded ({source_label}).\n\n"
        f"SO/PO: {so_po or '(none)'}\n"
        f"Recorded by: {staff_name or '(not assigned)'}\n\n"
        f"Open the app, 'Data Lookup' tab to view details: {REVIEW_APP_URL}"
    )
    mail_ok, mail_err = send_notification_email(
        recipients, f"[Nilorn Internal AI] New Complaint — {so_po or '(no SO/PO)'}", body,
    )
    _log_notify_attempt(conn, complaint_id, recipients, mail_ok, mail_err)



def check_and_notify_new_suggestions(conn):
    """Kiểm tra có đề xuất Pending nào MỚI kể từ lần thông báo gần nhất không — nếu có, gửi 1 email
    tổng hợp cho approver rồi cập nhật mốc thời gian last_notified_at. Cơ chế mốc thời gian tự chống
    gửi trùng lặp dù hàm này được gọi lại nhiều lần (mỗi lần app chạy lại script)."""
    try:
        with conn.cursor() as cur:
            cur.execute("select last_notified_at from app_notification_state where id = true;")
            row = cur.fetchone()
        if not row:
            return
        last_notified_at = row[0]
        with conn.cursor() as cur:
            cur.execute(
                "select suggestion_type, ai_suggested_name from taxonomy_suggestion "
                "where status = 'Pending' and created_at > %s order by created_at;",
                (last_notified_at,),
            )
            new_items = cur.fetchall()
        if not new_items:
            return
        lines = [f"- [{kind}] {name}" for kind, name in new_items]
        app_link_line = f"\n\nOpen the app: {REVIEW_APP_URL}" if REVIEW_APP_URL else ""
        body = (
            f"There are {len(new_items)} new suggestion(s) pending approval in Nilorn Internal AI:\n\n"
            + "\n".join(lines)
            + app_link_line
            + "\n\nOpen the app, 'Review Taxonomy' tab to view and process."
        )
        send_notification_email(
            APPROVER_EMAILS, f"[Nilorn Internal AI] {len(new_items)} new suggestion(s) need approval", body,
        )
        with conn.cursor() as cur:
            cur.execute("update app_notification_state set last_notified_at = now() where id = true;")
        conn.commit()
    except Exception:
        pass


def fetch_cs_staff(conn):
    with conn.cursor() as cur:
        cur.execute("select staff_id, name, role from cs_staff order by name;")
        return cur.fetchall()


def fetch_cs_staff_email(conn, staff_display_name):
    """Tra email của 1 CS staff từ chuỗi hiển thị dạng 'Tên (Vai trò)' (đúng định dạng cột
    'Ghi nhận bởi'). Trả về None nếu không có tên/chưa gán email — dùng để gửi bản nháp email
    (soạn cho khách hàng/nhà cung cấp) thẳng vào hộp thư Outlook của họ để review."""
    if not staff_display_name:
        return None
    plain_name = staff_display_name.split(" (")[0].strip()
    with conn.cursor() as cur:
        cur.execute("select email from cs_staff where name = %s limit 1;", (plain_name,))
        row = cur.fetchone()
    return row[0] if row and row[0] else None


# ============================================================
# XUẤT FILE — Excel / Word / PDF
# ============================================================
REPORT_TITLE = "Issue Report and Corrective Action"


def export_word_report(report: dict, supplier_signature_name=None, supplier_signature_image_b64=None) -> bytes:
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)
    # Ép mặc định TOÀN BỘ tài liệu về sát dòng, không chỉ riêng từng đoạn — để không có khoảng
    # trắng thừa dù có lỡ sót chỗ nào (mặc định của Word/python-docx thường tự thêm ~8-10pt sau
    # mỗi đoạn văn nếu không khai báo rõ). / Force the WHOLE document's default to tight spacing —
    # not just per-paragraph — so no stray extra gap remains anywhere (Word/python-docx's default
    # normally adds ~8-10pt after every paragraph unless explicitly overridden).
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(4)
    style.paragraph_format.line_spacing = 1.0

    title = doc.add_heading(REPORT_TITLE, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph(f"Exported on: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if subtitle.runs:
        subtitle.runs[0].italic = True
        subtitle.runs[0].font.size = Pt(10)

    doc.add_paragraph()

    def add_section(title_text):
        h = doc.add_heading(title_text, level=1)
        for run in h.runs:
            run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

    def add_field(label, key):
        rendered_keys.add(key)
        val = report.get(key)
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        p.add_run(f"{label}: ").bold = True
        p.add_run(str(val) if val not in (None, "") else "")

    def add_list_field(label, key):
        """Cho các trường có thể có NHIỀU giá trị (Root cause, CAPA) — mỗi giá trị xuống dòng riêng."""
        rendered_keys.add(key)
        items = report.get(key) or []
        p = doc.add_paragraph()
        p.add_run(f"{label}:").bold = True
        if not items:
            return
        for item in items:
            doc.add_paragraph(str(item), style="List Bullet")

    rendered_keys = {"Issue Description"}

    # I. Thông tin chung
    add_section("I. General Information")
    info_table = doc.add_table(rows=0, cols=2)
    info_table.style = "Light List Accent 1"
    for label, key in [
        ("System Complaint ID", "System Complaint ID"),
        ("Date Occurred", "Date Occurred"),
        ("Recorded By", "Recorded By"),
        ("SO/PO", "SO/PO"),
        ("Purchase Order No.", "Purchase Order No."),
    ]:
        rendered_keys.add(key)
        row = info_table.add_row().cells
        row[0].text = label
        if row[0].paragraphs[0].runs:
            row[0].paragraphs[0].runs[0].bold = True
        row[1].text = str(report.get(key)) if report.get(key) not in (None, "") else ""

    doc.add_paragraph()

    # II. Sản phẩm & Khách hàng
    add_section("II. Product & Customer")
    for label in ["Product", "Supplier", "Brand", "Customer", "Related Machine",
                  "Quantity Inspected", "Defect Quantity"]:
        add_field(label, label)

    # III. Mô tả sự cố
    add_section("III. Issue Description")
    doc.add_paragraph(str(report.get("Issue Description") or ""))

    rendered_keys.add("Defect Photo")
    defect_photo_b64 = report.get("Defect Photo")
    if defect_photo_b64:
        try:
            photo_bytes = base64.b64decode(defect_photo_b64)
            photo_stream = io.BytesIO(photo_bytes)
            caption_p = doc.add_paragraph()
            caption_p.add_run("Illustration:").bold = True
            doc.add_picture(photo_stream, width=Inches(4))
        except Exception:
            # Ảnh lỗi/không đọc được — bỏ qua phần ảnh, không làm gián đoạn việc xuất báo cáo.
            pass

    # IV. Phân loại & Hướng khắc phục
    add_section("IV. Classification & Corrective Action")
    add_field("Defect", "Defect")
    add_list_field("Root cause", "Root cause")
    add_list_field("CAPA", "CAPA")

    # V. Kết quả phân loại AI (chỉ hiện nếu có — báo cáo lúc nhập mới có, báo cáo tra cứu lại thì không)
    if "AI Classification Result" in report:
        rendered_keys.add("AI Classification Result")
        add_section("V. AI Classification Result")
        ai_text = str(report.get("AI Classification Result") or "")
        for line in ai_text.split(" | "):
            clean_line = line.replace("**", "")
            doc.add_paragraph(clean_line, style="List Bullet")

    # VI. Thông tin khác — bất kỳ trường nào chưa hiển thị ở trên (ví dụ khi tra cứu bản mới nhất)
    remaining = {k: v for k, v in report.items() if k not in rendered_keys}
    if remaining:
        add_section("VI. Other Information")
        for k, v in remaining.items():
            p = doc.add_paragraph()
            p.add_run(f"{k}: ").bold = True
            p.add_run(str(v) if v not in (None, "") else "")

    # Chữ ký
    doc.add_paragraph()
    doc.add_paragraph()
    sig_table = doc.add_table(rows=2, cols=2)
    sig_table.autofit = True
    c00 = sig_table.cell(0, 0).paragraphs[0]
    c00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c00.add_run("Prepared By").bold = True
    c01 = sig_table.cell(0, 1).paragraphs[0]
    c01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c01.add_run("Approved By").bold = True

    # Ô "Người lập" — nếu nhà cung cấp đã nộp báo cáo kèm ảnh chữ ký, tự động chèn ảnh đó vào đây
    # thay vì để trống chờ ký tay. Nếu không có ảnh (nhà cung cấp không tải ảnh, hoặc complaint
    # chưa có báo cáo nào từ nhà cung cấp), giữ nguyên chỗ trống để ký tay như trước.
    cell_lap = sig_table.cell(1, 0)
    p_lap = cell_lap.paragraphs[0]
    p_lap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if supplier_signature_image_b64:
        try:
            image_bytes = base64.b64decode(supplier_signature_image_b64)
            image_stream = io.BytesIO(image_bytes)
            run_img = p_lap.add_run()
            run_img.add_picture(image_stream, width=Inches(1.5))
        except Exception:
            # Ảnh lỗi/không đọc được — vẫn tiếp tục xuất báo cáo, chỉ bỏ qua phần ảnh.
            p_lap.add_run("\n\n\n(Signature, full name)").italic = True
        if supplier_signature_name:
            name_p = cell_lap.add_paragraph(supplier_signature_name)
            name_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p_lap.add_run("\n\n\n(Signature, full name)").italic = True

    p_duyet = sig_table.cell(1, 1).paragraphs[0]
    p_duyet.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_duyet.add_run("\n\n\n(Signature, full name)").italic = True

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def export_complaint_report_pdf(report: dict) -> bytes:
    """PDF cho complaint 'Complaint mới' (nhập tay) / complaint cũ nói chung — dùng chung
    fetch_complaint_full_report() làm nguồn dữ liệu, y hệt cách export_submission_pdf() làm cho
    luồng NCC. PDF luôn xuất bằng TIẾNG ANH — font mặc định reportlab không hỗ trợ tiếng Việt."""
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter)
    styles = getSampleStyleSheet()
    story = [Paragraph("Complaint Report", styles["Title"]), Spacer(1, 12)]

    def field(label, value):
        story.append(Paragraph(f"<b>{label}:</b> {value if value not in (None, '') else '-'}", styles["Normal"]))
        story.append(Spacer(1, 4))

    story.append(Paragraph("Order Info", styles["Heading2"]))
    field("Date Occurred", report.get("Date Occurred"))
    field("Supplier", report.get("Supplier"))
    field("Vendor No.", report.get("Vendor No."))
    field("Product", report.get("Product"))
    field("Product Group", report.get("Product Group"))
    field("Brand", report.get("Brand"))
    field("SO/PO", report.get("SO/PO"))
    field("Purchase Order No.", report.get("Purchase Order No."))
    field("Order Qty", report.get("Quantity Inspected"))
    field("Defect Qty", report.get("Defect Quantity"))
    field("Customer", report.get("Customer"))
    field("Client Code", report.get("Client Code"))

    story.append(Spacer(1, 8))
    story.append(Paragraph("Issue Details", styles["Heading2"]))
    field("Description", report.get("Issue Description"))
    field("Assessment", report.get("Assessment"))
    field("Recorded By", report.get("Recorded By"))
    field("Defect", report.get("Defect"))
    field("Root Cause", "; ".join(report.get("Root cause") or []) or None)
    field("CAPA", "; ".join(report.get("CAPA") or []) or None)
    field("Complaint Status", report.get("Complaint Status"))
    field("Responsible Party", report.get("Responsible Party"))
    field("Replacement Cost", report.get("Replacement Cost"))

    photo_b64 = report.get("Defect Photo")
    if photo_b64:
        try:
            photo_bytes = base64.b64decode(photo_b64)
            story.append(Spacer(1, 8))
            story.append(Paragraph("Photo", styles["Heading2"]))
            story.append(RLImage(io.BytesIO(photo_bytes), width=4 * inch, height=3 * inch, kind="proportional"))
        except Exception:
            pass

    doc.build(story)
    return buf.getvalue()


def export_chart_png(df: pd.DataFrame) -> bytes:
    fig, ax = plt.subplots(figsize=(8, 4))
    df.set_index(df.columns[0]).plot(kind="bar", ax=ax, legend=False)
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    return buf.getvalue()


# ============================================================
# KẾT NỐI DATABASE — mỗi phiên trình duyệt (session) có kết nối RIÊNG
# ============================================================
def _new_db_connection():
    return psycopg2.connect(
        host=DB_HOST, port=DB_PORT, dbname=DB_NAME,
        user=DB_USER, password=DB_PASSWORD
    )


def get_connection():
    if "db_conn" not in st.session_state or st.session_state.db_conn.closed:
        st.session_state.db_conn = _new_db_connection()
    return st.session_state.db_conn


def ensure_connection():
    """Kiểm tra kết nối của PHIÊN NÀY còn sống không, tự kết nối lại nếu đã bị ngắt."""
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT 1;")
    except Exception:
        try:
            st.session_state.db_conn.close()
        except Exception:
            pass
        st.session_state.db_conn = _new_db_connection()
    return st.session_state.db_conn


def fetch_reviewers(conn):
    with conn.cursor() as cur:
        cur.execute("select reviewer_id, name from reviewer where active = true order by name;")
        return cur.fetchall()


def fetch_pending(conn):
    with conn.cursor() as cur:
        cur.execute("""
            select s.suggestion_id, s.complaint_id, s.suggestion_type, s.ai_suggested_name,
                   s.ai_reasoning, s.closest_existing_code, s.created_at,
                   c.date_opened, c.so_po, s.responsible_party
            from taxonomy_suggestion s
            left join complaint c on s.complaint_id = c.complaint_id
            where s.status = 'Pending'
            order by s.created_at asc;
        """)
        return cur.fetchall()


def fetch_existing_codes(conn, kind):
    table, code_col, name_col = {
        "Defect": ("defect_taxonomy", "defect_code", "defect_name"),
        "Root Cause": ("root_cause_taxonomy", "root_cause_code", "root_cause"),
        "CAPA": ("capa_taxonomy", "capa_code", "capa_action"),
    }[kind]
    with conn.cursor() as cur:
        cur.execute(f"select {code_col}, {name_col} from {table} order by {code_col};")
        return cur.fetchall()


def approve_match_existing(conn, suggestion_id, existing_code, reviewer_id, complaint_id, kind, responsible_party=None):
    with conn.cursor() as cur:
        cur.execute("""
            update taxonomy_suggestion
            set status = 'Merged into existing', final_code = %s,
                reviewed_by = %s, reviewed_at = %s
            where suggestion_id = %s;
        """, (existing_code, reviewer_id, datetime.now(), suggestion_id))

        if complaint_id:
            col = {"Defect": "defect_code", "Root Cause": "root_cause_code"}.get(kind)
            if col:
                cur.execute(f"update complaint set {col} = %s where complaint_id = %s;",
                            (existing_code, complaint_id))
            elif kind == "CAPA":
                cur.execute("""
                    insert into capa_action (complaint_id, capa_code, responsible_party, verification_result)
                    values (%s, %s, %s, 'Pending');
                """, (complaint_id, existing_code, responsible_party))

            # Nếu complaint này đến từ luồng NCC tự khai báo qua link chung, cập nhật lại
            # root_cause/CAPA trong supplier_submissions bằng đúng nội dung đã khớp — để báo cáo
            # Word/PDF tải về luôn là bản mới nhất, không còn giữ nguyên văn NCC tự gõ.
            if kind in ("Root Cause", "CAPA"):
                table, code_col, name_col = {
                    "Root Cause": ("root_cause_taxonomy", "root_cause_code", "root_cause"),
                    "CAPA": ("capa_taxonomy", "capa_code", "capa_action"),
                }[kind]
                cur.execute(f"select {name_col} from {table} where {code_col} = %s;", (existing_code,))
                name_row = cur.fetchone()
                if name_row:
                    sub_col = "root_cause" if kind == "Root Cause" else "capa"
                    cur.execute(
                        f"update supplier_submissions set {sub_col} = %s "
                        f"where submission_id = (select source_submission_id from complaint where complaint_id = %s);",
                        (name_row[0], complaint_id),
                    )
    conn.commit()
    if complaint_id and kind == "CAPA":
        compute_and_update_complaint_status(conn, complaint_id)


def approve_new_code(conn, suggestion_id, new_code, name, description, reviewer_id, complaint_id, kind, extra, responsible_party=None):
    with conn.cursor() as cur:
        if kind == "Defect":
            cur.execute("""
                insert into defect_taxonomy (defect_code, product_group, defect_category, defect_name, description, default_severity)
                values (%s, %s, %s, %s, %s, %s);
            """, (new_code, extra["product_group"], extra.get("category", ""), name, description, extra["severity"]))
        elif kind == "Root Cause":
            cur.execute("""
                insert into root_cause_taxonomy (root_cause_code, category, root_cause, description, related_data_field)
                values (%s, %s, %s, %s, %s);
            """, (new_code, extra["category"], name, description, extra.get("data_field", "")))
        elif kind == "CAPA":
            cur.execute("""
                insert into capa_taxonomy (capa_code, capa_type, capa_action, description, verification_method, effectiveness_window)
                values (%s, %s, %s, %s, %s, %s);
            """, (new_code, extra["capa_type"], name, description, extra.get("verification_method", ""), extra.get("effectiveness_window", "")))

        cur.execute("""
            update taxonomy_suggestion
            set status = 'Approved', final_code = %s,
                reviewed_by = %s, reviewed_at = %s
            where suggestion_id = %s;
        """, (new_code, reviewer_id, datetime.now(), suggestion_id))

        if complaint_id:
            col = {"Defect": "defect_code", "Root Cause": "root_cause_code"}.get(kind)
            if col:
                cur.execute(f"update complaint set {col} = %s where complaint_id = %s;",
                            (new_code, complaint_id))
            elif kind == "CAPA":
                cur.execute("""
                    insert into capa_action (complaint_id, capa_code, responsible_party, verification_result)
                    values (%s, %s, %s, 'Pending');
                """, (complaint_id, new_code, responsible_party))

            # Giống hệt approve_match_existing — cập nhật lại supplier_submissions nếu áp dụng.
            if kind in ("Root Cause", "CAPA"):
                sub_col = "root_cause" if kind == "Root Cause" else "capa"
                cur.execute(
                    f"update supplier_submissions set {sub_col} = %s "
                    f"where submission_id = (select source_submission_id from complaint where complaint_id = %s);",
                    (name, complaint_id),
                )
    conn.commit()
    if complaint_id and kind == "CAPA":
        compute_and_update_complaint_status(conn, complaint_id)


def reject(conn, suggestion_id, reviewer_id):
    with conn.cursor() as cur:
        cur.execute("""
            update taxonomy_suggestion
            set status = 'Rejected', reviewed_by = %s, reviewed_at = %s
            where suggestion_id = %s;
        """, (reviewer_id, datetime.now(), suggestion_id))
    conn.commit()


# ============================================================
# 6 CÂU HỎI MỤC TIÊU (giống hệt phase2_queries.py)
# ============================================================
TARGET_QUERIES = {
    "Q1: Supplier with Most Defects (6 months)": """
        select s.name as nha_cung_cap, count(*) as so_luong
        from complaint c join supplier s on c.supplier_id = s.supplier_id
        where c.date_opened >= current_date - interval '6 months'
        group by s.name order by so_luong desc;
    """,
    "Q2: CAPA Effectiveness by Type": """
        select t.capa_type as loai_capa, count(*) as tong_so_lan,
               sum(case when a.verification_result = 'Effective' then 1 else 0 end) as so_lan_hieu_qua
        from capa_action a join capa_taxonomy t on a.capa_code = t.capa_code
        group by t.capa_type;
    """,
    "Q3: Top defect theo Product Group": """
        select coalesce(p.product_group, ss.product_group) as nhom_san_pham,
               d.defect_name as ten_loi, count(*) as so_luong
        from complaint c
          join defect_taxonomy d on c.defect_code = d.defect_code
          left join product p on c.product_id = p.product_id
          left join supplier_submissions ss on ss.submission_id = c.source_submission_id
        where coalesce(p.product_group, ss.product_group) is not null
        group by nhom_san_pham, d.defect_name order by nhom_san_pham, so_luong desc;
    """,
    "Q4: Most Common Root Cause Category (6 months)": """
        select rc.category as nhom_nguyen_nhan, count(*) as so_luong
        from complaint c join root_cause_taxonomy rc on c.root_cause_code = rc.root_cause_code
        where c.date_opened >= current_date - interval '6 months'
        group by rc.category order by so_luong desc;
    """,
}


def fetch_lookup(conn, table, id_col, name_col):
    with conn.cursor() as cur:
        cur.execute(f"select {id_col}, {name_col} from {table} order by {name_col};")
        return cur.fetchall()


def load_taxonomy_list(conn, kind):
    table, code_col, name_col = {
        "Defect": ("defect_taxonomy", "defect_code", "defect_name"),
        "Root Cause": ("root_cause_taxonomy", "root_cause_code", "root_cause"),
        "CAPA": ("capa_taxonomy", "capa_code", "capa_action"),
    }[kind]
    with conn.cursor() as cur:
        cur.execute(f"select {code_col}, {name_col}, description from {table} order by {code_col};")
        rows = cur.fetchall()
    return [{"code": r[0], "name": r[1], "description": r[2]} for r in rows]


def classify_description(client, description_vi, taxonomy_list, kind, exclude_codes=None):
    taxonomy_text = "\n".join(
        f"- {t['code']}: {t['name']} — {t['description']}" for t in taxonomy_list
    )
    exclude_note = ""
    if exclude_codes:
        exclude_note = (
            f"\nLƯU Ý QUAN TRỌNG: các mã sau ĐÃ được gán cho complaint này rồi — "
            f"KHÔNG chọn lại các mã này: {', '.join(exclude_codes)}. "
            f"Nếu mô tả chỉ khớp với đúng các mã đã loại trừ này (không có gì khác thêm), "
            f"trả về matched_code = null và is_new_suggestion = false.\n"
        )
    prompt = f"""Bạn là trợ lý QA cho nhà máy sản xuất nhãn/bao bì apparel branding.
Dưới đây là danh mục {kind} hiện có:

{taxonomy_text}

Mô tả sự cố mới (tiếng Việt, ghi tự do):
"{description_vi}"
{exclude_note}
Nhiệm vụ: xác định mô tả này có khớp với 1 mã {kind} nào có sẵn ở trên không.
Chỉ trả lời bằng JSON đúng định dạng sau, không thêm chữ nào khác:

{{
  "matched_code": "<mã nếu khớp rõ ràng, hoặc null nếu không>",
  "confidence": "<High|Medium|Low>",
  "is_new_suggestion": <true nếu nên đề xuất mã mới, false nếu đã khớp>,
  "suggested_name": "<tên ngắn gọn nếu là mã mới, tiếng Anh, theo văn phong các mã hiện có>",
  "closest_existing_code": "<mã gần giống nhất dù không khớp hoàn toàn, hoặc null>",
  "reasoning": "<1-2 câu giải thích, tiếng Việt>"
}}
"""
    resp = client.messages.create(
        model="claude-haiku-4-5-20251001", max_tokens=700,
        messages=[{"role": "user", "content": prompt}],
    )
    text = extract_text(resp).strip().replace("```json", "").replace("```", "").strip()
    return json.loads(text)


# ============================================================
# HỎI AI TỰ DO — text-to-SQL an toàn (chỉ SELECT)
# ============================================================
SCHEMA_DESCRIPTION = """
Bảng defect_taxonomy(defect_code PK, product_group, defect_category, defect_name, description, detection_point, default_severity)
Bảng root_cause_taxonomy(root_cause_code PK, category, root_cause, description, related_data_field)
Bảng capa_taxonomy(capa_code PK, capa_type, capa_action, description, verification_method, effectiveness_window)
Bảng supplier(supplier_id PK uuid, name, product_groups, parent_supplier_id FK->supplier, notes)
Bảng customer(customer_id PK uuid, name, notes)
Bảng cs_staff(staff_id PK uuid, name, role)
Bảng machine(machine_id PK uuid, name, machine_type, location)
Bảng material(material_id PK uuid, name, material_type, supplier_id FK->supplier)
Bảng product(product_id PK uuid, sku, name, product_group, customer_id FK->customer)
Bảng complaint(complaint_id PK uuid, complaint_no (CỘT NÀY LUÔN TRỐNG/NULL — chưa từng được điền trong hệ
  thống, TUYỆT ĐỐI không dùng cột này để hiển thị hay nhận diện complaint; khi cần hiển thị "complaint nào"
  cho người dùng dễ đọc, LUÔN dùng so_po thay thế, hoặc complaint_id nếu so_po cũng trống), date_opened,
  source, product_id FK->product, supplier_id FK->supplier,
  machine_id FK->machine, material_id FK->material, so_po, lot_number, quantity_inspected, quantity_affected,
  defect_code FK->defect_taxonomy, root_cause_code FK->root_cause_taxonomy,
  status (text, LUÔN là 1 trong đúng 5 giá trị: 'Thiếu Customer', 'Thiếu Client', 'Thiếu Replacement Cost',
  'Thiếu xác minh CAPA', 'Closed' — đây là trạng thái tự động tính theo dữ liệu còn thiếu, không phải trạng thái
  chọn tay tự do. 'Closed' nghĩa là đã đóng hồ sơ đầy đủ; 4 giá trị còn lại nghĩa là đang thiếu đúng phần đó,
  theo thứ tự ưu tiên Customer > Client > Replacement Cost > xác minh CAPA),
  brand (text, 1 trong 13 giá trị cố định: GYMSHARK, HESTRA, VAUDE, PASSENGER, NINEPINE, NOSO, SANDQVIST,
  MARTIN MAGNUSSON, STREET ONE & CECIL, FRED PERRY, HELLY HANSEN, JACK WOLFSKIN, LACOSTE),
  customer_id FK->customer, recorded_by FK->cs_staff,
  complaint_validity (text, 1 trong đúng 2 giá trị: 'Genuine Defect (Nilorn)' hoặc 'Customer-Caused Issue' —
  do CS tự đánh giá ngay lúc nhập complaint, phân biệt complaint THẬT SỰ do lỗi sản xuất/chất lượng của
  Nilorn, với complaint phát sinh do CHÍNH KHÁCH HÀNG gây ra (ví dụ khách bảo quản sai, dùng sai cách...)
  nhưng vẫn khiếu nại. CHỈ có ý nghĩa cho complaint có customer_id — không áp dụng cho complaint từ luồng
  NCC tự khai báo (đó là lỗi từ nhà cung cấp, không liên quan khái niệm này). Khi câu hỏi nhắc tới "tỷ lệ
  complaint thật/giả", "complaint do lỗi khách hàng", hay "khách hàng nào hay report sai" — LUÔN GROUP BY
  theo khách hàng (join bảng customer qua customer_id), đếm riêng theo complaint_validity, rồi tính tỷ lệ
  bằng cách ép kiểu số thực (count(...)::numeric / NULLIF(tổng, 0) * 100) để tránh chia số nguyên bị làm
  tròn sai và tránh lỗi chia cho 0.,
  dyne_level, humidity_pct, temperature_c, moisture_pct, delta_e, notes)

QUAN TRỌNG — giới hạn dữ liệu của machine_id, material_id, brand: 3 cột này CHỈ có giá trị cho
complaint nhập tay qua form "Complaint mới" (do CS tự chọn khi nhập) — LUÔN là NULL cho complaint
đến từ luồng NCC tự khai báo qua link chung (nhà cung cấp không có thông tin nội bộ này để khai).
Khi câu hỏi liên quan tới Machine, Material, hoặc Brand: BẮT BUỘC thêm câu lưu ý cuối câu trả lời
rằng kết quả chỉ phản ánh complaint nhập tay, KHÔNG bao gồm complaint từ luồng NCC tự khai báo
(vì thiếu thông tin này ở nguồn) — không được âm thầm bỏ qua giới hạn này.
Bảng complaint_event(event_id PK uuid, complaint_id FK->complaint, event_type, event_date, actor, notes)
Bảng capa_action(capa_action_id PK uuid, complaint_id FK->complaint, capa_code FK->capa_taxonomy,
  date_proposed, date_implemented, responsible_party, verification_result, verification_date)
Bảng complaint_measurement(measurement_id PK uuid, complaint_id FK->complaint, field_name, field_value, unit, recorded_date)
Bảng supplier_submissions(submission_id PK uuid, submitted_at, record_date, supplier_name_raw,
  vendor_code FK->vendor_lookup, vendor_name_matched, sales_order_no, purchase_order_no, item_no,
  order_qty, defect_qty, description, root_cause, capa, defect_image_urls, defect_video_url,
  customer_name, client_code FK->client_lookup, bear_the_claim, replacement_cost, replacement_cost_currency,
  product_group — AI tự phân loại từ item_no vào 1 trong các nhóm Product Group chuẩn giống hệt
  bảng product.product_group, dùng khi câu hỏi nhắc tới Product Group cho complaint đến từ luồng
  NCC (complaint.product_id luôn NULL cho luồng này — bảng product KHÔNG áp dụng được, phải dùng
  cột product_group ở đây thay thế))
  — dữ liệu do NHÀ CUNG CẤP TỰ KHAI BÁO qua link chung (luồng mới, khác với complaint nhập tay).
  MỖI dòng ở đây đều có 1 dòng complaint tương ứng (nối qua complaint.source_submission_id =
  supplier_submissions.submission_id) — ưu tiên dùng đúng bảng này khi câu hỏi nhắc tới Customer,
  Client Code, Replacement Cost, Responsible Party, Vendor No., Item No. (những cột chỉ có ở đây,
  KHÔNG có trong bảng complaint).
  QUAN TRỌNG khi câu hỏi liên quan tới Product Group: LUÔN dùng
  coalesce(product.product_group, supplier_submissions.product_group) — không chỉ join qua bảng
  product một mình, vì sẽ bỏ sót toàn bộ complaint từ luồng NCC.
Bảng vendor_lookup(vendor_code PK, vendor_name, country_code) — danh sách mã nhà cung cấp chuẩn.
Bảng client_lookup(client_code PK, client_name, country_code) — danh sách mã khách hàng chuẩn.

QUAN TRỌNG — phân biệt "Customer" và "Brand" (cả 2 đều có thể được gọi là "khách hàng" trong tiếng Việt, nhưng KHÁC NHAU):
- customer (bảng customer, cột complaint.customer_id): công ty NHẬN sản phẩm từ Nilorn để tiếp tục may thành phẩm
  (ví dụ "Fashion Garments 2 Co., Ltd"). Đây là đối tác sản xuất, không phải người tiêu dùng cuối.
- brand (cột complaint.brand): thương hiệu THỰC SỰ bán ra thị trường cho người tiêu dùng cuối (GYMSHARK, HESTRA,
  VAUDE, PASSENGER, NINEPINE, NOSO). Đây thường là ý người dùng muốn nói khi hỏi "khách hàng" một cách chung chung
  trong ngành báo cáo chất lượng theo thương hiệu.
NẾU câu hỏi dùng từ "khách hàng" một cách CHUNG CHUNG, không nói rõ là công ty sản xuất hay thương hiệu cuối —
hãy viết SQL trả về CẢ HAI cột (c.brand as brand VÀ cu.name as customer, LEFT JOIN customer cu),
để người dùng thấy đủ cả 2 góc nhìn, thay vì tự đoán chỉ 1 trong 2.

QUAN TRỌNG — khi GROUP BY để xếp hạng/tìm "nhiều nhất" theo 1 cột có thể để trống (brand, customer_id,
defect_code, root_cause_code...): LUÔN thêm điều kiện WHERE <cột> IS NOT NULL trước khi GROUP BY, để tránh
nhóm "chưa gán/để trống" (hiện ra là NULL/None) chiếm vị trí đầu bảng xếp hạng một cách vô nghĩa — nhóm NULL
không phải là câu trả lời hợp lệ cho câu hỏi "cái nào nhiều nhất".
"""


STATUS_LABELS = {
    "Thiếu Customer": "Missing Customer Information",
    "Thiếu Client": "Missing Client Code",
    "Thiếu Replacement Cost": "Missing Replacement Cost",
    "Thiếu xác minh CAPA": "CAPA Not Yet Verified",
    "Closed": "Closed — Case Fully Resolved",
}


MAX_DEFECT_REFERENCE_IMAGES = 3


def fetch_defect_reference_images(conn, defect_code):
    """Lấy TỐI ĐA 3 ảnh MẪU/tham chiếu đại diện cho 1 mã Defect (không gắn complaint cụ thể nào) —
    dùng để người dùng hình dung lỗi này trông thế nào, hiện ở Hỏi AI và Duyệt Taxonomy. Trả về
    danh sách chuỗi base64 (có thể rỗng)."""
    if not defect_code:
        return []
    with conn.cursor() as cur:
        cur.execute("select reference_images from defect_taxonomy where defect_code = %s;", (defect_code,))
        row = cur.fetchone()
    if not row or not row[0]:
        return []
    return list(row[0])


def add_defect_reference_images(conn, defect_code, new_image_b64_list):
    """Thêm 1 hoặc nhiều ảnh mới vào danh sách ảnh mẫu của 1 mã Defect — tự động giới hạn
    KHÔNG QUÁ 3 ảnh (ảnh mới thêm vào sau, nếu vượt quá 3 thì ảnh cũ nhất bị loại bớt để
    nhường chỗ). Trả về (danh_sách_sau_khi_thêm, số_ảnh_bị_loại_bớt_do_vượt_giới_hạn)."""
    current = fetch_defect_reference_images(conn, defect_code)
    combined = current + list(new_image_b64_list)
    dropped = max(0, len(combined) - MAX_DEFECT_REFERENCE_IMAGES)
    final_list = combined[-MAX_DEFECT_REFERENCE_IMAGES:]
    with conn.cursor() as cur:
        cur.execute(
            "update defect_taxonomy set reference_images = %s::jsonb where defect_code = %s;",
            (json.dumps(final_list), defect_code),
        )
    conn.commit()
    return final_list, dropped


def remove_defect_reference_image(conn, defect_code, index):
    """Xóa 1 ảnh cụ thể (theo vị trí) khỏi danh sách ảnh mẫu của 1 mã Defect."""
    current = fetch_defect_reference_images(conn, defect_code)
    if 0 <= index < len(current):
        current.pop(index)
    with conn.cursor() as cur:
        cur.execute(
            "update defect_taxonomy set reference_images = %s::jsonb where defect_code = %s;",
            (json.dumps(current), defect_code),
        )
    conn.commit()
    return current


def fetch_latest_supplier_signature(conn, complaint_id):
    """Lấy chữ ký (tên + ảnh nếu có) từ lần nhà cung cấp nộp báo cáo GẦN NHẤT cho complaint này —
    1 complaint có thể có nhiều token/lần yêu cầu, chỉ lấy bản đã nộp (submitted_at IS NOT NULL) mới nhất."""
    with conn.cursor() as cur:
        cur.execute("""
            select supplier_signature, supplier_signature_image
            from supplier_report_token
            where complaint_id = %s and submitted_at is not null
            order by submitted_at desc
            limit 1;
        """, (complaint_id,))
        return cur.fetchone()


def fetch_complaint_full_report(conn, complaint_id):
    with conn.cursor() as cur:
        cur.execute("""
            select c.date_opened, c.notes, p.name, s.name, s.vendor_code, c.brand, cu.name, m.name,
                   c.so_po, c.lot_number, c.quantity_inspected, c.quantity_affected,
                   d.defect_code, d.defect_name,
                   st.name, st.role, c.status, c.defect_photo,
                   c.client_code, c.bear_the_claim, c.replacement_cost, c.replacement_cost_currency,
                   c.root_cause_code, c.complaint_validity, coalesce(p.product_group, ss.product_group)
            from complaint c
            left join product p on c.product_id = p.product_id
            left join supplier s on c.supplier_id = s.supplier_id
            left join customer cu on c.customer_id = cu.customer_id
            left join machine m on c.machine_id = m.machine_id
            left join defect_taxonomy d on c.defect_code = d.defect_code
            left join cs_staff st on c.recorded_by = st.staff_id
            left join supplier_submissions ss on ss.submission_id = c.source_submission_id
            where c.complaint_id = %s;
        """, (complaint_id,))
        row = cur.fetchone()

        # Root Cause có 2 nguồn: bảng phụ complaint_root_cause (do "Complaint mới" tự ghi lúc tạo
        # mới) VÀ cột complaint.root_cause_code (do Duyệt Taxonomy ghi khi duyệt sau này) — 2
        # đường này KHÔNG đồng bộ với nhau, phải gộp cả 2 mới ra đúng, tránh mất dữ liệu.
        cur.execute("""
            select rc.root_cause_code, rc.root_cause
            from complaint_root_cause crc
            join root_cause_taxonomy rc on crc.root_cause_code = rc.root_cause_code
            where crc.complaint_id = %s
            order by crc.created_at;
        """, (complaint_id,))
        rc_rows = cur.fetchall()

        cur.execute("""
            select t.capa_code, t.capa_action, a.responsible_party
            from capa_action a join capa_taxonomy t on a.capa_code = t.capa_code
            where a.complaint_id = %s;
        """, (complaint_id,))
        capa_rows = cur.fetchall()

        cur.execute("""
            select suggestion_type, status from taxonomy_suggestion
            where complaint_id = %s and status = 'Pending';
        """, (complaint_id,))
        pending_rows = cur.fetchall()

    (date_opened, notes, product_name, supplier_name, vendor_code, brand, customer_name, machine_name,
     so_po, lot_number, qty_inspected, qty_affected, defect_code, defect_name,
     staff_name, staff_role, status, defect_photo,
     client_code, bear_the_claim, replacement_cost, replacement_cost_currency,
     direct_root_cause_code, complaint_validity, product_group) = row

    root_cause_list = [
        f"{code} — {name}" for code, name in rc_rows
    ]
    # Nếu bảng phụ complaint_root_cause trống nhưng complaint.root_cause_code đã được duyệt qua
    # Duyệt Taxonomy (ghi thẳng cột này, không qua bảng phụ) — vẫn phải lấy ra, không được bỏ sót.
    if not root_cause_list and direct_root_cause_code:
        with conn.cursor() as cur:
            cur.execute("select root_cause from root_cause_taxonomy where root_cause_code = %s;", (direct_root_cause_code,))
            direct_rc_row = cur.fetchone()
        if direct_rc_row:
            root_cause_list = [f"{direct_root_cause_code} — {direct_rc_row[0]}"]
    capa_list_display = [
        f"{code} — {name}" + (f" (Responsible: {resp})" if resp else "")
        for code, name, resp in capa_rows
    ]
    pending_text = ", ".join(t for t, _ in pending_rows) if pending_rows else ""

    return {
        "System Complaint ID": str(complaint_id),
        "Date Occurred": date_opened.strftime("%d/%m/%Y") if date_opened else "",
        "Issue Description": notes or "",
        "Assessment": complaint_validity or COMPLAINT_VALIDITY_OPTIONS[0],
        "Product": product_name or "",
        "Product Group": product_group or "",
        "Supplier": supplier_name or "",
        "Vendor No.": vendor_code or "",
        "Brand": brand or "",
        "Customer": customer_name or "",
        "Client Code": client_code or "",
        "Related Machine": machine_name or "",
        "SO/PO": so_po or "",
        "Purchase Order No.": lot_number or "",
        "Quantity Inspected": qty_inspected,
        "Defect Quantity": qty_affected,
        "Defect": f"{defect_code} — {defect_name}" if defect_code else "",
        "Root cause": root_cause_list,
        "CAPA": capa_list_display,
        "Responsible Party": bear_the_claim or "",
        "Replacement Cost": f"{replacement_cost} {replacement_cost_currency}" if replacement_cost is not None else "",
        "Recorded By": f"{staff_name} ({staff_role})" if staff_name else "",
        "Complaint Status": STATUS_LABELS.get(status, status or ""),
        "Pending Approval": pending_text,
        "Defect Photo": defect_photo,
    }


def fetch_capa_actions_for_complaint(conn, complaint_id):
    with conn.cursor() as cur:
        cur.execute("""
            select a.capa_action_id, t.capa_code, t.capa_action, a.date_proposed,
                   a.date_implemented, a.verification_result, a.responsible_party
            from capa_action a join capa_taxonomy t on a.capa_code = t.capa_code
            where a.complaint_id = %s
            order by a.date_proposed;
        """, (complaint_id,))
        return cur.fetchall()


def mark_capa_implemented(conn, capa_action_id, date_implemented):
    with conn.cursor() as cur:
        cur.execute(
            "update capa_action set date_implemented = %s where capa_action_id = %s;",
            (date_implemented, capa_action_id),
        )
    conn.commit()


def compute_and_update_complaint_status(conn, complaint_id):
    """Tính lại TOÀN BỘ danh sách mục còn thiếu (không chỉ 1 mục ưu tiên cao nhất) — trả về list
    rỗng nếu đã đủ hết (Closed), hoặc list gồm mọi mục đang thiếu cùng lúc (ví dụ có thể vừa
    "Thiếu Client" vừa "Thiếu Replacement Cost" cùng lúc). complaint.status trong database vẫn
    chỉ lưu 1 giá trị duy nhất (mục ưu tiên cao nhất trong list, hoặc 'Closed' nếu list rỗng) —
    dùng cho Excel/Ask AI; còn danh sách ĐẦY ĐỦ trả về đây dùng để hiển thị badge trên UI (card
    Dashboard, trang chi tiết, Truy xuất dữ liệu) — đảm bảo mọi nơi hiện NHẤT QUÁN.

    Bọc try/except + rollback quanh toàn bộ — nếu update thất bại (ví dụ constraint database
    chưa khớp giá trị mới), tránh để transaction bị "kẹt" làm hỏng mọi câu lệnh SQL sau đó trong
    cùng 1 lần chạy trang (đã từng xảy ra thật, gây crash toàn trang)."""
    if not complaint_id:
        return []
    try:
        with conn.cursor() as cur:
            cur.execute("""
                select c.source_submission_id, c.customer_id, c.client_code, c.replacement_cost,
                       ss.customer_name, ss.client_code, ss.replacement_cost
                from complaint c
                left join supplier_submissions ss on ss.submission_id = c.source_submission_id
                where c.complaint_id = %s;
            """, (complaint_id,))
            row = cur.fetchone()
            if not row:
                return []
            (source_submission_id, c_customer_id, c_client_code, c_replacement_cost,
             ss_customer_name, ss_client_code, ss_replacement_cost) = row

            if source_submission_id:
                has_customer = ss_customer_name is not None
                has_client = ss_client_code is not None
                has_cost = ss_replacement_cost is not None
            else:
                has_customer = c_customer_id is not None
                has_client = c_client_code is not None
                has_cost = c_replacement_cost is not None

            cur.execute("""
                select
                    exists(select 1 from capa_action a where a.complaint_id = %s
                           and (a.verification_result is null or a.verification_result = 'Pending')),
                    exists(select 1 from capa_action a where a.complaint_id = %s and a.date_implemented is not null);
            """, (complaint_id, complaint_id))
            has_pending_verif, has_implemented_capa = cur.fetchone()

            missing_tags = []
            if not has_customer:
                missing_tags.append("Thiếu Customer")
            if not has_client:
                missing_tags.append("Thiếu Client")
            if not has_cost:
                missing_tags.append("Thiếu Replacement Cost")
            if has_pending_verif or not has_implemented_capa:
                missing_tags.append("Thiếu xác minh CAPA")

            new_status = missing_tags[0] if missing_tags else "Closed"
            cur.execute("update complaint set status = %s where complaint_id = %s;", (new_status, complaint_id))
        conn.commit()
        return missing_tags
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return []


VERIFICATION_RESULTS = ["Effective", "Not Effective"]


def fetch_capa_pending_verification(conn):
    with conn.cursor() as cur:
        cur.execute("""
            select a.capa_action_id, c.complaint_id, c.so_po, c.date_opened,
                   t.capa_code, t.capa_action, t.effectiveness_window,
                   a.date_implemented, a.responsible_party
            from capa_action a
            join capa_taxonomy t on a.capa_code = t.capa_code
            join complaint c on a.complaint_id = c.complaint_id
            where a.date_implemented is not null
              and (a.verification_result is null or a.verification_result = 'Pending')
            order by a.date_implemented asc;
        """)
        return cur.fetchall()


def submit_capa_verification(conn, capa_action_id, result, verification_date):
    with conn.cursor() as cur:
        cur.execute(
            "update capa_action set verification_result = %s, verification_date = %s where capa_action_id = %s;",
            (result, verification_date, capa_action_id),
        )
        cur.execute("select complaint_id from capa_action where capa_action_id = %s;", (capa_action_id,))
        row = cur.fetchone()
    conn.commit()
    if row:
        compute_and_update_complaint_status(conn, row[0])


def add_capa_action_to_complaint(conn, complaint_id, capa_code, date_proposed, responsible_party):
    with conn.cursor() as cur:
        cur.execute(
            """insert into capa_action (complaint_id, capa_code, date_proposed, responsible_party, verification_result)
               values (%s, %s, %s, %s, 'Pending');""",
            (complaint_id, capa_code, date_proposed, responsible_party),
        )
    conn.commit()


def add_root_cause_to_complaint(conn, complaint_id, root_cause_code, notes=None):
    with conn.cursor() as cur:
        cur.execute(
            "insert into complaint_root_cause (complaint_id, root_cause_code, notes) values (%s, %s, %s);",
            (complaint_id, root_cause_code, notes),
        )
    conn.commit()


def delete_capa_action(conn, capa_action_id):
    with conn.cursor() as cur:
        cur.execute("delete from capa_action where capa_action_id = %s;", (capa_action_id,))
    conn.commit()


def suggest_next_code(existing_codes, prefix):
    max_num = 0
    for c in existing_codes:
        m = re.match(rf"{prefix}-(\d+)", c)
        if m:
            max_num = max(max_num, int(m.group(1)))
    return f"{prefix}-{max_num + 1:03d}"


def normalize_code(text):
    return re.sub(r"[^a-z0-9]", "", text.lower())


def fetch_taxonomy_context(conn):
    with conn.cursor() as cur:
        cur.execute("select defect_code, defect_name, description from defect_taxonomy order by defect_code;")
        defects = cur.fetchall()
        cur.execute("select root_cause_code, root_cause, description from root_cause_taxonomy order by root_cause_code;")
        rcs = cur.fetchall()
    defect_text = "\n".join(f"- {c}: {n} — {d}" for c, n, d in defects)
    rc_text = "\n".join(f"- {c}: {n} — {d}" for c, n, d in rcs)
    return f"Danh mục Defect hiện có:\n{defect_text}\n\nDanh mục Root Cause hiện có:\n{rc_text}"


def fetch_master_data_context(conn):
    with conn.cursor() as cur:
        cur.execute("select name from supplier order by name;")
        suppliers = [r[0] for r in cur.fetchall()]
        cur.execute("select name from customer order by name;")
        customers = [r[0] for r in cur.fetchall()]
        cur.execute("select name from product order by name;")
        products = [r[0] for r in cur.fetchall()]
        cur.execute("select name from machine order by name;")
        machines = [r[0] for r in cur.fetchall()]
    return (
        "Danh sách Supplier (tên thật trong database):\n" + (", ".join(suppliers) or "(chưa có)") +
        "\n\nDanh sách Customer (tên thật trong database):\n" + (", ".join(customers) or "(chưa có)") +
        "\n\nDanh sách Product (tên thật trong database):\n" + (", ".join(products) or "(chưa có)") +
        "\n\nDanh sách Machine (tên thật trong database):\n" + (", ".join(machines) or "(chưa có)")
    )


def is_safe_select(sql):
    s = sql.strip().rstrip(";").upper()
    # Cho phép cả câu bắt đầu bằng WITH (CTE, ví dụ "WITH x AS (...) SELECT ...") — vẫn chỉ đọc
    # dữ liệu, hoàn toàn an toàn, nhưng trước đây bị chặn nhầm chỉ vì không bắt đầu đúng bằng chữ
    # SELECT. AI hay dùng CTE khi câu hỏi cần gộp dữ liệu từ nhiều bảng (ví dụ tính tỷ lệ lỗi
    # gộp cả complaint và supplier_submissions).
    if not (s.startswith("SELECT") or s.startswith("WITH")):
        return False
    forbidden = ["INSERT", "UPDATE", "DELETE", "DROP", "ALTER", "TRUNCATE", "GRANT", "REVOKE", "CREATE"]
    return not any(kw in s for kw in forbidden)

def extract_text(response):
    for block in response.content:
        if getattr(block, "type", None) == "text":
            return block.text
    return ""


@st.cache_resource
def get_ai_client():
    return anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


def generate_sql(client, question, taxonomy_context, master_data_context):
    prompt = f"""Bạn là chuyên gia SQL cho hệ thống quản lý chất lượng nhà máy. Schema database (Postgres):
{SCHEMA_DESCRIPTION}

{taxonomy_context}

{master_data_context}

Câu hỏi của người dùng (tiếng Việt): "{question}"

Viết 1 câu lệnh SQL SELECT (CHỈ SELECT, tuyệt đối không INSERT/UPDATE/DELETE/DROP) để trả lời câu hỏi này.

QUAN TRỌNG 1 — khớp tên: câu hỏi có thể viết bằng tiếng Việt, viết tắt, hoặc khác chính tả (ví dụ 'Samson VN' hay
'Samson Viet Nam' đều có thể là 'Samson Vietnam' trong danh sách Supplier ở trên; 'Fashion Garment 2' — thiếu chữ 's'
— vẫn phải nhận ra là 'Fashion Garments 2' nếu tên đó có trong danh sách Customer thật ở trên). Đối chiếu KỸ CÀNG
với danh sách Supplier/Customer/Product/Machine/danh mục Defect/Root Cause thật ở trên để tìm ĐÚNG tên/mã trong
database — kể cả khi chỉ lệch 1-2 ký tự (số ít/số nhiều, dấu câu, khoảng trắng) — rồi dùng ILIKE '%tên đúng đã tìm
được từ danh sách%' (không phải nguyên văn người dùng gõ) để lọc — không dùng dấu = cứng nhắc trừ khi chắc chắn khớp tuyệt đối.

QUAN TRỌNG 2 — hiển thị tên, không hiển thị UUID: nếu câu hỏi liên quan đến supplier/customer/product/machine,
LUÔN JOIN sang đúng bảng đó và SELECT cột "name" (đặt alias dễ hiểu, ví dụ s.name as nha_cung_cap) — tuyệt đối
không trả về các cột *_id (uuid) trong kết quả, vì người dùng không đọc hiểu được UUID.

QUAN TRỌNG 3 — dùng mã taxonomy: nếu câu hỏi nhắc đến loại lỗi/nguyên nhân bằng mô tả (ví dụ "polybag bị ẩm"),
dựa vào danh mục Defect/Root Cause thật ở trên để xác định đúng mã, lọc bằng mã đó (c.defect_code = 'DEF-034').

QUAN TRỌNG 4 — tính tỷ lệ/phần trăm: nếu câu hỏi hỏi về tỷ lệ lỗi, % defect, defect rate...:
(a) LUÔN ép kiểu số thực trước khi chia (ví dụ SUM(defect_qty)::numeric / NULLIF(SUM(order_qty), 0) * 100) —
tuyệt đối KHÔNG chia trực tiếp 2 số nguyên (Postgres sẽ làm tròn về 0 một cách âm thầm, sai kết quả).
(b) LUÔN dùng NULLIF(mẫu_số, 0) để tránh lỗi chia cho 0 khi 1 supplier/khách hàng chưa có đơn hàng nào.
(c) Nếu câu hỏi không giới hạn rõ chỉ 1 nguồn dữ liệu, GỘP cả 2 nguồn số lượng bằng UNION ALL — 1 phần từ
bảng complaint (quantity_inspected/quantity_affected, cho complaint nhập tay/cũ) VÀ 1 phần từ bảng
supplier_submissions (order_qty/defect_qty, cho luồng NCC tự khai báo) — rồi mới tính tổng/tỷ lệ trên tập
đã gộp, để không bỏ sót dữ liệu từ 1 trong 2 nguồn.

Trong SQL, chỉ dùng dấu nháy đơn (') cho chuỗi văn bản, KHÔNG dùng dấu ngoặc kép (") ở bất kỳ đâu.

Trả lời CHÍNH XÁC theo đúng 6 dòng sau, BẮT ĐẦU NGAY bằng "SQL:" — không viết thêm bất kỳ câu mở đầu,
giải thích, hay suy luận nào trước đó, không thêm chữ nào khác ngoài 6 dòng, không dùng markdown hay JSON:
SQL: <câu lệnh SELECT, viết trên 1 dòng>
EXPLANATION: <1 câu giải thích bằng tiếng Việt>
QUESTION_TYPE: <classification nếu câu hỏi đang MÔ TẢ 1 sự cố/lỗi cụ thể cần xác định loại lỗi/nguyên nhân (ví dụ "lỗi túi bị ẩm là gì", "nguyên nhân của lỗi X"); retrieval cho MỌI câu hỏi tra cứu khác (theo supplier, khách hàng, thời gian, số lượng, lịch sử...)>
MATCHED_DEFECT_CODE: <mã DEF-xxx nếu khớp rõ ràng, hoặc NONE>
MATCHED_ROOT_CAUSE_CODE: <mã RC-xxx nếu khớp rõ ràng, hoặc NONE>
SUGGEST_NEW_COMPLAINT: <YES nếu câu hỏi đang mô tả 1 sự cố/lỗi CỤ THỂ trên 1 sản phẩm/lô hàng/đối tượng cụ thể — dù đang
hỏi dưới dạng tra cứu ("lỗi này đã từng gặp chưa", "nguyên nhân của lỗi X trên sản phẩm Y là gì") — vì đây rất có thể
là 1 occurrence MỚI đáng ghi nhận thành complaint; NO nếu câu hỏi là dạng báo cáo/tổng hợp/thống kê/so sánh chung
chung, KHÔNG gắn với 1 sự cố cụ thể nào (ví dụ "truy xuất lỗi của NB VINA 6 tháng qua", "so sánh lỗi giữa các nhà
cung cấp", "supplier nào lỗi nhiều nhất", "CAPA nào hiệu quả nhất") — những câu hỏi thuần phân tích/báo cáo luôn là NO>
"""
    resp = client.messages.create(
        model="claude-sonnet-5", max_tokens=3000,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = extract_text(resp).strip()
    text = raw.replace("**", "")

    def _extract_field(label, next_labels):
        next_pattern = "|".join(next_labels)
        if next_pattern:
            pattern = rf"{label}\s*:\s*(.*?)(?=\n\s*(?:{next_pattern})\s*:|\Z)"
        else:
            pattern = rf"{label}\s*:\s*(.*)"
        m = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        return m.group(1).strip() if m else None

    all_labels = ["EXPLANATION", "QUESTION_TYPE", "MATCHED_DEFECT_CODE", "MATCHED_ROOT_CAUSE_CODE", "SUGGEST_NEW_COMPLAINT"]
    sql = _extract_field("SQL", all_labels)
    explanation = _extract_field("EXPLANATION", all_labels[1:]) or ""
    question_type = _extract_field("QUESTION_TYPE", all_labels[2:]) or "retrieval"
    matched_defect = _extract_field("MATCHED_DEFECT_CODE", all_labels[3:])
    matched_rc = _extract_field("MATCHED_ROOT_CAUSE_CODE", all_labels[4:])
    suggest_new_complaint = _extract_field("SUGGEST_NEW_COMPLAINT", [])

    if not sql:
        m = re.search(r"(select\s.+?)(?:;|\Z)", raw, re.IGNORECASE | re.DOTALL)
        if m:
            sql = m.group(1).strip()

    if not sql:
        raise ValueError(f"AI không trả về câu SQL hợp lệ.\n\n[DEBUG] Nội dung AI trả về:\n{raw[:500]}")

    matched_defect = None if not matched_defect or matched_defect.strip().upper() == "NONE" else matched_defect
    matched_rc = None if not matched_rc or matched_rc.strip().upper() == "NONE" else matched_rc

    return {
        "sql": sql.rstrip(";").strip(),
        "explanation": explanation,
        "question_type": question_type.strip().lower() if question_type else "retrieval",
        "matched_defect_code": matched_defect,
        "matched_root_cause_code": matched_rc,
        "suggest_new_complaint": bool(suggest_new_complaint and suggest_new_complaint.strip().upper() == "YES"),
        "raw": raw,
    }


def truncate_at_word(text, limit):
    text = text or ""
    if len(text) <= limit:
        return text
    cut = text[:limit].rsplit(" ", 1)[0]
    return (cut or text[:limit]) + "..."


def find_best_label_match(value, labels):
    if not value:
        return None
    if value in labels:
        return value
    value_norm = re.sub(r"\s+", " ", value.strip().lower())
    for lbl in labels:
        if re.sub(r"\s+", " ", lbl.strip().lower()) == value_norm:
            return lbl
    for lbl in labels:
        lbl_norm = re.sub(r"\s+", " ", lbl.strip().lower())
        if value_norm and (value_norm in lbl_norm or lbl_norm in value_norm):
            return lbl
    return None


def _build_extraction_instructions(today_str, product_names, customer_names, supplier_names, staff_labels, rc_list, capa_list):
    def _fmt_list(items):
        return ", ".join(items) if items else "(not in the system)"

    product_list_text = _fmt_list(product_names)
    customer_list_text = _fmt_list(customer_names)
    supplier_list_text = _fmt_list(supplier_names)
    staff_list_text = _fmt_list(staff_labels)
    rc_list_text = _fmt_list([f"{c}: {n}" for c, n in rc_list])
    capa_list_text = _fmt_list([f"{c}: {n}" for c, n in capa_list])
    brand_list_text = ", ".join(BRANDS)

    rc_capa_fields = []
    for i in range(1, 5):
        rc_capa_fields.append(
            f"ROOT_CAUSE_{i}_CODE: mã Root Cause thứ {i} khớp rõ ràng từ danh mục (ví dụ RC-012), hoặc NONE nếu "
            f"không có nguyên nhân thứ {i} hoặc không khớp mã nào"
        )
        rc_capa_fields.append(
            f"ROOT_CAUSE_{i}_NEW: nếu nguyên nhân thứ {i} có thật trong nội dung nhưng KHÔNG khớp mã nào ở trên, "
            f"viết lại ngắn gọn mô tả nguyên nhân đó; nếu không có nguyên nhân thứ {i} hoặc đã khớp mã ở trên, để NONE"
        )
    for i in range(1, 5):
        rc_capa_fields.append(
            f"CAPA_{i}_CODE: mã CAPA thứ {i} khớp rõ ràng từ danh mục (ví dụ CAPA-005), hoặc NONE nếu không có "
            f"hành động khắc phục thứ {i} hoặc không khớp mã nào"
        )
        rc_capa_fields.append(
            f"CAPA_{i}_NEW: nếu hành động khắc phục thứ {i} có thật trong nội dung nhưng KHÔNG khớp mã nào ở trên, "
            f"viết lại ngắn gọn hành động đó; nếu không có hoặc đã khớp mã ở trên, để NONE"
        )
        rc_capa_fields.append(
            f"CAPA_{i}_RESPONSIBLE: người phụ trách hành động khắc phục thứ {i} nếu nội dung có nhắc rõ tên, hoặc NONE"
        )
    rc_capa_block = "\n".join(rc_capa_fields)

    return f"""Hôm nay là ngày: {today_str}

Danh sách sản phẩm có thật: {product_list_text}
Danh sách khách hàng có thật (công ty nhận sản phẩm từ Nilorn — không phải brand): {customer_list_text}
Danh sách nhà cung cấp có thật: {supplier_list_text}
Danh sách thương hiệu cố định: {brand_list_text}
Danh sách nhân viên CS có thật (định dạng "Tên (Vai trò)"): {staff_list_text}
Danh mục Root Cause có thật (mã: tên): {rc_list_text}
Danh mục CAPA có thật (mã: tên): {capa_list_text}

Trích xuất các thông tin sau từ nội dung được cung cấp (có thể là tiếng Việt hoặc tiếng Anh — ví dụ email từ
nhà cung cấp/khách hàng nước ngoài). CHỈ điền khi thật sự chắc chắn/rõ ràng — nếu không chắc hoặc nội dung
không nhắc tới, để đúng chữ NONE, tuyệt đối không đoán bừa:

DESC: viết lại thành 1 mô tả sự cố ngắn gọn, súc tích BẰNG TIẾNG VIỆT (dù nội dung gốc là tiếng Anh), bỏ các
phần không liên quan tới lỗi (chữ ký, disclaimer, lời chào...)
PRODUCT: tên sản phẩm khớp nguyên văn từ danh sách sản phẩm, hoặc NONE
CUSTOMER: tên khách hàng khớp nguyên văn từ danh sách khách hàng, hoặc NONE
SUPPLIER: tên nhà cung cấp khớp nguyên văn từ danh sách nhà cung cấp, hoặc NONE
BRAND: 1 trong 6 thương hiệu cố định khớp nguyên văn, hoặc NONE
QUANTITY: số lượng sản phẩm bị lỗi nếu có nhắc con số cụ thể (chỉ số nguyên, không đơn vị), hoặc NONE
DATE: ngày phát sinh sự cố dạng YYYY-MM-DD nếu có nhắc ngày cụ thể hoặc từ chỉ thời gian tương đối (hôm nay/
hôm qua/tuần trước...) — tự tính ra ngày thật dựa vào hôm nay đã cho ở trên; nếu không gợi ý gì về thời gian
(kể cả ngày gửi email, nếu không phải ngày phát sinh lỗi), để NONE
STAFF: tên nhân viên CS khớp nguyên văn (đúng định dạng "Tên (Vai trò)") từ danh sách nhân viên CS, hoặc NONE
SO_PO: nếu có nhắc mã SO/PO cụ thể (thường là chuỗi chữ+số liền nhau, ví dụ VNWSO1003456, VNPO1005433, PO#
4502797449), ghi nguyên văn mã đó, hoặc NONE

QUAN TRỌNG — nội dung có thể mô tả NHIỀU nguyên nhân (root cause) và NHIỀU hành động khắc phục (CAPA) cùng lúc.
Liệt kê TỐI ĐA 4 nguyên nhân và 4 CAPA riêng biệt theo đúng thứ tự được nhắc — đừng gộp nhiều nguyên nhân khác
nhau vào chung 1 dòng. Với email từ nhà cung cấp/khách hàng, thường CHƯA có root cause/CAPA rõ ràng (họ chỉ
báo hiện tượng lỗi) — nếu vậy, để tất cả các trường ROOT_CAUSE_*/CAPA_* là NONE, đừng suy diễn:
{rc_capa_block}

Trả lời CHÍNH XÁC theo đúng thứ tự các trường trên (9 trường đầu + {len(rc_capa_fields)} trường Root Cause/CAPA), không thêm chữ nào khác."""


def _parse_complaint_extraction(text, fallback_desc):
    base_fields = ["DESC", "PRODUCT", "CUSTOMER", "SUPPLIER", "BRAND", "QUANTITY", "DATE", "STAFF", "SO_PO"]
    rc_capa_field_names = []
    for i in range(1, 5):
        rc_capa_field_names += [f"ROOT_CAUSE_{i}_CODE", f"ROOT_CAUSE_{i}_NEW"]
    for i in range(1, 5):
        rc_capa_field_names += [f"CAPA_{i}_CODE", f"CAPA_{i}_NEW", f"CAPA_{i}_RESPONSIBLE"]
    fields = base_fields + rc_capa_field_names

    def _extract(label, next_labels):
        next_pattern = "|".join(next_labels)
        if next_pattern:
            pattern = rf"{label}\s*:\s*(.*?)(?=\n\s*(?:{next_pattern})\s*:|\Z)"
        else:
            pattern = rf"{label}\s*:\s*(.*)"
        m = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        return m.group(1).strip() if m else None

    values = {}
    for i, label in enumerate(fields):
        values[label] = _extract(label, fields[i + 1:])

    def _clean(v):
        return None if not v or v.strip().upper() == "NONE" else v.strip()

    desc = _clean(values["DESC"]) or fallback_desc
    quantity = _clean(values["QUANTITY"])
    try:
        quantity = int(re.sub(r"[^\d]", "", quantity)) if quantity else None
    except ValueError:
        quantity = None
    date_str = _clean(values["DATE"])
    parsed_date = None
    if date_str:
        try:
            parsed_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            parsed_date = None

    root_causes = []
    for i in range(1, 5):
        code = _clean(values[f"ROOT_CAUSE_{i}_CODE"])
        new_desc = _clean(values[f"ROOT_CAUSE_{i}_NEW"])
        if code or new_desc:
            root_causes.append({"code": code, "new": new_desc})

    capas = []
    for i in range(1, 5):
        code = _clean(values[f"CAPA_{i}_CODE"])
        new_desc = _clean(values[f"CAPA_{i}_NEW"])
        responsible = _clean(values[f"CAPA_{i}_RESPONSIBLE"])
        if code or new_desc:
            capas.append({"code": code, "new": new_desc, "responsible": responsible})

    so_po_final, po_final = split_so_po_text(_clean(values["SO_PO"]))

    return {
        "desc": desc,
        "product": _clean(values["PRODUCT"]),
        "customer": _clean(values["CUSTOMER"]),
        "supplier": _clean(values["SUPPLIER"]),
        "brand": _clean(values["BRAND"]),
        "quantity": quantity,
        "date": parsed_date,
        "staff": _clean(values["STAFF"]),
        "so_po": so_po_final,
        "purchase_order_no": po_final,
        "root_causes": root_causes,
        "capas": capas,
    }


def split_so_po_text(so_po_raw):
    """Tách 1 chuỗi SO/PO gộp chung (vd 'VNWSO1009871 / VNPO1007640') thành (so_po, purchase_order) —
    mã bắt đầu bằng VNPO hoặc MOR (không phân biệt hoa/thường) được coi là Purchase Order No.,
    phần còn lại giữ nguyên là Sales Order No."""
    if not so_po_raw:
        return so_po_raw, None
    parts = re.split(r"\s*/\s*", so_po_raw.strip())
    so_parts, po_parts = [], []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        (po_parts if re.match(r"(?i)^(VNPO|MOR)", p) else so_parts).append(p)
    return (" / ".join(so_parts) or None), (" / ".join(po_parts) or None)


def refine_complaint_prefill(client, question, today_str, product_names, customer_names,
                              supplier_names, staff_labels, rc_list, capa_list):
    instructions = _build_extraction_instructions(
        today_str, product_names, customer_names, supplier_names, staff_labels, rc_list, capa_list,
    )
    prompt = f'Câu hỏi gốc (tiếng Việt, dạng hỏi tự do): "{question}"\n\n{instructions}'
    resp = client.messages.create(
        model="claude-haiku-4-5-20251001", max_tokens=1200,
        messages=[{"role": "user", "content": prompt}],
    )
    text = extract_text(resp).strip()
    return _parse_complaint_extraction(text, fallback_desc=question)


def extract_complaint_from_email(client, today_str, product_names, customer_names, supplier_names,
                                  staff_labels, rc_list, capa_list, email_text=None, images=None):
    """images: list các dict {"b64": ..., "media_type": ...} — cho phép gửi NHIỀU ảnh cùng lúc
    (ví dụ vừa ảnh chụp email vừa ảnh chụp lỗi thực tế) để AI đọc tổng hợp, điền chính xác hơn."""
    instructions = _build_extraction_instructions(
        today_str, product_names, customer_names, supplier_names, staff_labels, rc_list, capa_list,
    )
    content = []
    if images:
        for img in images:
            content.append({
                "type": "image",
                "source": {"type": "base64", "media_type": img["media_type"], "data": img["b64"]},
            })
        n = len(images)
        plural_note = f"({n} ảnh — có thể là email, ảnh chụp lỗi thực tế, hoặc nhiều trang cùng 1 email) " if n > 1 else ""
        content.append({
            "type": "text",
            "text": f"Đọc nội dung trong (các) ảnh trên {plural_note}(kể cả chữ trong ảnh đính kèm nếu có) và "
                    f"trích xuất thông tin theo đúng hướng dẫn sau — nếu có nhiều ảnh, tổng hợp thông tin từ "
                    f"TẤT CẢ ảnh vào cùng 1 kết quả duy nhất, không bỏ sót ảnh nào:\n\n{instructions}",
        })
        fallback_desc = "(mô tả từ ảnh — xem lại nội dung gốc)"
    else:
        content.append({
            "type": "text",
            "text": f'Nội dung email gốc:\n"""\n{email_text}\n"""\n\n{instructions}',
        })
        fallback_desc = (email_text or "")[:200]

    resp = client.messages.create(
        model="claude-sonnet-5", max_tokens=1500,
        messages=[{"role": "user", "content": content}],
    )
    text = extract_text(resp).strip()
    return _parse_complaint_extraction(text, fallback_desc=fallback_desc)


def fetch_intake_lookup_lists(conn):
    with conn.cursor() as cur:
        cur.execute("select name from product order by name;")
        product_names = [r[0] for r in cur.fetchall()]
        cur.execute("select name from customer order by name;")
        customer_names = [r[0] for r in cur.fetchall()]
        cur.execute("select name from supplier order by name;")
        supplier_names = [r[0] for r in cur.fetchall()]
        cur.execute("select name, role from cs_staff order by name;")
        staff_labels = [f"{n} ({r})" for n, r in cur.fetchall()]
        cur.execute("select root_cause_code, root_cause from root_cause_taxonomy order by root_cause_code;")
        rc_list = cur.fetchall()
        cur.execute("select capa_code, capa_action from capa_taxonomy order by capa_code;")
        capa_list = cur.fetchall()
    return product_names, customer_names, supplier_names, staff_labels, rc_list, capa_list


def suggest_supplier_for_product(conn, product_name):
    if not product_name:
        return None
    with conn.cursor() as cur:
        cur.execute("""
            select s.name, count(*) as so_lan
            from complaint c
            join product p on c.product_id = p.product_id
            join supplier s on c.supplier_id = s.supplier_id
            where p.name = %s
            group by s.name
            order by so_lan desc
            limit 1;
        """, (product_name,))
        row = cur.fetchone()
    return row[0] if row else None


def draft_supplier_request_email(client, fields, supplier_name, portal_link, signer_name=None):
    signer = signer_name or "Customer Service Team"
    prompt = f"""Bạn là nhân viên CS (Customer Service) tại 1 nhà máy sản xuất nhãn/bao bì apparel branding.
Viết 1 email TIẾNG ANH, chuyên nghiệp, gửi cho nhà cung cấp "{supplier_name}", thông báo về 1 complaint chất
lượng liên quan tới họ và yêu cầu họ điều tra, dựa trên:

Mô tả sự cố: {fields.get("desc") or "(none)"}
Sản phẩm: {fields.get("product") or "(không rõ)"}
SO/PO: {fields.get("so_po") or "(none)"}
Số lượng lỗi: {fields.get("quantity") if fields.get("quantity") is not None else "(không rõ)"}
Link báo cáo: {portal_link}

QUAN TRỌNG — tại thời điểm gửi email này, nhà cung cấp CHƯA biết nguyên nhân hay hướng khắc phục là gì (đây
là lý do chúng ta đang liên hệ họ) — TUYỆT ĐỐI không viết kiểu "xác nhận nguyên nhân quý vị đã xác định" hay
ngụ ý họ đã biết sẵn điều gì. Phải viết đúng là YÊU CẦU họ ĐIỀU TRA và XÁC ĐỊNH: (1) nguyên nhân gốc (root
cause), (2) hành động khắc phục (CAPA) đề xuất.

Nội dung cần có: thông báo ngắn gọn về sự cố → yêu cầu điều tra và xác định root cause + CAPA như trên →
hướng dẫn rõ ràng: bấm vào đúng link báo cáo đã cho ở trên để điền đầy đủ thông tin (chèn NGUYÊN VĂN link đó
vào email, không đổi khác) → nêu rõ deadline mong muốn nếu hợp lý. Giọng văn chuyên nghiệp, lịch sự nhưng rõ
ràng về mức độ khẩn cấp.

Trả lời CHÍNH XÁC theo format sau, không thêm chữ nào khác:
SUBJECT: <tiêu đề email, tiếng Anh>
BODY: <toàn bộ nội dung email, tiếng Anh, có lời chào, link báo cáo nguyên văn ở đúng chỗ, và ký tên "{signer}" ở cuối>
"""
    resp = client.messages.create(
        model="claude-sonnet-5", max_tokens=1000,
        messages=[{"role": "user", "content": prompt}],
    )
    text = extract_text(resp).strip()
    subject_m = re.search(r"SUBJECT\s*:\s*(.*?)(?=\n\s*BODY\s*:|\Z)", text, re.IGNORECASE | re.DOTALL)
    body_m = re.search(r"BODY\s*:\s*(.*)", text, re.IGNORECASE | re.DOTALL)
    subject = subject_m.group(1).strip() if subject_m else f"Quality issue — {supplier_name}"
    body = body_m.group(1).strip() if body_m else text
    body += (
        "\n\nNote: If the link takes a moment to load or shows a \"waking up\" message, "
        "this is normal — please wait about 30 seconds and the form will load automatically."
    )
    return subject, body


def apply_complaint_prefill_fields(fields):
    st.session_state["prefill_complaint_desc"] = fields["desc"]
    matched_bits = []

    st.session_state["prefill_complaint_product"] = fields["product"] or "-- unknown --"
    if fields["product"]:
        matched_bits.append(f"product **{fields['product']}**")

    st.session_state["prefill_complaint_customer"] = fields["customer"] or "-- unknown --"
    if fields["customer"]:
        matched_bits.append(f"customer **{fields['customer']}**")

    st.session_state["prefill_complaint_supplier"] = fields["supplier"] or "-- unknown --"
    if fields["supplier"]:
        matched_bits.append(f"supplier **{fields['supplier']}**")

    st.session_state["prefill_complaint_brand"] = fields["brand"] or "-- unknown --"
    if fields["brand"]:
        matched_bits.append(f"brand **{fields['brand']}**")

    st.session_state["prefill_complaint_quantity"] = fields["quantity"] if fields["quantity"] is not None else 0
    if fields["quantity"] is not None:
        matched_bits.append(f"defect quantity **{fields['quantity']}**")

    st.session_state["prefill_complaint_date"] = fields["date"] or datetime.now().date()
    if fields["date"]:
        matched_bits.append(f"date occurred **{fields['date'].strftime('%d/%m/%Y')}**")

    st.session_state["prefill_complaint_staff"] = fields["staff"] or "-- not selected --"
    if fields["staff"]:
        matched_bits.append(f"recorded by **{fields['staff']}**")

    root_causes = fields["root_causes"]
    if root_causes:
        first_rc = root_causes[0]
        st.session_state["prefill_complaint_rc"] = first_rc["code"] or "-- let AI classify --"
        st.session_state["prefill_complaint_rc_new"] = first_rc["new"] or ""
        if first_rc["code"]:
            matched_bits.append(f"root cause **{first_rc['code']}**")
        elif first_rc["new"]:
            matched_bits.append(f"new root cause: *{first_rc['new']}*")
    else:
        st.session_state["prefill_complaint_rc"] = "-- let AI classify --"
        st.session_state["prefill_complaint_rc_new"] = ""

    extra_rcs = root_causes[1:]
    st.session_state["prefill_complaint_extra_rc_count"] = len(extra_rcs)
    for i, rc in enumerate(extra_rcs):
        st.session_state[f"prefill_complaint_extra_rc_{i}_code"] = rc["code"] or "-- let AI classify --"
        st.session_state[f"prefill_complaint_extra_rc_{i}_new"] = rc["new"] or ""
        if rc["code"]:
            matched_bits.append(f"other root cause **{rc['code']}**")
        elif rc["new"]:
            matched_bits.append(f"other new root cause: *{rc['new']}*")

    capas = fields["capas"]
    if capas:
        first_capa = capas[0]
        st.session_state["prefill_complaint_capa"] = first_capa["code"] or "-- none --"
        st.session_state["prefill_complaint_capa_new"] = first_capa["new"] or ""
        st.session_state["prefill_complaint_capa_responsible"] = first_capa["responsible"] or ""
        if first_capa["code"]:
            matched_bits.append(f"CAPA **{first_capa['code']}**")
        elif first_capa["new"]:
            matched_bits.append(f"new CAPA: *{first_capa['new']}*")
        if first_capa["responsible"]:
            matched_bits.append(f"CAPA responsible person **{first_capa['responsible']}**")
    else:
        st.session_state["prefill_complaint_capa"] = "-- none --"
        st.session_state["prefill_complaint_capa_new"] = ""
        st.session_state["prefill_complaint_capa_responsible"] = ""

    extra_capas = capas[1:]
    st.session_state["prefill_complaint_extra_capa_count"] = len(extra_capas)
    for i, cp in enumerate(extra_capas):
        st.session_state[f"prefill_complaint_extra_capa_{i}_code"] = cp["code"] or "-- none --"
        st.session_state[f"prefill_complaint_extra_capa_{i}_new"] = cp["new"] or ""
        st.session_state[f"prefill_complaint_extra_capa_{i}_resp"] = cp["responsible"] or ""
        if cp["code"]:
            matched_bits.append(f"other CAPA **{cp['code']}**")
        elif cp["new"]:
            matched_bits.append(f"other new CAPA: *{cp['new']}*")

    st.session_state["prefill_complaint_so_po"] = fields["so_po"] or ""
    if fields["so_po"]:
        matched_bits.append(f"SO/PO **{fields['so_po']}**")

    st.session_state["prefill_complaint_lot"] = fields.get("purchase_order_no") or ""
    if fields.get("purchase_order_no"):
        matched_bits.append(f"Purchase Order No. **{fields['purchase_order_no']}**")

    return matched_bits


def render_create_complaint_suggestion(conn, question, button_key):
    if st.button("📝 Create New Complaint with this description", key=button_key):
        with st.spinner("AI is preparing the description..."):
            ai_client = get_ai_client()
            (product_names, customer_names, supplier_names,
             staff_labels, rc_list, capa_list) = fetch_intake_lookup_lists(conn)
            conn = ensure_connection()
            fields = refine_complaint_prefill(
                ai_client, question, datetime.now().strftime("%Y-%m-%d"),
                product_names, customer_names, supplier_names, staff_labels, rc_list, capa_list,
            )

        matched_bits = apply_complaint_prefill_fields(fields)

        if matched_bits:
            st.success(
                "Description prefilled (AI summarized) and auto-filled: " + ", ".join(matched_bits) + " — go to the '📝 New Complaint' tab to continue, review then save."
            )
        else:
            st.success(
                "Description prefilled (AI summarized) — go to the '📝 New Complaint' tab to continue (no other fields auto-detected, please fill in the rest)."
            )


def generate_supplier_report_link(conn, complaint_id, supplier_name, requested_by_name):
    staff_id = None
    if requested_by_name:
        plain_name = requested_by_name.split(" (")[0].strip()
        with conn.cursor() as cur:
            cur.execute("select staff_id from cs_staff where name = %s limit 1;", (plain_name,))
            row = cur.fetchone()
            staff_id = row[0] if row else None
    with conn.cursor() as cur:
        cur.execute(
            """insert into supplier_report_token (complaint_id, supplier_name, requested_by_staff_id)
               values (%s, %s, %s) returning token;""",
            (complaint_id, supplier_name, staff_id),
        )
        token = cur.fetchone()[0]
    conn.commit()
    return f"{SUPPLIER_PORTAL_BASE_URL}/?token={token}"


def draft_customer_reply_email(client, report, compensation_qty):
    root_causes_text = "; ".join(report.get("Root cause") or []) or "(chưa xác định)"
    capa_text = "; ".join(report.get("CAPA") or []) or "(chưa xác định)"
    recorded_by = report.get("Recorded By") or ""
    signer_name = recorded_by.split(" (")[0].strip() if recorded_by else "Customer Service Team"
    prompt = f"""Bạn là nhân viên CS (Customer Service) tại 1 nhà máy sản xuất nhãn/bao bì apparel branding.
Viết 1 email TIẾNG ANH, chuyên nghiệp, lịch sự, trả lời complaint của khách hàng, dựa trên thông tin sau:

Mô tả sự cố: {report.get("Issue Description") or "(none)"}
Sản phẩm: {report.get("Product") or "(không rõ)"}
SO/PO: {report.get("SO/PO") or "(none)"}
Loại lỗi (Defect): {report.get("Defect") or "(chưa xác định)"}
Nguyên nhân gốc (Root Cause): {root_causes_text}
Hành động khắc phục (CAPA) đã/đang triển khai: {capa_text}
Số lượng sản phẩm bị lỗi: {report.get("Defect Quantity") or "(không rõ)"}
Số lượng đề xuất bù hàng: {compensation_qty}
Người ký tên cuối email: {signer_name}

QUAN TRỌNG — TUYỆT ĐỐI KHÔNG được nhắc tên nhà cung cấp hay bất kỳ thông tin nào có thể nhận diện được nhà
cung cấp cụ thể trong email — chỉ nói chung chung kiểu "our production process" hoặc "our manufacturing team"
nếu cần nhắc tới nguồn gốc lỗi.

Cấu trúc email: cảm ơn khách hàng đã báo cáo → xác nhận đã điều tra → nêu rõ nguyên nhân và hành động khắc
phục đã/đang triển khai (viết lại tự nhiên bằng tiếng Anh, không liệt kê mã kỹ thuật thô) → đề xuất số lượng
bù hàng → lời xin lỗi và cam kết cải thiện. Giọng văn chuyên nghiệp, ấm áp, không quá cứng nhắc.

Trả lời CHÍNH XÁC theo format sau, không thêm chữ nào khác:
SUBJECT: <tiêu đề email, tiếng Anh>
BODY: <toàn bộ nội dung email, tiếng Anh, đầy đủ lời chào và ký tên đúng "{signer_name}" ở cuối>
"""
    resp = client.messages.create(
        model="claude-sonnet-5", max_tokens=1200,
        messages=[{"role": "user", "content": prompt}],
    )
    text = extract_text(resp).strip()
    subject_m = re.search(r"SUBJECT\s*:\s*(.*?)(?=\n\s*BODY\s*:|\Z)", text, re.IGNORECASE | re.DOTALL)
    body_m = re.search(r"BODY\s*:\s*(.*)", text, re.IGNORECASE | re.DOTALL)
    subject = subject_m.group(1).strip() if subject_m else "Re: your complaint"
    body = body_m.group(1).strip() if body_m else text
    return subject, body


def synthesize_answer(client, question, df):
    data_text = df.to_csv(index=False) if not df.empty else "(không có dữ liệu phù hợp)"
    prompt = f"""Câu hỏi: "{question}"
Dữ liệu truy vấn được từ database (dạng CSV):
{data_text}

Trả lời câu hỏi bằng tiếng Việt, ngắn gọn, dựa đúng vào dữ liệu trên.
Nếu dữ liệu không đủ để trả lời chắc chắn, nói rõ điều đó thay vì đoán.
Nếu phù hợp, đề xuất 1 hành động tiếp theo.
"""
    resp = client.messages.create(
        model="claude-sonnet-5", max_tokens=1200,
        messages=[{"role": "user", "content": prompt}],
    )
    return extract_text(resp).strip()


# ============================================================
# COMPLAINT REGISTER — luồng mới: NCC tự khai báo qua link chung, CS không cần duyệt, chỉ điền
# thêm Customer/Replacement Cost rồi tải Excel tổng hợp bất kỳ lúc nào.
# ============================================================
REGISTER_EXCEL_COLUMNS = [
    "Record Date", "Year-Q", "Vendor No.", "Vendor Name", "Customer", "Client",
    "Sales Order No.", "Purchase Order No.", "Item No.", "Product Group", "Order Qty", "Defect Qty",
    "Reason", "Remarks", "Solution", "Result",
    "Responsible Party", "Replacement Cost", "Currency", "Recorded By",
    "Prepared By", "Position", "Report Link",
]


def fetch_all_submissions(conn):
    with conn.cursor() as cur:
        cur.execute("""
            select ss.submission_id, ss.submitted_at, ss.record_date, ss.supplier_name_raw, ss.vendor_code,
                   ss.vendor_name_matched, ss.match_confidence, ss.sales_order_no, ss.purchase_order_no,
                   ss.item_no, ss.order_qty, ss.defect_qty, ss.description, ss.root_cause, ss.capa,
                   ss.capa_status, ss.defect_image_urls, ss.defect_video_url,
                   ss.customer_name, ss.client_code, ss.bear_the_claim, ss.replacement_cost,
                   ss.replacement_cost_currency, ss.product_group,
                   c.status as complaint_status, st.name as recorded_by_name,
                   ss.prepared_by, ss.prepared_by_position,
                   exists(
                       select 1 from capa_action a where a.complaint_id = c.complaint_id
                       and (a.verification_result is null or a.verification_result = 'Pending')
                   ) as has_pending_verification,
                   exists(
                       select 1 from capa_action a where a.complaint_id = c.complaint_id
                       and a.date_implemented is not null
                   ) as has_implemented_capa
            from supplier_submissions ss
            left join complaint c on c.source_submission_id = ss.submission_id
            left join cs_staff st on c.recorded_by = st.staff_id
            order by ss.submitted_at desc;
        """)
        cols = [d[0] for d in cur.description]
        rows = cur.fetchall()
    return cols, rows


def fetch_legacy_complaints_for_dashboard(conn, cols):
    """Complaint tạo tay qua tab 'Complaint mới' — trả về CÙNG thứ tự cột với fetch_all_submissions()
    (dùng chuỗi 'legacy:<complaint_id>' làm submission_id giả để Dashboard/card/link chi tiết phân
    biệt được và điều hướng đúng chỗ), để gộp chung 1 danh sách duy nhất cho Dashboard hiển thị."""
    with conn.cursor() as cur:
        cur.execute("""
            select c.complaint_id, c.date_opened, s.name, s.vendor_code, cu.name, p.name, c.so_po, c.lot_number,
                   c.quantity_inspected, c.quantity_affected, c.notes, c.status,
                   c.client_code, c.bear_the_claim, c.replacement_cost, c.replacement_cost_currency, p.product_group,
                   st.name as recorded_by_name,
                   (select string_agg(distinct x.rc_text, '; ') from (
                        select rc.root_cause as rc_text
                        from complaint_root_cause crc join root_cause_taxonomy rc on crc.root_cause_code = rc.root_cause_code
                        where crc.complaint_id = c.complaint_id
                        union
                        select rc2.root_cause from root_cause_taxonomy rc2 where rc2.root_cause_code = c.root_cause_code
                   ) x) as root_causes,
                   string_agg(distinct cat.capa_action, '; ') as capas,
                   bool_or(a.verification_result is null or a.verification_result = 'Pending') as has_pending_verification,
                   bool_or(a.date_implemented is not null) as has_implemented_capa
            from complaint c
            left join supplier s on c.supplier_id = s.supplier_id
            left join customer cu on c.customer_id = cu.customer_id
            left join product p on c.product_id = p.product_id
            left join cs_staff st on c.recorded_by = st.staff_id
            left join capa_action a on a.complaint_id = c.complaint_id
            left join capa_taxonomy cat on a.capa_code = cat.capa_code
            where c.source is distinct from 'Nhà cung cấp tự khai báo (link chung)'
            group by c.complaint_id, c.date_opened, s.name, s.vendor_code, cu.name, p.name, c.so_po, c.lot_number,
                     c.quantity_inspected, c.quantity_affected, c.notes, c.status,
                     c.client_code, c.bear_the_claim, c.replacement_cost, c.replacement_cost_currency, p.product_group,
                     st.name
            order by c.date_opened desc;
        """)
        raw_rows = cur.fetchall()

    idx = {name: i for i, name in enumerate(cols)}
    result = []
    for (complaint_id, date_opened, supplier_name, vendor_code_db, customer_name, product_name, so_po, lot_number,
         qty_inspected, qty_affected, notes, status, client_code_db, bear_the_claim_db,
         replacement_cost_db, replacement_cost_currency_db, product_group_val, recorded_by_name_val, root_causes, capas,
         has_pending_verif, has_implemented) in raw_rows:
        vendor_code = vendor_code_db or _fuzzy_match_lookup_code(conn, "vendor_lookup", "vendor_code", "vendor_name", supplier_name)
        row = [None] * len(cols)
        row[idx["submission_id"]] = f"legacy:{complaint_id}"
        row[idx["submitted_at"]] = datetime.combine(date_opened, datetime.min.time()) if date_opened else datetime.now()
        row[idx["record_date"]] = date_opened
        row[idx["supplier_name_raw"]] = supplier_name or "(unknown)"
        row[idx["vendor_code"]] = vendor_code
        row[idx["vendor_name_matched"]] = supplier_name
        row[idx["match_confidence"]] = None
        row[idx["sales_order_no"]] = so_po or ""
        row[idx["purchase_order_no"]] = lot_number or ""
        row[idx["item_no"]] = product_name or ""
        row[idx["product_group"]] = product_group_val or ""
        row[idx["recorded_by_name"]] = recorded_by_name_val or ""
        row[idx["order_qty"]] = qty_inspected
        row[idx["defect_qty"]] = qty_affected
        row[idx["description"]] = notes or ""
        row[idx["root_cause"]] = root_causes or ""
        row[idx["capa"]] = capas or ""
        row[idx["capa_status"]] = ""
        row[idx["defect_image_urls"]] = []
        row[idx["defect_video_url"]] = None
        row[idx["customer_name"]] = customer_name
        row[idx["client_code"]] = client_code_db
        row[idx["bear_the_claim"]] = bear_the_claim_db
        row[idx["replacement_cost"]] = replacement_cost_db
        row[idx["replacement_cost_currency"]] = replacement_cost_currency_db
        row[idx["complaint_status"]] = status
        row[idx["has_pending_verification"]] = bool(has_pending_verif)
        row[idx["has_implemented_capa"]] = bool(has_implemented)
        result.append(tuple(row))
    return result


def update_submission_cs_fields(conn, submission_id, customer_name, client_code, bear_the_claim,
                                 replacement_cost, replacement_cost_currency, recorded_by_staff_id=None):
    with conn.cursor() as cur:
        cur.execute(
            """update supplier_submissions
               set customer_name = %s, client_code = %s, bear_the_claim = %s,
                   replacement_cost = %s, replacement_cost_currency = %s
               where submission_id = %s;""",
            (customer_name or None, client_code or None, bear_the_claim,
             replacement_cost, replacement_cost_currency, submission_id),
        )
    conn.commit()
    with conn.cursor() as cur:
        cur.execute(
            "select complaint_id, recorded_by, so_po from complaint where source_submission_id = %s;",
            (submission_id,),
        )
        row = cur.fetchone()
    if row:
        complaint_id, old_recorded_by, so_po = row
        if recorded_by_staff_id is not None and recorded_by_staff_id != old_recorded_by:
            with conn.cursor() as cur:
                cur.execute(
                    "update complaint set recorded_by = %s where complaint_id = %s;",
                    (recorded_by_staff_id, complaint_id),
                )
            conn.commit()
            notify_recorded_by_assignment(conn, complaint_id, recorded_by_staff_id, so_po)
        compute_and_update_complaint_status(conn, complaint_id)


def normalize_submission_row_for_excel(cols, row):
    idx = {name: i for i, name in enumerate(cols)}
    record_date = row[idx["record_date"]]
    year_q = f"{record_date.year}-Q{(record_date.month - 1) // 3 + 1}" if record_date else ""
    return {
        "Record Date": record_date,
        "Year-Q": year_q,
        "Vendor No.": row[idx["vendor_code"]] or "",
        "Vendor Name": row[idx["vendor_name_matched"]] or row[idx["supplier_name_raw"]],
        "Customer": row[idx["customer_name"]] or "",
        "Client": row[idx["client_code"]] or "",
        "Sales Order No.": row[idx["sales_order_no"]],
        "Purchase Order No.": row[idx["purchase_order_no"]],
        "Item No.": row[idx["item_no"]],
        "Product Group": row[idx["product_group"]] or "",
        "Order Qty": row[idx["order_qty"]],
        "Defect Qty": row[idx["defect_qty"]],
        "Reason": row[idx["description"]],
        "Remarks": row[idx["root_cause"]],
        "Solution": row[idx["capa"]],
        "Result": row[idx["complaint_status"]] or "Thiếu Customer",
        "Responsible Party": row[idx["bear_the_claim"]],
        "Replacement Cost": row[idx["replacement_cost"]],
        "Currency": row[idx["replacement_cost_currency"]],
        "Recorded By": row[idx["recorded_by_name"]] or "",
        "Prepared By": row[idx["prepared_by"]] or "",
        "Position": row[idx["prepared_by_position"]] or "",
        "_report_link": f"{REVIEW_APP_URL}/?submission_id={row[idx['submission_id']]}",
    }


def _fuzzy_match_lookup_code(conn, table, code_col, name_col, typed_name):
    """Đối chiếu gần đúng 1 tên (Vendor/Client) với danh sách mã chuẩn — dùng lại đúng nguyên lý
    match_vendor() bên supplier_portal.py, áp dụng cho cả complaint nhập tay (trước đây bị bỏ
    trống hoàn toàn, chưa từng tra cứu)."""
    if not typed_name:
        return None
    with conn.cursor() as cur:
        cur.execute(f"select {code_col}, {name_col} from {table};")
        rows = cur.fetchall()
    typed_norm = re.sub(r"\s+", " ", typed_name.strip().lower())

    # 1) Khớp tuyệt đối
    for code, name in rows:
        if re.sub(r"\s+", " ", (name or "").strip().lower()) == typed_norm:
            return code

    # 2) Khớp kiểu "tên ngắn nằm trong tên đầy đủ" (hoặc ngược lại) — xem giải thích chi tiết ở
    # match_vendor() bên supplier_portal.py, cùng 1 nguyên lý.
    if len(typed_norm) >= 4:
        substring_candidates = []
        for code, name in rows:
            name_norm = re.sub(r"\s+", " ", (name or "").strip().lower())
            if typed_norm in name_norm or name_norm in typed_norm:
                substring_candidates.append((code, len(name_norm)))
        if substring_candidates:
            substring_candidates.sort(key=lambda x: x[1])
            return substring_candidates[0][0]

    # 3) Khớp gần đúng theo tỉ lệ tương đồng chuỗi — dành cho lỗi chính tả nhẹ.
    best_code, best_ratio = None, 0.0
    for code, name in rows:
        ratio = difflib.SequenceMatcher(None, re.sub(r"\s+", " ", (name or "").strip().lower()), typed_norm).ratio()
        if ratio > best_ratio:
            best_ratio, best_code = ratio, code
    return best_code if best_ratio >= 0.55 else None


def fetch_legacy_complaints_for_excel(conn):
    """Complaint tạo qua tab 'Complaint mới' (nhập tay) — KHÔNG lấy lại các complaint đã tự tạo qua
    cầu nối từ supplier_submissions (source cụ thể), tránh hiện trùng 2 lần trong Excel."""
    with conn.cursor() as cur:
        cur.execute("""
            select c.complaint_id, c.date_opened, s.name, s.vendor_code, cu.name, c.so_po, c.lot_number, p.name,
                   c.quantity_inspected, c.quantity_affected, c.notes, c.status,
                   c.client_code, c.bear_the_claim, c.replacement_cost, c.replacement_cost_currency, p.product_group,
                   st.name as recorded_by_name,
                   (select string_agg(distinct x.rc_text, '; ') from (
                        select rc.root_cause as rc_text
                        from complaint_root_cause crc join root_cause_taxonomy rc on crc.root_cause_code = rc.root_cause_code
                        where crc.complaint_id = c.complaint_id
                        union
                        select rc2.root_cause from root_cause_taxonomy rc2 where rc2.root_cause_code = c.root_cause_code
                   ) x) as root_causes,
                   string_agg(distinct cat.capa_action, '; ') as capas
            from complaint c
            left join supplier s on c.supplier_id = s.supplier_id
            left join customer cu on c.customer_id = cu.customer_id
            left join product p on c.product_id = p.product_id
            left join cs_staff st on c.recorded_by = st.staff_id
            left join capa_action ca on ca.complaint_id = c.complaint_id
            left join capa_taxonomy cat on ca.capa_code = cat.capa_code
            where c.source is distinct from 'Nhà cung cấp tự khai báo (link chung)'
            group by c.complaint_id, c.date_opened, s.name, s.vendor_code, cu.name, c.so_po, c.lot_number, p.name,
                     c.quantity_inspected, c.quantity_affected, c.notes, c.status,
                     c.client_code, c.bear_the_claim, c.replacement_cost, c.replacement_cost_currency, p.product_group,
                     st.name
            order by c.date_opened desc;
        """)
        rows = cur.fetchall()

    result = []
    for (complaint_id, date_opened, supplier_name, vendor_code_db, customer_name, so_po, lot_number, product_name,
         qty_inspected, qty_affected, notes, status, client_code_db, bear_the_claim_db,
         replacement_cost_db, replacement_cost_currency_db, product_group_val, recorded_by_name_val, root_causes, capas) in rows:
        year_q = f"{date_opened.year}-Q{(date_opened.month - 1) // 3 + 1}" if date_opened else ""
        vendor_code = vendor_code_db or _fuzzy_match_lookup_code(conn, "vendor_lookup", "vendor_code", "vendor_name", supplier_name)
        client_code = client_code_db or _fuzzy_match_lookup_code(conn, "client_lookup", "client_code", "client_name", customer_name)
        result.append({
            "Record Date": date_opened,
            "Year-Q": year_q,
            "Vendor No.": vendor_code or "",
            "Vendor Name": supplier_name or "",
            "Customer": customer_name or "",
            "Client": client_code or "",
            "Sales Order No.": so_po or "",
            "Purchase Order No.": lot_number or "",
            "Item No.": product_name or "",
            "Product Group": product_group_val or "",
            "Order Qty": qty_inspected,
            "Defect Qty": qty_affected,
            "Reason": notes or "",
            "Remarks": root_causes or "",
            "Solution": capas or "",
            "Result": STATUS_LABELS.get(status, status or ""),
            "Responsible Party": bear_the_claim_db or "",
            "Replacement Cost": replacement_cost_db,
            "Currency": replacement_cost_currency_db or "",
            "Recorded By": recorded_by_name_val or "",
            "_report_link": f"{REVIEW_APP_URL}/?submission_id=legacy:{complaint_id}",
        })
    return result


def build_register_excel(normalized_rows):
    """Xuất Excel tổng hợp — normalized_rows là list các dict đã chuẩn hóa (từ
    normalize_submission_row_for_excel hoặc fetch_legacy_complaints_for_excel), gộp chung từ cả
    2 nguồn: luồng NCC tự khai báo VÀ complaint tạo tay qua tab 'Complaint mới'."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Supplier Return Register"

    n_cols = len(REGISTER_EXCEL_COLUMNS)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    title_cell = ws.cell(row=1, column=1, value="Supplier Return Register")
    title_cell.font = Font(bold=True, size=16, color="185FA5")
    title_cell.alignment = Alignment(horizontal="left")
    ws.row_dimensions[1].height = 28

    header_row_num = 3
    for c, col_name in enumerate(REGISTER_EXCEL_COLUMNS, start=1):
        ws.cell(row=header_row_num, column=c, value=col_name)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="185FA5", end_color="185FA5", fill_type="solid")
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=header_row_num, column=c)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for norm_row in normalized_rows:
        excel_row = [norm_row.get(col, "") for col in REGISTER_EXCEL_COLUMNS]
        report_link = norm_row.get("_report_link")
        excel_row[-1] = "View detail" if report_link else ""
        ws.append(excel_row)
        r = ws.max_row
        for c in range(1, n_cols + 1):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if report_link:
            link_cell = ws.cell(row=r, column=len(REGISTER_EXCEL_COLUMNS))
            link_cell.hyperlink = report_link
            link_cell.font = Font(color="185FA5", underline="single")

    for c in range(1, len(REGISTER_EXCEL_COLUMNS) + 1):
        col_letter = get_column_letter(c)
        max_len = max(
            [len(str(REGISTER_EXCEL_COLUMNS[c - 1]))]
            + [len(str(ws.cell(row=r, column=c).value or "")) for r in range(header_row_num + 1, ws.max_row + 1)]
        )
        ws.column_dimensions[col_letter].width = min(max_len + 3, 45)

    # Bật filter cho đúng hàng tiêu đề (row 3) — cho phép lọc/sắp xếp từng cột khi mở file Excel.
    last_col_letter = get_column_letter(n_cols)
    ws.auto_filter.ref = f"A{header_row_num}:{last_col_letter}{ws.max_row}"
    ws.freeze_panes = f"A{header_row_num + 1}"  # cố định luôn tiêu đề khi cuộn xuống, tiện lọc/xem

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _fetch_submission_by_id(conn, submission_id):
    with conn.cursor() as cur:
        cur.execute("""
            select ss.submission_id, ss.submitted_at, ss.record_date, ss.supplier_name_raw, ss.vendor_code,
                   ss.vendor_name_matched, ss.match_confidence, ss.sales_order_no, ss.purchase_order_no,
                   ss.item_no, ss.order_qty, ss.defect_qty, ss.description, ss.root_cause, ss.capa,
                   ss.defect_image_urls, ss.defect_video_url,
                   ss.customer_name, ss.client_code, ss.bear_the_claim, ss.replacement_cost,
                   ss.replacement_cost_currency, ss.product_group, c.status as complaint_status, c.complaint_id,
                   c.recorded_by, st.name as recorded_by_name,
                   ss.prepared_by, ss.prepared_by_position, ss.signature_image_url
            from supplier_submissions ss
            left join complaint c on c.source_submission_id = ss.submission_id
            left join cs_staff st on c.recorded_by = st.staff_id
            where ss.submission_id = %s;
        """, (submission_id,))
        row = cur.fetchone()
        if not row:
            return None
        cols = [d[0] for d in cur.description]
    return dict(zip(cols, row))


def export_submission_word(sub):
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(4)

    title = doc.add_heading("Supplier Quality Report — Detail", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def field(label, value):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.add_run(f"{label}: ").bold = True
        p.add_run(str(value) if value not in (None, "") else "")

    doc.add_heading("Order Info", level=1)
    field("Supplier (as reported)", sub["supplier_name_raw"])
    field("Vendor No.", sub["vendor_code"] or "unmatched)")
    field("Vendor Name", sub["vendor_name_matched"])
    field("Sales Order No.", sub["sales_order_no"])
    field("Purchase Order No.", sub["purchase_order_no"])
    field("Item No.", sub["item_no"])
    field("Product Group", sub.get("product_group"))
    field("Recorded By", sub.get("recorded_by_name"))
    field("Order Qty", sub["order_qty"])
    field("Defect Qty", sub["defect_qty"])
    field("Customer", sub["customer_name"])
    field("Client Code", sub["client_code"])

    doc.add_heading("Issue Details", level=1)
    field("Description", sub["description"])
    field("Root Cause", sub["root_cause"])
    field("CAPA", sub["capa"])
    field("Complaint Status", sub.get("complaint_status") or "Thiếu Customer")
    field("Responsible Party", sub["bear_the_claim"])
    field("Replacement Cost", f"{sub['replacement_cost']} {sub['replacement_cost_currency']}" if sub["replacement_cost"] is not None else "")

    image_urls = sub.get("defect_image_urls") or []
    if image_urls:
        doc.add_heading("Photos", level=1)
        for url in image_urls:
            try:
                img_resp = requests.get(url, timeout=15)
                if img_resp.status_code == 200:
                    doc.add_picture(io.BytesIO(img_resp.content), width=Inches(4))
                    doc.add_paragraph()
            except Exception:
                pass

    if sub.get("defect_video_url"):
        doc.add_paragraph().add_run(f"Video: {sub['defect_video_url']}").italic = True

    doc.add_paragraph()
    doc.add_heading("IV. Prepared By", level=1)
    field("Record Date", sub.get("record_date"))
    if sub.get("signature_image_url"):
        try:
            sig_resp = requests.get(sub["signature_image_url"], timeout=15)
            if sig_resp.status_code == 200:
                doc.add_picture(io.BytesIO(sig_resp.content), width=Inches(1.5))
        except Exception:
            pass
    field("Name", sub.get("prepared_by"))
    field("Position", sub.get("prepared_by_position"))

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def export_submission_pdf(sub):
    """PDF luôn xuất bằng TIẾNG ANH — font mặc định của reportlab không hỗ trợ tiếng Việt đầy đủ,
    dùng tiếng Việt sẽ lỗi hiển thị dấu."""
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter)
    styles = getSampleStyleSheet()
    story = [Paragraph("Supplier Quality Report", styles["Title"]), Spacer(1, 12)]

    def field(label, value):
        story.append(Paragraph(f"<b>{label}:</b> {value if value not in (None, '') else '-'}", styles["Normal"]))
        story.append(Spacer(1, 4))

    story.append(Paragraph("Order Info", styles["Heading2"]))
    field("Supplier (as reported)", sub["supplier_name_raw"])
    field("Vendor No.", sub["vendor_code"] or "(unmatched)")
    field("Vendor Name", sub["vendor_name_matched"])
    field("Sales Order No.", sub["sales_order_no"])
    field("Purchase Order No.", sub["purchase_order_no"])
    field("Item No.", sub["item_no"])
    field("Product Group", sub.get("product_group"))
    field("Recorded By", sub.get("recorded_by_name"))
    field("Order Qty", sub["order_qty"])
    field("Defect Qty", sub["defect_qty"])
    field("Customer", sub["customer_name"])
    field("Client Code", sub["client_code"])

    story.append(Spacer(1, 8))
    story.append(Paragraph("Issue Details", styles["Heading2"]))
    field("Description", sub["description"])
    field("Root Cause", sub["root_cause"])
    field("CAPA", sub["capa"])
    field("Complaint Status", sub.get("complaint_status") or "Thiếu Customer")
    field("Responsible Party", sub["bear_the_claim"])
    field("Replacement Cost", f"{sub['replacement_cost']} {sub['replacement_cost_currency']}" if sub["replacement_cost"] is not None else "-")

    image_urls = sub.get("defect_image_urls") or []
    if image_urls:
        story.append(Spacer(1, 8))
        story.append(Paragraph("Photos", styles["Heading2"]))
        for url in image_urls:
            try:
                img_resp = requests.get(url, timeout=15)
                if img_resp.status_code == 200:
                    story.append(RLImage(io.BytesIO(img_resp.content), width=4 * inch, height=3 * inch, kind="proportional"))
                    story.append(Spacer(1, 8))
            except Exception:
                pass

    story.append(Spacer(1, 12))
    story.append(Paragraph("Prepared By", styles["Heading2"]))
    field("Record Date", sub.get("record_date"))
    if sub.get("signature_image_url"):
        try:
            sig_resp = requests.get(sub["signature_image_url"], timeout=15)
            if sig_resp.status_code == 200:
                story.append(RLImage(io.BytesIO(sig_resp.content), width=1.5 * inch, height=1 * inch, kind="proportional"))
                story.append(Spacer(1, 4))
        except Exception:
            pass
    field("Name", sub.get("prepared_by"))
    field("Position", sub.get("prepared_by_position"))

    doc.build(story)
    return buf.getvalue()


def get_or_create_customer(conn, name):
    if not name:
        return None
    with conn.cursor() as cur:
        cur.execute("select customer_id from customer where name = %s limit 1;", (name,))
        row = cur.fetchone()
        if row:
            return row[0]
        cur.execute("insert into customer (name) values (%s) returning customer_id;", (name,))
        new_id = cur.fetchone()[0]
    conn.commit()
    return new_id


def get_or_create_supplier(conn, name, vendor_code=None):
    if not name:
        return None
    with conn.cursor() as cur:
        cur.execute("select supplier_id from supplier where name = %s limit 1;", (name,))
        row = cur.fetchone()
        if row:
            if vendor_code:
                cur.execute(
                    "update supplier set vendor_code = %s where supplier_id = %s and vendor_code is null;",
                    (vendor_code, row[0]),
                )
            conn.commit()
            return row[0]
        cur.execute(
            "insert into supplier (name, vendor_code) values (%s, %s) returning supplier_id;",
            (name, vendor_code),
        )
        new_id = cur.fetchone()[0]
    conn.commit()
    return new_id


def get_or_create_product(conn, item_no_text):
    if not item_no_text:
        return None
    try:
        with conn.cursor() as cur:
            cur.execute("select product_id from product where name = %s limit 1;", (item_no_text,))
            row = cur.fetchone()
            if row:
                return row[0]
            cur.execute(
                "insert into product (name, product_group) values (%s, 'OTHER') returning product_id;",
                (item_no_text,),
            )
            new_id = cur.fetchone()[0]
        conn.commit()
        return new_id
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return None


def ensure_client_lookup_exists(conn, client_code, fallback_name=None):
    """complaint.client_code có ràng buộc khóa ngoại tới client_lookup — nếu mã chưa tồn tại (ví
    dụ khi import từ Excel cũ, có mã chưa từng nạp vào client_lookup), tự tạo tạm 1 dòng để
    tránh lỗi insert. Reviewer có thể sửa tên chuẩn sau nếu cần."""
    if not client_code:
        return
    with conn.cursor() as cur:
        cur.execute("select 1 from client_lookup where client_code = %s;", (client_code,))
        if cur.fetchone():
            return
        cur.execute(
            "insert into client_lookup (client_code, client_name, country_code) values (%s, %s, '');",
            (client_code, fallback_name or client_code),
        )
    conn.commit()


def parse_replacement_cost_text(text):
    """Tách 'VND 2,000,000' / 'Usd 112.17' thành (số tiền, đơn vị) — trả (None, None) nếu không
    đọc được hoặc trống."""
    if not text or not isinstance(text, str):
        return None, None
    text = text.strip()
    m = re.match(r"([A-Za-z]+)\s*([\d,\.]+)", text)
    if not m:
        return None, None
    currency_raw, amount_raw = m.group(1).upper(), m.group(2).replace(",", "")
    currency = "VND" if currency_raw == "VND" else "USD"
    try:
        return float(amount_raw), currency
    except ValueError:
        return None, None


def parse_multi_qty_text(text):
    """Cộng dồn số lượng nếu 1 ô có nhiều dòng (nhiều item trong cùng 1 complaint), ví dụ
    '25,717\\n25,717' → 51434. Trả 0 nếu không đọc được."""
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return 0
    parts = str(text).replace(",", "").split("\n")
    total = 0
    for p in parts:
        p = p.strip()
        if p:
            try:
                total += int(float(p))
            except ValueError:
                pass
    return total


def find_orphaned_submissions(conn):
    """Quét supplier_submissions chưa có complaint tương ứng (do lỗi constraint database khiến
    bridge_to_legacy_tables() âm thầm thất bại trước khi được sửa) — trả về danh sách cần khôi phục."""
    with conn.cursor() as cur:
        cur.execute("""
            select ss.submission_id, ss.record_date, ss.supplier_name_raw, ss.vendor_name_matched,
                   ss.sales_order_no, ss.purchase_order_no, ss.order_qty, ss.defect_qty,
                   ss.description, ss.root_cause, ss.capa
            from supplier_submissions ss
            left join complaint c on c.source_submission_id = ss.submission_id
            where c.complaint_id is null
            order by ss.submitted_at asc;
        """)
        return cur.fetchall()


def repair_orphaned_submission(conn, ai_client, submission_id, record_date, supplier_name_raw,
                                vendor_name_matched, sales_order_no, purchase_order_no,
                                order_qty, defect_qty, description, root_cause_text, capa_text):
    """Tạo lại đúng complaint + 3 đề xuất taxonomy (Defect/Root Cause/CAPA) cho 1 submission bị mồ
    côi — làm y hệt logic bridge_to_legacy_tables() bên supplier_portal.py."""
    supplier_id = get_or_create_supplier(conn, vendor_name_matched or supplier_name_raw)

    def _classify_safe(text, kind):
        try:
            taxonomy_list = load_taxonomy_list(conn, kind)
            return classify_description(ai_client, text, taxonomy_list, kind)
        except Exception:
            try:
                conn.rollback()
            except Exception:
                pass
            return None

    defect_result = _classify_safe(description, "Defect")
    rc_result = _classify_safe(root_cause_text, "Root Cause")
    capa_result = _classify_safe(capa_text, "CAPA")

    with conn.cursor() as cur:
        cur.execute(
            """insert into complaint
                   (date_opened, source, supplier_id, so_po, lot_number, quantity_inspected, quantity_affected,
                    notes, status, source_submission_id, bear_the_claim)
               values (%s, %s, %s, %s, %s, %s, %s, %s, 'Thiếu Customer', %s, %s)
               returning complaint_id;""",
            (record_date, "Nhà cung cấp tự khai báo (link chung)", supplier_id,
             sales_order_no, purchase_order_no, order_qty, defect_qty, description, submission_id, "100% Nilorn"),
        )
        complaint_id = cur.fetchone()[0]

        for kind, text, result in [
            ("Defect", description, defect_result),
            ("Root Cause", root_cause_text, rc_result),
            ("CAPA", capa_text, capa_result),
        ]:
            closest = (result or {}).get("closest_existing_code") or (result or {}).get("matched_code")
            reasoning = (
                f"Typed by the supplier when submitting via the shared link (recovered from a technical issue). "
                f"AI suggestion: {(result or {}).get('reasoning', '(could not classify)')}"
            )
            cur.execute(
                """insert into taxonomy_suggestion
                       (complaint_id, suggestion_type, ai_suggested_name, ai_reasoning, closest_existing_code)
                   values (%s, %s, %s, %s, %s);""",
                (complaint_id, kind, text[:200], reasoning, closest),
            )
    conn.commit()
    return complaint_id


def find_complaints_missing_root_cause(conn, source_filter=None):
    """Quét complaint chưa có root_cause_code — dùng để bổ sung đề xuất Root Cause cho các
    complaint import từ nguồn không có sẵn cột này (ví dụ Excel cũ)."""
    with conn.cursor() as cur:
        if source_filter:
            cur.execute(
                "select complaint_id, notes from complaint where root_cause_code is null and source = %s;",
                (source_filter,),
            )
        else:
            cur.execute("select complaint_id, notes from complaint where root_cause_code is null;")
        return cur.fetchall()


def suggest_root_cause_for_complaint(conn, ai_client, complaint_id, description_text):
    """Cho AI đọc Description của 1 complaint đã tồn tại, tạo đề xuất Root Cause vào hàng chờ
    duyệt ở Duyệt Taxonomy — giữ đúng quy trình kiểm soát chất lượng chung, không tự gán."""
    if not description_text:
        return False
    try:
        taxonomy_list = load_taxonomy_list(conn, "Root Cause")
        result = classify_description(ai_client, description_text, taxonomy_list, "Root Cause")
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        result = None
    closest = (result or {}).get("closest_existing_code") or (result or {}).get("matched_code")
    reasoning = (
        f"Added later (original complaint was missing Root Cause on import). "
        f"AI suggestion: {(result or {}).get('reasoning', '(could not classify)')}"
    )
    with conn.cursor() as cur:
        cur.execute(
            """insert into taxonomy_suggestion
                   (complaint_id, suggestion_type, ai_suggested_name, ai_reasoning, closest_existing_code)
               values (%s, 'Root Cause', %s, %s, %s);""",
            (complaint_id, description_text[:200], reasoning, closest),
        )
    conn.commit()
    return True


def update_legacy_complaint_cs_fields(conn, complaint_id, customer_name, client_code, bear_the_claim,
                                       replacement_cost, replacement_cost_currency, complaint_validity=None,
                                       recorded_by_staff_id=None):
    customer_id = get_or_create_customer(conn, customer_name) if customer_name else None
    with conn.cursor() as cur:
        cur.execute("select recorded_by, so_po from complaint where complaint_id = %s;", (complaint_id,))
        old_recorded_by, so_po = cur.fetchone()
    with conn.cursor() as cur:
        cur.execute(
            """update complaint
               set customer_id = %s, client_code = %s, bear_the_claim = %s,
                   replacement_cost = %s, replacement_cost_currency = %s, complaint_validity = %s
               where complaint_id = %s;""",
            (customer_id, client_code or None, bear_the_claim, replacement_cost,
             replacement_cost_currency, complaint_validity, complaint_id),
        )
    conn.commit()
    if recorded_by_staff_id is not None and recorded_by_staff_id != old_recorded_by:
        with conn.cursor() as cur:
            cur.execute(
                "update complaint set recorded_by = %s where complaint_id = %s;",
                (recorded_by_staff_id, complaint_id),
            )
        conn.commit()
        notify_recorded_by_assignment(conn, complaint_id, recorded_by_staff_id, so_po)
    compute_and_update_complaint_status(conn, complaint_id)


def render_legacy_complaint_detail_page(conn, complaint_id):
    """Trang chi tiết cho complaint tạo tay qua 'Complaint mới' — giờ có đầy đủ Customer/Client
    Code/Responsible Party/Replacement Cost y hệt luồng NCC tự khai báo, tự điền sẵn nếu đã có dữ liệu
    (Customer từ lúc nhập complaint, Vendor No./Client Code tự dò bằng đối chiếu gần đúng)."""
    if st.session_state.get("selected_submission_id"):
        if st.button("Back to Dashboard"):
            st.session_state.selected_submission_id = None
            st.session_state.current_page = "dashboard"
            st.rerun()

    with conn.cursor() as cur:
        cur.execute("""
            select c.date_opened, s.name, s.vendor_code, cu.name, p.name, c.so_po, c.lot_number,
                   c.quantity_inspected, c.quantity_affected, c.notes, c.status,
                   c.client_code, c.bear_the_claim, c.replacement_cost, c.replacement_cost_currency,
                   c.complaint_validity, coalesce(p.product_group, ss.product_group) as product_group_val,
                   c.recorded_by, st.name as recorded_by_name,
                   (select string_agg(distinct x.rc_text, '; ') from (
                        select rc.root_cause as rc_text
                        from complaint_root_cause crc join root_cause_taxonomy rc on crc.root_cause_code = rc.root_cause_code
                        where crc.complaint_id = c.complaint_id
                        union
                        select rc2.root_cause from root_cause_taxonomy rc2 where rc2.root_cause_code = c.root_cause_code
                   ) x) as root_causes,
                   string_agg(distinct cat.capa_action, '; ') as capas
            from complaint c
            left join supplier s on c.supplier_id = s.supplier_id
            left join customer cu on c.customer_id = cu.customer_id
            left join product p on c.product_id = p.product_id
            left join supplier_submissions ss on ss.submission_id = c.source_submission_id
            left join cs_staff st on c.recorded_by = st.staff_id
            left join capa_action a on a.complaint_id = c.complaint_id
            left join capa_taxonomy cat on a.capa_code = cat.capa_code
            where c.complaint_id = %s
            group by c.complaint_id, c.date_opened, s.name, s.vendor_code, cu.name, p.name, c.so_po, c.lot_number,
                     c.quantity_inspected, c.quantity_affected, c.notes, c.status,
                     c.client_code, c.bear_the_claim, c.replacement_cost, c.replacement_cost_currency,
                     c.complaint_validity, ss.product_group, p.product_group, c.recorded_by, st.name;
        """, (complaint_id,))
        row = cur.fetchone()
    if not row:
        st.error("Complaint not found.")
        return
    (date_opened, supplier_name, vendor_code_db, customer_name, product_name, so_po, lot_number,
     qty_inspected, qty_affected, notes, status, client_code_db, bear_the_claim_db,
     replacement_cost_db, replacement_cost_currency_db, complaint_validity_db, product_group_val,
     recorded_by_db, recorded_by_name, root_causes, capas) = row

    vendor_code = vendor_code_db or _fuzzy_match_lookup_code(conn, "vendor_lookup", "vendor_code", "vendor_name", supplier_name)
    client_code_suggested = client_code_db or _fuzzy_match_lookup_code(
        conn, "client_lookup", "client_code", "client_name", customer_name
    )

    st.title("Complaint Detail (manual entry)")
    st.caption(f"Date occurred: {date_opened}")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### Order Information")
        st.write(f"**Supplier:** {supplier_name or '(unknown)'}")
        st.write(f"**Vendor No.:** {vendor_code or '(unmatched)'}")
        st.write(f"**Product (Item No.):** {product_name or '(unknown)'}")
        st.write(f"**Product Group:** {product_group_val or '-'}")
        st.write(f"**Recorded by:** {recorded_by_name or 'not assigned)'}")
        st.write(f"**Sales Order No.:** {so_po or '(empty)'}")
        st.write(f"**Purchase Order No.:** {lot_number or '(empty)'}")
        st.write(f"**Order Qty:** {qty_inspected} | **Defect Qty:** {qty_affected}")
    with col2:
        st.markdown("#### Issue Details")
        st.write(f"**Description:** {notes or '(empty)'}")
        st.write(f"**Assessment:** {complaint_validity_db or COMPLAINT_VALIDITY_OPTIONS[0]}")
        st.write(f"**Root Cause:** {root_causes or '(not yet approved / none)'}")
        st.write(f"**CAPA:** {capas or '(not yet approved / none)'}")

    st.markdown("---")
    with zone_card("amber"):
        section_header("✏️", "Additional Information — Customer & Replacement Cost", "amber")
        with st.form(f"legacy_edit_form_{complaint_id}"):
            e_col1, e_col2 = st.columns(2)
            with e_col1:
                customer_name_in = st.text_input("Customer", value=customer_name or "")
                client_code_in = st.text_input("Client Code", value=client_code_suggested or "")
                cs_staff_list_legacy = fetch_cs_staff(conn)
                staff_labels_legacy = ["not selected --"] + [f"{name} ({role})" for _, name, role in cs_staff_list_legacy]
                current_recorded_by_label_legacy = next(
                    (f"{name} ({role})" for _, name, role in cs_staff_list_legacy if name == recorded_by_name),
                    None,
                )
                recorded_by_choice_legacy = st.selectbox(
                    "Recorded by", staff_labels_legacy,
                    index=staff_labels_legacy.index(current_recorded_by_label_legacy) if current_recorded_by_label_legacy in staff_labels_legacy else 0,
                    help="Select the CS in charge — notification emails will also go to this person, in addition to the shared CS inbox.",
                )
            with e_col2:
                bear_options = ["100% Nilorn", "100% Customer", "50/50", "Other"]
                bear_the_claim_in = st.selectbox(
                    "Responsible Party", bear_options,
                    index=bear_options.index(bear_the_claim_db) if bear_the_claim_db in bear_options else 0,
                )
                rc_col1, rc_col2 = st.columns([2, 1])
                with rc_col1:
                    replacement_cost_in = st.number_input(
                        "Replacement Cost", min_value=0.0, value=float(replacement_cost_db or 0.0), step=1.0,
                    )
                with rc_col2:
                    currency_in = st.selectbox(
                        "Currency", ["USD", "VND"],
                        index=["USD", "VND"].index(replacement_cost_currency_db or "USD"),
                    )
                no_cost_in = st.checkbox(
                    "No replacement cost for this complaint",
                    value=(replacement_cost_db == 0),
                    help="Check this to explicitly confirm zero, to avoid confusion with 'not filled in' — if left blank/0 WITHOUT checking, the system still treats it as undetermined and keeps the reminder badge.",
                )
            with validity_radio_box():
                complaint_validity_in = st.radio(
                    "Assessment", COMPLAINT_VALIDITY_OPTIONS,
                    index=COMPLAINT_VALIDITY_OPTIONS.index(complaint_validity_db) if complaint_validity_db in COMPLAINT_VALIDITY_OPTIONS else 0,
                    horizontal=True,
                    help="Update this if further investigation reveals a different conclusion than the initial entry.",
                )
            if st.form_submit_button("Save"):
                final_cost_in = 0.0 if no_cost_in else (replacement_cost_in or None)
                recorded_by_staff_id_legacy = (
                    None if recorded_by_choice_legacy.startswith("--")
                    else cs_staff_list_legacy[staff_labels_legacy.index(recorded_by_choice_legacy) - 1][0]
                )
                update_legacy_complaint_cs_fields(
                    conn, complaint_id, customer_name_in, client_code_in,
                    bear_the_claim_in, final_cost_in, currency_in, complaint_validity_in,
                    recorded_by_staff_id_legacy,
                )
                st.success("Saved.")
                st.rerun()

    st.markdown("---")
    full_report_legacy = fetch_complaint_full_report(conn, complaint_id)
    col_word, col_pdf = st.columns(2)
    with col_word:
        st.download_button(
            "Download Word",
            data=export_word_report(full_report_legacy),
            file_name=f"Complaint_{complaint_id}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    with col_pdf:
        st.download_button(
            "⬇️ Download PDF (English)",
            data=export_complaint_report_pdf(full_report_legacy),
            file_name=f"Complaint_{complaint_id}.pdf",
            mime="application/pdf",
        )

    with st.expander("Replace defect photo"):
        if full_report_legacy.get("Defect Photo"):
            render_zoomable_image(full_report_legacy["Defect Photo"], width=220, caption="Current photo")
        new_defect_photo_legacy = st.file_uploader(
            "Defect photo", type=["png", "jpg", "jpeg"],
            key=f"attach_defect_photo_legacy_{complaint_id}",
        )
        if new_defect_photo_legacy is not None and st.button("Save photo", key=f"btn_save_defect_photo_legacy_{complaint_id}"):
            photo_b64_legacy = base64.b64encode(new_defect_photo_legacy.getvalue()).decode("utf-8")
            with conn.cursor() as cur:
                cur.execute("update complaint set defect_photo = %s where complaint_id = %s;", (photo_b64_legacy, complaint_id))
            conn.commit()
            st.success("Photo saved.")
            st.rerun()

    st.markdown("---")
    missing_tags_now = compute_and_update_complaint_status(conn, complaint_id)
    badges_html_now = "".join(status_badge_html(t) for t in missing_tags_now) if missing_tags_now else status_badge_html("Closed")
    st.markdown(
        f"Current status (auto-calculated based on missing data): {badges_html_now}",
        unsafe_allow_html=True,
    )


def render_submission_detail_page(conn, submission_id):
    sub = _fetch_submission_by_id(conn, submission_id)
    if not sub:
        st.error("Report not found.")
        return

    if st.session_state.get("selected_submission_id"):
        if st.button("Back to Dashboard"):
            st.session_state.selected_submission_id = None
            st.session_state.current_page = "dashboard"
            st.rerun()

    st.title("Report Detail")
    st.caption(f"Record Date: {sub.get('record_date')}")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### Order Information")
        st.write(f"**Supplier (self-reported):** {sub['supplier_name_raw']}")
        st.write(f"**Vendor No.:** {sub['vendor_code'] or '(unmatched)'}")
        st.write(f"**Vendor Name:** {sub['vendor_name_matched'] or '-'}")
        st.write(f"**Sales Order No.:** {sub['sales_order_no']}")
        st.write(f"**Purchase Order No.:** {sub['purchase_order_no']}")
        st.write(f"**Item No.:** {sub['item_no']}")
        st.write(f"**Product Group:** {sub.get('product_group') or '-'}")
        st.write(f"**Recorded by:** {sub.get('recorded_by_name') or 'not assigned)'}")
        st.write(f"**Prepared by:** {sub.get('prepared_by') or '-'} ({sub.get('prepared_by_position') or '-'})")
        if sub.get("signature_image_url"):
            st.image(sub["signature_image_url"], width=150, caption="Signature")
        st.write(f"**Order Qty:** {sub['order_qty']} | **Defect Qty:** {sub['defect_qty']}")
    with col2:
        st.markdown("#### Issue Details")
        st.write(f"**Description:** {sub['description']}")
        st.write(f"**Root Cause:** {sub['root_cause']}")
        st.write(f"**CAPA:** {sub['capa']}")
        st.write(f"**Complaint Status:** {sub.get('complaint_status') or 'Thiếu Customer'}")

    image_urls = sub.get("defect_image_urls") or []
    if image_urls:
        st.markdown("#### 📷 Defect Photos")
        img_cols = st.columns(len(image_urls))
        for col, url in zip(img_cols, image_urls):
            with col:
                render_zoomable_image(requests.get(url, timeout=15).content, width=200)

    if sub.get("defect_video_url"):
        st.markdown("#### 🎥 Defect Video")
        st.video(sub["defect_video_url"])

    st.markdown("---")
    with zone_card("amber"):
        section_header("✏️", "Additional Information — Customer & Replacement Cost", "amber")
        with st.form(f"edit_form_{submission_id}"):
            e_col1, e_col2 = st.columns(2)
            with e_col1:
                customer_name_in = st.text_input("Customer", value=sub.get("customer_name") or "")
                client_code_in = st.text_input("Client Code", value=sub.get("client_code") or "")
                cs_staff_list_sub = fetch_cs_staff(conn)
                staff_labels_sub = ["not selected --"] + [f"{name} ({role})" for _, name, role in cs_staff_list_sub]
                current_recorded_by_label = None
                if sub.get("recorded_by_name"):
                    current_recorded_by_label = next(
                        (lbl for _, name, role in cs_staff_list_sub for lbl in [f"{name} ({role})"] if name == sub["recorded_by_name"]),
                        None,
                    )
                recorded_by_choice_sub = st.selectbox(
                    "Recorded by", staff_labels_sub,
                    index=staff_labels_sub.index(current_recorded_by_label) if current_recorded_by_label in staff_labels_sub else 0,
                    help="Select the CS in charge — notification emails will also go to this person, in addition to the shared CS inbox.",
                )
            with e_col2:
                bear_the_claim_in = st.selectbox(
                    "Responsible Party",
                    ["100% Nilorn", "100% Customer", "50/50", "Other"],
                    index=["100% Nilorn", "100% Customer", "50/50", "Other"].index(sub.get("bear_the_claim") or "100% Nilorn")
                    if (sub.get("bear_the_claim") or "100% Nilorn") in ["100% Nilorn", "100% Customer", "50/50", "Other"] else 0,
                )
                rc_col1, rc_col2 = st.columns([2, 1])
                with rc_col1:
                    replacement_cost_in = st.number_input(
                        "Replacement Cost", min_value=0.0, value=float(sub.get("replacement_cost") or 0.0), step=1.0,
                    )
                with rc_col2:
                    currency_in = st.selectbox(
                        "Currency", ["USD", "VND"],
                        index=["USD", "VND"].index(sub.get("replacement_cost_currency") or "USD"),
                    )
                no_cost_in = st.checkbox(
                    "No replacement cost for this complaint",
                    value=(sub.get("replacement_cost") == 0),
                    help="Check this to explicitly confirm zero, to avoid confusion with 'not filled in' — if left blank/0 WITHOUT checking, the system still treats it as undetermined and keeps the reminder badge.",
                )
            if st.form_submit_button("Save"):
                final_cost_in = 0.0 if no_cost_in else (replacement_cost_in or None)
                recorded_by_staff_id_sub = (
                    None if recorded_by_choice_sub.startswith("--")
                    else cs_staff_list_sub[staff_labels_sub.index(recorded_by_choice_sub) - 1][0]
                )
                update_submission_cs_fields(
                    conn, submission_id, customer_name_in, client_code_in,
                    bear_the_claim_in, final_cost_in, currency_in, recorded_by_staff_id_sub,
                )
                st.success("Saved.")
                st.rerun()

    st.markdown("---")
    if sub.get("complaint_id"):
        missing_tags_now = compute_and_update_complaint_status(conn, sub["complaint_id"])
        badges_html_now = "".join(status_badge_html(t) for t in missing_tags_now) if missing_tags_now else status_badge_html("Closed")
        st.markdown(
            f"Current status (automatic): {badges_html_now}",
            unsafe_allow_html=True,
        )
    else:
        st.error(
            "This report has no linked complaint — status cannot be computed. Go to 'Data Lookup' to repair it."
        )

    with st.expander("🖼️ Replace defect photo/video"):
        st.caption(
            "Select new photo/video and click Update — this will FULLY REPLACE the current photo/video for this report (not additive). Leave blank to keep as-is."
        )
        new_images = st.file_uploader(
            "New photos (up to 3, leave blank to keep current)", type=["png", "jpg", "jpeg"],
            accept_multiple_files=True, key=f"replace_images_{submission_id}",
        )
        new_video = st.file_uploader(
            "New video (leave blank to keep current)", type=["mp4", "mov", "avi", "webm"],
            key=f"replace_video_{submission_id}",
        )
        if st.button("🔄 Update photo/video", key=f"btn_replace_media_{submission_id}"):
            if not new_images and not new_video:
                st.warning("No photo/video selected to replace.")
            else:
                try:
                    import json as _json
                    with st.spinner("Uploading..."):
                        final_image_urls = sub.get("defect_image_urls") or []
                        if new_images:
                            final_image_urls = []
                            for f in new_images[:3]:
                                url = upload_to_storage(f.getvalue(), f.name, f.type)
                                final_image_urls.append(url)
                        final_video_url = sub.get("defect_video_url")
                        if new_video:
                            final_video_url = upload_to_storage(new_video.getvalue(), new_video.name, new_video.type)
                        with conn.cursor() as cur:
                            cur.execute(
                                "update supplier_submissions set defect_image_urls = %s, defect_video_url = %s "
                                "where submission_id = %s;",
                                (_json.dumps(final_image_urls), final_video_url, submission_id),
                            )
                        conn.commit()
                    st.success("Media updated.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not upload: {e}")

    st.markdown("---")
    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        try:
            st.download_button(
                "Download Word", data=export_submission_word(sub),
                file_name=f"Report_{sub['sales_order_no']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        except Exception as e:
            st.error(f"Could not create Word file: {e}")
    with dl_col2:
        try:
            st.download_button(
                "⬇️ Download PDF (English)", data=export_submission_pdf(sub),
                file_name=f"Report_{sub['sales_order_no']}.pdf", mime="application/pdf",
            )
        except Exception as e:
            st.error(f"Could not create PDF: {e}")


# ============================================================
# GIAO DIỆN
# ============================================================
st.set_page_config(page_title="Factory Copilot", layout="wide")

# ------------------------------------------------------------
# Nền trang, font, và style chung — cho giao diện gọn gàng, chuyên nghiệp hơn: nền xám nhạt để
# card nổi bật lên, font dễ đọc, nút bấm/khung viền mềm mại hơn. / Page background, font, and
# shared styling — light gray page canvas so cards stand out, readable font, softer buttons/borders.
# ------------------------------------------------------------
st.markdown(
    """<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"]  {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
}
.stApp {
    background: linear-gradient(160deg, #f3f0fb 0%, #f7f8fc 55%);
}
[data-testid="stHeader"] {
    background: transparent;
}
[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 14px !important;
    box-shadow: 0 1px 3px rgba(20, 20, 30, 0.05);
}
.stButton > button {
    border-radius: 8px;
    font-weight: 500;
    transition: border-color 0.15s ease, color 0.15s ease;
}
.stButton > button:hover {
    border-color: #185fa5;
    color: #185fa5;
}
[data-testid="stTabs"] {
    background: #ffffff;
    border-radius: 12px;
    padding: 0.4rem 0.6rem 0;
    box-shadow: 0 1px 3px rgba(20, 20, 30, 0.05);
    margin-bottom: 0.5rem;
}
[data-testid="stTabs"] button[role="tab"] {
    font-weight: 500;
    font-size: 14.5px;
    padding: 10px 18px;
}
[data-testid="stTabs"] button[aria-selected="true"] {
    background: #eef2f7;
    border-radius: 8px 8px 0 0;
}
.stTextInput input, .stTextArea textarea, .stNumberInput input, .stDateInput input {
    border-radius: 8px;
}
.fc-badge {
    display: inline-block; padding: 2px 10px; border-radius: 12px;
    font-size: 12px; font-weight: 600; text-transform: uppercase;
    letter-spacing: 0.02em; border: 1px solid transparent;
}
/* Zone card dùng chung — nền màu rõ ràng (không trong suốt) + viền trái đậm, để phân vùng lớn
   thực sự nổi bật thay vì chỉ mờ mờ. Dùng chọn theo tiền tố class (key bắt đầu bằng "zone_<màu>_")
   nên áp dụng được cho BẤT KỲ container nào ở BẤT KỲ tab nào, không cần khai CSS riêng từng chỗ. */
div[class*="st-key-zone_blue_"] {
    background: #dfedfa; border-left: 4px solid #185fa5;
    border-radius: 0 12px 12px 0; padding: 0.4rem 1.2rem 1.1rem; margin-bottom: 0.6rem;
}
div[class*="st-key-zone_amber_"] {
    background: #f8e9cd; border-left: 4px solid #ba7517;
    border-radius: 0 12px 12px 0; padding: 0.4rem 1.2rem 1.1rem; margin-bottom: 0.6rem;
}
div[class*="st-key-zone_teal_"] {
    background: #d9f0e6; border-left: 4px solid #0f6e56;
    border-radius: 0 12px 12px 0; padding: 0.4rem 1.2rem 1.1rem; margin-bottom: 0.6rem;
}
div[class*="st-key-zone_coral_"] {
    background: #f9e2d8; border-left: 4px solid #993c1d;
    border-radius: 0 12px 12px 0; padding: 0.4rem 1.2rem 1.1rem; margin-bottom: 0.6rem;
}
div[class*="st-key-zone_gray_"] {
    background: #e9e7de; border-left: 4px solid #5f5e5a;
    border-radius: 0 12px 12px 0; padding: 0.4rem 1.2rem 1.1rem; margin-bottom: 0.6rem;
}
/* Ô chọn "Đánh giá ban đầu" (Lỗi thật / Lỗi khách hàng) — hiện thành 2 thẻ lớn rõ ràng thay vì
   nút radio nhỏ, xanh cho Nilorn, đỏ cho khách hàng, dùng chung 1 style cho mọi nơi xuất hiện ô này
   (form Complaint mới, mục Bổ sung thông tin ở trang chi tiết) nhờ đặt tên key cùng tiền tố. */
div[class*="st-key-validity_radio_"] div[data-testid="stRadio"] > div[role="radiogroup"] {
    display: flex; gap: 12px; margin-top: 2px;
}
div[class*="st-key-validity_radio_"] div[data-testid="stRadio"] label {
    flex: 1; border: 2px solid #e5e3da; border-radius: 10px; padding: 14px 18px !important;
    margin: 0 !important; cursor: pointer; transition: all 0.15s ease; background: #ffffff;
}
div[class*="st-key-validity_radio_"] div[data-testid="stRadio"] label > div:first-child {
    display: none;
}
div[class*="st-key-validity_radio_"] div[data-testid="stRadio"] label div[data-testid="stMarkdownContainer"] {
    text-align: center; width: 100%;
}
div[class*="st-key-validity_radio_"] div[data-testid="stRadio"] label div[data-testid="stMarkdownContainer"] p {
    font-size: 15px; font-weight: 600; margin: 0; color: #4a4a45;
}
div[class*="st-key-validity_radio_"] div[data-testid="stRadio"] label:nth-of-type(1) {
    border-color: #bcdcc0;
}
div[class*="st-key-validity_radio_"] div[data-testid="stRadio"] label:nth-of-type(1):hover {
    background: #f3f9ee;
}
div[class*="st-key-validity_radio_"] div[data-testid="stRadio"] label:nth-of-type(1):has(input:checked) {
    background: #eaf3de; border-color: #3b6d11; box-shadow: 0 0 0 1px #3b6d11;
}
div[class*="st-key-validity_radio_"] div[data-testid="stRadio"] label:nth-of-type(1):has(input:checked) p {
    color: #27500a;
}
div[class*="st-key-validity_radio_"] div[data-testid="stRadio"] label:nth-of-type(2) {
    border-color: #f0b8b8;
}
div[class*="st-key-validity_radio_"] div[data-testid="stRadio"] label:nth-of-type(2):hover {
    background: #fdf3f3;
}
div[class*="st-key-validity_radio_"] div[data-testid="stRadio"] label:nth-of-type(2):has(input:checked) {
    background: #fcebeb; border-color: #a32d2d; box-shadow: 0 0 0 1px #a32d2d;
}
div[class*="st-key-validity_radio_"] div[data-testid="stRadio"] label:nth-of-type(2):has(input:checked) p {
    color: #791f1f;
}
</style>""",
    unsafe_allow_html=True,
)

_validity_radio_counter = [0]


def validity_radio_box():
    """Context manager riêng cho ô 'Đánh giá ban đầu' — bọc trong 1 container có key mang tiền tố
    'validity_radio_' để CSS ở trên tự nhận diện và tô màu, dùng được ở bất kỳ đâu trong app mà
    không cần khai báo CSS riêng từng chỗ (giống nguyên lý zone_card())."""
    n = _validity_radio_counter[0]
    _validity_radio_counter[0] += 1
    return st.container(key=f"validity_radio_{n}")

KIND_BADGE_STYLE = {
    "Defect": ("#e6f1fb", "#0c447c", "#185fa5"),
    "Root Cause": ("#faeeda", "#854f0b", "#ba7517"),
    "CAPA": ("#e1f5ee", "#085041", "#0f6e56"),
}
STATUS_BADGE_STYLE = {
    "Chưa hoàn thành": ("#fcebeb", "#791f1f", "#a32d2d"),
    "Thiếu Customer": ("#fcebeb", "#791f1f", "#a32d2d"),
    "Thiếu Client": ("#fcebeb", "#791f1f", "#a32d2d"),
    "Thiếu Replacement Cost": ("#fcebeb", "#791f1f", "#a32d2d"),
    "Thiếu xác minh CAPA": ("#fcebeb", "#791f1f", "#a32d2d"),
    "Closed": ("#eaf3de", "#27500a", "#3b6d11"),
}


def render_badge(text, bg, color, border):
    return f'<span class="fc-badge" style="background:{bg};color:{color};border-color:{border};">{text}</span>'


def kind_badge_html(kind):
    bg, color, border = KIND_BADGE_STYLE.get(kind, ("#f1efe8", "#2c2c2a", "#5f5e5a"))
    return render_badge(kind, bg, color, border)


STATUS_BADGE_DISPLAY_LABELS = {
    "Thiếu Customer": "Missing Customer",
    "Thiếu Client": "Missing Client",
    "Thiếu Replacement Cost": "Missing Cost",
    "Thiếu xác minh CAPA": "CAPA Not Verified",
    "Closed": "Closed",
    "Chưa hoàn thành": "Not Closed",
}


def status_badge_html(status):
    bg, color, border = STATUS_BADGE_STYLE.get(status, ("#f1efe8", "#2c2c2a", "#5f5e5a"))
    display_text = STATUS_BADGE_DISPLAY_LABELS.get(status, status)
    return render_badge(display_text, bg, color, border)


SECTION_COLOR_STYLE = {
    "blue": ("#e6f1fb", "#0c447c"),
    "amber": ("#faeeda", "#854f0b"),
    "teal": ("#e1f5ee", "#085041"),
    "coral": ("#faece7", "#993c1d"),
    "gray": ("#f1efe8", "#2c2c2a"),
}

DASHBOARD_BADGE_STYLE = {
    "coral": ("#faece7", "#993c1d", "#f0997b"),
    "amber": ("#faeeda", "#854f0b", "#ef9f27"),
    "teal": ("#e1f5ee", "#085041", "#5dcaa5"),
    "gray": ("#f1efe8", "#2c2c2a", "#b4b2a9"),
}


def section_header(icon, title, color="gray"):
    """Tiêu đề khu vực có icon màu — dùng thay cho st.subheader() + st.markdown('---') để phân tách
    các khu vực rõ ràng hơn bằng màu sắc thay vì chỉ bằng đường kẻ ngang. / Colored section header —
    replaces st.subheader() + a plain divider, telling sections apart by color instead of just a line."""
    bg, text = SECTION_COLOR_STYLE.get(color, SECTION_COLOR_STYLE["gray"])
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:10px;margin:1.75rem 0 0.9rem;">'
        f'<span style="background:{bg};color:{text};min-width:34px;height:34px;border-radius:9px;'
        f'display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;">{icon}</span>'
        f'<span style="font-size:1.08rem;font-weight:600;color:var(--text-primary, #2c2c2a);">{title}</span>'
        f'</div>',
        unsafe_allow_html=True,
    )


def stat_card_html(label, value, color="gray"):
    bg, text = SECTION_COLOR_STYLE.get(color, SECTION_COLOR_STYLE["gray"])
    return (
        f'<div style="background:{bg};border-radius:10px;padding:0.9rem 1rem;">'
        f'<p style="font-size:12px;color:{text};opacity:0.75;margin:0 0 4px;font-weight:500;">{label}</p>'
        f'<p style="font-size:26px;font-weight:600;margin:0;color:{text};">{value}</p>'
        f'</div>'
    )


_zoom_img_counter = [0]


def render_zoomable_image(image_bytes_or_b64, width=220, caption="", media_type="image/png"):
    """Hiện ảnh thu nhỏ, bấm vào để phóng to toàn màn hình (lightbox) — dùng thuần CSS (checkbox
    ẩn + label), không cần JavaScript, nên chạy ổn định trong mọi ngữ cảnh (kể cả trong form).
    image_bytes_or_b64: chấp nhận cả bytes thô lẫn chuỗi base64 sẵn có. / Shows a thumbnail that
    expands to a fullscreen lightbox on click — pure CSS (hidden checkbox + label trick), no
    JavaScript needed, so it works reliably everywhere (including inside forms)."""
    if isinstance(image_bytes_or_b64, (bytes, bytearray)):
        b64_data = base64.b64encode(image_bytes_or_b64).decode("utf-8")
    else:
        b64_data = image_bytes_or_b64
    _zoom_img_counter[0] += 1
    uid = f"fc_zoom_{_zoom_img_counter[0]}"
    caption_html = (
        f'<div style="font-size:12px;color:#73726c;margin-top:4px;">{caption} '
        f'<span style="opacity:0.7;">(click to zoom)</span></div>'
        if caption else
        '<div style="font-size:12px;color:#73726c;margin-top:4px;opacity:0.7;">(click to zoom)</div>'
    )
    st.markdown(
        f"""
        <input type="checkbox" id="{uid}" style="display:none;">
        <label for="{uid}" style="cursor:zoom-in;display:inline-block;">
            <img src="data:{media_type};base64,{b64_data}" style="width:{width}px;border-radius:8px;display:block;box-shadow:0 1px 3px rgba(20,20,30,0.15);">
        </label>
        {caption_html}
        <label for="{uid}" class="fc-zoom-overlay-{uid}" style="display:none;position:fixed;inset:0;
            background:rgba(20,20,25,0.88);z-index:9999;align-items:center;justify-content:center;cursor:zoom-out;">
            <img src="data:{media_type};base64,{b64_data}" style="max-width:92vw;max-height:92vh;border-radius:10px;box-shadow:0 8px 30px rgba(0,0,0,0.4);">
        </label>
        <style>#{uid}:checked ~ .fc-zoom-overlay-{uid} {{ display: flex !important; }}</style>
        """,
        unsafe_allow_html=True,
    )


def render_paste_zone(target_label_substring, height=100):
    """Vùng dán ảnh Ctrl+V — tự tìm ĐÚNG ô upload cần điền dựa theo 1 đoạn nhãn duy nhất của ô đó
    (target_label_substring), thay vì lấy ô upload ĐẦU TIÊN tìm thấy trên trang. Cần thiết vì app
    này có NHIỀU ô tải ảnh khác nhau cùng lúc (ảnh minh họa lỗi, ảnh chụp email, ảnh mẫu Defect...)
    — nếu chỉ lấy ô đầu tiên sẽ dễ dán nhầm sang ô khác. / A Ctrl+V paste zone that finds the
    CORRECT upload field by matching a unique substring of its label, instead of grabbing the
    first upload field found on the page — necessary since this app has multiple image uploaders
    active at once."""
    label_js = json.dumps(target_label_substring)
    components.html(
        f"""
        <p style="margin:0 0 6px;font-size:13px;color:#73726c;font-family:-apple-system,sans-serif;">
            📋 Click the box below, then press <b>Ctrl+V</b> to paste a screenshot (e.g. Win+Shift+S)
        </p>
        <div id="fc-paste-zone" contenteditable="true" spellcheck="false" style="
            border: 1.5px dashed #9c9a92; border-radius: 8px; padding: 10px 14px;
            text-align: center; cursor: text; color: #9c9a92;
            font-family: -apple-system, sans-serif; font-size: 13px; outline: none;
            transition: border-color 0.15s ease, color 0.15s ease;
        ">(bấm vào đây / click here)</div>
        <div id="fc-paste-status" style="margin-top:4px;font-size:12px;font-family:-apple-system,sans-serif;"></div>
        <script>
        const targetLabel = {label_js};
        const zone = document.getElementById('fc-paste-zone');
        const status = document.getElementById('fc-paste-status');
        const placeholder = '(bấm vào đây / click here)';

        function findTargetInput() {{
            const parentDoc = window.parent.document;
            const candidates = parentDoc.querySelectorAll('[data-testid="stFileUploader"]');
            for (const el of candidates) {{
                if (el.textContent.includes(targetLabel)) {{
                    return el.querySelector('input[type="file"]');
                }}
            }}
            return null;
        }}

        zone.addEventListener('focus', () => {{
            zone.style.borderColor = '#185fa5';
            zone.style.color = '#185fa5';
            if (zone.innerText.trim() === placeholder) {{ zone.innerText = ''; }}
        }});
        zone.addEventListener('blur', () => {{
            zone.style.borderColor = '#9c9a92';
            zone.style.color = '#9c9a92';
            if (!zone.innerText.trim()) {{ zone.innerText = placeholder; }}
        }});
        zone.addEventListener('paste', async (e) => {{
            e.preventDefault();
            zone.innerText = '';
            const items = e.clipboardData ? e.clipboardData.items : [];
            let imageFile = null;
            for (const item of items) {{
                if (item.type && item.type.startsWith('image/')) {{
                    imageFile = item.getAsFile();
                    break;
                }}
            }}
            if (!imageFile) {{
                status.textContent = 'Không tìm thấy ảnh trong clipboard. / No image found in clipboard.';
                status.style.color = '#a32d2d';
                return;
            }}
            try {{
                const input = findTargetInput();
                if (!input) {{
                    status.textContent = 'Không tìm thấy đúng ô upload — thử cuộn tới đó trước. / Could not find the matching upload field.';
                    status.style.color = '#a32d2d';
                    return;
                }}
                const dt = new DataTransfer();
                dt.items.add(imageFile);
                input.files = dt.files;
                input.dispatchEvent(new Event('change', {{ bubbles: true }}));
                status.textContent = '✅ Đã dán — kiểm tra ô upload bên trên. / Pasted — check the uploader above.';
                status.style.color = '#3b6d11';
            }} catch (err) {{
                status.textContent = 'Lỗi / Error: ' + err.message;
                status.style.color = '#a32d2d';
            }}
        }});
        </script>
        """,
        height=height,
    )


def tab_banner(icon, title, subtitle, color="gray"):
    """Banner lớn màu ở đầu mỗi tab — giúp phân vùng rõ ràng đang ở khu vực nào ngay từ cái nhìn
    đầu tiên, thay vì chỉ dựa vào tên tab nhỏ ở trên. / Large colored banner at the top of each
    tab — makes it obvious at a glance which section you're in, not just the small tab label."""
    bg, text = SECTION_COLOR_STYLE.get(color, SECTION_COLOR_STYLE["gray"])
    st.markdown(
        f'<div style="background:{bg};border-radius:14px;padding:1.1rem 1.4rem;margin-bottom:1.25rem;">'
        f'<div style="display:flex;align-items:center;gap:12px;">'
        f'<span style="font-size:22px;">{icon}</span>'
        f'<span style="font-size:1.2rem;font-weight:600;color:{text};">{title}</span>'
        f'</div>'
        f'<p style="margin:6px 0 0 34px;font-size:13px;color:{text};opacity:0.85;">{subtitle}</p>'
        f'</div>',
        unsafe_allow_html=True,
    )


_zone_card_counters = {}


def zone_card(color="gray"):
    """Context manager dùng chung — bọc nội dung vào 1 khối nền màu rõ ràng + viền trái đậm, để
    phân vùng lớn thực sự nổi bật (không chỉ tiêu đề có màu). Dùng ở bất kỳ tab nào:
        with zone_card("blue"):
            section_header("📦", "...", "blue")
            <các widget bên trong>
    Mỗi lần gọi tự sinh 1 key riêng (Streamlit yêu cầu key duy nhất) nhưng CSS áp dụng theo TIỀN TỐ
    màu nên không cần khai báo CSS riêng cho từng chỗ gọi. / Shared context manager — wraps content
    in a clearly colored zone (background + left border), reusable across every tab without
    per-call CSS. Each call gets a unique key (Streamlit requirement) but CSS matches by color
    prefix, so no extra CSS is needed per call site."""
    n = _zone_card_counters.get(color, 0)
    _zone_card_counters[color] = n + 1
    return st.container(key=f"zone_{color}_{n}")


try:
    conn = ensure_connection()
except Exception as e:
    st.error(
        "⚠️ Could not connect to the database. This may be a network issue, or the Supabase project "
        "may be paused (Free tier auto-pauses after ~1 week of inactivity — check at "
        "supabase.com/dashboard). Please check and click retry below."
    )
    with st.expander("Technical details"):
        st.code(str(e))
    if st.button("Retry"):
        st.rerun()
    st.stop()


# ============================================================
# TRANG CHI TIẾT 1 BÁO CÁO — mở qua link ?submission_id=xxx trong Excel tổng hợp, hoặc bấm vào
# card ở Dashboard (lưu vào session_state). Kiểm tra NGAY trước khi vẽ sidebar. Nếu id có tiền tố
# "legacy:" nghĩa là complaint tạo tay qua "Complaint mới" — dùng trang chi tiết đơn giản hơn
# (không có Customer/Client Code/Replacement Cost vì schema cũ không có các cột này).
# ============================================================
_qp_submission_id = st.query_params.get("submission_id")
_ss_submission_id = st.session_state.get("selected_submission_id")
_active_id = _qp_submission_id or _ss_submission_id
if _active_id:
    if str(_active_id).startswith("legacy:"):
        render_legacy_complaint_detail_page(conn, str(_active_id).split("legacy:", 1)[1])
    else:
        render_submission_detail_page(conn, _active_id)
    st.stop()


# ============================================================
# SIDEBAR — điều hướng chính: Dashboard + Hỏi AI, các mục cũ gom trong "More".
# ============================================================
st.markdown(
    """<style>
    section[data-testid="stSidebar"] { background-color: #1a1a3e; }
    section[data-testid="stSidebar"] * { color: #ffffff !important; }
    section[data-testid="stSidebar"] .stButton button {
        background: transparent; border: none; text-align: left; padding: 8px 10px;
        border-radius: 8px; width: 100%; font-size: 14px;
    }
    section[data-testid="stSidebar"] .stButton button:hover { background: rgba(255,255,255,0.10); }
    </style>""",
    unsafe_allow_html=True,
)

if "current_page" not in st.session_state:
    st.session_state.current_page = "dashboard"

with st.sidebar:
    st.markdown("### 🏭 Nilorn Internal AI")
    st.markdown("")
    if st.button("📊  Dashboard", key="nav_dashboard", use_container_width=True):
        st.session_state.current_page = "dashboard"
        st.rerun()
    if st.button("💬  Ask AI", key="nav_ask", use_container_width=True):
        st.session_state.current_page = "ask_ai"
        st.rerun()
    with st.expander("More"):
        if st.button("✅ Review Taxonomy", key="nav_taxonomy", use_container_width=True):
            st.session_state.current_page = "taxonomy"
            st.rerun()
        if st.button("📝 New Complaint (Manual)", key="nav_new_complaint", use_container_width=True):
            st.session_state.current_page = "new_complaint"
            st.rerun()
        if st.button("🗄️ Data Lookup", key="nav_data_lookup", use_container_width=True):
            st.session_state.current_page = "data_lookup"
            st.rerun()

page = st.session_state.current_page

# ------------------------------------------------------------
# TAB 1 — DUYỆT TAXONOMY (cần chọn reviewer)
# ------------------------------------------------------------
if page == "taxonomy":
    conn = ensure_connection()
    tab_banner(
        "🔍", "Review Taxonomy",
        "Review and approve new Defect / Root Cause / CAPA suggestions detected by AI.",
        "amber",
    )

    reviewers = fetch_reviewers(conn)
    if not reviewers:
        st.error("No reviewer found in the 'reviewer' table. Add at least one before using this.")
        st.stop()

    reviewer_names = {name: rid for rid, name in reviewers}
    selected_name = st.selectbox(
        "Who are you?",
        list(reviewer_names.keys()),
        help="Select your name for the review record.",
    )
    reviewer_id = reviewer_names[selected_name]

    if st.button("Refresh"):
        st.rerun()

    with st.expander("Attach reference images for an existing Defect code"):
        st.caption(
            "Use this for existing Defect codes that never got an image when they were first approved. Each code holds up to 3 images."
        )
        defect_codes_for_image = fetch_existing_codes(conn, "Defect")
        defect_image_labels = ["select code --"] + [f"{c} — {n}" for c, n in defect_codes_for_image]
        picked_defect_label = st.selectbox("Defect code", defect_image_labels, key="picked_defect_for_image")
        if picked_defect_label != "select code --":
            picked_defect_code = picked_defect_label.split(" — ")[0]
            existing_ref_imgs = fetch_defect_reference_images(conn, picked_defect_code)
            if existing_ref_imgs:
                st.caption(f"Currently have {len(existing_ref_imgs)}/{MAX_DEFECT_REFERENCE_IMAGES}{len(existing_ref_imgs)}/{MAX_DEFECT_REFERENCE_IMAGES} images")
                img_cols = st.columns(len(existing_ref_imgs))
                for i, (col, img_b64) in enumerate(zip(img_cols, existing_ref_imgs)):
                    with col:
                        render_zoomable_image(img_b64, width=150, caption=f"Image {i + 1} / Image {i + 1}")
                        if st.button("Delete", key=f"del_ref_img_{picked_defect_code}_{i}"):
                            remove_defect_reference_image(conn, picked_defect_code, i)
                            st.rerun()
            slots_left = MAX_DEFECT_REFERENCE_IMAGES - len(existing_ref_imgs)
            if slots_left > 0:
                new_ref_img_files = st.file_uploader(
                    f"Add reference images (remaining {slots_left}Add reference images ({slots_left} slot(s) left)",
                    type=["png", "jpg", "jpeg"], accept_multiple_files=True, key=f"new_ref_img_{picked_defect_code}",
                )
                if new_ref_img_files and st.button("Save images", key=f"btn_save_ref_img_{picked_defect_code}"):
                    new_b64_list = [base64.b64encode(f.getvalue()).decode("utf-8") for f in new_ref_img_files]
                    _, dropped = add_defect_reference_images(conn, picked_defect_code, new_b64_list)
                    if dropped:
                        st.warning(
                            f"Saved, but exceeded the limit of {MAX_DEFECT_REFERENCE_IMAGES} images, so {dropped}Saved, but {dropped} oldest image(s) were dropped to stay within "
                            f"the {MAX_DEFECT_REFERENCE_IMAGES}-image limit."
                        )
                    else:
                        st.success(f"Reference images saved for {picked_defect_code}.")
                    st.rerun()
            else:
                st.caption(f"Already have {MAX_DEFECT_REFERENCE_IMAGES}Already at the {MAX_DEFECT_REFERENCE_IMAGES}-image limit — delete one to add a new one.")

    @st.fragment
    def render_pending_queue():
        conn_f = ensure_connection()
        pending = fetch_pending(conn_f)
        # check_and_notify_new_suggestions(conn_f)  # đã tắt — supplier_portal.py giờ gửi email
        # báo reviewer NGAY lúc NCC nộp, không cần cơ chế "kiểm tra khi mở tab" này nữa.
        card_color = "amber" if pending else "teal"
        st.markdown(
            stat_card_html("Pending", len(pending), card_color),
            unsafe_allow_html=True,
        )
        st.write("")

        if not pending:
            st.success(
                "No suggestions pending — queue is empty."
            )

        for row in pending:
            (sid, complaint_id, kind, suggested_name, reasoning,
             closest_code, created_at, date_opened, so_po, responsible_party) = row

            with st.container(border=True):
                st.markdown(
                    f'{kind_badge_html(kind)}&nbsp;&nbsp;<span style="font-size:1.15rem;font-weight:600;">{suggested_name}</span>',
                    unsafe_allow_html=True,
                )
                st.write(f"AI reasoning:** {reasoning}")
                meta_bits = [f"📄 {so_po or '—'}", f"📅 {date_opened or '?'}"]
                if closest_code:
                    meta_bits.append(f"🔍 {closest_code}")
                meta_bits.append(f"🕓 {created_at}")
                if responsible_party:
                    meta_bits.append(f"👤 {responsible_party}")
                st.caption(" | ".join(meta_bits))

                if kind == "Defect" and closest_code:
                    closest_ref_imgs = fetch_defect_reference_images(conn_f, closest_code)
                    if closest_ref_imgs:
                        cmp_cols = st.columns(len(closest_ref_imgs))
                        for i, (col, img_b64) in enumerate(zip(cmp_cols, closest_ref_imgs)):
                            with col:
                                render_zoomable_image(
                                    img_b64, width=180,
                                    caption=f"{closest_code} #{i + 1}" if len(closest_ref_imgs) > 1 else f"Reference image {closest_code}",
                                )

                existing_codes = fetch_existing_codes(conn_f, kind)
                code_options = ["select --"] + [f"{c} — {n}" for c, n in existing_codes]

                action = st.radio(
                    "Decision",
                    ["Match existing code",
                     "This is new — add new code",
                     "Reject"],
                    key=f"action_{sid}", horizontal=True
                )

                if action == "Match existing code":
                    chosen = st.selectbox("Select the best matching code",
                                           code_options, key=f"match_{sid}")
                    if st.button("Confirm match", key=f"btn_match_{sid}"):
                        if chosen == "select --":
                            st.warning("You haven't selected a code yet.")
                        else:
                            code_only = chosen.split(" — ")[0]
                            approve_match_existing(conn_f, sid, code_only, reviewer_id, complaint_id, kind, responsible_party)
                            st.session_state["needs_extra_refresh"] = True
                            st.rerun()

                elif action == "This is new — add new code":
                    prefix_map = {"Defect": "DEF", "Root Cause": "RC", "CAPA": "CAPA"}
                    suggested_code = suggest_next_code([c for c, _ in existing_codes], prefix_map[kind])
                    new_code = st.text_input(
                        "New code (next code suggested, editable)",
                        value=suggested_code, key=f"newcode_{sid}",
                    )
                    name_final = st.text_input("Name", value=suggested_name, key=f"name_{sid}")
                    desc_final = st.text_area("Description", value=reasoning, key=f"desc_{sid}")

                    extra = {}
                    new_defect_ref_images = None
                    if kind == "Defect":
                        extra["product_group"] = st.selectbox("Product Group", PRODUCT_GROUPS, key=f"pg_{sid}")
                        extra["severity"] = st.selectbox("Default Severity", SEVERITIES, key=f"sev_{sid}")
                        new_defect_ref_images = st.file_uploader(
                            "Reference images for this new Defect code, up to 3 (optional)",
                            type=["png", "jpg", "jpeg"], accept_multiple_files=True, key=f"refimg_{sid}",
                            help="Shown whenever this code is matched/suggested in Ask AI and Review Taxonomy.",
                        )
                    elif kind == "Root Cause":
                        extra["category"] = st.selectbox("Category (6M+)", RC_CATEGORIES, key=f"cat_{sid}")
                    elif kind == "CAPA":
                        extra["capa_type"] = st.selectbox("CAPA Type", CAPA_TYPES, key=f"capatype_{sid}")

                    if st.button("Add new code & Approve", key=f"btn_new_{sid}"):
                        if not new_code.strip():
                            st.warning("You haven't entered a new code.")
                        else:
                            approve_new_code(conn_f, sid, new_code.strip(), name_final, desc_final,
                                              reviewer_id, complaint_id, kind, extra, responsible_party)
                            if kind == "Defect" and new_defect_ref_images:
                                ref_b64_list = [base64.b64encode(f.getvalue()).decode("utf-8") for f in new_defect_ref_images[:MAX_DEFECT_REFERENCE_IMAGES]]
                                add_defect_reference_images(conn_f, new_code.strip(), ref_b64_list)
                            st.session_state["needs_extra_refresh"] = True
                            st.rerun()

                else:  # Từ chối / Reject
                    if st.button("Reject", key=f"btn_reject_{sid}"):
                        reject(conn_f, sid, reviewer_id)
                        st.rerun()

    render_pending_queue()

# ------------------------------------------------------------
# TAB 2 — TRUY XUẤT DỮ LIỆU (ai xem cũng được, không cần chọn reviewer)
# ------------------------------------------------------------
if page == "data_lookup":
    conn = ensure_connection()
    tab_banner(
        "📊", "Data Lookup",
        "Look up data, export Word reports, and draft emails for existing complaints.",
        "blue",
    )
    if st.session_state.pop("needs_extra_refresh", False):
        st.rerun()
    if st.button("Refresh", key="btn_refresh_data_tab"):
        st.rerun()

    orphaned = find_orphaned_submissions(conn)
    if orphaned:
        with zone_card("coral"):
            section_header(
                "🛠️", f"Found {len(orphaned)} supplier report(s) missing a matching complaint — need repair",
                "coral",
            )
            st.caption(
                "Due to a previous technical issue (now fixed), some supplier reports submitted via the shared link failed to create a matching complaint. Click the button below to automatically restore them (including Defect/Root Cause/CAPA suggestions for the reviewer)."
            )
            if st.button(f"🛠️ Repair {len(orphaned)}Repair {len(orphaned)} missing complaints"):
                ai_client_repair = get_ai_client()
                progress = st.progress(0.0)
                for i, row in enumerate(orphaned):
                    repair_orphaned_submission(conn, ai_client_repair, *row)
                    progress.progress((i + 1) / len(orphaned))
                st.success(f"Repaired {len(orphaned)} complaint. / Repaired {len(orphaned)} complaints.")
                st.rerun()

    with conn.cursor() as cur:
        cur.execute("select status, count(*) from complaint group by status;")
        status_counts = dict(cur.fetchall())
    total_complaints = sum(status_counts.values())
    not_closed_count = total_complaints - status_counts.get("Closed", 0)
    closed_count = status_counts.get("Closed", 0)
    stat_col1, stat_col2, stat_col3 = st.columns(3)
    with stat_col1:
        st.markdown(stat_card_html("Total", total_complaints, "gray"), unsafe_allow_html=True)
    with stat_col2:
        st.markdown(stat_card_html("Not closed", not_closed_count, "amber"), unsafe_allow_html=True)
    with stat_col3:
        st.markdown(stat_card_html("Closed", closed_count, "teal"), unsafe_allow_html=True)

    with zone_card("blue"):
        section_header("📊", "Target questions", "blue")
        chosen_q = st.selectbox("Choose a question to view", list(TARGET_QUERIES.keys()))
        df = pd.read_sql(TARGET_QUERIES[chosen_q], conn)
        if df.empty:
            with conn.cursor() as cur:
                cur.execute("select count(*) from complaint where defect_code is null or root_cause_code is null;")
                n_pending_codes = cur.fetchone()[0]
            hint = ""
            if "6 months" in chosen_q or "6 months" in chosen_q.lower():
                hint = " This question only counts complaints from the last 6 months — check the Record Date."
            elif chosen_q.startswith(("Q2", "Q3", "Q4")) and n_pending_codes:
                hint = f" Currently there are {n_pending_codes} complaint(s) missing an approved Defect/Root Cause — go to Review Taxonomy to process first."
            st.info(f"No data matches the filter for this question.{hint}")
        else:
            df.index = range(1, len(df) + 1)
            st.dataframe(df, use_container_width=True)

            excel_buf = io.BytesIO()
            with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Data")
                ws = writer.sheets["Data"]
                for cell in ws[1]:
                    cell.font = cell.font.copy(bold=True)
                for col_cells in ws.columns:
                    max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
                    ws.column_dimensions[col_cells[0].column_letter].width = max_len + 4
            st.download_button(
                "Export Excel", data=excel_buf.getvalue(),
                file_name="truy_xuat_du_lieu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            if df.shape[1] == 2 and df.shape[0] > 0:
                st.bar_chart(df.set_index(df.columns[0]))
                st.download_button(
                    "Export chart (PNG)", data=export_chart_png(df),
                    file_name="bieu_do.png", mime="image/png",
                )

    with zone_card("coral"):
        section_header("📋", "Export full report for a complaint", "coral")
        with conn.cursor() as cur:
            cur.execute("""
                select complaint_id, so_po, date_opened, notes
                from complaint order by date_opened desc, complaint_id desc limit 100;
            """)
            recent_complaints = cur.fetchall()

        if not recent_complaints:
            st.info("No complaints in the system yet.")
        else:
            complaint_id_to_row = {row[0]: row for row in recent_complaints}
            valid_ids = [row[0] for row in recent_complaints]

            sticky_id = st.session_state.get("sticky_picked_complaint_id")
            if sticky_id and sticky_id not in complaint_id_to_row:
                with conn.cursor() as cur:
                    cur.execute(
                        "select complaint_id, so_po, date_opened, notes from complaint where complaint_id = %s;",
                        (sticky_id,),
                    )
                    extra_row = cur.fetchone()
                if extra_row:
                    complaint_id_to_row[extra_row[0]] = extra_row
                    valid_ids = [extra_row[0]] + valid_ids

            if "sticky_picked_complaint_id" not in st.session_state or st.session_state.sticky_picked_complaint_id not in valid_ids:
                st.session_state.sticky_picked_complaint_id = valid_ids[0]
            default_index = valid_ids.index(st.session_state.sticky_picked_complaint_id)

            picked_id = st.selectbox(
                "Select complaint to export",
                options=valid_ids,
                format_func=lambda cid: (
                    f"{complaint_id_to_row[cid][1] or '(no SO/PO)'} — "
                    f"{complaint_id_to_row[cid][2].strftime('%d/%m/%Y') if complaint_id_to_row[cid][2] else '?'} — "
                    f"{truncate_at_word(complaint_id_to_row[cid][3], 120)}"
                ),
                index=default_index,
                key="picked_complaint_id_widget",
            )
            st.session_state.sticky_picked_complaint_id = picked_id
            full_report = fetch_complaint_full_report(conn, picked_id)

            sig_row = fetch_latest_supplier_signature(conn, picked_id)
            supplier_sig_name, supplier_sig_image_b64 = sig_row if sig_row else (None, None)

            try:
                st.download_button(
                    "Word (latest version)",
                    data=export_word_report(
                        full_report,
                        supplier_signature_name=supplier_sig_name,
                        supplier_signature_image_b64=supplier_sig_image_b64,
                    ),
                    file_name=f"BaoCao_{picked_id}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="full_word",
                )
            except Exception as e:
                st.error(f"Could not create Word file: {e}")

            with st.expander(
                "Attach or replace this complaint's defect photo"
            ):
                if full_report.get("Defect Photo"):
                    render_zoomable_image(full_report["Defect Photo"], width=220, caption="Current photo")
                new_defect_photo = st.file_uploader(
                    "Defect photo", type=["png", "jpg", "jpeg"],
                    key=f"attach_defect_photo_{picked_id}",
                )
                if new_defect_photo is not None and st.button("Save photo", key=f"btn_save_defect_photo_{picked_id}"):
                    photo_b64 = base64.b64encode(new_defect_photo.getvalue()).decode("utf-8")
                    with conn.cursor() as cur:
                        cur.execute("update complaint set defect_photo = %s where complaint_id = %s;", (photo_b64, picked_id))
                    conn.commit()
                    st.success("Photo saved — re-download the Word file to get the latest version.")
                    st.rerun()

            has_root_cause = bool(full_report.get("Root cause"))
            has_capa = bool(full_report.get("CAPA"))

            st.markdown("---")
            with st.expander("Add a CAPA action to this complaint"):
                capa_add_result = st.session_state.pop("capa_add_result", None)
                if capa_add_result and capa_add_result.get("complaint_id") == picked_id:
                    st.success(capa_add_result["message"])
                capa_codes_add = fetch_existing_codes(conn, "CAPA")
                capa_labels_add = ["select existing code --"] + [f"{c} — {n}" for c, n in capa_codes_add]
                with st.form(f"add_capa_form_{picked_id}"):
                    capa_choice_add = st.selectbox("Existing CAPA", capa_labels_add)
                    new_capa_text_add = st.text_input(
                        "...or describe a NEW CAPA (if no existing code fits — will be sent to the approval queue)"
                    )
                    responsible_add = st.text_input("Responsible person")
                    date_proposed_add = st.date_input("CAPA proposed date")
                    submitted_capa_add = st.form_submit_button("Add CAPA")
                    if submitted_capa_add:
                        if capa_choice_add != "select existing code --":
                            code_only_add = capa_choice_add.split(" — ")[0]
                            add_capa_action_to_complaint(conn, picked_id, code_only_add, date_proposed_add, responsible_add or None)
                            st.session_state["capa_add_result"] = {
                                "complaint_id": picked_id,
                                "message": f"✅ Added CAPA **{code_only_add}Added CAPA **{code_only_add}** — now in effect, see the list below.",
                            }
                            st.rerun()
                        elif new_capa_text_add.strip():
                            ai_client = get_ai_client()
                            capa_taxonomy_list = load_taxonomy_list(conn, "CAPA")
                            check_result = classify_description(ai_client, new_capa_text_add.strip(), capa_taxonomy_list, "CAPA")
                            conn = ensure_connection()
                            closest = check_result.get("closest_existing_code") or check_result.get("matched_code")
                            note = (
                                f"CAPA added later (not at complaint entry time) — manually described by staff, no matching code yet. "
                                f"Expected responsible person: {responsible_add or '(unknown)'}. "
                                f"AI cross-check: {check_result.get('reasoning', '(no notes)')}"
                            )
                            with conn.cursor() as cur:
                                cur.execute(
                                    """insert into taxonomy_suggestion
                                           (complaint_id, suggestion_type, ai_suggested_name, ai_reasoning, closest_existing_code, responsible_party)
                                       values (%s, %s, %s, %s, %s, %s);""",
                                    (picked_id, "CAPA", new_capa_text_add.strip(), note, closest, responsible_add or None),
                                )
                            conn.commit()
                            st.session_state["capa_add_result"] = {
                                "complaint_id": picked_id,
                                "message": "Sent the new CAPA to the **approval queue** — go to 'Review Taxonomy' to approve it before it applies.",
                            }
                            st.rerun()
                        else:
                            st.warning("You haven't selected an existing code or entered a new CAPA description.")

            st.markdown("---")
            section_header("🔄", "Complaint status", "gray")
            capa_rows_status = fetch_capa_actions_for_complaint(conn, picked_id)

            if not capa_rows_status:
                st.caption("No CAPA action yet — add one before moving to Resolved/Closed.")
            else:
                for capa_action_id, capa_code, capa_action_name, date_proposed, date_implemented, verif, resp in capa_rows_status:
                    with st.container(border=True):
                        st.write(f"**{capa_code} — {capa_action_name}responsible: {resp or 'unknown)'})")
                        st.caption(
                            f"Proposed on: {date_proposed.strftime('%d/%m/%Y') if date_proposed else '?'}Effectiveness verification: {verif or 'Pending'}"
                        )
                        if date_implemented:
                            st.success(f"Implemented on {date_implemented.strftime('%d/%m/%Y')}")
                        else:
                            impl_date = st.date_input(
                                "Mark implemented on date", key=f"impl_date_{capa_action_id}"
                            )
                            if st.button("Mark implemented", key=f"btn_impl_{capa_action_id}"):
                                mark_capa_implemented(conn, capa_action_id, impl_date)
                                compute_and_update_complaint_status(conn, picked_id)
                                st.rerun()

                        confirm_del = st.checkbox(
                            "Confirm delete this CAPA (cannot be undone)",
                            key=f"confirmdel_{capa_action_id}",
                        )
                        if st.button("Delete CAPA action", key=f"btn_delcapa_{capa_action_id}", disabled=not confirm_del):
                            delete_capa_action(conn, capa_action_id)
                            compute_and_update_complaint_status(conn, picked_id)
                            st.success("CAPA action deleted.")
                            st.rerun()

            missing_tags_now = compute_and_update_complaint_status(conn, picked_id)
            simple_status = "Closed" if not missing_tags_now else "Chưa hoàn thành"
            st.markdown(
                f"Status: {status_badge_html(simple_status)}",
                unsafe_allow_html=True,
            )

        st.markdown("---")
    with st.expander("CAPA pending effectiveness verification"):
        pending_verifications = fetch_capa_pending_verification(conn)
        if not pending_verifications:
            st.info("No CAPA pending verification.")
        else:
            for (capa_action_id, cid, so_po_v, date_opened_v, capa_code, capa_action_name,
                 eff_window, date_implemented_v, resp_v) in pending_verifications:
                with st.container(border=True):
                    st.write(f"**{capa_code} — {capa_action_name}**")
                    st.caption(
                        f"Complaint: {so_po_v or '(no SO/PO)'}opened {date_opened_v.strftime('%d/%m/%Y') if date_opened_v else '?'}Implemented: {date_implemented_v.strftime('%d/%m/%Y')}Responsible: {resp_v or 'unknown)'}Suggested monitoring window: {eff_window or 'none)'}"
                    )
                    col_v1, col_v2, col_v3 = st.columns([2, 2, 1])
                    with col_v1:
                        verif_choice = st.selectbox(
                            "Verification result", VERIFICATION_RESULTS, key=f"verifres_{capa_action_id}"
                        )
                    with col_v2:
                        verif_date = st.date_input("Verification date", key=f"verifdate_{capa_action_id}")
                    with col_v3:
                        st.write("")
                        st.write("")
                        if st.button("Save result", key=f"btn_verif_{capa_action_id}"):
                            submit_capa_verification(conn, capa_action_id, verif_choice, verif_date)
                            st.success("Verification result saved.")
                            st.rerun()


# ------------------------------------------------------------
# TAB 3 — HỎI AI TỰ DO (câu hỏi mở, không giới hạn 6 câu mục tiêu)
# ------------------------------------------------------------
if page == "ask_ai":
    conn = ensure_connection()
    tab_banner(
        "💬", "Ask AI",
        "Ask anything about the quality data — not limited to the 4 target questions.",
        "teal",
    )
    question = st.text_input(
        "Your question",
        placeholder="Example: What is the root cause of damp polybag defects?",
        help="Ask anything — not limited to the 4 target questions.",
    )

    if st.button("Ask AI") and question.strip():
        ai_client = get_ai_client()
        with st.spinner("AI is looking this up..."):
            try:
                conn = ensure_connection()
                taxonomy_context = fetch_taxonomy_context(conn)
                master_data_context = fetch_master_data_context(conn)
                gen = generate_sql(ai_client, question, taxonomy_context, master_data_context)
                conn = ensure_connection()
                sql = gen["sql"]

                if not is_safe_select(sql):
                    st.session_state["ask_result"] = {"error": "unsafe_sql", "sql": sql}
                else:
                    conn = ensure_connection()
                    try:
                        df_ans = pd.read_sql(sql, conn)
                    except Exception as sql_err:
                        st.session_state["ask_result"] = {
                            "error": "sql_execution_error",
                            "sql": sql,
                            "explanation": gen.get("explanation", ""),
                            "debug": str(sql_err),
                            "question": question,
                            "raw": gen.get("raw", ""),
                        }
                    else:
                        answer_text = None
                        if not df_ans.empty:
                            answer_text = synthesize_answer(ai_client, question, df_ans)
                            conn = ensure_connection()
                        st.session_state["ask_result"] = {
                            "question": question,
                            "sql": sql,
                            "explanation": gen.get("explanation", ""),
                            "question_type": gen.get("question_type", "retrieval"),
                            "matched_defect": gen.get("matched_defect_code"),
                            "matched_rc": gen.get("matched_root_cause_code"),
                            "suggest_new_complaint": gen.get("suggest_new_complaint", False),
                            "df": df_ans,
                            "answer": answer_text,
                        }
            except ValueError as e:
                st.session_state["ask_result"] = {"error": "not_a_question", "debug": str(e), "question": question}
            except Exception as e:
                st.session_state["ask_result"] = {"error": str(e), "question": question}

    result = st.session_state.get("ask_result")

    if result:
        if result.get("error") == "unsafe_sql":
            st.error(
                "AI generated an unsafe statement (contains a data-modifying keyword) — blocked, not executed."
            )
            if result.get("sql"):
                with st.expander("View the blocked SQL"):
                    st.code(result["sql"], language="sql")
        elif result.get("error") == "not_a_question":
            st.info(
                "This looks more like a **description of a new issue** than a lookup question, OR the AI replied in an unexpected format. The 'Ask AI' tab is for **read-only lookups** only."
            )
            render_create_complaint_suggestion(conn, result.get("question", ""), "btn_goto_new_complaint")
            st.caption(
                "If this is really a lookup question, try rephrasing it."
            )
            if result.get("debug"):
                with st.expander("Technical details (for bug reports)"):
                    st.code(result["debug"])
        elif result.get("error") == "sql_execution_error":
            st.error(
                "AI generated SQL but it failed to run — possibly a complex question or a wrong column/table reference. Try rephrasing, or share the debug details below."
            )
            with st.expander("Technical details (failed SQL + DB error)"):
                st.code(result.get("sql", ""), language="sql")
                st.caption(result.get("explanation", ""))
                st.code(result.get("debug", ""))
                if result.get("raw"):
                    st.caption(
                        "Full raw AI output (to check for truncation):"
                    )
                    st.code(result["raw"])
        elif "error" in result:
            st.error(f"An error occurred: {result['error']}")
        else:
            with st.expander("View the SQL AI used"):
                st.code(result["sql"], language="sql")
                st.caption(result["explanation"])

            df_ans = result["df"]
            matched_defect = result["matched_defect"]
            matched_rc = result["matched_rc"]

            if matched_defect:
                ref_imgs = fetch_defect_reference_images(conn, matched_defect)
                if ref_imgs:
                    ask_img_cols = st.columns(len(ref_imgs))
                    for i, (col, img_b64) in enumerate(zip(ask_img_cols, ref_imgs)):
                        with col:
                            render_zoomable_image(
                                img_b64, width=220,
                                caption=f"{matched_defect} #{i + 1}" if len(ref_imgs) > 1 else f"Reference image {matched_defect} / Reference image",
                            )

            if not df_ans.empty:
                df_ans.index = range(1, len(df_ans) + 1)
                st.dataframe(df_ans, use_container_width=True)

                # Tự vẽ biểu đồ nếu dữ liệu có dạng phù hợp (nhiều dòng, có ít nhất 1 cột phân loại
                # + 1 cột số) — giúp trả lời trực quan cho câu hỏi so sánh/tổng hợp/xu hướng mà
                # không cần người dùng phải nói rõ "vẽ biểu đồ" mới có.
                if len(df_ans) > 1:
                    numeric_cols = df_ans.select_dtypes(include="number").columns.tolist()
                    non_numeric_cols = [c for c in df_ans.columns if c not in numeric_cols]
                    if numeric_cols and non_numeric_cols:
                        try:
                            chart_df = df_ans[[non_numeric_cols[0], numeric_cols[0]]].set_index(non_numeric_cols[0])
                            st.bar_chart(chart_df, use_container_width=True)
                        except Exception:
                            pass

                st.markdown(f"Answer\n{result['answer']}")

                if result.get("suggest_new_complaint"):
                    st.markdown("---")
                    render_create_complaint_suggestion(conn, result.get("question", ""), "btn_goto_new_complaint_retrieval")

            elif result.get("question_type") != "classification":
                if result.get("suggest_new_complaint"):
                    st.info(
                        "No matching complaint found yet — this may be the first time it's being recorded."
                    )
                    st.markdown("This reads like **a specific complaint/issue**:")
                    render_create_complaint_suggestion(conn, result.get("question", ""), "btn_goto_new_complaint_zerorows")
                else:
                    st.info(
                        "No matching data found. This could be because: (1) a name (supplier/customer/product/machine...) didn't match exactly, or (2) there's simply no complaint matching these conditions yet — not an error, just insufficient data."
                    )

            elif matched_defect or matched_rc:
                code_shown = matched_defect or matched_rc
                kind_shown = "Defect" if matched_defect else "Root Cause"
                st.markdown(
                    f'{kind_badge_html(kind_shown)}&nbsp;&nbsp;<span style="font-weight:600;">{code_shown}</span>',
                    unsafe_allow_html=True,
                )
                st.warning(
                    "This already **exists in the taxonomy** — it's not unclassified. The system just doesn't have a real complaint recorded with enough detail yet (often a missing date)."
                )
                st.markdown(
                    "If this is a specific complaint, use the button below to create a full New Complaint (AI auto-fills product, customer, quantity...) — or use the quick form below for a simple record with the matched code."
                )
                render_create_complaint_suggestion(conn, result.get("question", ""), "btn_goto_new_complaint_matched")

                with st.form("quickform"):
                    st.write("Quickly log a complaint for this case")
                    new_date = st.date_input("Date occurred (required)")

                    defect_options = fetch_existing_codes(conn, "Defect")
                    rc_options = fetch_existing_codes(conn, "Root Cause")
                    defect_labels = ["unknown --"] + [f"{c} — {n}" for c, n in defect_options]
                    rc_labels = ["unknown --"] + [f"{c} — {n}" for c, n in rc_options]

                    def _default_index(labels, code):
                        if not code:
                            return 0
                        for i, lbl in enumerate(labels):
                            if lbl.startswith(code + " —"):
                                return i
                        return 0

                    defect_choice = st.selectbox(
                        "Defect code", defect_labels,
                        index=_default_index(defect_labels, matched_defect),
                    )
                    rc_choice = st.selectbox(
                        "Root cause code (if known — usually the missing piece)", rc_labels,
                        index=_default_index(rc_labels, matched_rc),
                    )
                    new_notes = st.text_area("Notes / additional description", value=result["question"])
                    submitted = st.form_submit_button("Save new complaint")
                    if submitted:
                        defect_final = None if defect_choice.startswith("--") else defect_choice.split(" — ")[0]
                        rc_final = None if rc_choice.startswith("--") else rc_choice.split(" — ")[0]
                        with conn.cursor() as cur:
                            cur.execute(
                                """insert into complaint (date_opened, defect_code, root_cause_code, notes, status)
                                   values (%s, %s, %s, %s, 'Thiếu Customer')
                                   returning complaint_id;""",
                                (new_date, defect_final, rc_final, new_notes),
                            )
                            new_id_quicksave = cur.fetchone()[0]
                        conn.commit()
                        compute_and_update_complaint_status(conn, new_id_quicksave)
                        st.success("Saved — click 'Ask AI' again to see the updated result.")

            else:
                st.warning(
                    "No matching code found in the taxonomy — **this may be a new type of defect/root cause** not yet classified."
                )

                if result.get("suggest_new_complaint"):
                    st.markdown(
                        "This reads more like **a specific complaint/issue** — better to **create a New Complaint** directly (with AI-prefilled fields) rather than just submitting a single taxonomy suggestion:"
                    )
                    render_create_complaint_suggestion(conn, result.get("question", ""), "btn_goto_new_complaint_classification")
                    st.caption(
                        "Only use the form below if you really just want to propose a new Defect/Root Cause/CAPA concept — without an actual complaint."
                    )

                sugg_result = st.session_state.pop("sugg_result", None)
                if sugg_result:
                    st.success(sugg_result)

                with st.expander(
                    "Only propose a new taxonomy concept (no complaint)"
                ):
                    with st.form("suggform"):
                        sugg_type = st.radio("Suggestion type", ["Defect", "Root Cause", "CAPA"], horizontal=True)
                        sugg_name = st.text_input("Suggested name (short)", value=result["question"])
                        sugg_reason = st.text_area("Description / reason", value=result["question"])
                        submitted2 = st.form_submit_button("Submit to queue")
                        if submitted2:
                            with conn.cursor() as cur:
                                cur.execute(
                                    """insert into taxonomy_suggestion (suggestion_type, ai_suggested_name, ai_reasoning)
                                       values (%s, %s, %s);""",
                                    (sugg_type, sugg_name, sugg_reason),
                                )
                            conn.commit()
                            st.session_state["sugg_result"] = (
                                "Sent to the queue — go to 'Review Taxonomy' to view and process it."
                            )
                            st.rerun()

# ------------------------------------------------------------
# TAB 4 — NHẬP COMPLAINT MỚI (luôn lưu lịch sử, tự phân loại hoặc gửi duyệt)
# ------------------------------------------------------------
if page == "new_complaint":
    conn = ensure_connection()
    tab_banner(
        "📝", "New Complaint",
        "Log a new complaint — the system auto-classifies, matches an existing code, or sends it to the approval queue.",
        "coral",
    )

    with st.expander(
        "Quick-create from an email (paste text or upload a screenshot)"
    ):
        email_text_input = st.text_area(
            "Paste the full email content here",
            height=150, key="email_paste_input",
        )
        email_image_input = st.file_uploader(
            "...or upload email screenshot(s) / real defect photo(s) (multiple allowed)",
            type=["png", "jpg", "jpeg"], key="email_image_input", accept_multiple_files=True,
            help="You can select multiple images at once (e.g. an email screenshot plus real defect photos) — AI reads all of them together for more accurate results. Max 5 images per run.",
        )
        render_paste_zone("upload email screenshot")
        if st.button("Extract from email", key="btn_extract_email"):
            if not email_text_input.strip() and not email_image_input:
                st.warning("You haven't pasted any content or uploaded an image.")
            else:
                with st.spinner("AI is reading the email..."):
                    ai_client = get_ai_client()
                    (product_names, customer_names, supplier_names,
                     staff_labels, rc_list, capa_list) = fetch_intake_lookup_lists(conn)
                    conn = ensure_connection()
                    if email_image_input:
                        images_for_ai = [
                            {
                                "b64": base64.b64encode(f.getvalue()).decode("utf-8"),
                                "media_type": f.type or "image/png",
                            }
                            for f in email_image_input[:5]
                        ]
                        if len(email_image_input) > 5:
                            st.warning(
                                f"You uploaded {len(email_image_input)}You uploaded {len(email_image_input)} images — only the first 5 are sent to AI."
                            )
                        fields = extract_complaint_from_email(
                            ai_client, datetime.now().strftime("%Y-%m-%d"),
                            product_names, customer_names, supplier_names, staff_labels, rc_list, capa_list,
                            images=images_for_ai,
                        )
                    else:
                        fields = extract_complaint_from_email(
                            ai_client, datetime.now().strftime("%Y-%m-%d"),
                            product_names, customer_names, supplier_names, staff_labels, rc_list, capa_list,
                            email_text=email_text_input.strip(),
                        )
                matched_bits = apply_complaint_prefill_fields(fields)
                if matched_bits:
                    st.session_state["email_extract_result"] = (
                        "Email read and prefilled: " + ", ".join(matched_bits) + "Read the email and prefilled: " + ", ".join(matched_bits) +
                        " — review the fields below before saving."
                    )
                else:
                    st.session_state["email_extract_result"] = (
                        "Read the email but couldn't match any field to real data — the description was prefilled, please fill in the rest."
                    )

                supplier_suggestion = fields.get("supplier")
                suggestion_source = "email" if supplier_suggestion else None
                if not supplier_suggestion and fields.get("product"):
                    supplier_suggestion = suggest_supplier_for_product(conn, fields["product"])
                    suggestion_source = "history" if supplier_suggestion else None
                st.session_state["email_extract_context"] = {
                    "desc": fields.get("desc"),
                    "product": fields.get("product"),
                    "so_po": fields.get("so_po"),
                    "quantity": fields.get("quantity"),
                    "supplier_suggestion": supplier_suggestion,
                    "suggestion_source": suggestion_source,
                }
                st.rerun()

    email_extract_result = st.session_state.pop("email_extract_result", None)
    if email_extract_result:
        st.success(email_extract_result)

    email_context = st.session_state.get("email_extract_context")
    if email_context and email_context.get("supplier_suggestion"):
        suggestion = email_context["supplier_suggestion"]
        source = email_context.get("suggestion_source")
        hint = "mentioned directly in the email" if source == "email" else "based on past complaint history"
        st.info(
            f"💡 Suggested related supplier ({hint}): **{suggestion}**. Save the complaint below first, then go to "
            f"'📊 Data Lookup' → export report to confirm and send them a request email."
        )

    try:
        products = fetch_lookup(conn, "product", "product_id", "name")
        suppliers = fetch_lookup(conn, "supplier", "supplier_id", "name")
        machines = fetch_lookup(conn, "machine", "machine_id", "name")
        customers = fetch_lookup(conn, "customer", "customer_id", "name")
        cs_staff_list = fetch_cs_staff(conn)
    except Exception:
        conn = ensure_connection()
        products = fetch_lookup(conn, "product", "product_id", "name")
        suppliers = fetch_lookup(conn, "supplier", "supplier_id", "name")
        machines = fetch_lookup(conn, "machine", "machine_id", "name")
        customers = fetch_lookup(conn, "customer", "customer_id", "name")
        cs_staff_list = fetch_cs_staff(conn)

    product_labels = ["unknown --"] + [name for _, name in products]
    supplier_labels = ["unknown --"] + [name for _, name in suppliers]
    machine_labels = ["unknown --"] + [name for _, name in machines]
    customer_labels = ["unknown --"] + [name for _, name in customers]
    staff_labels = ["not selected --"] + [f"{name} ({role})" for _, name, role in cs_staff_list]
    brand_labels = ["unknown --"] + BRANDS

    rc_codes = fetch_existing_codes(conn, "Root Cause")
    capa_codes = fetch_existing_codes(conn, "CAPA")
    rc_labels = ["let AI classify --"] + [f"{c} — {n}" for c, n in rc_codes]
    capa_labels = ["none --"] + [f"{c} — {n}" for c, n in capa_codes]

    with zone_card("amber"):
        st.caption(
            "Extra Root Cause & CAPA counts (if known upfront) — placed outside the form due to a Streamlit limitation."
        )
        col_extra1, col_extra2 = st.columns(2)
        with col_extra1:
            if "prefill_complaint_extra_rc_count" in st.session_state:
                st.session_state["extra_rc_count"] = st.session_state.pop("prefill_complaint_extra_rc_count")
            extra_rc_count = st.number_input(
                "Other root causes",
                min_value=0, max_value=10, step=1, value=0, key="extra_rc_count",
                help="Number of other root causes — extra fields appear next to the main one below.",
            )
        with col_extra2:
            if "prefill_complaint_extra_capa_count" in st.session_state:
                st.session_state["extra_capa_count"] = st.session_state.pop("prefill_complaint_extra_capa_count")
            extra_capa_count = st.number_input(
                "Other CAPAs",
                min_value=0, max_value=10, step=1, value=0, key="extra_capa_count",
                help="Number of other CAPAs — extra fields appear next to the main one below.",
            )

    with st.form("newcomplaintform"):
        if "prefill_complaint_desc" in st.session_state:
            st.session_state["newcomplaint_desc_value"] = st.session_state.pop("prefill_complaint_desc")
            st.caption(
                "The fields below were prefilled from your question in the 'Ask AI' tab — review before saving."
            )
        with zone_card("teal"):
            section_header("📝", "Issue description & photo", "teal")
            desc = st.text_area(
                "Issue description",
                height=100, key="newcomplaint_desc_value",
                help="The more detail the better — this is what the system uses to auto-classify and match a code.",
            )
            with validity_radio_box():
                complaint_validity_new = st.radio(
                    "Initial assessment",
                    COMPLAINT_VALIDITY_OPTIONS,
                    horizontal=True, key="newcomplaint_validity_value",
                    help="Is this a genuine Nilorn defect, or an issue caused by the customer themselves that still resulted in a complaint? Can be corrected later on the detail page.",
                )
            defect_photo_input = st.file_uploader(
                "Defect photo (optional)",
                type=["png", "jpg", "jpeg"], key="newcomplaint_defect_photo",
                help="A real photo of this defect — will be embedded in the Word report for easier understanding.",
            )
            render_paste_zone("Defect illustration photo (optional)")

        col1, col2 = st.columns(2)
        with col1:
            with zone_card("blue"):
                section_header("📦", "Product & supplier", "blue")
                if "prefill_complaint_date" in st.session_state:
                    st.session_state["newcomplaint_date_value"] = st.session_state.pop("prefill_complaint_date")
                date_opened_new = st.date_input("Date occurred (required)", key="newcomplaint_date_value")

                if "prefill_complaint_product" in st.session_state:
                    match_label = find_best_label_match(st.session_state.pop("prefill_complaint_product"), product_labels)
                    if match_label:
                        st.session_state["newcomplaint_product_value"] = match_label
                product_choice = st.selectbox("Product (if existing)", product_labels, key="newcomplaint_product_value")
                new_product_name = st.text_input("...or enter a NEW product name/code")
                new_product_group = st.selectbox("Product Group for the new product (only if entered above)", PRODUCT_GROUPS)

                if "prefill_complaint_supplier" in st.session_state:
                    match_label = find_best_label_match(st.session_state.pop("prefill_complaint_supplier"), supplier_labels)
                    if match_label:
                        st.session_state["newcomplaint_supplier_value"] = match_label
                supplier_choice = st.selectbox("Supplier", supplier_labels, key="newcomplaint_supplier_value")

                if "prefill_complaint_brand" in st.session_state:
                    match_label = find_best_label_match(st.session_state.pop("prefill_complaint_brand"), brand_labels)
                    if match_label:
                        st.session_state["newcomplaint_brand_value"] = match_label
                brand_choice = st.selectbox("Brand", brand_labels, key="newcomplaint_brand_value")
        with col2:
            with zone_card("coral"):
                section_header("🧾", "Order & customer", "coral")
                if "prefill_complaint_so_po" in st.session_state:
                    st.session_state["newcomplaint_so_po_value"] = st.session_state.pop("prefill_complaint_so_po")
                so_po_new = st.text_input("Sales Order No.", key="newcomplaint_so_po_value")
                if "prefill_complaint_lot" in st.session_state:
                    st.session_state["newcomplaint_lot_value"] = st.session_state.pop("prefill_complaint_lot")
                lot_new = st.text_input("Purchase Order No.", key="newcomplaint_lot_value")
                if "prefill_complaint_customer" in st.session_state:
                    match_label = find_best_label_match(st.session_state.pop("prefill_complaint_customer"), customer_labels)
                    if match_label:
                        st.session_state["newcomplaint_customer_value"] = match_label
                customer_choice = st.selectbox("Customer", customer_labels, key="newcomplaint_customer_value")
                new_customer_name = st.text_input("...or enter a NEW customer name")

        with zone_card("gray"):
            section_header("👤", "Recorded by & quantities", "gray")
            if "prefill_complaint_staff" in st.session_state:
                match_label = find_best_label_match(st.session_state.pop("prefill_complaint_staff"), staff_labels)
                if match_label:
                    st.session_state["newcomplaint_staff_value"] = match_label
            staff_choice = st.selectbox("Recorded by", staff_labels, key="newcomplaint_staff_value")

            col3, col4 = st.columns(2)
            with col3:
                qty_inspected_new = st.number_input("Quantity inspected", min_value=0, step=1, value=0)
            with col4:
                if "prefill_complaint_quantity" in st.session_state:
                    st.session_state["newcomplaint_qty_affected_value"] = st.session_state.pop("prefill_complaint_quantity")
                qty_affected_new = st.number_input(
                    "Quantity affected", min_value=0, step=1, key="newcomplaint_qty_affected_value",
                )

        st.markdown("---")
        with zone_card("amber"):
            section_header("🔎", "optional)", "amber")
            st.caption(
                "Fill in if known, skip otherwise (AI will auto-classify the root cause by default)."
            )
            if "prefill_complaint_rc" in st.session_state:
                prefill_rc_code = st.session_state.pop("prefill_complaint_rc")
                match_label = next((lbl for lbl in rc_labels if lbl.startswith(prefill_rc_code + " —")), None)
                if match_label:
                    st.session_state["newcomplaint_rc_value"] = match_label
            rc_choice = st.selectbox("Root cause (if code already known)", rc_labels, key="newcomplaint_rc_value")
            if "prefill_complaint_rc_new" in st.session_state:
                st.session_state["newcomplaint_new_rc_text_value"] = st.session_state.pop("prefill_complaint_rc_new")
            new_rc_text = st.text_input(
                "...or describe a NEW root cause (if no code fits — will be sent to the approval queue)",
                key="newcomplaint_new_rc_text_value",
            )

            extra_rc_entries = []
            for i in range(extra_rc_count):
                st.caption(f"Other Root Cause #{i + 1}")
                if f"prefill_complaint_extra_rc_{i}_code" in st.session_state:
                    prefill_code_i = st.session_state.pop(f"prefill_complaint_extra_rc_{i}_code")
                    match_label_i = next((lbl for lbl in rc_labels if lbl.startswith(prefill_code_i + " —")), None)
                    st.session_state[f"extra_rc_choice_{i}"] = match_label_i or "let AI classify --"
                if f"prefill_complaint_extra_rc_{i}_new" in st.session_state:
                    st.session_state[f"extra_rc_text_{i}"] = st.session_state.pop(f"prefill_complaint_extra_rc_{i}_new")
                ec1, ec2 = st.columns(2)
                with ec1:
                    choice_i = st.selectbox(f"Root cause #{i + 1}(if code already known)", rc_labels, key=f"extra_rc_choice_{i}")
                with ec2:
                    text_i = st.text_input(f"...or describe a NEW root cause #{i + 1}", key=f"extra_rc_text_{i}")
                extra_rc_entries.append((choice_i, text_i))

            col5, col6 = st.columns(2)
            with col5:
                if "prefill_complaint_capa" in st.session_state:
                    prefill_capa_code = st.session_state.pop("prefill_complaint_capa")
                    match_label = next((lbl for lbl in capa_labels if lbl.startswith(prefill_capa_code + " —")), None)
                    if match_label:
                        st.session_state["newcomplaint_capa_value"] = match_label
                capa_choice = st.selectbox(
                    "CAPA — corrective action (if code already known)",
                    capa_labels, key="newcomplaint_capa_value",
                )
                if "prefill_complaint_capa_new" in st.session_state:
                    st.session_state["newcomplaint_new_capa_text_value"] = st.session_state.pop("prefill_complaint_capa_new")
                new_capa_text = st.text_input(
                    "...or describe a NEW CAPA (if no code fits — will be sent to the approval queue)",
                    key="newcomplaint_new_capa_text_value",
                )
                if "prefill_complaint_capa_responsible" in st.session_state:
                    st.session_state["newcomplaint_capa_responsible_value"] = st.session_state.pop("prefill_complaint_capa_responsible")
                capa_responsible = st.text_input("CAPA responsible person", key="newcomplaint_capa_responsible_value")
            with col6:
                capa_date = st.date_input("CAPA proposed date", value=date_opened_new)

            extra_capa_entries = []
            for i in range(extra_capa_count):
                st.caption(f"Other CAPA #{i + 1}")
                if f"prefill_complaint_extra_capa_{i}_code" in st.session_state:
                    prefill_code_i = st.session_state.pop(f"prefill_complaint_extra_capa_{i}_code")
                    match_label_i = next((lbl for lbl in capa_labels if lbl.startswith(prefill_code_i + " —")), None)
                    st.session_state[f"extra_capa_choice_{i}"] = match_label_i or "none --"
                if f"prefill_complaint_extra_capa_{i}_new" in st.session_state:
                    st.session_state[f"extra_capa_text_{i}"] = st.session_state.pop(f"prefill_complaint_extra_capa_{i}_new")
                if f"prefill_complaint_extra_capa_{i}_resp" in st.session_state:
                    st.session_state[f"extra_capa_resp_{i}"] = st.session_state.pop(f"prefill_complaint_extra_capa_{i}_resp")
                ec3, ec4 = st.columns(2)
                with ec3:
                    capa_choice_i = st.selectbox(
                        f"CAPA #{i + 1}(if code already known)", capa_labels, key=f"extra_capa_choice_{i}",
                    )
                    capa_text_i = st.text_input(f"...or describe a NEW CAPA #{i + 1}", key=f"extra_capa_text_{i}")
                with ec4:
                    capa_resp_i = st.text_input(f"CAPA Responsible Person #{i + 1}", key=f"extra_capa_resp_{i}")
                    capa_date_i = st.date_input(f"CAPA Proposed Date #{i + 1}", value=date_opened_new, key=f"extra_capa_date_{i}")
                extra_capa_entries.append((capa_choice_i, capa_text_i, capa_resp_i, capa_date_i))

        submitted_new = st.form_submit_button("Classify & Save")

    if submitted_new:
        if not desc.strip():
            st.warning("You haven't entered an issue description.")
        else:
            ai_client = get_ai_client()
            product_id_new = None if product_choice.startswith("--") else products[product_labels.index(product_choice) - 1][0]

            if new_product_name.strip():
                typed_norm = normalize_code(new_product_name)
                existing_match = next(
                    (pid for pid, name in products if normalize_code(name) == typed_norm),
                    None,
                )
                if existing_match:
                    product_id_new = existing_match
                    st.info(f"Product '{new_product_name.strip()}Product already exists (even if formatted differently) — reusing it, not creating a duplicate.")
                else:
                    close_matches = [
                        name for _, name in products
                        if difflib.SequenceMatcher(None, normalize_code(name), typed_norm).ratio() > 0.85
                    ]
                    if close_matches:
                        st.warning(
                            f"⚠️ New product '{new_product_name.strip()}' looks similar to an existing one: "
                            f"**{', '.join(close_matches)}** — still creating it, please verify "
                            f"it's not a duplicate."
                        )
                    with conn.cursor() as cur:
                        cur.execute(
                            "insert into product (name, product_group) values (%s, %s) returning product_id;",
                            (new_product_name.strip(), new_product_group),
                        )
                        product_id_new = cur.fetchone()[0]
                    conn.commit()

            supplier_id_new = None if supplier_choice.startswith("--") else suppliers[supplier_labels.index(supplier_choice) - 1][0]
            machine_id_new = None
            customer_id_new = None if customer_choice.startswith("--") else customers[customer_labels.index(customer_choice) - 1][0]

            if new_customer_name.strip():
                typed_norm_c = normalize_code(new_customer_name)
                existing_match_c = next(
                    (cid for cid, name in customers if normalize_code(name) == typed_norm_c),
                    None,
                )
                if existing_match_c:
                    customer_id_new = existing_match_c
                    st.info(f"Customer '{new_customer_name.strip()}Customer already exists — reusing it, not creating a duplicate.")
                else:
                    close_matches_c = [
                        name for _, name in customers
                        if difflib.SequenceMatcher(None, normalize_code(name), typed_norm_c).ratio() > 0.85
                    ]
                    if close_matches_c:
                        st.warning(
                            f"⚠️ New customer '{new_customer_name.strip()}' looks similar to an existing one: "
                            f"**{', '.join(close_matches_c)}** — please verify it's not a duplicate."
                        )
                    with conn.cursor() as cur:
                        cur.execute(
                            "insert into customer (name) values (%s) returning customer_id;",
                            (new_customer_name.strip(),),
                        )
                        customer_id_new = cur.fetchone()[0]
                    conn.commit()

            staff_id_new = None if staff_choice.startswith("--") else cs_staff_list[staff_labels.index(staff_choice) - 1][0]
            brand_new = None if brand_choice.startswith("--") else brand_choice
            defect_photo_b64 = None
            if defect_photo_input is not None:
                defect_photo_b64 = base64.b64encode(defect_photo_input.getvalue()).decode("utf-8")

            with st.spinner("Saving and classifying..."):
                with conn.cursor() as cur:
                    cur.execute(
                        """insert into complaint
                               (date_opened, source, product_id, supplier_id, machine_id,
                                so_po, lot_number, quantity_inspected, quantity_affected, notes, status,
                                brand, customer_id, recorded_by, defect_photo, bear_the_claim, complaint_validity)
                           values (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 'Thiếu Customer', %s, %s, %s, %s, %s, %s)
                           returning complaint_id;""",
                        (date_opened_new, "Nhập tay (form)", product_id_new, supplier_id_new, machine_id_new,
                         so_po_new or None, lot_new or None,
                         qty_inspected_new or None, qty_affected_new or None, desc,
                         brand_new, customer_id_new, staff_id_new, defect_photo_b64, "100% Nilorn",
                         complaint_validity_new),
                    )
                    new_complaint_id = cur.fetchone()[0]
                conn.commit()
                compute_and_update_complaint_status(conn, new_complaint_id)

                summary_lines = []
                used_rc_codes = []
                used_capa_codes = []
                for kind, column in [("Defect", "defect_code"), ("Root Cause", "root_cause_code")]:
                    if kind == "Root Cause" and not rc_choice.startswith("--"):
                        manual_rc_code = rc_choice.split(" — ")[0]
                        with conn.cursor() as cur:
                            cur.execute(
                                f"update complaint set {column} = %s where complaint_id = %s;",
                                (manual_rc_code, new_complaint_id),
                            )
                        conn.commit()
                        add_root_cause_to_complaint(conn, new_complaint_id, manual_rc_code)
                        used_rc_codes.append(manual_rc_code)
                        summary_lines.append(f"✅ **{kind}**: manually selected — **{manual_rc_code}manually selected — **{manual_rc_code}** (AI skipped)")
                        continue

                    if kind == "Root Cause" and new_rc_text.strip():
                        rc_taxonomy_list = load_taxonomy_list(conn, "Root Cause")
                        check_result = classify_description(ai_client, new_rc_text.strip(), rc_taxonomy_list, "Root Cause")
                        conn = ensure_connection()
                        closest = check_result.get("closest_existing_code") or check_result.get("matched_code")
                        note = (
                            "Manually described by staff when entering the complaint — no matching code in the taxonomy yet. "
                            f"AI cross-check: {check_result.get('reasoning', '(no notes)')}"
                        )
                        with conn.cursor() as cur:
                            cur.execute(
                                """insert into taxonomy_suggestion
                                       (complaint_id, suggestion_type, ai_suggested_name, ai_reasoning, closest_existing_code)
                                   values (%s, %s, %s, %s, %s);""",
                                (new_complaint_id, "Root Cause", new_rc_text.strip(), note, closest),
                            )
                        conn.commit()
                        if closest:
                            summary_lines.append(
                                f"⚠️ **{kind}**: sent to the queue — AI found it **may relate to code {closest}sent to the queue — AI found it **may relate to code "
                                f"{closest}**, the reviewer will check during approval"
                            )
                        else:
                            summary_lines.append(
                                f"🕓 **{kind}new root cause as described — AI confirmed no related code exists, sent to the approval queue"
                            )
                        continue

                    taxonomy_list = load_taxonomy_list(conn, kind)
                    result = classify_description(ai_client, desc, taxonomy_list, kind)
                    conn = ensure_connection()

                    if result.get("matched_code") and result.get("confidence") == "High":
                        with conn.cursor() as cur:
                            cur.execute(
                                f"update complaint set {column} = %s where complaint_id = %s;",
                                (result["matched_code"], new_complaint_id),
                            )
                        conn.commit()
                        if kind == "Root Cause":
                            add_root_cause_to_complaint(conn, new_complaint_id, result["matched_code"])
                            used_rc_codes.append(result["matched_code"])
                        summary_lines.append(f"✅ **{kind}**: matched existing — **{result['matched_code']}** / matched existing code — **{result['matched_code']}**")
                    else:
                        # AI đôi khi không tự đặt được tên đề xuất (suggested_name = null) khi mô tả quá
                        # mơ hồ — cột ai_suggested_name trong DB không cho phép NULL, nên luôn cần giá trị
                        # dự phòng lấy từ mô tả gốc, tránh insert thất bại giữa chừng. / AI sometimes
                        # returns no suggested_name when the description is too vague — the DB column is
                        # NOT NULL, so always fall back to the original description to avoid a failed insert.
                        fallback_name = truncate_at_word(desc, 100) if desc else "unnamed)"
                        safe_name = (result.get("suggested_name") or "").strip() or fallback_name
                        with conn.cursor() as cur:
                            cur.execute(
                                """insert into taxonomy_suggestion
                                       (complaint_id, suggestion_type, ai_suggested_name, ai_reasoning, closest_existing_code)
                                   values (%s, %s, %s, %s, %s);""",
                                (new_complaint_id, kind, safe_name,
                                 result.get("reasoning"), result.get("closest_existing_code")),
                            )
                        conn.commit()
                        summary_lines.append(
                            f"🕓 **{kind}uncertain (may be new) — sent to the approval queue, see the 'Review Taxonomy' tab"
                        )

                if not capa_choice.startswith("--"):
                    manual_capa_code = capa_choice.split(" — ")[0]
                    with conn.cursor() as cur:
                        cur.execute(
                            """insert into capa_action
                                   (complaint_id, capa_code, date_proposed, responsible_party, verification_result)
                               values (%s, %s, %s, %s, 'Pending');""",
                            (new_complaint_id, manual_capa_code, capa_date, capa_responsible or None),
                        )
                    conn.commit()
                    summary_lines.append(f"📌 **CAPA**: added — **{manual_capa_code}**")
                    used_capa_codes.append(manual_capa_code)
                elif new_capa_text.strip():
                    capa_taxonomy_list = load_taxonomy_list(conn, "CAPA")
                    check_result = classify_description(ai_client, new_capa_text.strip(), capa_taxonomy_list, "CAPA")
                    conn = ensure_connection()
                    closest = check_result.get("closest_existing_code") or check_result.get("matched_code")
                    note = (
                        f"Proposed by staff when entering the complaint. Expected responsible person: {capa_responsible or '(unknown)'}. "
                        f"AI cross-check: {check_result.get('reasoning', '(no notes)')}"
                    )
                    with conn.cursor() as cur:
                        cur.execute(
                            """insert into taxonomy_suggestion
                                   (complaint_id, suggestion_type, ai_suggested_name, ai_reasoning, closest_existing_code, responsible_party)
                               values (%s, %s, %s, %s, %s, %s);""",
                            (new_complaint_id, "CAPA", new_capa_text.strip(), note, closest, capa_responsible or None),
                        )
                    conn.commit()
                    if closest:
                        summary_lines.append(
                            f"⚠️ **CAPA**: sent to the queue — AI found it may relate to **{closest}**, the reviewer will check during approval"
                        )
                    else:
                        summary_lines.append("new action as described — sent to the approval queue")

                for idx, (rc_choice_i, rc_text_i) in enumerate(extra_rc_entries, start=1):
                    if not rc_choice_i.startswith("--"):
                        rc_code_i = rc_choice_i.split(" — ")[0]
                        add_root_cause_to_complaint(conn, new_complaint_id, rc_code_i)
                        used_rc_codes.append(rc_code_i)
                        summary_lines.append(f"✅ **Other Root Cause #{idx}**: manually selected — **{rc_code_i}** / manually selected — **{rc_code_i}**")
                    elif rc_text_i.strip():
                        rc_taxonomy_list_i = load_taxonomy_list(conn, "Root Cause")
                        check_result_i = classify_description(
                            ai_client, rc_text_i.strip(), rc_taxonomy_list_i, "Root Cause", exclude_codes=used_rc_codes,
                        )
                        conn = ensure_connection()
                        if check_result_i.get("matched_code") and check_result_i.get("confidence") == "High":
                            add_root_cause_to_complaint(conn, new_complaint_id, check_result_i["matched_code"])
                            used_rc_codes.append(check_result_i["matched_code"])
                            summary_lines.append(
                                f"✅ **Other Root Cause #{idx}**: AI auto-matched existing — **{check_result_i['matched_code']}**"
                            )
                        else:
                            closest_i = check_result_i.get("closest_existing_code") or check_result_i.get("matched_code")
                            note_i = (
                                f"Other Root Cause #{idx} — manually described by staff when entering the complaint, no matching code yet. "
                                f"AI cross-check: {check_result_i.get('reasoning', '(no notes)')}"
                            )
                            with conn.cursor() as cur:
                                cur.execute(
                                    """insert into taxonomy_suggestion
                                           (complaint_id, suggestion_type, ai_suggested_name, ai_reasoning, closest_existing_code)
                                       values (%s, %s, %s, %s, %s);""",
                                    (new_complaint_id, "Root Cause", rc_text_i.strip(), note_i, closest_i),
                                )
                            conn.commit()
                            if closest_i:
                                summary_lines.append(
                                    f"⚠️ **Other Root Cause #{idx}**: sent to the queue — "
                                    f"may relate to **{closest_i}**"
                                )
                            else:
                                summary_lines.append(f"🕓 **Other Root Cause #{idx}**: new root cause, sent to the approval queue")
                    else:
                        rc_taxonomy_list_i = load_taxonomy_list(conn, "Root Cause")
                        check_result_i = classify_description(
                            ai_client, desc, rc_taxonomy_list_i, "Root Cause", exclude_codes=used_rc_codes,
                        )
                        conn = ensure_connection()
                        if (check_result_i.get("matched_code")
                                and check_result_i.get("confidence") == "High"
                                and check_result_i["matched_code"] not in used_rc_codes):
                            add_root_cause_to_complaint(conn, new_complaint_id, check_result_i["matched_code"])
                            used_rc_codes.append(check_result_i["matched_code"])
                            summary_lines.append(
                                f"✅ **Other Root Cause #{idx}**: AI additionally found from the main description — **{check_result_i['matched_code']}**"
                            )
                        else:
                            summary_lines.append(f"➖ **Other Root Cause #{idx}AI found no additional root cause — left blank")

                for idx, (capa_choice_i, capa_text_i, capa_resp_i, capa_date_i) in enumerate(extra_capa_entries, start=1):
                    if not capa_choice_i.startswith("--"):
                        capa_code_i = capa_choice_i.split(" — ")[0]
                        add_capa_action_to_complaint(conn, new_complaint_id, capa_code_i, capa_date_i, capa_resp_i or None)
                        used_capa_codes.append(capa_code_i)
                        summary_lines.append(f"📌 **Other CAPA #{idx}**: added — **{capa_code_i}** / added — **{capa_code_i}**")
                    elif capa_text_i.strip():
                        capa_taxonomy_list_i = load_taxonomy_list(conn, "CAPA")
                        check_result_i = classify_description(
                            ai_client, capa_text_i.strip(), capa_taxonomy_list_i, "CAPA", exclude_codes=used_capa_codes,
                        )
                        conn = ensure_connection()
                        if check_result_i.get("matched_code") and check_result_i.get("confidence") == "High":
                            add_capa_action_to_complaint(conn, new_complaint_id, check_result_i["matched_code"], capa_date_i, capa_resp_i or None)
                            used_capa_codes.append(check_result_i["matched_code"])
                            summary_lines.append(
                                f"📌 **Other CAPA #{idx}**: AI auto-matched existing — **{check_result_i['matched_code']}**"
                            )
                        else:
                            closest_i = check_result_i.get("closest_existing_code") or check_result_i.get("matched_code")
                            note_i = (
                                f"Other CAPA #{idx} — proposed by staff when entering the complaint. "
                                f"Expected responsible person: {capa_resp_i or '(unknown)'}. "
                                f"AI cross-check: {check_result_i.get('reasoning', '(no notes)')}"
                            )
                            with conn.cursor() as cur:
                                cur.execute(
                                    """insert into taxonomy_suggestion
                                           (complaint_id, suggestion_type, ai_suggested_name, ai_reasoning, closest_existing_code, responsible_party)
                                       values (%s, %s, %s, %s, %s, %s);""",
                                    (new_complaint_id, "CAPA", capa_text_i.strip(), note_i, closest_i, capa_resp_i or None),
                                )
                            conn.commit()
                            if closest_i:
                                summary_lines.append(
                                    f"⚠️ **Other CAPA #{idx}**: sent to the queue — "
                                    f"may relate to **{closest_i}**"
                                )
                            else:
                                summary_lines.append(f"🕓 **Other CAPA #{idx}**: new action, sent to the approval queue")

            # Gợi ý nhà cung cấp liên quan (dựa theo lịch sử) — chỉ khi CS KHÔNG chọn sẵn nhà cung cấp lúc
            # nhập complaint, để hướng dẫn bước tiếp theo (gửi email yêu cầu điều tra) ngay cả khi complaint
            # được tạo bằng form nhập tay (trước đây tính năng này chỉ hiện khi tạo từ luồng đọc email).
            supplier_suggestion_new = None
            if supplier_choice.startswith("--"):
                product_name_for_suggestion = (
                    new_product_name.strip() if new_product_name.strip()
                    else (None if product_choice.startswith("--") else product_choice)
                )
                if product_name_for_suggestion:
                    supplier_suggestion_new = suggest_supplier_for_product(conn, product_name_for_suggestion)

            # Báo reviewer NGAY nếu có mục nào vừa được đưa vào hàng đợi chờ duyệt (khớp cơ chế
            # "gửi ngay lúc nộp" đã dùng cho luồng NCC tự khai báo — không còn chờ ai mở tab Duyệt
            # Taxonomy mới được báo nữa).
            pending_lines_new = [
                ln for ln in summary_lines
                if "approval queue" in ln or "sent to the queue" in ln
            ]
            if pending_lines_new and APPROVER_EMAILS:
                clean_lines = [re.sub(r"\*\*", "", ln) for ln in pending_lines_new]
                approver_mail_ok_new, approver_mail_err_new = send_notification_email(
                    APPROVER_EMAILS,
                    f"[Nilorn Internal AI] {len(pending_lines_new)} new suggestion(s) need approval (New Complaint, manual entry)",
                    "New suggestion(s) pending approval, just created from a manually entered complaint:\n\n"
                    + "\n".join(f"- {ln}" for ln in clean_lines)
                    + f"\n\nOpen the app, 'Review Taxonomy' tab to view and process: {REVIEW_APP_URL}",
                )
                _log_notify_attempt(conn, new_complaint_id, APPROVER_EMAILS, approver_mail_ok_new, approver_mail_err_new)

            # Báo CS chung (LUÔN gửi) + đúng người được ghi nhận (nếu có chọn) ngay khi complaint mới được tạo.
            notify_new_complaint_recorded(conn, new_complaint_id, staff_id_new, so_po_new)

            st.session_state["new_complaint_result"] = {
                "complaint_id": new_complaint_id,
                "summary_lines": summary_lines,
                "date_opened": date_opened_new,
                "desc": desc,
                "product": new_product_name.strip() if new_product_name.strip() else product_choice,
                "supplier": supplier_choice,
                "supplier_suggestion": supplier_suggestion_new,
                "brand": brand_choice,
                "customer": new_customer_name.strip() if new_customer_name.strip() else customer_choice,
                "so_po": so_po_new,
                "lot": lot_new,
                "qty_inspected": qty_inspected_new,
                "qty_affected": qty_affected_new,
                "rc_choice": rc_choice,
                "capa_choice": capa_choice,
                "capa_responsible": capa_responsible,
                "staff": staff_choice,
            }
            st.rerun()

    result_new = st.session_state.get("new_complaint_result")
    if result_new:
        st.success(f"New complaint saved — system ID: `{result_new['complaint_id']}`")
        for line in result_new["summary_lines"]:
            st.write(line)

        if result_new.get("supplier_suggestion"):
            st.info(
                f"💡 Suggested related supplier (based on past complaint history): "
                f"**{result_new['supplier_suggestion']}**. Contact them directly to request an investigation."
            )
        else:
            st.info(
                "No related supplier identified yet — check the **'📊 Data Lookup'** tab if needed."
            )

        st.info(
            "To download a Word report for this complaint — even after the Root Cause/CAPA get approved — go to **'📊 Data Lookup'** → **'Export full report for a complaint'**, and select the complaint you just entered."
        )



# ------------------------------------------------------------
# DASHBOARD — trang chính: thẻ thống kê, biểu đồ, top 5 NCC, card complaint (bấm để xem chi tiết).
# ------------------------------------------------------------
if page == "dashboard":
    from collections import Counter
    from datetime import date as _date

    conn = ensure_connection()

    st.markdown(
        """<style>
div[data-testid="stRadio"] > div[role="radiogroup"] {
    background: #ffffff; border-radius: 24px; padding: 3px; border: 1px solid #e5e3da;
    display: inline-flex; width: fit-content; gap: 0;
}
div[data-testid="stRadio"] label {
    background: transparent; border-radius: 20px; padding: 5px 16px !important; margin: 0 !important;
    cursor: pointer; transition: background 0.15s ease;
}
div[data-testid="stRadio"] label:has(input:checked) { background: #7F77DD; }
div[data-testid="stRadio"] label:has(input:checked) p { color: #ffffff !important; font-weight: 600; }
div[data-testid="stRadio"] label div[data-testid="stMarkdownContainer"] p { font-size: 13px; margin: 0; }
div[data-testid="stRadio"] label > div:first-child { display: none; }
</style>""",
        unsafe_allow_html=True,
    )

    st.markdown(
        f'<div style="font-size:30px;font-weight:800;line-height:1.15;">Dashboard</div>',
        unsafe_allow_html=True,
    )
    _subtitle_placeholder = st.empty()

    cols_db_early, rows_db_early = fetch_all_submissions(conn)
    rows_db_early = rows_db_early + fetch_legacy_complaints_for_dashboard(conn, cols_db_early)

    period_col, picker_col, filler_col, dl_col = st.columns([1.3, 1.6, 1.9, 1.3])
    with period_col:
        period = st.radio(
            "Time Period", ["Month", "Quarter", "Year"],
            horizontal=True, key="dashboard_period", label_visibility="collapsed",
        )

    _today_for_label = _date.today()

    def _month_start(y, m):
        return _date(y, m, 1)

    def _month_end(y, m):
        import calendar as _cal
        return _date(y, m, _cal.monthrange(y, m)[1])

    period_options = []  # list of (label, start_date, end_date)
    if period == "Month":
        y, m = _today_for_label.year, _today_for_label.month
        for _ in range(12):
            period_options.append((_month_start(y, m).strftime("%B %Y"), _month_start(y, m), _month_end(y, m)))
            m -= 1
            if m == 0:
                m = 12
                y -= 1
    elif period == "Quarter":
        y = _today_for_label.year
        q = (_today_for_label.month - 1) // 3 + 1
        for _ in range(8):
            q_start_month = (q - 1) * 3 + 1
            q_end_month = q_start_month + 2
            period_options.append((f"Q{q} {y}", _month_start(y, q_start_month), _month_end(y, q_end_month)))
            q -= 1
            if q == 0:
                q = 4
                y -= 1
    else:
        y = _today_for_label.year
        for _ in range(5):
            period_options.append((str(y), _date(y, 1, 1), _date(y, 12, 31)))
            y -= 1

    # Luôn sắp xếp lại theo ngày bắt đầu giảm dần (mới nhất trước) — lớp bảo vệ độc lập với thứ tự
    # sinh ra ở trên, đảm bảo dropdown không bao giờ hiện lộn xộn dù logic sinh danh sách có đổi sau này.
    period_options.sort(key=lambda x: x[1], reverse=True)

    with picker_col:
        picked_idx = st.selectbox(
            "Select Period", list(range(len(period_options))),
            format_func=lambda i: period_options[i][0],
            key=f"period_picker_{period}", label_visibility="collapsed",
        )
    picked_label, period_start, period_end_selected = period_options[picked_idx]
    period_end = min(period_end_selected, _today_for_label)

    _subtitle_placeholder.markdown(
        f'<div style="font-size:13px;color:#73726c;margin:2px 0 14px;">{picked_label}</div>',
        unsafe_allow_html=True,
    )

    with dl_col:
        try:
            idx_early = {name: i for i, name in enumerate(cols_db_early)}
            normalized_all = [
                normalize_submission_row_for_excel(cols_db_early, r) for r in rows_db_early
                if not str(r[idx_early["submission_id"]]).startswith("legacy:")
            ] + fetch_legacy_complaints_for_excel(conn)
            if normalized_all:
                excel_bytes_early = build_register_excel(normalized_all)
                st.download_button(
                    "⬇️ Excel", data=excel_bytes_early,
                    file_name=f"Supplier_Return_Register_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        except Exception as e:
            st.caption(f"Excel error: {e}")

    cols_db, rows_db = cols_db_early, rows_db_early

    if not rows_db:
        st.info("No supplier reports yet.")
    else:
        idx = {name: i for i, name in enumerate(cols_db)}
        period_rows = [
            r for r in rows_db
            if r[idx["record_date"]] and period_start <= r[idx["record_date"]] <= period_end
        ]

        def _compute_missing_tags(r):
            """Tính trực tiếp từ dữ liệu hiện tại (không đọc complaint.status đã lưu) — đảm bảo
            luôn đúng ngay lập tức sau khi CS lưu bất kỳ trường nào, không bị trễ/lệch."""
            tags = []
            if r[idx["customer_name"]] is None:
                tags.append("Thiếu Customer")
            if r[idx["client_code"]] is None:
                tags.append("Thiếu Client")
            if r[idx["replacement_cost"]] is None:
                tags.append("Thiếu Replacement Cost")
            if r[idx["has_pending_verification"]] or not r[idx["has_implemented_capa"]]:
                tags.append("Thiếu xác minh CAPA")
            return tags

        total = len(period_rows)
        suppliers_involved = len(set(
            r[idx["vendor_code"]] or r[idx["supplier_name_raw"]] for r in period_rows
        ))
        missing_info = sum(1 for r in period_rows if _compute_missing_tags(r))

        stat_col1, stat_col2, stat_col3 = st.columns(3)
        with stat_col1:
            st.markdown(stat_card_html(f"Complaints this {period.lower()}", total, "gray"), unsafe_allow_html=True)
        with stat_col2:
            st.markdown(stat_card_html("Suppliers involved", suppliers_involved, "gray"), unsafe_allow_html=True)
        with stat_col3:
            st.markdown(
                stat_card_html("Not closed yet", missing_info, "coral" if missing_info else "teal"),
                unsafe_allow_html=True,
            )

        chart_col, top5_col = st.columns([1.4, 1])
        with chart_col:
            section_header("📈", "Complaints over time (last 6 months)", "blue")
            month_labels = []
            month_counts = []
            for i in range(5, -1, -1):
                m = _today_for_label.month - i
                y = _today_for_label.year
                while m <= 0:
                    m += 12
                    y -= 1
                label = f"{m:02d}/{y}"
                month_labels.append(label)
                count = sum(
                    1 for r in rows_db
                    if r[idx["record_date"]] and r[idx["record_date"]].year == y and r[idx["record_date"]].month == m
                )
                month_counts.append(count)
            chart_df = pd.DataFrame({"Complaints": month_counts}, index=month_labels)
            st.bar_chart(chart_df, use_container_width=True)

        with top5_col:
            section_header("🏆", "Suppliers with the Most Complaints", "amber")
            supplier_counter = Counter()
            for r in period_rows:
                name = r[idx["vendor_name_matched"]] or r[idx["supplier_name_raw"]]
                supplier_counter[name] += 1
            top5 = supplier_counter.most_common(5)
            if not top5:
                st.caption("No data in this time period.")
            else:
                for rank, (name, count) in enumerate(top5, start=1):
                    badge_color = "coral" if rank == 1 else ("amber" if rank <= 3 else "teal")
                    st.markdown(
                        f'<div style="display:flex;justify-content:space-between;align-items:center;'
                        f'padding:6px 0;">'
                        f'<span style="font-size:13px;">{rank}. {name}</span>'
                        f'{render_badge(str(count), *DASHBOARD_BADGE_STYLE.get(badge_color, DASHBOARD_BADGE_STYLE["gray"]))}'
                        f'</div>',
                        unsafe_allow_html=True,
                    )

        st.markdown("---")
        section_header("📥", "Recent complaints — click to view details", "gray")

        n_missing = sum(1 for r in period_rows if _compute_missing_tags(r))
        n_closed = len(period_rows) - n_missing

        filter_pick = st.radio(
            "Filter by status",
            [f"Not Closed ({n_missing})", f"Closed ({n_closed})", f"All ({len(period_rows)})"],
            horizontal=True, key="dashboard_card_filter", label_visibility="collapsed",
        )

        if filter_pick.startswith("Chưa hoàn thành"):
            filtered_card_rows = [r for r in period_rows if _compute_missing_tags(r)]
        elif filter_pick.startswith("Closed"):
            filtered_card_rows = [r for r in period_rows if not _compute_missing_tags(r)]
        else:
            filtered_card_rows = period_rows

        if not filtered_card_rows:
            st.caption("No complaints match the current filter.")

        card_cols = st.columns(2)
        for i, r in enumerate(filtered_card_rows):
            with card_cols[i % 2]:
                missing_tags = _compute_missing_tags(r)
                zone_color = "coral" if missing_tags else "teal"
                with zone_card(zone_color):
                    if missing_tags:
                        badges_html = "".join(status_badge_html(t) for t in missing_tags)
                    else:
                        badges_html = status_badge_html("Closed")
                    supplier_display = r[idx["vendor_name_matched"]] or r[idx["supplier_name_raw"]]
                    st.markdown(
                        f'<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:6px;">'
                        f'<span style="font-weight:600;font-size:14px;">{supplier_display}</span>'
                        f'<div style="display:flex;flex-wrap:wrap;gap:4px;justify-content:flex-end;">{badges_html}</div></div>',
                        unsafe_allow_html=True,
                    )
                    st.caption(f"SO {r[idx['sales_order_no']]} · Item {r[idx['item_no']]}")
                    st.caption(
                        f"Record Date {r[idx['record_date']].strftime('%d/%m/%Y') if r[idx['record_date']] else '(unknown)'}"
                    )
                    if st.button("View details →", key=f"card_detail_{r[idx['submission_id']]}"):
                        st.session_state.selected_submission_id = str(r[idx["submission_id"]])
                        st.rerun()
