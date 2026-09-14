import os
import re
import sqlite3
import calendar
from io import BytesIO
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

# Import modul untuk membaca dokumen (diperlukan untuk auto-fill judul)
try:
    import docx
except ImportError:
    docx = None

try:
    import PyPDF2
except ImportError:
    PyPDF2 = None


# =========================================================
# KONFIGURASI APLIKASI
# =========================================================

APP_TITLE = "Editorial Assistant — CT / HM / LC"
DB_FILE = "editorial_assistant.db"

JOURNALS = {
    "CT": {
        "name": "Computers, Technology and Education",
        "ojs_url": "https://journal.binus.ac.id/index.php/comtech/index",
    },
    "HM": {
        "name": "Humaniora",
        "ojs_url": "https://journal.binus.ac.id/index.php/humaniora/index",
    },
    "LC": {
        "name": "Lingua Cultura",
        "ojs_url": "https://journal.binus.ac.id/index.php/lingua/index",
    },
}

ARTICLE_STATUSES = [
    "New Submission",
    "Desk Review",
    "Assign Reviewer",
    "Under Review",
    "Review Feedback Received",
    "Revision Requested",
    "Revision Submitted",
    "Accepted",
    "APC Pending",
    "Payment Proof Submitted",
    "Payment Verified",
    "Ready Copy Edit",
    "Copy Editing",
    "Layout Editing",
    "Proofreading",
    "Scheduled",
    "Published",
    "Rejected",
    "Withdrawn",
]

REVIEWER_STATUSES = [
    "Invited",
    "Accepted",
    "Declined",
    "Review Submitted",
    "Unresponsive",
]

REVIEW_RECOMMENDATIONS = [
    "",
    "Accept",
    "Minor Revision",
    "Major Revision",
    "Reject",
]

REVIEWER_CATEGORIES = ["Nasional", "Binus"]

FOLDER_STRUCTURE = [
    "01_Submission",
    "02_Desk_Review",
    "03_Reviewers",
    "04_Review_Results",
    "05_Decision_Letters",
    "06_Revisi_Author",
    "07_APC",
    "08_Copyediting",
    "09_Layout_Proof",
    "10_Published",
]

DOCUMENT_CHECKLIST = {
    "Manuskrip Asli": ["manuskrip"],
    "Format Form": ["format form", "format_form"],
    "Turnitin Report": ["turnitin"],
    "Desk Review 1": ["desk review_1", "desk_review_1"],
    "Revisi Desk Review 1": ["revisi desk review_1", "revisi_desk_review_1"],
    "Desk Review 2": ["desk review_2", "desk_review_2"],
    "Revisi Desk Review 2": ["revisi desk review_2", "revisi_desk_review_2"],
    "Desk Review 3": ["desk review_3", "desk_review_3"],
    "Revisi Desk Review 3": ["revisi desk review_3", "revisi_desk_review_3"],
    "Reviewer Form 1": ["reviewer form_1", "reviewer_form_1"],
    "Reviewer Feedback 1": ["reviewer feedback_1", "reviewer_feedback_1"],
    "Reviewer Form 2": ["reviewer form_2", "reviewer_form_2"],
    "Reviewer Feedback 2": ["reviewer feedback_2", "reviewer_feedback_2"],
    "Revision Form (Author's Response)": ["revision form", "author's response"],
    "Review Revision 1": ["review revision_1", "review_revision_1"],
    "Review Revision 2": ["review revision_2", "review_revision_2"],
    "Review Revision 3": ["review revision_3", "review_revision_3"],
    "APC": ["apc"],
    "Copyediting Ready": ["copyediting ready", "copyediting_ready"],
    "Copyedit Revision 1": ["copyedit revision_1", "copyedit_revision_1"],
    "Copyedit Revision 2": ["copyedit revision_2", "copyedit_revision_2"],
    "Copyedit Revision 3": ["copyedit revision_3", "copyedit_revision_3"],
    "Copyedit Revision 4": ["copyedit revision_4", "copyedit_revision_4"],
    "Copyedit Revision 5": ["copyedit revision_5", "copyedit_revision_5"],
}

# =========================================================
# KONFIGURASI UPLOAD DOKUMEN
# =========================================================

ALLOWED_UPLOAD_EXTENSIONS = [
    "doc", "docx", "pdf", "xlsx", "xls", "csv", "png", "jpg", "jpeg", "zip", "txt",
]

UPLOAD_FOLDER_RULES = {
    "01_Submission": [
        "manuskrip", "format form", "format_form", "turnitin"
    ],
    "02_Desk_Review": [
        "desk review", "desk_review", "revisi desk review", "revisi_desk_review"
    ],
    "03_Reviewers": ["reviewer invitation", "reviewer_invitation", "invitation", "undangan reviewer"],
    "04_Review_Results": ["reviewer form", "reviewer_form", "reviewer feedback", "reviewer_feedback"],
    "05_Decision_Letters": ["decision letter", "decision_letter", "editor decision", "keputusan editor"],
    "06_Revisi_Author": ["revision form", "author's response", "review revision", "review_revision"],
    "07_APC": ["apc", "payment", "proof of payment", "payment proof", "bukti bayar", "invoice"],
    "08_Copyediting": ["copyediting ready", "copyediting_ready", "copyedit revision", "copyedit_revision"],
    "09_Layout_Proof": ["layout", "proof", "galley"],
    "10_Published": ["published", "publication", "final published"],
}

# =========================================================
# KONFIGURASI JADWAL RUTIN
# =========================================================

DAILY_SCHEDULE_TEMPLATE = [
    {"Waktu": "09.00–09.30", "Kegiatan Utama": "Cek Dashboard, OJS, email resmi, dan email pribadi untuk CT, HM, dan LC. Tandai pekerjaan mendesak.", "Fokus Email": "Semua email baru. Prioritaskan yang urgent."},
    {"Waktu": "09.30–11.30", "Kegiatan Utama": "{tugas_pagi}", "Fokus Email": "-"},
    {"Waktu": "11.30–12.00", "Kegiatan Utama": "Membalas email author/reviewer dan mencatat komunikasi.", "Fokus Email": "Balasan singkat dan permintaan dokumen."},
    {"Waktu": "12.00–13.00", "Kegiatan Utama": "Istirahat", "Fokus Email": "-"},
    {"Waktu": "13.00–15.00", "Kegiatan Utama": "{tugas_siang}", "Fokus Email": "-"},
    {"Waktu": "15.00–15.30", "Kegiatan Utama": "Cek email dan menangani pertanyaan singkat.", "Fokus Email": "Cek email lanjutan & email pribadi author."},
    {"Waktu": "15.30–16.15", "Kegiatan Utama": "{tugas_sore}", "Fokus Email": "-"},
    {"Waktu": "16.15–16.45", "Kegiatan Utama": "Membalas email yang belum selesai.", "Fokus Email": "Menyelesaikan balasan & catat di log."},
    {"Waktu": "16.45–17.00", "Kegiatan Utama": "Update OJS, aplikasi, Excel, log komunikasi, dan daftar besok.", "Fokus Email": "Update Excel & buat daftar besok."}
]

WEEKLY_ROUTINE = {
    0: {"Hari": "Senin", "Jurnal": "HM", "09.30–11.30": "Desk review dan cek format artikel HM", "13.00–15.00": "Pemeriksaan revisi author HM", "15.30–16.15": "Kirim keputusan atau permintaan perbaikan HM"},
    1: {"Hari": "Selasa", "Jurnal": "CT", "09.30–11.30": "Desk review dan cek format artikel CT", "13.00–15.00": "Evaluasi hasil reviewer dan revisi author CT", "15.30–16.15": "Kirim keputusan atau komentar editorial CT"},
    2: {"Hari": "Rabu", "Jurnal": "LC", "09.30–11.30": "Desk review dan cek format artikel LC", "13.00–15.00": "Evaluasi hasil reviewer dan revisi author LC", "15.30–16.15": "Kirim keputusan atau komentar editorial LC"},
    3: {"Hari": "Kamis", "Jurnal": "Semua jurnal", "09.30–11.30": "Evaluasi hasil reviewer yang sudah masuk dari CT, HM, dan LC", "13.00–15.00": "Copy editing artikel accepted dari jurnal yang paling mendesak", "15.30–16.15": "Follow-up reviewer, author, dan editor lintas-jurnal"},
    4: {"Hari": "Jumat", "Jurnal": "Semua jurnal", "09.30–11.30": "Menyelesaikan backlog dari CT, HM, dan LC", "13.00–15.00": "Copy editing dan final check artikel accepted", "15.30–16.15": "Audit status, sinkronisasi data, dan rencana minggu depan"}
}

# =========================================================
# DATABASE
# =========================================================

def get_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            setting_key TEXT PRIMARY KEY,
            setting_value TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS articles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            article_code TEXT UNIQUE NOT NULL,
            journal_code TEXT NOT NULL,
            submission_id TEXT,
            title TEXT NOT NULL,
            corresponding_author TEXT,
            author_email TEXT,
            affiliation_1 TEXT,
            affiliation_2 TEXT,
            status TEXT NOT NULL DEFAULT 'New Submission',
            submission_date TEXT,
            desk_review_deadline TEXT,
            author_revision_deadline TEXT,
            author_reminder_count INTEGER DEFAULT 0,
            apc_deadline TEXT,
            payment_proof_date TEXT,
            payment_verified_date TEXT,
            volume TEXT,
            issue TEXT,
            publication_year TEXT,
            folder_path TEXT,
            notes TEXT,
            is_international INTEGER DEFAULT 0,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    try: cur.execute("ALTER TABLE articles ADD COLUMN is_international INTEGER DEFAULT 0")
    except sqlite3.OperationalError: pass

    cur.execute("""
        CREATE TABLE IF NOT EXISTS reviewers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            article_id INTEGER NOT NULL,
            reviewer_name TEXT NOT NULL,
            reviewer_email TEXT,
            reviewer_category TEXT NOT NULL DEFAULT 'Nasional',
            invitation_date TEXT,
            invitation_deadline TEXT,
            acceptance_date TEXT,
            review_deadline TEXT,
            review_submitted_date TEXT,
            status TEXT NOT NULL DEFAULT 'Invited',
            recommendation TEXT,
            reminder_count INTEGER DEFAULT 0,
            notes TEXT,
            FOREIGN KEY(article_id) REFERENCES articles(id)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS communication_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            article_id INTEGER NOT NULL,
            reviewer_id INTEGER,
            communication_date TEXT NOT NULL,
            recipient_type TEXT NOT NULL,
            recipient_name TEXT,
            communication_type TEXT NOT NULL,
            subject TEXT,
            notes TEXT,
            FOREIGN KEY(article_id) REFERENCES articles(id),
            FOREIGN KEY(reviewer_id) REFERENCES reviewers(id)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS contacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT,
            affiliation TEXT,
            role TEXT,
            UNIQUE(name, email)
        )
    """)

    conn.commit()
    conn.close()

def get_setting(key, default=""):
    conn = get_connection()
    row = conn.execute("SELECT setting_value FROM settings WHERE setting_key = ?", (key,)).fetchone()
    conn.close()
    return row["setting_value"] if row else default

def save_setting(key, value):
    conn = get_connection()
    conn.execute("INSERT INTO settings (setting_key, setting_value) VALUES (?, ?) ON CONFLICT(setting_key) DO UPDATE SET setting_value = excluded.setting_value", (key, value))
    conn.commit()
    conn.close()


# =========================================================
# HELPER KONTAK (AUTO-SUGGEST)
# =========================================================

def save_contact(name, email, affiliation, role):
    if not name: return
    conn = get_connection()
    try:
        conn.execute("""
            INSERT INTO contacts (name, email, affiliation, role)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(name, email) DO UPDATE SET
            affiliation = excluded.affiliation, role = excluded.role
        """, (name.strip(), email.strip() if email else "", affiliation.strip() if affiliation else "", role))
        conn.commit()
    except Exception: pass
    finally: conn.close()

def get_contacts(role=None):
    conn = get_connection()
    query = "SELECT * FROM contacts"
    params = ()
    if role:
        query += " WHERE role LIKE ?"
        params = (f"%{role}%",)
    query += " ORDER BY name ASC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


# =========================================================
# HELPER DATA
# =========================================================

def safe_date(value):
    if not value: return None
    try: return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()
    except ValueError: return None

def status_color(status):
    color_map = {
        "New Submission": "blue", "Desk Review": "blue", "Assign Reviewer": "orange",
        "Under Review": "orange", "Review Feedback Received": "violet", "Revision Requested": "orange",
        "Revision Submitted": "violet", "Accepted": "green", "APC Pending": "orange",
        "Payment Proof Submitted": "blue", "Payment Verified": "green", "Ready Copy Edit": "green",
        "Copy Editing": "green", "Layout Editing": "green", "Proofreading": "green",
        "Scheduled": "green", "Published": "green", "Rejected": "red", "Withdrawn": "red",
    }
    return color_map.get(status, "gray")

def get_articles_dataframe():
    conn = get_connection()
    df = pd.read_sql_query("""
        SELECT a.*, COUNT(r.id) AS reviewer_count
        FROM articles a
        LEFT JOIN reviewers r ON r.article_id = a.id
        GROUP BY a.id
        ORDER BY CASE WHEN a.status = 'Published' THEN 1 ELSE 0 END, a.updated_at DESC, a.id DESC
    """, conn)
    conn.close()
    return df

def get_article(article_id):
    conn = get_connection()
    row = conn.execute("SELECT * FROM articles WHERE id = ?", (article_id,)).fetchone()
    conn.close()
    return dict(row) if row else None

def get_article_id_by_code(article_code):
    conn = get_connection()
    row = conn.execute("SELECT id FROM articles WHERE article_code = ?", (article_code,)).fetchone()
    conn.close()
    return row["id"] if row else None

def get_reviewers(article_id=None):
    conn = get_connection()
    query = "SELECT r.*, a.article_code, a.journal_code, a.title, a.status AS article_status FROM reviewers r JOIN articles a ON a.id = r.article_id"
    params = []
    if article_id:
        query += " WHERE r.article_id = ?"
        params.append(article_id)
    query += " ORDER BY r.id DESC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(row) for row in rows]

def log_communication(article_id, recipient_type, recipient_name, communication_type, subject="", notes="", reviewer_id=None):
    conn = get_connection()
    conn.execute("""
        INSERT INTO communication_log (article_id, reviewer_id, communication_date, recipient_type, recipient_name, communication_type, subject, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (article_id, reviewer_id, date.today().isoformat(), recipient_type, recipient_name, communication_type, subject, notes))
    conn.commit()
    conn.close()

def create_article_folders(article_code, root_folder):
    if not root_folder: return None, "Folder utama belum diatur."
    root_folder = os.path.abspath(root_folder)
    article_folder = os.path.join(root_folder, article_code)
    try:
        os.makedirs(article_folder, exist_ok=True)
        for subfolder in FOLDER_STRUCTURE:
            os.makedirs(os.path.join(article_folder, subfolder), exist_ok=True)
        return article_folder, None
    except Exception as error:
        return None, str(error)


# =========================================================
# HELPER AUTO-FILL & EKSTRAKSI DOKUMEN
# =========================================================

def parse_metadata_from_filename(filename, upload_year):
    info = {}
    name_upper = filename.upper()
    name_lower = filename.lower()

    detected_journal = None
    for j_code in JOURNALS.keys():
        if j_code in name_upper:
            detected_journal = j_code
            info["journal_code"] = j_code
            break

    submission_id = ""
    if detected_journal:
        pattern = rf"{detected_journal}[^A-Z0-9]*(\d+)"
        match = re.search(pattern, name_upper)
        if match:
            submission_id = match.group(1)
            
    if not submission_id:
        fallback_match = re.search(r'\b\d{4,6}\b', filename)
        if fallback_match: submission_id = fallback_match.group()

    if submission_id: info["submission_id"] = submission_id

    if "desk review 1" in name_lower or "desk_review_1" in name_lower:
        info["status"] = "Desk Review"
        info["notes"] = "Urutan: Desk Review 1"
    elif "desk review 2" in name_lower or "desk_review_2" in name_lower:
        info["status"] = "Desk Review"
        info["notes"] = "Urutan: Desk Review 2"
    elif "desk review" in name_lower or "desk_review" in name_lower:
        info["status"] = "Desk Review"

    if detected_journal and submission_id:
        info["article_code"] = f"{detected_journal}-{upload_year}-{submission_id}"

    return info

def extract_title_from_document(uploaded_file):
    title = ""
    filename = uploaded_file.name.lower()
    file_bytes = uploaded_file.read()

    try:
        if filename.endswith(".docx") and docx:
            doc = docx.Document(BytesIO(file_bytes))
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    title = text
                    break
        elif filename.endswith(".pdf") and PyPDF2:
            reader = PyPDF2.PdfReader(BytesIO(file_bytes))
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    lines = [line.strip() for line in text.split('\n') if line.strip()]
                    if lines:
                        title = lines[0]
                        break
        elif filename.endswith(".txt"):
            content = file_bytes.decode("utf-8")
            lines = [line.strip() for line in content.split('\n') if line.strip()]
            if lines: title = lines[0]
    except Exception:
        pass 
    uploaded_file.seek(0)
    return title


# =========================================================
# HELPER UPLOAD DOKUMEN
# =========================================================

def get_default_storage_root():
    configured_root = get_setting("root_folder", "").strip()
    if configured_root: return os.path.abspath(configured_root)
    return os.path.abspath("Artikel_Jurnal")

def sanitize_filename(filename):
    return re.sub(r'[<>:"/\\|?*]+', "_", os.path.basename(filename))

def get_unique_file_path(destination_folder, filename):
    filename = sanitize_filename(filename)
    name, extension = os.path.splitext(filename)
    target_path = os.path.join(destination_folder, filename)
    counter = 1
    while os.path.exists(target_path):
        target_path = os.path.join(destination_folder, f"{name}__{counter:02d}{extension}")
        counter += 1
    return target_path

def detect_document_subfolder(filename):
    filename_lower = filename.lower()
    for folder_name, keywords in UPLOAD_FOLDER_RULES.items():
        if any(keyword.lower() in filename_lower for keyword in keywords):
            return folder_name
    return "01_Submission"

def save_uploaded_files(uploaded_files, article_folder):
    saved_files, errors = [], []
    if not uploaded_files: return saved_files, errors
    for uploaded_file in uploaded_files:
        try:
            original_name = uploaded_file.name
            extension = os.path.splitext(original_name)[1].lower().replace(".", "")
            if extension not in ALLOWED_UPLOAD_EXTENSIONS:
                errors.append(f"{original_name}: format file tidak didukung.")
                continue
            target_subfolder = detect_document_subfolder(original_name)
            destination_folder = os.path.join(article_folder, target_subfolder)
            os.makedirs(destination_folder, exist_ok=True)
            target_path = get_unique_file_path(destination_folder, original_name)
            
            with open(target_path, "wb") as file:
                file.write(uploaded_file.getbuffer())
                
            saved_files.append(os.path.relpath(target_path, article_folder))
        except Exception as error:
            errors.append(f"{uploaded_file.name}: {error}")
    return saved_files, errors


# =========================================================
# EXPORT EXCEL
# =========================================================

def build_excel_export():
    conn = get_connection()
    query = """
        SELECT
            a.article_code AS "Nomor", a.journal_code AS "Jurnal", a.submission_id AS "Submission ID",
            a.corresponding_author AS "Penulis", a.author_email AS "Email Penulis",
            a.affiliation_1 AS "Afiliasi 1", a.affiliation_2 AS "Afiliasi 2",
            a.title AS "Judul", a.status AS "Status Terbit", r.reviewer_name AS "Reviewer",
            r.reviewer_email AS "Email Reviewer", r.reviewer_category AS "Status reviewer",
            r.status AS "Status Proses Reviewer", r.invitation_date AS "Mulai Undangan",
            r.invitation_deadline AS "Batas Respons", r.acceptance_date AS "Reviewer Accept",
            r.review_deadline AS "Selesai Review", r.review_submitted_date AS "Hasil Review Masuk",
            r.recommendation AS "Rekomendasi Reviewer", r.reminder_count AS "Jumlah Reminder Reviewer",
            a.author_revision_deadline AS "Selesai Revisi Author", a.author_reminder_count AS "Jumlah Reminder Author",
            a.apc_deadline AS "Batas APC", a.payment_proof_date AS "Bukti Bayar Masuk",
            a.payment_verified_date AS "Pembayaran Terverifikasi", a.volume AS "Volume",
            a.issue AS "Issue", a.publication_year AS "Tahun", a.folder_path AS "Folder Artikel",
            a.notes AS "Catatan",
            CASE WHEN a.is_international = 1 THEN 'Ya' ELSE 'Tidak' END AS "Afiliasi Luar Negeri"
        FROM articles a
        LEFT JOIN reviewers r ON r.article_id = a.id
        ORDER BY a.journal_code, a.article_code, r.id
    """
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def dataframe_to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Monitoring All")
        worksheet = writer.book["Monitoring All"]
        for column_cells in worksheet.columns:
            max_length = max((len(str(cell.value or "")) for cell in column_cells), default=0)
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(max_length + 2, 45)
        worksheet.freeze_panes = "A2"
    return output.getvalue()


# =========================================================
# DASHBOARD
# =========================================================

def get_dashboard_tasks():
    today = date.today()
    tasks = []
    conn = get_connection()
    
    reviewer_rows = conn.execute("""
        SELECT r.*, a.article_code, a.journal_code, a.title, a.status AS article_status
        FROM reviewers r JOIN articles a ON a.id = r.article_id
        WHERE r.status IN ('Invited', 'Accepted', 'Unresponsive')
    """).fetchall()

    for row in reviewer_rows:
        reviewer = dict(row)
        if reviewer["status"] == "Invited":
            deadline = safe_date(reviewer["invitation_deadline"])
            if deadline and deadline <= today:
                days_late = (today - deadline).days
                if reviewer["reminder_count"] >= 2:
                    action, priority = "Maksimum reminder tercapai — cari reviewer pengganti.", "Kritis"
                else:
                    action, priority = "Kirim reminder respons undangan reviewer.", "Terlambat" if days_late > 0 else "Hari ini"
                tasks.append({"Prioritas": priority, "Jurnal": reviewer["journal_code"], "Artikel": reviewer["article_code"], "Judul": reviewer["title"], "Penerima": reviewer["reviewer_name"], "Tindakan": action, "Deadline": deadline.isoformat(), "Jenis": "Reviewer Invitation"})
        elif reviewer["status"] == "Accepted":
            deadline = safe_date(reviewer["review_deadline"])
            if deadline and deadline <= today:
                days_late = (today - deadline).days
                if reviewer["reminder_count"] >= 2:
                    action, priority = "Maksimum reminder tercapai — evaluasi atau ganti reviewer.", "Kritis"
                else:
                    action, priority = "Kirim reminder hasil review.", "Terlambat" if days_late > 0 else "Hari ini"
                tasks.append({"Prioritas": priority, "Jurnal": reviewer["journal_code"], "Artikel": reviewer["article_code"], "Judul": reviewer["title"], "Penerima": reviewer["reviewer_name"], "Tindakan": action, "Deadline": deadline.isoformat(), "Jenis": "Reviewer Deadline"})

    author_rows = conn.execute("SELECT * FROM articles WHERE status = 'Revision Requested' AND author_revision_deadline IS NOT NULL").fetchall()
    for row in author_rows:
        article = dict(row)
        deadline = safe_date(article["author_revision_deadline"])
        if deadline and deadline <= today:
            if article["author_reminder_count"] >= 3:
                action, priority = "Final reminder expired — Withdraw atau Extend Deadline.", "Kritis"
            else:
                action, priority = "Kirim reminder revisi kepada author.", "Terlambat" if deadline < today else "Hari ini"
            tasks.append({"Prioritas": priority, "Jurnal": article["journal_code"], "Artikel": article["article_code"], "Judul": article["title"], "Penerima": article["corresponding_author"], "Tindakan": action, "Deadline": deadline.isoformat(), "Jenis": "Author Revision"})

    apc_rows = conn.execute("SELECT * FROM articles WHERE status = 'APC Pending' AND apc_deadline IS NOT NULL").fetchall()
    for row in apc_rows:
        article = dict(row)
        deadline = safe_date(article["apc_deadline"])
        if deadline and deadline <= today:
            tasks.append({"Prioritas": "Terlambat" if deadline < today else "Hari ini", "Jurnal": article["journal_code"], "Artikel": article["article_code"], "Judul": article["title"], "Penerima": article["corresponding_author"], "Tindakan": "Follow-up pembayaran APC.", "Deadline": deadline.isoformat(), "Jenis": "APC"})
            
    conn.close()
    priority_order = {"Kritis": 0, "Terlambat": 1, "Hari ini": 2}
    tasks.sort(key=lambda item: (priority_order.get(item["Prioritas"], 9), item["Deadline"]))
    return tasks

def page_dashboard():
    st.subheader("📊 Dashboard Tindakan & Jadwal")
    articles = get_articles_dataframe()
    tasks = get_dashboard_tasks()
    active_articles = len(articles[~articles["status"].isin(["Published", "Rejected", "Withdrawn"])]) if not articles.empty else 0
        
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Artikel Aktif", active_articles)
    col2.metric("Total Artikel", len(articles))
    col3.metric("Tindakan Diproses", len(tasks))
    col4.metric("Tindakan Kritis", len([t for t in tasks if t["Prioritas"] == "Kritis"]))

    st.markdown("---")
    st.markdown("### 📅 Jadwal Kerja Bulanan & Harian")
    tab_today, tab_month, tab_weekly = st.tabs(["📌 Jadwal Hari Ini", "🗓️ Kalender Bulan Ini", "📋 Jadwal Mingguan Penuh"])

    with tab_today:
        weekday = datetime.today().weekday()
        if weekday <= 4:
            focus = WEEKLY_ROUTINE[weekday]
            st.info(f"**Fokus Hari Ini ({focus['Hari']}):** Jurnal **{focus['Jurnal']}**")
            
            st.markdown("#### 💡 Rekomendasi Tindakan Cepat (Sesuai Fokus Hari Ini)")
            target_journals = [focus['Jurnal']] if focus['Jurnal'] != "Semua jurnal" else list(JOURNALS.keys())
            action_needed_statuses = ["New Submission", "Desk Review", "Review Feedback Received", "Revision Submitted"]
            
            if not articles.empty:
                recommended_df = articles[
                    (articles["journal_code"].isin(target_journals)) & 
                    (articles["status"].isin(action_needed_statuses))
                ]
                
                if not recommended_df.empty:
                    st.write(f"Berikut adalah **{len(recommended_df)} artikel** yang sedang menunggu tindakan Anda (Desk Review, Cek Hasil Review, atau Evaluasi Revisi):")
                    display_recom = recommended_df[["article_code", "title", "status", "updated_at"]].copy()
                    display_recom.rename(columns={"article_code": "Kode Artikel", "title": "Judul", "status": "Status Menunggu", "updated_at": "Terakhir Update"}, inplace=True)
                    st.dataframe(display_recom, use_container_width=True, hide_index=True)
                else:
                    st.success(f"Kerja bagus! Saat ini tidak ada antrean artikel dari jurnal **{focus['Jurnal']}** yang menumpuk di meja Anda.")
            else:
                st.info("Belum ada artikel di dalam database.")

            st.markdown("#### 📋 Agenda Rutin")
            today_schedule = []
            for item in DAILY_SCHEDULE_TEMPLATE:
                kegiatan = item["Kegiatan Utama"].format(tugas_pagi=focus["09.30–11.30"], tugas_siang=focus["13.00–15.00"], tugas_sore=focus["15.30–16.15"])
                today_schedule.append({"Waktu": item["Waktu"], "Kegiatan Utama": kegiatan, "Fokus Email": item["Fokus Email"]})
            st.dataframe(pd.DataFrame(today_schedule), use_container_width=True, hide_index=True)
        else:
            st.success("🎉 Selamat berakhir pekan! Tidak ada jadwal rutin editorial hari ini.")

    with tab_month:
        now = datetime.now()
        cal = calendar.monthcalendar(now.year, now.month)
        bulan_indo = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
        st.markdown(f"**Jadwal Alokasi Jurnal Bulan {bulan_indo[now.month]} {now.year}**")
        df_cal = pd.DataFrame(cal, columns=["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]).replace(0, "")
        for idx, row in df_cal.iterrows():
            if row["Senin"] != "": df_cal.at[idx, "Senin"] = f"{row['Senin']} (HM)"
            if row["Selasa"] != "": df_cal.at[idx, "Selasa"] = f"{row['Selasa']} (CT)"
            if row["Rabu"] != "": df_cal.at[idx, "Rabu"] = f"{row['Rabu']} (LC)"
            if row["Kamis"] != "": df_cal.at[idx, "Kamis"] = f"{row['Kamis']} (All)"
            if row["Jumat"] != "": df_cal.at[idx, "Jumat"] = f"{row['Jumat']} (All)"
        st.dataframe(df_cal, use_container_width=True, hide_index=True)

    with tab_weekly:
        st.dataframe(pd.DataFrame.from_dict(WEEKLY_ROUTINE, orient="index"), use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 🚨 Daftar Tindakan Deadline & Reminder")
    if not tasks:
        st.success("Tidak ada deadline atau reminder yang perlu diproses hari ini.")
    else:
        st.dataframe(pd.DataFrame(tasks), use_container_width=True, hide_index=True)


# =========================================================
# HALAMAN TAMBAH ARTIKEL
# =========================================================

def page_add_article():
    st.subheader("📝 Tambah Artikel Baru")
    active_root_folder = get_default_storage_root()
    st.caption(f"Folder penyimpanan aktif: `{active_root_folder}`")

    st.info("Unggah file di sini *sebelum* mengetik untuk mengisi informasi secara otomatis berdasarkan nama file dan baris pertama di dalam dokumen.")
    
    if "auto_uploader_key" not in st.session_state: st.session_state["auto_uploader_key"] = 0
    auto_uploaded_files = st.file_uploader("Pilih satu/beberapa dokumen", type=ALLOWED_UPLOAD_EXTENSIONS, accept_multiple_files=True, key=f"auto_uploader_{st.session_state['auto_uploader_key']}")

    def_journal, def_sub_id, def_status, def_title, def_notes, def_article_code = list(JOURNALS.keys())[0], "", "New Submission", "", "", ""

    if auto_uploaded_files:
        first_file = auto_uploaded_files[0]
        meta_info = parse_metadata_from_filename(first_file.name, str(date.today().year))
        def_journal = meta_info.get("journal_code", def_journal)
        def_sub_id = meta_info.get("submission_id", def_sub_id)
        def_status = meta_info.get("status", def_status)
        def_notes = meta_info.get("notes", def_notes)
        def_article_code = meta_info.get("article_code", def_article_code)
        extracted_title = extract_title_from_document(first_file)
        if extracted_title: def_title = extracted_title
        st.success(f"File **{first_file.name}** berhasil dibaca.")

    try: journal_index = list(JOURNALS.keys()).index(def_journal)
    except ValueError: journal_index = 0
    try: status_index = ARTICLE_STATUSES.index(def_status)
    except ValueError: status_index = 0

    st.markdown("#### Informasi Artikel")
    
    authors_db = get_contacts("Author")
    author_opts = ["--- Input Manual ---"] + [f"{c['name']} | {c['email']}" for c in authors_db]
    selected_auth_opt = st.selectbox("💡 Pilih dari History Author Tersimpan (Opsional)", author_opts)

    with st.form("form_add_article", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            journal_code = st.selectbox("Jurnal", list(JOURNALS.keys()), index=journal_index)
            article_code = st.text_input("Kode Artikel *", value=def_article_code)
            submission_id = st.text_input("Submission ID OJS", value=def_sub_id)
            
        with col2:
            def_auth_name, def_auth_email, def_affil = "", "", ""
            if selected_auth_opt != "--- Input Manual ---":
                matched = next(c for c in authors_db if f"{c['name']} | {c['email']}" == selected_auth_opt)
                def_auth_name, def_auth_email, def_affil = matched["name"], matched["email"], matched["affiliation"]
                
            corresponding_author = st.text_input("Corresponding Author", value=def_auth_name)
            author_email = st.text_input("Email Author", value=def_auth_email)
            submission_date = st.date_input("Tanggal Submit", value=date.today())
            
        with col3:
            affiliation_1 = st.text_input("Afiliasi 1", value=def_affil)
            affiliation_2 = st.text_input("Afiliasi 2")
            status = st.selectbox("Status Awal", ARTICLE_STATUSES, index=status_index)

        title = st.text_area("Judul Artikel *", value=def_title, height=100)
        
        is_international = st.checkbox("🌍 Terdapat Afiliasi Luar Negeri (Internasional)")
        
        notes = st.text_area("Catatan Internal", value=def_notes, height=80)
        create_folder = st.checkbox("Buat struktur folder artikel otomatis", value=True)
        submitted = st.form_submit_button("Simpan Artikel", type="primary")

    if submitted:
        if not article_code.strip() or not title.strip():
            st.error("Kode artikel dan judul artikel wajib diisi.")
            return

        clean_article_code = article_code.strip()
        folder_path = ""
        if create_folder or bool(auto_uploaded_files):
            folder_path, error = create_article_folders(clean_article_code, active_root_folder)
            if error:
                st.error(f"Gagal membuat folder: {error}")
                return

        try:
            conn = get_connection()
            conn.execute("""
                INSERT INTO articles (
                    article_code, journal_code, submission_id, title, corresponding_author, author_email,
                    affiliation_1, affiliation_2, status, submission_date, folder_path, notes, is_international, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                clean_article_code, journal_code, submission_id.strip(), title.strip(), corresponding_author.strip(),
                author_email.strip(), affiliation_1.strip(), affiliation_2.strip(), status, submission_date.isoformat(),
                folder_path, notes.strip(), 1 if is_international else 0, datetime.now().isoformat(timespec="seconds")
            ))
            conn.commit()
            conn.close()
            article_id = get_article_id_by_code(clean_article_code)

            if auto_uploaded_files and folder_path:
                saved_files, upload_errors = save_uploaded_files(auto_uploaded_files, folder_path)
                if article_id:
                    log_communication(article_id, "Internal", "Editorial Team", "Article Created", notes=f"{len(saved_files)} dokumen diunggah.")

            save_contact(corresponding_author, author_email, affiliation_1, "Author")

            st.success(f"Artikel {clean_article_code} berhasil disimpan.")
            st.session_state["auto_uploader_key"] += 1
            st.rerun()
        except sqlite3.IntegrityError:
            st.error("Kode artikel sudah digunakan.")


# =========================================================
# HALAMAN KELOLA ARTIKEL
# =========================================================

def highlight_journal_row(row):
    jurnal = row['Jurnal']
    if jurnal == 'CT': return ['background-color: rgba(13, 110, 253, 0.15)'] * len(row)
    elif jurnal == 'HM': return ['background-color: rgba(25, 135, 84, 0.15)'] * len(row)
    elif jurnal == 'LC': return ['background-color: rgba(255, 193, 7, 0.25)'] * len(row)
    return [''] * len(row)

def page_manage_articles():
    st.subheader("⚙️ Kelola Artikel")
    articles = get_articles_dataframe()
    if articles.empty:
        st.info("Belum ada artikel. Tambahkan artikel terlebih dahulu.")
        return

    search = st.text_input("Cari kode artikel, judul, author, atau jurnal", placeholder="Contoh: HM-2026-12345 atau nama author")
    display_df = articles.copy()
    
    if search:
        search_lower = search.lower()
        mask = (display_df["article_code"].fillna("").str.lower().str.contains(search_lower) |
                display_df["title"].fillna("").str.lower().str.contains(search_lower) |
                display_df["corresponding_author"].fillna("").str.lower().str.contains(search_lower) |
                display_df["journal_code"].fillna("").str.lower().str.contains(search_lower))
        display_df = display_df[mask]

    if display_df.empty:
        st.warning("Artikel tidak ditemukan.")
        return

    display_df['is_international'] = display_df['is_international'].fillna(0).astype(int)
    display_df['article_code_display'] = display_df.apply(lambda x: f"{x['article_code']} 🌍" if x['is_international'] else x['article_code'], axis=1)

    display_columns = ["article_code_display", "journal_code", "title", "corresponding_author", "status", "reviewer_count", "author_revision_deadline"]
    renamed_df = display_df[display_columns].rename(columns={
        "article_code_display": "Kode Artikel", "journal_code": "Jurnal", "title": "Judul", "corresponding_author": "Author",
        "status": "Status", "reviewer_count": "Reviewer", "author_revision_deadline": "Deadline Revisi"
    })
    
    styled_df = renamed_df.style.apply(highlight_journal_row, axis=1)
    st.dataframe(styled_df, use_container_width=True, hide_index=True)

    selected_code = st.selectbox("Pilih artikel (tanpa emoji) untuk dikelola", display_df["article_code"].tolist())
    selected_row = articles[articles["article_code"] == selected_code].iloc[0]
    article = get_article(int(selected_row["id"]))

    international_marker = " 🌍 (Afiliasi Luar Negeri)" if article.get("is_international", 0) else ""
    st.markdown("---")
    st.markdown(f"### {article['article_code']}{international_marker} — {article['journal_code']}")
    st.write(article["title"])

    tabs = st.tabs(["Status & Metadata", "Reviewer", "Revisi / APC", "Folder & Checklist", "Riwayat Komunikasi"])

    # --- TAB STATUS & METADATA ---
    with tabs[0]:
        current_index = ARTICLE_STATUSES.index(article["status"]) if article["status"] in ARTICLE_STATUSES else 0
        with st.form(f"form_article_status_{article['id']}"):
            status = st.selectbox("Status Artikel", ARTICLE_STATUSES, index=current_index)
            is_international = st.checkbox("Terdapat Afiliasi Luar Negeri (Internasional) 🌍", value=bool(article.get("is_international", 0)))
            
            col1, col2, col3 = st.columns(3)
            with col1: volume = st.text_input("Volume", value=article["volume"] or "")
            with col2: issue = st.text_input("Issue", value=article["issue"] or "")
            with col3: publication_year = st.text_input("Tahun", value=article["publication_year"] or "")
            
            notes = st.text_area("Catatan Internal", value=article["notes"] or "", height=120)
            if st.form_submit_button("Simpan Perubahan", type="primary"):
                conn = get_connection()
                conn.execute("""
                    UPDATE articles
                    SET status=?, volume=?, issue=?, publication_year=?, notes=?, is_international=?, updated_at=?
                    WHERE id=?
                """, (status, volume.strip(), issue.strip(), publication_year.strip(), notes.strip(), 1 if is_international else 0, datetime.now().isoformat(timespec="seconds"), article["id"]))
                conn.commit()
                conn.close()
                st.success("Perubahan artikel berhasil disimpan.")
                st.rerun()

    # --- TAB REVIEWER ---
    with tabs[1]:
        reviewers = get_reviewers(article["id"])
        st.markdown("#### Reviewer Terdaftar")
        if reviewers:
            r_df = pd.DataFrame(reviewers)
            st.dataframe(r_df[["reviewer_name", "reviewer_email", "status", "invitation_deadline", "review_deadline", "recommendation", "reminder_count"]].rename(columns={"reviewer_name": "Nama", "reviewer_email": "Email", "status": "Status", "invitation_deadline": "Batas Respons", "review_deadline": "Deadline Review", "recommendation": "Rekomendasi", "reminder_count": "Reminder"}), use_container_width=True, hide_index=True)
        else:
            st.warning("Belum ada reviewer untuk artikel ini.")

        st.markdown("#### Tambah Reviewer")
        
        revs_db = get_contacts("Reviewer")
        rev_opts = ["--- Input Manual ---"] + [f"{c['name']} | {c['email']}" for c in revs_db]
        selected_rev_opt = st.selectbox("💡 Pilih dari History Reviewer (Opsional)", rev_opts)
        
        with st.form(f"form_add_reviewer_{article['id']}", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                def_rev_name, def_rev_email = "", ""
                if selected_rev_opt != "--- Input Manual ---":
                    matched_r = next(c for c in revs_db if f"{c['name']} | {c['email']}" == selected_rev_opt)
                    def_rev_name, def_rev_email = matched_r["name"], matched_r["email"]

                r_name = st.text_input("Nama Reviewer", value=def_rev_name)
                r_email = st.text_input("Email Reviewer", value=def_rev_email)
                r_cat = st.selectbox("Status reviewer untuk Excel", REVIEWER_CATEGORIES)
            with c2:
                r_date = st.date_input("Tanggal Undangan", value=date.today())
                r_notes = st.text_area("Catatan Reviewer", height=100)
            
            if st.form_submit_button("Tambah Reviewer"):
                if not r_name.strip(): st.error("Nama reviewer wajib diisi.")
                else:
                    r_dead = r_date + timedelta(days=7)
                    conn = get_connection()
                    conn.execute("INSERT INTO reviewers (article_id, reviewer_name, reviewer_email, reviewer_category, invitation_date, invitation_deadline, status, notes) VALUES (?, ?, ?, ?, ?, ?, 'Invited', ?)", (article["id"], r_name.strip(), r_email.strip(), r_cat, r_date.isoformat(), r_dead.isoformat(), r_notes.strip()))
                    conn.execute("UPDATE articles SET status = 'Under Review' WHERE id = ?", (article["id"],))
                    conn.commit()
                    conn.close()
                    
                    save_contact(r_name, r_email, "", "Reviewer")
                    log_communication(article["id"], "Reviewer", r_name.strip(), "Reviewer Invitation", notes="Reviewer ditambahkan ke aplikasi.")
                    st.success("Reviewer ditambahkan!")
                    st.rerun()

        if reviewers:
            st.markdown("#### Perbarui Status Reviewer")
            r_opts = {f"{r['reviewer_name']} — {r['status']}": r for r in reviewers}
            selected_r = r_opts[st.selectbox("Pilih reviewer", list(r_opts.keys()))]
            
            with st.form(f"form_update_reviewer_{selected_r['id']}"):
                c1, c2 = st.columns(2)
                with c1:
                    r_stat = st.selectbox("Status", REVIEWER_STATUSES, index=REVIEWER_STATUSES.index(selected_r["status"]) if selected_r["status"] in REVIEWER_STATUSES else 0)
                    r_rec = st.selectbox("Rekomendasi", REVIEW_RECOMMENDATIONS, index=REVIEW_RECOMMENDATIONS.index(selected_r["recommendation"] or "") if (selected_r["recommendation"] or "") in REVIEW_RECOMMENDATIONS else 0)
                with c2:
                    acc_date = st.date_input("Tanggal Accept", value=safe_date(selected_r["acceptance_date"]) or date.today())
                    sub_date = st.date_input("Tanggal Hasil Masuk", value=safe_date(selected_r["review_submitted_date"]) or date.today())
                r_up_notes = st.text_area("Catatan", value=selected_r["notes"] or "")
                if st.form_submit_button("Simpan Status"):
                    r_dead, acc_db, sub_db, a_stat = selected_r["review_deadline"], selected_r["acceptance_date"], selected_r["review_submitted_date"], None
                    if r_stat == "Accepted":
                        acc_db = acc_date.isoformat()
                        r_dead = (acc_date + timedelta(days=14)).isoformat()
                    elif r_stat == "Review Submitted":
                        sub_db = sub_date.isoformat()
                        a_stat = "Review Feedback Received"
                    conn = get_connection()
                    conn.execute("UPDATE reviewers SET status=?, recommendation=?, acceptance_date=?, review_deadline=?, review_submitted_date=?, notes=? WHERE id=?", (r_stat, r_rec, acc_db, r_dead, sub_db, r_up_notes.strip(), selected_r["id"]))
                    if a_stat: conn.execute("UPDATE articles SET status=? WHERE id=?", (a_stat, article["id"]))
                    conn.commit()
                    conn.close()
                    st.success("Tersimpan!")
                    st.rerun()

            with st.expander("🔔 Kirim & Catat Reminder Reviewer"):
                st.write(f"Reminder reviewer terkirim: **{selected_r['reminder_count']} / 2**")
                new_rev_deadline = st.date_input("Set Deadline Baru (Opsional)", value=safe_date(selected_r["review_deadline"]) or date.today(), key=f"dead_rev_{selected_r['id']}")
                if st.button("Catat Reminder & Perbarui Deadline", key=f"btn_rem_rev_{selected_r['id']}", type="primary"):
                    if selected_r["reminder_count"] >= 2:
                        st.error("Maksimum 2 reminder tercapai. Pertimbangkan mencari pengganti.")
                    else:
                        new_cnt = selected_r["reminder_count"] + 1
                        conn = get_connection()
                        conn.execute("UPDATE reviewers SET reminder_count=?, review_deadline=? WHERE id=?", (new_cnt, new_rev_deadline.isoformat(), selected_r["id"]))
                        conn.commit()
                        conn.close()
                        log_communication(article["id"], "Reviewer", selected_r["reviewer_name"], "Reviewer Reminder", notes=f"Reminder ke-{new_cnt}. Deadline update: {new_rev_deadline}")
                        st.success(f"Reminder dicatat, deadline diubah menjadi {new_rev_deadline}.")
                        st.rerun()

    # --- TAB REVISI DAN APC ---
    with tabs[2]:
        st.markdown("#### Revisi Author")
        cur_rev_dead = safe_date(article["author_revision_deadline"])
        with st.form(f"form_revision_{article['id']}"):
            set_dead = st.checkbox("Atur deadline revisi author", value=cur_rev_dead is not None)
            rev_dead = st.date_input("Deadline Revisi Author", value=cur_rev_dead or (date.today() + timedelta(days=14)))
            if st.form_submit_button("Simpan Deadline Revisi"):
                val = rev_dead.isoformat() if set_dead else None
                conn = get_connection()
                conn.execute("UPDATE articles SET author_revision_deadline=?, status=CASE WHEN ? IS NOT NULL THEN 'Revision Requested' ELSE status END WHERE id=?", (val, val, article["id"]))
                conn.commit()
                conn.close()
                st.success("Tersimpan!")
                st.rerun()

        with st.expander("🔔 Kirim & Catat Reminder Author"):
            st.write(f"Reminder revisi terkirim: **{article['author_reminder_count']} / 3**")
            new_auth_deadline = st.date_input("Set Deadline Baru (Opsional)", value=cur_rev_dead or date.today(), key=f"dead_auth_{article['id']}")
            if st.button("Catat Reminder & Perbarui Deadline", key=f"btn_rem_auth_{article['id']}", type="primary"):
                if article["author_reminder_count"] >= 3:
                    st.error("Maksimum reminder tercapai. Pilih Withdraw atau Extend Deadline.")
                else:
                    new_cnt = article["author_reminder_count"] + 1
                    conn = get_connection()
                    conn.execute("UPDATE articles SET author_reminder_count=?, author_revision_deadline=? WHERE id=?", (new_cnt, new_auth_deadline.isoformat(), article["id"]))
                    conn.commit()
                    conn.close()
                    log_communication(article["id"], "Author", article["corresponding_author"], "Author Revision Reminder", notes=f"Reminder ke-{new_cnt}. Deadline update: {new_auth_deadline}")
                    st.success(f"Reminder dicatat, deadline diubah menjadi {new_auth_deadline}.")
                    st.rerun()
            
            if st.button("Withdraw Artikel", key=f"withdraw_{article['id']}"):
                conn = get_connection()
                conn.execute("UPDATE articles SET status = 'Withdrawn' WHERE id = ?", (article["id"],))
                conn.commit()
                conn.close()
                log_communication(article["id"], "Author", article["corresponding_author"], "Withdrawal Decision", notes="Manual Withdraw.")
                st.rerun()

        st.markdown("---")
        st.markdown("#### APC")
        with st.form(f"form_apc_{article['id']}"):
            c1, c2, c3 = st.columns(3)
            with c1: apc_dead = st.date_input("Deadline APC", value=safe_date(article["apc_deadline"]) or (date.today()+timedelta(days=14)))
            with c2: prf_date = st.date_input("Tgl Bukti Bayar", value=safe_date(article["payment_proof_date"]) or date.today())
            with c3: vrf_date = st.date_input("Tgl Verifikasi", value=safe_date(article["payment_verified_date"]) or date.today())
            apc_stat = st.selectbox("Status APC", ["APC Pending", "Payment Proof Submitted", "Payment Verified", "Ready Copy Edit"])
            if st.form_submit_button("Simpan Status APC"):
                conn = get_connection()
                conn.execute("UPDATE articles SET status=?, apc_deadline=?, payment_proof_date=?, payment_verified_date=? WHERE id=?", (apc_stat, apc_dead.isoformat(), prf_date.isoformat() if "Proof" in apc_stat or "Veri" in apc_stat or "Copy" in apc_stat else None, vrf_date.isoformat() if "Veri" in apc_stat or "Copy" in apc_stat else None, article["id"]))
                conn.commit()
                conn.close()
                st.success("Tersimpan!")
                st.rerun()

    # --- TAB FOLDER DAN CHECKLIST ---
    with tabs[3]:
        folder_path = article["folder_path"]
        if not folder_path:
            st.warning("Folder belum dibuat.")
            if st.button("Buat Struktur Folder"):
                new_path, err = create_article_folders(article["article_code"], get_default_storage_root())
                if err: st.error(err)
                else:
                    conn = get_connection()
                    conn.execute("UPDATE articles SET folder_path=? WHERE id=?", (new_path, article["id"]))
                    conn.commit()
                    conn.close()
                    st.rerun()
        else:
            st.write(f"`{folder_path}`")
            up_files = st.file_uploader("Upload dokumen tambahan", type=ALLOWED_UPLOAD_EXTENSIONS, accept_multiple_files=True)
            if st.button("Simpan Dokumen") and up_files:
                svd, errs = save_uploaded_files(up_files, folder_path)
                if svd: st.success(f"{len(svd)} disimpan.")
                st.rerun()

            all_files = []
            if os.path.exists(folder_path):
                for r, d, f in os.walk(folder_path):
                    for fn in f: all_files.append(os.path.relpath(os.path.join(r, fn), folder_path))
            c_rows = []
            for d_name, kwds in DOCUMENT_CHECKLIST.items():
                fnd = [fn for fn in all_files if any(k.lower() in fn.lower() for k in kwds)]
                c_rows.append({"Dokumen": d_name, "Status": "Tersedia" if fnd else "Belum ditemukan", "File Terdeteksi": ", ".join(fnd) if fnd else "-"})
            st.dataframe(pd.DataFrame(c_rows), use_container_width=True, hide_index=True)

    # --- TAB RIWAYAT KOMUNIKASI ---
    with tabs[4]:
        conn = get_connection()
        logs = pd.read_sql_query("SELECT communication_date, recipient_type, recipient_name, communication_type, subject, notes FROM communication_log WHERE article_id = ? ORDER BY communication_date DESC, id DESC", conn, params=(article["id"],))
        conn.close()
        if not logs.empty: st.dataframe(logs, use_container_width=True, hide_index=True)
        else: st.info("Belum ada komunikasi.")


# =========================================================
# HALAMAN IMPOR DOKUMEN & EMAIL & EKSPOR
# =========================================================

def page_import_documents():
    st.subheader("📂 Impor Dokumen Artikel")
    arts = get_articles_dataframe()
    if arts.empty: return st.info("Belum ada artikel.")
    opts = {f"{r['article_code']} — {r['title']}": int(r["id"]) for _, r in arts.iterrows()}
    art = get_article(opts[st.selectbox("Pilih Artikel", list(opts.keys()))])
    if not art["folder_path"]: return st.warning("Folder belum dibuat.")
    
    files = st.file_uploader("Upload dokumen", type=ALLOWED_UPLOAD_EXTENSIONS, accept_multiple_files=True)
    if st.button("Simpan ke Folder", type="primary") and files:
        svd, errs = save_uploaded_files(files, art["folder_path"])
        if svd: st.success(f"{len(svd)} file tersimpan.")

def page_email_drafts():
    st.subheader("📧 Generator Draft Email Copy-Paste")
    st.markdown("Kamu bisa menyesuaikan isi text di bawah sebelum menyalinnya menggunakan tombol khusus yang ada di kotak abu-abu.")
    
    arts = get_articles_dataframe()
    if arts.empty: return st.info("Belum ada artikel yang tersedia.")
    
    art = get_article(int(arts[arts["article_code"] == st.selectbox("Artikel", arts["article_code"])].iloc[0]["id"]))
    
    rmode = st.radio("Penerima", ["Author", "Reviewer"])
    rname, remail = art["corresponding_author"], art["author_email"]
    
    if rmode == "Reviewer":
        revs = get_reviewers(art["id"])
        if not revs: return st.warning("Artikel ini belum memiliki reviewer ditugaskan.")
        ropts = {f"{r['reviewer_name']}": r for r in revs}
        sel_r = ropts[st.selectbox("Pilih Reviewer Terdaftar", list(ropts.keys()))]
        rname, remail = sel_r["reviewer_name"], sel_r["reviewer_email"]

    ttype = st.selectbox("Pilih Template", ["Reviewer Invitation", "Reviewer Reminder", "Author Revision Reminder", "APC Reminder"])
    
    # Text Templates
    if ttype == "Reviewer Invitation":
        default_subject = f"Invitation to Review — {art['article_code']}"
        default_body = f"Dear {rname},\n\nWe would like to invite you to review the manuscript below.\n\nArticle ID: {art['article_code']}\nTitle: {art['title']}\n\nPlease confirm whether you are available to review this manuscript.\n\nThank you for your consideration.\n\nBest regards,\nEditorial Team\n{art['journal_code']}"
    elif ttype == "Reviewer Reminder":
        default_subject = f"Review Reminder — {art['article_code']}"
        default_body = f"Dear {rname},\n\nThis is a friendly reminder regarding the review of the following manuscript.\n\nArticle ID: {art['article_code']}\nTitle: {art['title']}\n\nWe would greatly appreciate it if you could submit your review at your earliest convenience.\n\nThank you for your valuable contribution.\n\nBest regards,\nEditorial Team\n{art['journal_code']}"
    elif ttype == "Author Revision Reminder":
        default_subject = f"Revision Reminder — {art['article_code']}"
        default_body = f"Dear {rname},\n\nThis is a reminder regarding the revision of your manuscript.\n\nArticle ID: {art['article_code']}\nTitle: {art['title']}\n\nPlease submit your revised manuscript and response to reviewers through OJS before the assigned deadline.\n\nBest regards,\nEditorial Team\n{art['journal_code']}"
    elif ttype == "APC Reminder":
        default_subject = f"APC Payment Reminder — {art['article_code']}"
        default_body = f"Dear {rname},\n\nThis is a reminder regarding the Article Processing Charge (APC) for the following accepted manuscript.\n\nArticle ID: {art['article_code']}\nTitle: {art['title']}\n\nPlease submit the payment proof according to the editorial instruction.\n\nBest regards,\nEditorial Team\n{art['journal_code']}"

    # Form Editor Draft
    st.markdown("#### 1. Sesuaikan Isi Email")
    subject_edit = st.text_input("Subjek", value=default_subject)
    body_edit = st.text_area("Isi Email", value=default_body, height=250)
    
    st.markdown("#### 2. Copy Hasil Jadi & Catat Riwayat")
    st.info("Arahkan mouse ke sudut kanan atas kotak di bawah ini untuk memunculkan tombol 'Copy to clipboard'.")
    st.code(f"To: {remail}\nSubject: {subject_edit}\n\n{body_edit}", language="text")

    if st.button("Simpan ke Riwayat Komunikasi", type="primary"):
        log_communication(
            article_id=art["id"],
            recipient_type=rmode,
            recipient_name=rname,
            communication_type=ttype,
            subject=subject_edit,
            notes=f"Draft dicopy dari Generator Email. Penerima: {remail}"
        )
        st.success("Tercatat ke Riwayat Komunikasi artikel!")


def page_export():
    st.subheader("📥 Ekspor Excel")
    df = build_excel_export()
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button("Unduh Excel", data=dataframe_to_excel_bytes(df), file_name=f"Monitoring_All_{date.today().isoformat()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")


# =========================================================
# HALAMAN ANALISIS REVIEWER & AUTHOR (FITUR DIPERBARUI)
# =========================================================

def page_reviewer_analytics():
    st.subheader("📈 Analisis Performa Reviewer & Afiliasi Author")
    col1, col2 = st.columns(2)
    with col1: start_date = st.date_input("Mulai Tanggal", value=date.today() - timedelta(days=90))
    with col2: end_date = st.date_input("Sampai Tanggal", value=date.today())

    tab1, tab2 = st.tabs(["📊 Reviewer", "🏢 Author & Afiliasi"])
    
    with tab1:
        st.markdown("Menghitung seberapa banyak beban & kontribusi yang diselesaikan reviewer.")
        conn = get_connection()
        query_rev = """
            SELECT 
                reviewer_name AS "Nama Reviewer",
                reviewer_email AS "Email Reviewer",
                COUNT(id) AS "Total Ditugaskan",
                SUM(CASE WHEN status IN ('Review Submitted', 'Accepted', 'Declined') THEN 1 ELSE 0 END) AS "Total Direspons",
                SUM(CASE WHEN status = 'Review Submitted' THEN 1 ELSE 0 END) AS "Selesai Mereview"
            FROM reviewers
            WHERE date(invitation_date) BETWEEN ? AND ?
            GROUP BY reviewer_email, reviewer_name
            ORDER BY "Total Ditugaskan" DESC
        """
        df_rev = pd.read_sql_query(query_rev, conn, params=(start_date.isoformat(), end_date.isoformat()))
        conn.close()

        if df_rev.empty:
            st.info("Tidak ada data reviewer pada rentang waktu yang dipilih.")
        else:
            st.dataframe(df_rev, use_container_width=True, hide_index=True)
            st.bar_chart(df_rev.set_index("Nama Reviewer")["Total Ditugaskan"])

    with tab2:
        st.markdown("Mendata author dan instansi yang paling sering mengirim naskah ke jurnal.")
        conn = get_connection()
        query_auth = """
            SELECT 
                corresponding_author AS "Nama Author",
                affiliation_1 AS "Instansi / Afiliasi",
                COUNT(id) AS "Total Naskah Terkirim"
            FROM articles
            WHERE date(submission_date) BETWEEN ? AND ?
            GROUP BY corresponding_author, affiliation_1
            ORDER BY "Total Naskah Terkirim" DESC
        """
        df_auth = pd.read_sql_query(query_auth, conn, params=(start_date.isoformat(), end_date.isoformat()))
        conn.close()

        if df_auth.empty:
            st.info("Tidak ada data pengiriman artikel pada rentang waktu yang dipilih.")
        else:
            st.dataframe(df_auth, use_container_width=True, hide_index=True)


# =========================================================
# HALAMAN PENGATURAN
# =========================================================

def page_settings():
    st.subheader("🔧 Pengaturan")
    rf = st.text_input("Path Folder Utama", value=get_setting("root_folder", ""))
    if st.button("Simpan", type="primary"):
        save_setting("root_folder", rf.strip())
        st.success("Tersimpan!")

# =========================================================
# MAIN APP
# =========================================================

def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="📚", layout="wide", initial_sidebar_state="expanded")
    init_db()

    MENU_OPTIONS = {
        "📊 Dashboard": page_dashboard,
        "📝 Tambah Artikel": page_add_article,
        "⚙️ Kelola Artikel": page_manage_articles,
        "📂 Impor Dokumen": page_import_documents,
        "📧 Generator Email": page_email_drafts,
        "📈 Analisis Reviewer & Author": page_reviewer_analytics,
        "📥 Ekspor Excel": page_export,
        "🔧 Pengaturan": page_settings
    }

    with st.sidebar:
        st.markdown("<h2 style='text-align: center; color: #4CAF50;'>📚 Editorial App</h2>", unsafe_allow_html=True)
        st.markdown("---")
        menu_selection = st.radio("Menu Navigasi", list(MENU_OPTIONS.keys()))
        
        st.markdown("---")
        st.markdown("### 📌 Info Jurnal")
        for code, config in JOURNALS.items():
            st.markdown(f"**{code}** — {config['name']}")

        st.markdown("---")
        with st.expander("📂 Standar Penamaan File", expanded=False):
            st.markdown("""
            **Format Dasar:** `[Jurnal]_[ID]_[Nama]`
            
            **Fase Submisi:**
            - `..._Manuskrip` (Submisi asli)
            - `..._Format form` (Form cek format)
            - `..._turnitin` (Bukti similaritas)
            
            **Fase Desk Review:**
            - `..._Desk Review_1` (Maks 3)
            - `..._Revisi Desk review_1` (Maks 3)
            
            **Proses Review:**
            - `..._Reviewer Form_1` (Form rev 1/2)
            - `..._Reviewer Feedback_1` (Komen rev dlm naskah)
            - `..._Revision form (author's response)`
            - `..._Review Revision_1` (Revisi pasca review maks 3)
            
            **Fase Pasca Penerimaan:**
            - `..._APC`
            - `..._copyediting ready`
            - `..._copyedit revision_1` (Maks 5)
            """)

    # Jalankan fungsi halaman
    MENU_OPTIONS[menu_selection]()

if __name__ == "__main__":
    main()