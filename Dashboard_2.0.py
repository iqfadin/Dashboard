import os
import re
import sqlite3
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
    "Submission File": ["submission", "manuscript", "artikel"],
    "Author Checklist": ["checklist"],
    "Turnitin / Similarity Report": ["turnitin", "similarity", "plagiarism"],
    "Cover Letter": ["cover"],
}

# =========================================================
# KONFIGURASI UPLOAD DOKUMEN
# =========================================================

ALLOWED_UPLOAD_EXTENSIONS = [
    "doc",
    "docx",
    "pdf",
    "xlsx",
    "xls",
    "csv",
    "png",
    "jpg",
    "jpeg",
    "zip",
    "txt",
]

UPLOAD_FOLDER_RULES = {
    "02_Desk_Review": [
        "turnitin",
        "similarity",
        "plagiarism",
        "checklist",
        "desk review",
        "desk_review",
    ],
    "03_Reviewers": [
        "reviewer invitation",
        "reviewer_invitation",
        "invitation",
        "undangan reviewer",
    ],
    "04_Review_Results": [
        "review result",
        "review_results",
        "reviewer report",
        "review form",
        "hasil review",
        "reviewer comment",
        "comments reviewer",
    ],
    "05_Decision_Letters": [
        "decision letter",
        "decision_letter",
        "editor decision",
        "keputusan editor",
    ],
    "06_Revisi_Author": [
        "revision",
        "revisi",
        "revised manuscript",
        "response to reviewer",
        "response_to_reviewer",
    ],
    "07_APC": [
        "apc",
        "payment",
        "proof of payment",
        "payment proof",
        "bukti bayar",
        "invoice",
    ],
    "08_Copyediting": [
        "copyedit",
        "copy edit",
        "copyediting",
    ],
    "09_Layout_Proof": [
        "layout",
        "proof",
        "galley",
    ],
    "10_Published": [
        "published",
        "publication",
        "final published",
    ],
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

            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

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

    conn.commit()
    conn.close()


def get_setting(key, default=""):
    conn = get_connection()
    row = conn.execute(
        "SELECT setting_value FROM settings WHERE setting_key = ?",
        (key,)
    ).fetchone()
    conn.close()
    return row["setting_value"] if row else default


def save_setting(key, value):
    conn = get_connection()
    conn.execute("""
        INSERT INTO settings (setting_key, setting_value)
        VALUES (?, ?)
        ON CONFLICT(setting_key)
        DO UPDATE SET setting_value = excluded.setting_value
    """, (key, value))
    conn.commit()
    conn.close()


# =========================================================
# HELPER DATA
# =========================================================

def safe_date(value):
    """Mengubah string tanggal menjadi objek date bila memungkinkan."""
    if not value:
        return None
    try:
        return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None

def date_to_str(value):
    if isinstance(value, date):
        return value.isoformat()
    return value

def status_color(status):
    color_map = {
        "New Submission": "blue",
        "Desk Review": "blue",
        "Assign Reviewer": "orange",
        "Under Review": "orange",
        "Review Feedback Received": "violet",
        "Revision Requested": "orange",
        "Revision Submitted": "violet",
        "Accepted": "green",
        "APC Pending": "orange",
        "Payment Proof Submitted": "blue",
        "Payment Verified": "green",
        "Ready Copy Edit": "green",
        "Copy Editing": "green",
        "Layout Editing": "green",
        "Proofreading": "green",
        "Scheduled": "green",
        "Published": "green",
        "Rejected": "red",
        "Withdrawn": "red",
    }
    return color_map.get(status, "gray")

def get_articles_dataframe():
    conn = get_connection()
    df = pd.read_sql_query("""
        SELECT
            a.*,
            COUNT(r.id) AS reviewer_count
        FROM articles a
        LEFT JOIN reviewers r ON r.article_id = a.id
        GROUP BY a.id
        ORDER BY
            CASE WHEN a.status = 'Published' THEN 1 ELSE 0 END,
            a.updated_at DESC,
            a.id DESC
    """, conn)
    conn.close()
    return df

def get_article(article_id):
    conn = get_connection()
    row = conn.execute(
        "SELECT * FROM articles WHERE id = ?",
        (article_id,)
    ).fetchone()
    conn.close()
    return dict(row) if row else None

def get_article_id_by_code(article_code):
    conn = get_connection()
    row = conn.execute(
        "SELECT id FROM articles WHERE article_code = ?",
        (article_code,)
    ).fetchone()
    conn.close()
    return row["id"] if row else None

def get_reviewers(article_id=None):
    conn = get_connection()
    query = """
        SELECT
            r.*,
            a.article_code,
            a.journal_code,
            a.title,
            a.status AS article_status
        FROM reviewers r
        JOIN articles a ON a.id = r.article_id
    """
    params = []
    if article_id:
        query += " WHERE r.article_id = ?"
        params.append(article_id)
    query += " ORDER BY r.id DESC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(row) for row in rows]

def log_communication(
    article_id,
    recipient_type,
    recipient_name,
    communication_type,
    subject="",
    notes="",
    reviewer_id=None,
):
    conn = get_connection()
    conn.execute("""
        INSERT INTO communication_log (
            article_id, reviewer_id, communication_date,
            recipient_type, recipient_name, communication_type,
            subject, notes
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        article_id,
        reviewer_id,
        date.today().isoformat(),
        recipient_type,
        recipient_name,
        communication_type,
        subject,
        notes,
    ))
    conn.commit()
    conn.close()

def create_article_folders(article_code, root_folder):
    if not root_folder:
        return None, "Folder utama belum diatur."
    root_folder = os.path.abspath(root_folder)
    article_folder = os.path.join(root_folder, article_code)
    try:
        os.makedirs(article_folder, exist_ok=True)
        for subfolder in FOLDER_STRUCTURE:
            os.makedirs(
                os.path.join(article_folder, subfolder),
                exist_ok=True
            )
        return article_folder, None
    except Exception as error:
        return None, str(error)


# =========================================================
# HELPER AUTO-FILL & EKSTRAKSI DOKUMEN
# =========================================================

def parse_metadata_from_filename(filename, upload_year):
    """
    Mengambil Journal Code, ID Submission, Status, Catatan Urutan, dan Kode Artikel otomatis.
    Format yang diharapkan: HM_12345_Desk review_1
    """
    info = {}
    name_upper = filename.upper()
    name_lower = filename.lower()

    # 1. Mendeteksi Jurnal
    detected_journal = None
    for j_code in JOURNALS.keys():
        if j_code in name_upper:
            detected_journal = j_code
            info["journal_code"] = j_code
            break

    # 2. Mendeteksi ID OJS (Mencari angka tepat setelah inisial jurnal)
    submission_id = ""
    if detected_journal:
        # Pola pencarian: Inisial jurnal, diikuti oleh karakter pemisah (misal underscore/strip), lalu tangkap angkanya
        # Contoh: Jika 'HM_12345' -> tangkap '12345'
        pattern = rf"{detected_journal}[^A-Z0-9]*(\d+)"
        match = re.search(pattern, name_upper)
        if match:
            submission_id = match.group(1)
            
    # Jika cara di atas gagal, fallback mencari deretan angka dengan minimal 4 digit (misal: 12345)
    if not submission_id:
        fallback_match = re.search(r'\b\d{4,6}\b', filename)
        if fallback_match:
            submission_id = fallback_match.group()

    if submission_id:
        info["submission_id"] = submission_id

    # 3. Mendeteksi Status & Urutan (Misal: Desk Review 1)
    if "desk review 1" in name_lower or "desk_review_1" in name_lower:
        info["status"] = "Desk Review"
        info["notes"] = "Urutan: Desk Review 1"
    elif "desk review 2" in name_lower or "desk_review_2" in name_lower:
        info["status"] = "Desk Review"
        info["notes"] = "Urutan: Desk Review 2"
    elif "desk review" in name_lower or "desk_review" in name_lower:
        info["status"] = "Desk Review"

    # 4. Membuat Kode Artikel otomatis [InisialJurnal]-[TahunUpload]-[ID_OJS]
    # Contoh: HM-2026-12345
    if detected_journal and submission_id:
        info["article_code"] = f"{detected_journal}-{upload_year}-{submission_id}"

    return info


def extract_title_from_document(uploaded_file):
    """Membaca isi teks dokumen dan mengembalikan baris/kalimat pertama sebagai judul."""
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
            if lines:
                title = lines[0]
                
    except Exception:
        pass 

    uploaded_file.seek(0)
    return title


# =========================================================
# HELPER UPLOAD DOKUMEN
# =========================================================

def get_default_storage_root():
    configured_root = get_setting("root_folder", "").strip()
    if configured_root:
        return os.path.abspath(configured_root)
    return os.path.abspath("Artikel_Jurnal")

def sanitize_filename(filename):
    filename = os.path.basename(filename)
    return re.sub(r'[<>:"/\\|?*]+', "_", filename)

def get_unique_file_path(destination_folder, filename):
    filename = sanitize_filename(filename)
    name, extension = os.path.splitext(filename)
    target_path = os.path.join(destination_folder, filename)
    counter = 1
    while os.path.exists(target_path):
        target_path = os.path.join(
            destination_folder,
            f"{name}__{counter:02d}{extension}"
        )
        counter += 1
    return target_path

def detect_document_subfolder(filename):
    filename_lower = filename.lower()
    for folder_name, keywords in UPLOAD_FOLDER_RULES.items():
        if any(keyword.lower() in filename_lower for keyword in keywords):
            return folder_name
    return "01_Submission"

def save_uploaded_files(uploaded_files, article_folder):
    saved_files = []
    errors = []
    if not uploaded_files:
        return saved_files, errors
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


def create_outlook_draft(to_email, subject, body, attachment_path=None):
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to_email
        mail.Subject = subject
        mail.Body = body
        if attachment_path and os.path.exists(attachment_path):
            mail.Attachments.Add(os.path.abspath(attachment_path))
        mail.Display()
        return True, "Draf email Outlook berhasil dibuka."
    except Exception as error:
        return False, (
            "Draf Outlook tidak dapat dibuat. Pastikan Anda memakai Windows, "
            "Outlook desktop sudah terpasang, dan pywin32 sudah diinstal.\n\n"
            f"Detail: {error}"
        )


# =========================================================
# EXPORT EXCEL
# =========================================================

def build_excel_export():
    conn = get_connection()
    query = """
        SELECT
            a.article_code AS "Nomor",
            a.journal_code AS "Jurnal",
            a.submission_id AS "Submission ID",
            a.corresponding_author AS "Penulis",
            a.author_email AS "Email Penulis",
            a.affiliation_1 AS "Afiliasi 1",
            a.affiliation_2 AS "Afiliasi 2",
            a.title AS "Judul",
            a.status AS "Status Terbit",

            r.reviewer_name AS "Reviewer",
            r.reviewer_email AS "Email Reviewer",
            r.reviewer_category AS "Status reviewer",
            r.status AS "Status Proses Reviewer",
            r.invitation_date AS "Mulai Undangan",
            r.invitation_deadline AS "Batas Respons",
            r.acceptance_date AS "Reviewer Accept",
            r.review_deadline AS "Selesai Review",
            r.review_submitted_date AS "Hasil Review Masuk",
            r.recommendation AS "Rekomendasi Reviewer",
            r.reminder_count AS "Jumlah Reminder Reviewer",

            a.author_revision_deadline AS "Selesai Revisi Author",
            a.author_reminder_count AS "Jumlah Reminder Author",
            a.apc_deadline AS "Batas APC",
            a.payment_proof_date AS "Bukti Bayar Masuk",
            a.payment_verified_date AS "Pembayaran Terverifikasi",
            a.volume AS "Volume",
            a.issue AS "Issue",
            a.publication_year AS "Tahun",
            a.folder_path AS "Folder Artikel",
            a.notes AS "Catatan"
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
            max_length = 0
            column_letter = column_cells[0].column_letter
            for cell in column_cells:
                try:
                    max_length = max(max_length, len(str(cell.value or "")))
                except Exception:
                    pass
            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 45)
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
        SELECT
            r.*,
            a.article_code,
            a.journal_code,
            a.title,
            a.status AS article_status
        FROM reviewers r
        JOIN articles a ON a.id = r.article_id
        WHERE r.status IN ('Invited', 'Accepted', 'Unresponsive')
    """).fetchall()

    for row in reviewer_rows:
        reviewer = dict(row)
        if reviewer["status"] == "Invited":
            deadline = safe_date(reviewer["invitation_deadline"])
            if deadline and deadline <= today:
                days_late = (today - deadline).days
                if reviewer["reminder_count"] >= 2:
                    action = "Maksimum reminder tercapai — putuskan untuk mencari reviewer pengganti."
                    priority = "Kritis"
                else:
                    action = "Kirim reminder respons undangan reviewer."
                    priority = "Terlambat" if days_late > 0 else "Hari ini"
                tasks.append({
                    "Prioritas": priority,
                    "Jurnal": reviewer["journal_code"],
                    "Artikel": reviewer["article_code"],
                    "Judul": reviewer["title"],
                    "Penerima": reviewer["reviewer_name"],
                    "Tindakan": action,
                    "Deadline": deadline.isoformat(),
                    "Jenis": "Reviewer Invitation",
                })
        elif reviewer["status"] == "Accepted":
            deadline = safe_date(reviewer["review_deadline"])
            if deadline and deadline <= today:
                days_late = (today - deadline).days
                if reviewer["reminder_count"] >= 2:
                    action = "Maksimum reminder tercapai — evaluasi reviewer atau tetapkan pengganti."
                    priority = "Kritis"
                else:
                    action = "Kirim reminder hasil review."
                    priority = "Terlambat" if days_late > 0 else "Hari ini"
                tasks.append({
                    "Prioritas": priority,
                    "Jurnal": reviewer["journal_code"],
                    "Artikel": reviewer["article_code"],
                    "Judul": reviewer["title"],
                    "Penerima": reviewer["reviewer_name"],
                    "Tindakan": action,
                    "Deadline": deadline.isoformat(),
                    "Jenis": "Reviewer Deadline",
                })

    author_rows = conn.execute("""
        SELECT *
        FROM articles
        WHERE status = 'Revision Requested'
          AND author_revision_deadline IS NOT NULL
    """).fetchall()

    for row in author_rows:
        article = dict(row)
        deadline = safe_date(article["author_revision_deadline"])
        if deadline and deadline <= today:
            if article["author_reminder_count"] >= 3:
                action = "Final reminder expired — putuskan: Withdraw atau Extend Deadline."
                priority = "Kritis"
            else:
                action = "Kirim reminder revisi kepada author."
                priority = "Terlambat" if deadline < today else "Hari ini"
            tasks.append({
                "Prioritas": priority,
                "Jurnal": article["journal_code"],
                "Artikel": article["article_code"],
                "Judul": article["title"],
                "Penerima": article["corresponding_author"],
                "Tindakan": action,
                "Deadline": deadline.isoformat(),
                "Jenis": "Author Revision",
            })

    apc_rows = conn.execute("""
        SELECT *
        FROM articles
        WHERE status = 'APC Pending'
          AND apc_deadline IS NOT NULL
    """).fetchall()

    for row in apc_rows:
        article = dict(row)
        deadline = safe_date(article["apc_deadline"])
        if deadline and deadline <= today:
            tasks.append({
                "Prioritas": "Terlambat" if deadline < today else "Hari ini",
                "Jurnal": article["journal_code"],
                "Artikel": article["article_code"],
                "Judul": article["title"],
                "Penerima": article["corresponding_author"],
                "Tindakan": "Follow-up pembayaran APC / bukti pembayaran.",
                "Deadline": deadline.isoformat(),
                "Jenis": "APC",
            })
            
    conn.close()
    priority_order = {"Kritis": 0, "Terlambat": 1, "Hari ini": 2}
    tasks.sort(key=lambda item: (priority_order.get(item["Prioritas"], 9), item["Deadline"]))
    return tasks


def page_dashboard():
    st.subheader("Dashboard Tindakan Hari Ini")
    articles = get_articles_dataframe()
    tasks = get_dashboard_tasks()
    active_articles = 0
    if not articles.empty:
        active_articles = len(articles[~articles["status"].isin(["Published", "Rejected", "Withdrawn"])])
        
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Artikel Aktif", active_articles)
    col2.metric("Total Artikel", len(articles))
    col3.metric("Tindakan Perlu Diproses", len(tasks))
    col4.metric("Tindakan Kritis", len([task for task in tasks if task["Prioritas"] == "Kritis"]))

    st.markdown("### Daftar Tindakan")
    if not tasks:
        st.success("Tidak ada deadline atau reminder yang perlu diproses hari ini.")
    else:
        task_df = pd.DataFrame(tasks)
        st.dataframe(task_df, use_container_width=True, hide_index=True)

    st.markdown("### Ringkasan Status Artikel")
    if not articles.empty:
        status_summary = (
            articles.groupby(["journal_code", "status"])
            .size()
            .reset_index(name="Jumlah")
            .rename(columns={"journal_code": "Jurnal", "status": "Status Artikel"})
        )
        st.dataframe(status_summary, use_container_width=True, hide_index=True)

    st.info("Gunakan dashboard ini sebagai meja kerja harian. OJS dan Excel resmi tetap menjadi sumber administrasi utama.")


# =========================================================
# HALAMAN TAMBAH ARTIKEL
# =========================================================

def page_add_article():
    st.subheader("Tambah Artikel Baru")

    configured_root = get_setting("root_folder", "")
    active_root_folder = get_default_storage_root()

    st.caption(f"Folder penyimpanan aktif: `{active_root_folder}`")

    if not configured_root:
        st.info(
            "Folder utama belum diatur pada menu Pengaturan. "
            "Dokumen akan sementara disimpan pada folder lokal `Artikel_Jurnal`."
        )

    # 1. UPLOAD FILE TERLEBIH DAHULU UNTUK AUTO-FILL
    st.markdown("#### 1. Upload Dokumen Awal (Untuk Auto-fill Info)")
    st.info(
        "Unggah file di sini *sebelum* mengetik untuk mengisi informasi secara "
        "otomatis berdasarkan nama file dan baris pertama di dalam dokumen."
    )
    
    if "auto_uploader_key" not in st.session_state:
        st.session_state["auto_uploader_key"] = 0

    auto_uploaded_files = st.file_uploader(
        "Pilih satu atau beberapa dokumen",
        type=ALLOWED_UPLOAD_EXTENSIONS,
        accept_multiple_files=True,
        key=f"auto_uploader_{st.session_state['auto_uploader_key']}",
        help=(
            "Contoh dokumen: manuscript, checklist, Turnitin, "
            "cover letter, hasil review, revisi author, atau bukti pembayaran."
        )
    )

    # Nilai default variabel form
    def_journal = list(JOURNALS.keys())[0]
    def_sub_id = ""
    def_status = "New Submission"
    def_title = ""
    def_notes = ""
    def_article_code = ""

    if auto_uploaded_files:
        first_file = auto_uploaded_files[0]
        current_year = str(date.today().year)
        
        # Ekstrak dari nama file menggunakan fungsi yang telah diperbarui
        meta_info = parse_metadata_from_filename(first_file.name, current_year)
        if "journal_code" in meta_info:
            def_journal = meta_info["journal_code"]
        if "submission_id" in meta_info:
            def_sub_id = meta_info["submission_id"]
        if "status" in meta_info:
            def_status = meta_info["status"]
        if "notes" in meta_info:
            def_notes = meta_info["notes"]
        if "article_code" in meta_info:
            def_article_code = meta_info["article_code"]
            
        # Ekstrak judul dari isi dokumen
        extracted_title = extract_title_from_document(first_file)
        if extracted_title:
            def_title = extracted_title
            
        st.success(f"File **{first_file.name}** berhasil dibaca. Silakan periksa isian form di bawah.")

    try:
        journal_index = list(JOURNALS.keys()).index(def_journal)
    except ValueError:
        journal_index = 0

    try:
        status_index = ARTICLE_STATUSES.index(def_status)
    except ValueError:
        status_index = 0

    st.markdown("#### 2. Informasi Artikel")

    with st.form("form_add_article", clear_on_submit=True):
        col1, col2, col3 = st.columns(3)

        with col1:
            journal_code = st.selectbox("Jurnal", list(JOURNALS.keys()), index=journal_index)
            article_code = st.text_input("Kode Artikel *", value=def_article_code, placeholder="Contoh: HM-2026-12345")
            submission_id = st.text_input("Submission ID OJS", value=def_sub_id)

        with col2:
            corresponding_author = st.text_input("Corresponding Author")
            author_email = st.text_input("Email Author")
            submission_date = st.date_input("Tanggal Submit", value=date.today())

        with col3:
            affiliation_1 = st.text_input("Afiliasi 1")
            affiliation_2 = st.text_input("Afiliasi 2")
            status = st.selectbox("Status Awal", ARTICLE_STATUSES, index=status_index)

        title = st.text_area("Judul Artikel *", value=def_title, height=100)
        notes = st.text_area("Catatan Internal", value=def_notes, height=80)

        create_folder = st.checkbox(
            "Buat struktur folder artikel otomatis",
            value=True
        )

        submitted = st.form_submit_button(
            "Simpan Artikel",
            type="primary"
        )

    if submitted:
        if not article_code.strip() or not title.strip():
            st.error("Kode artikel dan judul artikel wajib diisi.")
            return

        clean_article_code = article_code.strip()
        folder_path = ""

        should_create_folder = create_folder or bool(auto_uploaded_files)

        if should_create_folder:
            folder_path, error = create_article_folders(
                clean_article_code,
                active_root_folder
            )
            if error:
                st.error(
                    f"Artikel tidak dapat disimpan karena folder gagal dibuat: {error}"
                )
                return

        try:
            conn = get_connection()
            conn.execute("""
                INSERT INTO articles (
                    article_code, journal_code, submission_id, title,
                    corresponding_author, author_email,
                    affiliation_1, affiliation_2, status,
                    submission_date, folder_path, notes, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                clean_article_code,
                journal_code,
                submission_id.strip(),
                title.strip(),
                corresponding_author.strip(),
                author_email.strip(),
                affiliation_1.strip(),
                affiliation_2.strip(),
                status,
                submission_date.isoformat(),
                folder_path,
                notes.strip(),
                datetime.now().isoformat(timespec="seconds"),
            ))
            conn.commit()
            conn.close()

            article_id = get_article_id_by_code(clean_article_code)

            saved_files = []
            upload_errors = []

            if auto_uploaded_files and folder_path:
                saved_files, upload_errors = save_uploaded_files(
                    auto_uploaded_files,
                    folder_path
                )

            if article_id:
                log_communication(
                    article_id=article_id,
                    recipient_type="Internal",
                    recipient_name="Editorial Team",
                    communication_type="Article Created",
                    subject=f"Artikel baru dibuat — {clean_article_code}",
                    notes=(
                        "Artikel dibuat pada aplikasi. "
                        f"Jumlah dokumen awal diunggah: {len(saved_files)}."
                    )
                )

            st.success(f"Artikel {clean_article_code} berhasil disimpan.")

            if folder_path:
                st.caption(f"Folder artikel: `{folder_path}`")

            if saved_files:
                st.success(f"{len(saved_files)} dokumen berhasil diunggah.")
                with st.expander("Lihat dokumen yang tersimpan"):
                    for saved_file in saved_files:
                        st.write(f"✅ {saved_file}")

            if upload_errors:
                st.warning("Beberapa file tidak dapat diunggah:")
                for error_message in upload_errors:
                    st.write(f"⚠️ {error_message}")

            st.session_state["auto_uploader_key"] += 1
            st.rerun()

        except sqlite3.IntegrityError:
            st.error(
                "Kode artikel sudah digunakan. Gunakan kode yang berbeda."
            )


# =========================================================
# HALAMAN KELOLA ARTIKEL
# =========================================================

def page_manage_articles():
    st.subheader("Kelola Artikel")
    articles = get_articles_dataframe()
    if articles.empty:
        st.info("Belum ada artikel. Tambahkan artikel terlebih dahulu.")
        return

    search = st.text_input(
        "Cari kode artikel, judul, author, atau jurnal",
        placeholder="Contoh: HM-2026-12345 atau nama author"
    )

    display_df = articles.copy()
    if search:
        search_lower = search.lower()
        mask = (
            display_df["article_code"].fillna("").str.lower().str.contains(search_lower) |
            display_df["title"].fillna("").str.lower().str.contains(search_lower) |
            display_df["corresponding_author"].fillna("").str.lower().str.contains(search_lower) |
            display_df["journal_code"].fillna("").str.lower().str.contains(search_lower)
        )
        display_df = display_df[mask]

    if display_df.empty:
        st.warning("Artikel tidak ditemukan.")
        return

    display_columns = [
        "article_code", "journal_code", "title", "corresponding_author",
        "status", "reviewer_count", "author_revision_deadline", "apc_deadline"
    ]
    st.dataframe(
        display_df[display_columns].rename(columns={
            "article_code": "Kode",
            "journal_code": "Jurnal",
            "title": "Judul",
            "corresponding_author": "Author",
            "status": "Status",
            "reviewer_count": "Jumlah Reviewer",
            "author_revision_deadline": "Deadline Revisi",
            "apc_deadline": "Deadline APC",
        }),
        use_container_width=True, hide_index=True
    )

    selected_code = st.selectbox("Pilih artikel untuk dikelola", display_df["article_code"].tolist())
    selected_row = articles[articles["article_code"] == selected_code].iloc[0]
    article = get_article(int(selected_row["id"]))

    st.markdown("---")
    st.markdown(f"### {article['article_code']} — {article['journal_code']}")
    st.write(article["title"])

    tabs = st.tabs([
        "Status & Metadata",
        "Reviewer",
        "Revisi / APC",
        "Folder & Checklist",
        "Riwayat Komunikasi",
    ])

    # --- TAB STATUS & METADATA ---
    with tabs[0]:
        current_index = ARTICLE_STATUSES.index(article["status"]) if article["status"] in ARTICLE_STATUSES else 0
        with st.form(f"form_article_status_{article['id']}"):
            status = st.selectbox("Status Artikel", ARTICLE_STATUSES, index=current_index)
            col1, col2, col3 = st.columns(3)
            with col1:
                volume = st.text_input("Volume", value=article["volume"] or "")
            with col2:
                issue = st.text_input("Issue", value=article["issue"] or "")
            with col3:
                publication_year = st.text_input("Tahun", value=article["publication_year"] or "")
            notes = st.text_area("Catatan Internal", value=article["notes"] or "", height=120)
            save_status = st.form_submit_button("Simpan Perubahan", type="primary")

        if save_status:
            conn = get_connection()
            conn.execute("""
                UPDATE articles
                SET status = ?, volume = ?, issue = ?, publication_year = ?, notes = ?, updated_at = ?
                WHERE id = ?
            """, (
                status, volume.strip(), issue.strip(), publication_year.strip(),
                notes.strip(), datetime.now().isoformat(timespec="seconds"), article["id"]
            ))
            conn.commit()
            conn.close()
            st.success("Perubahan artikel berhasil disimpan.")
            st.rerun()

        if article["status"] not in ["Ready Copy Edit", "Copy Editing", "Layout Editing", "Proofreading", "Scheduled", "Published"]:
            st.caption("Volume, issue, dan tahun bersifat opsional dan umumnya diisi mulai tahap Ready Copy Edit.")

    # --- TAB REVIEWER ---
    with tabs[1]:
        reviewers = get_reviewers(article["id"])
        st.markdown("#### Reviewer Terdaftar")
        if reviewers:
            reviewer_df = pd.DataFrame(reviewers)
            st.dataframe(
                reviewer_df[[
                    "reviewer_name", "reviewer_email", "reviewer_category", "status",
                    "invitation_date", "invitation_deadline", "acceptance_date",
                    "review_deadline", "review_submitted_date", "recommendation", "reminder_count"
                ]].rename(columns={
                    "reviewer_name": "Nama", "reviewer_email": "Email",
                    "reviewer_category": "Kategori", "status": "Status Proses",
                    "invitation_date": "Undangan", "invitation_deadline": "Batas Respons",
                    "acceptance_date": "Accept", "review_deadline": "Deadline Review",
                    "review_submitted_date": "Review Masuk", "recommendation": "Rekomendasi",
                    "reminder_count": "Reminder"
                }),
                use_container_width=True, hide_index=True
            )
        else:
            st.warning("Belum ada reviewer untuk artikel ini.")

        if len(reviewers) >= 2:
            st.info("Artikel ini sudah memiliki 2 reviewer. Tambahkan reviewer baru hanya jika memang diperlukan.")

        st.markdown("#### Tambah Reviewer")
        with st.form(f"form_add_reviewer_{article['id']}", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                reviewer_name = st.text_input("Nama Reviewer")
                reviewer_email = st.text_input("Email Reviewer")
                reviewer_category = st.selectbox("Status reviewer untuk Excel", REVIEWER_CATEGORIES)
            with col2:
                invitation_date = st.date_input("Tanggal Undangan", value=date.today())
                reviewer_notes = st.text_area("Catatan Reviewer", height=100)
            add_reviewer = st.form_submit_button("Tambah Reviewer")

        if add_reviewer:
            if not reviewer_name.strip():
                st.error("Nama reviewer wajib diisi.")
            else:
                invitation_deadline = invitation_date + timedelta(days=7)
                conn = get_connection()
                conn.execute("""
                    INSERT INTO reviewers (
                        article_id, reviewer_name, reviewer_email, reviewer_category,
                        invitation_date, invitation_deadline, status, notes
                    ) VALUES (?, ?, ?, ?, ?, ?, 'Invited', ?)
                """, (
                    article["id"], reviewer_name.strip(), reviewer_email.strip(),
                    reviewer_category, invitation_date.isoformat(),
                    invitation_deadline.isoformat(), reviewer_notes.strip()
                ))
                conn.execute("""
                    UPDATE articles SET status = 'Under Review', updated_at = ? WHERE id = ?
                """, (datetime.now().isoformat(timespec="seconds"), article["id"]))
                conn.commit()
                conn.close()
                log_communication(
                    article_id=article["id"], reviewer_id=None, recipient_type="Reviewer",
                    recipient_name=reviewer_name.strip(), communication_type="Reviewer Invitation",
                    subject=f"Invitation to Review — {article['article_code']}",
                    notes="Reviewer ditambahkan ke aplikasi."
                )
                st.success(f"Reviewer ditambahkan. Batas respons otomatis: {invitation_deadline.isoformat()}")
                st.rerun()

        if reviewers:
            st.markdown("#### Perbarui Status Reviewer")
            reviewer_options = {f"{r['reviewer_name']} — {r['status']}": r for r in reviewers}
            reviewer_label = st.selectbox("Pilih reviewer", list(reviewer_options.keys()))
            selected_reviewer = reviewer_options[reviewer_label]
            current_reviewer_status_index = REVIEWER_STATUSES.index(selected_reviewer["status"]) if selected_reviewer["status"] in REVIEWER_STATUSES else 0
            current_recommendation = selected_reviewer["recommendation"] or ""
            recommendation_index = REVIEW_RECOMMENDATIONS.index(current_recommendation) if current_recommendation in REVIEW_RECOMMENDATIONS else 0

            with st.form(f"form_update_reviewer_{selected_reviewer['id']}"):
                col1, col2 = st.columns(2)
                with col1:
                    reviewer_status = st.selectbox("Status Proses Reviewer", REVIEWER_STATUSES, index=current_reviewer_status_index)
                    recommendation = st.selectbox("Rekomendasi Reviewer", REVIEW_RECOMMENDATIONS, index=recommendation_index)
                with col2:
                    accepted_value = safe_date(selected_reviewer["acceptance_date"]) or date.today()
                    submitted_value = safe_date(selected_reviewer["review_submitted_date"]) or date.today()
                    acceptance_date = st.date_input("Tanggal Accept", value=accepted_value)
                    review_submitted_date = st.date_input("Tanggal Hasil Review Masuk", value=submitted_value)
                reviewer_update_notes = st.text_area("Catatan", value=selected_reviewer["notes"] or "")
                update_reviewer = st.form_submit_button("Simpan Status Reviewer")

            if update_reviewer:
                review_deadline = selected_reviewer["review_deadline"]
                acceptance_date_db = selected_reviewer["acceptance_date"]
                submitted_date_db = selected_reviewer["review_submitted_date"]
                article_status_update = None
                if reviewer_status == "Accepted":
                    acceptance_date_db = acceptance_date.isoformat()
                    review_deadline = (acceptance_date + timedelta(days=14)).isoformat()
                elif reviewer_status == "Review Submitted":
                    submitted_date_db = review_submitted_date.isoformat()
                    article_status_update = "Review Feedback Received"
                elif reviewer_status == "Declined":
                    article_status_update = "Assign Reviewer"

                conn = get_connection()
                conn.execute("""
                    UPDATE reviewers
                    SET status = ?, recommendation = ?, acceptance_date = ?,
                        review_deadline = ?, review_submitted_date = ?, notes = ?
                    WHERE id = ?
                """, (
                    reviewer_status, recommendation, acceptance_date_db,
                    review_deadline, submitted_date_db, reviewer_update_notes.strip(),
                    selected_reviewer["id"]
                ))
                if article_status_update:
                    conn.execute("""
                        UPDATE articles SET status = ?, updated_at = ? WHERE id = ?
                    """, (
                        article_status_update, datetime.now().isoformat(timespec="seconds"), article["id"]
                    ))
                conn.commit()
                conn.close()
                st.success("Status reviewer berhasil diperbarui.")
                st.rerun()

            st.markdown("#### Reminder Reviewer")
            col1, col2 = st.columns(2)
            with col1:
                st.write(f"Jumlah reminder saat ini: **{selected_reviewer['reminder_count']} / 2**")
            with col2:
                if st.button("Catat Reminder Reviewer", key=f"reviewer_reminder_{selected_reviewer['id']}"):
                    if selected_reviewer["reminder_count"] >= 2:
                        st.error("Maksimum 2 reminder sudah tercapai. Pertimbangkan mencari reviewer pengganti.")
                    else:
                        new_count = selected_reviewer["reminder_count"] + 1
                        conn = get_connection()
                        conn.execute("UPDATE reviewers SET reminder_count = ? WHERE id = ?", (new_count, selected_reviewer["id"]))
                        conn.commit()
                        conn.close()
                        log_communication(
                            article_id=article["id"], reviewer_id=selected_reviewer["id"],
                            recipient_type="Reviewer", recipient_name=selected_reviewer["reviewer_name"],
                            communication_type="Reviewer Reminder",
                            subject=f"Review Reminder — {article['article_code']}",
                            notes=f"Reminder reviewer ke-{new_count}."
                        )
                        st.success(f"Reminder reviewer ke-{new_count} dicatat.")
                        st.rerun()

    # --- TAB REVISI DAN APC ---
    with tabs[2]:
        st.markdown("#### Revisi Author")
        current_revision_deadline = safe_date(article["author_revision_deadline"])
        with st.form(f"form_revision_{article['id']}"):
            set_revision_deadline = st.checkbox("Atur / ubah deadline revisi author", value=current_revision_deadline is not None)
            revision_deadline = st.date_input("Deadline Revisi Author", value=(current_revision_deadline or (date.today() + timedelta(days=14))))
            save_revision = st.form_submit_button("Simpan Deadline Revisi")

        if save_revision:
            deadline_value = revision_deadline.isoformat() if set_revision_deadline else None
            conn = get_connection()
            conn.execute("""
                UPDATE articles
                SET author_revision_deadline = ?,
                    status = CASE WHEN ? IS NOT NULL THEN 'Revision Requested' ELSE status END,
                    updated_at = ?
                WHERE id = ?
            """, (
                deadline_value, deadline_value, datetime.now().isoformat(timespec="seconds"), article["id"]
            ))
            conn.commit()
            conn.close()
            st.success("Deadline revisi author berhasil disimpan.")
            st.rerun()

        st.write(f"Jumlah reminder author: **{article['author_reminder_count']} / 3**")
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("Catat Reminder Author", key=f"author_reminder_{article['id']}"):
                if article["author_reminder_count"] >= 3:
                    st.error("Maksimum reminder author sudah tercapai. Pilih Withdraw atau Extend Deadline.")
                else:
                    new_count = article["author_reminder_count"] + 1
                    conn = get_connection()
                    conn.execute("UPDATE articles SET author_reminder_count = ? WHERE id = ?", (new_count, article["id"]))
                    conn.commit()
                    conn.close()
                    log_communication(
                        article_id=article["id"], recipient_type="Author",
                        recipient_name=article["corresponding_author"],
                        communication_type="Author Revision Reminder",
                        subject=f"Revision Reminder — {article['article_code']}",
                        notes=f"Reminder revisi author ke-{new_count}."
                    )
                    st.success(f"Reminder author ke-{new_count} dicatat.")
                    st.rerun()
        with col2:
            if st.button("Extend Deadline", key=f"extend_{article['id']}"):
                st.session_state[f"show_extend_{article['id']}"] = True
        with col3:
            if st.button("Withdraw Artikel", key=f"withdraw_{article['id']}"):
                conn = get_connection()
                conn.execute("""
                    UPDATE articles SET status = 'Withdrawn', updated_at = ? WHERE id = ?
                """, (datetime.now().isoformat(timespec="seconds"), article["id"]))
                conn.commit()
                conn.close()
                log_communication(
                    article_id=article["id"], recipient_type="Author",
                    recipient_name=article["corresponding_author"],
                    communication_type="Withdrawal Decision",
                    subject=f"Withdrawal — {article['article_code']}",
                    notes="Artikel diubah menjadi Withdrawn secara manual."
                )
                st.warning("Status artikel diubah menjadi Withdrawn.")
                st.rerun()

        if st.session_state.get(f"show_extend_{article['id']}", False):
            new_deadline = st.date_input("Pilih deadline revisi baru", value=(current_revision_deadline or date.today()) + timedelta(days=7), key=f"new_deadline_{article['id']}")
            if st.button("Simpan Deadline Baru", key=f"save_extend_{article['id']}"):
                conn = get_connection()
                conn.execute("""
                    UPDATE articles
                    SET author_revision_deadline = ?, author_reminder_count = 0,
                        status = 'Revision Requested', updated_at = ?
                    WHERE id = ?
                """, (
                    new_deadline.isoformat(), datetime.now().isoformat(timespec="seconds"), article["id"]
                ))
                conn.commit()
                conn.close()
                log_communication(
                    article_id=article["id"], recipient_type="Author",
                    recipient_name=article["corresponding_author"],
                    communication_type="Revision Deadline Extended",
                    subject=f"Extended Revision Deadline — {article['article_code']}",
                    notes=f"Deadline revisi diperpanjang sampai {new_deadline.isoformat()}."
                )
                st.session_state[f"show_extend_{article['id']}"] = False
                st.success("Deadline revisi diperpanjang dan reminder direset.")
                st.rerun()

        st.markdown("---")
        st.markdown("#### APC")
        apc_deadline_current = safe_date(article["apc_deadline"])
        payment_proof_current = safe_date(article["payment_proof_date"])
        payment_verified_current = safe_date(article["payment_verified_date"])

        with st.form(f"form_apc_{article['id']}"):
            col1, col2, col3 = st.columns(3)
            with col1:
                apc_deadline = st.date_input("Deadline APC", value=(apc_deadline_current or (date.today() + timedelta(days=14))))
            with col2:
                payment_proof_date = st.date_input("Tanggal Bukti Bayar Masuk", value=payment_proof_current or date.today())
            with col3:
                payment_verified_date = st.date_input("Tanggal Pembayaran Diverifikasi", value=payment_verified_current or date.today())

            apc_status = st.selectbox("Status APC", ["APC Pending", "Payment Proof Submitted", "Payment Verified", "Ready Copy Edit"], index=0)
            save_apc = st.form_submit_button("Simpan Status APC")

        if save_apc:
            proof_date = payment_proof_date.isoformat() if apc_status in ["Payment Proof Submitted", "Payment Verified", "Ready Copy Edit"] else None
            verified_date = payment_verified_date.isoformat() if apc_status in ["Payment Verified", "Ready Copy Edit"] else None
            conn = get_connection()
            conn.execute("""
                UPDATE articles
                SET status = ?, apc_deadline = ?, payment_proof_date = ?,
                    payment_verified_date = ?, updated_at = ?
                WHERE id = ?
            """, (
                apc_status, apc_deadline.isoformat(), proof_date, verified_date,
                datetime.now().isoformat(timespec="seconds"), article["id"]
            ))
            conn.commit()
            conn.close()
            st.success("Status APC berhasil disimpan.")
            st.rerun()

    # --- TAB FOLDER DAN CHECKLIST ---
    with tabs[3]:
        st.markdown("#### Folder Artikel")
        root_folder = get_default_storage_root()
        st.write(f"Folder utama saat ini: `{root_folder}`")
        folder_path = article["folder_path"]
        if not folder_path:
            st.warning("Folder artikel belum dibuat.")
        else:
            st.write(f"Folder artikel: `{folder_path}`")

        if st.button("Buat / Perbaiki Struktur Folder", key=f"folder_{article['id']}"):
            new_folder_path, error = create_article_folders(article["article_code"], root_folder)
            if error:
                st.error(f"Gagal membuat folder: {error}")
            else:
                conn = get_connection()
                conn.execute("UPDATE articles SET folder_path = ?, updated_at = ? WHERE id = ?", (new_folder_path, datetime.now().isoformat(timespec="seconds"), article["id"]))
                conn.commit()
                conn.close()
                st.success("Struktur folder berhasil dibuat.")
                st.rerun()

        st.markdown("#### Upload Dokumen untuk Artikel Ini")
        if folder_path and os.path.exists(folder_path):
            uploaded_files_from_tab = st.file_uploader("Pilih dokumen tambahan", type=ALLOWED_UPLOAD_EXTENSIONS, accept_multiple_files=True, key=f"manage_upload_{article['id']}")
            if uploaded_files_from_tab:
                preview_rows = [{"Nama File": uf.name, "Ukuran": f"{uf.size / 1024:.1f} KB", "Folder Tujuan": detect_document_subfolder(uf.name)} for uf in uploaded_files_from_tab]
                st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)
            if st.button("Simpan Dokumen", key=f"save_manage_upload_{article['id']}", type="primary", disabled=not uploaded_files_from_tab):
                saved_files, upload_errors = save_uploaded_files(uploaded_files_from_tab, folder_path)
                if saved_files:
                    log_communication(
                        article_id=article["id"], recipient_type="Internal",
                        recipient_name="Editorial Team", communication_type="Document Upload",
                        subject=f"Dokumen diunggah — {article['article_code']}",
                        notes=f"{len(saved_files)} dokumen diunggah: " + "; ".join(saved_files)
                    )
                    st.success(f"{len(saved_files)} dokumen berhasil disimpan.")
                if upload_errors:
                    st.warning("Beberapa dokumen gagal diunggah:")
                    for error_message in upload_errors:
                        st.write(f"⚠️ {error_message}")
                if saved_files:
                    st.rerun()
        else:
            st.info("Buat struktur folder artikel terlebih dahulu agar dokumen dapat diunggah.")

        st.markdown("#### Checklist Dokumen")
        if folder_path and os.path.exists(folder_path):
            all_files = []
            for root, _, files in os.walk(folder_path):
                for file_name in files:
                    all_files.append(os.path.relpath(os.path.join(root, file_name), folder_path))

            checklist_rows = []
            for document_name, keywords in DOCUMENT_CHECKLIST.items():
                found_files = [fn for fn in all_files if any(keyword.lower() in fn.lower() for keyword in keywords)]
                checklist_rows.append({
                    "Dokumen": document_name,
                    "Status": "Tersedia" if found_files else "Belum ditemukan",
                    "File Terdeteksi": ", ".join(found_files) if found_files else "-"
                })
            st.dataframe(pd.DataFrame(checklist_rows), use_container_width=True, hide_index=True)
            with st.expander("Lihat semua file dalam folder artikel"):
                if all_files:
                    st.write(all_files)
                else:
                    st.caption("Belum ada file di dalam folder artikel.")
        else:
            st.info("Checklist otomatis akan tersedia setelah folder artikel dibuat dan path folder valid.")

    # --- TAB RIWAYAT KOMUNIKASI ---
    with tabs[4]:
        conn = get_connection()
        logs = pd.read_sql_query("""
            SELECT communication_date, recipient_type, recipient_name, communication_type, subject, notes
            FROM communication_log WHERE article_id = ? ORDER BY communication_date DESC, id DESC
        """, conn, params=(article["id"],))
        conn.close()

        if logs.empty:
            st.info("Belum ada riwayat komunikasi.")
        else:
            st.dataframe(logs, use_container_width=True, hide_index=True)

        st.markdown("#### Tambah Catatan Komunikasi Manual")
        with st.form(f"form_manual_log_{article['id']}", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                recipient_type = st.selectbox("Jenis Penerima", ["Author", "Reviewer", "Editor", "Other"])
                recipient_name = st.text_input("Nama Penerima")
            with col2:
                communication_type = st.text_input("Jenis Komunikasi", placeholder="Contoh: Email follow-up")
                subject = st.text_input("Subjek")
            log_notes = st.text_area("Catatan")
            save_log = st.form_submit_button("Simpan Catatan")

        if save_log:
            if not communication_type.strip():
                st.error("Jenis komunikasi wajib diisi.")
            else:
                log_communication(
                    article_id=article["id"], recipient_type=recipient_type,
                    recipient_name=recipient_name.strip(), communication_type=communication_type.strip(),
                    subject=subject.strip(), notes=log_notes.strip()
                )
                st.success("Catatan komunikasi disimpan.")
                st.rerun()


# =========================================================
# HALAMAN IMPOR DOKUMEN
# =========================================================

def page_import_documents():
    st.subheader("Impor Dokumen Artikel")
    st.write("Gunakan menu ini untuk mengunggah dokumen tambahan ke artikel yang sudah terdaftar...")

    articles = get_articles_dataframe()
    if articles.empty:
        st.info("Belum ada artikel. Tambahkan artikel terlebih dahulu.")
        return

    article_options = {f"{row['article_code']} — {row['title']}": int(row["id"]) for _, row in articles.iterrows()}
    selected_label = st.selectbox("Pilih Artikel", list(article_options.keys()))
    article_id = article_options[selected_label]
    article = get_article(article_id)

    if not article:
        st.error("Artikel tidak ditemukan.")
        return

    st.markdown(f"**Artikel:** {article['article_code']}")
    st.markdown(f"**Jurnal:** {article['journal_code']}")
    st.markdown(f"**Judul:** {article['title']}")

    active_root_folder = get_default_storage_root()
    folder_path = article["folder_path"]

    if not folder_path or not os.path.exists(folder_path):
        st.warning("Folder artikel belum tersedia. Buat struktur folder terlebih dahulu.")
        if st.button("Buat Struktur Folder Artikel", type="primary"):
            folder_path, error = create_article_folders(article["article_code"], active_root_folder)
            if error:
                st.error(f"Gagal membuat folder: {error}")
                return
            conn = get_connection()
            conn.execute("UPDATE articles SET folder_path = ?, updated_at = ? WHERE id = ?", (folder_path, datetime.now().isoformat(timespec="seconds"), article["id"]))
            conn.commit()
            conn.close()
            st.success("Folder artikel berhasil dibuat.")
            st.rerun()
        return

    st.caption(f"Lokasi folder artikel: `{folder_path}`")

    uploaded_files = st.file_uploader(
        "Pilih dokumen untuk diunggah",
        type=ALLOWED_UPLOAD_EXTENSIONS,
        accept_multiple_files=True,
        key=f"upload_documents_{article['id']}",
        help="Folder tujuan ditentukan otomatis dari nama file."
    )

    if uploaded_files:
        st.markdown("#### File yang Dipilih")
        preview_rows = [{"Nama File": uf.name, "Ukuran": f"{uf.size / 1024:.1f} KB", "Folder Tujuan Otomatis": detect_document_subfolder(uf.name)} for uf in uploaded_files]
        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

    if st.button("Simpan Dokumen ke Folder Artikel", type="primary", disabled=not uploaded_files):
        saved_files, upload_errors = save_uploaded_files(uploaded_files, folder_path)
        if saved_files:
            log_communication(
                article_id=article["id"], recipient_type="Internal",
                recipient_name="Editorial Team", communication_type="Document Upload",
                subject=f"Dokumen diunggah — {article['article_code']}",
                notes=f"{len(saved_files)} dokumen diunggah: " + "; ".join(saved_files)
            )
            st.success(f"{len(saved_files)} dokumen berhasil disimpan.")
            with st.expander("Lihat lokasi dokumen yang tersimpan"):
                for saved_file in saved_files:
                    st.write(f"✅ {saved_file}")
        if upload_errors:
            st.warning("Beberapa dokumen gagal diunggah:")
            for error_message in upload_errors:
                st.write(f"⚠️ {error_message}")
        if saved_files:
            st.rerun()


# =========================================================
# HALAMAN GENERATOR EMAIL
# =========================================================

def page_email_drafts():
    st.subheader("Generator Draf Email Outlook")
    articles = get_articles_dataframe()
    if articles.empty:
        st.info("Belum ada artikel.")
        return

    selected_code = st.selectbox("Pilih Artikel", articles["article_code"].tolist(), key="email_article_select")
    selected_row = articles[articles["article_code"] == selected_code].iloc[0]
    article = get_article(int(selected_row["id"]))
    reviewers = get_reviewers(article["id"])

    recipient_mode = st.radio("Penerima", ["Author", "Reviewer"])
    recipient_name = ""
    recipient_email = ""

    if recipient_mode == "Author":
        recipient_name = article["corresponding_author"] or ""
        recipient_email = article["author_email"] or ""
    else:
        if not reviewers:
            st.warning("Artikel ini belum memiliki reviewer.")
            return
        reviewer_options = {f"{r['reviewer_name']} — {r['reviewer_email']}": r for r in reviewers}
        selected_reviewer_label = st.selectbox("Pilih Reviewer", list(reviewer_options.keys()))
        selected_reviewer = reviewer_options[selected_reviewer_label]
        recipient_name = selected_reviewer["reviewer_name"] or ""
        recipient_email = selected_reviewer["reviewer_email"] or ""

    template_type = st.selectbox("Template", ["Reviewer Invitation", "Reviewer Reminder", "Author Revision Reminder", "APC Reminder", "Custom"])

    if template_type == "Reviewer Invitation":
        default_subject = f"Invitation to Review — {article['article_code']}"
        default_body = f"Dear {recipient_name},\n\nWe would like to invite you to review the manuscript below.\n\nArticle ID: {article['article_code']}\nTitle: {article['title']}\n\nPlease confirm whether you are available to review this manuscript.\n\nThank you for your consideration.\n\nBest regards,\nEditorial Team\n{article['journal_code']}"
    elif template_type == "Reviewer Reminder":
        default_subject = f"Review Reminder — {article['article_code']}"
        default_body = f"Dear {recipient_name},\n\nThis is a friendly reminder regarding the review of the following manuscript.\n\nArticle ID: {article['article_code']}\nTitle: {article['title']}\n\nWe would greatly appreciate it if you could submit your review at your earliest convenience.\n\nThank you for your valuable contribution.\n\nBest regards,\nEditorial Team\n{article['journal_code']}"
    elif template_type == "Author Revision Reminder":
        default_subject = f"Revision Reminder — {article['article_code']}"
        default_body = f"Dear {recipient_name},\n\nThis is a reminder regarding the revision of your manuscript.\n\nArticle ID: {article['article_code']}\nTitle: {article['title']}\n\nPlease submit your revised manuscript and response to reviewers through OJS before the assigned deadline.\n\nBest regards,\nEditorial Team\n{article['journal_code']}"
    elif template_type == "APC Reminder":
        default_subject = f"APC Payment Reminder — {article['article_code']}"
        default_body = f"Dear {recipient_name},\n\nThis is a reminder regarding the Article Processing Charge (APC) for the following accepted manuscript.\n\nArticle ID: {article['article_code']}\nTitle: {article['title']}\n\nPlease submit the payment proof according to the editorial instruction.\n\nBest regards,\nEditorial Team\n{article['journal_code']}"
    else:
        default_subject = ""
        default_body = f"Dear {recipient_name},\n\n"

    to_email = st.text_input("Email Penerima", value=recipient_email)
    subject = st.text_input("Subjek Email", value=default_subject)
    body = st.text_area("Isi Email", value=default_body, height=300)
    attachment_path = st.text_input("Path Lampiran Opsional", placeholder=r"Contoh: C:\Editorial\HM-2026-12345\05_Decision_Letters\decision.docx")

    if st.button("Buka Draf di Outlook", type="primary"):
        if not to_email.strip():
            st.error("Email penerima belum diisi.")
        elif not subject.strip():
            st.error("Subjek email belum diisi.")
        else:
            success, message = create_outlook_draft(to_email.strip(), subject.strip(), body, attachment_path.strip() or None)
            if success:
                st.success(message)
                log_communication(
                    article_id=article["id"], recipient_type=recipient_mode, recipient_name=recipient_name,
                    communication_type=template_type, subject=subject.strip(), notes="Draf Outlook dibuat dari aplikasi."
                )
            else:
                st.error(message)


# =========================================================
# HALAMAN EKSPOR EXCEL
# =========================================================

def page_export():
    st.subheader("Ekspor Excel Monitoring All")
    st.write("Format ekspor menggunakan prinsip **satu reviewer = satu baris**. Apabila sebuah artikel memiliki dua reviewer, informasi artikelnya akan tampil dua kali dengan nama reviewer yang berbeda.")
    df = build_excel_export()
    if df.empty:
        st.info("Belum ada data yang dapat diekspor.")
        return
    st.dataframe(df, use_container_width=True, hide_index=True)
    excel_bytes = dataframe_to_excel_bytes(df)
    filename = f"Monitoring_All_{date.today().isoformat()}.xlsx"
    st.download_button("Unduh Excel Monitoring All", data=excel_bytes, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")


# =========================================================
# HALAMAN PENGATURAN
# =========================================================

def page_settings():
    st.subheader("Pengaturan")

    st.markdown("#### Folder Utama Editorial")
    current_root_folder = get_setting("root_folder", "")
    root_folder = st.text_input("Path Folder Utama", value=current_root_folder, placeholder=r"Contoh: C:\Editorial_Journals atau folder OneDrive/SharePoint")

    if st.button("Simpan Pengaturan Folder", type="primary"):
        if root_folder.strip():
            try:
                os.makedirs(os.path.abspath(root_folder.strip()), exist_ok=True)
                save_setting("root_folder", root_folder.strip())
                st.success("Folder utama berhasil disimpan.")
            except Exception as error:
                st.error(f"Folder tidak dapat digunakan: {error}")
        else:
            save_setting("root_folder", "")
            st.warning("Path folder dikosongkan.")

    st.markdown("#### Informasi OJS")
    ojs_rows = [{"Kode": code, "Jurnal": config["name"], "URL OJS": config["ojs_url"]} for code, config in JOURNALS.items()]
    st.dataframe(pd.DataFrame(ojs_rows), use_container_width=True, hide_index=True)

    st.markdown("#### Lokasi Database")
    st.code(os.path.abspath(DB_FILE))

    st.markdown("#### Folder Penyimpanan Default")
    st.code(get_default_storage_root())

    st.warning("Lakukan backup berkala terhadap file `editorial_assistant.db`. File tersebut menyimpan seluruh artikel, reviewer, deadline, dan log komunikasi.")


# =========================================================
# MAIN APP
# =========================================================

def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="📚", layout="wide", initial_sidebar_state="expanded")
    init_db()

    st.title("📚 Editorial Assistant")
    st.caption("Meja kerja lokal untuk pengelolaan CT, HM, dan LC.")

    with st.sidebar:
        st.markdown("## Menu")
        menu = st.radio("Navigasi", ["Dashboard", "Tambah Artikel", "Kelola Artikel", "Impor Dokumen", "Generator Email", "Ekspor Excel", "Pengaturan"])
        st.markdown("---")
        st.markdown("### Jurnal")
        for code, config in JOURNALS.items():
            st.markdown(f"**{code}** — {config['name']}")
        st.markdown("---")
        st.caption("Aplikasi ini adalah asisten administrasi. OJS dan Excel resmi tetap menjadi data operasional utama.")

    if menu == "Dashboard":
        page_dashboard()
    elif menu == "Tambah Artikel":
        page_add_article()
    elif menu == "Kelola Artikel":
        page_manage_articles()
    elif menu == "Impor Dokumen":
        page_import_documents()
    elif menu == "Generator Email":
        page_email_drafts()
    elif menu == "Ekspor Excel":
        page_export()
    elif menu == "Pengaturan":
        page_settings()

if __name__ == "__main__":
    main()