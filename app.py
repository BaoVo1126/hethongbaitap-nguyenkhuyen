import os
import re
import json
import sqlite3
import random
import hashlib
from datetime import datetime, date
import requests 
from flask import Flask, g, render_template, request, redirect, url_for, flash, abort, send_from_directory
from werkzeug.utils import secure_filename
os.environ["FLAGS_use_mkldnn"] = "0"
os.environ["FLAGS_enable_pir_api"] = "0"

try:
    import pymupdf as fitz  
    PDF_AVAILABLE = True
except ImportError:
    try:
        import fitz
        PDF_AVAILABLE = True
    except ImportError:
        PDF_AVAILABLE = False

try:
    from paddleocr import PaddleOCR
    PADDLE_AVAILABLE = True
    _ocr = PaddleOCR(use_textline_orientation=True, lang='en')
except ImportError:
    PADDLE_AVAILABLE = False
    _ocr = None


try:
    import pytesseract
    from PIL import Image
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False



OLLAMA_URL = os.environ.get("OLLAMA_URL", "http://localhost:11434/api/generate")
OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "llama3.2:latest")  


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "center.db")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
DOCS_DIR = os.path.join(UPLOAD_DIR, "documents")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(DOCS_DIR, exist_ok=True)

CENTER_NAME = "Trung Tâm Nguyễn Khuyến"
CENTER_INFO = {
    "tagline": "Hệ thống luyện tập & kiểm tra trực tuyến",
    "subjects": ["Toán", "Hoá", "Lý", "Anh văn"],
    "description": (
        "Trung tâm tổ chức các buổi luyện tập và kiểm tra theo từng chuyên đề. "
        "Học sinh làm bài trực tiếp trên hệ thống, được chấm điểm và xem lại lỗi sai ngay sau khi nộp."
    ),
}

# MẬT KHẨU BẢO VỆ DÀNH CHO GIÁO VIÊN
TEACHER_PASSWORD = os.environ.get("TEACHER_PASSWORD", "123456")

app = Flask(__name__)
app.config["SECRET_KEY"] = "demo-secret-key-doi-khi-deploy-that"
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30MB

# BẬT MẶC ĐỊNH CHẾ ĐỘ XÁO TRỘN CÂU HỎI VÀ ĐÁP ÁN THEO MÃ SỐ HỌC SINH (RANDOMIZE_EXAM = True)
RANDOMIZE_EXAM = True


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    db = sqlite3.connect(DB_PATH)
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS exams (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            subject TEXT,
            source_image TEXT,
            raw_ocr_text TEXT,
            status TEXT NOT NULL DEFAULT 'draft',
            available_date TEXT,
            time_limit_minutes INTEGER,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_id INTEGER NOT NULL REFERENCES exams(id) ON DELETE CASCADE,
            order_index INTEGER NOT NULL,
            question_text TEXT NOT NULL,
            option_a TEXT NOT NULL,
            option_b TEXT NOT NULL,
            option_c TEXT NOT NULL,
            option_d TEXT NOT NULL,
            correct_answer TEXT
        );

        CREATE TABLE IF NOT EXISTS tests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            exam_id INTEGER NOT NULL REFERENCES exams(id) ON DELETE CASCADE,
            student_name TEXT NOT NULL,
            student_code TEXT NOT NULL,
            question_order TEXT NOT NULL,
            option_maps TEXT NOT NULL,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            test_id INTEGER NOT NULL REFERENCES tests(id) ON DELETE CASCADE,
            answers TEXT NOT NULL,
            score INTEGER NOT NULL,
            total INTEGER NOT NULL,
            submitted_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            original_filename TEXT NOT NULL,
            stored_filename TEXT NOT NULL,
            uploaded_at TEXT NOT NULL
        );
        """
    )
    cols = [r[1] for r in db.execute("PRAGMA table_info(exams)").fetchall()]
    if "available_date" not in cols:
        db.execute("ALTER TABLE exams ADD COLUMN available_date TEXT")
    if "time_limit_minutes" not in cols:
        db.execute("ALTER TABLE exams ADD COLUMN time_limit_minutes INTEGER")
    db.commit()
    db.close()


# --------------------------------------------------------------------------
# Pipeline: deterministic document -> exact question extraction
# --------------------------------------------------------------------------


def _clean_extracted_line(line: str) -> str:
    """Remove PDF line noise without changing the actual wording."""
    line = line.replace("\u00a0", " ")
    return re.sub(r"\s+", " ", line).strip()


def _join_wrapped_lines(text: str) -> str:
    """Join visual line wraps while preserving the words/content."""
    lines = [_clean_extracted_line(x) for x in text.splitlines()]
    lines = [x for x in lines if x]
    return " ".join(lines).strip()


def _read_txt(file_path: str) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1258", "cp1252"):
        try:
            with open(file_path, "r", encoding=encoding) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    return ""


def _read_docx(file_path: str) -> str:
    try:
        from docx import Document
    except ImportError:
        return ""

    doc = Document(file_path)
    parts = []
    for p in doc.paragraphs:
        if p.text.strip():
            parts.append(p.text)
    for table in doc.tables:
        for row in table.rows:
            parts.append(" | ".join(cell.text for cell in row.cells))
    return "\n".join(parts)


def _read_pptx(file_path: str) -> str:
    try:
        from pptx import Presentation
    except ImportError:
        return ""

    prs = Presentation(file_path)
    parts = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                parts.append(shape.text)
    return "\n".join(parts)


def _ocr_image_paddle(image) -> str:
    if not PADDLE_AVAILABLE or _ocr is None:
        return ""
    try:
        import numpy as np
        img_np = np.array(image)
        result = _ocr.predict(img_np)

        chunks = []
        for page_result in result or []:
            data = None
            if hasattr(page_result, "json"):
                try:
                    data = page_result.json
                    if callable(data):
                        data = data()
                except Exception:
                    data = None
            if data is None and hasattr(page_result, "res"):
                data = page_result.res
            if isinstance(data, str):
                try:
                    data = json.loads(data)
                except Exception:
                    pass
            if isinstance(data, dict):
                data = data.get("res", data)
                texts = data.get("rec_texts") or data.get("texts") or []
                chunks.extend(str(x) for x in texts if str(x).strip())
            elif isinstance(page_result, dict):
                texts = page_result.get("rec_texts") or page_result.get("texts") or []
                chunks.extend(str(x) for x in texts if str(x).strip())

        if chunks:
            return "\n".join(chunks)
    except Exception as e:
        print(f"[PaddleOCR Warning] {e}")

    return ""


def _ocr_image_tesseract(image) -> str:
    if not OCR_AVAILABLE:
        return ""
    try:
        return pytesseract.image_to_string(image, lang="vie+eng")
    except Exception:
        try:
            return pytesseract.image_to_string(image)
        except Exception:
            return ""


def _pdf_has_real_text(text: str) -> bool:
    if not text or len(text.strip()) < 40:
        return False
    question_hits = len(re.findall(r"(?m)^\s*\d+\.\d+\.\s+", text))
    word_hits = len(re.findall(r"[A-Za-zÀ-ỹ]{3,}", text))
    return question_hits >= 1 or word_hits >= 20


def extract_text_from_file(file_path: str) -> str:
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".txt":
        return _read_txt(file_path)

    if ext == ".docx":
        return _read_docx(file_path)

    if ext == ".pptx":
        return _read_pptx(file_path)

    if ext == ".pdf" and PDF_AVAILABLE:
        doc = fitz.open(file_path)
        native_pages = []
        for page in doc:
            native_pages.append(page.get_text("text", sort=True))
        native_text = "\n".join(native_pages)

        if _pdf_has_real_text(native_text):
            return native_text

        if PDF2IMAGE_AVAILABLE:
            try:
                images = convert_from_path(file_path, dpi=220)
                ocr_pages = []
                for img in images:
                    page_text = _ocr_image_paddle(img)
                    if not page_text:
                        page_text = _ocr_image_tesseract(img)
                    ocr_pages.append(page_text)
                return "\n".join(ocr_pages)
            except Exception as e:
                print(f"[PDF OCR Warning] {e}")
        return native_text

    if ext in {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}:
        if not OCR_AVAILABLE and not PADDLE_AVAILABLE:
            return ""
        try:
            from PIL import Image
            image = Image.open(file_path)
            text = _ocr_image_paddle(image)
            return text or _ocr_image_tesseract(image)
        except Exception as e:
            print(f"[Image OCR Warning] {e}")
            return ""

    return ""


def _parse_option_lines(text: str):
    options = []
    current_letter = None
    current_parts = []

    for raw_line in text.splitlines():
        line = _clean_extracted_line(raw_line)
        if not line:
            continue

        m = re.match(r"^([ABCD])(?:\s+(.*))?$", line)
        if m:
            if current_letter is not None:
                options.append((current_letter, _join_wrapped_lines(" ".join(current_parts))))
            current_letter = m.group(1)
            current_parts = [m.group(2)] if m.group(2) else []
        elif current_letter is not None:
            current_parts.append(line)

    if current_letter is not None:
        options.append((current_letter, _join_wrapped_lines(" ".join(current_parts))))

    return options


def _extract_questions_from_pdf_layout(file_path: str):
    if not PDF_AVAILABLE:
        return []

    doc = fitz.open(file_path)
    questions = []
    current = None

    def flush_current():
        nonlocal current
        if current is None:
            return

        parsed_options = {}
        orphan_text = []

        for block_text in current["option_blocks"]:
            parsed = _parse_option_lines(block_text)
            if parsed:
                for letter, value in parsed:
                    if letter not in parsed_options:
                        parsed_options[letter] = value
                    elif value and not parsed_options[letter]:
                        parsed_options[letter] = value
            else:
                cleaned = _join_wrapped_lines(block_text)
                if cleaned and not _looks_like_header_or_footer(cleaned):
                    orphan_text.append(cleaned)

        if len(parsed_options) == 4 and orphan_text:
            tail = " ".join(orphan_text).strip()
            if tail:
                parsed_options["D"] = (parsed_options.get("D", "") + " " + tail).strip()

        if all(k in parsed_options and parsed_options[k] for k in "ABCD"):
            qtext = _join_wrapped_lines("\n".join(current["question_text"]))
            questions.append({
                "question_text": f"{current['number']}. {qtext}",
                "option_a": parsed_options["A"],
                "option_b": parsed_options["B"],
                "option_c": parsed_options["C"],
                "option_d": parsed_options["D"],
                "correct_answer": None,
                "source_page": current["page"],
            })

        current = None

    for page_number, page in enumerate(doc, start=1):
        blocks = page.get_text("blocks", sort=True)
        for block in blocks:
            text = block[4] or ""
            lines = text.splitlines()
            if not lines:
                continue

            m = re.match(r"^\s*(\d+\.\d+)\.\s*(.*)$", lines[0])
            if m:
                flush_current()
                current = {
                    "number": m.group(1),
                    "question_text": [m.group(2)] + lines[1:],
                    "option_blocks": [],
                    "page": page_number,
                }
                continue

            if current is not None:
                current["option_blocks"].append(text)

    flush_current()
    return questions


def _looks_like_header_or_footer(text: str) -> bool:
    t = text.lower().strip()
    if not t:
        return True
    header_words = (
        "hóa học 11",
        "chương 6:",
        "gv. nguyễn trung kiên",
        "mức độ",
        "trắc nghiệm khách quan",
        "bài tập tự luận",
    )
    if any(x in t for x in header_words):
        return True
    if re.fullmatch(r"[-–—\s]*\d+[-–—\s]*", t):
        return True
    return False


def _parse_questions_from_text_generic(raw_text: str):
    if not raw_text:
        return []

    lines = raw_text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    starts = []
    for i, line in enumerate(lines):
        m = re.match(r"^\s*(\d+\.\d+)\.\s*(.*)$", line)
        if m:
            starts.append((i, m.group(1), m.group(2)))

    questions = []
    for idx, (start, number, first_line) in enumerate(starts):
        end = starts[idx + 1][0] if idx + 1 < len(starts) else len(lines)
        block = lines[start:end]
        qtext_lines = [first_line]
        option_source = []
        seen_option = False

        for line in block[1:]:
            if re.match(r"^\s*[ABCD]\s+", line):
                seen_option = True
            if seen_option:
                option_source.append(line)
            else:
                qtext_lines.append(line)

        parsed_options = {}
        for letter, value in _parse_option_lines("\n".join(option_source)):
            if letter not in parsed_options:
                parsed_options[letter] = value

        if all(k in parsed_options and parsed_options[k] for k in "ABCD"):
            questions.append({
                "question_text": f"{number}. {_join_wrapped_lines(' '.join(qtext_lines))}",
                "option_a": parsed_options["A"],
                "option_b": parsed_options["B"],
                "option_c": parsed_options["C"],
                "option_d": parsed_options["D"],
                "correct_answer": None,
                "source_page": None,
            })

    return questions


def _detect_answer_key(raw_text: str):
    if not raw_text:
        return {}
    answers = {}
    patterns = [
        r"(?:Đáp\s*án|Dap\s*an|Answer)\s*(?:câu\s*)?(\d+(?:\.\d+)?)\s*[:.)-]?\s*([ABCD])\b",
        r"(?:Câu\s*)?(\d+(?:\.\d+)?)\s*[:.)-]\s*(?:Đáp\s*án|Dap\s*an|Answer)\s*[:.)-]?\s*([ABCD])\b",
    ]
    for pattern in patterns:
        for m in re.finditer(pattern, raw_text, flags=re.IGNORECASE):
            answers[m.group(1)] = m.group(2).upper()
    return answers


def parse_questions_pipeline(raw_text: str, file_path: str = None):
    if not raw_text or not raw_text.strip():
        return []

    questions = []
    if file_path and os.path.splitext(file_path)[1].lower() == ".pdf" and PDF_AVAILABLE:
        questions = _extract_questions_from_pdf_layout(file_path)

    if not questions:
        questions = _parse_questions_from_text_generic(raw_text)

    explicit_answers = _detect_answer_key(raw_text)
    for q in questions:
        number_match = re.match(r"^(\d+\.\d+)\.", q["question_text"])
        number = number_match.group(1) if number_match else None
        q["correct_answer"] = explicit_answers.get(number)

    return questions


def parse_questions_from_text_regex(raw_text: str):
    return _parse_questions_from_text_generic(raw_text)

# --------------------------------------------------------------------------
# Randomizer & Grading
# --------------------------------------------------------------------------

def seeded_random(exam_id: int, student_code: str) -> random.Random:
    """Tạo seed ngẫu nhiên duy nhất dựa trên Mã số học sinh và Đề thi."""
    key = f"{exam_id}:{student_code.strip().lower()}".encode("utf-8")
    seed = int(hashlib.sha256(key).hexdigest(), 16) % (2**32)
    return random.Random(seed)


def generate_test(exam_id: int, student_name: str, student_code: str, num_questions: int = None):
    """Tạo đề thi ngẫu nhiên hóa câu hỏi & đáp án A/B/C/D dựa trên MSSV."""
    db = get_db()
    all_questions = db.execute(
        "SELECT * FROM questions WHERE exam_id = ? ORDER BY order_index", (exam_id,)
    ).fetchall()
    if not all_questions:
        return None

    # Khởi tạo thuật toán tráo ngẫu nhiên dựa theo MSSV
    rng = seeded_random(exam_id, student_code)
    q_ids = [q["id"] for q in all_questions]

    # 1. Xáo trộn thứ tự CÂU HỎI dựa trên MSSV
    if RANDOMIZE_EXAM:
        rng.shuffle(q_ids)

    if num_questions:
        q_ids = q_ids[:num_questions]

    # 2. Xáo trộn thứ tự CÁC ĐÁP ÁN (A, B, C, D) dựa trên MSSV
    option_maps = {}
    for q in all_questions:
        if q["id"] not in q_ids:
            continue
        letters = ["A", "B", "C", "D"]
        if RANDOMIZE_EXAM:
            shuffled = letters[:]
            rng.shuffle(shuffled)
        else:
            shuffled = letters[:]
        option_maps[str(q["id"])] = {display: original for display, original in zip(letters, shuffled)}

    cur = db.execute(
        "INSERT INTO tests (exam_id, student_name, student_code, question_order, option_maps, created_at) "
        "VALUES (?, ?, ?, ?, ?, ?)",
        (
            exam_id,
            student_name,
            student_code,
            json.dumps(q_ids),
            json.dumps(option_maps),
            datetime.now().isoformat(timespec="seconds"),
        ),
    )
    db.commit()
    return cur.lastrowid


def grade_submission(test_row, answers: dict):
    db = get_db()
    q_ids = json.loads(test_row["question_order"])
    option_maps = json.loads(test_row["option_maps"])

    questions = {
        q["id"]: q
        for q in db.execute(
            f"SELECT * FROM questions WHERE id IN ({','.join('?' * len(q_ids))})", q_ids
        ).fetchall()
    }

    results = []
    score = 0
    for qid in q_ids:
        q = questions[qid]
        display_map = option_maps[str(qid)]
        chosen_display = answers.get(str(qid))
        chosen_original = display_map.get(chosen_display) if chosen_display else None
        is_correct = chosen_original is not None and chosen_original == q["correct_answer"]
        if is_correct:
            score += 1

        orig_texts = {
            "A": q["option_a"], "B": q["option_b"], "C": q["option_c"], "D": q["option_d"]
        }
        displayed_options = {display: orig_texts[orig] for display, orig in display_map.items()}
        correct_display = next((d for d, orig in display_map.items() if orig == q["correct_answer"]), None)

        results.append(
            {
                "question_text": q["question_text"],
                "options": displayed_options,
                "chosen": chosen_display,
                "correct_display": correct_display,
                "is_correct": is_correct,
            }
        )

    return score, len(q_ids), results


# --------------------------------------------------------------------------
# Pipeline Step 5 & 6: Database & Web Routes (Học sinh & Admin)
# --------------------------------------------------------------------------

@app.route("/")
def home():
    db = get_db()
    exams = db.execute(
        "SELECT * FROM exams WHERE status = 'published' ORDER BY created_at DESC"
    ).fetchall()
    today = date.today().isoformat()
    today_exams = [e for e in exams if not e["available_date"] or e["available_date"] <= today]
    upcoming_exams = [e for e in exams if e["available_date"] and e["available_date"] > today]
    return render_template(
        "home.html", center_name=CENTER_NAME, today_exams=today_exams, upcoming_exams=upcoming_exams,
        active="kiemtra",
    )


@app.route("/exam/<int:exam_id>")
def exam_entry(exam_id):
    db = get_db()
    exam = db.execute("SELECT * FROM exams WHERE id = ? AND status='published'", (exam_id,)).fetchone()
    if not exam:
        abort(404)
    today = date.today().isoformat()
    if exam["available_date"] and exam["available_date"] > today:
        flash(f"Bài kiểm tra này chưa mở, sẽ mở vào ngày {exam['available_date']}.")
        return redirect(url_for("home"))
    return render_template("entry.html", center_name=CENTER_NAME, exam=exam, active="kiemtra")


@app.route("/exam/<int:exam_id>/start", methods=["POST"])
def exam_start(exam_id):
    db = get_db()
    exam = db.execute("SELECT * FROM exams WHERE id = ? AND status='published'", (exam_id,)).fetchone()
    if not exam:
        abort(404)
    today = date.today().isoformat()
    if exam["available_date"] and exam["available_date"] > today:
        flash(f"Bài kiểm tra này chưa mở, sẽ mở vào ngày {exam['available_date']}.")
        return redirect(url_for("home"))

    student_name = request.form.get("student_name", "").strip()
    student_code = request.form.get("student_code", "").strip()
    if not student_name or not student_code:
        flash("Vui lòng nhập đầy đủ họ tên và mã sinh viên.")
        return redirect(url_for("exam_entry", exam_id=exam_id))

    test_id = generate_test(exam_id, student_name, student_code)
    if test_id is None:
        flash("Đề này chưa có câu hỏi nào, vui lòng báo giáo viên.")
        return redirect(url_for("home"))

    return redirect(url_for("take_test", test_id=test_id))


@app.route("/test/<int:test_id>")
def take_test(test_id):
    db = get_db()
    test = db.execute("SELECT * FROM tests WHERE id = ?", (test_id,)).fetchone()
    if not test:
        abort(404)
    exam = db.execute("SELECT * FROM exams WHERE id = ?", (test["exam_id"],)).fetchone()

    q_ids = json.loads(test["question_order"])
    option_maps = json.loads(test["option_maps"])
    questions = {
        q["id"]: q
        for q in db.execute(
            f"SELECT * FROM questions WHERE id IN ({','.join('?' * len(q_ids))})", q_ids
        ).fetchall()
    }

    ordered_questions = []
    for idx, qid in enumerate(q_ids, start=1):
        q = questions[qid]
        display_map = option_maps[str(qid)]
        orig_texts = {"A": q["option_a"], "B": q["option_b"], "C": q["option_c"], "D": q["option_d"]}
        displayed = [(letter, orig_texts[orig]) for letter, orig in sorted(display_map.items())]
        ordered_questions.append(
            {"index": idx, "id": qid, "question_text": q["question_text"], "options": displayed}
        )

    return render_template(
        "test.html", center_name=CENTER_NAME, exam=exam, test=test, questions=ordered_questions,
        active="kiemtra",
    )


@app.route("/test/<int:test_id>/submit", methods=["POST"])
def submit_test(test_id):
    db = get_db()
    test = db.execute("SELECT * FROM tests WHERE id = ?", (test_id,)).fetchone()
    if not test:
        abort(404)

    q_ids = json.loads(test["question_order"])
    answers = {}
    for qid in q_ids:
        val = request.form.get(f"q_{qid}")
        if val:
            answers[str(qid)] = val

    score, total, results = grade_submission(test, answers)

    db.execute(
        "INSERT INTO submissions (test_id, answers, score, total, submitted_at) VALUES (?, ?, ?, ?, ?)",
        (test_id, json.dumps(answers), score, total, datetime.now().isoformat(timespec="seconds")),
    )
    db.commit()
    submission_id = db.execute("SELECT last_insert_rowid() AS id").fetchone()["id"]

    return redirect(url_for("result", submission_id=submission_id))


@app.route("/result/<int:submission_id>")
def result(submission_id):
    db = get_db()
    submission = db.execute("SELECT * FROM submissions WHERE id = ?", (submission_id,)).fetchone()
    if not submission:
        abort(404)
    test = db.execute("SELECT * FROM tests WHERE id = ?", (submission["test_id"],)).fetchone()
    exam = db.execute("SELECT * FROM exams WHERE id = ?", (test["exam_id"],)).fetchone()

    answers = json.loads(submission["answers"])
    _, _, results = grade_submission(test, answers)

    duration_str = None
    try:
        started = datetime.fromisoformat(test["created_at"])
        finished = datetime.fromisoformat(submission["submitted_at"])
        total_seconds = max(0, int((finished - started).total_seconds()))
        m, s = divmod(total_seconds, 60)
        duration_str = f"{m} phút {s} giây" if m else f"{s} giây"
    except (ValueError, TypeError):
        duration_str = None

    return render_template(
        "result.html",
        center_name=CENTER_NAME,
        exam=exam,
        test=test,
        submission=submission,
        results=results,
        duration_str=duration_str,
        active="kiemtra",
    )


# --------------------------------------------------------------------------
# ROUTE XÁC THỰC MẬT KHẨU GIÁO VIÊN
# --------------------------------------------------------------------------

@app.route("/verify-teacher-password", methods=["POST"])
def verify_teacher_pass():
    password_input = request.form.get("teacher_password", "")
    next_url = request.form.get("next_url") or url_for("admin_home")
    
    if password_input == TEACHER_PASSWORD:
        return redirect(next_url)
    else:
        flash("❌ Mật khẩu giáo viên không chính xác. Vui lòng thử lại!")
        return redirect(request.referrer or url_for("home"))


@app.route("/admin")
def admin_home():
    db = get_db()
    exams = db.execute("SELECT * FROM exams ORDER BY created_at DESC").fetchall()
    exam_stats = {}
    for e in exams:
        n_questions = db.execute("SELECT COUNT(*) c FROM questions WHERE exam_id=?", (e["id"],)).fetchone()["c"]
        n_submissions = db.execute(
            "SELECT COUNT(*) c FROM submissions s JOIN tests t ON s.test_id=t.id WHERE t.exam_id=?", (e["id"],)
        ).fetchone()["c"]
        exam_stats[e["id"]] = {"n_questions": n_questions, "n_submissions": n_submissions}
    return render_template(
        "admin_home.html", center_name=CENTER_NAME, exams=exams, exam_stats=exam_stats,
        ocr_available=(PADDLE_AVAILABLE or OCR_AVAILABLE), pdf_available=PDF_AVAILABLE,
        active="admin", admin_tab="exams",
    )


@app.route("/admin/upload", methods=["POST"])
def admin_upload():
    file = request.files.get("exam_image")
    title = request.form.get("title", "").strip() or "Đề chưa đặt tên"
    subject = request.form.get("subject", "").strip()
    available_date = request.form.get("available_date", "").strip() or None
    time_limit_raw = request.form.get("time_limit_minutes", "").strip()
    time_limit_minutes = int(time_limit_raw) if time_limit_raw.isdigit() and int(time_limit_raw) > 0 else None

    if not file or file.filename == "":
        flash("Vui lòng chọn tệp đề thi (PDF, DOCX, PPTX, TXT hoặc Ảnh).")
        return redirect(url_for("admin_home"))

    allowed_extensions = {".pdf", ".txt", ".docx", ".pptx", ".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in allowed_extensions:
        flash("Định dạng chưa được hỗ trợ. Hãy dùng PDF, TXT, DOCX, PPTX hoặc ảnh.")
        return redirect(url_for("admin_home"))

    filename = secure_filename(file.filename)
    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
    saved_name = f"{stamp}_{filename}"
    saved_path = os.path.join(UPLOAD_DIR, saved_name)
    file.save(saved_path)

    raw_text = extract_text_from_file(saved_path)
    parsed = parse_questions_pipeline(raw_text, saved_path) if raw_text else []

    db = get_db()
    cur = db.execute(
        "INSERT INTO exams (title, subject, source_image, raw_ocr_text, status, available_date, "
        "time_limit_minutes, created_at) VALUES (?, ?, ?, ?, 'draft', ?, ?, ?)",
        (title, subject, saved_name, raw_text, available_date, time_limit_minutes,
         datetime.now().isoformat(timespec="seconds")),
    )
    exam_id = cur.lastrowid

    for idx, q in enumerate(parsed):
        db.execute(
            "INSERT INTO questions (exam_id, order_index, question_text, option_a, option_b, option_c, "
            "option_d, correct_answer) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (
                exam_id, idx, q["question_text"], q["option_a"], q["option_b"], q["option_c"], q["option_d"],
                q["correct_answer"],
            ),
        )
    db.commit()

    if parsed:
        flash(f"Hệ thống đã tự động trích xuất thành công {len(parsed)} câu hỏi từ tệp {filename}!")
    else:
        flash("Không thể tự động nhận diện câu hỏi từ tệp này. Bạn có thể thêm câu hỏi thủ công ở trang bên dưới.")

    return redirect(url_for("admin_edit_exam", exam_id=exam_id))


@app.route("/admin/exam/<int:exam_id>/delete", methods=["POST"])
def admin_delete_exam(exam_id):
    db = get_db()
    exam = db.execute("SELECT * FROM exams WHERE id=?", (exam_id,)).fetchone()
    if exam:
        if exam["source_image"]:
            file_path = os.path.join(UPLOAD_DIR, exam["source_image"])
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except OSError:
                    pass
        db.execute("DELETE FROM exams WHERE id=?", (exam_id,))
        db.commit()
        flash(f"Đã xóa thành công đề thi '{exam['title']}'.")
    return redirect(url_for("admin_home"))


@app.route("/admin/exam/<int:exam_id>/edit")
def admin_edit_exam(exam_id):
    db = get_db()
    exam = db.execute("SELECT * FROM exams WHERE id=?", (exam_id,)).fetchone()
    if not exam:
        abort(404)
    questions = db.execute(
        "SELECT * FROM questions WHERE exam_id=? ORDER BY order_index", (exam_id,)
    ).fetchall()
    return render_template(
        "admin_edit.html", center_name=CENTER_NAME, exam=exam, questions=questions, active="admin", admin_tab="exams",
    )


@app.route("/admin/exam/<int:exam_id>/schedule", methods=["POST"])
def admin_update_schedule(exam_id):
    db = get_db()
    available_date = request.form.get("available_date", "").strip() or None
    db.execute("UPDATE exams SET available_date=? WHERE id=?", (available_date, exam_id))
    db.commit()
    return redirect(url_for("admin_edit_exam", exam_id=exam_id))


@app.route("/admin/exam/<int:exam_id>/time_limit", methods=["POST"])
def admin_update_time_limit(exam_id):
    db = get_db()
    raw = request.form.get("time_limit_minutes", "").strip()
    time_limit_minutes = int(raw) if raw.isdigit() and int(raw) > 0 else None
    db.execute("UPDATE exams SET time_limit_minutes=? WHERE id=?", (time_limit_minutes, exam_id))
    db.commit()
    return redirect(url_for("admin_edit_exam", exam_id=exam_id))


@app.route("/admin/exam/<int:exam_id>/add_question", methods=["POST"])
def admin_add_question(exam_id):
    db = get_db()
    max_idx = db.execute(
        "SELECT COALESCE(MAX(order_index), -1) m FROM questions WHERE exam_id=?", (exam_id,)
    ).fetchone()["m"]
    db.execute(
        "INSERT INTO questions (exam_id, order_index, question_text, option_a, option_b, option_c, "
        "option_d, correct_answer) VALUES (?, ?, '', '', '', '', '', 'A')",
        (exam_id, max_idx + 1),
    )
    db.commit()
    return redirect(url_for("admin_edit_exam", exam_id=exam_id))


@app.route("/admin/question/<int:question_id>/update", methods=["POST"])
def admin_update_question(question_id):
    db = get_db()
    q = db.execute("SELECT * FROM questions WHERE id=?", (question_id,)).fetchone()
    if not q:
        abort(404)
    correct = request.form.get("correct_answer", "").strip().upper() or None
    db.execute(
        "UPDATE questions SET question_text=?, option_a=?, option_b=?, option_c=?, option_d=?, "
        "correct_answer=? WHERE id=?",
        (
            request.form.get("question_text", "").strip(),
            request.form.get("option_a", "").strip(),
            request.form.get("option_b", "").strip(),
            request.form.get("option_c", "").strip(),
            request.form.get("option_d", "").strip(),
            correct,
            question_id,
        ),
    )
    db.commit()
    return redirect(url_for("admin_edit_exam", exam_id=q["exam_id"]))


@app.route("/admin/exam/<int:exam_id>/save_all_questions", methods=["POST"])
def admin_save_all_questions(exam_id):
    db = get_db()
    question_ids = request.form.getlist("question_ids")
    
    for qid in question_ids:
        q_text = request.form.get(f"question_text_{qid}", "").strip()
        opt_a = request.form.get(f"option_a_{qid}", "").strip()
        opt_b = request.form.get(f"option_b_{qid}", "").strip()
        opt_c = request.form.get(f"option_c_{qid}", "").strip()
        opt_d = request.form.get(f"option_d_{qid}", "").strip()
        correct = request.form.get(f"correct_answer_{qid}", "").strip().upper() or None
        
        db.execute(
            """
            UPDATE questions 
            SET question_text=?, option_a=?, option_b=?, option_c=?, option_d=?, correct_answer=? 
            WHERE id=? AND exam_id=?
            """,
            (q_text, opt_a, opt_b, opt_c, opt_d, correct, qid, exam_id)
        )
        
    db.commit()
    flash(f"Đã lưu thành công toàn bộ {len(question_ids)} câu hỏi!")
    return redirect(url_for("admin_edit_exam", exam_id=exam_id))


@app.route("/admin/question/<int:question_id>/delete", methods=["POST"])
def admin_delete_question(question_id):
    db = get_db()
    q = db.execute("SELECT * FROM questions WHERE id=?", (question_id,)).fetchone()
    if not q:
        abort(404)
    db.execute("DELETE FROM questions WHERE id=?", (question_id,))
    db.commit()
    return redirect(url_for("admin_edit_exam", exam_id=q["exam_id"]))


@app.route("/admin/exam/<int:exam_id>/publish", methods=["POST"])
def admin_publish_exam(exam_id):
    db = get_db()
    n = db.execute("SELECT COUNT(*) c FROM questions WHERE exam_id=?", (exam_id,)).fetchone()["c"]
    if n == 0:
        flash("Đề chưa có câu hỏi nào, không thể đăng bài.")
        return redirect(url_for("admin_edit_exam", exam_id=exam_id))
    n_missing = db.execute(
        "SELECT COUNT(*) c FROM questions WHERE exam_id=? AND (correct_answer IS NULL OR correct_answer='')",
        (exam_id,),
    ).fetchone()["c"]
    if n_missing:
        flash(f"Còn {n_missing} câu chưa chọn đáp án đúng — vui lòng điền đủ trước khi đăng.")
        return redirect(url_for("admin_edit_exam", exam_id=exam_id))
    db.execute("UPDATE exams SET status='published' WHERE id=?", (exam_id,))
    db.commit()
    return redirect(url_for("admin_home"))


@app.route("/admin/exam/<int:exam_id>/unpublish", methods=["POST"])
def admin_unpublish_exam(exam_id):
    db = get_db()
    db.execute("UPDATE exams SET status='draft' WHERE id=?", (exam_id,))
    db.commit()
    return redirect(url_for("admin_home"))


@app.route("/admin/exam/<int:exam_id>/submissions")
def admin_exam_submissions(exam_id):
    db = get_db()
    exam = db.execute("SELECT * FROM exams WHERE id=?", (exam_id,)).fetchone()
    if not exam:
        abort(404)
    rows = db.execute(
        """
        SELECT t.student_name, t.student_code, s.score, s.total, s.submitted_at
        FROM submissions s JOIN tests t ON s.test_id = t.id
        WHERE t.exam_id = ?
        ORDER BY s.submitted_at DESC
        """,
        (exam_id,),
    ).fetchall()
    return render_template(
        "admin_submissions.html", center_name=CENTER_NAME, exam=exam, rows=rows, active="admin", admin_tab="exams",
    )


@app.route("/admin/documents")
def admin_documents():
    db = get_db()
    docs = db.execute("SELECT * FROM documents ORDER BY uploaded_at DESC").fetchall()
    return render_template(
        "admin_documents.html", center_name=CENTER_NAME, docs=docs, active="admin", admin_tab="documents",
    )


@app.route("/documents")
def public_documents():
    db = get_db()
    docs = db.execute("SELECT * FROM documents ORDER BY uploaded_at DESC").fetchall()
    return render_template(
        "documents.html", center_name=CENTER_NAME, docs=docs, active="tailieu",
    )


@app.route("/info")
def info_page():
    db = get_db()
    history_rows = db.execute(
        """
        SELECT 
            s.id AS submission_id,
            t.student_name,
            t.student_code,
            e.title AS exam_title,
            e.subject,
            s.score,
            s.total,
            t.created_at AS started_at,
            s.submitted_at
        FROM submissions s
        JOIN tests t ON s.test_id = t.id
        JOIN exams e ON t.exam_id = e.id
        ORDER BY s.submitted_at DESC
        """
    ).fetchall()

    processed_history = []
    for row in history_rows:
        duration_str = "-"
        try:
            started = datetime.fromisoformat(row["started_at"])
            finished = datetime.fromisoformat(row["submitted_at"])
            total_seconds = max(0, int((finished - started).total_seconds()))
            m, s = divmod(total_seconds, 60)
            duration_str = f"{m} phút {s} giây" if m else f"{s} giây"
        except (ValueError, TypeError):
            pass

        processed_history.append({
            "submission_id": row["submission_id"],
            "student_name": row["student_name"],
            "student_code": row["student_code"],
            "exam_title": row["exam_title"],
            "subject": row["subject"],
            "score": row["score"],
            "total": row["total"],
            "submitted_at": row["submitted_at"],
            "duration": duration_str
        })

    return render_template(
        "info.html", 
        center_name=CENTER_NAME, 
        history=processed_history, 
        active="thongtin"
    )

@app.route("/info/submission/<int:submission_id>/delete", methods=["POST"])
def delete_submission(submission_id):
    db = get_db()
    # Tìm thông tin lượt nộp bài
    sub = db.execute("SELECT * FROM submissions WHERE id = ?", (submission_id,)).fetchone()
    if sub:
        # Xóa lượt nộp bài (submission) và bản ghi khởi tạo đề thi tương ứng (tests)
        db.execute("DELETE FROM submissions WHERE id = ?", (submission_id,))
        db.execute("DELETE FROM tests WHERE id = ?", (sub["test_id"],))
        db.commit()
        flash("Đã xóa lượt làm bài của học sinh thành công.")
    return redirect(url_for("info_page"))



@app.route("/admin/documents/upload", methods=["POST"])
def admin_documents_upload():
    file = request.files.get("document_file")
    if not file or file.filename == "":
        flash("Vui lòng chọn tệp tài liệu.")
        return redirect(url_for("admin_documents"))

    filename = secure_filename(file.filename)
    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
    saved_name = f"{stamp}_{filename}"
    file.save(os.path.join(DOCS_DIR, saved_name))

    db = get_db()
    db.execute(
        "INSERT INTO documents (original_filename, stored_filename, uploaded_at) VALUES (?, ?, ?)",
        (filename, saved_name, datetime.now().isoformat(timespec="seconds")),
    )
    db.commit()
    return redirect(url_for("admin_documents"))


@app.route("/admin/documents/<int:doc_id>/delete", methods=["POST"])
def admin_documents_delete(doc_id):
    db = get_db()
    doc = db.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
    if doc:
        try:
            os.remove(os.path.join(DOCS_DIR, doc["stored_filename"]))
        except OSError:
            pass
        db.execute("DELETE FROM documents WHERE id=?", (doc_id,))
        db.commit()
    return redirect(url_for("admin_documents"))


@app.route("/documents/file/<path:filename>")
def serve_document(filename):
    return send_from_directory(DOCS_DIR, filename, as_attachment=True)


if __name__ == "__main__":
    init_db()
    if PDF_AVAILABLE:
        print("-> Document Parser (PyMuPDF) đã sẵn sàng.")
    if PADDLE_AVAILABLE:
        print("-> PaddleOCR đã sẵn sàng.")
    elif OCR_AVAILABLE:
        print("-> Tesseract OCR (dự phòng) đã sẵn sàng.")
    app.run(debug=True, use_reloader=False, port=5000)
