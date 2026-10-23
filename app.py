"""
Certificate Generation System - Main Flask Application
نظام إصدار الشهادات الإلكترونية
"""

import os
import re
import json
import hashlib
from io import BytesIO
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for, session
import openpyxl
from PIL import Image, ImageDraw, ImageFont
import qrcode
import arabic_reshaper
from bidi.algorithm import get_display
from fuzzywuzzy import fuzz, process

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'knu-certificates-secret-key-2026')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max upload

ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'knu2026admin')
BASE_DIR = os.path.dirname(__file__)

# ─────────────────────────────────────────
#  CONFIG — إعدادات المؤتمرات
# ─────────────────────────────────────────
CONFERENCES = {
    'ai_emergency': {
        'name': 'مؤتمر الذكاء الاصطناعي في الطوارئ والإسعافات الأولية',
        'display': 'مؤتمر: الذكاء الاصطناعي في الطوارئ والإسعافات الأولية',
        'slogan': '« الذكاء الاصطناعي في الطوارئ والإسعافات الأولية - إنقاذ الأرواح بالتقنيات الذكية »',
        'date': '19 أكتوبر 2026',
        'code': '01',
        'excel': os.path.join(BASE_DIR, 'data', 'ai_emergency.xlsx'),
        'conf_color': (10, 48, 92),      # كحلي سيبراني عميق
        'slogan_color': (3, 105, 161),   # أزرق تقني هادئ
    },
    'accounting': {
        'name': 'مؤتمر المحاسبة والإدارة المالية',
        'display': 'مؤتمر: المحاسبة والإدارة المالية',
        'slogan': '« آفاق المحاسبة والإدارة المالية الرقمية ودورها في التنمية المستدامة »',
        'date': '22 أكتوبر 2026',
        'code': '02',
        'excel': os.path.join(BASE_DIR, 'data', 'accounting.xlsx'),
        'conf_color': (12, 37, 72),      # كحلي رسمي كلاسيكي
        'slogan_color': (180, 83, 9),    # ذهبي كهرماني ملكي
    },
}

CERT_BASE_PATH = os.path.join(BASE_DIR, 'static', 'images', 'certificate_base.jpg')
FONTS_DIR      = os.path.join(BASE_DIR, 'static', 'fonts')
FUZZY_THRESHOLD = 70

# ─────────────────────────────────────────
#  قائمة الكليات الرسمية بجامعة كفر الشيخ الأهلية
# ─────────────────────────────────────────
OFFICIAL_FACULTIES = [
    'كلية الطب البشري',
    'كلية طب الفم والأسنان',
    'كلية العلاج الطبيعي',
    'كلية الصيدلة',
    'كلية التمريض',
    'كلية الطب البيطري',
    'كلية الزراعة',
    'كلية الهندسة',
    'كلية الحاسبات والمعلومات والذكاء الاصطناعي',
    'كلية العلوم',
    'كلية اللغات والعلوم الإنسانية',
    'كلية الأعمال',
    'كلية الفنون والتصميم',
]

FACULTY_KEYWORD_MAP = {
    'حاسب': 'كلية الحاسبات والمعلومات والذكاء الاصطناعي',
    'ذكاء': 'كلية الحاسبات والمعلومات والذكاء الاصطناعي',
    'معلومات': 'كلية الحاسبات والمعلومات والذكاء الاصطناعي',
    'كمبيوتر': 'كلية الحاسبات والمعلومات والذكاء الاصطناعي',
    'اسنان': 'كلية طب الفم والأسنان',
    'أسنان': 'كلية طب الفم والأسنان',
    'فم': 'كلية طب الفم والأسنان',
    'بشري': 'كلية الطب البشري',
    'بشرى': 'كلية الطب البشري',
    'علاج': 'كلية العلاج الطبيعي',
    'صيدل': 'كلية الصيدلة',
    'تمريض': 'كلية التمريض',
    'بيطر': 'كلية الطب البيطري',
    'زراع': 'كلية الزراعة',
    'هندس': 'كلية الهندسة',
    'علوم': 'كلية العلوم',
    'لغات': 'كلية اللغات والعلوم الإنسانية',
    'لسان': 'كلية اللغات والعلوم الإنسانية',
    'انسان': 'كلية اللغات والعلوم الإنسانية',
    'إنسان': 'كلية اللغات والعلوم الإنسانية',
    'أداب': 'كلية اللغات والعلوم الإنسانية',
    'اداب': 'كلية اللغات والعلوم الإنسانية',
    'أعمال': 'كلية الأعمال',
    'اعمال': 'كلية الأعمال',
    'تجار': 'كلية الأعمال',
    'بزنس': 'كلية الأعمال',
    'ادارة': 'كلية الأعمال',
    'إدارة': 'كلية الأعمال',
    'فنون': 'كلية الفنون والتصميم',
    'تصميم': 'كلية الفنون والتصميم',
}


def match_faculty(query: str) -> str:
    """يتعرف بذكاء على أقرب كلية رسمية من الكليات المعتمدة"""
    if not query:
        return 'جامعة كفر الشيخ الأهلية'
    q = query.strip()

    if q in OFFICIAL_FACULTIES:
        return q
    if f'كلية {q}' in OFFICIAL_FACULTIES:
        return f'كلية {q}'

    # البحث بالكلمات الدلالية المفتاحية
    q_norm = re.sub(r'[أإآٱ]', 'ا', q)
    q_norm = re.sub(r'[ةه]', 'ه', q_norm)
    for kw, target in FACULTY_KEYWORD_MAP.items():
        kw_norm = re.sub(r'[أإآٱ]', 'ا', kw)
        kw_norm = re.sub(r'[ةه]', 'ه', kw_norm)
        if kw_norm in q_norm:
            return target

    # مطابقة تقريبية بالتشابه الذكي (Fuzzy Matching)
    match, score = process.extractOne(q, OFFICIAL_FACULTIES, scorer=fuzz.token_set_ratio)
    if score >= 50:
        return match

    return f"كلية {q}" if not q.startswith('كلية') else q

# قاعدة بيانات خفيفة في الذاكرة لتسجيل الشهادات للتحقق الفوري
REGISTRY_FILE = os.path.join(BASE_DIR, 'data', 'registry.json')


# ─────────────────────────────────────────
#  Certificate ID Generation & Registry
# ─────────────────────────────────────────
def generate_cert_id(student_name: str, conference_id: str) -> str:
    """يولّد رقم توثيق فريد وثابت لكل طالب ومؤتمر: KNU-HONOR-2026-XXX"""
    conf = CONFERENCES.get(conference_id, CONFERENCES['ai_emergency'])
    code = conf['code']
    h = int(hashlib.md5(f"{student_name}_{conference_id}".encode('utf-8')).hexdigest(), 16)
    seq = (h % 900) + 100
    return f"KNU-HONOR-2026-{seq}"


def save_certificate_record(cert_id: str, student_name: str, faculty: str, conference_id: str):
    """يحفظ بيانات الشهادة للتحقق منها عبر الباركود"""
    try:
        data = {}
        if os.path.exists(REGISTRY_FILE):
            with open(REGISTRY_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
        data[cert_id] = {
            'student_name': student_name,
            'faculty': faculty or 'جامعة كفر الشيخ الأهلية',
            'conference_id': conference_id,
        }
        os.makedirs(os.path.dirname(REGISTRY_FILE), exist_ok=True)
        with open(REGISTRY_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Warning: could not save registry record: {e}")


def get_certificate_record(cert_id: str) -> dict | None:
    """يسترجع بيانات الشهادة برقم التوثيق"""
    if os.path.exists(REGISTRY_FILE):
        try:
            with open(REGISTRY_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if cert_id in data:
                    return data[cert_id]
        except Exception:
            pass
    return None


# ─────────────────────────────────────────
#  Arabic Text Normalization
# ─────────────────────────────────────────
def normalize_arabic(text: str) -> str:
    if not text:
        return ''
    text = text.strip()
    text = re.sub(r'[\u064B-\u065F\u0670]', '', text)
    text = re.sub(r'[أإآٱ]', 'ا', text)
    text = re.sub(r'[ةه]', 'ه', text)
    text = re.sub(r'[يى]', 'ي', text)
    text = re.sub(r'[ؤو]', 'و', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


# ─────────────────────────────────────────
#  Excel Loading & Fuzzy Search
# ─────────────────────────────────────────
_STUDENTS_CACHE = {}

def load_students(conference_id: str, force_reload: bool = False) -> dict:
    if not force_reload and conference_id in _STUDENTS_CACHE:
        return _STUDENTS_CACHE[conference_id]

    students = {}
    conf = CONFERENCES.get(conference_id)
    if not conf:
        return students

    excel_path = conf['excel']
    if not os.path.exists(excel_path):
        return students

    try:
        wb = openpyxl.load_workbook(excel_path, read_only=True)
        ws = wb.active

        first_row = True
        for row in ws.iter_rows():
            cell = row[0]
            if cell.value is None:
                continue

            name = str(cell.value).strip()

            if first_row and ('اسم' in name or 'الاسم' in name or 'name' in name.lower()):
                first_row = False
                continue
            first_row = False

            if name:
                normalized = normalize_arabic(name)
                students[normalized] = name

        wb.close()
    except Exception as e:
        print(f"❌ خطأ في قراءة الإكسيل: {e}")

    _STUDENTS_CACHE[conference_id] = students
    return students


def find_student(query: str, students: dict) -> dict:
    if not query or not students:
        return {'found': False, 'original_name': None, 'score': 0, 'suggestions': []}

    normalized_query = normalize_arabic(query)
    normalized_list  = list(students.keys())

    if normalized_query in students:
        return {
            'found': True,
            'original_name': students[normalized_query],
            'score': 100,
            'suggestions': []
        }

    if normalized_list:
        results = process.extract(
            normalized_query,
            normalized_list,
            scorer=fuzz.token_sort_ratio,
            limit=3
        )
        top_match, top_score = results[0][0], results[0][1]
        if top_score >= FUZZY_THRESHOLD:
            original = students[top_match]
            suggestions = [
                {'name': students[r[0]], 'score': r[1]}
                for r in results
                if r[1] >= FUZZY_THRESHOLD
            ]
            return {
                'found': True,
                'original_name': original,
                'score': top_score,
                'suggestions': suggestions if top_score < 100 else []
            }

    return {'found': False, 'original_name': None, 'score': 0, 'suggestions': []}


# ─────────────────────────────────────────
#  Font Loading
# ─────────────────────────────────────────
def _load_arabic_font(size: int, bold: bool = True):
    preferred_names = ['arialbd.ttf', 'tahomabd.ttf', 'timesbd.ttf', 'arial.ttf'] if bold else ['arial.ttf', 'tahoma.ttf']
    if os.path.exists(FONTS_DIR):
        for fname in preferred_names:
            p = os.path.join(FONTS_DIR, fname)
            if os.path.exists(p):
                try:
                    return ImageFont.truetype(p, size)
                except Exception:
                    continue

    win_fonts = [r'C:\Windows\Fonts\arialbd.ttf', r'C:\Windows\Fonts\tahoma.ttf', r'C:\Windows\Fonts\arial.ttf']
    for path in win_fonts:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue

    return ImageFont.load_default()


# ─────────────────────────────────────────
#  Full High-Res Certificate Rendering
# ─────────────────────────────────────────
def render_certificate_image(student_name: str, faculty: str, conference_id: str) -> Image.Image:
    """
    يبني صورة الشهادة كاملة بدقة فائقة مطابقة للشهادة الرسمية
    مع الباركود الفعلي ورقم التوثيق واسم الطالب وكليته واسم المؤتمر
    """
    conf = CONFERENCES.get(conference_id, CONFERENCES['ai_emergency'])
    conf_name   = conf['name']
    conf_slogan = conf['slogan']

    # التعرف الذكي على أقرب كلية رسمية من قائمة الجامعة
    fac_text = match_faculty(faculty)

    # رقم التوثيق الفريد
    cert_id = generate_cert_id(student_name, conference_id)
    save_certificate_record(cert_id, student_name, fac_text, conference_id)

    # تحميل قالب الشهادة النقي عالي الدقة المستخرج من الـ PDF
    if os.path.exists(CERT_BASE_PATH):
        img = Image.open(CERT_BASE_PATH).convert('RGB')
    else:
        # Fallback لو القالب مش موجود
        img = Image.new('RGB', (2480, 1754), (255, 255, 255))

    draw = ImageDraw.Draw(img)
    w, h = img.size

    # تحميل الخطوط
    f_name    = _load_arabic_font(68, bold=True)
    f_faculty = _load_arabic_font(36, bold=True)
    f_apprec  = _load_arabic_font(32, bold=True)
    f_conf    = _load_arabic_font(56, bold=True)
    f_slogan  = _load_arabic_font(30, bold=True)
    f_wish    = _load_arabic_font(28, bold=True)
    f_serial  = _load_arabic_font(22, bold=True)

    C_NAVY_DARK  = (12, 37, 72)
    C_BLUE       = (16, 53, 101)
    C_RED_ACCENT = (195, 40, 40)
    C_GRAY       = (100, 110, 120)

    def draw_centered(text: str, y: int, font, fill, max_w: int = 2100):
        reshaped = arabic_reshaper.reshape(text)
        bidi = get_display(reshaped)
        bbox = draw.textbbox((0, 0), bidi, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]

        # تصغير الخط تلقائياً لو النص أطول من المساحة المسموحة
        cur_font = font
        while tw > max_w and getattr(cur_font, 'size', 30) > 18:
            new_size = cur_font.size - 2
            cur_font = _load_arabic_font(new_size, bold=True)
            bbox = draw.textbbox((0, 0), bidi, font=cur_font)
            tw = bbox[2] - bbox[0]
            th = bbox[3] - bbox[1]

        draw.text((w // 2 - tw // 2, y - th // 2), bidi, font=cur_font, fill=fill)

    # 1. اسم الطالب (مكتوب بوضوح: الطالب / [الاسم])
    student_display = f"الطالب / {student_name}"
    draw_centered(student_display, 615, f_name, C_NAVY_DARK, max_w=1900)

    # 2. الكلية التي كتبها الطالب
    draw_centered(fac_text, 690, f_faculty, C_BLUE, max_w=1800)

    # 3. صيغة التقدير والمشاركة
    draw_centered('تقديراً لمشاركته الفعالة وجهوده المتميزة في خدمة المنظومة الطلابية وإنجاح فعاليات:', 765, f_apprec, C_GRAY, max_w=2000)

    # 4. اسم المؤتمر
    c_conf   = conf.get('conf_color', C_NAVY_DARK)
    c_slogan = conf.get('slogan_color', C_RED_ACCENT)
    draw_centered(conf_name, 840, f_conf, c_conf, max_w=2000)

    # 5. شعار المؤتمر
    draw_centered(conf_slogan, 915, f_slogan, c_slogan, max_w=2000)

    # 6. الخاتمة
    draw_centered('متمنين له دوام التوفيق والتميز والنجاح المستمر', 980, f_wish, C_GRAY, max_w=2000)

    # 7. توليد رمز الاستجابة السريعة (QR Code) بدقة عالية ومسافات متناسقة
    qr_url = f"https://testimonials-virid.vercel.app/verify/{cert_id}"
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=4,
        border=1
    )
    qr.add_data(qr_url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color='#0c2548', back_color='white').convert('RGB')
    qr_img = qr_img.resize((150, 150), Image.NEAREST)
    qr_w, qr_h = qr_img.size

    qr_x = w // 2 - qr_w // 2
    qr_y = 1315
    img.paste(qr_img, (qr_x, qr_y))

    # 8. بيانات التوثيق أسفل الباركود بتباعد مريح وبدون أي تداخل
    draw_centered('امسح الباركود للتحقق من صحة الشهادة', 1492, f_serial, C_NAVY_DARK)
    draw_centered(f'رقم التوثيق: {cert_id}', 1522, f_serial, C_BLUE)
    draw_centered('• شهادة رسمية معتمدة ومسجلة بسجلات الجامعة •', 1552, f_serial, (22, 130, 60))

    return img


def generate_certificate(student_name: str, faculty: str, conference_id: str) -> bytes:
    """توليد ملف PDF في الذاكرة بالكامل"""
    img = render_certificate_image(student_name, faculty, conference_id)
    pdf_buffer = BytesIO()
    img.save(pdf_buffer, format='PDF', resolution=150.0)
    pdf_buffer.seek(0)
    return pdf_buffer.read()


# ─────────────────────────────────────────
#  Flask Routes
# ─────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html', conferences=CONFERENCES)


@app.route('/search', methods=['POST'])
def search():
    data          = request.get_json(silent=True) or {}
    query         = data.get('name', '').strip()
    conference_id = data.get('conference', '').strip()

    if not query:
        return jsonify({'status': 'error', 'message': 'يرجى إدخال الاسم'}), 400

    if conference_id not in CONFERENCES:
        return jsonify({'status': 'error', 'message': 'مؤتمر غير معروف'}), 400

    students = load_students(conference_id)
    result   = find_student(query, students)

    if result['found']:
        if result['score'] < 100 and result['suggestions']:
            return jsonify({
                'status': 'confirm',
                'suggestions': result['suggestions'],
                'entered_name': query
            })
        return jsonify({
            'status': 'found',
            'original_name': result['original_name'],
            'entered_name': query
        })
    else:
        return jsonify({
            'status': 'not_found',
            'message': 'لم يتم العثور على اسمك في هذا المؤتمر. تأكد من كتابته بشكل صحيح أو اختيار المؤتمر الصحيح.'
        })


@app.route('/preview', methods=['POST'])
def preview():
    data          = request.get_json(silent=True) or {}
    name          = data.get('name', '').strip()
    faculty       = data.get('faculty', '').strip()
    conference_id = data.get('conference', 'ai_emergency').strip()

    if not name:
        return jsonify({'status': 'error', 'message': 'اسم غير صحيح'}), 400

    try:
        img = render_certificate_image(name, faculty, conference_id)
        cert_id = generate_cert_id(name, conference_id)
        verify_url = f"https://testimonials-virid.vercel.app/verify/{cert_id}"
        conf = CONFERENCES.get(conference_id, CONFERENCES['ai_emergency'])
        conf_name = conf['name']

        preview_img = img.copy()
        preview_img.thumbnail((1200, 850), Image.LANCZOS)

        buf = BytesIO()
        preview_img.save(buf, 'JPEG', quality=90)
        buf.seek(0)

        import base64
        img_b64 = base64.b64encode(buf.read()).decode('utf-8')
        return jsonify({
            'status': 'ok',
            'image': f'data:image/jpeg;base64,{img_b64}',
            'cert_id': cert_id,
            'verify_url': verify_url,
            'conf_name': conf_name
        })

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/download', methods=['POST'])
def download():
    data          = request.get_json(silent=True) or {}
    name          = data.get('name', '').strip()
    faculty       = data.get('faculty', '').strip()
    conference_id = data.get('conference', 'ai_emergency').strip()

    if not name:
        return jsonify({'status': 'error', 'message': 'اسم غير صحيح'}), 400

    if conference_id not in CONFERENCES:
        return jsonify({'status': 'error', 'message': 'مؤتمر غير معروف'}), 400

    students = load_students(conference_id)
    result   = find_student(name, students)
    if not result['found']:
        return jsonify({'status': 'error', 'message': 'الاسم غير موجود'}), 403

    official_name = result['original_name']

    try:
        pdf_bytes = generate_certificate(official_name, faculty, conference_id)
        safe_name = re.sub(r'[^\w\s-]', '', official_name).strip().replace(' ', '_')
        filename  = f'شهادة_{safe_name}.pdf'

        return send_file(
            BytesIO(pdf_bytes),
            mimetype='application/pdf',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


# ─────────────────────────────────────────
#  Verification Route — صفحة التوثيق بالباركود
# ─────────────────────────────────────────
@app.route('/verify/<cert_id>')
def verify_certificate(cert_id):
    """
    صفحة التحقق الإلكتروني الذكي (QR Verified)
    تفتح عند مسح الباركود بالكاميرا أو الضغط على الرابط
    """
    record = get_certificate_record(cert_id)

    # لو غير موجود في الـ registry، نستخرج البيانات من معلمات الرابط Query Params
    student_name  = request.args.get('name')
    faculty       = request.args.get('fac')
    conference_id = request.args.get('conf', 'ai_emergency')

    if record:
        student_name  = record.get('student_name', student_name or 'طالب بجامعة كفر الشيخ الأهلية')
        faculty       = record.get('faculty', faculty or 'جامعة كفر الشيخ الأهلية')
        conference_id = record.get('conference_id', conference_id)

    if not student_name:
        student_name = "مشارك معتمد"
    if not faculty:
        faculty = "جامعة كفر الشيخ الأهلية"

    conf = CONFERENCES.get(conference_id, CONFERENCES['ai_emergency'])
    conf_name = conf['name']
    conf_date = conf['date']

    download_url = url_for('download_by_cert', cert_id=cert_id, name=student_name, fac=faculty, conf=conference_id)

    return render_template(
        'verify.html',
        cert_id=cert_id,
        student_name=student_name,
        faculty=faculty,
        conf_name=conf_name,
        conf_date=conf_date,
        download_url=download_url
    )


@app.route('/download_cert/<cert_id>')
def download_by_cert(cert_id):
    """تحميل الشهادة مباشرة من صفحة التحقق"""
    student_name  = request.args.get('name', 'طالب')
    faculty       = request.args.get('fac', 'جامعة كفر الشيخ الأهلية')
    conference_id = request.args.get('conf', 'ai_emergency')

    record = get_certificate_record(cert_id)
    if record:
        student_name  = record.get('student_name', student_name)
        faculty       = record.get('faculty', faculty)
        conference_id = record.get('conference_id', conference_id)

    pdf_bytes = generate_certificate(student_name, faculty, conference_id)
    safe_name = re.sub(r'[^\w\s-]', '', student_name).strip().replace(' ', '_')
    filename  = f'شهادة_{safe_name}.pdf'

    return send_file(
        BytesIO(pdf_bytes),
        mimetype='application/pdf',
        as_attachment=True,
        download_name=filename
    )


# ─────────────────────────────────────────
#  Admin Dashboard Routes — لوحة التحكم للمسؤول
# ─────────────────────────────────────────
@app.route('/admin', methods=['GET', 'POST'])
def admin_dashboard():
    error = None
    if request.method == 'POST':
        pwd = request.form.get('password', '').strip()
        if pwd == ADMIN_PASSWORD:
            session['is_admin'] = True
            return redirect(url_for('admin_dashboard'))
        else:
            error = 'كلمة المرور غير صحيحة، يرجى المحاولة مجدداً.'

    if not session.get('is_admin'):
        return render_template('admin.html', is_authenticated=False, error=error)

    # إحصائيات وقوائم الطلاب
    students_ai = load_students('ai_emergency')
    students_acc = load_students('accounting')

    registry_count = 0
    if os.path.exists(REGISTRY_FILE):
        try:
            with open(REGISTRY_FILE, 'r', encoding='utf-8') as f:
                registry_count = len(json.load(f))
        except Exception:
            pass

    return render_template(
        'admin.html',
        is_authenticated=True,
        conferences=CONFERENCES,
        count_ai=len(students_ai),
        count_acc=len(students_acc),
        total_students=len(students_ai) + len(students_acc),
        registry_count=registry_count,
        students_ai=list(students_ai.values()),
        students_acc=list(students_acc.values())
    )


@app.route('/admin/logout')
def admin_logout():
    session.pop('is_admin', None)
    return redirect(url_for('admin_dashboard'))


@app.route('/api/admin/add_student', methods=['POST'])
def admin_add_student():
    if not session.get('is_admin'):
        return jsonify({'status': 'error', 'message': 'غير مصرح لك بالوصول'}), 401

    data = request.get_json(silent=True) or {}
    name = data.get('name', '').strip()
    conf_id = data.get('conference', '').strip()

    if not name or conf_id not in CONFERENCES:
        return jsonify({'status': 'error', 'message': 'يرجى إدخال اسم الطالب واختيار المؤتمر'}), 400

    # حفظ في ملف الإكسيل إن أمكن
    try:
        conf = CONFERENCES[conf_id]
        excel_path = conf['excel']
        if os.path.exists(excel_path):
            wb = openpyxl.load_workbook(excel_path)
            ws = wb.active
            ws.append([name])
            wb.save(excel_path)
            wb.close()
    except Exception as e:
        print(f"Notice: couldn't write to disk excel (normal on serverless): {e}")

    # تحديث الكاش بالذاكرة
    normalized = normalize_arabic(name)
    students = load_students(conf_id)
    students[normalized] = name
    _STUDENTS_CACHE[conf_id] = students

    # تسجيل شهادة الطالب بالـ registry
    cert_id = generate_cert_id(name, conf_id)
    save_certificate_record(cert_id, name, 'جامعة كفر الشيخ الأهلية', conf_id)

    return jsonify({
        'status': 'ok',
        'message': f'تمت إضافة الطالب «{name}» بنجاح!',
        'cert_id': cert_id
    })


@app.route('/api/admin/upload_excel', methods=['POST'])
def admin_upload_excel():
    if not session.get('is_admin'):
        return jsonify({'status': 'error', 'message': 'غير مصرح لك بالوصول'}), 401

    conf_id = request.form.get('conference', '').strip()
    file = request.files.get('file')

    if not file or not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'status': 'error', 'message': 'يرجى اختيار ملف Excel صالح (.xlsx)'}), 400

    if conf_id not in CONFERENCES:
        return jsonify({'status': 'error', 'message': 'مؤتمر غير معروف'}), 400

    try:
        wb = openpyxl.load_workbook(file, read_only=True)
        ws = wb.active
        new_students = {}
        first_row = True
        for row in ws.iter_rows():
            cell = row[0]
            if cell.value is None:
                continue
            name = str(cell.value).strip()
            if first_row and ('اسم' in name or 'الاسم' in name or 'name' in name.lower()):
                first_row = False
                continue
            first_row = False
            if name:
                normalized = normalize_arabic(name)
                new_students[normalized] = name
        wb.close()

        # تحديث الكاش
        _STUDENTS_CACHE[conf_id] = new_students

        # محاولة الحفظ على القرص (إذا كان متاحاً)
        try:
            excel_path = CONFERENCES[conf_id]['excel']
            file.seek(0)
            file.save(excel_path)
        except Exception:
            pass

        return jsonify({
            'status': 'ok',
            'message': f'تم تحديث قائمة الطلاب بنجاح! تم استيراد {len(new_students)} طالب.'
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'خطأ أثناء قراءة الملف: {str(e)}'}), 500


if __name__ == '__main__':
    print("Starting Certificate System on http://localhost:5000 ...")
    app.run(debug=True, host='0.0.0.0', port=5000)
