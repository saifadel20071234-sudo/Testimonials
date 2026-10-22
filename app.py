"""
Certificate Generation System - Main Flask Application
نظام إصدار الشهادات الإلكترونية
"""

import os
import re
import unicodedata
from io import BytesIO
from flask import Flask, render_template, request, jsonify, send_file
import openpyxl
from PIL import Image, ImageDraw, ImageFont
from fuzzywuzzy import fuzz, process

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max upload

# ─────────────────────────────────────────
#  CONFIG — إعدادات المؤتمرات
# ─────────────────────────────────────────
BASE_DIR = os.path.dirname(__file__)

CONFERENCES = {
    'ai_emergency': {
        'name': 'مؤتمر: الذكاء الاصطناعي في الطوارئ والإسعافات الأولية',
        'excel': os.path.join(BASE_DIR, 'data', 'ai_emergency.xlsx'),
    },
    'accounting': {
        'name': 'مؤتمر: المحاسبة والإدارة المالية',
        'excel': os.path.join(BASE_DIR, 'data', 'accounting.xlsx'),
    },
}

CERT_TEMPLATE_PATH = os.path.join(BASE_DIR, 'static', 'images', 'certificate_template.png')
LOGO_PATH          = os.path.join(BASE_DIR, 'static', 'images', 'logo.png')
FONTS_DIR          = os.path.join(BASE_DIR, 'static', 'fonts')

FUZZY_THRESHOLD    = 70


# ─────────────────────────────────────────
#  Arabic Text Normalization
#  توحيد الحروف العربية
# ─────────────────────────────────────────
def normalize_arabic(text: str) -> str:
    """
    يوحّد الكتابة العربية:
    - يزيل الهمزات والمدة من الألف
    - يوحّد التاء المربوطة والهاء
    - يوحّد الياء والألف المقصورة
    - يزيل التشكيل والتنوين
    - يزيل المسافات الزائدة
    """
    if not text:
        return ''
    text = text.strip()
    # إزالة التشكيل (الحركات والتنوين)
    text = re.sub(r'[\u064B-\u065F\u0670]', '', text)
    # توحيد الألف وأشكالها (أ إ آ أ) → ا
    text = re.sub(r'[أإآٱ]', 'ا', text)
    # توحيد التاء المربوطة مع الهاء
    text = re.sub(r'[ةه]', 'ه', text)
    # توحيد الياء والألف المقصورة
    text = re.sub(r'[يى]', 'ي', text)
    # توحيد الواو مع همزتها
    text = re.sub(r'[ؤو]', 'و', text)
    # إزالة المسافات المتعددة
    text = re.sub(r'\s+', ' ', text).strip()
    return text


# ─────────────────────────────────────────
#  Excel Loading
#  تحميل بيانات الطلاب من الإكسيل
# ─────────────────────────────────────────
def load_students(conference_id: str) -> dict:
    """يحمّل الطلاب من ملف الإكسيل الخاص بالمؤتمر ويعيد قاموساً: normalized_name → original_name"""
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
            cell = row[0]  # العمود A
            if cell.value is None:
                continue

            name = str(cell.value).strip()

            if first_row and ('اسم' in name or 'الاسم' in name or 'name' in name.lower()):
                first_row = False
                continue
            first_row = False

            if name:
                normalized = normalize_arabic(name)
                students[normalized] = name  # normalized → original

        wb.close()
    except Exception as e:
        print(f"❌ خطأ في قراءة الإكسيل: {e}")

    return students


# ─────────────────────────────────────────
#  Smart Name Search
#  البحث الذكي عن الاسم
# ─────────────────────────────────────────
def find_student(query: str, students: dict) -> dict:
    if not query or not students:
        return {'found': False, 'original_name': None, 'score': 0, 'suggestions': []}

    normalized_query = normalize_arabic(query)
    normalized_list  = list(students.keys())

    # أولاً: تطابق تام
    if normalized_query in students:
        return {
            'found': True,
            'original_name': students[normalized_query],
            'score': 100,
            'suggestions': []
        }

    # ثانياً: Fuzzy Matching
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
    """يحمّل خط عربي بالحجم المطلوب، مع إعطاء الأولوية للخطوط المرفقة في المشروع"""
    preferred_names = ['arialbd.ttf', 'tahomabd.ttf', 'timesbd.ttf', 'arial.ttf', 'tahoma.ttf'] if bold else ['arial.ttf', 'tahoma.ttf']

    # 1. الخطوط المرفقة مع المشروع
    if os.path.exists(FONTS_DIR):
        for fname in preferred_names:
            p = os.path.join(FONTS_DIR, fname)
            if os.path.exists(p):
                try:
                    return ImageFont.truetype(p, size)
                except Exception:
                    continue

    # 2. خطوط ويندوز كبديل
    win_fonts = [
        r'C:\Windows\Fonts\arialbd.ttf',
        r'C:\Windows\Fonts\tahoma.ttf',
        r'C:\Windows\Fonts\arial.ttf',
        r'C:\Windows\Fonts\times.ttf',
    ]
    for path in win_fonts:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue

    # 3. Fallback
    return ImageFont.load_default()


# ─────────────────────────────────────────
#  Certificate Image Rendering
# ─────────────────────────────────────────
def render_certificate_image(student_name: str, conference_id: str) -> Image.Image:
    """يبني صورة الشهادة كاملة بدمج القالب مع اسم الطالب والمؤتمر"""
    import arabic_reshaper
    from bidi.algorithm import get_display

    conf = CONFERENCES.get(conference_id, CONFERENCES['ai_emergency'])
    conf_title = conf['name']

    has_template = os.path.exists(CERT_TEMPLATE_PATH)
    if has_template:
        img = Image.open(CERT_TEMPLATE_PATH).convert('RGB')
    else:
        img = _create_placeholder_certificate()

    draw = ImageDraw.Draw(img)
    img_w, img_h = img.size
    scale = img_w / 1000.0

    # ── 1. اسم الطالب ──
    name_font_size = max(24, int(32 * scale))
    f_name = _load_arabic_font(name_font_size, bold=True)
    reshaped_name = arabic_reshaper.reshape(student_name)
    bidi_name     = get_display(reshaped_name)

    bbox_name = draw.textbbox((0, 0), bidi_name, font=f_name)
    nw = bbox_name[2] - bbox_name[0]
    nh = bbox_name[3] - bbox_name[1]

    max_name_width = int(580 * scale)
    while nw > max_name_width and name_font_size > int(18 * scale):
        name_font_size -= 2
        f_name = _load_arabic_font(name_font_size, bold=True)
        bbox_name = draw.textbbox((0, 0), bidi_name, font=f_name)
        nw = bbox_name[2] - bbox_name[0]
        nh = bbox_name[3] - bbox_name[1]

    name_center_x = int(img_w * 0.40)
    name_center_y = int(img_h * 0.547)
    draw.text(
        (name_center_x - nw // 2, name_center_y - nh // 2),
        bidi_name,
        font=f_name,
        fill=(12, 37, 72)
    )

    # ── 2. اسم المؤتمر ──
    conf_font_size = max(20, int(25 * scale))
    f_conf = _load_arabic_font(conf_font_size, bold=True)
    reshaped_conf = arabic_reshaper.reshape(conf_title)
    bidi_conf     = get_display(reshaped_conf)

    bbox_conf = draw.textbbox((0, 0), bidi_conf, font=f_conf)
    cw = bbox_conf[2] - bbox_conf[0]
    ch = bbox_conf[3] - bbox_conf[1]

    max_conf_width = int(820 * scale)
    while cw > max_conf_width and conf_font_size > int(16 * scale):
        conf_font_size -= 2
        f_conf = _load_arabic_font(conf_font_size, bold=True)
        bbox_conf = draw.textbbox((0, 0), bidi_conf, font=f_conf)
        cw = bbox_conf[2] - bbox_conf[0]
        ch = bbox_conf[3] - bbox_conf[1]

    conf_center_x = int(img_w * 0.50)
    conf_center_y = int(img_h * 0.709)
    draw.text(
        (conf_center_x - cw // 2, conf_center_y - ch // 2),
        bidi_conf,
        font=f_conf,
        fill=(16, 53, 101)
    )

    return img


# ─────────────────────────────────────────
#  Certificate PDF Generation
# ─────────────────────────────────────────
def generate_certificate(student_name: str, conference_id: str) -> bytes:
    """يولّد ملف PDF للشهادة مع اسم الطالب واسم المؤتمر في الذاكرة"""
    img = render_certificate_image(student_name, conference_id)
    pdf_buffer = BytesIO()
    img.save(pdf_buffer, format='PDF', resolution=150.0)
    pdf_buffer.seek(0)
    return pdf_buffer.read()


def _create_placeholder_certificate() -> Image.Image:
    """fallback في حال عدم وجود القالب"""
    w, h = 3508, 2480
    img = Image.new('RGB', (w, h), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    for y in range(h):
        ratio = y / h
        r = int(255 - ratio * 15)
        g = int(248 - ratio * 10)
        b = int(220 - ratio * 20)
        draw.line([(0, y), (w, y)], fill=(r, g, b))
    gold  = (184, 134, 11)
    gold2 = (218, 165, 32)
    border = 60
    draw.rectangle([border, border, w-border, h-border], outline=gold, width=8)
    draw.rectangle([border+20, border+20, w-border-20, h-border-20], outline=gold2, width=3)
    return img


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
    conference_id = data.get('conference', 'ai_emergency').strip()

    if not name:
        return jsonify({'status': 'error', 'message': 'اسم غير صحيح'}), 400

    try:
        img = render_certificate_image(name, conference_id)

        # تصغير للمعاينة السريعة
        preview_img = img.copy()
        preview_img.thumbnail((1200, 850), Image.LANCZOS)

        buf = BytesIO()
        preview_img.save(buf, 'JPEG', quality=90)
        buf.seek(0)

        import base64
        img_b64 = base64.b64encode(buf.read()).decode('utf-8')
        return jsonify({'status': 'ok', 'image': f'data:image/jpeg;base64,{img_b64}'})

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/download', methods=['POST'])
def download():
    data          = request.get_json(silent=True) or {}
    name          = data.get('name', '').strip()
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
        pdf_bytes = generate_certificate(official_name, conference_id)
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


if __name__ == '__main__':
    print("Starting Certificate System on http://localhost:5000 ...")
    app.run(debug=True, host='0.0.0.0', port=5000)
