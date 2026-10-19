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
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from fuzzywuzzy import fuzz, process

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max upload

# ─────────────────────────────────────────
#  CONFIG — سهل التعديل لاحقاً
# ─────────────────────────────────────────
EXCEL_PATH         = os.path.join(os.path.dirname(__file__), 'data', 'students.xlsx')
CERT_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'static', 'images', 'certificate_template.png')
LOGO_PATH          = os.path.join(os.path.dirname(__file__), 'static', 'images', 'logo.png')
OUTPUT_DIR         = os.path.join(os.path.dirname(__file__), 'generated_certs')

# موضع اسم الطالب على الشهادة (X، Y) — عدّل حسب تصميمك
NAME_POSITION      = (0.5, 0.52)   # نسبة من عرض/ارتفاع الصورة (المنتصف أفقياً + 52% رأسياً)
NAME_FONT_SIZE     = 80            # حجم الخط
NAME_FONT_COLOR    = (0, 0, 0)     # أسود — عدّل حسب تصميمك

FUZZY_THRESHOLD    = 70            # نسبة التشابه المطلوبة (0-100)
STUDENTS_NAME_COL  = 'A'          # عمود الاسم في الإكسيل


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
    # توحيد التاء المربوطة → ة  (نتركها هي) لكن نطابقها مع الهاء
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
def load_students() -> dict:
    """يحمّل الطلاب من ملف الإكسيل ويعيد قاموساً: normalized_name → original_name"""
    students = {}
    if not os.path.exists(EXCEL_PATH):
        return students
    
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True)
        ws = wb.active
        
        # نتخطى السطر الأول لو كان header (يحتوي على كلمة "اسم" أو "الاسم")
        first_row = True
        for row in ws.iter_rows():
            cell = row[0]  # العمود A
            if cell.value is None:
                continue
            
            name = str(cell.value).strip()
            
            # تخطي الـ header
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
    """
    يبحث عن الاسم بذكاء ويعيد:
    - found: True/False
    - original_name: الاسم الأصلي في الملف
    - score: نسبة التشابه
    - suggestions: اقتراحات لو مافيش تطابق كامل
    """
    if not query or not students:
        return {'found': False, 'original_name': None, 'score': 0, 'suggestions': []}

    normalized_query = normalize_arabic(query)
    normalized_list  = list(students.keys())
    
    # أولاً: تطابق تام (بعد التوحيد)
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
            scorer=fuzz.token_sort_ratio,  # أنسب scorer للأسماء العربية
            limit=3
        )
        
        top_match, top_score = results[0][0], results[0][1]
        
        if top_score >= FUZZY_THRESHOLD:
            # تطابق كافٍ — نقترح الاسم للتأكيد
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
#  Certificate Generation
#  توليد الشهادة
# ─────────────────────────────────────────
def generate_certificate(student_name: str) -> bytes:
    """
    يولّد ملف PDF للشهادة مع اسم الطالب.
    لو في template صورة يستخدمها، لو لأ يعمل placeholder جميل.
    """
    import arabic_reshaper
    from bidi.algorithm import get_display

    # ─── تجهيز النص العربي للعرض الصحيح ───
    reshaped_name = arabic_reshaper.reshape(student_name)
    bidi_name     = get_display(reshaped_name)

    # ─── تحديد إذا في template ───
    has_template = os.path.exists(CERT_TEMPLATE_PATH)
    
    # ─── بناء الصورة ───
    if has_template:
        img = Image.open(CERT_TEMPLATE_PATH).convert('RGBA')
    else:
        img = _create_placeholder_certificate()
    
    draw = ImageDraw.Draw(img)
    img_w, img_h = img.size

    # ─── تحميل الخط ───
    font = _load_arabic_font(NAME_FONT_SIZE)

    # ─── حساب موضع الاسم ───
    name_x = int(img_w * NAME_POSITION[0])
    name_y = int(img_h * NAME_POSITION[1])

    # ─── رسم الاسم في المنتصف ───
    bbox = draw.textbbox((0, 0), bidi_name, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    draw.text(
        (name_x - text_w // 2, name_y - text_h // 2),
        bidi_name,
        font=font,
        fill=NAME_FONT_COLOR
    )

    # ─── تحويل لـ PDF مباشرة في الذاكرة بدون لمس القرص ───
    img_rgb = img.convert('RGB')
    pdf_buffer = BytesIO()
    img_rgb.save(pdf_buffer, format='PDF', resolution=150.0)
    pdf_buffer.seek(0)
    return pdf_buffer.read()



def _create_placeholder_certificate() -> Image.Image:
    """ينشئ شهادة placeholder جميلة لحد ما يجي التصميم الحقيقي"""
    # حجم A4 أفقي بدقة عالية (300 DPI)
    w, h = 3508, 2480  # A4 landscape at 300 DPI
    img = Image.new('RGBA', (w, h), (255, 255, 255, 255))
    draw = ImageDraw.Draw(img)
    
    # خلفية متدرجة كريمية
    for y in range(h):
        ratio = y / h
        r = int(255 - ratio * 15)
        g = int(248 - ratio * 10)
        b = int(220 - ratio * 20)
        draw.line([(0, y), (w, y)], fill=(r, g, b, 255))
    
    # إطار خارجي ذهبي
    border = 60
    gold  = (184, 134, 11)
    gold2 = (218, 165, 32)
    draw.rectangle([border,     border,     w-border,     h-border    ], outline=gold,  width=8)
    draw.rectangle([border+20,  border+20,  w-border-20,  h-border-20 ], outline=gold2, width=3)
    
    # نقاط زينة في الزوايا
    corner_size = 80
    for cx, cy in [(border+10, border+10), (w-border-10, border+10),
                   (border+10, h-border-10), (w-border-10, h-border-10)]:
        draw.ellipse([cx-corner_size//2, cy-corner_size//2,
                      cx+corner_size//2, cy+corner_size//2],
                     fill=gold, outline=gold2, width=3)
    
    # خط فاصل أفقي تحت العنوان
    draw.line([(border+100, h//3 + 40), (w-border-100, h//3 + 40)], fill=gold, width=4)
    draw.line([(border+200, h//3 + 50), (w-border-200, h//3 + 50)], fill=gold2, width=2)
    
    return img


def _load_arabic_font(size: int):
    """يحمّل خط عربي بالحجم المطلوب"""
    font_paths = [
        # Windows Arabic fonts
        r'C:\Windows\Fonts\arial.ttf',
        r'C:\Windows\Fonts\arialbd.ttf',
        r'C:\Windows\Fonts\tahoma.ttf',
        r'C:\Windows\Fonts\trebucbd.ttf',
        r'C:\Windows\Fonts\times.ttf',
        # Amiri / Cairo (لو متثبتة)
        r'C:\Windows\Fonts\Amiri-Regular.ttf',
        r'C:\Windows\Fonts\Cairo-Regular.ttf',
    ]
    for path in font_paths:
        if os.path.exists(path):
            try:
                from PIL import ImageFont
                return ImageFont.truetype(path, size)
            except Exception:
                continue
    # fallback
    from PIL import ImageFont
    return ImageFont.load_default()


# ─────────────────────────────────────────
#  Flask Routes
# ─────────────────────────────────────────

@app.route('/')
def index():
    """الصفحة الرئيسية — البحث عن الاسم"""
    return render_template('index.html')


@app.route('/search', methods=['POST'])
def search():
    """API: البحث عن اسم الطالب"""
    data  = request.get_json(silent=True) or {}
    query = data.get('name', '').strip()
    
    if not query:
        return jsonify({'status': 'error', 'message': 'يرجى إدخال الاسم'}), 400
    
    students = load_students()
    result   = find_student(query, students)
    
    if result['found']:
        # لو التشابه أقل من 100% — نطلب تأكيد
        if result['score'] < 100 and result['suggestions']:
            return jsonify({
                'status': 'confirm',
                'suggestions': result['suggestions'],
                'entered_name': query
            })
        # تطابق كامل أو واضح
        return jsonify({
            'status': 'found',
            'original_name': result['original_name'],
            'entered_name': query
        })
    else:
        return jsonify({
            'status': 'not_found',
            'message': 'لم يتم العثور على اسمك. تأكد من كتابته بشكل صحيح أو تواصل مع المسؤول.'
        })


@app.route('/preview', methods=['POST'])
def preview():
    """
    API: معاينة الشهادة — يعيد الشهادة كـ PNG للعرض في المتصفح
    """
    data = request.get_json(silent=True) or {}
    name = data.get('name', '').strip()
    
    if not name:
        return jsonify({'status': 'error', 'message': 'اسم غير صحيح'}), 400

    try:
        import arabic_reshaper
        from bidi.algorithm import get_display

        # بناء الصورة للمعاينة
        has_template = os.path.exists(CERT_TEMPLATE_PATH)
        img = Image.open(CERT_TEMPLATE_PATH).convert('RGBA') if has_template else _create_placeholder_certificate()

        draw  = ImageDraw.Draw(img)
        img_w, img_h = img.size

        reshaped_name = arabic_reshaper.reshape(name)
        bidi_name     = get_display(reshaped_name)

        font   = _load_arabic_font(NAME_FONT_SIZE)
        name_x = int(img_w * NAME_POSITION[0])
        name_y = int(img_h * NAME_POSITION[1])

        bbox   = draw.textbbox((0, 0), bidi_name, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        draw.text(
            (name_x - text_w // 2, name_y - text_h // 2),
            bidi_name,
            font=font,
            fill=NAME_FONT_COLOR
        )

        # تصغير للمعاينة (جودة عرض لا طباعة)
        preview_img = img.convert('RGB')
        preview_img.thumbnail((1200, 850), Image.LANCZOS)

        buf = BytesIO()
        preview_img.save(buf, 'JPEG', quality=88)
        buf.seek(0)

        import base64
        img_b64 = base64.b64encode(buf.read()).decode('utf-8')
        return jsonify({'status': 'ok', 'image': f'data:image/jpeg;base64,{img_b64}'})

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/download', methods=['POST'])
def download():
    """API: تحميل الشهادة كـ PDF"""
    data = request.get_json(silent=True) or {}
    name = data.get('name', '').strip()
    
    if not name:
        return jsonify({'status': 'error', 'message': 'اسم غير صحيح'}), 400
    
    # تحقق أن الاسم موجود في الملف فعلاً
    students = load_students()
    result   = find_student(name, students)
    if not result['found']:
        return jsonify({'status': 'error', 'message': 'الاسم غير موجود'}), 403
    
    official_name = result['original_name']
    
    try:
        pdf_bytes = generate_certificate(official_name)
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
    # إنشاء البيانات التجريبية لو مافيش
    if not os.path.exists(EXCEL_PATH):
        print("Creating sample data...")
        import subprocess
        subprocess.run(['py', '-3.13', 'create_sample_data.py'],
                       cwd=os.path.dirname(__file__), check=False)
    
    print("Starting Certificate System on http://localhost:5000 ...")
    app.run(debug=True, host='0.0.0.0', port=5000)
