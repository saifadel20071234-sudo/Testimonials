"""
import_names.py
==============
سكريبت استيراد الأسماء من ملف Excel رئيسي وتوزيعها على ملفات المؤتمرات

طريقة الاستخدام:
    py -3.13 import_names.py <اسم_الملف.xlsx>

شكل الملف المطلوب:
    عمود A → اسم الطالب
    عمود B → اسم المؤتمر (أو جزء منه)

مثال:
    | الاسم              | المؤتمر                     |
    |--------------------|------------------------------|
    | أحمد محمد علي     | الذكاء الاصطناعي            |
    | فاطمة إبراهيم     | المحاسبة                    |
"""

import sys
import os
import openpyxl
import re

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# ─────────────────────────────────────────
#  تعريف المؤتمرات وكلمات التعرف عليها
# ─────────────────────────────────────────
CONFERENCES = {
    'ai_emergency': {
        'file': os.path.join(os.path.dirname(__file__), 'data', 'ai_emergency.xlsx'),
        'display': 'مؤتمر: الذكاء الاصطناعي في الطوارئ والإسعافات الأولية',
        # أي كلمة من دي تظهر في اسم المؤتمر → يتعرف عليه
        'keywords': ['ذكاء', 'اصطناعي', 'طوارئ', 'اسعاف', 'إسعاف', 'ai', 'artificial'],
    },
    'accounting': {
        'file': os.path.join(os.path.dirname(__file__), 'data', 'accounting.xlsx'),
        'display': 'مؤتمر: المحاسبة والإدارة المالية',
        'keywords': ['محاسب', 'ادار', 'إدار', 'مالي', 'مال', 'accounting', 'finance'],
    },
}


# ─────────────────────────────────────────
#  تعرف على المؤتمر من النص
# ─────────────────────────────────────────
def detect_conference(text: str) -> str | None:
    """يحاول يعرف المؤتمر المقصود من النص، يعيد conference_id أو None"""
    if not text:
        return None
    text_lower = text.strip().lower()

    for conf_id, conf in CONFERENCES.items():
        for keyword in conf['keywords']:
            if keyword.lower() in text_lower:
                return conf_id
    return None


# ─────────────────────────────────────────
#  قراءة الأسماء الموجودة في ملف المؤتمر
# ─────────────────────────────────────────
def load_existing_names(filepath: str) -> set:
    """يحمّل الأسماء الموجودة في الملف لتجنب التكرار"""
    existing = set()
    if not os.path.exists(filepath):
        return existing
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True)
        ws = wb.active
        for row in ws.iter_rows():
            cell = row[0]
            if cell.value:
                existing.add(str(cell.value).strip())
        wb.close()
    except Exception as e:
        print(f"  ⚠️ تحذير: تعذّر قراءة {filepath}: {e}")
    return existing


# ─────────────────────────────────────────
#  إضافة الأسماء لملف المؤتمر
# ─────────────────────────────────────────
def append_names_to_file(filepath: str, names: list[str], display_name: str):
    """يضيف الأسماء الجديدة لملف الإكسيل"""
    # لو الملف مش موجود، نعمله من الصفر
    if not os.path.exists(filepath):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws['A1'] = 'الاسم'
        print(f"  📄 تم إنشاء ملف جديد: {os.path.basename(filepath)}")
    else:
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active

    existing = load_existing_names(filepath)
    added = 0
    skipped = 0

    for name in names:
        name = name.strip()
        if not name:
            continue
        if name in existing:
            skipped += 1
            continue
        ws.append([name])
        existing.add(name)
        added += 1

    wb.save(filepath)
    wb.close()

    print(f"\n  ✅ {display_name}")
    print(f"     ➕ أُضيف: {added} اسم")
    if skipped > 0:
        print(f"     ⏭️  تخطّي (مكرر): {skipped} اسم")


# ─────────────────────────────────────────
#  الدالة الرئيسية
# ─────────────────────────────────────────
def import_from_excel(input_file: str):
    """يقرأ الملف الرئيسي ويوزع الأسماء"""

    if not os.path.exists(input_file):
        print(f"❌ الملف غير موجود: {input_file}")
        sys.exit(1)

    print(f"\n📂 جاري قراءة: {os.path.basename(input_file)}")
    print("=" * 55)

    try:
        wb = openpyxl.load_workbook(input_file, read_only=True)
        ws = wb.active
    except Exception as e:
        print(f"❌ تعذّر فتح الملف: {e}")
        sys.exit(1)

    # توزيع الأسماء: conference_id → list of names
    distribution: dict[str, list[str]] = {k: [] for k in CONFERENCES}
    unknown_conf = []
    header_skipped = False
    total_rows = 0

    for row in ws.iter_rows():
        # العمود A = الاسم، العمود B = اسم المؤتمر
        name_cell = row[0].value if len(row) > 0 else None
        conf_cell = row[1].value if len(row) > 1 else None

        if not name_cell:
            continue

        name = str(name_cell).strip()

        # تخطي الـ header (السطر الأول لو فيه كلمة اسم)
        if not header_skipped:
            if any(kw in name.lower() for kw in ['اسم', 'name', 'الاسم']):
                header_skipped = True
                continue
            header_skipped = True  # مافيش header — ابدأ مباشرة

        # لو مافيش عمود ب، يطلب التوضيح
        if conf_cell is None:
            unknown_conf.append(name)
            continue

        conf_text = str(conf_cell).strip()
        conf_id   = detect_conference(conf_text)

        if conf_id:
            distribution[conf_id].append(name)
        else:
            unknown_conf.append(f"{name} (مؤتمر: {conf_text})")

        total_rows += 1

    wb.close()

    print(f"📊 إجمالي الصفوف المقروءة: {total_rows}")

    # ── إضافة الأسماء لكل مؤتمر ──
    for conf_id, names in distribution.items():
        if names:
            conf = CONFERENCES[conf_id]
            append_names_to_file(conf['file'], names, conf['display'])

    # ── أسماء غير معروفة ──
    if unknown_conf:
        print(f"\n⚠️  أسماء لم يُتعرَّف على مؤتمرها ({len(unknown_conf)}):")
        for item in unknown_conf:
            print(f"    • {item}")
        print("\n  👉 تأكد إن عمود B فيه اسم المؤتمر بشكل واضح.")

    print("\n" + "=" * 55)
    print("✅ انتهى الاستيراد بنجاح!\n")


# ─────────────────────────────────────────
#  تشغيل السكريبت
# ─────────────────────────────────────────
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("\n❌ يرجى تحديد مسار الملف:")
        print("   py -3.13 import_names.py <اسم_الملف.xlsx>\n")
        sys.exit(1)

    import_from_excel(sys.argv[1])
