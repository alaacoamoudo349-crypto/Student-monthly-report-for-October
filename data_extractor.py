import pandas as pd
import json
import numpy as np
import os

# ==============================================================================
# الإعدادات: يجب تعديل هذه المتغيرات عند تغيير الملف
# ==============================================================================

# اسم ملف الإكسيل الذي سيتم معالجته.
# تأكد من أن هذا الملف موجود في نفس المجلد الذي يحتوي على السكريبت.
EXCEL_FILE_NAME = 'التقريرالشهرى.xlsx' 

# اسم ملف JSON الناتج.
JSON_OUTPUT_FILE = 'report_data_corrected.json'

# أسماء أوراق العمل (Sheets) التي تحتوي على بيانات الطلاب.
SHEET_NAMES = ['GRADE 5', 'GRADE 6', 'GRADE 7', 'GRADE 8']

# ==============================================================================
# هيكل البيانات: لا تقم بتعديل هذا الجزء إلا إذا تغير تنسيق ملف الإكسيل
# ==============================================================================

# خريطة المواد والترجمة الإنجليزية (للاستخدام في كود الويب)
SUBJECT_COLUMNS_MAP = [
    ('التربية الإسلامية', 'islamic_education'),
    ('اللغة العربية', 'arabic_language'),
    ('اللغة الانجليزية', 'english_language'),
    ('الدراسات الاجتماعية', 'social_studies'),
    ('الرياضيات', 'mathematics'),
    ('العلوم', 'science'),
    ('التربية البدنية', 'physical_education'),
    ('الفنون', 'arts'),
    ('الموسيقى', 'music'),
    ('التصميم والتكنولوجيا', 'design_technology'),
    ('اللغة الفرنسية', 'french_language'),
    ('اللغة الألمانية', 'german_language')
]

# عدد الأعمدة الفرعية لكل مادة (5 أعمدة)
SUB_COLUMNS_COUNT = 5

# مؤشرات الأعمدة الرئيسية (0-based index)
NATIONAL_ID_COL_IDX = 0  # رقم الهوية الوطنية (العمود الأول)
STUDENT_ID_COL_IDX = 1   # رقم الطالب (العمود الثاني)
STUDENT_NAME_COL_IDX = 2 # اسم الطالب (العمود الثالث)
CLASS_COL_IDX = 3        # الصف (العمود الرابع)
BEHAVIOR_COL_IDX = 4     # السلوك (العمود الخامس)

# مؤشر بداية أعمدة الدرجات (العمود السادس)
GRADES_START_COL_IDX = 5

# ==============================================================================
# الدالة الرئيسية لاستخراج البيانات
# ==============================================================================

def extract_data_to_json(file_path):
    """
    يقرأ ملف الإكسيل ويستخرج بيانات الطلاب إلى ملف JSON بهدف البحث المزدوج.
    """
    all_students_data = {}
    
    if not os.path.exists(file_path):
        print(f"خطأ: لم يتم العثور على ملف الإكسيل بالاسم '{EXCEL_FILE_NAME}'.")
        print("يرجى التأكد من تسمية الملف بشكل صحيح ووضعه في نفس المجلد.")
        return

    print(f"بدء معالجة ملف الإكسيل: {file_path}")

    for sheet_name in SHEET_NAMES:
        try:
            print(f"  - معالجة ورقة العمل: {sheet_name}...")
            
            # قراءة ورقة العمل كبيانات خام (بدون رأس)
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            
            # تحديد الصف الذي تبدأ منه البيانات (الصف الرابع في الإكسيل، أي index 3 في pandas)
            data_start_row = 3 
            
            # تصفية الصفوف التي لا تحتوي على رقم طالب أو رقم هوية وطنية
            # نستخدم الأعمدة بناءً على index 0-based
            df_data = df.iloc[data_start_row:].copy()
            
            # التكرار على صفوف البيانات (الطلاب)
            for index, row in df_data.iterrows():
                try:
                    student_id_raw = row[STUDENT_ID_COL_IDX]
                    national_id_raw = row[NATIONAL_ID_COL_IDX]
                    
                    # تنظيف رقم الطالب (تحويله إلى عدد صحيح ثم إلى نص)
                    student_id = str(int(student_id_raw)) if pd.notna(student_id_raw) and str(student_id_raw).replace('.', '', 1).isdigit() else None
                    
                    # تنظيف رقم الهوية الوطنية (إزالة .0 إذا وجدت)
                    national_id = str(national_id_raw).strip()
                    if national_id.endswith('.0'):
                        national_id = national_id[:-2]
                    
                    # التحقق من صحة البيانات الأساسية
                    if not student_id or not national_id or national_id == 'nan':
                        continue
                        
                    # المفتاح المزدوج للبحث: "رقم الطالب_رقم الهوية الوطنية"
                    combined_key = f"{student_id}_{national_id}"
                        
                    student_data = {
                        'student_name': str(row[STUDENT_NAME_COL_IDX]).strip(),
                        'class_name': str(row[CLASS_COL_IDX]).strip(),
                        'general_behavior': str(row[BEHAVIOR_COL_IDX]).strip(),
                        'grades': {}
                    }
                    
                    # استخراج الدرجات
                    current_col_idx = GRADES_START_COL_IDX
                    for ar_subject, en_subject in SUBJECT_COLUMNS_MAP:
                        subject_grades = {}
                        
                        # التحقق من حدود الأعمدة
                        if current_col_idx + SUB_COLUMNS_COUNT > len(row):
                            break 
                            
                        # استخراج الأعمدة الفرعية الخمسة للمادة الحالية
                        sub_column_keys = ['formative_exam', 'academic_level', 'participation', 'doing_tasks', 'attending_books']
                        
                        for i in range(SUB_COLUMNS_COUNT):
                            value = str(row[current_col_idx + i]).strip() if pd.notna(row[current_col_idx + i]) else 'N/A'
                            subject_grades[sub_column_keys[i]] = value
                        
                        # إضافة المادة فقط إذا كانت تحتوي على درجات فعلية
                        if any(v != 'N/A' for v in subject_grades.values()):
                            student_data['grades'][en_subject] = subject_grades
                        
                        # الانتقال إلى بداية المادة التالية
                        current_col_idx += SUB_COLUMNS_COUNT

                    all_students_data[combined_key] = student_data
                
                except Exception as e:
                    # طباعة خطأ في صف معين للمساعدة في التصحيح
                    print(f"    [خطأ في الصف] حدث خطأ أثناء معالجة الصف رقم {index + 1}: {e}")
                    continue

        except Exception as e:
            print(f"  [خطأ عام] حدث خطأ أثناء معالجة ورقة {sheet_name}: {e}")

    # حفظ البيانات في ملف JSON
    with open(JSON_OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(all_students_data, f, ensure_ascii=False, indent=4)

    print("\n" + "="*50)
    print(f"✓ تم الانتهاء من معالجة البيانات بنجاح.")
    print(f"✓ إجمالي عدد الطلاب المستخرجين: {len(all_students_data)}")
    print(f"✓ تم حفظ ملف البيانات الجديد في: {JSON_OUTPUT_FILE}")
    print("="*50)

# تنفيذ الدالة
if __name__ == "__main__":
    extract_data_to_json(EXCEL_FILE_NAME)
