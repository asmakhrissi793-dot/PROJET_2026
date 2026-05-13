import os
import re
import json
import shutil
import pandas as pd
from datetime import datetime
from tqdm import tqdm

# ------------------------------------------------------------------
# 1. الإعدادات - غير المسار إذا لزم الأمر
# ------------------------------------------------------------------
ROOT_DRIVE = r"C:\Users\SIGMA IT\PFE_PROJET\cv_selection_pfe\backend\001_Recruiting"
UPLOAD_CV_FOLDER = r"uploads/cvs"
os.makedirs(UPLOAD_CV_FOLDER, exist_ok=True)

IGNORE_FOLDERS = ["00_Procedure de recrutement", "001_CV", "002_Clients Share Point",
                  "003_Liste des Clients", "Candidature spontanée", "Stagiaires"]

entreprises = []
postes = []
candidats = []
candidatures = []

comp_counter = 0
poste_counter = 0
cand_counter = 0
appli_counter = 0

def clean_text(text):
    if pd.isna(text) or text is None:
        return ""
    return str(text).strip()

def clean_email(email):
    email = clean_text(email)
    if "@" not in email:
        return ""
    return email

def generate_id(prefix, counter):
    return f"{prefix}_{counter:04d}"

def copy_cv_file(src_path, dest_filename):
    if not os.path.exists(src_path):
        return ""
    safe_name = re.sub(r'[\\/*?:"<>|]', "_", dest_filename)
    if not safe_name.lower().endswith('.pdf'):
        safe_name += ".pdf"
    dest_path = os.path.join(UPLOAD_CV_FOLDER, safe_name)
    shutil.copy2(src_path, dest_path)
    return safe_name

def find_column(df, keywords):
    """البحث عن عمود يحتوي على أي من الكلمات المفتاحية (غير حساس لحالة الأحرف)"""
    for col in df.columns:
        col_lower = str(col).lower()
        for kw in keywords:
            if kw in col_lower:
                return col
    return None

print("🔍 جارٍ مسح المجلد الرئيسي...")
all_items = os.listdir(ROOT_DRIVE)
company_folders = [item for item in all_items if os.path.isdir(os.path.join(ROOT_DRIVE, item)) and item not in IGNORE_FOLDERS]
print(f"✅ تم العثور على {len(company_folders)} شركة: {', '.join(company_folders[:10])}...")

for comp_name in tqdm(company_folders, desc="معالجة الشركات"):
    comp_path = os.path.join(ROOT_DRIVE, comp_name)
    postes_path = os.path.join(comp_path, "Postes")
    if not os.path.isdir(postes_path):
        print(f"  ⚠️ الشركة {comp_name} ليس بها مجلد 'Postes'، سيتم تخطيها.")
        continue

    comp_counter += 1
    company_id = generate_id("COMP", comp_counter)
    entreprises.append({
        "_id": company_id,
        "name": comp_name,
        "secteur": "",
        "location": "",
        "email": "",
        "status": "Active",
        "logo": "assets/default-logo.png"
    })

    job_folders = [f for f in os.listdir(postes_path) if os.path.isdir(os.path.join(postes_path, f))]
    for job_folder in job_folders:
        job_path = os.path.join(postes_path, job_folder)
        job_title = job_folder.strip()

        # البحث عن أي ملف Excel
        excel_files = [f for f in os.listdir(job_path) if f.endswith(('.xlsx', '.xls'))]
        if not excel_files:
            print(f"    ⚠️ المنصب {job_title} لا يحتوي على ملف Excel، تم تخطيه.")
            continue

        excel_path = os.path.join(job_path, excel_files[0])
        df = pd.read_excel(excel_path, header=0)

        # تحديد الأعمدة بالكلمات المفتاحية
        col_candidate = find_column(df, ['candidate', 'nom', 'name', 'candidat'])
        col_email = find_column(df, ['email', 'e-mail', 'mail'])
        col_phone = find_column(df, ['phone', 'téléphone', 'tel'])
        col_salaire = find_column(df, ['required salary', 'salaire demandé', 'prétention'])
        col_residence = find_column(df, ['address', 'adresse', 'residence', 'résidence'])
        col_position = find_column(df, ['position', 'poste', 'job title'])  # قد لا يكون موجوداً

        if not col_candidate:
            print(f"    ⚠️ ملف Excel {excel_files[0]} لا يحتوي على عمود خاص بالمرشحين، تم تخطيه.")
            continue

        # إنشاء كيان الوظيفة (مرة واحدة)
        poste_counter += 1
        job_id = generate_id("JOB", poste_counter)
        postes.append({
            "_id": job_id,
            "title": job_title,
            "description": "",
            "location": "",
            "type": "CDI",
            "keywords": [],
            "companyId": company_id,
            "companyName": comp_name,
            "status": "active",
            "created_at": datetime.now().isoformat(),
            "salary": ""
        })

        # معالجة كل مرشح
        for idx, row in df.iterrows():
            candidate_name = clean_text(row[col_candidate]) if col_candidate else ""
            if not candidate_name:
                continue

            # إذا كان هناك عمود للمنصب في الملف، تحقق من تطابقه مع اسم المجلد (اختياري)
            if col_position:
                pos_in_excel = clean_text(row[col_position])
                if pos_in_excel and pos_in_excel.lower() != job_title.lower():
                    # يمكن تخطي أو الاحتفاظ حسب الحاجة
                    pass

            email = clean_email(row[col_email]) if col_email else ""
            phone = clean_text(row[col_phone]) if col_phone else ""
            salaire = clean_text(row[col_salaire]) if col_salaire else ""
            residence = clean_text(row[col_residence]) if col_residence else ""

            # البحث عن ملف PDF مرتبط بالمرشح
            cv_filename = ""
            sub_folders = [f for f in os.listdir(job_path) if os.path.isdir(os.path.join(job_path, f))]
            candidate_norm = candidate_name.lower().replace(" ", "").replace("-", "")
            for sub in sub_folders:
                sub_path = os.path.join(job_path, sub)
                for file in os.listdir(sub_path):
                    if file.lower().endswith('.pdf'):
                        file_norm = os.path.splitext(file)[0].lower().replace(" ", "").replace("-", "")
                        if candidate_norm in file_norm or file_norm in candidate_norm:
                            src_cv = os.path.join(sub_path, file)
                            cv_filename = copy_cv_file(src_cv, f"{candidate_name}_{job_title}.pdf")
                            break
                if cv_filename:
                    break

            # إضافة المرشح (إذا له بريد إلكتروني أو اسم)
            existing_cand = None
            if email:
                existing_cand = next((c for c in candidats if c.get("email") == email), None)
            if existing_cand:
                candidate_id = existing_cand["_id"]
                if not existing_cand.get("cv_path") and cv_filename:
                    existing_cand["cv_path"] = cv_filename
            else:
                cand_counter += 1
                candidate_id = generate_id("CAND", cand_counter)
                first_name = candidate_name.split()[0] if candidate_name else ""
                last_name = " ".join(candidate_name.split()[1:]) if candidate_name else ""
                candidats.append({
                    "_id": candidate_id,
                    "first_name": first_name,
                    "last_name": last_name,
                    "email": email,
                    "nationalite": "",
                    "password": "hashed_for_real_data",
                    "role": "candidate",
                    "cv_path": cv_filename,
                    "photo_url": "",
                    "created_at": datetime.now().isoformat()
                })

            # إضافة الترشيح
            appli_counter += 1
            candidatures.append({
                "_id": generate_id("APPLI", appli_counter),
                "poste": job_title,
                "company": comp_name,
                "companyId": company_id,
                "posteId": job_id,
                "candidateEmail": email,
                "candidateName": candidate_name,
                "cvPath": cv_filename,
                "status": "pending",
                "score": 0,
                "date": datetime.now().isoformat(),
                "nationalite": "",
                "salaire": salaire,
                "devise": "DT" if salaire else "",
                "residence": residence,
                "langues": []
            })

# تصدير إلى JSON
output_dir = "mongo_import"
os.makedirs(output_dir, exist_ok=True)

with open(os.path.join(output_dir, "entreprises.json"), "w", encoding="utf-8") as f:
    json.dump(entreprises, f, ensure_ascii=False, indent=2)

with open(os.path.join(output_dir, "postes.json"), "w", encoding="utf-8") as f:
    json.dump(postes, f, ensure_ascii=False, indent=2)

with open(os.path.join(output_dir, "candidats.json"), "w", encoding="utf-8") as f:
    json.dump(candidats, f, ensure_ascii=False, indent=2)

with open(os.path.join(output_dir, "candidatures.json"), "w", encoding="utf-8") as f:
    json.dump(candidatures, f, ensure_ascii=False, indent=2)

print(f"\n✅ تم التصدير بنجاح!")
print(f"   📁 المجلد: {output_dir}")
print(f"   🏢 الشركات: {len(entreprises)}")
print(f"   💼 المناصب: {len(postes)}")
print(f"   👥 المرشحون: {len(candidats)}")
print(f"   📄 الترشيحات: {len(candidatures)}")
print(f"   📎 تم نسخ السير الذاتية إلى: {UPLOAD_CV_FOLDER}")