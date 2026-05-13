import os
import random
import uuid
import smtplib
import hashlib
import re
import io



import pdfkit

from functools import wraps
from datetime import datetime, timedelta

from flask import Flask, current_app, request, jsonify, send_from_directory, render_template_string, send_file
from flask_pymongo import PyMongo
from flask_cors import CORS
from flask_jwt_extended import create_access_token, jwt_required, get_jwt_identity, JWTManager
from flask_socketio import SocketIO, emit
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from bson import ObjectId
from dotenv import load_dotenv
import pdfplumber
import PyPDF2
import spacy
import openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from google import genai
from groq import Groq

from flask import current_app, url_for

# ------------------------------------------------------------------
# Load environment variables FIRST
# ------------------------------------------------------------------
load_dotenv()

# ------------------------------------------------------------------
# Flask App Initialization - MUST BE BEFORE SocketIO
# ------------------------------------------------------------------
app = Flask(__name__)
CORS(app)

app.config["MONGO_URI"] = os.getenv("MONGO_URI", "mongodb://localhost:27017/LCA")
app.config['UPLOAD_FOLDER'] = os.getenv('UPLOAD_FOLDER', 'uploads/cvs')
app.config["JWT_SECRET_KEY"] = os.getenv("JWT_SECRET_KEY", "super-secret-key-change-this-in-production")

# ------------------------------------------------------------------
# Database & JWT
# ------------------------------------------------------------------
mongo = PyMongo(app)
jwt = JWTManager(app)

# ------------------------------------------------------------------
# Gemini Client (Nouvelle API)
# ------------------------------------------------------------------
try:
    client = genai.Client()  # Utilise GOOGLE_API_KEY du .env
    print("✅ Gemini client initialized successfully")
except Exception as e:
    print(f"⚠️ Gemini client initialization error: {e}")
    client = None

# ------------------------------------------------------------------
# wkhtmltopdf Configuration (UNIQUE VERSION)
# ------------------------------------------------------------------
wkhtmltopdf_path = os.getenv("WKHTMLTOPDF_PATH", r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
try:
    if os.path.exists(wkhtmltopdf_path):
        config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
        print(f"✅ wkhtmltopdf configured at: {wkhtmltopdf_path}")
    else:
        config = pdfkit.configuration()
        print("⚠️ Using default wkhtmltopdf path (may not work)")
except Exception as e:
    print(f"⚠️ wkhtmltopdf configuration warning: {e}")
    config = None

# ------------------------------------------------------------------
# SocketIO - MUST BE AFTER app is created
# ------------------------------------------------------------------
socketio = SocketIO(app, cors_allowed_origins="*", async_mode='threading')

# Create necessary folders
for folder in ['uploads/cvs', 'uploads/photos', 'static/uploads/admins']:
    os.makedirs(folder, exist_ok=True)

# Email configuration
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USERNAME = os.getenv("SMTP_USERNAME", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
DEFAULT_FROM_EMAIL = os.getenv("DEFAULT_FROM_EMAIL", "noreply@lca.com")

# NLP
try:
    nlp = spacy.load("en_core_web_sm")
except OSError:
    print("⏳ Downloading spacy model...")
    os.system("python -m spacy download en_core_web_sm")
    nlp = spacy.load("en_core_web_sm")

# Groq client
# Au début du fichier, après les imports
groq_api_key = os.getenv("GROQ_API_KEY")
groq_client = Groq(api_key=groq_api_key) if groq_api_key else None
if groq_api_key:
    groq_client = Groq(api_key=groq_api_key)
    print("✅ Groq client initialized")
else:
    print("⚠️ GROQ_API_KEY not found in .env file")

# ------------------------------------------------------------------
# SocketIO Events
# ------------------------------------------------------------------
@socketio.on('connect')
def handle_connect():
    print('Client connected')

@socketio.on('disconnect')
def handle_disconnect():
    print('Client disconnected')


# ------------------------------------------------------------------
# Utility functions
# ------------------------------------------------------------------
def send_real_email(to_email, subject, html_body):
    """Envoie un email avec HTML correctement formaté"""
    if not SMTP_USERNAME or not SMTP_PASSWORD:
        print("⚠️ SMTP not configured")
        return False
    
    try:
        # Créer le message avec MIMEText en HTML
        msg = MIMEText(html_body, 'html', 'utf-8')
        msg['Subject'] = subject
        msg['From'] = DEFAULT_FROM_EMAIL
        msg['To'] = to_email
        
        # Envoi avec timeout
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as server:
            server.starttls()
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.send_message(msg)
        
        print(f"✅ Email envoyé à {to_email}")
        return True
        
    except smtplib.SMTPAuthenticationError:
        print(f"❌ Erreur d'authentification SMTP - Vérifiez vos identifiants")
        return False
    except smtplib.SMTPException as e:
        print(f"❌ Erreur SMTP: {e}")
        return False
    except Exception as e:
        print(f"❌ Erreur email: {e}")
        return False


def send_email_with_attachment(to_email, subject, html_body, pdf_buffer, filename):
    if not SMTP_USERNAME or not SMTP_PASSWORD:
        print("⚠️ SMTP not configured")
        return False
    try:
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = DEFAULT_FROM_EMAIL
        msg['To'] = to_email
        msg.attach(MIMEText(html_body, 'html'))

        pdf_data = pdf_buffer.read()
        if not pdf_data:
            raise Exception("PDF buffer is empty")
        part = MIMEBase('application', 'pdf')
        part.set_payload(pdf_data)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(part)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USERNAME, SMTP_PASSWORD)
            server.send_message(msg)
        print(f"✅ Email with PDF sent to {to_email}")
        return True
    except Exception as e:
        print(f"❌ Error sending email with PDF: {e}")
        return False

def extract_cv_info_spacy(cv_path):
    text = ""
    with pdfplumber.open(cv_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
    doc = nlp(text.lower())
    skill_keywords = ["python", "javascript", "angular", "react", "node", "mongodb", "sql",
                      "docker", "kubernetes", "machine learning", "ia", "java", "c++", "php",
                      "laravel", "symfony", "django", "flask", "fastapi", "vue", "typescript",
                      "html", "css", "git", "linux", "aws", "azure"]
    found_skills = [token.text for token in doc if token.text in skill_keywords]
    exp_pattern = r"(\d+)\s*(?:ans|années|years|year)"
    exp_matches = re.findall(exp_pattern, text.lower())
    max_exp = max([int(x) for x in exp_matches]) if exp_matches else 0
    diploma_keywords = ["licence", "master", "doctorat", "bachelor", "diplôme", "ingénieur", "phd"]
    found_diplomas = [token.text for token in doc if token.text in diploma_keywords]
    return {
        "skills": list(set(found_skills)),
        "max_experience": max_exp,
        "diplomas": found_diplomas,
        "full_text": text.lower()
    }

def extract_detailed_cv(cv_path):
    text = ""
    with pdfplumber.open(cv_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
    lines = text.split('\n')
    experiences = []
    formations = []
    competences = []
    langues = []
    centres_interet = []

    for i, line in enumerate(lines):
        date_match = re.search(r'(\d{4})\s*[-–]\s*(\d{4}|aujourd’hui|présent)', line, re.IGNORECASE)
        if date_match:
            titre = lines[i-1] if i>0 else "Experience"
            entreprise = lines[i+1] if i+1 < len(lines) else ""
            description = lines[i+2] if i+2 < len(lines) else ""
            experiences.append({"titre": titre.strip(), "date": date_match.group(0), "entreprise": entreprise.strip(), "description": description.strip()})

    formation_keywords = ['licence', 'master', 'doctorat', 'bachelor', 'ingénieur', 'diplôme', 'bac', 'bts', 'dut']
    for i, line in enumerate(lines):
        if any(kw in line.lower() for kw in formation_keywords):
            annee_match = re.search(r'(\d{4})', line)
            annee = annee_match.group(1) if annee_match else ""
            formations.append({"diplome": line.strip(), "annee": annee, "etablissement": lines[i+1] if i+1 < len(lines) else "", "description": ""})

    skill_keywords = ["python", "javascript", "angular", "react", "node", "mongodb", "sql",
                      "docker", "kubernetes", "machine learning", "ia", "java", "c++", "php",
                      "laravel", "symfony", "django", "flask", "fastapi", "vue", "typescript",
                      "html", "css", "git", "linux", "aws", "azure"]
    for sk in skill_keywords:
        if sk in text.lower() and sk not in competences:
            competences.append(sk)

    langue_keywords = {'anglais': 'Fluent', 'français': 'Fluent', 'arabe': 'Native', 'espagnol': 'Intermediate', 'allemand': 'Beginner'}
    for langue, niveau in langue_keywords.items():
        if langue in text.lower():
            langues.append({"nom": langue.capitalize(), "niveau": niveau})

    if "centre d'intérêt" in text.lower() or "hobbies" in text.lower():
        centres_interet = ["Sports", "Reading", "Travel"]

    exp_pattern = r"(\d+)\s*(?:ans|années|years|year)"
    exp_matches = re.findall(exp_pattern, text.lower())
    max_exp = max([int(x) for x in exp_matches]) if exp_matches else 0
    diplomas = [f['diplome'] for f in formations] if formations else []

    return {
        "experiences": experiences,
        "formations": formations,
        "competences": competences,
        "langues": langues,
        "centres_interet": centres_interet,
        "max_experience": max_exp,
        "diplomas": diplomas
    }

def calculate_advanced_score(cv_filename, job_keywords):
    cv_path = os.path.join(app.config['UPLOAD_FOLDER'], cv_filename)
    if not os.path.exists(cv_path):
        return 20
    info = extract_cv_info_spacy(cv_path)
    if job_keywords and len(job_keywords) > 0:
        keyword_matches = sum(1 for kw in job_keywords if kw.lower() in info["full_text"])
        keyword_score = (keyword_matches / len(job_keywords)) * 50
    else:
        keyword_score = 0
    exp_bonus = min(info["max_experience"] * 5, 30)
    diploma_bonus = 20 if info["diplomas"] else 0
    total = keyword_score + exp_bonus + diploma_bonus
    return max(10, min(100, round(total, 2)))

def get_status_title(status):
    titles = {
        "validated": "✅ Application validated",
        "contacted": "📞 Contact made",
        "archived": "📁 Application archived",
        "pending": "⏳ Under review",
        "to_contact": "📞 To be contacted",
        "interview_scheduled": "📅 Interview scheduled",
        "sent_to_client": "📤 Sent to client"
    }
    return titles.get(status, f"Status: {status}")




def get_status_message_text(status, poste):
    """Retourne un message texte simple pour la notification"""
    messages = {
        "validated": f"✅ Votre candidature pour le poste '{poste}' a été validée avec succès !",
        "contacted": f"📞 Un recruteur va vous contacter pour le poste '{poste}'.",
        "archived": f"📁 Votre candidature pour '{poste}' a été archivée. Nous vous recontacterons pour d'autres opportunités.",
        "pending": f"⏳ Votre candidature pour '{poste}' est en cours d'examen.",
        "to_contact": f"📞 Votre profil a été présélectionné pour le poste '{poste}'. Un recruteur vous contactera.",
        "interview_scheduled": f"📅 Un entretien a été programmé pour votre candidature au poste '{poste}'.",
        "sent_to_client": f"📤 Votre candidature pour '{poste}' a été envoyée au client."
    }
    return messages.get(status, f"📌 Statut mis à jour: {status} pour '{poste}'")



def is_status_value(value):
    return value in ["pending", "validated", "contacted", "archived", "en attente", "validé", "contacté", "archivé"]

def admin_required():
    def wrapper(fn):
        @wraps(fn)
        @jwt_required()
        def decorator(*args, **kwargs):
            current_user = get_jwt_identity()
            admin = mongo.db.Users.find_one({"email": current_user, "role": "admin"})
            if not admin:
                return jsonify({"error": "Admin access required"}), 403
            return fn(*args, **kwargs)
        return decorator
    return wrapper

def load_evaluation_template():
    try:
        wb = openpyxl.load_workbook('Evaluation sheet Energy engineer.xlsx')
        sheet = wb.active
        questions = []
        reading = False
        for row in sheet.iter_rows(values_only=True):
            if row[0] and "technical skills" in str(row[0]).lower():
                reading = True
                continue
            if reading and row[0]:
                q = str(row[0]).strip()
                if q:
                    questions.append(q)
                if len(questions) >= 10:
                    break
    except FileNotFoundError:
        questions = [
            "Describe your main technical experience.",
            "How do you handle tight deadlines?"
        ]
    soft_skills = [
        "First impression",
        "Oral presentation",
        "Communication skills",
        "Team spirit",
        "Problem solving",
        "Adaptability"
    ]
    return {"technical": questions, "soft": soft_skills}










def generate_cv_pdf(candidature, candidate_info, cv_details, answers, is_anonymous=False):
    """
    Génère un dossier PDF professionnel avec ReportLab - Design Premium
    - is_anonymous=True : Masque nom, email, téléphone, nationalité
    - is_anonymous=False : Affiche toutes les informations personnelles
    """
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm, cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, PageBreak
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
    import io
    
    try:
        # Police moderne
        try:
            pdfmetrics.registerFont(TTFont('Poppins', 'Poppins-Regular.ttf'))
            pdfmetrics.registerFont(TTFont('Poppins-Bold', 'Poppins-Bold.ttf'))
            font_family = 'Poppins'
            font_bold = 'Poppins-Bold'
        except:
            font_family = 'Helvetica'
            font_bold = 'Helvetica-Bold'
        
        styles = getSampleStyleSheet()
        
        # Styles Premium
        style_title = ParagraphStyle(
            'CustomTitle', parent=styles['Heading1'],
            fontName=font_bold, fontSize=28, textColor=colors.white,
            alignment=TA_CENTER, spaceAfter=25, spaceBefore=15
        )
        
        style_subtitle = ParagraphStyle(
            'CustomSubtitle', parent=styles['Normal'],
            fontName=font_family, fontSize=11, textColor=colors.HexColor('#cbd5e1'),
            alignment=TA_CENTER, spaceAfter=10
        )
        
        style_section = ParagraphStyle(
            'SectionTitle', parent=styles['Heading2'],
            fontName=font_bold, fontSize=18, textColor=colors.HexColor('#0f172a'),
            spaceAfter=12, spaceBefore=20, backColor=colors.HexColor('#f1f5f9'),
            borderPadding=8, borderRadius=8
        )
        
        style_card_title = ParagraphStyle(
            'CardTitle', parent=styles['Normal'],
            fontName=font_bold, fontSize=14, textColor=colors.HexColor('#1e293b'),
            spaceAfter=6
        )
        
        style_label = ParagraphStyle(
            'Label', parent=styles['Normal'],
            fontName=font_bold, fontSize=10, textColor=colors.HexColor('#64748b'),
            alignment=TA_LEFT
        )
        
        style_value = ParagraphStyle(
            'Value', parent=styles['Normal'],
            fontName=font_family, fontSize=11, textColor=colors.HexColor('#1e293b'),
            alignment=TA_LEFT, leading=16
        )
        
        style_normal = ParagraphStyle(
            'NormalCustom', parent=styles['Normal'],
            fontName=font_family, fontSize=10, textColor=colors.HexColor('#334155'),
            alignment=TA_LEFT, leading=14
        )
        
        style_skill = ParagraphStyle(
            'Skill', parent=styles['Normal'],
            fontName=font_family, fontSize=9, textColor=colors.HexColor('#00695c'),
            alignment=TA_CENTER
        )
        
        style_question = ParagraphStyle(
            'Question', parent=styles['Normal'],
            fontName=font_bold, fontSize=11, textColor=colors.HexColor('#00a6a6'),
            alignment=TA_LEFT, spaceAfter=4
        )
        
        style_answer = ParagraphStyle(
            'Answer', parent=styles['Normal'],
            fontName=font_family, fontSize=10, textColor=colors.HexColor('#334155'),
            alignment=TA_LEFT, leading=14, leftIndent=15
        )
        
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer, pagesize=A4,
            leftMargin=1.8*cm, rightMargin=1.8*cm,
            topMargin=1.5*cm, bottomMargin=1.5*cm
        )
        
        story = []
        
        # ========== EN-TÊTE LCA PREMIUM ==========
        header_bg = colors.HexColor('#0f172a')
        header_data = [[
            Paragraph("<font color='white' size='26'><b>LAURA CONNECTING AGENCY</b></font>", style_title),
        ]]
        header_table = Table(header_data, colWidths=[doc.width])
        header_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), header_bg),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 25),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 15),
        ]))
        story.append(header_table)
        story.append(Paragraph("Talent Consulting & Recruitment", style_subtitle))
        story.append(Spacer(1, 0.3*cm))
        
        # ========== BANDEAU SCORE AVEC Dégradé ==========
        score = candidature.get('score', 0)
        if score >= 70:
            score_color = '#10b981'
            score_icon = '🏆'
            score_text = 'Excellent - Très bon profil'
        elif score >= 50:
            score_color = '#f59e0b'
            score_icon = '📈'
            score_text = 'Potentiel - À approfondir'
        else:
            score_color = '#ef4444'
            score_icon = '📊'
            score_text = 'À améliorer - Profil en développement'
        
        score_data = [[
            Paragraph(f"<font color='white' size='16'><b>{score_icon} SCORE IA: {score}% - {score_text}</b></font>", style_title)
        ]]
        score_table = Table(score_data, colWidths=[doc.width])
        score_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(score_color)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ]))
        story.append(score_table)
        story.append(Spacer(1, 0.4*cm))
        
        # ========== TITRE CANDIDAT ==========
        if is_anonymous:
            story.append(Paragraph("📄 DOSSIER DE CANDIDATURE", 
                ParagraphStyle('CandidateTitle', parent=style_section, fontSize=20, alignment=TA_CENTER, textColor=colors.HexColor('#00a6a6'))))
            story.append(Paragraph("<i>Version anonyme - Informations personnelles masquées</i>", style_subtitle))
        else:
            first_name = candidate_info.get('first_name', '')
            last_name = candidate_info.get('last_name', '')
            candidate_full_name = f"{first_name} {last_name}".strip() or "Candidat"
            story.append(Paragraph(f"👤 DOSSIER DE CANDIDATURE - <b>{candidate_full_name}</b>", 
                ParagraphStyle('CandidateTitle', parent=style_section, fontSize=20, alignment=TA_CENTER, textColor=colors.HexColor('#00a6a6'))))
        story.append(Spacer(1, 0.3*cm))
        
        # ========== SECTION 1: INFORMATIONS PERSONNELLES ==========
        story.append(Paragraph("📋 INFORMATIONS PERSONNELLES", style_section))
        
        if is_anonymous:
            info_data = [
                ["🔒 Version anonyme", "Les informations personnelles sont confidentielles"],
                ["📊 Évaluation basée sur", "Compétences techniques et expérience uniquement"],
            ]
        else:
            email = candidate_info.get('email', 'Non renseigné')
            phone = candidate_info.get('phone', 'Non renseigné')
            nationalite = candidate_info.get('nationalite', 'Non renseignée')
            poste = candidature.get('poste', 'Non spécifié')
            application_date = candidature.get('date', datetime.now())
            if isinstance(application_date, datetime):
                application_date = application_date.strftime('%d/%m/%Y')
            
            info_data = [
                ["👤 Nom complet", f"{candidate_info.get('first_name', '')} {candidate_info.get('last_name', '')}"],
                ["📧 Email", email],
                ["📞 Téléphone", phone],
                ["🌍 Nationalité", nationalite],
                ["💼 Poste visé", poste],
                ["📅 Date candidature", application_date],
            ]
        
        info_table = Table(info_data, colWidths=[4.5*cm, 11*cm])
        info_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f8fafc')),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            ('LEFTPADDING', (0, 0), (-1, -1), 12),
            ('FONTNAME', (0, 0), (0, -1), font_bold),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
        ]))
        story.append(info_table)
        story.append(Spacer(1, 0.4*cm))
        
        # ========== SECTION 2: COMPÉTENCES TECHNIQUES ==========
        story.append(Paragraph("💻 COMPÉTENCES TECHNIQUES", style_section))
        skills = cv_details.get('competences', [])
        if skills:
            skill_rows = []
            row = []
            for i, skill in enumerate(skills[:24]):
                skill_text = Paragraph(f"▹ {skill}", style_skill)
                row.append(skill_text)
                if (i + 1) % 3 == 0:
                    skill_rows.append(row)
                    row = []
            if row:
                skill_rows.append(row)
            
            skill_table = Table(skill_rows, colWidths=[doc.width/3, doc.width/3, doc.width/3])
            skill_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#e0f2f1')),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                ('LEFTPADDING', (0, 0), (-1, -1), 12),
            ]))
            story.append(skill_table)
        else:
            story.append(Paragraph("⚠️ Aucune compétence détectée dans le CV", style_normal))
        story.append(Spacer(1, 0.4*cm))
        
        # ========== SECTION 3: EXPÉRIENCE PROFESSIONNELLE ==========
        experiences = cv_details.get('experiences', [])
        if experiences:
            story.append(Paragraph("💼 EXPÉRIENCE PROFESSIONNELLE", style_section))
            for exp in experiences[:5]:
                exp_title = exp.get('titre', 'Expérience')
                exp_date = exp.get('date', '')
                exp_company = exp.get('entreprise', '')
                exp_desc = exp.get('description', '')
                
                exp_html = f"""
                <b><font color='#0f172a' size='11'>{exp_title}</font></b><br/>
                <font color='#00a6a6' size='9'>{exp_date}</font><br/>
                <font color='#475569' size='9'>{exp_company}</font><br/>
                <font color='#64748b' size='9'>{exp_desc[:200]}</font>
                """
                story.append(Paragraph(exp_html, style_normal))
                story.append(Spacer(1, 0.25*cm))
        else:
            story.append(Paragraph("⚠️ Aucune expérience détectée dans le CV", style_normal))
            story.append(Spacer(1, 0.4*cm))
        
        # ========== SECTION 4: FORMATION ==========
        formations = cv_details.get('formations', [])
        if formations:
            story.append(Paragraph("🎓 FORMATION & DIPLÔMES", style_section))
            for f in formations[:5]:
                diploma = f.get('diplome', 'Diplôme')
                annee = f.get('annee', '')
                etablissement = f.get('etablissement', '')
                
                edu_html = f"""
                <b><font color='#0f172a' size='11'>{diploma}</font></b><br/>
                <font color='#00a6a6' size='9'>{annee}</font><br/>
                <font color='#475569' size='9'>{etablissement}</font>
                """
                story.append(Paragraph(edu_html, style_normal))
                story.append(Spacer(1, 0.25*cm))
        else:
            story.append(Paragraph("⚠️ Aucune formation détectée dans le CV", style_normal))
            story.append(Spacer(1, 0.4*cm))
        
        # ========== SECTION 5: LANGUES ==========
        langues = cv_details.get('langues', [])
        if langues:
            story.append(Paragraph("🌐 LANGUES", style_section))
            for langue in langues:
                nom = langue.get('nom', '')
                niveau = langue.get('niveau', '')
                niveau_icon = "🔵" if niveau == "Débutant" else "🟡" if niveau == "Intermédiaire" else "🟠" if niveau == "Avancé" else "🟢" if niveau == "Courant" else "⭐" if niveau == "Natif" else "📖"
                langue_html = f"<b>{nom}</b> {niveau_icon} {niveau}"
                story.append(Paragraph(langue_html, style_normal))
                story.append(Spacer(1, 0.15*cm))
        else:
            story.append(Paragraph("⚠️ Aucune langue renseignée", style_normal))
            story.append(Spacer(1, 0.4*cm))
        
        # ========== SECTION 6: RÉPONSES QUESTIONNAIRE (AMÉLIORÉE) ==========
        if answers and len(answers) > 0:
            story.append(PageBreak())
            story.append(Paragraph("📋 RÉPONSES AU QUESTIONNAIRE", style_section))
            
            valid_answers = []
            for idx, a in enumerate(answers):
                if not isinstance(a, dict):
                    continue
                
                # Récupérer question et réponse
                question = None
                for key in ['question', 'questionText', 'question_text', 'text']:
                    if key in a and a[key] and str(a[key]) != 'null':
                        question = str(a[key]).strip()
                        break
                
                if not question:
                    continue
                
                # Ignorer les champs techniques
                if question.lower().startswith('gestion') and len(question) < 20:
                    continue
                if question in ['field', 'gestion', 'Gestion']:
                    continue
                
                answer = None
                for key in ['answer', 'answerText', 'response', 'value']:
                    if key in a and a[key] and str(a[key]) != 'null':
                        answer = str(a[key]).strip()
                        break
                
                if not answer or answer == '':
                    continue
                
                valid_answers.append({
                    'idx': len(valid_answers) + 1,
                    'question': question,
                    'answer': answer
                })
            
            if valid_answers:
                for qa in valid_answers:
                    answer_text = qa['answer']
                    if len(answer_text) > 250:
                        answer_text = answer_text[:250] + "..."
                    
                    # Carte réponse
                    story.append(Paragraph(f"<b>❓ Question {qa['idx']}</b>", style_question))
                    story.append(Paragraph(qa['question'], style_label))
                    story.append(Spacer(1, 0.1*cm))
                    story.append(Paragraph(f"<b>✏️ Réponse :</b> {answer_text}", style_answer))
                    story.append(Spacer(1, 0.25*cm))
                    
                    if qa['idx'] % 8 == 0 and qa['idx'] < len(valid_answers):
                        story.append(PageBreak())
                
                story.append(Spacer(1, 0.3*cm))
                story.append(Paragraph(f"📊 <b>Total:</b> {len(valid_answers)} réponses valides", 
                    ParagraphStyle('Summary', parent=style_normal, alignment=TA_CENTER, textColor=colors.HexColor('#00a6a6'))))
            else:
                story.append(Paragraph("⚠️ Aucune réponse valide trouvée", style_normal))
        
        # ========== PIED DE PAGE PREMIUM ==========
        story.append(Spacer(1, 0.8*cm))
        footer_data = [[
            Paragraph("Ce document est confidentiel et appartient à <b>Laura Connecting Agency</b><br/>© 2025 LCA - Tous droits réservés", 
                ParagraphStyle('Footer', parent=style_subtitle, fontSize=8, textColor=colors.HexColor('#94a3b8'), alignment=TA_CENTER))
        ]]
        footer_table = Table(footer_data, colWidths=[doc.width])
        footer_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#0f172a')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('TOPPADDING', (0, 0), (-1, -1), 15),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 15),
        ]))
        story.append(footer_table)
        
        # Générer le PDF
        doc.build(story)
        buffer.seek(0)
        pdf_bytes = buffer.getvalue()
        
        print(f"✅ PDF généré avec succès, taille: {len(pdf_bytes)} bytes")
        return pdf_bytes
        
    except Exception as e:
        print(f"❌ Erreur génération PDF: {e}")
        import traceback
        traceback.print_exc()
        
        # Fallback
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=A4)
        c.setFont("Helvetica-Bold", 16)
        if is_anonymous:
            c.drawString(50, 800, "Dossier de candidature (Version anonyme)")
            c.drawString(50, 780, "Score IA: " + str(candidature.get('score', 0)) + "%")
        else:
            first_name = candidate_info.get('first_name', '')
            last_name = candidate_info.get('last_name', '')
            c.drawString(50, 800, f"Dossier - {first_name} {last_name}")
        c.save()
        buffer.seek(0)
        return buffer.getvalue()








def init_daily_tips():
    if mongo.db.daily_tips.count_documents({}) == 0:
        tips = [
            "Update your CV every month with new skills.",
            "Customize your cover letter for each application.",
            "Take online courses to stay competitive.",
            "Create a GitHub portfolio to showcase your projects.",
            "Networking on LinkedIn increases your chances by 40%.",
            "Attend webinars to expand your network.",
            "Ask former colleagues for recommendations.",
            "Apply even if you don't meet 100% of the requirements.",
            "Prepare 3 relevant questions for the interview.",
            "Proofread your CV to remove typos."
        ]
        for tip in tips:
            mongo.db.daily_tips.insert_one({"text": tip, "used_dates": []})


def serialize_doc(doc):
    """Convertit un document MongoDB en dictionnaire JSON-sérialisable"""
    if isinstance(doc, dict):
        return {k: serialize_doc(v) for k, v in doc.items()}
    elif isinstance(doc, list):
        return [serialize_doc(item) for item in doc]
    elif isinstance(doc, ObjectId):
        return str(doc)
    elif isinstance(doc, datetime):
        return doc.isoformat()
    else:
        return doc

   
# ------------------------------------------------------------------
# Public routes
# ------------------------------------------------------------------
@app.route('/api/jobs', methods=['GET'])
def get_jobs():
    try:
        jobs = list(mongo.db.postes.find({"status": "active"}))
        out = []
        for j in jobs:
            company = mongo.db.entreprises.find_one({"_id": ObjectId(j.get('companyId'))})
            out.append({
                'id': str(j['_id']),
                'title': j.get('title', ''),
                'location': j.get('location', 'Remote'),
                'sector': company.get('secteur') or company.get('sector', 'General') if company else 'General',
                'type': j.get('type', 'Full-time'),
                'description': j.get('description', ''),
                'company': company.get('name', 'LCA') if company else 'LCA',
                'companyId': str(j.get('companyId')) if j.get('companyId') else '',
                'salary': j.get('salary', 'N/A'),
                'responsibilities': j.get('responsibilities', []),
                'featured': j.get('featured', False),
                'created_at': j.get('created_at').isoformat() if j.get('created_at') else None
            })
        return jsonify(out), 200
    except Exception as e:
        print(f"❌ get_jobs: {e}")
        return jsonify([]), 200




@app.route('/api/stats', methods=['GET'])
def get_stats():
    return jsonify({
        "cv_scored": f"{mongo.db.candidatures.count_documents({})}+",
        "accuracy": "95%",
        "partners": f"{mongo.db.entreprises.count_documents({})}+",
        "speed": "3x",
        "total_candidates": mongo.db.candidats.count_documents({})
    })

@app.route('/api/contact', methods=['POST'])
def handle_contact():
    data = request.json
    mongo.db.messages.insert_one({
        "fullName": data.get('fullName'),
        "email": data.get('email'),
        "phone": data.get('phone'),
        "subject": data.get('subject'),
        "message": data.get('message'),
        "status": "unread",
        "date": datetime.utcnow()
    })
    return jsonify({"message": "Message sent!"}), 201

@app.route('/api/catalog', methods=['GET'])
def get_catalog():
    catalog = []
    jobs = list(mongo.db.postes.find({"status": "active"}))
    for job in jobs:
        count = mongo.db.candidatures.count_documents({"posteId": str(job["_id"])})
        catalog.append({
            "role_title": job.get("title", "Offer"),
            "category": job.get("sector") or job.get("title", "general"),
            "short_desc": (job.get("description") or "")[:120],
            "count": count,
            "icon": "fas fa-briefcase"
        })
    if not catalog:
        catalog = [
            {"role_title": "Developer", "category": "dev", "short_desc": "Technical profile", "count": 0, "icon": "fas fa-code"},
            {"role_title": "Sales", "category": "commerce", "short_desc": "Sales profile", "count": 0, "icon": "fas fa-chart-line"},
            {"role_title": "Manager", "category": "gestion", "short_desc": "Administrative profile", "count": 0, "icon": "fas fa-building"}
        ]
    return jsonify(catalog), 200

def get_icon_for_category(category):
    mapping = {
        "informatique": "fas fa-laptop-code",
        "vente": "fas fa-chart-line",
        "commerce": "fas fa-shopping-cart",
        "gestion": "fas fa-file-invoice",
        "technique": "fas fa-microchip",
        "ingénierie": "fas fa-cogs",
        "support": "fas fa-headset"
    }
    category_lower = category.lower()
    for key, icon in mapping.items():
        if key in category_lower:
            return icon
    return "fas fa-briefcase"

# ------------------------------------------------------------------
# Admin routes
# ------------------------------------------------------------------
@app.route('/admin-login', methods=['POST'])
def admin_login():
    data = request.json
    admin = mongo.db.Users.find_one({"email": data['email'], "role": "admin"})
    if admin and check_password_hash(admin['password'], data['password']):
        access_token = create_access_token(identity=admin['email'])
        return jsonify({
            "token": access_token,
            "nom": admin.get('nom', 'Admin'),
            "image": admin.get('image', 'default-avatar.png')
        }), 200
    return jsonify({"error": "Invalid credentials"}), 401


@app.route('/api/admin/verify-token', methods=['GET'])
@jwt_required()
def verify_admin_token():
    """Vérifie si le token admin est valide"""
    current_user = get_jwt_identity()
    admin = mongo.db.Users.find_one({"email": current_user, "role": "admin"})
    if not admin:
        return jsonify({"error": "Invalid token"}), 401
    return jsonify({"valid": True, "email": current_user}), 200








@app.route('/api/admin/refresh-token', methods=['POST'])
@admin_required()
def refresh_token():
    current_user = get_jwt_identity()
    new_token = create_access_token(identity=current_user)
    return jsonify({"token": new_token}), 200

@app.route('/api/admin/update-photo', methods=['POST'])
def update_admin_photo():
    file = request.files.get('photo')
    email = request.form.get('email')
    if file:
        filename = secure_filename(f"updated_{email}_{file.filename}")
        file.save(os.path.join('static/uploads/admins', filename))
        mongo.db.Users.update_one({"email": email}, {"$set": {"image": filename}})
        return jsonify({"new_image": filename}), 200
    return jsonify({"error": "No file"}), 400

@app.route('/api/admin/profile', methods=['GET'])
def get_admin_profile():
    email = request.args.get('email')
    if not email:
        return jsonify({"error": "Email required"}), 400
    admin = mongo.db.Users.find_one({"email": email, "role": "admin"})
    if admin:
        return jsonify({"nom": admin.get('nom'), "image": admin.get('image', 'default-avatar.png')}), 200
    return jsonify({"error": "Admin not found"}), 404

@app.route('/api/admin/update-profile', methods=['PUT'])
def update_admin_profile():
    data = request.json
    email = data.get('email')
    nom = data.get('nom')
    if email:
        mongo.db.Users.update_one({"email": email}, {"$set": {"nom": nom}})
    return jsonify({"message": "Profile updated"}), 200

@app.route('/api/admin/main-stats', methods=['GET'])
def get_main_stats():
    return jsonify({
        "totalCandidates": mongo.db.candidats.count_documents({}),
        "activeJobs": mongo.db.postes.count_documents({"status": "active"}),
        "totalCompanies": mongo.db.entreprises.count_documents({}),
        "matchedCVs": mongo.db.candidatures.count_documents({"score": {"$gte": 70}})
    }), 200

@app.route('/api/admin/recent-activities', methods=['GET'])
def get_recent_activities():
    acts = list(mongo.db.candidatures.find().sort("date", -1).limit(5))
    for a in acts:
        stored_poste = a.get('poste', '')
        if not stored_poste or is_status_value(stored_poste.lower()):
            job_id = a.get('posteId') or a.get('jobId')
            if job_id and ObjectId.is_valid(job_id):
                job = mongo.db.postes.find_one({"_id": ObjectId(job_id)})
                if job and job.get('title'):
                    a['poste'] = job['title']
                    mongo.db.candidatures.update_one({"_id": a['_id']}, {"$set": {"poste": job['title']}})
            else:
                a['poste'] = "Deleted offer"
    return jsonify([{
        "candidate": a.get("candidateName", "Anonymous"),
        "job_title": a.get("poste", ""),
        "score": a.get("score", 0),
        "date": a.get("date", datetime.utcnow()).isoformat()
    } for a in acts]), 200

@app.route('/api/admin/sidebar-counts', methods=['GET'])
def get_sidebar_counts():
    return jsonify({
        "pending": mongo.db.candidatures.count_documents({"status": "pending", "hidden": {"$ne": True}}),
        "to_contact": mongo.db.candidatures.count_documents({"status": "to_contact", "hidden": {"$ne": True}}),
        "interview_scheduled": mongo.db.candidatures.count_documents({"status": "interview_scheduled", "hidden": {"$ne": True}}),
        "sent_to_client": mongo.db.candidatures.count_documents({"status": "sent_to_client", "hidden": {"$ne": True}}),
        "archived": mongo.db.candidatures.count_documents({"status": "archived", "hidden": {"$ne": True}})
    }), 200


@app.route('/api/candidatures/update-status', methods=['PUT'])
def update_candidature_status():
    data = request.json
    cand_id = data.get("id")
    new_status = data.get("status")
    
    if not cand_id or not new_status:
        return jsonify({"error": "ID and status required"}), 400
    
    if not ObjectId.is_valid(cand_id):
        return jsonify({"error": "Invalid ID"}), 400
    
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(cand_id)})
    if not candidature:
        return jsonify({"error": "Application not found"}), 404
    
    # Mise à jour du statut
    update_data = {"status": new_status, "updated_at": datetime.utcnow()}
    
    if new_status == "hired":
        update_data["final_decision"] = "hired"
        update_data["final_decision_date"] = datetime.utcnow()
    elif new_status == "rejected":
        update_data["final_decision"] = "rejected"
        update_data["final_decision_date"] = datetime.utcnow()
    
    mongo.db.candidatures.update_one(
        {"_id": ObjectId(cand_id)},
        {"$set": update_data}
    )
    
    candidate_email = candidature.get("candidateEmail")
    candidate_name = candidature.get("candidateName", "Dear candidate")
    poste = candidature.get("poste", "your application")
    
    # ✅ NOTIFICATION AVEC TEXTE SIMPLE (PAS HTML)
    mongo.db.notifications.insert_one({
        "recipient": "candidate",
        "recipientEmail": candidate_email,
        "type": "status_update",
        "title": get_status_title(new_status),
        "message": get_status_message_text(new_status, poste),  # ← Texte simple !
        "old_status": candidature.get("status", "pending"),
        "new_status": new_status,
        "status": "unread",
        "date": datetime.utcnow()
    })
    
    # ✅ EMAIL PERSONNALISÉ AVEC HTML (get_status_message existe déjà)
    if new_status == "hired":
        subject = "🎉 Application Accepted - Laura Connecting Agency"
        body = get_status_message("validated", poste, candidate_name)
        send_real_email(candidate_email, subject, body)
    elif new_status == "rejected":
        subject = "📌 Application Update - Laura Connecting Agency"
        body = get_status_message("archived", poste, candidate_name)
        send_real_email(candidate_email, subject, body)
    elif new_status == "archived":
        subject = "📌 Application Update - Laura Connecting Agency"
        body = get_status_message("archived", poste, candidate_name)
        send_real_email(candidate_email, subject, body)
    elif new_status == "to_contact":
        subject = "📞 Profile Selected - Laura Connecting Agency"
        body = get_status_message("to_contact", poste, candidate_name)
        send_real_email(candidate_email, subject, body)
    elif new_status == "interview_scheduled":
        subject = "📅 Interview Scheduled - Laura Connecting Agency"
        body = get_status_message("interview_scheduled", poste, candidate_name)
        send_real_email(candidate_email, subject, body)
    elif new_status == "sent_to_client":
        subject = "📤 Application Forwarded - Laura Connecting Agency"
        body = get_status_message("sent_to_client", poste, candidate_name)
        send_real_email(candidate_email, subject, body)
    elif new_status == "pending":
        subject = "⏳ Application Received - Laura Connecting Agency"
        body = get_status_message("pending", poste, candidate_name)
        send_real_email(candidate_email, subject, body)
    elif new_status == "validated":
        subject = "✅ Application Validated - Laura Connecting Agency"
        body = get_status_message("validated", poste, candidate_name)
        send_real_email(candidate_email, subject, body)
    else:
        subject = f"📌 Application Update: {poste} - Laura Connecting Agency"
        body = get_status_message("pending", poste, candidate_name)
        send_real_email(candidate_email, subject, body)
    
    # Si le statut est "validated", notifier le client
    if new_status == "validated":
        job = mongo.db.postes.find_one({"_id": ObjectId(candidature.get('posteId'))}) if candidature.get('posteId') else None
        if job and job.get('companyId'):
            company = mongo.db.entreprises.find_one({"_id": ObjectId(job['companyId'])})
            if company and company.get('email'):
                token = hashlib.md5(f"{candidate_email}{cand_id}".encode()).hexdigest()[:16]
                mongo.db.candidatures.update_one(
                    {"_id": ObjectId(cand_id)},
                    {"$set": {"consultation_token": token}}
                )
                consult_url = f"http://localhost:4200/company/status?token={token}"
                subject_company = f"✅ Candidate selected for {poste}"
                body_company = f"""
                <html><body><h2>Hello {company.get('name', '')},</h2>
                <p>The candidate <strong>{candidate_name}</strong> has been <strong>selected</strong> for the position <strong>{poste}</strong>.</p>
                <p><a href="{consult_url}">👉 View application status</a></p>
                <p>Best regards,<br>The LCA team</p></body></html>
                """
                send_real_email(company['email'], subject_company, body_company)
    
    # Notification pour l'admin si décision finale
    if new_status in ["hired", "rejected"]:
        mongo.db.notifications.insert_one({
            "recipient": "admin",
            "recipientEmail": "lca.platformlaura@gmail.com",
            "type": "final_decision",
            "title": "🎯 Décision finale enregistrée",
            "message": f"{candidate_name} - {'✅ Retenu' if new_status == 'hired' else '❌ Non retenu'} pour {poste}",
            "status": "unread",
            "date": datetime.utcnow(),
            "candidatureId": cand_id
        })
    
    return jsonify({"message": "Status updated"}), 200


def get_status_message(status, poste, candidate_name):
    """Retourne un message HTML pour l'email"""
    messages = {
        "validated": f"""
        <html>
        <body>
            <h2>Félicitations {candidate_name} !</h2>
            <p>Votre candidature pour le poste <strong>{poste}</strong> a été validée.</p>
            <p>Nous vous contacterons prochainement.</p>
        </body>
        </html>
        """,
        "archived": f"""
        <html>
        <body>
            <h2>Bonjour {candidate_name},</h2>
            <p>Votre candidature pour le poste <strong>{poste}</strong> a été archivée.</p>
            <p>Nous vous remercions de votre intérêt.</p>
        </body>
        </html>
        """,
        "pending": f"""
        <html>
        <body>
            <h2>Bonjour {candidate_name},</h2>
            <p>Votre candidature pour le poste <strong>{poste}</strong> est en cours d'examen.</p>
            <p>Nous vous tiendrons informé.</p>
        </body>
        </html>
        """
    }
    return messages.get(status, f"<p>Status mis à jour: {status} pour {poste}</p>")



def get_status_message_text(status, poste):
    """Retourne un message texte simple pour la notification"""
    messages = {
        "validated": f"✅ Votre candidature pour le poste '{poste}' a été validée avec succès !",
        "contacted": f"📞 Un recruteur va vous contacter pour le poste '{poste}'.",
        "archived": f"📁 Votre candidature pour '{poste}' a été archivée. Nous vous recontacterons pour d'autres opportunités.",
        "pending": f"⏳ Votre candidature pour '{poste}' est en cours d'examen.",
        "to_contact": f"📞 Votre profil a été présélectionné pour le poste '{poste}'. Un recruteur vous contactera.",
        "interview_scheduled": f"📅 Un entretien a été programmé pour votre candidature au poste '{poste}'.",
        "sent_to_client": f"📤 Votre candidature pour '{poste}' a été envoyée au client."
    }
    return messages.get(status, f"📌 Statut mis à jour: {status} pour '{poste}'")




@app.route('/api/subscribe', methods=['POST', 'OPTIONS'])
def subscribe_newsletter():
    if request.method == 'OPTIONS':
        return '', 200
    data = request.json
    email = data.get('email') if data else None
    if not email:
        return jsonify({"error": "Email required"}), 400
    
    # Check for existing subscription
    existing = mongo.db.newsletter_subscribers.find_one({"email": email})
    if existing:
        return jsonify({"message": "You are already subscribed!"}), 200
    
    # Insert new subscriber
    mongo.db.newsletter_subscribers.insert_one({
        "email": email,
        "subscribed_at": datetime.utcnow()
    })
    return jsonify({"message": "Subscribed successfully!"}), 200

@app.route('/api/company/status', methods=['GET'])
def company_check_status():
    token = request.args.get('token')
    if not token:
        return jsonify({"error": "Missing token"}), 400
    candidature = mongo.db.candidatures.find_one({"consultation_token": token})
    if not candidature:
        return jsonify({"error": "Invalid or expired link"}), 404
    candidate_name = candidature.get('candidateName')
    poste = candidature.get('poste')
    status = candidature.get('status')
    company_id = candidature.get('companyId')
    company = mongo.db.entreprises.find_one({"_id": ObjectId(company_id)}) if company_id else None
    return jsonify({
        "candidate_name": candidate_name,
        "poste": poste,
        "status": status,
        "company_name": company.get('name') if company else "Unknown",
        "message": "Candidate selected" if status == "validated" else "Status in progress"
    }), 200

# Companies
@app.route('/api/companies', methods=['GET'])
def get_companies():
    try:
        companies = list(mongo.db.entreprises.find())
        result = []
        for c in companies:
            # Convertir ObjectId en string
            company_data = {
                '_id': str(c['_id']),
                'name': c.get('name', ''),
                'secteur': c.get('secteur') or c.get('sector', ''),
                'location': c.get('location', ''),
                'email': c.get('email', ''),
                'status': c.get('status', 'Active'),
                'logo': c.get('logo', 'assets/default-logo.png'),
                'ville': c.get('ville', ''),
                'sector': c.get('sector', '')
            }
            
            # Récupérer les jobs
            jobs = list(mongo.db.postes.find({"companyId": str(c['_id'])}))
            company_data['jobs'] = []
            for job in jobs:
                job_data = {
                    '_id': str(job['_id']),
                    'title': job.get('title', ''),
                    'type': job.get('type', ''),
                    'location': job.get('location', ''),
                    'description': job.get('description', ''),
                    'status': job.get('status', 'active'),
                    'keywords': job.get('keywords', []),
                    'candidatures_count': mongo.db.candidatures.count_documents({"posteId": str(job['_id'])}),
                    'candidatures': []
                }
                company_data['jobs'].append(job_data)
            
            company_data['jobs_count'] = len(company_data['jobs'])
            result.append(company_data)
        
         # ⭐ AJOUTEZ serialize_doc ICI ⭐
        return jsonify(serialize_doc(result)), 200
    except Exception as e:
        print(f"❌ Erreur get_companies: {e}")
        return jsonify([]), 200


@app.route('/api/companies/add', methods=['POST'])
def add_company():
    data = request.json
    mongo.db.entreprises.insert_one({
        "name": data.get("name"),
        "secteur": data.get("sector"),
        "location": data.get("location"),
        "email": data.get("email"),
        "status": data.get("status", "Active"),
        "logo": "assets/default-logo.png"
    })
    return jsonify({"message": "Success"}), 201




@app.route('/api/companies/delete/<id>', methods=['DELETE'])
def delete_company(id):
    if not ObjectId.is_valid(id):
        return jsonify({"error": "Invalid ID"}), 400
    mongo.db.entreprises.delete_one({"_id": ObjectId(id)})
    return jsonify({"message": "Deleted"}), 200

@app.route('/api/companies/update/<id>', methods=['PUT'])
def update_company(id):
    data = request.json
    mongo.db.entreprises.update_one({"_id": ObjectId(id)}, {"$set": {
        "name": data.get('name'),
        "secteur": data.get('sector'),
        "location": data.get('location'),
        "email": data.get('email'),
        "status": data.get('status')
    }})
    return jsonify({"message": "Updated"}), 200

# Jobs
@app.route('/api/companies/<company_id>/jobs', methods=['GET'])
def get_company_jobs(company_id):
    jobs = list(mongo.db.postes.find({"companyId": company_id}))
    for j in jobs:
        j['_id'] = str(j['_id'])
        j['candidatures_count'] = mongo.db.candidatures.count_documents({"posteId": j['_id']})
    return jsonify(jobs), 200

@app.route('/api/companies/<company_id>/jobs/add', methods=['POST'])
def add_company_job(company_id):
    data = request.json
    company = mongo.db.entreprises.find_one({"_id": ObjectId(company_id)})
    if not company:
        return jsonify({"error": "Company not found"}), 404
    new_job = {
        "companyId": company_id,
        "companyName": company.get('name'),
        "title": data.get('title'),
        "description": data.get('description', ''),
        "location": data.get('location', company.get('location')),
        "type": data.get('type', 'Full-time'),
        "keywords": data.get('keywords', []),
        "evaluation_questions": data.get('evaluation_questions', []),
        "status": "active",
        "featured": data.get('featured', False),
        "created_at": datetime.utcnow()
    }
    result = mongo.db.postes.insert_one(new_job)
    return jsonify({"message": "Job added", "jobId": str(result.inserted_id)}), 201


@app.route('/api/jobs/update/<job_id>', methods=['PUT'])
def update_job(job_id):
    data = request.json
    update_data = {
        "title": data.get('title'),
        "description": data.get('description'),
        "location": data.get('location'),
        "type": data.get('type'),
        "keywords": data.get('keywords'),
    }
    if 'featured' in data:
        update_data['featured'] = data['featured']
    update_data = {k: v for k, v in update_data.items() if v is not None}
    if 'evaluation_questions' in data:
        update_data['evaluation_questions'] = data.get('evaluation_questions')
    if not update_data:
        return jsonify({"error": "No data to update"}), 400
    mongo.db.postes.update_one({"_id": ObjectId(job_id)}, {"$set": update_data})
    return jsonify({"message": "Job updated"}), 200

@app.route('/api/jobs/delete/<job_id>', methods=['DELETE'])
def delete_job(job_id):
    mongo.db.postes.delete_one({"_id": ObjectId(job_id)})
    mongo.db.candidatures.delete_many({"posteId": job_id})
    return jsonify({"message": "Deleted"}), 200






@app.route('/api/jobs/<job_id>', methods=['GET'])
def get_job_by_id(job_id):
    """Récupère un job par son ID avec ses questionnaires"""
    try:
        if not ObjectId.is_valid(job_id):
            return jsonify({"error": "Invalid ID"}), 400
        
        job = mongo.db.postes.find_one({"_id": ObjectId(job_id)})
        if not job:
            return jsonify({"error": "Job not found"}), 404
        
        # Convertir ObjectId en string pour la réponse JSON
        job['_id'] = str(job['_id'])
        
        # S'assurer que customQuestionSections existe
        if 'customQuestionSections' not in job:
            job['customQuestionSections'] = []
        
        return jsonify(job), 200
    except Exception as e:
        print(f"❌ Error: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/api/jobs/<job_id>/candidates', methods=['GET'])
def get_job_candidates(job_id):
    candidatures = list(mongo.db.candidatures.find({"posteId": job_id}))
    for c in candidatures:
        c['_id'] = str(c['_id'])
        email = c.get('candidateEmail')
        if not email:
            email = c.get('candidate_email') or c.get('email')
            if not email:
                email = 'Unknown email'
            c['candidateEmail'] = email
        cand = mongo.db.candidats.find_one({"email": email}) if email and email != 'Unknown email' else None
        if cand:
            c['candidate_first_name'] = cand.get('first_name', '')
            c['candidate_last_name'] = cand.get('last_name', '')
            c['candidate_photo'] = cand.get('photo_url', '')
        else:
            c['candidate_first_name'] = c.get('candidateName', 'Anonymous')
            c['candidate_last_name'] = ''
            if email == 'Unknown email':
                c['candidate_first_name'] = 'Missing email'
        if 'date' not in c or c['date'] is None:
            c['date'] = c.get('created_at', datetime.utcnow())
        if isinstance(c['date'], datetime):
            c['date'] = c['date'].isoformat()
        if not isinstance(c['date'], str) or c['date'] == '':
            c['date'] = datetime.utcnow().isoformat()
    return jsonify(candidatures), 200

@app.route('/api/jobs/all', methods=['GET'])
def get_all_jobs():
    jobs = list(mongo.db.postes.find())
    for j in jobs:
        j['_id'] = str(j['_id'])
    return jsonify(jobs), 200

@app.route('/api/admin/top-matches', methods=['GET'])
def get_top_matches():
    pipeline = [
        {"$match": {"score": {"$exists": True}}},
        {"$sort": {"score": -1}},
        {"$limit": 5},
        {"$lookup": {"from": "candidats", "localField": "candidateEmail", "foreignField": "email", "as": "candidate_info"}},
        {"$unwind": {"path": "$candidate_info", "preserveNullAndEmptyArrays": True}}
    ]
    top = list(mongo.db.candidatures.aggregate(pipeline))
    seen = set()
    unique = []
    for t in top:
        email = t.get("candidateEmail")
        if email not in seen:
            seen.add(email)
            unique.append(t)
    result = []
    for t in unique[:4]:
        cand = t.get("candidate_info", {})
        result.append({
            "name": t.get("candidateName") or f"{cand.get('first_name', '')} {cand.get('last_name', '')}".strip() or "Anonymous",
            "score": t.get("score", 0),
            "target_job": t.get("poste", ""),
            "top_skills": t.get("keywords", [])[:3],
            "photo_url": cand.get("photo_url", ""),
            "email": t.get("candidateEmail", "")
        })
    return jsonify(result), 200



# ========== MASQUER UNE CANDIDATURE ==========
@app.route('/api/candidatures/hide/<candidature_id>', methods=['PATCH', 'OPTIONS'])
def hide_candidature(candidature_id):
    """Masque une candidature (soft delete)"""
    
    # Gestion CORS pour OPTIONS
    if request.method == 'OPTIONS':
        response = jsonify({'message': 'OK'})
        response.headers.add('Access-Control-Allow-Origin', '*')
        response.headers.add('Access-Control-Allow-Headers', 'Content-Type, Authorization')
        response.headers.add('Access-Control-Allow-Methods', 'PATCH, OPTIONS')
        return response
    
    # Vérifier si l'ID est valide
    if not ObjectId.is_valid(candidature_id):
        return jsonify({"error": "ID invalide"}), 400
    
    try:
        # Mettre à jour la candidature avec hidden = True
        result = mongo.db.candidatures.update_one(
            {"_id": ObjectId(candidature_id)},
            {"$set": {"hidden": True, "hidden_at": datetime.utcnow()}}
        )
        
        if result.modified_count:
            return jsonify({"message": "Candidature masquée avec succès"}), 200
        else:
            return jsonify({"error": "Candidature non trouvée"}), 404
            
    except Exception as e:
        print(f"❌ Erreur lors du masquage: {e}")
        return jsonify({"error": str(e)}), 500



# Candidates admin
@app.route('/api/admin/candidates', methods=['GET'])
def get_all_candidates():
    candidates = list(mongo.db.candidats.find({}, {"password": 0}))
    for c in candidates:
        c['_id'] = str(c['_id'])
    return jsonify(candidates), 200

@app.route('/api/admin/candidates/add', methods=['POST'])
def add_candidate():
    try:
        cv_filename = ""
        photo_filename = ""
        
        # Récupérer les données
        if request.is_json:
            data = request.get_json()
            first_name = data.get('first_name')
            last_name = data.get('last_name')
            email = data.get('email')
            password = data.get('password')
            status = data.get('status', 'pending')
            cv_filename = data.get('cv_path', '')
            photo_filename = data.get('photo_url', '')
        else:
            first_name = request.form.get('first_name')
            last_name = request.form.get('last_name')
            email = request.form.get('email')
            password = request.form.get('password')
            status = request.form.get('status', 'pending')
            
            # Gérer le CV
            cv_file = request.files.get('cv')
            if cv_file and cv_file.filename:
                cv_filename = secure_filename(f"{email}_cv_{datetime.utcnow().timestamp()}.pdf")
                cv_file.save(os.path.join('uploads/cvs', cv_filename))
            
            # ⭐ GESTION CORRIGÉE DE LA PHOTO ⭐
            photo_file = request.files.get('photo')
            
            if photo_file and photo_file.filename:
                # Si une photo est uploadée
                ext = photo_file.filename.rsplit('.', 1)[1].lower()
                photo_filename = secure_filename(f"{email}_photo_{datetime.utcnow().timestamp()}.{ext}")
                photo_file.save(os.path.join('uploads/photos', photo_filename))
            else:
                # ⭐ AUCUNE PHOTO UPLOADÉE - NE RIEN FAIRE, LAISSER vide ou utiliser avatar par défaut
                # On ne fait rien ici, on utilisera un avatar par défaut dans le frontend si nécessaire
                photo_filename = ''
        
        # Validation
        if not first_name or not email:
            return jsonify({"error": "First name and email are required"}), 400
        
        if mongo.db.candidats.find_one({"email": email}):
            return jsonify({"error": "Email already exists"}), 400
        
        hashed = generate_password_hash(password) if password else ""
        
        candidate = {
            "first_name": first_name,
            "last_name": last_name,
            "email": email,
            "password": hashed,
            "status": status,
            "role": "candidate",
            "cv_path": cv_filename,
            "photo_url": photo_filename if photo_filename else "",
            "created_at": datetime.utcnow()
        }
        
        result = mongo.db.candidats.insert_one(candidate)
        
        return jsonify({
            "message": "Candidate added", 
            "id": str(result.inserted_id),
            "cvPath": cv_filename
        }), 201
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500
@app.route('/api/debug/candidates-all', methods=['GET'])
def debug_candidates_all():
    """Route debug pour voir tous les candidats"""
    candidates = list(mongo.db.candidats.find({}, {"password": 0}))
    for c in candidates:
        c['_id'] = str(c['_id'])
    return jsonify({
        "count": len(candidates),
        "candidates": candidates
    }), 200

@app.route('/api/candidatures/direct', methods=['POST'])
def create_candidature_direct():
    """Crée une candidature directement"""
    try:
        data = request.json
        
        # Nettoyer les données
        candidature = {
            "candidateEmail": data.get('candidateEmail'),
            "candidateName": data.get('candidateName', ''),
            "company": data.get('company', ''),
            "companyId": data.get('companyId', ''),
            "posteId": data.get('posteId', ''),
            "poste": data.get('poste', 'Manual Application'),
            "status": data.get('status', 'pending'),
            "date": data.get('date', datetime.utcnow().isoformat()),
            "cvPath": data.get('cvPath', ''),
            "score": data.get('score', 0),
            "created_at": datetime.utcnow()
        }
        
        result = mongo.db.candidatures.insert_one(candidature)
        
        return jsonify({
            "message": "Application created",
            "id": str(result.inserted_id)
        }), 201
        
    except Exception as e:
        print(f"❌ Error in create_candidature_direct: {e}")
        return jsonify({"error": str(e)}), 500
    

@app.route('/api/admin/fix-missing-applications', methods=['POST'])
def fix_missing_applications():
    """Crée des candidatures pour les candidats qui n'en ont pas"""
    try:
        # Récupérer tous les candidats
        all_candidates = list(mongo.db.candidats.find())
        
        # Récupérer tous les emails qui ont déjà une candidature
        existing_app_emails = set()
        for app in mongo.db.candidatures.find({}, {"candidateEmail": 1}):
            if app.get('candidateEmail'):
                existing_app_emails.add(app['candidateEmail'])
        
        created_count = 0
        missing_candidates = []
        
        for candidate in all_candidates:
            email = candidate.get('email')
            if email and email not in existing_app_emails:
                # Créer une candidature pour ce candidat
                new_application = {
                    "candidateEmail": email,
                    "candidateName": f"{candidate.get('first_name', '')} {candidate.get('last_name', '')}".strip() or email.split('@')[0],
                    "company": "Manual Entry",
                    "companyId": "",
                    "posteId": "",
                    "poste": "Candidature manuelle",
                    "status": candidate.get('status', 'pending'),
                    "date": candidate.get('created_at', datetime.utcnow()),
                    "created_at": candidate.get('created_at', datetime.utcnow()),
                    "cvPath": candidate.get('cv_path', ''),
                    "score": 0
                }
                mongo.db.candidatures.insert_one(new_application)
                created_count += 1
                missing_candidates.append(email)
        
        return jsonify({
            "message": f"{created_count} candidatures créées",
            "fixed_emails": missing_candidates
        }), 200
        
    except Exception as e:
        print(f"❌ Erreur: {e}")
        return jsonify({"error": str(e)}), 500
    


@app.route('/api/admin/candidates/update/<id>', methods=['PUT'])
def update_candidate(id):
    first_name = request.form.get('first_name')
    last_name = request.form.get('last_name')
    email = request.form.get('email')
    status = request.form.get('status')
    update_data = {}
    if first_name:
        update_data["first_name"] = first_name
    if last_name:
        update_data["last_name"] = last_name
    if email:
        update_data["email"] = email
    if status:
        update_data["status"] = status
    if 'cv' in request.files:
        cv = request.files['cv']
        if cv and cv.filename:
            cv_filename = secure_filename(f"{email}_cv_{datetime.utcnow().timestamp()}.pdf")
            cv.save(os.path.join('uploads/cvs', cv_filename))
            update_data["cv_path"] = cv_filename
    if 'photo' in request.files:
        photo = request.files['photo']
        if photo and photo.filename:
            ext = photo.filename.rsplit('.', 1)[1].lower()
            photo_filename = secure_filename(f"{email}_photo_{datetime.utcnow().timestamp()}.{ext}")
            photo.save(os.path.join('uploads/photos', photo_filename))
            update_data["photo_url"] = photo_filename
    if not update_data:
        return jsonify({"error": "No data to update"}), 400
    mongo.db.candidats.update_one({"_id": ObjectId(id)}, {"$set": update_data})
    return jsonify({"message": "Updated"}), 200

@app.route('/api/admin/candidates/delete/<id>', methods=['DELETE'])
def delete_candidate(id):
    cand = mongo.db.candidats.find_one({"_id": ObjectId(id)})
    if cand:
        if cand.get('cv_path'):
            p = os.path.join('uploads/cvs', cand['cv_path'])
            if os.path.exists(p):
                os.remove(p)
        if cand.get('photo_url'):
            p = os.path.join('uploads/photos', cand['photo_url'])
            if os.path.exists(p):
                os.remove(p)
    mongo.db.candidats.delete_one({"_id": ObjectId(id)})
    return jsonify({"message": "Deleted"}), 200

@app.route('/api/admin/add-comment', methods=['POST'])
def add_admin_comment():
    data = request.json
    mongo.db.candidats.update_one(
        {"email": data.get('candidateEmail')},
        {"$push": {"admin_comments": {"admin": data.get('adminName', 'Admin'), "comment": data.get('comment'), "date": datetime.utcnow()}}}
    )
    return jsonify({"message": "Comment added"}), 200

# Messages
@app.route('/api/admin/messages', methods=['GET'])
def get_messages():
    try:
        messages = list(mongo.db.messages.find({
            "$or": [
                {"type": {"$exists": False}},
                {"recipient": {"$exists": False}},
                {"old_status": {"$exists": False}},
                {"new_status": {"$exists": False}}
            ]
        }).sort("date", -1))
        for msg in messages:
            msg['_id'] = str(msg['_id'])
            if 'date' in msg and hasattr(msg['date'], 'isoformat'):
                msg['date'] = msg['date'].isoformat()
            if 'status' not in msg:
                msg['status'] = 'unread'
        return jsonify(messages), 200
    except Exception as e:
        print(f"❌ Error get_messages: {e}")
        return jsonify([]), 200

@app.route('/api/admin/messages/read/<id>', methods=['PUT'])
def mark_message_read(id):
    mongo.db.messages.update_one({"_id": ObjectId(id)}, {"$set": {"status": "read"}})
    return jsonify({"message": "OK"}), 200

@app.route('/api/admin/messages/delete/<id>', methods=['DELETE'])
def delete_message(id):
    mongo.db.messages.delete_one({"_id": ObjectId(id)})
    return jsonify({"message": "Deleted"}), 200

# Conversations
@app.route('/api/admin/conversations', methods=['GET'])
def get_admin_conversations():
    pipeline = [
        {"$match": {"$or": [{"from": "candidate", "to": "admin"}, {"from": "admin", "toEmail": {"$exists": True}}]}},
        {"$sort": {"date": -1}},
        {"$group": {
            "_id": {"$ifNull": ["$fromEmail", "$toEmail"]},
            "candidateEmail": {"$first": {"$ifNull": ["$fromEmail", "$toEmail"]}},
            "candidateName": {"$first": {"$ifNull": ["$fromName", "$toName"]}},
            "lastMessage": {"$first": "$message"},
            "lastMessageDate": {"$first": "$date"},
            "unreadCount": {"$sum": {"$cond": [{"$and": [{"$eq": ["$to", "admin"]}, {"$eq": ["$read", False]}]}, 1, 0]}}
        }},
        {"$sort": {"lastMessageDate": -1}}
    ]
    convs = list(mongo.db.messages.aggregate(pipeline))
    for c in convs:
        c['_id'] = str(c['_id'])
        if 'lastMessageDate' in c and hasattr(c['lastMessageDate'], 'isoformat'):
            c['lastMessageDate'] = c['lastMessageDate'].isoformat()
    return jsonify(convs), 200

@app.route('/api/admin/conversation/<email>', methods=['GET'])
def get_admin_conversation(email):
    msgs = list(mongo.db.messages.find({
        "$or": [
            {"fromEmail": email, "to": "admin"},
            {"from": "admin", "toEmail": email}
        ]
    }).sort("date", 1))
    for m in msgs:
        m['_id'] = str(m['_id'])
        if 'date' in m and hasattr(m['date'], 'isoformat'):
            m['date'] = m['date'].isoformat()
        m['sender'] = 'admin' if m.get('from') == 'admin' else 'candidate'
        m['senderName'] = 'Administrator' if m['sender'] == 'admin' else m.get('fromName', 'Candidate')
    cand = mongo.db.candidats.find_one({"email": email})
    cand_name = f"{cand.get('first_name', '')} {cand.get('last_name', '')}".strip() if cand else email
    mongo.db.messages.update_many({"from": "candidate", "to": "admin", "read": False}, {"$set": {"read": True}})
    return jsonify({"candidateEmail": email, "candidateName": cand_name, "messages": msgs}), 200






@app.route('/api/admin/send-message', methods=['POST'])
def admin_send_message():
    data = request.json
    msg = {
        "from": "admin",
        "fromName": data.get('fromName'),
        "to": "candidate",
        "toEmail": data.get('toEmail'),
        "message": data.get('message'),
        "date": datetime.utcnow(),
        "read": False
    }
    mongo.db.messages.insert_one(msg)
    mongo.db.notifications.insert_one({
        "type": "new_message",
        "recipient": "candidate",
        "recipientEmail": data.get('toEmail'),
        "title": "New message",
        "message": f"Reply: {data.get('message')[:50]}...",
        "status": "unread",
        "date": datetime.utcnow()
    })
    cand = mongo.db.candidats.find_one({"email": data.get('toEmail')})
    cand_name = f"{cand.get('first_name', '')} {cand.get('last_name', '')}".strip() if cand else "Candidate"
    send_real_email(data.get('toEmail'), "New message from administrator", f"<html><body><h2>Hello {cand_name}</h2><p>{data.get('message')}</p><a href='http://localhost:4200/login-candidate'>Reply</a></body></html>")
    return jsonify({"message": "Message sent"}), 201

@app.route('/api/admin/notifications', methods=['GET'])
def get_admin_notifications():
    notifs = list(mongo.db.notifications.find({"recipient": "admin"}).sort("date", -1))
    for n in notifs:
        n['_id'] = str(n['_id'])
        if 'date' in n and hasattr(n['date'], 'isoformat'):
            n['date'] = n['date'].isoformat()
    return jsonify(notifs), 200

@app.route('/api/admin/notifications/mark-all-read', methods=['PUT'])
def mark_admin_notifs_read():
    mongo.db.notifications.update_many({"recipient": "admin", "status": "unread"}, {"$set": {"status": "read"}})
    return jsonify({"message": "OK"}), 200

# Reports
@app.route('/api/admin/reports', methods=['GET'])
def get_reports():
    try:
        # ========== STATISTIQUES DE BASE ==========
        total_cvs = mongo.db.candidatures.count_documents({})
        validated = mongo.db.candidatures.count_documents({"status": "sent_to_client"})
        conversion_rate = round(validated / total_cvs * 100, 1) if total_cvs > 0 else 0
        
        # Score moyen
        avg_pipeline = [{"$group": {"_id": None, "avg": {"$avg": "$score"}}}]
        avg_result = list(mongo.db.candidatures.aggregate(avg_pipeline))
        average_score = round(avg_result[0]["avg"], 1) if avg_result else 0
        
        # ========== TOP SKILLS ==========
        skills_pipeline = [
            {"$match": {"keywords": {"$exists": True, "$ne": []}}},
            {"$unwind": "$keywords"},
            {"$group": {"_id": "$keywords", "count": {"$sum": 1}}},
            {"$sort": {"count": -1}},
            {"$limit": 10}
        ]
        top_skills = [{"skill": s["_id"], "count": s["count"]} for s in list(mongo.db.postes.aggregate(skills_pipeline))]
        
        # ========== RÉPARTITION PAR RÉGION (CORRIGÉ) ==========
        # Utiliser la collection candidatures et le champ company ou candidats.nationalite
        region_pipeline = [
            {"$match": {"nationalite": {"$exists": True, "$nin": [None, "", "Not specified"]}}},
            {"$group": {"_id": "$nationalite", "count": {"$sum": 1}}},
            {"$sort": {"count": -1}},
            {"$limit": 5}
        ]
        region_result = list(mongo.db.candidats.aggregate(region_pipeline))
        region_data = [{"region": r["_id"], "count": r["count"]} for r in region_result] if region_result else []
        
        # ========== TENDANCE DES SCORES ==========
        trend_pipeline = [
            {"$match": {"date": {"$exists": True}, "score": {"$exists": True}}},
            {"$group": {
                "_id": {"$dateToString": {"format": "%Y-%m", "date": "$date"}},
                "avg_score": {"$avg": "$score"}
            }},
            {"$sort": {"_id": 1}},
            {"$limit": 12}
        ]
        score_trend = [{"month": t["_id"], "avg_score": round(t["avg_score"], 1)} for t in list(mongo.db.candidatures.aggregate(trend_pipeline))]
        
        # ========== TAUX DE RÉPONSE ==========
        total_msgs = mongo.db.messages.count_documents({"to": "admin"})
        read_msgs = mongo.db.messages.count_documents({"to": "admin", "read": True})
        admin_response_rate = round(read_msgs / total_msgs * 100, 1) if total_msgs > 0 else 0
        
        # ========== CONVERSION PAR POSTE ==========
        jobs_conv = list(mongo.db.postes.aggregate([
            {"$lookup": {"from": "candidatures", "localField": "_id", "foreignField": "posteId", "as": "cands"}},
            {"$project": {
                "title": 1,
                "total": {"$size": "$cands"},
                "validated": {"$size": {"$filter": {"input": "$cands", "as": "c", "cond": {"$eq": ["$$c.status", "sent_to_client"]}}}}
            }},
            {"$addFields": {"rate": {"$cond": [{"$eq": ["$total", 0]}, 0, {"$multiply": [{"$divide": ["$validated", "$total"]}, 100]}]}}},
            {"$sort": {"rate": -1}},
            {"$limit": 5}
        ]))
        top_jobs_conversion = [{"title": j["title"], "total": j["total"], "validated": j["validated"], "rate": round(j["rate"], 1)} for j in jobs_conv]
        
        # ========== RÉPARTITION PAR CONTRAT ==========
        contract_result = list(mongo.db.postes.aggregate([{"$group": {"_id": "$type", "count": {"$sum": 1}}}]))
        contract_data = {c["_id"]: c["count"] for c in contract_result if c["_id"]}
        
        # ========== TEMPS MOYEN DE TRAITEMENT ==========
        processing_days = []
        for cand in mongo.db.candidatures.find({"status": "sent_to_client", "date": {"$exists": True}, "sent_to_client_at": {"$exists": True}}):
            if cand.get("date") and cand.get("sent_to_client_at"):
                start = cand["date"]
                end = cand["sent_to_client_at"]
                if isinstance(start, datetime) and isinstance(end, datetime):
                    days = (end - start).days
                    if days >= 0:
                        processing_days.append(days)
        avg_processing_days = round(sum(processing_days) / len(processing_days), 1) if processing_days else 0
        
        # ========== RÉPONSE ==========
        return jsonify({
            "totalCVs": total_cvs,
            "validatedCount": validated,
            "conversionRate": conversion_rate,
            "averageScore": average_score,
            "avgProcessingDays": avg_processing_days,
            "topSkills": top_skills,
            "regionData": region_data,
            "scoreTrend": score_trend,
            "adminResponseRate": admin_response_rate,
            "topJobsConversion": top_jobs_conversion,
            "contractData": contract_data,
            "monthlyData": [],
            "statusData": {},
            "topCompanies": []
        }), 200
        
    except Exception as e:
        print(f"❌ Erreur dans get_reports: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({
            "totalCVs": 0, "validatedCount": 0, "conversionRate": 0,
            "averageScore": 0, "avgProcessingDays": 0, "topSkills": [],
            "regionData": [], "scoreTrend": [], "adminResponseRate": 0,
            "topJobsConversion": [], "contractData": {}, "monthlyData": [],
            "statusData": {}, "topCompanies": []
        }), 200




@app.route('/api/test-candidate-origins', methods=['GET'])
def test_candidate_origins():
    """Test simple pour voir les données"""
    total = mongo.db.candidats.count_documents({})
    with_nationalite = mongo.db.candidats.count_documents({"nationalite": {"$nin": [None, "", "All candidates", "Not specified"]}})
    
    sample = list(mongo.db.candidats.find({}, {"nationalite": 1, "first_name": 1}).limit(5))
    
    return jsonify({
        "total_candidates": total,
        "candidates_with_nationalite": with_nationalite,
        "sample": sample
    }), 200




@app.route('/api/admin/kpi-advanced', methods=['GET'])
def get_advanced_kpi():
    """
    Retourne les KPIs avancés pour le dashboard
    """
    try:
        # Temps moyen d'embauche (entre candidature et envoi au client)
        processing_days = []
        for cand in mongo.db.candidatures.find({
            "status": "sent_to_client",
            "date": {"$exists": True},
            "sent_to_client_at": {"$exists": True}
        }):
            start = cand.get("date")
            end = cand.get("sent_to_client_at")
            if isinstance(start, datetime) and isinstance(end, datetime):
                days = (end - start).days
                if days >= 0:
                    processing_days.append(days)
        
        avg_hiring_days = round(sum(processing_days) / len(processing_days), 1) if processing_days else 0
        
        # Taux de conversion global
        total_cvs = mongo.db.candidatures.count_documents({})
        validated = mongo.db.candidatures.count_documents({"status": "sent_to_client"})
        conversion_rate = round(validated / total_cvs * 100, 1) if total_cvs > 0 else 0
        
        # Coût par embauche estimé
        total_candidates = mongo.db.candidats.count_documents({})
        cost_per_hire = round(validated * 250 / max(validated, 1), 0) if validated > 0 else 0
        
        return jsonify({
            "avgHiringDays": avg_hiring_days,
            "conversionRate": conversion_rate,
            "costPerHire": cost_per_hire
        }), 200
        
    except Exception as e:
        print(f"Erreur get_advanced_kpi: {e}")
        return jsonify({"avgHiringDays": 0, "conversionRate": 0, "costPerHire": 0}), 200
    


@app.route('/api/admin/ai-insights', methods=['GET'])
def get_ai_insights():
    """
    Retourne les insights IA pour le dashboard
    """
    try:
        # Meilleur poste du jour (celui avec le plus de candidatures récentes)
        last_7_days = datetime.utcnow() - timedelta(days=7)
        best_job_pipeline = [
            {"$match": {"date": {"$gte": last_7_days}}},
            {"$group": {"_id": "$poste", "count": {"$sum": 1}}},
            {"$sort": {"count": -1}},
            {"$limit": 1}
        ]
        best_job_result = list(mongo.db.candidatures.aggregate(best_job_pipeline))
        
        best_job = None
        if best_job_result and best_job_result[0].get('_id'):
            best_job = {
                "title": best_job_result[0]['_id'],
                "matchCount": best_job_result[0]['count']
            }
        
        # Baisse des candidatures (comparaison semaine précédente)
        last_week = datetime.utcnow() - timedelta(days=7)
        two_weeks_ago = datetime.utcnow() - timedelta(days=14)
        
        this_week_count = mongo.db.candidatures.count_documents({"date": {"$gte": last_week}})
        last_week_count = mongo.db.candidatures.count_documents({
            "date": {"$gte": two_weeks_ago, "$lt": last_week}
        })
        
        drop_alert = False
        drop_percent = 0
        recommended_action = "Continue monitoring applications."
        
        if last_week_count > 0:
            drop_percent = round(((last_week_count - this_week_count) / last_week_count) * 100, 1)
            if drop_percent > 20:
                drop_alert = True
                recommended_action = "Boost job postings on social media and LinkedIn."
        
        return jsonify({
            "bestJob": best_job,
            "dropAlert": drop_alert,
            "dropPercent": drop_percent,
            "recommendedAction": recommended_action
        }), 200
        
    except Exception as e:
        print(f"Erreur get_ai_insights: {e}")
        return jsonify({
            "bestJob": {"title": "Full Stack Developer", "matchCount": 12},
            "dropAlert": False,
            "dropPercent": 0,
            "recommendedAction": "Review pending candidates for Backend roles."
        }), 200
    
@app.route('/api/admin/matching-trend', methods=['GET'])
def get_matching_trend():
    """
    Retourne la tendance des scores de matching sur les 6 derniers mois
    """
    try:
        six_months_ago = datetime.utcnow() - timedelta(days=180)
        
        trend_pipeline = [
            {"$match": {
                "date": {"$gte": six_months_ago},
                "score": {"$exists": True}
            }},
            {"$group": {
                "_id": {
                    "$dateToString": {"format": "%Y-%m", "date": "$date"}
                },
                "avg_score": {"$avg": "$score"}
            }},
            {"$sort": {"_id": 1}}
        ]
        
        results = list(mongo.db.candidatures.aggregate(trend_pipeline))
        
        trend_data = []
        for r in results:
            trend_data.append({
                "month": r["_id"],
                "score": round(r["avg_score"], 1)
            })
        
        return jsonify(trend_data), 200
        
    except Exception as e:
        print(f"Erreur get_matching_trend: {e}")
        return jsonify([]), 200


@app.route('/api/admin/notifications/recent', methods=['GET'])
def get_recent_notifications():
    """
    Retourne les notifications récentes formatées pour l'overview
    """
    try:
        notifs = list(mongo.db.notifications.find({
            "recipient": "admin"
        }).sort("date", -1).limit(10))
        
        formatted = []
        icons_map = {
            "new_candidature": "fas fa-user-plus",
            "new_message": "fas fa-envelope",
            "new_feedback": "fas fa-star",
            "questionnaire_answered": "fas fa-file-alt",
            "client_notified": "fas fa-paper-plane",
            "candidature_viewed": "fas fa-eye"
        }
        
        for n in notifs:
            formatted.append({
                "id": str(n["_id"]),
                "message": n.get("message", ""),
                "icon": icons_map.get(n.get("type", ""), "fas fa-bell"),
                "date": n.get("date", datetime.utcnow()).isoformat(),
                "status": n.get("status", "unread")
            })
        
        return jsonify(formatted), 200
        
    except Exception as e:
        print(f"Erreur get_recent_notifications: {e}")
        return jsonify([]), 200
    

@app.route('/api/admin/today-interviews', methods=['GET'])
def get_today_interviews():
    """
    Retourne les entretiens programmés pour aujourd'hui
    """
    try:
        today = datetime.utcnow().strftime("%Y-%m-%d")
        
        candidatures = list(mongo.db.candidatures.find({
            "status": "interview_scheduled",
            "interview.date": today
        }))
        
        result = []
        for c in candidatures:
            interview = c.get("interview", {})
            result.append({
                "candidateName": c.get("candidateName", "Unknown"),
                "jobTitle": c.get("poste", "Unknown"),
                "date": interview.get("date", ""),
                "time": interview.get("time", ""),
                "type": interview.get("type", "Online"),
                "link": interview.get("link", "")
            })
        
        return jsonify(result), 200
        
    except Exception as e:
        print(f"Erreur get_today_interviews: {e}")
        return jsonify([]), 200
    

@app.route('/api/admin/status-distribution', methods=['GET'])
def get_status_distribution():
    """
    Retourne la répartition des candidatures par statut pour le graphique
    """
    try:
        statuses = ["pending", "to_contact", "interview_scheduled", "sent_to_client", "archived"]
        status_labels = {
            "pending": "En attente",
            "to_contact": "À contacter",
            "interview_scheduled": "Entretien",
            "sent_to_client": "Envoyé client",
            "archived": "Archivé"
        }
        
        data = []
        for status in statuses:
            count = mongo.db.candidatures.count_documents({"status": status})
            data.append({
                "status": status_labels.get(status, status),
                "value": count
            })
        
        return jsonify(data), 200
        
    except Exception as e:
        print(f"Erreur get_status_distribution: {e}")
        return jsonify([]), 200
    











# Dans app.py, remplacez la route get_top_companies par :

@app.route('/api/admin/top-companies', methods=['GET'])
def get_top_companies():
    """
    Retourne le classement des entreprises avec le plus de postes actifs
    """
    try:
        # Pipeline pour compter les postes par entreprise
        pipeline = [
            {"$match": {"status": "active"}},
            {"$group": {
                "_id": {
                    "companyId": "$companyId",
                    "companyName": "$companyName"
                },
                "totalJobs": {"$sum": 1}
            }},
            {"$sort": {"totalJobs": -1}},
            {"$limit": 5}
        ]
        
        result = list(mongo.db.postes.aggregate(pipeline))
        
        # Formater la réponse
        top_companies = []
        for item in result:
            company_id = item["_id"]["companyId"]
            company_name = item["_id"]["companyName"] or "Entreprise"
            
            # Récupérer les infos supplémentaires de l'entreprise
            company_info = {}
            if company_id and ObjectId.is_valid(company_id):
                company = mongo.db.entreprises.find_one({"_id": ObjectId(company_id)})
                if company:
                    company_info = {
                        "sector": company.get("secteur") or company.get("sector", "Non spécifié"),
                        "logo": company.get("logo", "assets/default-logo.png"),
                        "location": company.get("location", "Non spécifiée")
                    }
            
            top_companies.append({
                "name": company_name,
                "count": item["totalJobs"],
                "sector": company_info.get("sector", "Non spécifié"),
                "logo": company_info.get("logo", "assets/default-logo.png"),
                "location": company_info.get("location", "Non spécifiée"),
                "companyId": str(company_id) if company_id else None
            })
        
        return jsonify(top_companies), 200
        
    except Exception as e:
        print(f"❌ Erreur dans get_top_companies: {e}")
        import traceback
        traceback.print_exc()
        return jsonify([]), 200


@app.route('/api/admin/import-cvs', methods=['POST'])
def import_cvs():
    """Import multiple CVs and create candidates automatically"""
    try:
        files = request.files.getlist('cvs')
        if not files:
            return jsonify({"error": "No files uploaded"}), 400
        
        created_candidates = []
        errors = []
        
        for file in files:
            if not file.filename.endswith('.pdf'):
                errors.append(f"{file.filename}: Format non supporté (PDF requis)")
                continue
            
            # Extraire le nom du fichier sans extension
            base_name = os.path.splitext(file.filename)[0]
            parts = base_name.split('_')
            
            # Essayer d'extraire nom et prénom du nom du fichier
            if len(parts) >= 2:
                first_name = parts[0]
                last_name = parts[1]
            else:
                first_name = base_name
                last_name = ""
            
            # Générer un email unique
            email = f"{base_name.lower().replace(' ', '_')}_{datetime.utcnow().timestamp()}@candidate.com"
            
            # Sauvegarder le CV
            cv_filename = secure_filename(f"{email}_cv_{datetime.utcnow().timestamp()}.pdf")
            file.save(os.path.join('uploads/cvs', cv_filename))
            
            # Créer le candidat
            candidate = {
                "first_name": first_name,
                "last_name": last_name,
                "email": email,
                "password": generate_password_hash("default123"),
                "role": "candidate",
                "cv_path": cv_filename,
                "photo_url": "",
                "status": "pending",
                "created_at": datetime.utcnow()
            }
            
            # Vérifier si l'email existe déjà
            existing = mongo.db.candidats.find_one({"email": email})
            if existing:
                errors.append(f"{file.filename}: Email déjà existant")
                continue
            
            result = mongo.db.candidats.insert_one(candidate)
            created_candidates.append({
                "filename": file.filename,
                "email": email,
                "id": str(result.inserted_id)
            })
        
        message = f"{len(created_candidates)} CV(s) importés avec succès"
        if errors:
            message += f". {len(errors)} erreur(s): " + "; ".join(errors[:3])
        
        return jsonify({
            "message": message,
            "created": created_candidates,
            "errors": errors
        }), 201
        
    except Exception as e:
        print(f"❌ Erreur import CVs: {e}")
        return jsonify({"error": str(e)}), 500





# Settings
@app.route('/api/admin/settings/change-password', methods=['POST'])
def change_admin_password():
    data = request.json
    email = request.args.get('email') or data.get('email')
    if email:
        mongo.db.Users.update_one({"email": email, "role": "admin"}, {"$set": {"password": generate_password_hash(data.get('newPassword'))}})
    return jsonify({"message": "Password changed"}), 200

@app.route('/api/admin/settings/smtp', methods=['POST'])
def save_smtp_config():
    return jsonify({"message": "SMTP config saved"}), 200

@app.route('/api/admin/settings/test-email', methods=['POST'])
def test_email():
    data = request.json
    to = data.get('to')
    if to:
        send_real_email(to, "SMTP Test", "<p>This is a test of the SMTP configuration.</p>")
    return jsonify({"message": "Test email sent"}), 200

@app.route('/api/admin/settings/general', methods=['POST'])
def save_general_settings():
    return jsonify({"message": "General settings saved"}), 200

# User Management
@app.route('/api/admin/users', methods=['GET'])
def get_all_users():
    users = list(mongo.db.Users.find({}, {"password": 0}))
    for u in users:
        u['_id'] = str(u['_id'])
    return jsonify(users), 200

@app.route('/api/admin/users/add', methods=['POST'])
def add_user():
    data = request.json
    email = data.get('email')
    password = data.get('password')
    if not email or not password:
        return jsonify({"error": "Email and password required"}), 400
    if mongo.db.Users.find_one({"email": email}):
        return jsonify({"error": "Email already exists"}), 400
    mongo.db.Users.insert_one({
        "email": email,
        "password": generate_password_hash(password),
        "role": data.get('role', 'admin'),
        "nom": data.get('nom', ''),
        "image": "default-avatar.png"
    })
    return jsonify({"message": "User added"}), 201

@app.route('/api/admin/users/update/<id>', methods=['PUT'])
def update_user(id):
    data = request.json
    update = {"nom": data.get('nom'), "role": data.get('role')}
    if data.get('password'):
        update["password"] = generate_password_hash(data.get('password'))
    mongo.db.Users.update_one({"_id": ObjectId(id)}, {"$set": update})
    return jsonify({"message": "User updated"}), 200

@app.route('/api/admin/users/delete/<id>', methods=['DELETE'])
def delete_user(id):
    mongo.db.Users.delete_one({"_id": ObjectId(id)})
    return jsonify({"message": "User deleted"}), 200

@app.route('/api/admin/candidature/viewed', methods=['POST'])
def candidature_viewed():
    data = request.json
    candidature_id = data.get('candidatureId')
    candidate_email = data.get('candidateEmail')
    if not candidature_id or not candidate_email:
        return jsonify({"error": "Missing data"}), 400
    mongo.db.notifications.insert_one({
        "recipient": "candidate",
        "recipientEmail": candidate_email,
        "type": "candidature_viewed",
        "title": "Your application has been viewed",
        "message": "The administrator has reviewed your file. You will receive an update soon.",
        "status": "unread",
        "date": datetime.utcnow(),
        "candidatureId": candidature_id
    })
    candidate = mongo.db.candidats.find_one({"email": candidate_email})
    candidate_name = f"{candidate.get('first_name', '')} {candidate.get('last_name', '')}".strip() or "Candidate"
    subject = "Your application has been viewed"
    body = f"<html><body><h2>Hello {candidate_name},</h2><p>The administrator has viewed your application.</p><p>You will receive an update soon.</p><a href='http://localhost:4200/login-candidate'>Track my application</a></body></html>"
    send_real_email(candidate_email, subject, body)
    return jsonify({"message": "View notification sent"}), 200

@app.route('/api/funnel-stats', methods=['GET'])
def funnel_stats():
    interviews = mongo.db.candidatures.count_documents({"status": "interview_scheduled"})
    hired = mongo.db.candidatures.count_documents({"status": "sent_to_client"})
    return jsonify({"interviews": interviews, "hired": hired}), 200



@app.route('/api/candidate-origins', methods=['GET'])
def candidate_origins():
    """Retourne les origines géographiques des candidats pour le graphique"""
    try:
        # Pipeline pour regrouper par nationalité (ou ville si disponible)
        pipeline = [
            {"$match": {
                "nationalite": {"$nin": [None, "", "All candidates", "Not specified"]}
            }},
            {"$group": {
                "_id": "$nationalite",
                "count": {"$sum": 1}
            }},
            {"$sort": {"count": -1}},
            {"$limit": 10}
        ]
        
        result = list(mongo.db.candidats.aggregate(pipeline))
        
        if not result:
            # Données par défaut si pas de nationalités
            return jsonify([
                {"city": "Tunisie", "count": 0, "percentage": 0},
                {"city": "France", "count": 0, "percentage": 0},
                {"city": "Maroc", "count": 0, "percentage": 0}
            ]), 200
        
        total = sum(r["count"] for r in result)
        
        # Mapping des nationalités vers les pays/villes
        country_map = {
            'Tunisienne': 'Tunisie',
            'Française': 'France', 
            'Marocaine': 'Maroc',
            'Algérienne': 'Algérie',
            'Italienne': 'Italie',
            'Espagnole': 'Espagne',
            'Allemande': 'Allemagne',
            'Sfax': 'Sfax',
            'Tunis': 'Tunis',
            'Sousse': 'Sousse'
        }
        
        output = []
        for r in result:
            city_name = country_map.get(r["_id"], r["_id"])
            percentage = round((r["count"] / total) * 100, 1) if total > 0 else 0
            output.append({
                "city": city_name,
                "count": r["count"],
                "percentage": percentage
            })
        
        return jsonify(output), 200
        
    except Exception as e:
        print(f"❌ Erreur candidate_origins: {e}")
        import traceback
        traceback.print_exc()
        # Retourner des données par défaut
        return jsonify([
            {"city": "Tunisie", "count": 0, "percentage": 0},
            {"city": "France", "count": 0, "percentage": 0},
            {"city": "Maroc", "count": 0, "percentage": 0}
        ]), 200




# Route de debug pour vérifier les données
@app.route('/api/debug/candidate-nationalities', methods=['GET'])
def debug_candidate_nationalities():
    """Vérifie quels champs nationalité existent"""
    candidates = list(mongo.db.candidats.find({}, {"nationalite": 1, "first_name": 1}))
    
    nationalites = []
    for c in candidates:
        nat = c.get('nationalite')
        if nat and nat not in nationalites:
            nationalites.append(nat)
    
    return jsonify({
        "total_candidates": len(candidates),
        "candidates_with_nationalite": len([c for c in candidates if c.get('nationalite')]),
        "nationalites_found": nationalites,
        "sample": candidates[:5]
    }), 200



@app.route('/api/admin/fix-nationalities', methods=['POST'])
def fix_nationalities():
    """Ajoute des nationalités de test aux candidats qui n'en ont pas"""
    try:
        candidates = list(mongo.db.candidats.find({"nationalite": {"$in": [None, "", "All candidates", "Not specified"]}}))
        
        nationalities = ["Tunisienne", "Française", "Marocaine", "Algérienne", "Italienne"]
        updated = 0
        
        for cand in candidates:
            random_nat = random.choice(nationalities)
            mongo.db.candidats.update_one(
                {"_id": cand["_id"]},
                {"$set": {"nationalite": random_nat}}
            )
            updated += 1
        
        return jsonify({"message": f"{updated} candidats mis à jour avec des nationalités"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    



@app.route('/api/upcoming-interviews', methods=['GET'])
def upcoming_interviews():
    now = datetime.utcnow().strftime("%Y-%m-%d")
    candidatures = list(mongo.db.candidatures.find({
        "status": "interview_scheduled",
        "interview.date": {"$gte": now}
    }).limit(5))
    result = []
    for c in candidatures:
        interview = c.get('interview', {})
        result.append({
            "candidateName": c.get('candidateName', ''),
            "jobTitle": c.get('poste', ''),
            "date": interview.get('date', ''),
            "time": interview.get('time', ''),
            "type": interview.get('type', 'Online')
        })
    return jsonify(result), 200

@app.route('/api/avg-score', methods=['GET'])
def avg_score():
    pipeline = [{"$group": {"_id": None, "avg": {"$avg": "$score"}}}]
    result = list(mongo.db.candidatures.aggregate(pipeline))
    avg = round(result[0]["avg"], 1) if result else 0
    return jsonify({"avgScore": avg}), 200

# Advanced candidature management
@app.route('/api/candidatures/<id>/comment', methods=['POST'])
def add_candidature_comment(id):
    data = request.json
    comment = {
        "text": data.get('comment'),
        "admin": data.get('adminName', 'Admin'),
        "date": datetime.utcnow(),
        "type": data.get('type', 'general')
    }
    mongo.db.candidatures.update_one({"_id": ObjectId(id)}, {"$push": {"internal_comments": comment}})
    return jsonify({"message": "Comment added"}), 200


@app.route('/api/candidatures/<id>/detailed-evaluation', methods=['POST'])
@jwt_required()
def save_detailed_evaluation(id):
    """Sauvegarde l'évaluation détaillée d'un candidat"""
    try:
        data = request.json
        evaluation = data.get('evaluation', [])
        
        if not ObjectId.is_valid(id):
            return jsonify({"error": "Invalid ID"}), 400
            
        result = mongo.db.candidatures.update_one(
            {"_id": ObjectId(id)},
            {"$set": {"detailed_evaluation": evaluation, "detailed_evaluation_date": datetime.utcnow()}}
        )
        
        if result.modified_count:
            return jsonify({"message": "Evaluation saved successfully"}), 200
        else:
            return jsonify({"error": "No changes made"}), 400
            
    except Exception as e:
        print(f"Error saving detailed evaluation: {e}")
        return jsonify({"error": str(e)}), 500
    



@app.route('/api/candidatures/<id>/schedule-interview', methods=['POST'])
def schedule_interview(id):
    data = request.json
    interview = {
        "date": data.get('date'),
        "time": data.get('time'),
        "type": data.get('type'),
        "link": data.get('link', ''),
        "location": data.get('location', '')
    }
    mongo.db.candidatures.update_one(
        {"_id": ObjectId(id)},
        {"$set": {"interview": interview, "status": "interview_scheduled"}}
    )
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(id)})
    candidate_email = candidature.get('candidateEmail')
    candidate_name = candidature.get('candidateName')
    subject_candidate = f"Interview scheduled for {candidature.get('poste')}"
    body_candidate = f"""
    <html><body><h2>Hello {candidate_name},</h2>
    <p>An interview has been scheduled for your application.</p>
    <p><strong>Date:</strong> {interview['date']} at {interview['time']}</p>
    <p><strong>Type:</strong> {interview['type']}</p>
    {f'<p><strong>Link:</strong> <a href="{interview["link"]}">{interview["link"]}</a></p>' if interview.get('link') else ''}
    {f'<p><strong>Address:</strong> {interview["location"]}</p>' if interview.get('location') else ''}
    <a href="http://localhost:4200/login-candidate">Go to dashboard</a>
    </body></html>
    """
    send_real_email(candidate_email, subject_candidate, body_candidate)
    if not candidature.get("evaluation_token"):
        token = hashlib.md5(f"{candidature['_id']}{datetime.utcnow()}".encode()).hexdigest()[:16]
        mongo.db.candidatures.update_one(
            {"_id": ObjectId(id)},
            {"$set": {"evaluation_token": token}}
        )
    else:
        token = candidature["evaluation_token"]
    job = mongo.db.postes.find_one({"_id": ObjectId(candidature.get('posteId'))})
    client_email = None
    if job and job.get('companyId'):
        company = mongo.db.entreprises.find_one({"_id": ObjectId(job['companyId'])})
        client_email = company.get('email')
    if client_email:
        eval_link = f"http://localhost:4200/evaluation/{token}"
        dossier_link = f"http://localhost:4200/company/dossier/{token}"
        subject_client = f"Interview scheduled for {candidature.get('candidateName')}"
        body_client = f"""
        <html><body>
        <p>An interview has been scheduled for <strong>{candidature.get('candidateName')}</strong> on <strong>{interview['date']}</strong> at <strong>{interview['time']}</strong>.</p>
        <p><a href="{dossier_link}">📄 View candidate's full dossier (CV + questionnaire)</a></p>
        <p><a href="{eval_link}">📝 Evaluation form (to be filled by candidate)</a></p>
        <p>Best regards,<br>The LCA team</p>
        </body></html>
        """
        send_real_email(client_email, subject_client, body_client)
    return jsonify({"message": "Interview scheduled and client notified"}), 200

@app.route('/api/candidatures/<id>/evaluation', methods=['POST'])
def add_evaluation(id):
    data = request.json
    evaluation = {
        "score": data.get('score'),
        "comment": data.get('comment'),
        "skills": data.get('skills', []),
        "admin": data.get('adminName', 'Admin'),
        "date": datetime.utcnow()
    }
    mongo.db.candidatures.update_one({"_id": ObjectId(id)}, {"$set": {"evaluation": evaluation, "status": "shortlisted"}})
    return jsonify({"message": "Evaluation saved"}), 200






# Send to client (with PDF)

@app.route('/api/candidatures/<id>/send-to-client', methods=['POST'])
def send_to_client(id):
    try:
        data = request.json or {}
        override_client_email = data.get('client_email')
        is_anonymous = data.get('is_anonymous', False)
        message_perso = data.get('message_perso', '')
        
        candidature = mongo.db.candidatures.find_one({"_id": ObjectId(id)})
        if not candidature:
            return jsonify({"error": "Application not found"}), 404
        
        # Générer un token client
        if not candidature.get('client_token'):
            token = hashlib.md5(f"{candidature['candidateEmail']}{id}{datetime.utcnow()}".encode()).hexdigest()[:16]
            mongo.db.candidatures.update_one(
                {"_id": ObjectId(id)},
                {"$set": {"client_token": token}}
            )
            client_token = token
        else:
            client_token = candidature['client_token']
        
        client_email = override_client_email
        if not client_email:
            job = mongo.db.postes.find_one({"_id": ObjectId(candidature.get('posteId'))}) if candidature.get('posteId') else None
            if job and job.get('companyId'):
                company = mongo.db.entreprises.find_one({"_id": ObjectId(job['companyId'])})
                client_email = company.get('email') if company else None
        
        if not client_email:
            return jsonify({"error": "Client email not found. Please provide an email."}), 400
        
        candidate_info = mongo.db.candidats.find_one({"email": candidature.get('candidateEmail')})
        if not candidate_info:
            return jsonify({"error": "Candidate not found"}), 404
        
        cv_details = {}
        if candidate_info.get('cv_path'):
            cv_full_path = os.path.join(app.config['UPLOAD_FOLDER'], candidate_info['cv_path'])
            if os.path.exists(cv_full_path):
                cv_details = extract_detailed_cv(cv_full_path)
        
        # Générer le PDF avec option anonyme
        pdf_bytes = generate_cv_pdf(candidature, candidate_info, cv_details, candidature.get('answers', []), is_anonymous)
        pdf_buffer = io.BytesIO(pdf_bytes) if pdf_bytes else None
        
        mongo.db.candidatures.update_one(
            {"_id": ObjectId(id)},
            {"$set": {"status": "sent_to_client", "sent_to_client_at": datetime.utcnow()}}
        )
        
        # ⭐ CORRECTION ICI : Nom du fichier sans le nom du candidat en mode anonyme
        if is_anonymous:
            # Utiliser un ID unique au lieu du nom
            unique_id = str(uuid.uuid4())[:8]
            pdf_filename = f"dossier_anonyme_{unique_id}.pdf"
        else:
            candidate_name_raw = candidature.get('candidateName', 'candidate')
            safe_name = re.sub(r'[^a-zA-Z0-9_]', '_', candidate_name_raw)
            pdf_filename = f"dossier_{safe_name}.pdf"
        
        client_link = f"http://localhost:4200/client/access/{client_token}"
        
        subject = f"Candidature - {candidature.get('poste')}"
        
        # Construction du message selon anonyme ou complet
        if is_anonymous:
            body_client = f"""
            <html>
            <body style="font-family: Arial, sans-serif;">
                <div style="background: #0f172a; padding: 20px; text-align: center;">
                    <h2 style="color: white; margin: 0;">📁 Dossier de candidature (version anonyme)</h2>
                </div>
                <div style="padding: 20px;">
                    <p>Bonjour,</p>
                    <p>Veuillez trouver ci-joint le dossier de candidature pour le poste de <strong>{candidature.get('poste')}</strong>.</p>
                    <div style="background: #f0fdfa; padding: 15px; border-radius: 10px; margin: 15px 0;">
                        <p style="margin: 5px 0;">🎯 <strong>Score IA:</strong> {candidature.get('score', '?')}%</p>
                        <p style="margin: 5px 0;">📊 <strong>Compétences clés:</strong> {', '.join(cv_details.get('competences', [])[:5])}</p>
                        <p style="margin: 5px 0;">💼 <strong>Expérience:</strong> {cv_details.get('max_experience', 0)} ans</p>
                    </div>
                    <p><a href="{client_link}" style="background: #00a6a6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">📄 Voir le dossier en ligne</a></p>
                    {f'<p><em>{message_perso}</em></p>' if message_perso else ''}
                    <p>Cordialement,<br>L'équipe LCA</p>
                </div>
            </body>
            </html>
            """
        else:
            candidate_name_raw = candidature.get('candidateName', 'candidate')
            body_client = f"""
            <html>
            <body style="font-family: Arial, sans-serif;">
                <div style="background: #0f172a; padding: 20px; text-align: center;">
                    <h2 style="color: white; margin: 0;">📄 Dossier de candidature complet</h2>
                </div>
                <div style="padding: 20px;">
                    <p>Bonjour,</p>
                    <p>Veuillez trouver ci-joint le dossier complet de <strong>{candidate_name_raw}</strong> pour le poste de <strong>{candidature.get('poste')}</strong>.</p>
                    <div style="background: #f0fdfa; padding: 15px; border-radius: 10px; margin: 15px 0;">
                        <p style="margin: 5px 0;">👤 <strong>Candidat:</strong> {candidate_name_raw}</p>
                        <p style="margin: 5px 0;">📧 <strong>Email:</strong> {candidate_info.get('email', 'Non renseigné')}</p>
                        <p style="margin: 5px 0;">🌍 <strong>Nationalité:</strong> {candidate_info.get('nationalite', 'Non renseignée')}</p>
                        <p style="margin: 5px 0;">🎯 <strong>Score IA:</strong> {candidature.get('score', '?')}%</p>
                        <p style="margin: 5px 0;">📊 <strong>Compétences:</strong> {', '.join(cv_details.get('competences', [])[:8])}</p>
                        <p style="margin: 5px 0;">💼 <strong>Expérience:</strong> {cv_details.get('max_experience', 0)} ans</p>
                    </div>
                    <p><a href="{client_link}" style="background: #00a6a6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">📄 Voir le dossier complet en ligne</a></p>
                    {f'<p><em>{message_perso}</em></p>' if message_perso else ''}
                    <p>Cordialement,<br>L'équipe LCA</p>
                </div>
            </body>
            </html>
            """
        
        if pdf_buffer:
            success = send_email_with_attachment(client_email, subject, body_client, pdf_buffer, pdf_filename)
            if not success:
                send_real_email(client_email, subject, body_client)
        else:
            send_real_email(client_email, subject, body_client)
        
        return jsonify({"message": f"Dossier {'anonyme' if is_anonymous else 'complet'} envoyé au client"}), 200
        
    except Exception as e:
        print(f"❌ Error in send_to_client: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Internal error: {str(e)}"}), 500



@app.route('/api/candidate/dossier/<candidature_id>', methods=['GET'])
def candidate_dossier_pdf(candidature_id):
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(candidature_id)})
    if not candidature:
        return jsonify({"error": "Application not found"}), 404
    candidate_info = mongo.db.candidats.find_one({"email": candidature.get('candidateEmail')})
    if not candidate_info:
        return jsonify({"error": "Candidate not found"}), 404
    cv_details = { }
    if candidate_info.get('cv_path'):
        cv_full_path = os.path.join(app.config['UPLOAD_FOLDER'], candidate_info['cv_path'])
        if os.path.exists(cv_full_path):
            cv_details = extract_detailed_cv(cv_full_path)
    pdf_bytes = generate_cv_pdf(candidature, candidate_info, cv_details, candidature.get('answers', []))
    return send_file(
        io.BytesIO(pdf_bytes),
        as_attachment=True,
        download_name=f"dossier_{candidature_id}.pdf",
        mimetype='application/pdf'
    )

@app.route('/api/admin/candidatures-with-details', methods=['GET'])
def get_candidatures_with_details():
    try:
        pipeline = [
            # ⭐ AJOUTEZ CETTE LIGNE POUR EXCLURE LES CANDIDATURES MASQUÉES ⭐
            {"$match": {"hidden": {"$ne": True}}},  # <-- AJOUTEZ CETTE LIGNE
            
            {"$lookup": {"from": "postes", "localField": "posteId", "foreignField": "_id", "as": "job"}},
            {"$unwind": {"path": "$job", "preserveNullAndEmptyArrays": True}},
            {"$lookup": {"from": "entreprises", "localField": "job.companyId", "foreignField": "_id", "as": "company_from_job"}},
            {"$unwind": {"path": "$company_from_job", "preserveNullAndEmptyArrays": True}},
            {"$lookup": {"from": "candidats", "localField": "candidateEmail", "foreignField": "email", "as": "candidate"}},
            {"$unwind": {"path": "$candidate", "preserveNullAndEmptyArrays": True}},
            {"$sort": {"date": -1}}
        ]
        results = list(mongo.db.candidatures.aggregate(pipeline))

        
        output = []
        for r in results:
            # Créer un nouveau dictionnaire sérialisable
            output_data = {}
            
            for key, value in r.items():
                if isinstance(value, ObjectId):
                    output_data[key] = str(value)
                elif isinstance(value, datetime):
                    output_data[key] = value.isoformat()
                elif isinstance(value, dict):
                    # Nettoyer les sous-documents
                    clean_dict = {}
                    for sub_key, sub_value in value.items():
                        if isinstance(sub_value, ObjectId):
                            clean_dict[sub_key] = str(sub_value)
                        elif isinstance(sub_value, datetime):
                            clean_dict[sub_key] = sub_value.isoformat()
                        else:
                            clean_dict[sub_key] = sub_value
                    output_data[key] = clean_dict
                elif isinstance(value, list):
                    output_data[key] = value
                else:
                    output_data[key] = value
            
            # Ajouter l'ID
            output_data['_id'] = str(r['_id'])
            output_data['candidateEmail'] = r.get('candidateEmail', '')
            
            # Ajouter les infos candidat
            candidate = r.get('candidate', {})
            output_data['first_name'] = candidate.get('first_name', '')
            output_data['last_name'] = candidate.get('last_name', '')
            output_data['photo_url'] = candidate.get('photo_url', '')
            output_data['cv_path'] = candidate.get('cv_path', '')
            output_data['email'] = candidate.get('email', '')
            
            # Formater les dates
            date_val = r.get('date') or r.get('created_at')
            if date_val:
                if isinstance(date_val, datetime):
                    output_data['date'] = date_val.isoformat()
                else:
                    output_data['date'] = str(date_val)
            
            created_val = r.get('created_at')
            if created_val:
                if isinstance(created_val, datetime):
                    output_data['created_at'] = created_val.isoformat()
                else:
                    output_data['created_at'] = str(created_val)
            
            # Ajouter le nom de l'entreprise
            company = r.get('company_from_job', {})
            output_data['company'] = company.get('name', '')
            
            output.append(output_data)
        
        return jsonify(serialize_doc(output)), 200
    except Exception as e:
        print(f"❌ Error: {e}")
        return jsonify([]), 200




@app.route('/api/admin/send-reminder', methods=['POST'])
def send_reminder():
    data = request.json
    send_real_email(data['email'], 'Reminder from LCA platform', f'<p>Hello {data.get("name", "")}, we remind you to follow up on your file.</p>')
    return jsonify({"message": "OK"})

@app.route('/api/admin/send-interview-link', methods=['POST'])
def send_interview_link():
    data = request.json
    send_real_email(data['email'], 'Link to your interview', f'<p>Link to the interview: <a href="{data["link"]}">{data["link"]}</a></p>')
    return jsonify({"message": "OK"})

@app.route('/api/candidatures/<id>/comments', methods=['GET'])
def get_candidature_comments(id):
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(id)})
    if not candidature:
        return jsonify([]), 200
    comments = candidature.get('internal_comments', [])
    for c in comments:
        if 'date' in c and isinstance(c['date'], datetime):
            c['date'] = c['date'].isoformat()
    return jsonify(comments), 200

@app.route('/api/jobs/add-independent', methods=['POST'])
def add_independent_job():
    data = request.json
    company_id = data.get('companyId')
    if not company_id:
        return jsonify({"error": "companyId required"}), 400
    company = mongo.db.entreprises.find_one({"_id": ObjectId(company_id)})
    if not company:
        return jsonify({"error": "Company not found"}), 404
    keywords = data.get('keywords', [])
    if isinstance(keywords, str):
        keywords = [k.strip() for k in keywords.split(',') if k.strip()]
    new_job = {
        "companyId": company_id,
        "companyName": company.get('name'),
        "title": data.get('title'),
        "description": data.get('description', ''),
        "location": data.get('location', company.get('location', 'Remote')),
        "type": data.get('type', 'Full-time'),
        "keywords": keywords,
        "status": "active",
        "created_at": datetime.utcnow()
    }
    result = mongo.db.postes.insert_one(new_job)
    new_job['_id'] = str(result.inserted_id)
    for cand in mongo.db.candidats.find({}, {"email": 1, "first_name": 1}):
        email = cand.get('email')
        if email:
            send_real_email(email, f"New job: {new_job['title']}", f"<html><body><h2>{new_job['title']} at {new_job['companyName']}</h2><p>{new_job['description'][:200]}...</p><a href='http://localhost:4200/login-candidate'>Apply</a></body></html>")
            mongo.db.notifications.insert_one({
                "recipient": "candidate",
                "recipientEmail": email,
                "type": "new_job",
                "title": f"New position: {new_job['title']}",
                "message": f"New job at {new_job['companyName']}",
                "status": "unread",
                "date": datetime.utcnow(),
                "jobId": str(new_job['_id'])
            })
    return jsonify({"_id": str(result.inserted_id), "message": "Job added successfully"}), 201

# ========== CLIENT SPACE ==========
@app.route('/api/client/access/<token>', methods=['GET'])
def client_access(token):
    candidature = mongo.db.candidatures.find_one({"client_token": token})
    if not candidature:
        return jsonify({"error": "Invalid or expired link"}), 404
    candidat = mongo.db.candidats.find_one({"email": candidature.get("candidateEmail")})
    if not candidat:
        return jsonify({"error": "Candidate not found"}), 404
    cv_info = {}
    if candidat.get("cv_path"):
        cv_path = os.path.join(app.config['UPLOAD_FOLDER'], candidat["cv_path"])
        if os.path.exists(cv_path):
            cv_info = extract_cv_info_spacy(cv_path)
    client_questions = candidature.get("client_questions", [])
    return jsonify({
        "candidate_name": candidature.get("candidateName"),
        "poste": candidature.get("poste"),
        "cv_skills": cv_info.get("skills", []),
        "cv_experience": cv_info.get("max_experience", 0),
        "cv_diplomas": cv_info.get("diplomas", []),
        "evaluation": candidature.get("evaluation", {}),
        "candidate_answers": candidature.get("answers", []),
        "client_questions": client_questions,
        "client_answers": candidature.get("client_answers", [])
    }), 200

@app.route('/api/client/submit-answers', methods=['POST'])
def client_submit_answers():
    data = request.json
    token = data.get("token")
    answers = data.get("answers", [])
    if not token:
        return jsonify({"error": "Missing token"}), 400
    candidature = mongo.db.candidatures.find_one({"client_token": token})
    if not candidature:
        return jsonify({"error": "Invalid or expired link"}), 404
    mongo.db.candidatures.update_one(
        {"_id": candidature["_id"]},
        {"$set": {"client_answers": answers, "client_answered_at": datetime.utcnow()}}
    )
    return jsonify({"message": "Answers saved"}), 200




@app.route('/api/admin/global-search', methods=['POST'])
def global_search():
    data = request.json
    query = data.get('query', '').strip()
    if len(query) < 2:
        return jsonify({"candidates": [], "applications": []})
    
    regex = {"$regex": query, "$options": "i"}
    
    # Search candidates by name, email, skills (from CV text)
    candidates = list(mongo.db.candidats.find({
        "$or": [
            {"first_name": regex}, {"last_name": regex}, {"email": regex},
            {"nationalite": regex}
        ]
    }, {"password": 0}).limit(20))
    
    # Search applications by job title, company, status
    applications = list(mongo.db.candidatures.find({
        "$or": [
            {"poste": regex}, {"company": regex}, {"status": regex},
            {"candidateName": regex}, {"candidateEmail": regex}
        ]
    }).sort("date", -1).limit(20))
    
    # Optionally, search inside CV text (requires storing extracted text)
    # You could add a 'cv_text' field in candidats collection.
    
    return jsonify({
        "candidates": [c for c in candidates],
        "applications": [a for a in applications]
    })




@app.route('/api/admin/interviews', methods=['GET'])
def get_all_interviews():
    pipeline = [
        {"$match": {"interview.date": {"$exists": True}}},
        {"$project": {
            "title": "$poste",
            "start": "$interview.date",
            "time": "$interview.time",
            "candidateName": 1,
            "candidateEmail": 1,
            "status": 1
        }}
    ]
    interviews = list(mongo.db.candidatures.aggregate(pipeline))
    return jsonify(interviews)





# When a new candidature is inserted, emit an event
def emit_new_application(candidature):
    socketio.emit('new_application', {
        'id': str(candidature['_id']),
        'candidate': candidature['candidateName'],
        'job': candidature['poste']
    })

    



@app.route('/api/admin/weekly-stats', methods=['GET'])
def weekly_stats():
    now = datetime.utcnow()
    start_of_week = now - timedelta(days=now.weekday())
    new_candidates = mongo.db.candidats.count_documents({"created_at": {"$gte": start_of_week}})
    new_jobs = mongo.db.postes.count_documents({"created_at": {"$gte": start_of_week}})
    return jsonify({"newCandidates": new_candidates, "newJobs": new_jobs})



















































































# ========== CANDIDATE DASHBOARD ROUTES ==========
@app.route('/register', methods=['POST'])
def register_candidate():
    first_name = request.form.get('first_name') or request.form.get('nom', '')
    last_name = request.form.get('last_name') or request.form.get('prenom', '')
    nationalite = request.form.get('nationalite', '')
    email = request.form.get('email')
    password = request.form.get('password')
    if mongo.db.candidats.find_one({"email": email}):
        return jsonify({"error": "Email already used"}), 400
    if 'cv' not in request.files or not request.files['cv'].filename:
        return jsonify({"error": "CV required"}), 400
    cv = request.files['cv']
    cv_filename = secure_filename(f"{email}_cv_{datetime.utcnow().timestamp()}.pdf")
    cv.save(os.path.join('uploads/cvs', cv_filename))
    photo_filename = ""
    if 'photo' in request.files and request.files['photo'].filename:
        photo = request.files['photo']
        ext = photo.filename.rsplit('.', 1)[1].lower()
        photo_filename = secure_filename(f"{email}_photo_{datetime.utcnow().timestamp()}.{ext}")
        photo.save(os.path.join('uploads/photos', photo_filename))
    mongo.db.candidats.insert_one({
        "first_name": first_name, "last_name": last_name, "nationalite": nationalite,
        "email": email, "password": generate_password_hash(password), "role": "candidate",
        "cv_path": cv_filename, "photo_url": photo_filename, "created_at": datetime.utcnow()
    })
    return jsonify({"message": "Registration successful"}), 201

@app.route('/login-candidate', methods=['POST'])
def login_candidate():
    data = request.json
    cand = mongo.db.candidats.find_one({"email": data['email']})
    if not cand or not check_password_hash(cand['password'], data['password']):
        return jsonify({"error": "Invalid email or password"}), 401
    return jsonify({
        "token": str(uuid.uuid4()),
        "role": "candidate",
        "email": cand['email'],
        "first_name": cand.get('first_name', ''),
        "last_name": cand.get('last_name', ''),
        "photo_url": cand.get('photo_url', 'assets/default-avatar.png'),
        "cv_path": cand.get('cv_path', '')
    }), 200

@app.route('/api/candidate/profile/<email>', methods=['GET'])
def get_candidate_profile(email):
    cand = mongo.db.candidats.find_one({"email": email})
    if cand:
        photo_url = cand.get('photo_url', '')
        # NE PAS ajouter /static/ ici - retourner juste le nom du fichier
        # si photo_url contient déjà un chemin, extraire juste le nom
        if photo_url and '/' in photo_url:
            photo_url = photo_url.split('/')[-1]
            
        return jsonify({
            "email": cand.get('email'),
            "first_name": cand.get('first_name', ''),
            "last_name": cand.get('last_name', ''),
            "photo_url": photo_url,  # retourne juste le nom du fichier
            "cv_path": cand.get('cv_path', ''),
            "nationalite": cand.get('nationalite', '')
        }), 200
    return jsonify({"error": "Not found"}), 404


@app.route('/upload_cv', methods=['POST'])
def upload_cv():
    file = request.files.get('file')
    email = request.form.get('email')
    if not file or not file.filename.endswith('.pdf'):
        return jsonify({"error": "PDF required"}), 400
    filename = secure_filename(f"{email.replace('@', '_').replace('.', '_')}_{datetime.utcnow().timestamp()}.pdf")
    file.save(os.path.join('uploads/cvs', filename))
    mongo.db.candidats.update_one({"email": email}, {"$set": {"cv_path": filename}})
    return jsonify({"message": "CV uploaded", "file_name": filename}), 200

@app.route('/upload_candidate_photo', methods=['POST'])
def upload_candidate_photo():
    file = request.files.get('photo')
    email = request.form.get('email')
    if not file:
        return jsonify({"error": "No file"}), 400
    ext = file.filename.rsplit('.', 1)[1].lower()
    filename = f"{email.replace('@', '_').replace('.', '_')}_{datetime.utcnow().timestamp()}.{ext}"
    file.save(os.path.join('uploads/photos', filename))
    mongo.db.candidats.update_one({"email": email}, {"$set": {"photo_url": filename}})
    return jsonify({"photo_url": filename}), 200

@app.route('/update_profile', methods=['POST', 'PUT', 'PATCH'])
def update_profile():
    data = request.json
    email = data.get('email')
    if not email:
        return jsonify({"error": "Email required"}), 400
    update = {}
    if 'first_name' in data and 'last_name' in data:
        update['first_name'] = data['first_name']
        update['last_name'] = data['last_name']
    else:
        full = data.get('full_name') or data.get('userName')
        if full:
            parts = full.strip().split(' ')
            update['first_name'] = parts[0]
            update['last_name'] = ' '.join(parts[1:]) if len(parts) > 1 else ''
    if 'nationalite' in data:
        update['nationalite'] = data['nationalite']
    if update:
        mongo.db.candidats.update_one({"email": email}, {"$set": update})
        return jsonify({"message": "Profile updated", **update}), 200
    return jsonify({"message": "No changes"}), 200

@app.route('/api/candidate/change-password', methods=['POST'])
def change_candidate_password():
    data = request.json
    email = data.get('email')
    current_password = data.get('currentPassword')
    new_password = data.get('newPassword')
    if not email or not current_password or not new_password:
        return jsonify({"error": "Missing fields"}), 400
    candidate = mongo.db.candidats.find_one({"email": email})
    if not candidate or not check_password_hash(candidate['password'], current_password):
        return jsonify({"error": "Current password is incorrect"}), 401
    hashed_new = generate_password_hash(new_password)
    mongo.db.candidats.update_one({"email": email}, {"$set": {"password": hashed_new}})
    return jsonify({"message": "Password changed successfully"}), 200

@app.route('/get_entreprises', methods=['GET'])
def get_entreprises():
    ents = list(mongo.db.entreprises.find())
    for e in ents:
        e['_id'] = str(e['_id'])
    return jsonify(ents), 200

@app.route('/apply_job', methods=['POST'])
def apply_job():
    try:
        file = request.files.get('cv')
        poste_id = request.form.get('posteId')
        candidate_email = request.form.get('candidateEmail')
        candidate_name = request.form.get('candidateName', 'Candidat')
        
        # Si l'email est "candidat@lca.com" ou générique, c'est un utilisateur non connecté
        is_anonymous = candidate_email == 'candidat@lca.com' or candidate_email == 'test@candidat.com'
        
        print(f"📥 Reçu: posteId={poste_id}, email={candidate_email}, is_anonymous={is_anonymous}")
        
        if not file or not poste_id:
            return jsonify({"error": "CV and posteId required"}), 400
        
        # Sauvegarder le CV
        timestamp = datetime.utcnow().timestamp()
        safe_email = candidate_email.replace('@', '_').replace('.', '_')
        cv_filename = secure_filename(f"{safe_email}_cv_{timestamp}.pdf")
        file.save(os.path.join('uploads/cvs', cv_filename))
        
        # Récupérer le job
        if not ObjectId.is_valid(poste_id):
            return jsonify({"error": "Invalid job ID"}), 400
        
        job = mongo.db.postes.find_one({"_id": ObjectId(poste_id)})
        if not job:
            return jsonify({"error": "Job not found"}), 404
        
        # Calculer le score
        keywords = job.get('keywords', [])
        score = calculate_advanced_score(cv_filename, keywords)
        
        # ⭐⭐ EXTRAIRE LE NOM DU CV si c'est un candidat anonyme ⭐⭐
        first_name = ''
        last_name = ''
        
        if is_anonymous or candidate_name == 'Candidat' or candidate_name == 'Test Candidat':
            cv_full_path = os.path.join('uploads/cvs', cv_filename)
            if os.path.exists(cv_full_path):
                extracted_first, extracted_last = extract_candidate_name_from_cv(cv_full_path)
                if extracted_first:
                    first_name = extracted_first
                    last_name = extracted_last
                    candidate_name = f"{first_name} {last_name}".strip()
                    print(f"📝 Nom extrait du CV: {first_name} {last_name}")
        
        # Si toujours pas de nom, utiliser "Utilisateur" + ID
        if not first_name and not last_name:
            first_name = "Utilisateur"
            last_name = str(uuid.uuid4())[:6]
            candidate_name = f"{first_name} {last_name}"
        
        # Vérifier si le candidat existe déjà dans la base
        existing_candidate = mongo.db.candidats.find_one({"email": candidate_email})
        
        if existing_candidate:
            # Mettre à jour le nom si vide
            if not existing_candidate.get('first_name') and first_name:
                mongo.db.candidats.update_one(
                    {"_id": existing_candidate["_id"]},
                    {"$set": {"first_name": first_name, "last_name": last_name}}
                )
            candidate_id = existing_candidate["_id"]
        else:
            # Créer un nouveau candidat
            candidate = {
                "first_name": first_name,
                "last_name": last_name,
                "email": candidate_email,
                "password": generate_password_hash("temporary_" + str(uuid.uuid4())[:8]),
                "role": "candidate",
                "cv_path": cv_filename,
                "photo_url": "",
                "status": "pending",
                "created_at": datetime.utcnow()
            }
            result = mongo.db.candidats.insert_one(candidate)
            candidate_id = result.inserted_id
        
        # Créer la candidature
        candidature = {
            "poste": job.get('title'),
            "company": job.get('companyName', 'LCA'),
            "companyId": job.get('companyId'),
            "posteId": poste_id,
            "candidateEmail": candidate_email,
            "candidateName": candidate_name,
            "candidateId": str(candidate_id),
            "cvPath": cv_filename,
            "status": "pending",
            "score": score,
            "date": datetime.utcnow()
        }
        
        mongo.db.candidatures.insert_one(candidature)
        
        # Notification admin
        mongo.db.notifications.insert_one({
            "recipient": "admin",
            "recipientEmail": "admin@lca.com",
            "type": "new_candidature",
            "title": "📥 New application",
            "message": f"{candidate_name} applied for {job.get('title')}",
            "status": "unread",
            "date": datetime.utcnow()
        })
        
        return jsonify({
            "message": "Application sent",
            "score": score,
            "skills": keywords[:5],
            "candidate_name": candidate_name
        }), 200
        
    except Exception as e:
        print(f"❌ Error in apply_job: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


def extract_candidate_name_from_cv(cv_path):
    """Extrait le nom du candidat depuis son CV si possible"""
    try:
        text = ""
        with pdfplumber.open(cv_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
        
        # Chercher des patterns de nom (premières lignes)
        lines = text.split('\n')[:10]  # Premières 10 lignes
        
        words = []
        common_words = ['curriculum', 'vitae', 'resume', 'cv', 'name', 'email', 
                        'phone', 'address', 'tél', 'téléphone', 'adresse']
        
        for line in lines:
            line = line.strip()
            if len(line) > 3 and len(line) < 50:
                # Éviter les lignes qui ressemblent à des emails
                if '@' not in line and not any(cw in line.lower() for cw in common_words):
                    # Vérifier si ça ressemble à un nom (2-3 mots)
                    parts = line.split()
                    if 1 <= len(parts) <= 3:
                        words.extend(parts)
        
        # Si on a trouvé des mots, prendre les 2 premiers comme prénom/nom
        if len(words) >= 2:
            return words[0], ' '.join(words[1:3])
        elif len(words) == 1:
            return words[0], ''
        else:
            return '', ''
    except:
        return '', ''


@app.route('/get_candidatures/<email>', methods=['GET'])
def get_candidatures(email):
    candidatures = list(mongo.db.candidatures.find({
        "candidateEmail": email,
        "hidden": {"$ne": True}
    }))
    for c in candidatures:
        c['_id'] = str(c['_id'])
        c['questionnaire_token'] = c.get('questionnaire_token')
        c['questionnaire_status'] = c.get('questionnaire_status', 'pending')
        print(f"DEBUG: {c.get('candidateName')} -> token={c.get('questionnaire_token')}, status={c.get('questionnaire_status')}")
    updates = []
    for c in candidatures:
        stored_poste = c.get('poste', '')
        if not stored_poste or stored_poste.lower() in ['pending', 'to_contact', 'validated', 'archived', 'contacted', 'interview_scheduled']:
            job_id = c.get('posteId') or c.get('jobId')
            if job_id and ObjectId.is_valid(job_id):
                job = mongo.db.postes.find_one({"_id": ObjectId(job_id)})
                if job and job.get('title'):
                    nouveau_poste = job['title']
                    c['poste'] = nouveau_poste
                    updates.append((ObjectId(c['_id']), nouveau_poste))
            else:
                c['poste'] = "Deleted offer"
    for candidature_id, nouveau_titre in updates:
        mongo.db.candidatures.update_one({"_id": candidature_id}, {"$set": {"poste": nouveau_titre}})
    return jsonify(candidatures), 200

@app.route('/api/candidate/notifications/<email>', methods=['GET'])
def get_candidate_notifications(email):
    notifs = list(mongo.db.notifications.find({"recipientEmail": email}).sort("date", -1))
    for n in notifs:
        n['_id'] = str(n['_id'])
        if 'date' in n and hasattr(n['date'], 'isoformat'):
            n['date'] = n['date'].isoformat()
    unread = sum(1 for n in notifs if n.get('status') == 'unread')
    return jsonify({"notifications": notifs, "unread_count": unread}), 200

@app.route('/api/candidate/notifications/mark-read/<id>', methods=['PUT'])
def mark_notification_read(id):
    mongo.db.notifications.update_one({"_id": ObjectId(id)}, {"$set": {"status": "read"}})
    return jsonify({"message": "Marked read"}), 200

@app.route('/api/candidate/notifications/mark-all-read', methods=['PUT'])
def mark_all_read():
    email = request.json.get('email')
    mongo.db.notifications.update_many({"recipientEmail": email, "status": "unread"}, {"$set": {"status": "read"}})
    return jsonify({"message": "All read"}), 200

@app.route('/api/candidate/conversation/<email>', methods=['GET'])
def get_candidate_conversation(email):
    msgs = list(mongo.db.messages.find({
        "$or": [{"fromEmail": email, "to": "admin"}, {"from": "admin", "toEmail": email}]
    }).sort("date", 1))
    for m in msgs:
        m['_id'] = str(m['_id'])
        if 'date' in m and hasattr(m['date'], 'isoformat'):
            m['date'] = m['date'].isoformat()
        m['sender'] = 'candidate' if m.get('from') == 'candidate' else 'admin'
        m['senderName'] = 'Administrator' if m['sender'] == 'admin' else m.get('fromName', 'Candidate')
    return jsonify(msgs), 200

@app.route('/api/candidate/send-message', methods=['POST'])
def candidate_send_message():
    data = request.json
    msg = {
        "from": "candidate",
        "fromEmail": data.get('fromEmail'),
        "fromName": data.get('fromName'),
        "to": "admin",
        "toEmail": "admin@lca.com",
        "message": data.get('message'),
        "date": datetime.utcnow(),
        "read": False
    }
    mongo.db.messages.insert_one(msg)
    mongo.db.notifications.insert_one({
        "type": "new_message",
        "recipient": "admin",
        "recipientEmail": "admin@lca.com",
        "title": "New message from a candidate",
        "message": f"{data.get('fromName')} sent you a message.",
        "status": "unread",
        "date": datetime.utcnow()
    })
    return jsonify({"message": "Message sent"}), 201

# Saved jobs
@app.route('/api/candidate/saved-jobs/<email>', methods=['GET'])
def get_saved_jobs(email):
    saved = list(mongo.db.saved_jobs.find({"candidateEmail": email}))
    job_ids = []
    for s in saved:
        job_id = s.get('jobId')
        if job_id and ObjectId.is_valid(job_id):
            job_ids.append(ObjectId(job_id))
    jobs = list(mongo.db.postes.find({"_id": {"$in": job_ids}}))
    for job in jobs:
        job['_id'] = str(job['_id'])
        company = mongo.db.entreprises.find_one({"_id": ObjectId(job.get('companyId'))})
        job['company'] = company.get('name') if company else 'Unknown'
    return jsonify(jobs), 200

@app.route('/api/candidate/unsave-job/<email>/<job_id>', methods=['DELETE'])
def unsave_job(email, job_id):
    if not ObjectId.is_valid(job_id):
        return jsonify({"error": "Invalid job ID"}), 400
    result = mongo.db.saved_jobs.delete_one({"candidateEmail": email, "jobId": job_id})
    if result.deleted_count:
        return jsonify({"message": "Removed"}), 200
    else:
        return jsonify({"error": "Job not found in favorites"}), 404

@app.route('/api/candidate/save-job', methods=['POST'])
def save_job():
    data = request.json
    email = data.get('candidateEmail')
    job_id = data.get('jobId')
    if not email or not job_id:
        return jsonify({"error": "Missing data"}), 400
    if not ObjectId.is_valid(job_id):
        return jsonify({"error": "Invalid job ID"}), 400
    if mongo.db.saved_jobs.find_one({"candidateEmail": email, "jobId": job_id}):
        return jsonify({"message": "Job already saved"}), 200
    mongo.db.saved_jobs.insert_one({
        "candidateEmail": email,
        "jobId": job_id,
        "savedAt": datetime.utcnow()
    })
    return jsonify({"message": "Saved"}), 201

@app.route('/api/salary-by-nationality', methods=['GET'])
def get_salary_scale():
    return jsonify({
        "Tunisian": {"amount": 2500, "currency": "DT"},
        "Moroccan": {"amount": 8000, "currency": "MAD"},
        "French": {"amount": 2000, "currency": "EUR"},
        "Algerian": {"amount": 70000, "currency": "DZD"},
        "Other": {"amount": 1500, "currency": "EUR"}
    }), 200

@app.route('/apply_job_with_details', methods=['POST'])
def apply_job_with_details():
    try:
        data = request.json
        candidate_email = data.get('candidateEmail')
        if not candidate_email:
            return jsonify({"error": "candidateEmail required"}), 400
        cv_filename = data.get('cvPath')
        if not cv_filename:
            return jsonify({"error": "CV not found"}), 400
        job_id = data.get('jobId') or data.get('posteId')
        if not job_id:
            return jsonify({"error": "jobId required"}), 400
        job = mongo.db.postes.find_one({"_id": ObjectId(job_id)})
        if not job:
            return jsonify({"error": "Job not found"}), 404
        keywords = job.get('keywords', [])
        score = calculate_advanced_score(cv_filename, keywords)
        langues = data.get('langues', [])
        if not langues and data.get('langue') and data.get('niveauLangue'):
            langues = [{'nom': data.get('langue'), 'niveau': data.get('niveauLangue')}]
        candidature_dict = {
            "poste": job.get('title'),
            "company": data.get('company'),
            "companyId": job.get('companyId'),
            "posteId": job_id,
            "candidateEmail": candidate_email,
            "candidateName": data.get('candidateName'),
            "cvPath": cv_filename,
            "status": "pending",
            "score": score,
            "date": datetime.utcnow(),
            "langues": langues,
            "nationalite": data.get('nationalite'),
            "salaire": data.get('salaire'),
            "devise": data.get('devise'),
            "residence": data.get('residence')
        }
        result = mongo.db.candidatures.insert_one(candidature_dict)
        candidature_id = str(result.inserted_id)
        candidate_name = data.get('candidateName', '').strip()
        if candidate_name:
            mongo.db.notifications.insert_one({
                "recipient": "admin",
                "recipientEmail": "admin@lca.com",
                "type": "new_candidature",
                "title": "📥 New application",
                "message": f"{data.get('candidateName')} applied for {job.get('title')}",
                "status": "unread",
                "date": datetime.utcnow(),
                "candidatureId": candidature_id
            })
        return jsonify({"message": "Application sent"}), 201
    except Exception as e:
        print(f"❌ Error in apply_job_with_details: {e}")
        return jsonify({"error": str(e)}), 500

# ========== DAILY ADVICE ==========
@app.route('/api/daily-advice', methods=['GET'])
def get_daily_advice():
    init_daily_tips()
    today = datetime.utcnow().strftime("%Y-%m-%d")
    tip = mongo.db.daily_tips.find_one({"used_dates": {"$ne": today}})
    if not tip:
        mongo.db.daily_tips.update_many({}, {"$set": {"used_dates": []}})
        tip = mongo.db.daily_tips.find_one()
    if tip:
        mongo.db.daily_tips.update_one({"_id": tip["_id"]}, {"$push": {"used_dates": today}})
        return jsonify({"advice": tip["text"]}), 200
    return jsonify({"advice": "Update your CV today!"}), 200

# ========== PERSONALIZED INTERVIEW TIP ==========
@app.route('/api/interview-tip/<email>', methods=['GET'])
def personalized_interview_tip(email):
    past_interviews = mongo.db.candidatures.count_documents({
        "candidateEmail": email,
        "status": {"$in": ["interview_scheduled", "contacted"]}
    })
    if past_interviews >= 2:
        tip = "💪 You already have interview experience. For the next one, prepare a case study on your key project."
    elif past_interviews == 1:
        tip = "🎯 You've had one interview. Analyze what went well and prepare more pointed questions about the company."
    else:
        tip = "📝 First interview? Practice in front of a mirror, time your answers (max 2 minutes per question)."
    return jsonify({"tip": tip}), 200

# ========== WEEKLY ACTION ==========
@app.route('/api/user/complete-action/<email>', methods=['POST'])
def complete_weekly_action(email):
    current_week = datetime.utcnow().strftime("%Y-W%W")
    result = mongo.db.user_actions.update_one({"email": email, "week": current_week}, {"$set": {"completed": True}})
    if result.modified_count:
        mongo.db.candidats.update_one({"email": email}, {"$inc": {"gamification_points": 10}})
    return jsonify({"message": "Action completed!"}), 200

# ========== STATIC FILES ==========
@app.route('/static/uploads/admins/<filename>')
def get_admin_static(filename):
    return send_from_directory('static/uploads/admins', filename)

@app.route('/uploads/cvs/<filename>')
def serve_cv(filename):
    return send_from_directory('uploads/cvs', filename)


@app.route('/uploads/photos/<filename>')
def serve_photo(filename):
    return send_from_directory('uploads/photos', filename)

@app.route('/static/uploads/photos/<filename>')
def serve_static_photo(filename):
    return send_from_directory('uploads/photos', filename)


# ========== QUESTIONNAIRE ROUTES ==========
@app.route('/api/admin/candidatures/<id>/questions', methods=['POST'])
def set_candidate_questions(id):
    data = request.json
    questions = data.get('questions', [])
    
    print(f"📥 Received {len(questions)} questions for application {id}")
    
    if not questions:
        return jsonify({"error": "No questions provided"}), 400
    
    # Format correct des questions
    questions_formatted = []
    for q in questions:
        if isinstance(q, str):
            questions_formatted.append({"text": q, "type": "text"})
        else:
            q_type = q.get("type", "text")
            questions_formatted.append({
                "text": q.get("text", ""), 
                "type": q_type
            })
    
    print(f"📝 Formatted: {len(questions_formatted)} questions")

    
    # Générer un token pour le questionnaire
    import hashlib
    token = hashlib.md5(f"{id}{datetime.utcnow()}".encode()).hexdigest()[:16]
    
    # Mettre à jour avec les bons champs
    result = mongo.db.candidatures.update_one(
        {"_id": ObjectId(id)},
        {"$set": {
            "custom_questions": questions_formatted, 
            "questionnaire_status": "pending", 
            "answers": [],
            "questionnaire_token": token
        }}
    )
    
    if result.modified_count == 0:
        return jsonify({"error": "Candidature not found or not modified"}), 404
    
    print(f"✅ Saved to MongoDB: {len(questions_formatted)} questions, token={token}")
    
    # ========== ENVOYER L'EMAIL AU CANDIDAT ==========
    # Récupérer la candidature pour avoir l'email
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(id)})
    candidate_email = candidature.get('candidateEmail')
    candidate_name = candidature.get('candidateName', 'Candidat')
    
    # Construire le lien
    link = f"http://localhost:4200/questionnaire/{token}"
    
    subject = "📋 Formulaire d'évaluation détaillé"
    body = f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <h2 style="color: #00a6a6;">Formulaire d'évaluation détaillé</h2>
        <p>Bonjour {candidate_name},</p>
        <p>Merci de bien vouloir remplir ce formulaire d'évaluation détaillé pour finaliser votre candidature :</p>
        <p style="margin: 20px 0;">
            <a href="{link}" style="background-color: #00a6a6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">
                ➡️ Accéder au formulaire d'évaluation
            </a>
        </p>
        <p>Cordialement,<br>L'équipe LCA</p>
    </body>
    </html>
    """
    
    # Envoyer l'email
    send_real_email(candidate_email, subject, body)
    
    # Ajouter une notification pour le candidat
    mongo.db.notifications.insert_one({
        "recipient": "candidate",
        "recipientEmail": candidate_email,
        "type": "questionnaire_received",
        "title": "📋 Formulaire d'évaluation",
        "message": f"Un formulaire d'évaluation détaillé est disponible pour votre candidature.",
        "status": "unread",
        "date": datetime.utcnow()
    })
    
    print(f"📧 Email sent to {candidate_email} with token: {token}")

    
    return jsonify({
        "message": "Formulaire envoyé au candidat", 
        "count": len(questions_formatted),
        "token": token
    }), 200



@app.route('/api/debug/question-count/<candidature_id>', methods=['GET'])
def debug_question_count(candidature_id):
    if not ObjectId.is_valid(candidature_id):
        return jsonify({"error": "Invalid ID"}), 400
    
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(candidature_id)})
    if not candidature:
        return jsonify({"error": "Not found"}), 404
    
    questions = candidature.get("custom_questions", [])
    
    return jsonify({
        "candidature_id": candidature_id,
        "token": candidature.get("questionnaire_token"),
        "questions_count": len(questions),
        "questions": questions[:5],  # Afficher les 5 premières
        "full_questions": questions   # Toutes les questions
    }), 200

# À ajouter dans app.py, après les autres routes

@app.route('/api/candidatures/<id>/generate-client-token', methods=['POST'])
def generate_client_token(id):
    """Génère un token pour le client"""
    if not ObjectId.is_valid(id):
        return jsonify({"error": "Invalid ID"}), 400
    
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(id)})
    if not candidature:
        return jsonify({"error": "Not found"}), 404
    
    token = hashlib.md5(f"{candidature['candidateEmail']}{id}{datetime.utcnow()}".encode()).hexdigest()[:16]
    
    mongo.db.candidatures.update_one(
        {"_id": ObjectId(id)},
        {"$set": {"client_token": token}}
    )
    
    return jsonify({"token": token}), 200


@app.route('/api/candidatures/delete/<candidature_id>', methods=['DELETE', 'OPTIONS'])
def delete_candidature(candidature_id):
    """Supprime définitivement une candidature"""
    
    if request.method == 'OPTIONS':
        response = jsonify({'message': 'OK'})
        response.headers.add('Access-Control-Allow-Origin', '*')
        response.headers.add('Access-Control-Allow-Headers', 'Content-Type, Authorization')
        response.headers.add('Access-Control-Allow-Methods', 'DELETE, OPTIONS')
        return response
    
    if not ObjectId.is_valid(candidature_id):
        return jsonify({"error": "Invalid ID"}), 400
    
    try:
        result = mongo.db.candidatures.delete_one({"_id": ObjectId(candidature_id)})
        
        if result.deleted_count:
            return jsonify({"message": "Application deleted successfully"}), 200
        else:
            return jsonify({"error": "Application not found"}), 404
            
    except Exception as e:
        print(f"❌ Error deleting candidature: {e}")
        return jsonify({"error": str(e)}), 500




@app.route('/api/candidate/questionnaire/<token>', methods=['GET'])
def get_candidate_questionnaire_merged(token):
    """Récupère le questionnaire fusionné (évaluation + personnalisé)"""
    try:
        print(f"🔍 Recherche du token: {token}")
        
        # Chercher dans candidatures d'abord
        candidature = mongo.db.candidatures.find_one({"questionnaire_token": token})
        
        if not candidature:
            return jsonify({"error": "Lien invalide ou expiré"}), 404
        
        # ========== 1. RÉCUPÉRER LES QUESTIONS DE LA FICHE D'ÉVALUATION ==========
        evaluation_questions = []
        
        # Essayer de récupérer les questions d'évaluation depuis candidature
        if candidature.get('evaluation_questions') and len(candidature.get('evaluation_questions', [])) > 0:
            evaluation_questions = candidature.get('evaluation_questions')
        else:
            # Questions par défaut de la fiche d'évaluation (16 questions)
            evaluation_questions = [
                {"text": "Nom complet", "type": "text", "required": True},
                {"text": "Genre (Homme/Femme)", "type": "choice", "required": True, "options": ["Homme", "Femme"]},
                {"text": "Téléphone", "type": "tel", "required": True},
                {"text": "Email", "type": "email", "required": True},
                {"text": "Nationalité", "type": "text", "required": True},
                {"text": "Années d'expérience / Secteur", "type": "text", "required": True},
                {"text": "Type de contrat actuel", "type": "choice", "required": True, 
                 "options": ["CDI", "CDD", "Stage", "Freelance", "Indépendant", "Sans emploi"]},
                {"text": "Délai de préavis", "type": "text", "required": False},
                {"text": "Salaire actuel", "type": "text", "required": False},
                {"text": "Salaire demandé", "type": "text", "required": True},
                {"text": "Entreprise actuelle", "type": "text", "required": False},
                {"text": "Adresse actuelle", "type": "text", "required": True},
                {"text": "Diplôme", "type": "text", "required": True},
                {"text": "Raison de changement de poste", "type": "textarea", "required": True},
                {"text": "Certifications", "type": "textarea", "required": False},
                {"text": "Parcours professionnel", "type": "textarea", "required": False}
            ]
        
        # ========== 2. RÉCUPÉRER LES QUESTIONS DU QUESTIONNAIRE PERSONNALISÉ ==========
        custom_questions = candidature.get('custom_questions', [])
        
        # ========== 3. FUSIONNER LES DEUX ==========
        all_questions = []
        
        # Ajouter un en-tête pour la section Évaluation
        if evaluation_questions and len(evaluation_questions) > 0:
            all_questions.append({
                "text": "📋 SECTION 1: FICHE D'ÉVALUATION DÉTAILLÉE",
                "type": "section_header",
                "required": False
            })
            for q in evaluation_questions:
                if isinstance(q, dict):
                    all_questions.append({
                        "text": q.get('text', 'Question'),
                        "type": q.get('type', 'text'),
                        "required": q.get('required', True),
                        "options": q.get('options', [])
                    })
                else:
                    all_questions.append({"text": str(q), "type": "text", "required": True})
        
        # Ajouter un en-tête pour la section Questionnaire personnalisé
        if custom_questions and len(custom_questions) > 0:
            all_questions.append({
                "text": "📝 SECTION 2: QUESTIONNAIRE PERSONNALISÉ",
                "type": "section_header",
                "required": False
            })
            for q in custom_questions:
                if isinstance(q, dict):
                    # Extraire le texte de la question
                    question_text = q.get('text', '')
                    if not question_text:
                        question_text = q.get('textFr', '')
                    if not question_text:
                        question_text = q.get('question', '')
                    
                    all_questions.append({
                        "text": question_text,
                        "type": q.get('type', 'text'),
                        "required": q.get('required', True),
                        "options": q.get('options', [])
                    })
                elif isinstance(q, str):
                    all_questions.append({"text": q, "type": "text", "required": True})
        
        print(f"✅ {len(evaluation_questions)} questions d'évaluation + {len(custom_questions)} questions personnalisées = {len(all_questions)} questions totales")
        
        # Mettre à jour la candidature avec toutes les questions
        mongo.db.candidatures.update_one(
            {"_id": candidature["_id"]},
            {"$set": {"merged_questions": all_questions}}
        )
        
        return jsonify({
            "status": candidature.get('questionnaire_status', 'pending'),
            "candidatureId": str(candidature["_id"]),
            "questions": all_questions,
            "title": f"Questionnaire - {candidature.get('poste', 'Poste')}",
            "candidateName": candidature.get('candidateName', 'Candidat'),
            "evaluation_count": len(evaluation_questions),
            "custom_count": len(custom_questions)
        }), 200
        
    except Exception as e:
        print(f"❌ Erreur: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route('/api/debug/answers/<candidature_id>', methods=['GET'])
def debug_answers_public(candidature_id):
    """Route publique pour debug les réponses d'une candidature"""
    if not ObjectId.is_valid(candidature_id):
        return jsonify({"error": "Invalid ID"}), 400
    
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(candidature_id)})
    if not candidature:
        return jsonify({"error": "Not found"}), 404
    
    answers = candidature.get("answers", [])
    questions = candidature.get("custom_questions", [])
    
    return jsonify({
        "candidature_id": str(candidature["_id"]),
        "candidateName": candidature.get("candidateName"),
        "questions_count": len(questions),
        "answers_count": len(answers),
        "status": candidature.get("questionnaire_status"),
        "answers": answers
    }), 200



@app.route('/api/admin/force-update-questions/<candidature_id>', methods=['POST'])
def force_update_questions(candidature_id):
    all_questions = [
        {"text": "Nom du candidat", "type": "text"},
        {"text": "Genre (Homme/Femme)", "type": "choice"},
        {"text": "Téléphone", "type": "text"},
        {"text": "Email", "type": "text"},
        {"text": "Années d'expérience / Secteur", "type": "text"},
        {"text": "Type de contrat actuel", "type": "choice"},
        {"text": "Délai de préavis", "type": "text"},
        {"text": "Salaire actuel", "type": "text"},
        {"text": "Salaire demandé", "type": "text"},
        {"text": "Entreprise actuelle", "type": "text"},
        {"text": "Adresse actuelle", "type": "text"},
        {"text": "Diplôme", "type": "text"},
        {"text": "Raison de changement de poste", "type": "textarea"},
        {"text": "Langue pendant l'entretien", "type": "choice"},
        {"text": "Certifications", "type": "textarea"},
        {"text": "Parcours professionnel", "type": "textarea"}
       
    ]
    
    if not ObjectId.is_valid(candidature_id):
        return jsonify({"error": "Invalid candidature ID"}), 400
    
    result = mongo.db.candidatures.update_one(
        {"_id": ObjectId(candidature_id)},
        {"$set": {"custom_questions": all_questions, "questionnaire_status": "pending"}}
    )
    
    if result.modified_count:
        return jsonify({
            "message": f"✅ {len(all_questions)} questions mises à jour pour la candidature {candidature_id}",
            "count": len(all_questions)
        }), 200
    else:
        return jsonify({"error": "Candidature non trouvée ou non modifiée"}), 404





@app.route('/api/debug/candidature-questions/<candidature_id>', methods=['GET'])
def debug_candidature_questions(candidature_id):
    """Route de débogage pour voir les questions d'une candidature"""
    if not ObjectId.is_valid(candidature_id):
        return jsonify({"error": "Invalid ID"}), 400
    
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(candidature_id)})
    if not candidature:
        return jsonify({"error": "Not found"}), 404
    
    return jsonify({
        "candidature_id": str(candidature["_id"]),
        "candidate_email": candidature.get("candidateEmail"),
        "custom_questions": candidature.get("custom_questions", []),
        "questionnaire_token": candidature.get("questionnaire_token"),
        "questionnaire_status": candidature.get("questionnaire_status"),
        "has_questions": len(candidature.get("custom_questions", [])) > 0
    }), 200

@app.route('/api/candidate/questionnaire/submit', methods=['POST'])
def submit_questionnaire_enhanced():
    """Soumet les réponses du questionnaire"""
    try:
        data = request.json
        candidature_id = data.get('candidatureId')
        answers = data.get('answers', [])
        
        print(f"📥 Soumission questionnaire - ID: {candidature_id}, Réponses: {len(answers)}")
        
        if not candidature_id or not ObjectId.is_valid(candidature_id):
            return jsonify({"error": "ID de candidature invalide"}), 400
        
        candidature = mongo.db.candidatures.find_one({"_id": ObjectId(candidature_id)})
        if not candidature:
            return jsonify({"error": "Candidature non trouvée"}), 404
        
        # Formater les réponses
        formatted_answers = []
        for a in answers:
            if isinstance(a, dict):
                formatted_answers.append({
                    "question": a.get('question', a.get('questionText', '')),
                    "answer": a.get('answer', '')
                })
        
        print(f"📝 {len(formatted_answers)} réponses formatées")
        
        # Sauvegarder
        mongo.db.candidatures.update_one(
            {"_id": ObjectId(candidature_id)},
            {"$set": {
                "answers": formatted_answers,
                "questionnaire_status": "answered",
                "questionnaire_answered_at": datetime.utcnow()
            }}
        )
        
        # Notification pour l'admin
        mongo.db.notifications.insert_one({
            "recipient": "admin",
            "recipientEmail": "admin@lca.com",
            "type": "questionnaire_answered",
            "title": "📝 Questionnaire rempli",
            "message": f"{candidature.get('candidateName')} a répondu au questionnaire",
            "status": "unread",
            "date": datetime.utcnow(),
            "candidatureId": candidature_id
        })
        
        return jsonify({
            "message": "Réponses enregistrées avec succès",
            "answers_count": len(formatted_answers)
        }), 200
        
    except Exception as e:
        print(f"❌ Erreur: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500



@app.route('/api/admin/candidatures/<id>/answers', methods=['GET'])
def get_answers(id):
    print(f"🔍 Answers route called with id = {id}")
    if not ObjectId.is_valid(id):
        print("❌ Invalid ID")
        return jsonify({"error": "Invalid ID"}), 400
    
    cand = mongo.db.candidatures.find_one({"_id": ObjectId(id)})
    if not cand:
        print("❌ Application not found")
        return jsonify({"error": "Not found"}), 404
    
    # Récupérer les questions et réponses
    questions = cand.get("custom_questions", [])
    answers = cand.get("answers", [])
    
    print(f"✅ Found: {len(questions)} questions, {len(answers)} answers")
    
    # Normaliser les questions
    normalized_questions = []
    for q in questions:
        if isinstance(q, str):
            normalized_questions.append({"text": q, "type": "text"})
        elif isinstance(q, dict):
            normalized_questions.append({
                "text": q.get("text", q.get("question", "")),
                "type": q.get("type", "text")
            })
        else:
            normalized_questions.append({"text": str(q), "type": "text"})
    
    # S'assurer que les réponses sont dans le bon format
    formatted_answers = []
    for a in answers:
        if isinstance(a, dict):
            formatted_answers.append({
                "question": a.get("question", ""),
                "answer": a.get("answer", "")
            })
    
    return jsonify({
        "questions": normalized_questions,
        "answers": formatted_answers,
        "status": cand.get("questionnaire_status", "pending")
    }), 200


@app.route('/api/admin/candidatures/answers-by-email/<email>', methods=['GET'])
def get_answers_by_email(email):
    clean_email = email.strip().lower()
    candidature = mongo.db.candidatures.find_one(
        {"candidateEmail": {"$regex": f"^{clean_email}$", "$options": "i"}},
        sort=[("date", -1)]
    )
    if not candidature:
        return jsonify({"error": f"No application found for {email}"}), 404
    return jsonify({
        "questions": candidature.get("custom_questions", []),
        "answers": candidature.get("answers", []),
        "status": candidature.get("questionnaire_status", "pending"),
        "candidatureId": str(candidature["_id"])
    }), 200


@app.route('/api/admin/candidatures/<id>/send-questionnaire', methods=['POST'])
def send_questionnaire_to_candidate(id):
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(id)})
    if not candidature:
        return jsonify({"error": "Application not found"}), 404
    
    # S'assurer qu'il y a des questions
    if not candidature.get('custom_questions'):
        job = mongo.db.postes.find_one({"_id": ObjectId(candidature.get('posteId'))}) if candidature.get('posteId') else None
        if job and job.get('evaluation_questions'):
            custom_questions = [{"text": q, "type": "text"} for q in job['evaluation_questions']]
            mongo.db.candidatures.update_one(
                {"_id": ObjectId(id)},
                {"$set": {"custom_questions": custom_questions}}
            )
            candidature = mongo.db.candidatures.find_one({"_id": ObjectId(id)})
    
    # Générer un token s'il n'existe pas
    if not candidature.get('questionnaire_token'):
        token = hashlib.md5(f"{candidature['candidateEmail']}{id}{datetime.utcnow()}".encode()).hexdigest()[:16]
        mongo.db.candidatures.update_one(
            {"_id": ObjectId(id)},
            {"$set": {"questionnaire_token": token, "questionnaire_status": "pending"}}
        )
    else:
        token = candidature['questionnaire_token']
    
    # Construire le lien
    link = f"http://localhost:4200/questionnaire/{token}"
    
    subject = "📋 Formulaire d'évaluation pour votre candidature"
    body = f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <h2 style="color: #00a6a6;">Formulaire d'évaluation</h2>
        <p>Bonjour {candidature.get('candidateName')},</p>
        <p>Merci de bien vouloir répondre à ce questionnaire :</p>
        <p style="margin: 20px 0;">
            <a href="{link}" style="background-color: #00a6a6; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">
                ➡️ Accéder au formulaire d'évaluation
            </a>
        </p>
        <p>Cordialement,<br>L'équipe LCA</p>
    </body>
    </html>
    """
    
    send_real_email(candidature['candidateEmail'], subject, body)
    return jsonify({"message": "Questionnaire envoyé au candidat", "link": link}), 200



@app.route('/api/debug/questionnaire/<email>', methods=['GET'])
def debug_questionnaire(email):
    """Route de débogage pour vérifier le questionnaire"""
    candidature = mongo.db.candidatures.find_one({
        "candidateEmail": {"$regex": f"^{email}$", "$options": "i"},
        "questionnaire_token": {"$exists": True}
    })
    
    if candidature:
        return jsonify({
            "has_questionnaire": True,
            "token": candidature.get('questionnaire_token'),
            "status": candidature.get('questionnaire_status'),
            "questions_count": len(candidature.get('custom_questions', [])),
            "candidature_id": str(candidature['_id'])
        }), 200
    else:
        return jsonify({
            "has_questionnaire": False,
            "message": "No pending questionnaire found for this email"
        }), 200
    

    
@app.route('/api/candidate/has-questionnaire/<email>', methods=['GET'])
def has_questionnaire_fixed(email):
    """Vérifie si un candidat a un questionnaire - Version corrigée sans suppression"""
    print(f"🔍 Checking questionnaire for: {email}")
    
    # Chercher la candidature la plus récente avec un token
    candidature = mongo.db.candidatures.find_one(
        {
            "candidateEmail": {"$regex": f"^{email}$", "$options": "i"},
            "questionnaire_token": {"$exists": True},
        },
        sort=[("date", -1)]
    )
    
    if not candidature:
        print("❌ No questionnaire found")
        return jsonify({"has": False})
    
    questions = candidature.get("custom_questions", [])
    token = candidature.get("questionnaire_token")
    status = candidature.get("questionnaire_status", "pending")
    
    print(f"📋 Found token: {token} with {len(questions)} questions, status: {status}")
    
    # Si déjà répondu, ne pas renvoyer le token
    if status == "answered":
        return jsonify({"has": False, "answered": True})
    
    # ⭐ NE PAS SUPPRIMER LE TOKEN MÊME S'IL Y A PEU DE QUESTIONS ⭐
    # Juste retourner ce qu'on a
    if len(questions) < 10:
        print(f"⚠️ Questionnaire has only {len(questions)} questions, but keeping it")
        # On garde le token tel quel
    
    return jsonify({
        "has": True, 
        "token": token,
        "questions_count": len(questions)
    })

# Dans app.py, vers la ligne 2050, supprimez le fallback et remplacez par :



@app.route('/api/debug/candidature-tokens/<email>', methods=['GET'])
def debug_candidature_tokens(email):
    """Route de débogage pour voir tous les tokens d'un candidat"""
    candidatures = list(mongo.db.candidatures.find(
        {"candidateEmail": email},
        {"_id": 1, "poste": 1, "questionnaire_token": 1, "questionnaire_status": 1, "custom_questions": {"$size": 1}}
    ))
    
    for c in candidatures:
        c['_id'] = str(c['_id'])
        if 'custom_questions' in c:
            c['questions_count'] = len(c.get('custom_questions', []))
    
    return jsonify({
        "candidate": email,
        "candidatures": candidatures
    }), 200












@app.route('/api/admin/clean-questionnaire/<email>', methods=['POST'])
def clean_questionnaire(email):
    """Nettoie les anciens questionnaires pour un candidat"""
    try:
        # Supprimer les anciens tokens pour cet email
        result = mongo.db.candidatures.update_many(
            {"candidateEmail": email},
            {"$unset": {"questionnaire_token": "", "questionnaire_status": "", "custom_questions": ""}}
        )
        
        return jsonify({
            "message": f"Nettoyage effectué pour {email}",
            "modified_count": result.modified_count
        }), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500







# ========== CHATBOT ==========





def get_user_context(email, role):
    """Récupère le contexte utilisateur pour le chatbot"""
    if role == 'candidate':
        cand = mongo.db.candidats.find_one({"email": email})
        if not cand:
            return "Candidate not found."
        
        applications = list(mongo.db.candidatures.find({"candidateEmail": email}).sort("date", -1).limit(5))
        apps_text = "\n".join([f"- {app.get('poste', 'Unknown position')} (score: {app.get('score',0)}%, status: {app.get('status', 'unknown')})" for app in applications])
        
        cv_skills = []
        cv_path = cand.get('cv_path')
        if cv_path:
            full_path = os.path.join(app.config['UPLOAD_FOLDER'], cv_path)
            if os.path.exists(full_path):
                cv_info = extract_cv_info_spacy(full_path)
                cv_skills = cv_info.get('skills', [])
        
        all_jobs = list(mongo.db.postes.find({"status": "active"}))
        jobs_text = ""
        if all_jobs:
            jobs_text = "📋 Offres d'emploi disponibles:\n"
            for job in all_jobs[:10]:
                title = job.get('title', 'Sans titre')
                company = job.get('companyName') or job.get('company') or 'LCA'
                location = job.get('location', 'Non spécifié')
                jobs_text += f"- {title} chez {company} – {location}\n"
        else:
            jobs_text = "Aucune offre d'emploi en base de données.\n"
        
        context = f"""
Candidat: {cand.get('first_name', '')} {cand.get('last_name', '')}
Email: {email}
Nationalité: {cand.get('nationalite', 'Non spécifiée')}
CV: {'Oui' if cv_path else 'Non'}
Compétences extraites: {', '.join(cv_skills) if cv_skills else 'Aucune détectée'}

Candidatures récentes:
{apps_text if apps_text else 'Aucune candidature.'}

{jobs_text}
"""
        return context
    
    elif role == 'admin':
        total_candidates = mongo.db.candidats.count_documents({})
        total_apps = mongo.db.candidatures.count_documents({})
        total_companies = mongo.db.entreprises.count_documents({})
        total_jobs = mongo.db.postes.count_documents({"status": "active"})
        
        top_matches = list(mongo.db.candidatures.find().sort("score", -1).limit(5))
        top_text_lines = []
        for c in top_matches:
            name = c.get('candidateName', 'Inconnu')
            score = c.get('score', 0)
            poste = c.get('poste', 'Poste inconnu')
            top_text_lines.append(f"- {name}: {score}% pour {poste}")
        top_text = "\n".join(top_text_lines) if top_text_lines else "Aucun match."
        
        all_jobs = list(mongo.db.postes.find({"status": "active"}))
        jobs_text = ""
        if all_jobs:
            jobs_text = "📋 Offres d'emploi actives:\n"
            for job in all_jobs[:10]:
                title = job.get('title', 'Sans titre')
                company = job.get('companyName') or job.get('company') or 'LCA'
                location = job.get('location', 'Non spécifié')
                featured = "⭐" if job.get('featured') else ""
                jobs_text += f"- {featured} {title} chez {company} – {location}\n"
        else:
            jobs_text = "Aucune offre d'emploi en base de données.\n"
        
        context = f"""
👑 **Vue Admin LCA**

📊 **Statistiques globales:**
- Total candidats: {total_candidates}
- Total candidatures: {total_apps}
- Entreprises partenaires: {total_companies}
- Offres actives: {total_jobs}

🏆 **Top 5 meilleurs scores:**
{top_text}

{'-' * 40}
{ jobs_text }
{'-' * 40}

L'admin peut me demander de:
- Filtrer/trier les offres par mot-clé
- Analyser des candidats
- Consulter les statistiques
- Obtenir des recommandations
"""
        return context
    
    return "Rôle non reconnu."




@app.route('/api/chatbot/ask', methods=['POST'])
def chatbot_ask():
    """Route principale du chatbot avec détection d'intention améliorée"""
    try:
        data = request.json
        message = data.get('message', '').strip()
        role = data.get('role', 'candidate')
        email = data.get('email')
        
        if not message:
            return jsonify({"reply": "📝 Veuillez écrire un message."}), 400
        
        message_lower = message.lower()
        
        # === DÉTECTION D'INTENTION AMÉLIORÉE ===
        
        # 1. Détection des salutations
        greetings = ['bonjour', 'salut', 'hello', 'hi', 'coucou', 'hey']
        if any(greeting in message_lower for greeting in greetings):
            name = ""
            if role == 'candidate' and email:
                cand = mongo.db.candidats.find_one({"email": email})
                if cand:
                    name = f" {cand.get('first_name', '')}"
            return jsonify({
                "reply": f"👋 Bonjour{name} ! Je suis votre assistant IA LCA. Comment puis-je vous aider aujourd'hui ?\n\n💡 Suggestions :\n• Afficher les offres d'emploi\n• Filtrer par mot-clé (ex: 'filtrer Python')\n• Trier les offres (ex: 'trier par date')\n• Analyser mon CV\n• Voir le statut de mes candidatures"
            }), 200
        
        # 2. Détection "aide" ou "help"
        if any(word in message_lower for word in ['aide', 'help', 'commandes', 'que puis-je', 'possibilités', 'fonctionnalités']):
            help_text = """🤖 **Mes capacités :**

📋 **Offres d'emploi**
• `liste des offres` - Voir toutes les offres
• `filtrer [mot-clé]` - Filtrer par mot-clé (ex: "filtrer Python")
• `trier par date` - Trier par date de publication
• `trier par popularité` - Trier par nombre de candidatures

👤 **Candidats** (Admin uniquement)
• `top candidats` - Meilleurs scores de matching
• `analyser [email]` - Analyser un candidat
• `comparer [email1] et [email2]` - Comparer deux candidats

📊 **Statistiques**
• `statistiques` - Vue globale des données

💡 **Conseils**
• `conseil CV` - Améliorer votre CV
• `tendances` - Compétences recherchées

Que souhaitez-vous faire ?"""
            return jsonify({"reply": help_text}), 200
        
        # ⭐ NOUVEAU : Détection pour exécuter filtrer ET trier en une commande
        if ('filtr' in message_lower or 'offre' in message_lower) and ('trier' in message_lower or 'tri' in message_lower):
            import re
            
            # Extraire le mot-clé pour filtrer
            filter_keyword = None
            filter_patterns = [
                r'filtrer\s+([a-zA-Z0-9éèêëçàù]+)',
                r'offres?\s+([a-zA-Z0-9éèêëçàù]+)',
                r'([a-zA-Z0-9éèêëçàù]+)\s+(?:et|avec)',
            ]
            for pattern in filter_patterns:
                match = re.search(pattern, message_lower)
                if match:
                    filter_keyword = match.group(1)
                    break
            
            # 🔥 Test spécifique pour "filtrer et trier" sans mot-clé
            if not filter_keyword and ('filtrer et trier' in message_lower or 'filtrer et trier les offres' in message_lower):
                filter_keyword = 'développeur'
                default_note = "\n\n💡 (J'ai utilisé 'développeur' comme mot-clé par défaut pour vous montrer un exemple)"
            
            # Déterminer le type de tri
            sort_type = 'date'  # par défaut
            if 'popularité' in message_lower or 'candidature' in message_lower:
                sort_type = 'popularity'
            elif 'titre' in message_lower or 'alphab' in message_lower:
                sort_type = 'title'
            
            if filter_keyword:
                # Récupérer les jobs filtrés
                jobs = list(mongo.db.postes.find({
                    "status": "active",
                    "$or": [
                        {"title": {"$regex": filter_keyword, "$options": "i"}},
                        {"description": {"$regex": filter_keyword, "$options": "i"}},
                        {"keywords": {"$regex": filter_keyword, "$options": "i"}}
                    ]
                }).limit(10))
                
                if jobs:
                    # Trier les résultats
                    if sort_type == 'date':
                        jobs.sort(key=lambda x: x.get('created_at', datetime.min), reverse=True)
                        sort_name = "date de publication"
                    elif sort_type == 'popularity':
                        for j in jobs:
                            j['app_count'] = mongo.db.candidatures.count_documents({"posteId": str(j['_id'])})
                        jobs.sort(key=lambda x: x.get('app_count', 0), reverse=True)
                        sort_name = "nombre de candidatures"
                    else:
                        jobs.sort(key=lambda x: x.get('title', '').lower())
                        sort_name = "titre alphabétique"
                    
                    job_list = []
                    for j in jobs[:10]:
                        title = j.get('title', 'Sans titre')
                        company = j.get('companyName') or j.get('company', 'LCA')
                        location = j.get('location', 'Non spécifié')
                        app_count = j.get('app_count', 0)
                        if sort_type == 'popularity':
                            job_list.append(f"  • {title} - {company} ({location}) - {app_count} candidature(s)")
                        else:
                            job_list.append(f"  • {title} - {company} ({location})")
                    
                    reply = f"🔍 **{len(jobs)} offre(s) contenant '{filter_keyword}', triées par {sort_name}:**\n\n" + "\n".join(job_list)
                    if 'default_note' in locals():
                        reply += default_note
                    if len(jobs) >= 10:
                        reply += "\n\n📌 Il y a peut-être plus d'offres. Voulez-vous affiner votre recherche ?"
                else:
                    reply = f"🔍 Aucune offre ne contient '{filter_keyword}'. Essayez un autre mot-clé comme 'développeur', 'python' ou 'comptable'."
            else:
                # Pas de mot-clé trouvé, demander
                reply = "🔍 **Pour filtrer ET trier, indiquez un mot-clé !**\n\n"
                reply += "Exemples :\n"
                reply += "• `Filtrer Python et trier par date`\n"
                reply += "• `Offres développeur triées par popularité`\n"
                reply += "• `Affiche les offres Anglais triées par titre`\n\n"
                reply += "Quel mot-clé souhaitez-vous utiliser ?"
            return jsonify({"reply": reply}), 200
        
        # 3. Détection AVANCÉE pour "liste des offres"
        show_offers_keywords = [
            'liste des offres', 'afficher offres', 'toutes les offres', 
            'montre les offres', 'offres dispo', 'voir les offres',
            'les offres', 'afficher les offres', 'offres d emploi',
            'offres disponibles', 'liste offres', 'show offers'
        ]
        if any(keyword in message_lower for keyword in show_offers_keywords):
            jobs = list(mongo.db.postes.find({"status": "active"}).limit(10))
            if jobs:
                job_list = []
                for j in jobs:
                    title = j.get('title', 'Sans titre')
                    company = j.get('companyName') or j.get('company', 'LCA')
                    location = j.get('location', 'Distance')
                    job_list.append(f"  • {title} - {company} ({location})")
                reply = f"📋 Voici les dernières offres d'emploi ({len(jobs)}):\n\n" + "\n".join(job_list) + "\n\n✨ Pour filtrer, dites 'filtrer [mot-clé]' (ex: 'filtrer Python') ou 'trier par date'"
            else:
                reply = "📋 Aucune offre d'emploi disponible pour le moment."
            return jsonify({"reply": reply}), 200
        
        # 4. Détection pour "FILTRER" uniquement
        if any(kw in message_lower for kw in ['filtrer', 'filtre', 'filter']):
            import re
            keyword_match = re.search(r'(?:filtrer|filtre|filter)\s+(?:par\s+)?["\']?([a-zA-Z0-9\s#+éèêëçàù]+)["\']?', message_lower)
            if keyword_match:
                keyword = keyword_match.group(1).strip()
                jobs = list(mongo.db.postes.find({
                    "status": "active",
                    "$or": [
                        {"title": {"$regex": keyword, "$options": "i"}},
                        {"description": {"$regex": keyword, "$options": "i"}},
                        {"keywords": {"$regex": keyword, "$options": "i"}}
                    ]
                }).limit(10))
                if jobs:
                    job_list = []
                    for j in jobs:
                        title = j.get('title', 'Sans titre')
                        company = j.get('companyName') or j.get('company', 'LCA')
                        location = j.get('location', 'Non spécifié')
                        job_list.append(f"  • {title} - {company} ({location})")
                    reply = f"🔍 J'ai trouvé {len(jobs)} offre(s) contenant '{keyword}':\n\n" + "\n".join(job_list)
                else:
                    reply = f"🔍 Aucune offre ne contient '{keyword}'. Essayez 'développeur', 'python' ou 'javascript'."
            else:
                reply = "🔍 Pour filtrer les offres, dites 'filtrer [mot-clé]' (ex: 'filtrer Python')"
            return jsonify({"reply": reply}), 200
        
        # 5. Détection pour "TRIER" uniquement
        if any(kw in message_lower for kw in ['trier', 'tri', 'sort']):
            jobs = list(mongo.db.postes.find({"status": "active"}))
            sort_type = 'date'
            sort_name = "date de publication"
            sort_emoji = "📅"
            if 'popularité' in message_lower or 'candidature' in message_lower:
                for j in jobs:
                    j['app_count'] = mongo.db.candidatures.count_documents({"posteId": str(j['_id'])})
                jobs.sort(key=lambda x: x.get('app_count', 0), reverse=True)
                sort_name = "nombre de candidatures"
                sort_emoji = "📊"
            elif 'titre' in message_lower or 'alphab' in message_lower:
                jobs.sort(key=lambda x: x.get('title', '').lower())
                sort_name = "titre alphabétique"
                sort_emoji = "🔤"
            else:
                jobs.sort(key=lambda x: x.get('created_at', datetime.min), reverse=True)
            
            if jobs:
                job_list = []
                for j in jobs[:5]:
                    title = j.get('title', 'Sans titre')
                    company = j.get('companyName') or j.get('company', 'LCA')
                    job_list.append(f"  • {title} - {company}")
                reply = f"{sort_emoji} Offres triées par {sort_name}:\n\n" + "\n".join(job_list)
                if len(jobs) > 5:
                    reply += f"\n\n📌 +{len(jobs)-5} autres offres..."
            else:
                reply = "📊 Aucune offre d'emploi disponible."
            return jsonify({"reply": reply}), 200
        
        # ... (reste des intentions et fallback Groq)
        
        context = get_user_context(email, role) if email else ""
        
        system_prompt = f"""Tu es un assistant IA expert en recrutement pour la plateforme LCA (Laura Connecting Agency).

CONTEXTE UTILISATEUR:
{context}

RÈGLES IMPORTANTES:
1. Réponds TOUJOURS en FRANÇAIS
2. Sois concis, clair et professionnel (max 300 mots)
3. Utilise des émojis pour rendre la réponse plus agréable

Réponds de manière naturelle et utile."""
        
        try:
            if groq_client:
                response = groq_client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": message}
                    ],
                    temperature=0.7,
                    max_tokens=500
                )
                reply = response.choices[0].message.content.strip()
            else:
                reply = "🤖 Service IA en cours de configuration. Veuillez réessayer dans quelques instants."
        except Exception as e:
            print(f"Groq error: {e}")
            reply = "⚠️ Service IA temporairement indisponible. Posez votre question à l'administrateur."
        
        return jsonify({"reply": reply}), 200
        
    except Exception as e:
        print(f"❌ Chatbot error: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"reply": "❌ Une erreur technique est survenue. L'équipe a été notifiée."}), 500


















#============================================================================================




# Evaluation (template)
@app.route('/api/candidate/evaluation/<token>', methods=['GET'])
def get_evaluation_form(token):
    candidature = mongo.db.candidatures.find_one({"evaluation_token": token})
    if not candidature:
        return jsonify({"error": "Invalid or expired link"}), 404
    template = load_evaluation_template()
    return jsonify({
        "candidatureId": str(candidature["_id"]),
        "questions": template["technical"],
        "softSkills": template["soft"]
    })

@app.route('/api/candidate/evaluation/submit', methods=['POST'])
def submit_evaluation():
    data = request.json
    cand_id = data.get('candidatureId')
    answers = data.get('answers')
    soft_scores = data.get('softScores')
    mongo.db.candidatures.update_one(
        {"_id": ObjectId(cand_id)},
        {"$set": {
            "evaluation_answers": answers,
            "soft_skills_evaluation": soft_scores,
            "evaluation_filled_at": datetime.utcnow()
        }}
    )
    return jsonify({"message": "Evaluation saved"})

# Compatibility score
@app.route('/api/jobs/<job_id>/compatibility/<candidate_email>', methods=['GET'])
def get_compatibility_score(job_id, candidate_email):
    candidate = mongo.db.candidats.find_one({"email": candidate_email})
    if not candidate or not candidate.get('cv_path'):
        return jsonify({"score": 0}), 200
    job = mongo.db.postes.find_one({"_id": ObjectId(job_id)})
    if not job:
        return jsonify({"score": 0}), 200
    cv_path = os.path.join(app.config['UPLOAD_FOLDER'], candidate['cv_path'])
    if not os.path.exists(cv_path):
        return jsonify({"score": 0}), 200
    cv_info = extract_cv_info_spacy(cv_path)
    job_keywords = [kw.lower() for kw in job.get('keywords', [])]
    if not job_keywords:
        score = 0
    else:
        found = sum(1 for kw in job_keywords if kw in cv_info['full_text'])
        score = int((found / len(job_keywords)) * 100)
    return jsonify({"score": score}), 200

@app.route('/api/jobs/new-since/<last_date>', methods=['GET'])
def get_new_jobs_since(last_date):
    try:
        since = datetime.fromisoformat(last_date)
        new_jobs = list(mongo.db.postes.find({"created_at": {"$gt": since}, "status": "active"}))
        return jsonify([{
            "id": str(j['_id']),
            "title": j.get('title'),
            "description": j.get('description'),
            "keywords": j.get('keywords', [])
        } for j in new_jobs]), 200
    except:
        return jsonify([]), 200

# ========== INSIGHTS BACKEND ==========
@app.route('/api/trending-skills', methods=['GET'])
def get_trending_skills():
    pipeline = [
        {"$match": {"status": "active"}},
        {"$unwind": "$keywords"},
        {"$group": {"_id": "$keywords", "count": {"$sum": 1}}},
        {"$sort": {"count": -1}},
        {"$limit": 10}
    ]
    result = list(mongo.db.postes.aggregate(pipeline))
    skills = [item['_id'] for item in result if item['_id']]
    return jsonify(skills), 200

@app.route('/api/candidate/cv-strength/<email>', methods=['GET'])
def cv_strength(email):
    cand = mongo.db.candidats.find_one({"email": email})
    if not cand or not cand.get('cv_path'):
        return jsonify({"score": 0, "tip": "Upload your CV to evaluate its strength."}), 200
    path = os.path.join(app.config['UPLOAD_FOLDER'], cand['cv_path'])
    if not os.path.exists(path):
        return jsonify({"score": 20, "tip": "CV file not found."}), 200
    with open(path, 'rb') as f:
        pdf = PyPDF2.PdfReader(f)
        text = "".join([page.extract_text() or "" for page in pdf.pages])
    score = 30
    tip = "Basic CV."
    if len(text) > 1000:
        score += 20
        tip = "Good content volume."
    skill_keywords = ["python", "javascript", "angular", "react", "sql", "java", "php", "laravel"]
    found = sum(1 for kw in skill_keywords if kw in text.lower())
    score += min(found * 5, 30)
    if found >= 4:
        tip = "Excellent, several key skills detected."
    elif found >= 2:
        tip = "Good start, add more technologies."
    else:
        tip = "Add technical keywords to improve your score."
    return jsonify({"score": min(score, 100), "tip": tip}), 200

@app.route('/api/user/weekly-action/<email>', methods=['GET'])
def get_weekly_action(email):
    current_week = datetime.utcnow().strftime("%Y-W%W")
    action = mongo.db.user_actions.find_one({"email": email, "week": current_week})
    if not action:
        actions_pool = [
            "Add your nationality to your profile",
            "Upload a professional profile photo",
            "Apply to a job that interests you",
            "Send a message to the administrator",
            "Review your saved jobs and apply to one of them"
        ]
        new_action = {"email": email, "week": current_week, "action": random.choice(actions_pool), "completed": False, "reward_points": 10}
        mongo.db.user_actions.insert_one(new_action)
        action = new_action
    return jsonify({"action": action["action"], "completed": action["completed"]}), 200

@app.route('/api/candidate/skill-gap/<email>', methods=['GET'])
def skill_gap(email):
    cand = mongo.db.candidats.find_one({"email": email})
    if not cand or not cand.get('cv_path'):
        return jsonify({"missing_skills": []}), 200
    path = os.path.join(app.config['UPLOAD_FOLDER'], cand['cv_path'])
    cv_info = extract_cv_info_spacy(path) if os.path.exists(path) else {"skills": []}
    cv_skills = set([s.lower() for s in cv_info.get('skills', [])])
    all_jobs = list(mongo.db.postes.find({"status": "active"}))
    all_keywords = set()
    for job in all_jobs:
        for kw in job.get('keywords', []):
            all_keywords.add(kw.lower())
    missing = list(all_keywords - cv_skills)
    return jsonify({"missing_skills": missing[:8]}), 200

@app.route('/api/recommend-courses/<email>', methods=['GET'])
def recommend_courses(email):
    missing = ['Angular', 'React', 'Python']
    courses = [{"name": f"Master {skill} - Complete course", "platform": "Udemy / Coursera", "link": f"https://www.udemy.com/courses/search/?q={skill}"} for skill in missing]
    return jsonify(courses), 200

@app.route('/api/events/upcoming', methods=['GET'])
def upcoming_events():
    events = [
        {"date": "2025-05-20", "title": "Webinar: Negotiate your salary", "link": "https://lca.com/webinar"},
        {"date": "2025-06-05", "title": "CV & LinkedIn workshop", "link": "#"}
    ]
    return jsonify(events), 200






# ========== FEEDBACK & TESTIMONIALS ==========
@app.route('/api/feedback', methods=['POST'])
def submit_feedback():
    data = request.json
    email = data.get('email')
    rating = data.get('rating')
    comment = data.get('comment', '')
    if not email or not rating:
        return jsonify({"error": "Email and rating required"}), 400
    
    existing = mongo.db.feedback.find_one({"email": email})
    
    if existing:
        # Mettre à jour le feedback existant au lieu de bloquer
        mongo.db.feedback.update_one(
            {"email": email},
            {"$set": {
                "rating": rating,
                "comment": comment,
                "updated_at": datetime.utcnow()
            }}
        )
        return jsonify({"message": "Your feedback has been updated!"}), 200
    else:
        mongo.db.feedback.insert_one({
            "email": email,
            "rating": rating,
            "comment": comment,
            "status": "pending",
            "created_at": datetime.utcnow()
        })
        mongo.db.notifications.insert_one({
            "recipient": "admin",
            "recipientEmail": "admin@lca.com",
            "type": "new_feedback",
            "title": "⭐ New candidate feedback",
            "message": f"{email} gave a rating of {rating}/5 and a comment.",
            "status": "unread",
            "date": datetime.utcnow()
        })
        return jsonify({"message": "Thank you for your feedback!"}), 200
@app.route('/api/admin/feedback/pending', methods=['GET'])
@admin_required()
def get_pending_feedback():
    pending = list(mongo.db.feedback.find({"status": "pending"}))
    for fb in pending:
        fb['_id'] = str(fb['_id'])
        fb['created_at'] = fb['created_at'].isoformat() if fb.get('created_at') else None
    return jsonify(pending), 200

@app.route('/api/admin/feedback/approve/<feedback_id>', methods=['PUT'])
@admin_required()
def approve_feedback(feedback_id):
    if not ObjectId.is_valid(feedback_id):
        return jsonify({"error": "Invalid ID"}), 400
    fb = mongo.db.feedback.find_one({"_id": ObjectId(feedback_id)})
    if not fb:
        return jsonify({"error": "Feedback not found"}), 404
    mongo.db.feedback.update_one({"_id": ObjectId(feedback_id)}, {"$set": {"status": "approved"}})
    testimonial = {
        "name": fb.get("email").split('@')[0],
        "job_title": "LCA Candidate",
        "comment": fb.get("comment"),
        "photo": "assets/default-user.png",
        "rating": fb.get("rating"),
        "approved": True,
        "created_at": datetime.utcnow()
    }
    candidat = mongo.db.candidats.find_one({"email": fb.get("email")})
    if candidat:
        testimonial["name"] = f"{candidat.get('first_name', '')} {candidat.get('last_name', '')}".strip() or candidat.get('email')
        if candidat.get('photo_url'):
            testimonial["photo"] = f"/static/uploads/photos/{candidat.get('photo_url')}"
    mongo.db.testimonials.insert_one(testimonial)
    return jsonify({"message": "Feedback approved and added to testimonials"}), 200

@app.route('/api/admin/feedback/reject/<feedback_id>', methods=['DELETE'])
@admin_required()
def reject_feedback(feedback_id):
    if not ObjectId.is_valid(feedback_id):
        return jsonify({"error": "Invalid ID"}), 400
    result = mongo.db.feedback.delete_one({"_id": ObjectId(feedback_id), "status": "pending"})
    if result.deleted_count:
        return jsonify({"message": "Feedback deleted"}), 200
    return jsonify({"error": "Feedback not found or already processed"}), 404

@app.route('/api/testimonials', methods=['GET'])
def get_testimonials():
    base_url = request.host_url.rstrip('/')
    stories = list(mongo.db.testimonials.find({"approved": {"$ne": False}}).sort("created_at", -1).limit(3))
    return jsonify([{
        'comment': s.get('comment', ''),
        'candidateName': s.get('name', 'Anonymous'),
        'position': s.get('job_title', 'Candidate'),
        'photoUrl': base_url + s.get('photo', 'assets/default-user.png') if s.get('photo', '').startswith('/') else s.get('photo', 'assets/default-user.png')
    } for s in stories])







# ========== GEMINI CHATBOT (alternate) ==========
@app.route('/chatbot', methods=['POST'])
def chatbot():
    data = request.get_json()
    user_message = data.get('message')
    
    if not client:
        return jsonify({"reply": "⚠️ Service Gemini non disponible. Veuillez réessayer plus tard."}), 200
    
    try:
        prompt = f"You are the assistant of the LCA platform (Laura Connecting Agency). Answer candidates' questions about recruitment in French. Here is the question: {user_message}"
        response = client.models.generate_content(
            model="gemini-2.0-flash-exp",
            contents=prompt
        )
        return jsonify({"reply": response.text})
    except Exception as e:
        print(f"Gemini error: {e}")
        return jsonify({"reply": "⚠️ Erreur avec le service IA. Veuillez réessayer."}), 200



























# ========== ROUTE POUR SAUVEGARDER LE QUESTIONNAIRE D'UN JOB ==========
@app.route('/api/jobs/<job_id>/questionnaire', methods=['PATCH'])
def save_job_questionnaire(job_id):
    """Sauvegarde le questionnaire personnalisé d'un job"""
    try:
        if not ObjectId.is_valid(job_id):
            return jsonify({"error": "Invalid ID"}), 400
        
        data = request.get_json()
        question_sections = data.get('questionSections', [])
        
        result = mongo.db.postes.update_one(
            {"_id": ObjectId(job_id)},
            {"$set": {"customQuestionSections": question_sections}}
        )
        
        if result.matched_count == 0:
            return jsonify({"error": "Job not found"}), 404
        
        return jsonify({
            "message": "Questionnaire saved successfully",
            "sections_count": len(question_sections)
        }), 200
    except Exception as e:
        print(f"❌ Error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/debug/candidature-full/<candidature_id>', methods=['GET'])
def debug_candidature_full(candidature_id):
    """Debug complet d'une candidature"""
    if not ObjectId.is_valid(candidature_id):
        return jsonify({"error": "Invalid ID"}), 400
    
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(candidature_id)})
    if not candidature:
        return jsonify({"error": "Not found"}), 404
    
    return jsonify({
        "_id": str(candidature["_id"]),
        "candidateName": candidature.get("candidateName"),
        "candidateEmail": candidature.get("candidateEmail"),
        "poste": candidature.get("poste"),
        "custom_questions": candidature.get("custom_questions", []),
        "answers": candidature.get("answers", []),
        "questionnaire_status": candidature.get("questionnaire_status"),
        "questionnaire_token": candidature.get("questionnaire_token")
    }), 200



@app.route('/api/admin/debug-answers/<candidature_id>', methods=['GET'])
def debug_answers(candidature_id):
    """Debug pour voir les réponses d'une candidature"""
    if not ObjectId.is_valid(candidature_id):
        return jsonify({"error": "Invalid ID"}), 400
    
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(candidature_id)})
    if not candidature:
        return jsonify({"error": "Not found"}), 404
    
    answers = candidature.get("answers", [])
    questions = candidature.get("custom_questions", [])
    
    return jsonify({
        "candidature_id": str(candidature["_id"]),
        "candidateName": candidature.get("candidateName"),
        "questions_count": len(questions),
        "answers_count": len(answers),
        "status": candidature.get("questionnaire_status"),
        "answers": answers  # Retourner les réponses complètes
    }), 200






# ========== FINAL DECISION ROUTES (Retenu / Non retenu) ==========

@app.route('/api/candidatures/final-decision', methods=['PUT'])
def update_final_decision():
    """
    Met à jour la décision finale d'une candidature (retenu / non retenu)
    """
    try:
        data = request.json
        candidature_id = data.get('id')
        final_decision = data.get('final_decision')  # 'hired' ou 'rejected'
        
        if not candidature_id or not final_decision:
            return jsonify({"error": "Missing required fields"}), 400
        
        if final_decision not in ['hired', 'rejected']:
            return jsonify({"error": "Invalid decision. Must be 'hired' or 'rejected'"}), 400
        
        if not ObjectId.is_valid(candidature_id):
            return jsonify({"error": "Invalid ID"}), 400
        
        candidature = mongo.db.candidatures.find_one({"_id": ObjectId(candidature_id)})
        if not candidature:
            return jsonify({"error": "Application not found"}), 404
        
        # Mettre à jour la décision finale
        mongo.db.candidatures.update_one(
            {"_id": ObjectId(candidature_id)},
            {"$set": {
                "final_decision": final_decision,
                "final_decision_date": datetime.utcnow(),
                "status": "closed" if final_decision in ['hired', 'rejected'] else candidature.get('status')
            }}
        )
        
        # Envoyer un email au candidat
        candidate_email = candidature.get('candidateEmail')
        candidate_name = candidature.get('candidateName', 'Candidat')
        poste = candidature.get('poste', 'le poste')
        
        if final_decision == 'hired':
            subject = "🎉 Félicitations ! Vous avez été retenu"
            body = f"""
            <html>
            <body style="font-family: Arial, sans-serif;">
                <h2 style="color: #10b981;">Félicitations {candidate_name} !</h2>
                <p>Nous sommes ravis de vous annoncer que vous avez été <strong>retenu</strong> pour le poste de <strong>{poste}</strong>.</p>
                <p>Un recruteur vous contactera très prochainement pour finaliser votre intégration.</p>
                <p>Nous vous félicitons pour votre parcours et avons hâte de travailler avec vous !</p>
                <br>
                <p>Cordialement,<br>L'équipe LCA</p>
            </body>
            </html>
            """
        else:
            subject = "📌 Suite à votre candidature"
            body = f"""
            <html>
            <body style="font-family: Arial, sans-serif;">
                <h2 style="color: #ef4444;">Bonjour {candidate_name},</h2>
                <p>Nous vous remercions pour votre candidature au poste de <strong>{poste}</strong>.</p>
                <p>Après étude approfondie de votre dossier, nous avons décidé de ne pas donner suite à votre candidature.</p>
                <p>Cependant, votre profil nous a intéressés et nous vous encourageons vivement à postuler à d'autres offres qui correspondent mieux à votre parcours.</p>
                <p>Nous vous souhaitons beaucoup de succès dans vos recherches !</p>
                <br>
                <p>Cordialement,<br>L'équipe LCA</p>
            </body>
            </html>
            """
        
        send_real_email(candidate_email, subject, body)
        
        # Ajouter une notification
        mongo.db.notifications.insert_one({
            "recipient": "candidate",
            "recipientEmail": candidate_email,
            "type": "final_decision",
            "title": "✅ Décision finale" if final_decision == 'hired' else "📌 Décision finale",
            "message": f"Votre candidature pour {poste} a été {'retenue' if final_decision == 'hired' else 'non retenue'}.",
            "status": "unread",
            "date": datetime.utcnow(),
            "candidatureId": candidature_id
        })
        
        # Notification pour l'admin
        mongo.db.notifications.insert_one({
            "recipient": "admin",
            "recipientEmail": "admin@lca.com",
            "type": "final_decision_made",
            "title": "🎯 Décision finale enregistrée",
            "message": f"{candidate_name} - {'✅ Retenu' if final_decision == 'hired' else '❌ Non retenu'} pour {poste}",
            "status": "unread",
            "date": datetime.utcnow(),
            "candidatureId": candidature_id
        })
        
        return jsonify({
            "message": f"Décision '{'retenu' if final_decision == 'hired' else 'non retenu'}' enregistrée",
            "final_decision": final_decision
        }), 200
        
    except Exception as e:
        print(f"❌ Error in update_final_decision: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/api/admin/final-decisions-stats', methods=['GET'])
def get_final_decisions_stats():
    """
    Retourne les statistiques des décisions finales
    """
    try:
        hired_count = mongo.db.candidatures.count_documents({"final_decision": "hired"})
        rejected_count = mongo.db.candidatures.count_documents({"final_decision": "rejected"})
        pending_count = mongo.db.candidatures.count_documents({
            "status": "sent_to_client",
            "final_decision": {"$exists": False}
        })
        
        # Détail par poste
        pipeline_hired_by_job = [
            {"$match": {"final_decision": "hired"}},
            {"$group": {"_id": "$poste", "count": {"$sum": 1}}},
            {"$sort": {"count": -1}},
            {"$limit": 5}
        ]
        hired_by_job = list(mongo.db.candidatures.aggregate(pipeline_hired_by_job))
        
        # Détail par entreprise
        pipeline_hired_by_company = [
            {"$match": {"final_decision": "hired", "company": {"$exists": True}}},
            {"$group": {"_id": "$company", "count": {"$sum": 1}}},
            {"$sort": {"count": -1}},
            {"$limit": 5}
        ]
        hired_by_company = list(mongo.db.candidatures.aggregate(pipeline_hired_by_company))
        
        return jsonify({
            "hired": hired_count,
            "rejected": rejected_count,
            "pending": pending_count,
            "hired_by_job": [{"job": item["_id"], "count": item["count"]} for item in hired_by_job],
            "hired_by_company": [{"company": item["_id"], "count": item["count"]} for item in hired_by_company]
        }), 200
        
    except Exception as e:
        print(f"❌ Error in get_final_decisions_stats: {e}")
        return jsonify({"hired": 0, "rejected": 0, "pending": 0}), 200


@app.route('/api/admin/request-client-feedback', methods=['POST'])
def request_client_feedback():
    """
    Envoie une demande de feedback au client pour un candidat
    """
    try:
        data = request.json
        client_email = data.get('clientEmail')
        message = data.get('message')
        candidate_id = data.get('candidateId')
        candidate_name = data.get('candidateName')
        position = data.get('position')
        
        if not client_email or not message:
            return jsonify({"error": "Missing required fields"}), 400
        
        # Générer un token unique pour ce feedback
        token = hashlib.md5(f"{client_email}{candidate_id}{datetime.utcnow()}".encode()).hexdigest()[:16]
        
        # Sauvegarder la demande de feedback
        mongo.db.client_feedback_requests.insert_one({
            "client_email": client_email,
            "candidate_id": candidate_id,
            "candidate_name": candidate_name,
            "position": position,
            "token": token,
            "status": "pending",
            "created_at": datetime.utcnow()
        })
        
        # Construire le lien de feedback
        feedback_link = f"http://localhost:4200/client/feedback/{token}"
        
        # Envoyer l'email au client
        subject = f"📋 Demande de feedback - Candidat {candidate_name}"
        body = f"""
        <html>
        <body style="font-family: Arial, sans-serif;">
            <h2 style="color: #3b82f6;">Demande de feedback</h2>
            <p>Bonjour,</p>
            <p>Nous vous remercions d'avoir examiné le dossier de <strong>{candidate_name}</strong> pour le poste de <strong>{position}</strong>.</p>
            <p>Pourriez-vous nous faire part de votre décision concernant ce candidat ?</p>
            
            <p style="margin: 30px 0;">
                <a href="{feedback_link}" style="background-color: #3b82f6; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px;">
                    📝 Donner mon avis
                </a>
            </p>
            
            <p>Vous pouvez également répondre directement à cet email.</p>
            <br>
            <p>Cordialement,<br>L'équipe LCA</p>
        </body>
        </html>
        """
        
        send_real_email(client_email, subject, body)
        
        return jsonify({"message": "Demande de feedback envoyée au client", "token": token}), 200
        
    except Exception as e:
        print(f"❌ Error in request_client_feedback: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/client/feedback/<token>', methods=['GET'])
def get_client_feedback_form(token):
    """
    Affiche le formulaire de feedback pour le client
    """
    try:
        request_data = mongo.db.client_feedback_requests.find_one({"token": token})
        if not request_data:
            return jsonify({"error": "Invalid or expired link"}), 404
        
        return jsonify({
            "token": token,
            "candidate_name": request_data.get("candidate_name"),
            "position": request_data.get("position"),
            "status": request_data.get("status")
        }), 200
        
    except Exception as e:
        print(f"❌ Error in get_client_feedback_form: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/client/submit-feedback', methods=['POST'])
def submit_client_feedback():
    """
    Enregistre le feedback du client (retenu / non retenu)
    """
    try:
        data = request.json
        token = data.get('token')
        decision = data.get('decision')  # 'hired' ou 'rejected'
        comments = data.get('comments', '')
        
        if not token or not decision:
            return jsonify({"error": "Missing required fields"}), 400
        
        if decision not in ['hired', 'rejected']:
            return jsonify({"error": "Invalid decision"}), 400
        
        request_data = mongo.db.client_feedback_requests.find_one({"token": token})
        if not request_data:
            return jsonify({"error": "Invalid or expired link"}), 404
        
        # Mettre à jour la demande de feedback
        mongo.db.client_feedback_requests.update_one(
            {"token": token},
            {"$set": {
                "decision": decision,
                "comments": comments,
                "status": "completed",
                "completed_at": datetime.utcnow()
            }}
        )
        
        # Mettre à jour la candidature avec la décision du client
        candidate_id = request_data.get("candidate_id")
        if candidate_id and ObjectId.is_valid(candidate_id):
            mongo.db.candidatures.update_one(
                {"_id": ObjectId(candidate_id)},
                {"$set": {
                    "client_feedback": {
                        "decision": decision,
                        "comments": comments,
                        "received_at": datetime.utcnow()
                    }
                }}
            )
            
            # Si le client dit "retenu", on met aussi final_decision
            if decision == 'hired':
                mongo.db.candidatures.update_one(
                    {"_id": ObjectId(candidate_id)},
                    {"$set": {
                        "final_decision": "hired",
                        "final_decision_date": datetime.utcnow()
                    }}
                )
        
        # Notification à l'admin
        mongo.db.notifications.insert_one({
            "recipient": "admin",
            "recipientEmail": "admin@lca.com",
            "type": "client_feedback_received",
            "title": "📬 Feedback client reçu",
            "message": f"Le client a {'retenu' if decision == 'hired' else 'non retenu'} {request_data.get('candidate_name')}",
            "status": "unread",
            "date": datetime.utcnow()
        })
        
        return jsonify({
            "message": f"Merci pour votre retour ! Le candidat a été {'retenu' if decision == 'hired' else 'non retenu'}."
        }), 200
        
    except Exception as e:
        print(f"❌ Error in submit_client_feedback: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/admin/export-decisions', methods=['GET'])
@admin_required()
def export_decisions():
    """
    Exporte toutes les décisions finales en CSV
    """
    try:
        candidatures = list(mongo.db.candidatures.find({
            "final_decision": {"$exists": True}
        }).sort("final_decision_date", -1))
        
        data = []
        for c in candidatures:
            data.append({
                "Candidat": c.get("candidateName", ""),
                "Email": c.get("candidateEmail", ""),
                "Poste": c.get("poste", ""),
                "Entreprise": c.get("company", ""),
                "Décision": "Retenu" if c.get("final_decision") == "hired" else "Non retenu",
                "Date décision": c.get("final_decision_date").strftime("%d/%m/%Y") if c.get("final_decision_date") else "",
                "Score IA": c.get("score", 0)
            })
        
        if not data:
            return jsonify({"error": "No data to export"}), 404
        
        # Générer CSV
        import csv
        import io
        
        output = io.StringIO()
        writer = csv.DictWriter(output, fieldnames=data[0].keys(), delimiter=';')
        writer.writeheader()
        writer.writerows(data)
        
        output.seek(0)
        
        return send_file(
            io.BytesIO(output.getvalue().encode('utf-8-sig')),
            mimetype='text/csv',
            as_attachment=True,
            download_name=f"decisions_finales_{datetime.now().strftime('%Y%m%d')}.csv"
        )
        
    except Exception as e:
        print(f"❌ Error in export_decisions: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/admin/decisions-timeline', methods=['GET'])
def get_decisions_timeline():
    """
    Retourne l'historique des décisions pour le graphique
    """
    try:
        pipeline = [
            {"$match": {"final_decision": {"$exists": True}, "final_decision_date": {"$exists": True}}},
            {"$group": {
                "_id": {
                    "month": {"$dateToString": {"format": "%Y-%m", "date": "$final_decision_date"}},
                    "decision": "$final_decision"
                },
                "count": {"$sum": 1}
            }},
            {"$sort": {"_id.month": 1}}
        ]
        
        results = list(mongo.db.candidatures.aggregate(pipeline))
        
        # Organiser les données
        timeline = {}
        for r in results:
            month = r["_id"]["month"]
            decision = r["_id"]["decision"]
            count = r["count"]
            
            if month not in timeline:
                timeline[month] = {"hired": 0, "rejected": 0}
            
            if decision == "hired":
                timeline[month]["hired"] = count
            else:
                timeline[month]["rejected"] = count
        
        # Convertir en liste pour le frontend
        result = [
            {"month": month, "hired": data["hired"], "rejected": data["rejected"]}
            for month, data in sorted(timeline.items())
        ]
        
        return jsonify(result), 200
        
    except Exception as e:
        print(f"❌ Error in get_decisions_timeline: {e}")
        return jsonify([]), 200
    


# Dans app.py, ajoutez cette route
@app.route('/api/generate-default-avatar/<email>', methods=['GET'])
def generate_default_avatar(email):
    """Génère un avatar par défaut pour un candidat"""
    try:
        from PIL import Image, ImageDraw, ImageFont
        
        # Créer une image
        img = Image.new('RGB', (200, 200), color='#00a6a6')
        d = ImageDraw.Draw(img)
        
        # Prendre la première lettre de l'email
        letter = email[0].upper() if email else '?'
        
        # Dessiner la lettre
        d.text((100, 100), letter, fill="white", anchor="mm")
        
        # Sauvegarder
        filename = secure_filename(f"{email}_avatar.png")
        filepath = os.path.join('uploads/photos', filename)
        img.save(filepath)
        
        return jsonify({"photo_url": filename}), 200
    except Exception as e:
        print(f"Error generating avatar: {e}")
        return jsonify({"error": str(e)}), 500
    



# ========== ÉVALUATION COMPORTEMENTALE & LINGUISTIQUE ==========

@app.route('/api/candidatures/<id>/behavioral-evaluation', methods=['POST'])
def save_behavioral_evaluation(id):
    """
    Sauvegarde l'évaluation comportementale et linguistique d'un candidat
    """
    try:
        data = request.json
        evaluation = data.get('evaluation', {})
        candidate_name = data.get('candidateName', '')
        answers = data.get('answers', [])
        
        if not ObjectId.is_valid(id):
            return jsonify({"error": "Invalid ID"}), 400
        
        # Récupérer la candidature
        candidature = mongo.db.candidatures.find_one({"_id": ObjectId(id)})
        if not candidature:
            return jsonify({"error": "Application not found"}), 404
        
        # Sauvegarder l'évaluation
        mongo.db.candidatures.update_one(
            {"_id": ObjectId(id)},
            {"$set": {
                "behavioral_evaluation": evaluation,
                "behavioral_evaluation_date": datetime.utcnow(),
                "behavioral_answers": answers
            }}
        )
        
        return jsonify({
            "message": "Évaluation enregistrée avec succès",
            "evaluation": evaluation
        }), 200
        
    except Exception as e:
        print(f"❌ Error saving behavioral evaluation: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/api/candidatures/<id>/behavioral-evaluation', methods=['GET'])
def get_behavioral_evaluation(id):
    """
    Récupère l'évaluation comportementale existante
    """
    try:
        if not ObjectId.is_valid(id):
            return jsonify({"error": "Invalid ID"}), 400
        
        candidature = mongo.db.candidatures.find_one({"_id": ObjectId(id)})
        if not candidature:
            return jsonify({"error": "Application not found"}), 404
        
        evaluation = candidature.get("behavioral_evaluation", {})
        
        return jsonify({
            "evaluation": evaluation,
            "has_evaluation": bool(evaluation)
        }), 200
        
    except Exception as e:
        print(f"❌ Error getting behavioral evaluation: {e}")
        return jsonify({"error": str(e)}), 500  











@app.route('/api/admin/reports/annual-summary', methods=['GET'])
def get_annual_summary():
    """
    Retourne un résumé annuel complet pour l'export PDF
    """
    try:
        year = request.args.get('year', type=int)
        if not year:
            year = datetime.utcnow().year
        
        # Date range for the year
        start_date = datetime(year, 1, 1)
        end_date = datetime(year, 12, 31, 23, 59, 59)
        
        # 1. Candidates stats for the year
        total_candidates = mongo.db.candidats.count_documents({
            "created_at": {"$gte": start_date, "$lte": end_date}
        })
        
        # 2. Applications stats
        total_applications = mongo.db.candidatures.count_documents({
            "date": {"$gte": start_date, "$lte": end_date}
        })
        
        # 3. Hired candidates by year
        hired_candidates = list(mongo.db.candidatures.aggregate([
            {"$match": {
                "final_decision": "hired",
                "final_decision_date": {"$gte": start_date, "$lte": end_date}
            }},
            {"$group": {
                "_id": "$company",
                "count": {"$sum": 1},
                "candidates": {"$push": {
                    "name": "$candidateName",
                    "email": "$candidateEmail",
                    "poste": "$poste",
                    "date": "$final_decision_date"
                }}
            }},
            {"$sort": {"count": -1}}
        ]))
        
        # 4. Jobs per company
        jobs_by_company = list(mongo.db.postes.aggregate([
            {"$match": {"created_at": {"$gte": start_date, "$lte": end_date}}},
            {"$group": {
                "_id": "$companyName",
                "jobs_count": {"$sum": 1},
                "jobs": {"$push": {
                    "title": "$title",
                    "type": "$type",
                    "created_at": "$created_at"
                }}
            }},
            {"$sort": {"jobs_count": -1}}
        ]))
        
        # 5. Monthly recruitment trend
        monthly_hired = list(mongo.db.candidatures.aggregate([
            {"$match": {
                "final_decision": "hired",
                "final_decision_date": {"$gte": start_date, "$lte": end_date}
            }},
            {"$group": {
                "_id": {"$month": "$final_decision_date"},
                "count": {"$sum": 1}
            }},
            {"$sort": {"_id": 1}}
        ]))
        
        # 6. Conversion rate by sector
        sector_conversion = list(mongo.db.candidatures.aggregate([
            {"$match": {"date": {"$gte": start_date, "$lte": end_date}}},
            {"$lookup": {
                "from": "postes",
                "localField": "posteId",
                "foreignField": "_id",
                "as": "job"
            }},
            {"$unwind": {"path": "$job", "preserveNullAndEmptyArrays": True}},
            {"$lookup": {
                "from": "entreprises",
                "localField": "job.companyId",
                "foreignField": "_id",
                "as": "company"
            }},
            {"$unwind": {"path": "$company", "preserveNullAndEmptyArrays": True}},
            {"$group": {
                "_id": "$company.secteur",
                "total": {"$sum": 1},
                "hired": {"$sum": {"$cond": [{"$eq": ["$final_decision", "hired"]}, 1, 0]}}
            }},
            {"$addFields": {
                "conversion_rate": {
                    "$cond": [
                        {"$eq": ["$total", 0]},
                        0,
                        {"$multiply": [{"$divide": ["$hired", "$total"]}, 100]}
                    ]
                }
            }}
        ]))
        
        # 7. Candidate origin distribution
        candidate_origins = list(mongo.db.candidats.aggregate([
            {"$match": {
                "created_at": {"$gte": start_date, "$lte": end_date},
                "nationalite": {"$nin": [None, "", "Not specified"]}
            }},
            {"$group": {
                "_id": "$nationalite",
                "count": {"$sum": 1}
            }},
            {"$sort": {"count": -1}}
        ]))
        
        # Retourner directement le dictionnaire (Flask le convertira automatiquement en JSON)
        return {
            "year": year,
            "total_candidates": total_candidates,
            "total_applications": total_applications,
            "total_hired": sum(h.get('count', 0) for h in hired_candidates),
            "conversion_rate": round((sum(h.get('count', 0) for h in hired_candidates) / total_applications * 100), 1) if total_applications > 0 else 0,
            "hired_by_company": hired_candidates,
            "jobs_by_company": jobs_by_company,
            "monthly_hired": monthly_hired,
            "sector_conversion": sector_conversion,
            "candidate_origins": candidate_origins
        }
        
    except Exception as e:
        print(f"❌ Error in annual_summary: {e}")
        return {"error": str(e)}, 500




@app.route('/api/admin/export-annual-pdf/<int:year>', methods=['GET'])
def export_annual_pdf(year):
    """
    Exporte un rapport annuel complet en PDF
    """
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        
        # Calculer les dates pour l'année
        start_date = datetime(year, 1, 1)
        end_date = datetime(year, 12, 31, 23, 59, 59)
        
        # ========== RÉCUPÉRER LES DONNÉES DIRECTEMENT ==========
        
        # 1. Candidates stats for the year
        total_candidates = mongo.db.candidats.count_documents({
            "created_at": {"$gte": start_date, "$lte": end_date}
        })
        
        # 2. Applications stats
        total_applications = mongo.db.candidatures.count_documents({
            "date": {"$gte": start_date, "$lte": end_date}
        })
        
        # 3. Hired candidates by year
        hired_candidates = list(mongo.db.candidatures.aggregate([
            {"$match": {
                "final_decision": "hired",
                "final_decision_date": {"$gte": start_date, "$lte": end_date}
            }},
            {"$group": {
                "_id": "$company",
                "count": {"$sum": 1},
                "candidates": {"$push": {
                    "name": "$candidateName",
                    "email": "$candidateEmail",
                    "poste": "$poste",
                    "date": "$final_decision_date"
                }}
            }},
            {"$sort": {"count": -1}}
        ]))
        
        total_hired = sum(h.get('count', 0) for h in hired_candidates)
        conversion_rate = round((total_hired / total_applications * 100), 1) if total_applications > 0 else 0
        
        # 4. Jobs per company
        jobs_by_company = list(mongo.db.postes.aggregate([
            {"$match": {"created_at": {"$gte": start_date, "$lte": end_date}}},
            {"$group": {
                "_id": "$companyName",
                "jobs_count": {"$sum": 1},
                "jobs": {"$push": {
                    "title": "$title",
                    "type": "$type",
                    "created_at": "$created_at"
                }}
            }},
            {"$sort": {"jobs_count": -1}}
        ]))
        
        # 5. Monthly recruitment trend
        monthly_hired = list(mongo.db.candidatures.aggregate([
            {"$match": {
                "final_decision": "hired",
                "final_decision_date": {"$gte": start_date, "$lte": end_date}
            }},
            {"$group": {
                "_id": {"$month": "$final_decision_date"},
                "count": {"$sum": 1}
            }},
            {"$sort": {"_id": 1}}
        ]))
        
        # 6. Conversion rate by sector
        sector_conversion = list(mongo.db.candidatures.aggregate([
            {"$match": {"date": {"$gte": start_date, "$lte": end_date}}},
            {"$lookup": {
                "from": "postes",
                "localField": "posteId",
                "foreignField": "_id",
                "as": "job"
            }},
            {"$unwind": {"path": "$job", "preserveNullAndEmptyArrays": True}},
            {"$lookup": {
                "from": "entreprises",
                "localField": "job.companyId",
                "foreignField": "_id",
                "as": "company"
            }},
            {"$unwind": {"path": "$company", "preserveNullAndEmptyArrays": True}},
            {"$group": {
                "_id": "$company.secteur",
                "total": {"$sum": 1},
                "hired": {"$sum": {"$cond": [{"$eq": ["$final_decision", "hired"]}, 1, 0]}}
            }},
            {"$addFields": {
                "conversion_rate": {
                    "$cond": [
                        {"$eq": ["$total", 0]},
                        0,
                        {"$multiply": [{"$divide": ["$hired", "$total"]}, 100]}
                    ]
                }
            }}
        ]))
        
        # 7. Candidate origin distribution
        candidate_origins = list(mongo.db.candidats.aggregate([
            {"$match": {
                "created_at": {"$gte": start_date, "$lte": end_date},
                "nationalite": {"$nin": [None, "", "Not specified"]}
            }},
            {"$group": {
                "_id": "$nationalite",
                "count": {"$sum": 1}
            }},
            {"$sort": {"count": -1}}
        ]))
        
        # 8. Company performance
        company_performance = []
        companies = list(mongo.db.entreprises.find())
        for company in companies:
            apps = list(mongo.db.candidatures.aggregate([
                {"$match": {
                    "companyId": str(company['_id']),
                    "date": {"$gte": start_date, "$lte": end_date}
                }},
                {"$group": {
                    "_id": None,
                    "total": {"$sum": 1},
                    "hired": {"$sum": {"$cond": [{"$eq": ["$final_decision", "hired"]}, 1, 0]}}
                }}
            ]))
            if apps and apps[0]['total'] > 0:
                company_performance.append({
                    "name": company.get('name', 'Unknown'),
                    "sector": company.get('secteur') or company.get('sector', 'Non spécifié'),
                    "total": apps[0]['total'],
                    "hired": apps[0]['hired']
                })
        company_performance.sort(key=lambda x: x['hired'], reverse=True)
        
        # Créer le PDF
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer, 
            pagesize=A4,
            leftMargin=1.5*cm,
            rightMargin=1.5*cm,
            topMargin=1.5*cm,
            bottomMargin=1.5*cm
        )
        
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=24,
            textColor=colors.HexColor('#0f172a'),
            alignment=1,
            spaceAfter=20
        )
        
        subtitle_style = ParagraphStyle(
            'Subtitle',
            parent=styles['Normal'],
            fontSize=12,
            textColor=colors.HexColor('#64748b'),
            alignment=1,
            spaceAfter=30
        )
        
        section_style = ParagraphStyle(
            'Section',
            parent=styles['Heading2'],
            fontSize=16,
            textColor=colors.HexColor('#00a6a6'),
            spaceAfter=12,
            spaceBefore=20
        )
        
        story = []
        
        # En-tête
        story.append(Paragraph(f"<b>LAURA CONNECTING AGENCY</b>", title_style))
        story.append(Paragraph(f"Rapport Annuel {year}", subtitle_style))
        story.append(Spacer(1, 0.5*cm))
        
        # ========== STATISTIQUES GLOBALES ==========
        story.append(Paragraph("📊 STATISTIQUES GLOBALES", section_style))
        
        stats_data = [
            ["Total candidats", str(total_candidates)],
            ["Total candidatures", str(total_applications)],
            ["Candidats retenus", str(total_hired)],
            ["Taux de conversion", f"{conversion_rate}%"]
        ]
        
        stats_table = Table(stats_data, colWidths=[8*cm, 4*cm])
        stats_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#f8fafc')),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))
        story.append(stats_table)
        story.append(Spacer(1, 0.5*cm))
        
        # ========== ÉVOLUTION MENSUELLE ==========
        story.append(Paragraph("📈 ÉVOLUTION MENSUELLE DES RECRUTEMENTS", section_style))
        
        months = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']
        monthly_data = [0] * 12
        for item in monthly_hired:
            month_idx = item['_id'] - 1
            if 0 <= month_idx < 12:
                monthly_data[month_idx] = item['count']
        
        monthly_table_data = [["Mois", "Recrutements"]]
        for i, count in enumerate(monthly_data):
            if count > 0:
                monthly_table_data.append([months[i], str(count)])
        
        if len(monthly_table_data) > 1:
            monthly_table = Table(monthly_table_data, colWidths=[6*cm, 6*cm])
            monthly_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#00a6a6')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
                ('ALIGN', (1, 0), (1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]))
            story.append(monthly_table)
        else:
            story.append(Paragraph("<i>Aucun recrutement enregistré cette année</i>", styles['Normal']))
        
        story.append(Spacer(1, 0.5*cm))
        
        # ========== CANDIDATS RETENUS PAR ENTREPRISE ==========
        story.append(Paragraph("🏢 CANDIDATS RETENUS PAR ENTREPRISE", section_style))
        
        if hired_candidates:
            company_data = [["Entreprise", "Candidats retenus"]]
            for company in hired_candidates[:15]:
                company_name = company['_id'] or "Non spécifié"
                company_data.append([company_name, str(company['count'])])
            
            company_table = Table(company_data, colWidths=[10*cm, 4*cm])
            company_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0f172a')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]))
            story.append(company_table)
        else:
            story.append(Paragraph("<i>Aucune donnée disponible</i>", styles['Normal']))
        
        story.append(Spacer(1, 0.5*cm))
        
        # ========== PERFORMANCE PAR SECTEUR ==========
        story.append(Paragraph("📊 CONVERSION PAR SECTEUR", section_style))
        
        if sector_conversion:
            sector_data = [["Secteur", "Candidatures", "Retenus", "Taux"]]
            for sector in sector_conversion:
                secteur = sector['_id'] or "Non spécifié"
                sector_data.append([
                    secteur, 
                    str(sector['total']), 
                    str(sector['hired']), 
                    f"{round(sector['conversion_rate'], 1)}%"
                ])
            
            sector_table = Table(sector_data, colWidths=[5*cm, 3*cm, 3*cm, 3*cm])
            sector_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0f172a')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
                ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]))
            story.append(sector_table)
        
        story.append(PageBreak())
        
        # ========== CLASSEMENT DES ENTREPRISES ==========
        story.append(Paragraph("🏆 CLASSEMENT DES ENTREPRISES", section_style))
        
        if company_performance:
            ranking_data = [["Rang", "Entreprise", "Secteur", "Candidatures", "Retenus"]]
            for idx, company in enumerate(company_performance[:20], 1):
                ranking_data.append([
                    str(idx),
                    company['name'][:30],
                    company['sector'][:20],
                    str(company['total']),
                    str(company['hired'])
                ])
            
            ranking_table = Table(ranking_data, colWidths=[1.5*cm, 6*cm, 4*cm, 2.5*cm, 2.5*cm])
            ranking_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#00a6a6')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
                ('ALIGN', (0, 0), (0, -1), 'CENTER'),
                ('ALIGN', (3, 0), (4, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ]))
            story.append(ranking_table)
        
        story.append(Spacer(1, 0.5*cm))
        
        # ========== ORIGINE DES CANDIDATS ==========
        story.append(Paragraph("🌍 ORIGINE DES CANDIDATS", section_style))
        
        if candidate_origins:
            total_origins = sum(o['count'] for o in candidate_origins)
            origin_data = [["Nationalité", "Nombre", "Pourcentage"]]
            for origin in candidate_origins[:10]:
                percentage = round(origin['count'] / total_origins * 100, 1) if total_origins > 0 else 0
                origin_data.append([origin['_id'], str(origin['count']), f"{percentage}%"])
            
            origin_table = Table(origin_data, colWidths=[6*cm, 4*cm, 4*cm])
            origin_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0f172a')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#e2e8f0')),
                ('ALIGN', (1, 0), (2, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]))
            story.append(origin_table)
        
        # Pied de page
        story.append(Spacer(1, 1*cm))
        footer_style = ParagraphStyle(
            'Footer',
            parent=styles['Normal'],
            fontSize=8,
            textColor=colors.HexColor('#94a3b8'),
            alignment=1
        )
        story.append(Paragraph(f"Document généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')} - Laura Connecting Agency", footer_style))
        
        # Générer le PDF
        doc.build(story)
        buffer.seek(0)
        
        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"rapport_annuel_lca_{year}.pdf",
            mimetype='application/pdf'
        )
        
    except Exception as e:
        print(f"❌ Error generating annual PDF: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500
    



@app.route('/api/admin/reports/available-years', methods=['GET'])
def get_available_years():
    """
    Retourne les années disponibles pour les rapports
    """
    try:
        years = set()
        
        # Années des candidatures
        for cand in mongo.db.candidatures.find({"date": {"$exists": True}}, {"date": 1}):
            if cand.get('date'):
                try:
                    years.add(cand['date'].year)
                except:
                    pass
        
        # Années des décisions finales
        for cand in mongo.db.candidatures.find({"final_decision_date": {"$exists": True}}, {"final_decision_date": 1}):
            if cand.get('final_decision_date'):
                try:
                    years.add(cand['final_decision_date'].year)
                except:
                    pass
        
        # Si aucune année trouvée, ajouter l'année courante
        if not years:
            years.add(datetime.now().year)
        
        return jsonify(sorted(list(years), reverse=True)), 200
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return jsonify([datetime.now().year]), 200


@app.route('/api/admin/reports/company-performance/<year>', methods=['GET'])
def get_company_performance(year):
    """
    Retourne les performances par entreprise pour une année donnée
    """
    try:
        year = int(year)
        start_date = datetime(year, 1, 1)
        end_date = datetime(year, 12, 31, 23, 59, 59)
        
        companies = list(mongo.db.entreprises.find())
        
        result = []
        for company in companies:
            # Candidatures pour cette entreprise
            applications = list(mongo.db.candidatures.aggregate([
                {"$match": {
                    "companyId": str(company['_id']),
                    "date": {"$gte": start_date, "$lte": end_date}
                }},
                {"$group": {
                    "_id": None,
                    "total": {"$sum": 1},
                    "hired": {"$sum": {"$cond": [{"$eq": ["$final_decision", "hired"]}, 1, 0]}},
                    "rejected": {"$sum": {"$cond": [{"$eq": ["$final_decision", "rejected"]}, 1, 0]}}
                }}
            ]))
            
            stats = applications[0] if applications else {"total": 0, "hired": 0, "rejected": 0}
            
            # Offres actives
            active_jobs = mongo.db.postes.count_documents({
                "companyId": str(company['_id']),
                "status": "active"
            })
            
            conversion = round(stats['hired'] / stats['total'] * 100, 1) if stats['total'] > 0 else 0
            
            result.append({
                "company_id": str(company['_id']),
                "company_name": company.get('name', 'Unknown'),
                "sector": company.get('secteur') or company.get('sector', 'Non spécifié'),
                "total_applications": stats['total'],
                "hired": stats['hired'],
                "rejected": stats['rejected'],
                "conversion_rate": conversion,
                "active_jobs": active_jobs
            })
        
        # Trier par nombre d'embauches
        result.sort(key=lambda x: x['hired'], reverse=True)
        
        return jsonify(result), 200
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return jsonify([]), 500
    
















# ========== À AJOUTER DANS app.py ==========

# ========== NOUVELLES ROUTES QUESTIONNAIRE ADMIN ==========

@app.route('/api/admin/questionnaires', methods=['GET'])
@admin_required()
def get_all_questionnaires():
    """Récupère tous les questionnaires envoyés"""
    try:
        questionnaires = list(mongo.db.questionnaires.find().sort("created_at", -1))
        for q in questionnaires:
            q['_id'] = str(q['_id'])
            q['created_at'] = q['created_at'].isoformat() if q.get('created_at') else None
            q['answered_at'] = q['answered_at'].isoformat() if q.get('answered_at') else None
            # Ajouter infos candidat
            candidate = mongo.db.candidats.find_one({"email": q['candidate_email']})
            if candidate:
                q['candidate_name'] = f"{candidate.get('first_name', '')} {candidate.get('last_name', '')}".strip()
        return jsonify(questionnaires), 200
    except Exception as e:
        print(f"❌ Error: {e}")
        return jsonify([]), 200


@app.route('/api/admin/questionnaires/create', methods=['POST'])
@admin_required()
def create_questionnaire():
    """Crée un nouveau questionnaire et l'envoie au candidat"""
    try:
        data = request.json
        candidate_email = data.get('candidate_email')
        title = data.get('title', 'Questionnaire d\'évaluation')
        questions = data.get('questions', [])
        due_date = data.get('due_date')
        
        if not candidate_email or not questions:
            return jsonify({"error": "Email candidat et questions requis"}), 400
        
        # Vérifier que le candidat existe
        candidate = mongo.db.candidats.find_one({"email": candidate_email})
        if not candidate:
            return jsonify({"error": "Candidat non trouvé"}), 404
        
        # Générer un token unique
        token = hashlib.md5(f"{candidate_email}{datetime.utcnow()}{uuid.uuid4()}".encode()).hexdigest()[:32]
        
        # Créer le document questionnaire
        questionnaire = {
            "token": token,
            "candidate_email": candidate_email,
            "candidate_name": f"{candidate.get('first_name', '')} {candidate.get('last_name', '')}".strip(),
            "title": title,
            "questions": questions,
            "status": "pending",
            "created_at": datetime.utcnow(),
            "due_date": due_date,
            "answers": [],
            "answered_at": None
        }
        
        result = mongo.db.questionnaires.insert_one(questionnaire)
        
        # Mettre à jour la candidature si un job est associé (optionnel)
        if data.get('candidature_id'):
            mongo.db.candidatures.update_one(
                {"_id": ObjectId(data['candidature_id'])},
                {"$set": {"questionnaire_token": token, "questionnaire_status": "pending"}}
            )
        
        # Envoyer l'email au candidat
        candidate_name = f"{candidate.get('first_name', '')} {candidate.get('last_name', '')}".strip() or "Candidat"
        link = f"http://localhost:4200/questionnaire/{token}"
        
        # Construction des questions pour l'email
        questions_html = ""
        for i, q in enumerate(questions, 1):
            questions_html += f"<li><strong>{q.get('text', 'Question')}</strong> ({q.get('type', 'text')})</li>"
        
        subject = f"📋 {title} - LCA"
        body = f"""
        <html>
        <body style="font-family: Arial, sans-serif;">
            <div style="background: #0f172a; padding: 20px; text-align: center;">
                <h2 style="color: white; margin: 0;">{title}</h2>
            </div>
            <div style="padding: 20px;">
                <p>Bonjour <strong>{candidate_name}</strong>,</p>
                <p>Vous avez reçu un questionnaire d'évaluation de la part de Laura Connecting Agency.</p>
                
                <div style="background: #f0fdfa; padding: 15px; border-radius: 10px; margin: 15px 0;">
                    <p><strong>📝 Questions ({len(questions)}):</strong></p>
                    <ul>{questions_html}</ul>
                    {f'<p><strong>📅 Date limite:</strong> {due_date}</p>' if due_date else ''}
                </div>
                
                <p style="margin: 30px 0;">
                    <a href="{link}" style="background-color: #00a6a6; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px;">
                        ➡️ Répondre au questionnaire
                    </a>
                </p>
                
                <p>Cordialement,<br>L'équipe LCA</p>
            </div>
        </body>
        </html>
        """
        
        send_real_email(candidate_email, subject, body)
        
        # Notification pour le candidat
        mongo.db.notifications.insert_one({
            "recipient": "candidate",
            "recipientEmail": candidate_email,
            "type": "questionnaire_received",
            "title": f"📋 {title}",
            "message": f"Vous avez reçu un questionnaire d'évaluation ({len(questions)} questions).",
            "status": "unread",
            "date": datetime.utcnow(),
            "link": link
        })
        
        return jsonify({
            "message": "Questionnaire créé et envoyé avec succès",
            "questionnaire_id": str(result.inserted_id),
            "token": token
        }), 201
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route('/api/admin/questionnaires/<questionnaire_id>/answers', methods=['GET'])
@admin_required()
def get_questionnaire_answers(questionnaire_id):
    """Récupère les réponses d'un questionnaire"""
    try:
        if not ObjectId.is_valid(questionnaire_id):
            return jsonify({"error": "ID invalide"}), 400
        
        questionnaire = mongo.db.questionnaires.find_one({"_id": ObjectId(questionnaire_id)})
        if not questionnaire:
            return jsonify({"error": "Questionnaire non trouvé"}), 404
        
        return jsonify({
            "id": str(questionnaire['_id']),
            "title": questionnaire.get('title'),
            "candidate_email": questionnaire.get('candidate_email'),
            "candidate_name": questionnaire.get('candidate_name'),
            "questions": questionnaire.get('questions', []),
            "answers": questionnaire.get('answers', []),
            "status": questionnaire.get('status'),
            "created_at": questionnaire.get('created_at').isoformat() if questionnaire.get('created_at') else None,
            "answered_at": questionnaire.get('answered_at').isoformat() if questionnaire.get('answered_at') else None
        }), 200
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/admin/questionnaires/<questionnaire_id>/reminder', methods=['POST'])
@admin_required()
def send_questionnaire_reminder(questionnaire_id):
    """Envoie un rappel au candidat pour remplir le questionnaire"""
    try:
        if not ObjectId.is_valid(questionnaire_id):
            return jsonify({"error": "ID invalide"}), 400
        
        questionnaire = mongo.db.questionnaires.find_one({"_id": ObjectId(questionnaire_id)})
        if not questionnaire:
            return jsonify({"error": "Questionnaire non trouvé"}), 404
        
        if questionnaire.get('status') == 'answered':
            return jsonify({"error": "Le questionnaire a déjà été répondu"}), 400
        
        candidate = mongo.db.candidats.find_one({"email": questionnaire['candidate_email']})
        candidate_name = f"{candidate.get('first_name', '')} {candidate.get('last_name', '')}".strip() or "Candidat"
        link = f"http://localhost:4200/questionnaire/{questionnaire['token']}"
        
        subject = f"📋 Rappel : {questionnaire.get('title', 'Questionnaire')} - LCA"
        body = f"""
        <html>
        <body style="font-family: Arial, sans-serif;">
            <div style="background: #f59e0b; padding: 20px; text-align: center;">
                <h2 style="color: white; margin: 0;">⏰ Rappel - Questionnaire en attente</h2>
            </div>
            <div style="padding: 20px;">
                <p>Bonjour <strong>{candidate_name}</strong>,</p>
                <p>Nous vous rappelons que vous devez encore remplir le questionnaire <strong>{questionnaire.get('title')}</strong>.</p>
                
                <p style="margin: 30px 0;">
                    <a href="{link}" style="background-color: #f59e0b; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px;">
                        ➡️ Répondre au questionnaire
                    </a>
                </p>
                
                <p>Cordialement,<br>L'équipe LCA</p>
            </div>
        </body>
        </html>
        """
        
        send_real_email(questionnaire['candidate_email'], subject, body)
        
        return jsonify({"message": "Rappel envoyé avec succès"}), 200
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/api/admin/questionnaires/delete/<questionnaire_id>', methods=['DELETE'])
@admin_required()
def delete_questionnaire(questionnaire_id):
    """Supprime un questionnaire"""
    try:
        if not ObjectId.is_valid(questionnaire_id):
            return jsonify({"error": "ID invalide"}), 400
        
        result = mongo.db.questionnaires.delete_one({"_id": ObjectId(questionnaire_id)})
        
        if result.deleted_count:
            return jsonify({"message": "Questionnaire supprimé"}), 200
        else:
            return jsonify({"error": "Questionnaire non trouvé"}), 404
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return jsonify({"error": str(e)}), 500


# ========== MODIFICATION DE LA ROUTE CANDIDAT QUESTIONNAIRE ==========



@app.route('/api/candidate/questionnaire/submit-v2', methods=['POST'])
def submit_questionnaire_v2():
    """Soumet les réponses du questionnaire (version améliorée)"""
    try:
        data = request.json
        questionnaire_id = data.get('questionnaireId')
        token = data.get('token')
        answers = data.get('answers', [])
        
        questionnaire = None
        
        # Chercher par ID ou par token
        if questionnaire_id and ObjectId.is_valid(questionnaire_id):
            questionnaire = mongo.db.questionnaires.find_one({"_id": ObjectId(questionnaire_id)})
        elif token:
            questionnaire = mongo.db.questionnaires.find_one({"token": token})
        
        if not questionnaire:
            # Fallback: chercher dans candidatures
            if token:
                candidature = mongo.db.candidatures.find_one({"questionnaire_token": token})
                if candidature:
                    formatted_answers = []
                    for a in answers:
                        if isinstance(a, dict):
                            formatted_answers.append({
                                "question": a.get('question') or a.get('questionText', ''),
                                "answer": a.get('answer', '')
                            })
                    
                    mongo.db.candidatures.update_one(
                        {"_id": candidature["_id"]},
                        {"$set": {
                            "answers": formatted_answers,
                            "questionnaire_status": "answered",
                            "questionnaire_answered_at": datetime.utcnow()
                        }}
                    )
                    
                    # Notification admin
                    mongo.db.notifications.insert_one({
                        "recipient": "admin",
                        "recipientEmail": "admin@lca.com",
                        "type": "questionnaire_answered",
                        "title": "📝 Questionnaire rempli",
                        "message": f"{candidature.get('candidateName')} a répondu au questionnaire",
                        "status": "unread",
                        "date": datetime.utcnow()
                    })
                    
                    return jsonify({"message": "Réponses enregistrées", "answered_count": len(formatted_answers)}), 200
            
            return jsonify({"error": "Questionnaire non trouvé"}), 404
        
        # Vérifier si déjà répondu
        if questionnaire.get('status') == 'answered':
            return jsonify({"error": "Ce questionnaire a déjà été répondu"}), 400
        
        # Formater les réponses
        formatted_answers = []
        for i, a in enumerate(answers):
            if isinstance(a, dict):
                formatted_answers.append({
                    "question_id": a.get('questionId', i),
                    "question_text": a.get('questionText') or a.get('question', ''),
                    "answer": a.get('answer', '')
                })
        
        # Mettre à jour le questionnaire
        mongo.db.questionnaires.update_one(
            {"_id": questionnaire["_id"]},
            {"$set": {
                "answers": formatted_answers,
                "status": "answered",
                "answered_at": datetime.utcnow()
            }}
        )
        
        # Notification pour l'admin
        mongo.db.notifications.insert_one({
            "recipient": "admin",
            "recipientEmail": "admin@lca.com",
            "type": "questionnaire_answered",
            "title": "📝 Questionnaire rempli",
            "message": f"{questionnaire.get('candidate_name')} a répondu au questionnaire '{questionnaire.get('title')}'",
            "status": "unread",
            "date": datetime.utcnow(),
            "questionnaire_id": str(questionnaire["_id"])
        })
        
        return jsonify({
            "message": "Réponses enregistrées avec succès",
            "answered_count": len(formatted_answers)
        }), 200
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500
    



@app.route('/api/debug/job-questions/<candidature_id>', methods=['GET'])
def debug_job_questions(candidature_id):
    """Debug pour voir le job associé à une candidature"""
    try:
        if not ObjectId.is_valid(candidature_id):
            return jsonify({"error": "Invalid ID"}), 400
        
        candidature = mongo.db.candidatures.find_one({"_id": ObjectId(candidature_id)})
        if not candidature:
            return jsonify({"error": "Candidature not found"}), 404
        
        job_id = candidature.get('posteId')
        if not job_id:
            return jsonify({"error": "No job associated with this candidature"}), 404
        
        if not ObjectId.is_valid(job_id):
            return jsonify({"error": "Invalid job ID"}), 400
        
        job = mongo.db.postes.find_one({"_id": ObjectId(job_id)})
        if not job:
            return jsonify({"error": "Job not found"}), 404
        
        question_sections = job.get('questionSections', [])
        
        return jsonify({
            "candidature_id": candidature_id,
            "candidate_email": candidature.get('candidateEmail'),
            "job_id": str(job['_id']),
            "job_title": job.get('title'),
            "has_question_sections": len(question_sections) > 0,
            "question_sections_count": len(question_sections),
            "total_questions": sum(len(s.get('questions', [])) for s in question_sections),
            "question_sections": question_sections
        }), 200
    except Exception as e:
        print(f"❌ Error: {e}")
        return jsonify({"error": str(e)}), 500














@app.route('/api/debug/candidature-questionnaire/<candidature_id>', methods=['GET'])
def debug_candidature_questionnaire(candidature_id):
    """Debug pour voir l'état du questionnaire d'une candidature"""
    if not ObjectId.is_valid(candidature_id):
        return jsonify({"error": "Invalid ID"}), 400
    
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(candidature_id)})
    if not candidature:
        return jsonify({"error": "Not found"}), 404
    
    return jsonify({
        "candidature_id": str(candidature["_id"]),
        "candidateName": candidature.get("candidateName"),
        "candidateEmail": candidature.get("candidateEmail"),
        "has_custom_questions": len(candidature.get("custom_questions", [])) > 0,
        "custom_questions_count": len(candidature.get("custom_questions", [])),
        "has_answers": len(candidature.get("answers", [])) > 0,
        "answers_count": len(candidature.get("answers", [])),
        "questionnaire_token": candidature.get("questionnaire_token"),
        "questionnaire_status": candidature.get("questionnaire_status", "pending"),
        "questionnaire_sent_at": candidature.get("questionnaire_sent_at"),
        "questionnaire_answered_at": candidature.get("questionnaire_answered_at")
    }), 200





@app.route('/api/admin/force-send-questionnaire/<candidature_id>', methods=['POST'])
def force_send_questionnaire(candidature_id):
    """Force l'envoi d'un questionnaire au candidat"""
    if not ObjectId.is_valid(candidature_id):
        return jsonify({"error": "Invalid ID"}), 400
    
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(candidature_id)})
    if not candidature:
        return jsonify({"error": "Not found"}), 404
    
    # Récupérer le job associé
    job = None
    job_id = candidature.get('posteId')
    if job_id and ObjectId.is_valid(job_id):
        job = mongo.db.postes.find_one({"_id": ObjectId(job_id)})
    
    # Créer des questions par défaut si nécessaire
    questions = []
    if job and job.get('questionSections'):
        for section in job['questionSections']:
            section_title = section.get('titleFr', section.get('title', 'Question'))
            for q in section.get('questions', []):
                questions.append({
                    "text": f"[{section_title}] {q.get('textFr', q.get('text', 'Question'))}",
                    "type": q.get('type', 'text')
                })
    else:
        # Questions par défaut
        questions = [
            {"text": "Nom complet", "type": "text"},
            {"text": "Email", "type": "text"},
            {"text": "Téléphone", "type": "text"},
            {"text": "Ville de résidence", "type": "text"},
            {"text": "Années d'expérience", "type": "text"},
            {"text": "Dernier poste occupé", "type": "text"},
            {"text": "Langues parlées (avec niveau)", "type": "textarea"}
        ]
    
    # Générer un token
    token = hashlib.md5(f"{candidature_id}{datetime.utcnow()}".encode()).hexdigest()[:16]
    
    # Mettre à jour la candidature
    mongo.db.candidatures.update_one(
        {"_id": ObjectId(candidature_id)},
        {"$set": {
            "custom_questions": questions,
            "questionnaire_status": "pending",
            "answers": [],
            "questionnaire_token": token,
            "questionnaire_sent_at": datetime.utcnow()
        }}
    )
    
    # Envoyer l'email
    candidate_email = candidature.get('candidateEmail')
    candidate_name = candidature.get('candidateName', 'Candidat')
    link = f"http://localhost:4200/questionnaire/{token}"
    
    subject = "📋 Questionnaire d'évaluation"
    body = f"""
    <html>
    <body>
        <h2>Questionnaire d'évaluation</h2>
        <p>Bonjour {candidate_name},</p>
        <p>Merci de remplir ce questionnaire :</p>
        <a href="{link}">Cliquez ici</a>
    </body>
    </html>
    """
    
    send_real_email(candidate_email, subject, body)
    
    return jsonify({
        "message": "Questionnaire envoyé",
        "token": token,
        "questions_count": len(questions)
    }), 200





@app.route('/api/admin/reset-questionnaire/<candidature_id>', methods=['POST'])
def reset_questionnaire(candidature_id):
    """Réinitialise le questionnaire d'un candidat"""
    if not ObjectId.is_valid(candidature_id):
        return jsonify({"error": "Invalid ID"}), 400
    
    result = mongo.db.candidatures.update_one(
        {"_id": ObjectId(candidature_id)},
        {"$set": {
            "questionnaire_status": "pending",
            "answers": []
        }}
    )
    
    return jsonify({"message": "Questionnaire réinitialisé"}), 200








# ========== ROUTE AMÉLIORÉE POUR ENVOYER LE QUESTIONNAIRE DÉTAILLÉ ==========
@app.route('/api/admin/candidatures/<id>/send-detailed-questionnaire', methods=['POST', 'OPTIONS'])
def send_detailed_questionnaire(id):
    """Envoie un questionnaire détaillé au candidat"""
    
    if request.method == 'OPTIONS':
        response = jsonify({'message': 'OK'})
        response.headers.add('Access-Control-Allow-Origin', '*')
        response.headers.add('Access-Control-Allow-Headers', 'Content-Type, Authorization')
        response.headers.add('Access-Control-Allow-Methods', 'POST, OPTIONS')
        return response
    
    if not ObjectId.is_valid(id):
        return jsonify({"error": "ID invalide"}), 400
    
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(id)})
    if not candidature:
        return jsonify({"error": "Candidature non trouvée"}), 404
    
    data = request.json
    questions_data = data.get('questions', [])
    
    if not questions_data:
        return jsonify({"error": "Aucune question fournie"}), 400
    
    # Formater les questions pour MongoDB (format simple)
    formatted_questions = []
    for q in questions_data:
        if isinstance(q, dict):
            formatted_questions.append({
                "text": q.get('text', 'Question'),
                "type": q.get('type', 'text')
            })
        else:
            formatted_questions.append({"text": q, "type": "text"})
    
    # Générer un token unique
    token = hashlib.md5(f"{candidature['_id']}{datetime.utcnow()}{uuid.uuid4()}".encode()).hexdigest()[:16]
    
    # Mettre à jour la candidature
    mongo.db.candidatures.update_one(
        {"_id": ObjectId(id)},
        {"$set": {
            "custom_questions": formatted_questions,
            "questionnaire_status": "pending",
            "answers": [],
            "questionnaire_token": token,
            "questionnaire_sent_at": datetime.utcnow()
        }}
    )
    
    # Envoyer l'email au candidat
    candidate_email = candidature.get('candidateEmail')
    candidate_name = candidature.get('candidateName', 'Candidat')
    link = f"http://localhost:4200/questionnaire/{token}"
    
    subject = "📋 Questionnaire d'évaluation"
    body = f"""
    <html>
    <body style="font-family: Arial, sans-serif;">
        <div style="background: #0f172a; padding: 20px; text-align: center;">
            <h2 style="color: white; margin: 0;">📋 Questionnaire d'évaluation</h2>
        </div>
        <div style="padding: 20px;">
            <p>Bonjour <strong>{candidate_name}</strong>,</p>
            <p>Merci de bien vouloir remplir ce questionnaire d'évaluation :</p>
            <p style="margin: 30px 0;">
                <a href="{link}" style="background-color: #00a6a6; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px;">
                    ➡️ Accéder au questionnaire
                </a>
            </p>
            <p>Cordialement,<br>L'équipe LCA</p>
        </div>
    </body>
    </html>
    """
    
    send_real_email(candidate_email, subject, body)
    
    # Notification pour le candidat
    mongo.db.notifications.insert_one({
        "recipient": "candidate",
        "recipientEmail": candidate_email,
        "type": "questionnaire_received",
        "title": "📋 Questionnaire d'évaluation",
        "message": f"Vous avez reçu un questionnaire de {len(formatted_questions)} questions.",
        "status": "unread",
        "date": datetime.utcnow(),
        "link": link
    })
    
    response = jsonify({
        "message": "Questionnaire envoyé au candidat",
        "count": len(formatted_questions),
        "token": token
    })
    response.headers.add('Access-Control-Allow-Origin', '*')
    return response, 200



@app.route('/api/debug/questionnaire-token/<token>', methods=['GET'])
def debug_questionnaire_token(token):
    """Debug pour voir si un token existe"""
    # Chercher dans candidatures
    candidature = mongo.db.candidatures.find_one({"questionnaire_token": token})
    if candidature:
        return jsonify({
            "exists": True,
            "collection": "candidatures",
            "candidature_id": str(candidature["_id"]),
            "candidate_email": candidature.get("candidateEmail"),
            "questions_count": len(candidature.get("custom_questions", []))
        }), 200
    
    # Chercher dans questionnaires
    questionnaire = mongo.db.questionnaires.find_one({"token": token})
    if questionnaire:
        return jsonify({
            "exists": True,
            "collection": "questionnaires",
            "questionnaire_id": str(questionnaire["_id"]),
            "candidate_email": questionnaire.get("candidate_email"),
            "questions_count": len(questionnaire.get("questions", []))
        }), 200
    
    return jsonify({"exists": False, "token": token}), 404



@app.route('/api/admin/candidatures/<id>/create-questionnaire', methods=['POST'])
def create_questionnaire_only(id):
    """Crée uniquement le questionnaire sans envoyer d'email (pour test)"""
    if not ObjectId.is_valid(id):
        return jsonify({"error": "ID invalide"}), 400
    
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(id)})
    if not candidature:
        return jsonify({"error": "Candidature non trouvée"}), 404
    
    data = request.json
    questions = data.get('questions', [])
    
    if not questions:
        return jsonify({"error": "Aucune question fournie"}), 400
    
    # Générer un token unique
    token = hashlib.md5(f"{candidature['_id']}{datetime.utcnow()}{uuid.uuid4()}".encode()).hexdigest()[:16]
    
    # Mettre à jour la candidature
    result = mongo.db.candidatures.update_one(
        {"_id": ObjectId(id)},
        {"$set": {
            "custom_questions": questions,
            "questionnaire_status": "pending",
            "answers": [],
            "questionnaire_token": token,
            "questionnaire_sent_at": datetime.utcnow()
        }}
    )
    
    if result.modified_count == 0:
        return jsonify({"error": "Échec de la mise à jour"}), 500
    
    # Retourner le token pour test
    return jsonify({
        "message": "Questionnaire créé avec succès",
        "token": token,
        "link": f"http://localhost:4200/questionnaire/{token}",
        "questions_count": len(questions)
    }), 200




@app.route('/api/debug/verify-token/<token>', methods=['GET'])
def debug_verify_token(token):
    """Vérifie la validité d'un token et retourne des infos de debug"""
    result = {
        "token": token,
        "valid": False,
        "details": {}
    }
    
    # Vérifier dans candidatures
    candidature = mongo.db.candidatures.find_one({"questionnaire_token": token})
    if candidature:
        result["valid"] = True
        result["details"]["collection"] = "candidatures"
        result["details"]["candidature_id"] = str(candidature["_id"])
        result["details"]["candidate_email"] = candidature.get("candidateEmail")
        result["details"]["candidate_name"] = candidature.get("candidateName")
        result["details"]["questions_count"] = len(candidature.get("custom_questions", []))
        result["details"]["status"] = candidature.get("questionnaire_status", "pending")
        result["details"]["sent_at"] = candidature.get("questionnaire_sent_at", "").isoformat() if candidature.get("questionnaire_sent_at") else None
    
    return jsonify(result), 200




@app.route('/api/admin/candidatures/<id>/send-custom-questionnaire', methods=['POST'])
def send_custom_questionnaire_final(id):
    """Envoie un questionnaire personnalisé - Version finale"""
    
    if request.method == 'OPTIONS':
        response = jsonify({'message': 'OK'})
        response.headers.add('Access-Control-Allow-Origin', '*')
        response.headers.add('Access-Control-Allow-Headers', 'Content-Type, Authorization')
        response.headers.add('Access-Control-Allow-Methods', 'POST, OPTIONS')
        return response
    
    if not ObjectId.is_valid(id):
        return jsonify({"error": "ID invalide"}), 400
    
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(id)})
    if not candidature:
        return jsonify({"error": "Candidature non trouvée"}), 404
    
    data = request.json
    sections = data.get('sections', [])
    
    if not sections:
        return jsonify({"error": "Aucune section fournie"}), 400
    
    # Compter et préparer les questions
    total_questions = 0
    questions_plat = []
    
    for section in sections:
        section_title_fr = section.get('titleFr', '')
        section_title_en = section.get('titleEn', '')
        
        for q in section.get('questions', []):
            total_questions += 1
            
            # Récupérer le texte de la question
            question_text = q.get('textFr', '')
            if not question_text:
                question_text = q.get('text', '')
            
            # Ajouter le titre de section si pertinent
            if section_title_fr and section_title_fr not in question_text:
                question_text = f"[{section_title_fr}] {question_text}"
            
            questions_plat.append({
                "text": question_text,
                "textEn": q.get('textEn', ''),
                "type": q.get('type', 'text'),
                "required": q.get('required', False),
                "options": q.get('options', [])
            })
    
    print(f"📋 Questionnaire préparé: {total_questions} questions")
    
    # Générer un token unique
    import hashlib
    token = hashlib.md5(f"{candidature['_id']}{datetime.utcnow()}{uuid.uuid4()}".encode()).hexdigest()[:16]
    
    # Supprimer l'ancien token s'il existe
    mongo.db.candidatures.update_one(
        {"_id": ObjectId(id)},
        {"$unset": {"questionnaire_token": "", "questionnaire_status": "", "custom_questions": ""}}
    )
    
    # Mettre à jour avec les nouvelles données
    result = mongo.db.candidatures.update_one(
        {"_id": ObjectId(id)},
        {"$set": {
            "custom_questions": questions_plat,
            "questionnaire_status": "pending",
            "answers": [],
            "questionnaire_token": token,
            "questionnaire_sent_at": datetime.utcnow()
        }}
    )
    
    if result.modified_count == 0:
        return jsonify({"error": "Impossible de mettre à jour la candidature"}), 500
    
    # ========== ENVOI D'EMAIL ==========
    candidate_email = candidature.get('candidateEmail')
    candidate_name = candidature.get('candidateName', 'Candidat')
    poste_title = candidature.get('poste', 'votre candidature')
    link = f"http://localhost:4200/questionnaire/{token}"
    
    subject = f"📋 Questionnaire personnalisé - Laura Connecting Agency"
    
    # Corps HTML de l'email
    body = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Questionnaire LCA</title>
        <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
            .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
            .header {{ background: linear-gradient(135deg, #0f172a, #1e293b); padding: 25px; text-align: center; border-radius: 15px 15px 0 0; }}
            .header h2 {{ color: white; margin: 0; font-size: 24px; }}
            .content {{ background: #ffffff; padding: 30px; border: 1px solid #e2e8f0; border-radius: 0 0 15px 15px; }}
            .button {{ display: inline-block; background: #00a6a6; color: white; padding: 14px 28px; text-decoration: none; border-radius: 8px; margin: 20px 0; font-weight: bold; }}
            .button:hover {{ background: #008a8a; }}
            .footer {{ text-align: center; padding: 20px; font-size: 12px; color: #94a3b8; }}
            .info-box {{ background: #f0fdfa; padding: 15px; border-radius: 10px; margin: 20px 0; border-left: 4px solid #00a6a6; }}
            hr {{ border: none; border-top: 1px solid #e2e8f0; margin: 20px 0; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h2>📋 Questionnaire d'évaluation</h2>
            </div>
            <div class="content">
                <p>Bonjour <strong>{candidate_name}</strong>,</p>
                <p>Merci de votre intérêt pour le poste de <strong>{poste_title}</strong>.</p>
                
                <div class="info-box">
                    <p><strong>📝 Détails :</strong></p>
                    <ul>
                        <li><strong>{total_questions}</strong> questions au total</li>
                        <li>Réponse obligatoire pour toutes les questions</li>
                        <li>Temps estimé : environ 5-10 minutes</li>
                    </ul>
                </div>
                
                <div style="text-align: center;">
                    <a href="{link}" class="button">➡️ Accéder au questionnaire</a>
                </div>
                
                <hr>
                
                <p><strong>ℹ️ Informations importantes :</strong><br>
                - Ce lien est personnel et valable <strong>7 jours</strong>.<br>
                - Vous pouvez sauvegarder vos réponses et reprendre plus tard.<br>
                - En cas de problème, contactez-nous à <a href="mailto:info@lcateam.com">info@lcateam.com</a>
                </p>
                
                <p>Cordialement,<br>
                <strong>L'équipe Laura Connecting Agency</strong></p>
            </div>
            <div class="footer">
                © 2025 Laura Connecting Agency - Tous droits réservés
            </div>
        </div>
    </body>
    </html>
    """
    
    # Envoi de l'email
    email_sent = send_real_email(candidate_email, subject, body)
    
    # Notification pour le candidat
    mongo.db.notifications.insert_one({
        "recipient": "candidate",
        "recipientEmail": candidate_email,
        "type": "questionnaire_received",
        "title": "📋 Questionnaire disponible",
        "message": f"Un questionnaire d'évaluation ({total_questions} questions) vous attend pour le poste de {poste_title}.",
        "status": "unread",
        "date": datetime.utcnow(),
        "link": link
    })
    
    return jsonify({
        "message": "Questionnaire envoyé avec succès",
        "count": total_questions,
        "token": token,
        "email_sent": email_sent,
        "link": link
    }), 200 








@app.route('/api/admin/fix-questionnaire/<candidature_id>', methods=['POST'])
def fix_questionnaire(candidature_id):
    """Recrée le questionnaire pour une candidature"""
    if not ObjectId.is_valid(candidature_id):
        return jsonify({"error": "Invalid ID"}), 400
    
    candidature = mongo.db.candidatures.find_one({"_id": ObjectId(candidature_id)})
    if not candidature:
        return jsonify({"error": "Candidature not found"}), 404
    
    # Questions par défaut
    default_questions = [
        {"text": "Nom complet", "type": "text"},
        {"text": "Genre (Homme/Femme)", "type": "choice"},
        {"text": "Téléphone", "type": "text"},
        {"text": "Email", "type": "email"},
        {"text": "Nationalité", "type": "text"},
        {"text": "Années d'expérience / Secteur", "type": "text"},
        {"text": "Type de contrat actuel", "type": "choice"},
        {"text": "Délai de préavis", "type": "text"},
        {"text": "Salaire actuel", "type": "text"},
        {"text": "Salaire demandé", "type": "text"},
        {"text": "Entreprise actuelle", "type": "text"},
        {"text": "Adresse actuelle", "type": "text"},
        {"text": "Diplôme", "type": "text"},
        {"text": "Raison de changement de poste", "type": "textarea"},
        {"text": "Certifications", "type": "textarea"},
        {"text": "Parcours professionnel", "type": "textarea"}
    ]
    
    # Régénérer le token
    import hashlib
    token = hashlib.md5(f"{candidature['_id']}{datetime.utcnow()}{uuid.uuid4()}".encode()).hexdigest()[:16]
    
    # Mettre à jour
    mongo.db.candidatures.update_one(
        {"_id": ObjectId(candidature_id)},
        {"$set": {
            "custom_questions": default_questions,
            "questionnaire_status": "pending",
            "answers": [],
            "questionnaire_token": token,
            "questionnaire_sent_at": datetime.utcnow()
        }}
    )
    
    return jsonify({
        "message": "Questionnaire recréé avec succès",
        "token": token,
        "questions_count": len(default_questions),
        "link": f"http://localhost:4200/questionnaire/{token}"
    }), 200











#=======================================================================
if __name__ == '__main__':
    socketio.run(app, host='0.0.0.0', debug=True, port=5000)   















    
