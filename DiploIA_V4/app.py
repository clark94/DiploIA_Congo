import re
import json
import base64
import os
from io import BytesIO
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from reportlab.lib import colors as rl_colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    from docx import Document as DocxDocument
except ImportError:
    DocxDocument = None

try:
    from openai import OpenAI
    HAS_OPENAI = True
except ImportError:
    HAS_OPENAI = False

# ── CONSTANTES ────────────────────────────────────────────────────────
MINISTERE = "Ministère des Affaires Étrangères, de la Francophonie et des Congolais de l'Étranger"
VERSION   = "v5.0"

MOT_URGENCE   = ["urgence","crise","attaque","évacuation","danger","sécurité",
                  "conflit","otage","menace","frontière","incident","alerte"]
MOT_SENSIBLE  = ["confidentiel","ministre","ambassadeur","accord","négociation",
                  "stratégique","audience","coopération","cabinet","diplomatique"]
MOT_IMPORTANT = ["invitation","réunion","rapport","information","suivi",
                  "note","courrier","programme","mission"]

ALERTES = [
    {"pays":"RDC",            "flag":"🇨🇩","sujet":"Situation sécuritaire régionale et risques frontaliers","niveau":"Urgent",   "domaine":"Sécurité"},
    {"pays":"France",         "flag":"🇫🇷","sujet":"Coopération universitaire et mobilité étudiante",        "niveau":"Important","domaine":"Coopération"},
    {"pays":"Chine",          "flag":"🇨🇳","sujet":"Partenariat économique et infrastructures",              "niveau":"Sensible", "domaine":"Économie"},
    {"pays":"Union Africaine","flag":"🌍", "sujet":"Sommet diplomatique régional",                           "niveau":"Important","domaine":"Multilatéral"},
    {"pays":"Gabon",          "flag":"🇬🇦","sujet":"Suivi politique régional post-transition",               "niveau":"Sensible", "domaine":"Politique"},
    {"pays":"ONU",            "flag":"🇺🇳","sujet":"Résolution et coordination internationale",              "niveau":"Important","domaine":"Multilatéral"},
    {"pays":"Angola",         "flag":"🇦🇴","sujet":"Gestion des frontières et coopération sécuritaire",      "niveau":"Urgent",   "domaine":"Sécurité"},
    {"pays":"États-Unis",     "flag":"🇺🇸","sujet":"Relations bilatérales et accord commercial",             "niveau":"Sensible", "domaine":"Économie"},
]

DASH_DATA = [
    {"domaine":"Affaires consulaires",   "total":35,"urgents":7, "sensibles":9},
    {"domaine":"Diaspora",               "total":22,"urgents":3, "sensibles":4},
    {"domaine":"Coopération bilatérale", "total":18,"urgents":4, "sensibles":7},
    {"domaine":"Veille géopolitique",    "total":15,"urgents":6, "sensibles":5},
    {"domaine":"Ambassades",             "total":28,"urgents":5, "sensibles":8},
    {"domaine":"Protocoles",             "total":12,"urgents":1, "sensibles":2},
]

PAGES = ["Accueil","Importer","Résumé","Priorité","Note","Dashboard","Veille","Cabinet"]

# ── PAGE CONFIG ───────────────────────────────────────────────────────
st.set_page_config(
    page_title="DiploIA Congo",
    page_icon="🇨🇬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,700;0,900;1,700&family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
.stApp { background: #F0EDE6 !important; }
.block-container { padding: 0 !important; max-width: 100% !important; }
#MainMenu, footer, header, .stDeployButton { display: none !important; }

/* ── SIDEBAR ── */
section[data-testid="stSidebar"] {
    background: #1C4A2E !important;
    min-width: 240px !important;
    max-width: 240px !important;
}
section[data-testid="stSidebar"] * { color: rgba(255,255,255,0.85) !important; }
section[data-testid="stSidebar"] .stButton > button {
    background: transparent !important;
    color: rgba(255,255,255,0.75) !important;
    border: none !important;
    border-radius: 8px !important;
    text-align: left !important;
    padding: 8px 12px !important;
    font-size: 13.5px !important;
    font-weight: 400 !important;
    width: 100% !important;
    margin-bottom: 2px !important;
    transition: all 0.15s !important;
}
section[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(255,255,255,0.12) !important;
    color: #fff !important;
}

/* ── SECTIONS ── */
.s-cream { background: #F0EDE6; padding: 60px 48px; }
.s-white { background: #FFFFFF; padding: 60px 48px; }
.s-forest { background: #1C4A2E; padding: 60px 48px; }
.inner { max-width: 860px; margin: 0 auto; }

/* ── TYPOGRAPHY ── */
.eyebrow {
    font-size: 11px; font-weight: 700; letter-spacing: .18em;
    text-transform: uppercase; color: #1C4A2E;
    display: flex; align-items: center; gap: 8px; margin-bottom: 16px;
}
.dot { width: 7px; height: 7px; border-radius: 50%; background: #1C4A2E; display: inline-block; }
.eyebrow-w { font-size: 11px; font-weight: 700; letter-spacing: .18em;
    text-transform: uppercase; color: rgba(255,255,255,.55); margin-bottom: 16px; }

.hero {
    font-family: 'Playfair Display', serif;
    font-size: clamp(38px, 5vw, 64px); font-weight: 900;
    line-height: 1.05; letter-spacing: -.02em; color: #1A1A1A; margin-bottom: 20px;
}
.hero em { color: #1C4A2E; font-style: italic; }

.h2 { font-family: 'Playfair Display', serif; font-size: clamp(28px,4vw,46px);
    font-weight: 700; line-height: 1.1; color: #1A1A1A; margin-bottom: 16px; }
.h2-w { font-family: 'Playfair Display', serif; font-size: clamp(28px,4vw,46px);
    font-weight: 700; line-height: 1.1; color: #fff; margin-bottom: 16px; }
.h3 { font-family: 'Playfair Display', serif; font-size: 22px;
    font-weight: 700; color: #1A1A1A; margin-bottom: 10px; }

.body-lg { font-size: 17px; line-height: 1.75; color: #3D3D3D; margin-bottom: 24px; }
.body-md { font-size: 15px; line-height: 1.7; color: #3D3D3D; }
.body-sm { font-size: 13px; line-height: 1.65; color: #888; }
.body-w  { font-size: 15px; line-height: 1.7; color: rgba(255,255,255,.75); margin-bottom: 28px; }

/* ── CARDS ── */
.card { background: #fff; border-radius: 18px; padding: 26px;
    margin-bottom: 16px; box-shadow: 0 2px 14px rgba(0,0,0,.05); }

/* ── METRICS ── */
.metric-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 16px; margin: 32px 0; }
.metric { background: #fff; border-radius: 16px; padding: 22px 18px; text-align: center;
    box-shadow: 0 2px 12px rgba(0,0,0,.05); }
.metric-ey { font-size: 10px; font-weight: 700; letter-spacing: .14em;
    text-transform: uppercase; color: #aaa; margin-bottom: 8px; }
.metric-val { font-family: 'Playfair Display', serif; font-size: 44px;
    font-weight: 700; color: #1C4A2E; line-height: 1; }
.metric-sub { font-size: 12px; color: #aaa; margin-top: 6px; }
.up { color: #059669 !important; } .warn { color: #DC2626 !important; }

/* ── PROGRESS BARS ── */
.prog-row { display: flex; align-items: center; gap: 14px; margin-bottom: 12px; }
.prog-label { min-width: 130px; font-size: 14px; color: #444; text-align: right; }
.prog-track { flex: 1; background: #F0EDE6; border-radius: 100px; height: 34px; overflow: hidden; }
.prog-fill { height: 100%; border-radius: 100px; display: flex; align-items: center;
    padding-left: 14px; font-size: 13px; font-weight: 600; color: #fff; }
.g { background: #1C4A2E; } .b { background: #C4A882; } .p { background: #D4A5A5; }

/* ── BADGES ── */
.badge { display: inline-flex; align-items: center; gap: 5px;
    padding: 4px 12px; border-radius: 100px; font-size: 12px; font-weight: 600; }
.bu { background: #FEE2E2; color: #7F1D1D; }
.bs { background: #FEF3C7; color: #78350F; }
.bi { background: #DBEAFE; color: #1E3A8A; }
.bn { background: #D1FAE5; color: #065F46; }

/* ── PILLS ── */
.pill-row { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 14px; }
.pill { display: inline-flex; align-items: center; gap: 6px; padding: 8px 16px;
    background: #F0EDE6; border-radius: 100px; font-size: 14px; font-weight: 500; }
.ck { color: #1C4A2E; font-weight: 800; }
.pill-w { background: rgba(255,255,255,.12); color: rgba(255,255,255,.9) !important; }

/* ── TIMELINE ── */
.t-item { display: flex; align-items: flex-start; gap: 18px; margin-bottom: 18px; }
.t-icon { width: 50px; height: 50px; background: #1C4A2E; border-radius: 13px;
    display: flex; align-items: center; justify-content: center; font-size: 20px; flex-shrink: 0; }
.t-num { font-size: 11px; font-weight: 700; letter-spacing: .1em; color: #1C4A2E; margin-bottom: 2px; }
.t-title { font-family: 'Playfair Display', serif; font-size: 21px; font-weight: 700; margin-bottom: 4px; }
.t-desc { font-size: 14px; color: #666; line-height: 1.65; }

/* ── ALERT ROWS ── */
.alerte { display: flex; align-items: center; gap: 14px; padding: 13px 16px;
    background: #F8F6F1; border-radius: 12px; margin-bottom: 10px;
    border: 1px solid rgba(28,74,46,.08); }
.al-pays { font-weight: 700; font-size: 14px; margin-bottom: 3px; }
.al-sujet { font-size: 13px; color: #666; }

/* ── NOTE OUTPUT ── */
.note-out { background: #fff; border-left: 4px solid #1C4A2E; border-radius: 12px;
    padding: 22px 24px; font-family: 'Playfair Display', serif; font-size: 14.5px;
    line-height: 1.85; white-space: pre-wrap; box-shadow: 0 2px 14px rgba(0,0,0,.05);
    max-height: 440px; overflow-y: auto; color: #1A1A1A; }

/* ── RESULT BOX ── */
.result-box { background: #F8F6F1; border-radius: 12px; padding: 18px 20px;
    font-size: 15px; line-height: 1.8; color: #1A1A1A; }

/* ── KW PILL ── */
.kw-pill { display: inline-block; background: #F0EDE6;
    border: 1px solid rgba(28,74,46,.2); border-radius: 100px;
    padding: 3px 10px; font-size: 12px; margin: 3px 3px 0 0; color: #1C4A2E; }

/* ── MODULE LOCKED ── */
.mod-lock { display: flex; align-items: center; gap: 14px; padding: 15px 18px;
    background: #F8F6F1; border-radius: 12px; margin-bottom: 10px;
    border: 1px solid rgba(28,74,46,.08); opacity: .7; }

/* ── STREAMLIT OVERRIDES ── */
div[data-testid="stTextArea"] textarea {
    background: #fff !important; border: 1.5px solid rgba(28,74,46,.2) !important;
    border-radius: 10px !important; font-size: 14px !important; color: #1A1A1A !important; }
div[data-testid="stTextArea"] textarea:focus {
    border-color: #1C4A2E !important; box-shadow: 0 0 0 3px rgba(28,74,46,.08) !important; }
div[data-testid="stTextInput"] input {
    background: #fff !important; border: 1.5px solid rgba(28,74,46,.2) !important;
    border-radius: 10px !important; font-size: 14px !important; color: #1A1A1A !important; }
div[data-testid="stSelectbox"] > div > div {
    background: #fff !important; border: 1.5px solid rgba(28,74,46,.2) !important;
    border-radius: 10px !important; }
.stButton > button {
    background: #1C4A2E !important; color: #fff !important; border: none !important;
    border-radius: 10px !important; font-weight: 600 !important; font-size: 14px !important;
    padding: 10px 22px !important; transition: background .2s !important; width: 100%; }
.stButton > button:hover { background: #163d24 !important; }
.stDownloadButton > button {
    background: transparent !important; color: #1C4A2E !important;
    border: 2px solid rgba(28,74,46,.25) !important; border-radius: 10px !important;
    font-weight: 600 !important; }
.stDownloadButton > button:hover { background: #1C4A2E !important; color: #fff !important; }
div[data-testid="stFileUploader"] {
    background: #fff !important; border: 2px dashed rgba(28,74,46,.25) !important;
    border-radius: 13px !important; }
.stSlider > div > div > div { background: #1C4A2E !important; }
div[data-testid="metric-container"] {
    background: #fff; border-radius: 14px; padding: 18px !important;
    box-shadow: 0 2px 10px rgba(0,0,0,.05); }
div[data-testid="metric-container"] [data-testid="stMetricValue"] {
    font-family: 'Playfair Display', serif !important;
    font-size: 36px !important; color: #1C4A2E !important; font-weight: 700 !important; }
label[data-testid="stWidgetLabel"] p {
    font-size: 11px !important; font-weight: 600 !important;
    text-transform: uppercase !important; letter-spacing: .08em !important; color: #555 !important; }
[data-testid="column"] { padding: 0 8px !important; }
[data-testid="column"]:first-child { padding-left: 0 !important; }
[data-testid="column"]:last-child  { padding-right: 0 !important; }
</style>
""", unsafe_allow_html=True)


# ── SESSION STATE ─────────────────────────────────────────────────────
if "page"         not in st.session_state: st.session_state.page = "Accueil"
if "note_content" not in st.session_state: st.session_state.note_content = ""
if "historique"   not in st.session_state: st.session_state.historique = []

# Charge la clé depuis variable d'environnement Railway (priorité)
# ou depuis la saisie manuelle dans la sidebar
if "api_key" not in st.session_state:
    env_key = os.getenv("OPENAI_API_KEY", "")
    st.session_state.api_key = env_key if env_key else None


# ── HELPERS ───────────────────────────────────────────────────────────
def img_b64(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return ""

def detecter_priorite(texte):
    t = (texte or "").lower()
    if any(m in t for m in MOT_URGENCE):   return "Urgent"
    if any(m in t for m in MOT_SENSIBLE):  return "Sensible"
    if any(m in t for m in MOT_IMPORTANT): return "Important"
    return "Normal"

def badge_html(p):
    cls = {"Urgent":"bu","Sensible":"bs","Important":"bi","Normal":"bn"}
    ico = {"Urgent":"🔴","Sensible":"🟡","Important":"🔵","Normal":"🟢"}
    c = cls.get(p, "bn")
    i = ico.get(p, "")
    return f'<span class="badge {c}">{i} {p}</span>'

def resumer_simple(texte, n=5):
    texte = re.sub(r"\s+", " ", texte).strip()
    phrases = re.split(r"(?<=[.!?]) +", texte)
    if len(phrases) <= n:
        return texte
    mots = MOT_URGENCE + MOT_SENSIBLE + MOT_IMPORTANT
    scores = []
    for i, ph in enumerate(phrases):
        s = sum(1 for m in mots if m in ph.lower())
        s += min(len(ph.split()) / 20, 2)
        s += 0.3 if i < 3 else 0
        scores.append(s)
    top = sorted(range(len(scores)), key=lambda x: scores[x], reverse=True)[:n]
    return " ".join(phrases[i] for i in sorted(top))

def gpt_call(messages, api_key, max_tokens=600, json_mode=False):
    """Generic GPT-4o call."""
    if not HAS_OPENAI or not api_key:
        return None
    try:
        client = OpenAI(api_key=api_key)
        kwargs = dict(
            model="gpt-4o",
            messages=messages,
            max_tokens=max_tokens,
            temperature=0.2,
        )
        if json_mode:
            kwargs["response_format"] = {"type": "json_object"}
        resp = client.chat.completions.create(**kwargs)
        return resp.choices[0].message.content.strip()
    except Exception as e:
        st.warning(f"GPT-4o : {e}")
        return None

def resumer(texte, n=5):
    api_key = st.session_state.api_key
    result = gpt_call([
        {"role": "system", "content": (
            "Tu es un conseiller diplomatique senior de la République du Congo. "
            "Tu rédiges des résumés clairs et exploitables pour le cabinet du Ministre."
        )},
        {"role": "user", "content": (
            f"Résume ce document en exactement {n} phrases. "
            f"Identifie les enjeux stratégiques, les acteurs clés et les actions requises. "
            f"Réponds uniquement avec le résumé.\n\nDOCUMENT :\n{texte[:4000]}"
        )}
    ], api_key)
    return result if result else resumer_simple(texte, n)

def extraire_mots_cles(texte):
    api_key = st.session_state.api_key
    result = gpt_call([
        {"role": "system", "content": "Tu es expert en analyse diplomatique. Réponds uniquement en JSON valide."},
        {"role": "user", "content": (
            "Extrais les mots-clés diplomatiques de ce texte. "
            'Format JSON : {"mots_cles": ["mot1","mot2"], "entites": ["Pays/Org"], "actions_requises": ["action1"]}. '
            "Uniquement le JSON.\n\n" + texte[:3000]
        )}
    ], api_key, max_tokens=300, json_mode=True)
    if result:
        try:
            return json.loads(result)
        except Exception:
            pass
    mots = [m for m in MOT_URGENCE + MOT_SENSIBLE + MOT_IMPORTANT if m in texte.lower()]
    return {"mots_cles": mots[:8], "entites": [], "actions_requises": []}

def analyser_ia(texte):
    api_key = st.session_state.api_key
    result = gpt_call([
        {"role": "system", "content": "Tu es expert en diplomatie africaine. Réponds uniquement en JSON valide."},
        {"role": "user", "content": (
            "Analyse ce document diplomatique pour le Cabinet du Ministre du Congo. "
            'JSON : {"priorite":"Urgent|Sensible|Important|Normal","resume":"3 phrases",'
            '"risques":["r1","r2"],"opportunites":["o1"],"recommandation":"1 phrase",'
            '"pays_concernes":["pays1"]}. Uniquement le JSON.\n\n' + texte[:4000]
        )}
    ], api_key, max_tokens=500, json_mode=True)
    if result:
        try:
            return json.loads(result)
        except Exception:
            pass
    return None

def lire_fichier(f):
    nom = f.name.lower()
    if nom.endswith(".txt"):
        return f.read().decode("utf-8", errors="ignore"), None
    if nom.endswith(".csv"):
        df = pd.read_csv(f)
        return df.astype(str).to_string(index=False), df
    if nom.endswith((".xlsx",".xls")):
        df = pd.read_excel(f)
        return df.astype(str).to_string(index=False), df
    if nom.endswith(".pdf") and PdfReader:
        r = PdfReader(f)
        return "\n".join(p.extract_text() or "" for p in r.pages), None
    if nom.endswith(".docx") and DocxDocument:
        doc = DocxDocument(f)
        return "\n".join(p.text for p in doc.paragraphs), None
    return "", None

def generer_note_txt(titre, contexte, objectif, priorite):
    now = datetime.now().strftime("%d/%m/%Y à %H:%M")
    return f"""NOTE AU MINISTRE
{"─"*62}
Institution : {MINISTERE}
Date        : {now}
Objet       : {titre}
Priorité    : {priorite.upper()}
{"─"*62}

1. CONTEXTE

{contexte}

2. OBJECTIF

{objectif}

3. ANALYSE SYNTHÉTIQUE

Le dossier présente un intérêt stratégique pour le cabinet. Il nécessite une coordination institutionnelle rapide et une orientation claire pour la prise de décision ministérielle.

4. POINTS D'ATTENTION

- Vérifier les implications diplomatiques, politiques et consulaires.
- Identifier les partenaires nationaux et internationaux concernés.
- Évaluer les risques pour l'image, la sécurité et les intérêts du Congo.
- Préparer les éléments de langage nécessaires pour le Ministre.

5. RECOMMANDATION

Centraliser les informations disponibles et proposer une décision opérationnelle dans les meilleurs délais.

6. ACTIONS PROPOSÉES

- Désigner un point focal chargé du suivi du dossier.
- Informer les services concernés selon le niveau ({priorite}).
- Préparer une note complémentaire si de nouveaux éléments apparaissent.

{"─"*62}
DiploIA Congo {VERSION}  —  {now}"""

def pdf_note(titre, contenu):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=50, leftMargin=50, topMargin=50, bottomMargin=50)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle("Inst", parent=styles["Title"],
                               textColor=rl_colors.HexColor("#1C4A2E"), fontSize=13, spaceAfter=12))
    els = [Paragraph(MINISTERE, styles["Inst"]), Paragraph(titre, styles["Title"]), Spacer(1, 14)]
    for ligne in contenu.split("\n"):
        if ligne.strip():
            safe = ligne.strip().replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
            els += [Paragraph(safe, styles["BodyText"]), Spacer(1, 5)]
    doc.build(els)
    return buf.getvalue()

def pdf_table(titre, df):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), rightMargin=30, leftMargin=30, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    df2 = df.iloc[:, :8] if len(df.columns) > 8 else df.copy()
    data = [df2.columns.tolist()] + df2.astype(str).head(30).values.tolist()
    tbl = Table(data, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),rl_colors.HexColor("#1C4A2E")),
        ("TEXTCOLOR", (0,0),(-1,0),rl_colors.white),
        ("GRID",      (0,0),(-1,-1),0.4,rl_colors.HexColor("#E0DDD6")),
        ("FONTNAME",  (0,0),(-1,0),"Helvetica-Bold"),
        ("ALIGN",     (0,0),(-1,-1),"CENTER"),
        ("FONTSIZE",  (0,0),(-1,-1),7.5),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[rl_colors.white,rl_colors.HexColor("#F8F6F1")]),
    ]))
    doc.build([Paragraph(titre, styles["Title"]), Spacer(1,10), tbl])
    return buf.getvalue()

def to_excel(df, sheet="Rapport"):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        df.to_excel(w, index=False, sheet_name=sheet[:31])
        wb2 = w.book
        ws2 = w.sheets[sheet[:31]]
        fmt = wb2.add_format({"bold":True,"bg_color":"#1C4A2E","font_color":"#FFFFFF","border":1,"align":"center"})
        for i, col in enumerate(df.columns):
            ws2.write(0, i, col, fmt)
            ws2.set_column(i, i, 26)
    return out.getvalue()


# ── SIDEBAR ───────────────────────────────────────────────────────────
arm_b64 = img_b64("assets/armoiries.jpg")

with st.sidebar:
    # Logo
    if arm_b64:
        st.markdown(
            f'<div style="text-align:center;padding:20px 0 16px">'
            f'<img src="data:image/jpeg;base64,{arm_b64}" width="68" '
            f'style="border-radius:50%;border:3px solid rgba(255,255,255,.4)"/>'
            f'<div style="font-family:Playfair Display,serif;font-size:17px;'
            f'font-weight:700;color:#fff;margin-top:10px">DiploIA Congo</div>'
            f'<div style="font-size:10px;color:rgba(255,255,255,.5);margin-top:3px">'
            f'Cabinet du Ministre</div></div>',
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            '<div style="text-align:center;padding:20px 0 16px">'
            '<div style="font-size:38px">🇨🇬</div>'
            '<div style="font-family:Playfair Display,serif;font-size:17px;'
            'font-weight:700;color:#fff;margin-top:8px">DiploIA Congo</div></div>',
            unsafe_allow_html=True
        )

    st.markdown('<hr style="border:none;border-top:1px solid rgba(255,255,255,.15);margin:0 0 12px"/>', unsafe_allow_html=True)

    # Navigation buttons
    icons = {"Accueil":"🏛️","Importer":"📂","Résumé":"📄","Priorité":"🚨",
             "Note":"📝","Dashboard":"📊","Veille":"🌐","Cabinet":"🔐"}

    sections = {
        "GÉNÉRAL": ["Accueil","Dashboard"],
        "TRAITEMENT": ["Importer","Résumé","Priorité","Note"],
        "PILOTAGE": ["Veille","Cabinet"],
    }

    for sec_title, sec_pages in sections.items():
        st.markdown(
            f'<div style="font-size:9.5px;font-weight:700;letter-spacing:.15em;'
            f'color:rgba(255,255,255,.4);text-transform:uppercase;'
            f'padding:12px 4px 5px">{sec_title}</div>',
            unsafe_allow_html=True
        )
        for p in sec_pages:
            is_active = st.session_state.page == p
            label = f"{'▶  ' if is_active else '     '}{icons[p]}  {p}"
            if st.button(label, key=f"nav_{p}", use_container_width=True):
                st.session_state.page = p
                st.rerun()

    st.markdown('<hr style="border:none;border-top:1px solid rgba(255,255,255,.15);margin:12px 0"/>', unsafe_allow_html=True)

    # API Key
    with st.expander("🔑 Clé API GPT-4o"):
        env_key = os.getenv("OPENAI_API_KEY", "")
        if env_key:
            st.success("✅ Clé chargée depuis Railway")
            st.caption("Variable : OPENAI_API_KEY")
        else:
            key_in = st.text_input(
                "OpenAI API Key",
                type="password",
                placeholder="sk-proj-...",
                help="platform.openai.com/api-keys",
                label_visibility="collapsed"
            )
            if key_in:
                st.session_state.api_key = key_in
                st.success("✅ GPT-4o activé")
            else:
                st.caption("Sans clé : mode règles actif")

    ai_status = "🤖 GPT-4o actif" if st.session_state.api_key else "⚙️ Mode règles"
    st.markdown(
        f'<div style="font-size:10px;color:rgba(255,255,255,.4);'
        f'text-align:center;margin-top:8px;line-height:1.7">'
        f'{ai_status}<br>{datetime.now().strftime("%d %B %Y")}<br>v5.0</div>',
        unsafe_allow_html=True
    )

page = st.session_state.page


# ══════════════════════════════════════════════════════════════════════
# PAGE : ACCUEIL
# ══════════════════════════════════════════════════════════════════════
if page == "Accueil":

    st.markdown("""
    <div class="s-cream"><div class="inner">
        <div class="eyebrow"><span class="dot"></span> CABINET DU MINISTRE · BRAZZAVILLE 2026</div>
        <div class="hero">Ne laissons pas<br><em>la diplomatie</em><br>attendre.</div>
        <p class="body-lg">Analyse automatique des courriers, rapports et alertes diplomatiques
        permettant à <strong>tout conseiller</strong> de produire une note au Ministre
        en quelques secondes — propulsé par <strong>GPT-4o</strong>.</p>
    </div></div>""", unsafe_allow_html=True)

    _, c1, c2, _ = st.columns([1,2,2,1])
    with c1:
        if st.button("✦ Essayez la plateforme"):
            st.session_state.page = "Importer"; st.rerun()
    with c2:
        if st.button("Voir le dashboard →"):
            st.session_state.page = "Dashboard"; st.rerun()

    st.markdown("""
    <div class="s-white"><div class="inner">
        <div class="eyebrow"><span class="dot"></span> COMMENT ÇA MARCHE</div>
        <div class="h2">Du document à la décision.</div>
        <div class="t-item"><div class="t-icon">📂</div><div>
            <div class="t-num">01</div><div class="t-title">Importer</div>
            <div class="t-desc">PDF, Word, Excel, TXT — lecture et extraction automatiques.</div>
        </div></div>
        <div class="t-item"><div class="t-icon">🤖</div><div>
            <div class="t-num">02</div><div class="t-title">Analyser avec GPT-4o</div>
            <div class="t-desc">Priorité, risques, opportunités, mots-clés diplomatiques — analyse IA complète.</div>
        </div></div>
        <div class="t-item"><div class="t-icon">📄</div><div>
            <div class="t-num">03</div><div class="t-title">Résumer</div>
            <div class="t-desc">Résumé intelligent en N phrases adapté au cabinet du Ministre.</div>
        </div></div>
        <div class="t-item"><div class="t-icon">✅</div><div>
            <div class="t-num">04</div><div class="t-title">Décider</div>
            <div class="t-desc">Note en 6 sections générée et téléchargeable en PDF ou TXT officiel.</div>
        </div></div>
    </div></div>
    <div class="s-cream"><div class="inner">
        <div class="card">
            <p class="body-sm" style="margin-bottom:10px">vs. circuit administratif traditionnel</p>
            <div class="pill-row">
                <span class="pill"><span class="ck">✓</span> GPT-4o intégré</span>
                <span class="pill"><span class="ck">✓</span> Export PDF officiel</span>
                <span class="pill"><span class="ck">✓</span> Détection de priorité</span>
                <span class="pill"><span class="ck">✓</span> Mots-clés diplomatiques</span>
                <span class="pill"><span class="ck">✓</span> Veille géopolitique</span>
                <span class="pill"><span class="ck">✓</span> Temps divisé par 5</span>
            </div>
        </div>
    </div></div>
    <div class="s-forest"><div class="inner">
        <div class="eyebrow-w">DÉMO EN DIRECT</div>
        <div class="h2-w">Voyez-le en action.</div>
        <p class="body-w">Collez n'importe quel rapport diplomatique. GPT-4o l'analyse en temps réel.</p>
    </div></div>""", unsafe_allow_html=True)

    _, c1, _ = st.columns([1,3,4])
    with c1:
        if st.button("Accéder à la plateforme →", key="cta2"):
            st.session_state.page = "Importer"; st.rerun()

    # Metrics
    st.markdown('<div class="s-cream"><div class="inner">', unsafe_allow_html=True)
    _, c1,c2,c3,c4,_ = st.columns([.5,1,1,1,1,.5])
    with c1: st.metric("Documents analysés","128","↑ +18")
    with c2: st.metric("Alertes actives","17","⚠ 4 urgentes")
    with c3: st.metric("Notes générées","42","↑ +9")
    with c4: st.metric("Temps gagné","70 %","par dossier")
    st.markdown('</div></div>', unsafe_allow_html=True)

    # Progress bars
    st.markdown("""
    <div class="s-white"><div class="inner"><div class="card">
        <div class="h3">Dossiers par domaine</div>
        <p class="body-sm" style="margin-bottom:18px">130 dossiers actifs</p>
        <div class="prog-row"><span class="prog-label">Consulaires</span><div class="prog-track"><div class="prog-fill g" style="width:90%">35</div></div></div>
        <div class="prog-row"><span class="prog-label">Ambassades</span><div class="prog-track"><div class="prog-fill g" style="width:72%">28</div></div></div>
        <div class="prog-row"><span class="prog-label">Diaspora</span><div class="prog-track"><div class="prog-fill b" style="width:57%">22</div></div></div>
        <div class="prog-row"><span class="prog-label">Coopération</span><div class="prog-track"><div class="prog-fill b" style="width:46%">18</div></div></div>
        <div class="prog-row"><span class="prog-label">Protocoles</span><div class="prog-track"><div class="prog-fill p" style="width:31%">12</div></div></div>
    </div></div></div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# PAGE : IMPORTER
# ══════════════════════════════════════════════════════════════════════
elif page == "Importer":
    st.markdown("""
    <div class="s-cream"><div class="inner">
        <div class="eyebrow"><span class="dot"></span> TRAITEMENT DES DOSSIERS</div>
        <div class="h2">Importez votre document.</div>
        <p class="body-md">PDF · Word · TXT · Excel · CSV — analyse IA automatique.</p>
    </div></div>""", unsafe_allow_html=True)

    _, col, _ = st.columns([.5,8,.5])
    with col:
        st.markdown('<div style="padding:0 48px">', unsafe_allow_html=True)
        uploaded = st.file_uploader("Déposez votre fichier", type=["pdf","docx","txt","xlsx","xls","csv"])

        if uploaded:
            try:
                texte, df = lire_fichier(uploaded)
                st.success(f"✅ Fichier chargé — **{uploaded.name}**")

                ai_result = None
                with st.spinner("Analyse GPT-4o en cours..." if st.session_state.api_key else "Analyse en cours..."):
                    ai_result = analyser_ia(texte[:6000])
                    priorite  = ai_result["priorite"] if ai_result else detecter_priorite(texte)
                    resume_   = ai_result.get("resume","") if ai_result else ""
                    if not resume_:
                        resume_ = resumer(texte[:6000], 5)
                    kw_data   = extraire_mots_cles(texte[:6000])

                c1,c2,c3 = st.columns(3)
                c1.metric("Priorité", priorite)
                c2.metric("Mots", f"{len(texte.split()):,}")
                c3.metric("Format", uploaded.name.split(".")[-1].upper())

                ai_label = "🤖 ANALYSE GPT-4o" if st.session_state.api_key else "⚙️ ANALYSE"
                kw_pills = "".join(
                    f'<span class="kw-pill">{k}</span>'
                    for k in kw_data.get("mots_cles", [])[:8]
                )
                risques_html = ""
                if ai_result and ai_result.get("risques"):
                    risques_html = "".join(
                        f'<div style="font-size:13px;color:#92400E;margin-bottom:4px">⚠ {r}</div>'
                        for r in ai_result["risques"][:3]
                    )

                st.markdown(f"""
                <div class="card" style="margin-top:16px">
                    <p style="font-size:11px;font-weight:700;letter-spacing:.1em;
                       text-transform:uppercase;color:#888;margin-bottom:12px">{ai_label} · {priorite.upper()}</p>
                    <div class="result-box">{resume_}</div>
                    <div style="margin-top:10px;display:flex;align-items:center;gap:8px">
                        <span style="font-size:13px;color:#888">Niveau :</span>
                        {badge_html(priorite)}
                    </div>
                    {'<div style="margin-top:12px"><p style="font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#888;margin-bottom:6px">MOTS-CLÉS DIPLOMATIQUES</p>' + kw_pills + '</div>' if kw_pills else ''}
                    {'<div style="margin-top:12px;padding:12px;background:#FEF3C7;border-radius:8px">' + risques_html + '</div>' if risques_html else ''}
                </div>""", unsafe_allow_html=True)

                # Save to history
                st.session_state.historique.append({
                    "date": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "fichier": uploaded.name,
                    "priorite": priorite,
                    "resume": resume_[:200] + "…"
                })

                if df is not None:
                    st.markdown('<div class="card"><p style="font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#888;margin-bottom:10px">APERÇU DU TABLEAU</p>', unsafe_allow_html=True)
                    st.dataframe(df, use_container_width=True, hide_index=True)
                    st.markdown('</div>', unsafe_allow_html=True)

                note_txt = generer_note_txt(f"Analyse — {uploaded.name}", resume_, "Informer le cabinet.", priorite)
                c1,c2,c3 = st.columns(3)
                with c1:
                    st.download_button("📄 Note PDF", pdf_note("Note au Ministre", note_txt), "note.pdf", "application/pdf")
                with c2:
                    st.download_button("📝 Note TXT", note_txt.encode(), "note.txt", "text/plain")
                if df is not None:
                    with c3:
                        st.download_button("📥 Excel", to_excel(df), "analyse.xlsx",
                                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Erreur : {e}")
        st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# PAGE : RÉSUMÉ
# ══════════════════════════════════════════════════════════════════════
elif page == "Résumé":
    st.markdown("""
    <div class="s-cream"><div class="inner">
        <div class="eyebrow"><span class="dot"></span> RÉSUMÉ AUTOMATIQUE</div>
        <div class="h2">L'essentiel en quelques phrases.</div>
    </div></div>""", unsafe_allow_html=True)

    _, col, _ = st.columns([.5,8,.5])
    with col:
        st.markdown('<div style="padding:0 48px">', unsafe_allow_html=True)
        texte_in = st.text_area("Rapport ou courrier diplomatique", height=180,
            value="Le Ministère reçoit un rapport concernant une situation sécuritaire régionale impliquant plusieurs ressortissants congolais. Le document signale des tensions frontalières et la nécessité d'une coordination rapide avec l'ambassade concernée.")
        nb = st.slider("Nombre de phrases", 3, 10, 5)
        if st.button("✦ Générer le résumé", key="btn_res"):
            with st.spinner("GPT-4o analyse..." if st.session_state.api_key else "Résumé en cours..."):
                res  = resumer(texte_in, nb)
                prio = detecter_priorite(texte_in)
            st.markdown(f"""
            <div class="card" style="margin-top:16px">
                <p style="font-size:11px;font-weight:700;letter-spacing:.1em;
                   text-transform:uppercase;color:#888;margin-bottom:10px">RÉSUMÉ GÉNÉRÉ</p>
                <div class="result-box">{res}</div>
                <div style="margin-top:10px">{badge_html(prio)}</div>
            </div>""", unsafe_allow_html=True)
            c1,c2 = st.columns(2)
            with c1:
                st.download_button("📄 PDF", pdf_note("Résumé", f"Priorité : {prio}\n\n{res}"), "resume.pdf", "application/pdf")
            with c2:
                st.download_button("📝 TXT", res.encode(), "resume.txt", "text/plain")
        st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# PAGE : PRIORITÉ
# ══════════════════════════════════════════════════════════════════════
elif page == "Priorité":
    st.markdown("""
    <div class="s-cream"><div class="inner">
        <div class="eyebrow"><span class="dot"></span> ANALYSE DE PRIORITÉ</div>
        <div class="h2">Quel niveau de traitement ?</div>
    </div></div>""", unsafe_allow_html=True)

    _, col, _ = st.columns([.5,8,.5])
    with col:
        st.markdown('<div style="padding:0 48px">', unsafe_allow_html=True)
        c1,c2 = st.columns(2)
        with c1:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            texte_p = st.text_area("Texte à analyser", height=200,
                placeholder="Alerte urgente — Des ressortissants congolais sont bloqués à la frontière…")
            if st.button("🚨 Analyser", key="btn_prio"):
                p = detecter_priorite(texte_p)
                reco = {
                    "Urgent":    "⚡ Traitement immédiat. Alerter les services sans délai.",
                    "Sensible":  "⚠️ Validation par un responsable avant transmission.",
                    "Important": "📌 Suivi diplomatique dans les 24 heures.",
                    "Normal":    "✅ Circuit administratif standard."
                }
                st.markdown(f"""
                <div style="margin-top:14px">
                    <div style="font-family:'Playfair Display',serif;font-size:30px;
                         font-weight:700;color:#1C4A2E;margin-bottom:8px">{p}</div>
                    {badge_html(p)}
                    <div class="result-box" style="margin-top:12px">{reco[p]}</div>
                </div>""", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with c2:
            st.markdown("""
            <div class="card">
                <p style="font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#888;margin-bottom:14px">GRILLE DE PRIORITÉS</p>
                <div style="padding:13px;border-radius:9px;background:#FEE2E2;border-left:3px solid #DC2626;margin-bottom:10px">
                    <strong style="color:#7F1D1D">🔴 URGENT</strong>
                    <p style="font-size:12px;color:#991B1B;margin-top:3px">Évacuation, crise, attaque, danger sécuritaire</p>
                </div>
                <div style="padding:13px;border-radius:9px;background:#FEF3C7;border-left:3px solid #D97706;margin-bottom:10px">
                    <strong style="color:#78350F">🟡 SENSIBLE</strong>
                    <p style="font-size:12px;color:#92400E;margin-top:3px">Accord, ministre, ambassadeur, confidentiel</p>
                </div>
                <div style="padding:13px;border-radius:9px;background:#DBEAFE;border-left:3px solid #2563EB;margin-bottom:10px">
                    <strong style="color:#1E3A8A">🔵 IMPORTANT</strong>
                    <p style="font-size:12px;color:#1E40AF;margin-top:3px">Réunion, rapport, invitation, mission</p>
                </div>
                <div style="padding:13px;border-radius:9px;background:#D1FAE5;border-left:3px solid #059669">
                    <strong style="color:#065F46">🟢 NORMAL</strong>
                    <p style="font-size:12px;color:#047857;margin-top:3px">Courrier ordinaire, information générale</p>
                </div>
            </div>""", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# PAGE : NOTE
# ══════════════════════════════════════════════════════════════════════
elif page == "Note":
    st.markdown("""
    <div class="s-cream"><div class="inner">
        <div class="eyebrow"><span class="dot"></span> GÉNÉRATEUR DE NOTE</div>
        <div class="h2">Une note officielle en 30 secondes.</div>
    </div></div>""", unsafe_allow_html=True)

    _, col, _ = st.columns([.5,8,.5])
    with col:
        st.markdown('<div style="padding:0 48px">', unsafe_allow_html=True)
        c1, c2 = st.columns(2)

        with c1:
            titre_n    = st.text_input("Objet de la note", "Situation sécuritaire régionale")
            contexte_n = st.text_area("Contexte", height=110,
                value="Une alerte diplomatique signale une situation sensible nécessitant une coordination rapide.")
            objectif_n = st.text_area("Objectif", height=80,
                value="Informer le Ministre et proposer une orientation stratégique.")
            prio_n     = st.selectbox("Niveau de priorité", ["Urgent","Sensible","Important","Normal"], index=1)
            if st.button("✦ Générer la note officielle", key="btn_note"):
                st.session_state.note_content = generer_note_txt(titre_n, contexte_n, objectif_n, prio_n)
                # Save to history
                st.session_state.historique.append({
                    "date": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "fichier": f"Note : {titre_n}",
                    "priorite": prio_n,
                    "resume": contexte_n[:200]
                })

        with c2:
            note_ph = st.session_state.note_content
            if not note_ph:
                note_ph = "La note apparaîtra ici après génération.\n\n1. Contexte\n2. Objectif\n3. Analyse\n4. Points d'attention\n5. Recommandation\n6. Actions"
            st.markdown(f'<div class="note-out">{note_ph}</div>', unsafe_allow_html=True)
            if st.session_state.note_content:
                c1b, c2b = st.columns(2)
                with c1b:
                    st.download_button("📄 PDF officiel",
                                       pdf_note("Note au Ministre", st.session_state.note_content),
                                       "note_ministre.pdf", "application/pdf")
                with c2b:
                    st.download_button("📝 TXT",
                                       st.session_state.note_content.encode(),
                                       "note_ministre.txt", "text/plain")
        st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# PAGE : DASHBOARD
# ══════════════════════════════════════════════════════════════════════
elif page == "Dashboard":
    st.markdown("""
    <div class="s-cream"><div class="inner">
        <div class="eyebrow"><span class="dot"></span> TABLEAU DE BORD</div>
        <div class="h2">Vue stratégique des dossiers actifs.</div>
    </div></div>""", unsafe_allow_html=True)

    df_d = pd.DataFrame(DASH_DATA)
    _, c1,c2,c3,c4,_ = st.columns([.5,1,1,1,1,.5])
    with c1: st.metric("Total dossiers", df_d["total"].sum())
    with c2: st.metric("Urgents", df_d["urgents"].sum())
    with c3: st.metric("Sensibles", df_d["sensibles"].sum())
    with c4: st.metric("Domaines", len(df_d))

    _, col, _ = st.columns([.5,8,.5])
    with col:
        st.markdown('<div style="padding:20px 48px">', unsafe_allow_html=True)
        max_t = df_d["total"].max()
        bars = "".join(
            f'<div class="prog-row"><span class="prog-label">{row["domaine"][:16]}</span>'
            f'<div class="prog-track"><div class="prog-fill g" style="width:{int(row["total"]/max_t*100)}%">'
            f'{row["total"]} dossiers</div></div></div>'
            for _, row in df_d.iterrows()
        )
        st.markdown(f'<div class="card"><div class="h3">Dossiers par domaine</div>{bars}</div>', unsafe_allow_html=True)

        # Historique des notes
        if st.session_state.historique:
            st.markdown('<div class="card"><p style="font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#888;margin-bottom:12px">HISTORIQUE DES ANALYSES</p>', unsafe_allow_html=True)
            df_hist = pd.DataFrame(st.session_state.historique[::-1])
            st.dataframe(df_hist, use_container_width=True, hide_index=True)
            st.markdown('</div>', unsafe_allow_html=True)

        c1,c2 = st.columns(2)
        with c1:
            st.download_button("📥 Excel", to_excel(df_d), "dashboard.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with c2:
            st.download_button("📄 PDF", pdf_table("Tableau de bord", df_d), "dashboard.pdf", "application/pdf")
        st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# PAGE : VEILLE
# ══════════════════════════════════════════════════════════════════════
elif page == "Veille":
    st.markdown("""
    <div class="s-cream"><div class="inner">
        <div class="eyebrow"><span class="dot"></span> VEILLE GÉOPOLITIQUE</div>
        <div class="h2">Alertes par pays, niveau et domaine.</div>
    </div></div>""", unsafe_allow_html=True)

    _, col, _ = st.columns([.5,8,.5])
    with col:
        st.markdown('<div style="padding:0 48px">', unsafe_allow_html=True)
        c1,c2 = st.columns(2)
        with c1:
            niv_f = st.selectbox("Niveau", ["Tous","Urgent","Sensible","Important","Normal"])
        with c2:
            dom_f = st.selectbox("Domaine", ["Tous","Sécurité","Coopération","Économie","Multilatéral","Politique"])

        fil = [a for a in ALERTES
               if (niv_f=="Tous" or a["niveau"]==niv_f)
               and (dom_f=="Tous" or a["domaine"]==dom_f)]

        rows = "".join(f"""
        <div class="alerte">
            <span style="font-size:24px">{a['flag']}</span>
            <div style="flex:1">
                <div class="al-pays">{a['pays']} {badge_html(a['niveau'])}
                    <span style="font-size:11px;color:#aaa;margin-left:8px">{a['domaine']}</span>
                </div>
                <div class="al-sujet">{a['sujet']}</div>
            </div>
        </div>""" for a in fil) or '<p class="body-sm" style="padding:12px">Aucune alerte.</p>'

        st.markdown(f'<div class="card"><p style="font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#888;margin-bottom:14px">ALERTES ({len(fil)})</p>{rows}</div>', unsafe_allow_html=True)

        if fil:
            df_v = pd.DataFrame(fil)
            c1,c2 = st.columns(2)
            with c1:
                st.download_button("📥 Excel", to_excel(df_v,"Veille"), "veille.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with c2:
                st.download_button("📄 PDF", pdf_table("Veille géopolitique", df_v), "veille.pdf", "application/pdf")

        st.markdown('<div class="card" style="margin-top:16px"><p style="font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:#888;margin-bottom:12px">NOTE RAPIDE</p>', unsafe_allow_html=True)
        if fil:
            choix = st.selectbox("Sélectionner une alerte",
                                 [f"{a['flag']} {a['pays']} — {a['sujet']}" for a in fil])
            al = next(a for a in fil if f"{a['flag']} {a['pays']} — {a['sujet']}" == choix)
            if st.button("📝 Générer la note", key="btn_vn"):
                note = generer_note_txt(
                    f"Note — {al['pays']}",
                    f"Alerte : {al['sujet']}. Domaine : {al['domaine']}.",
                    "Informer le Ministre.", al["niveau"]
                )
                st.markdown(f'<div class="note-out">{note}</div>', unsafe_allow_html=True)
                st.download_button("📄 PDF", pdf_note("Note", note), "note_veille.pdf", "application/pdf")
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# PAGE : CABINET
# ══════════════════════════════════════════════════════════════════════
elif page == "Cabinet":
    st.markdown("""
    <div class="s-cream"><div class="inner">
        <div class="eyebrow"><span class="dot"></span> ESPACE CABINET</div>
        <div class="h2">Accès restreint.</div>
        <p class="body-md">Fonctionnalités sécurisées — déploiement V6.</p>
    </div></div>
    <div class="s-forest"><div class="inner">
        <div class="eyebrow-w">MODULES PLANIFIÉS — V6</div>
        <div class="h2-w">La V6 arrive.</div>
        <p class="body-w">Authentification, journal d'audit, circuit de validation, PostgreSQL, tableau de bord temps réel.</p>
        <div class="pill-row">
            <span class="pill pill-w"><span class="ck" style="color:#6EE7B7">✓</span> Comptes hiérarchisés</span>
            <span class="pill pill-w"><span class="ck" style="color:#6EE7B7">✓</span> Chiffrement AES-256</span>
            <span class="pill pill-w"><span class="ck" style="color:#6EE7B7">✓</span> Journal d'audit</span>
            <span class="pill pill-w"><span class="ck" style="color:#6EE7B7">✓</span> Base PostgreSQL</span>
            <span class="pill pill-w"><span class="ck" style="color:#6EE7B7">✓</span> Dashboard temps réel</span>
        </div>
    </div></div>""", unsafe_allow_html=True)

    st.markdown('<div class="s-white"><div class="inner">', unsafe_allow_html=True)
    for ico,titre,desc in [
        ("👤","Comptes hiérarchisés","Ministre · Directeur · Conseillers · Secrétariat"),
        ("📋","Journal d'audit","Traçabilité des accès et validations"),
        ("🔒","Stockage chiffré AES-256","Serveur privé · Sauvegarde automatique"),
        ("✅","Validation hiérarchique","Circuit avant transmission au Ministre"),
        ("🗄️","Base de données PostgreSQL","Persistance des dossiers et historique"),
        ("📡","Tableau de bord temps réel","Mise à jour automatique des indicateurs"),
    ]:
        st.markdown(
            f'<div class="mod-lock"><span style="font-size:22px">{ico}</span>'
            f'<div><strong style="display:block;font-size:14px">{titre}</strong>'
            f'<span class="body-sm">{desc}</span></div></div>',
            unsafe_allow_html=True
        )
    st.markdown('</div></div>', unsafe_allow_html=True)
