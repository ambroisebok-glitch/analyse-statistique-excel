"""
Étudiants Learn (version Gemini) – cours (PDF/Word) -> résumé approfondi, QCM par chapitre (jusqu'à 100),
assistant IA et exports (PDF, Word, HTML, Markdown). Utilisation directe, sans compte.
Services : traitement et analyse de données, accompagnement à la rédaction de mémoire.
Auteur : BOKOUMBO Tibetobe Ambroise
Contact : ambroisebok@gmail.com | +228 91942948
"""
import html
import io
import json
import os
import random
import re
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from functools import lru_cache

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from google import genai
from google.genai import types
from pypdf import PdfReader

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
# Modèle préféré (modifiable sans toucher au code via la variable GEMINI_MODEL / secrets Streamlit)
MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite")
MAX_CHARS = 700_000                 # le cours est lu en entier (≈ 175 000 tokens)
MAX_PDF_OCTETS = 10 * 1024 * 1024   # taille max d'un PDF envoyé directement à Gemini
LOT_QCM = 20                        # questions générées par appel
MAX_QCM = 100
LETTRES = "ABCD"
PAGES = ["Accueil", "Mon cours", "Résumé", "QCM", "Assistant", "Historique", "Services & contact"]

AUTEUR = "BOKOUMBO Tibetobe Ambroise"
EMAIL = "ambroisebok@gmail.com"
TEL_REP = ("+228 96286244", "+22896286244")   # contact services (données, mémoire)
TEL_AMB = ("+228 91942948", "+22891942948")   # contact principal
APP_URL = "https://etudiantslearn.streamlit.app/"
SEO_TITRE = "Étudiants Learn – Résumés de cours, QCM et assistant IA"
SEO_DESC = ("Étudiants Learn aide les étudiants à réviser leurs cours PDF ou Word grâce à des résumés "
            "approfondis, des QCM corrigés jusqu'à 100 questions et un assistant IA. Accompagnement "
            "en traitement de données et en rédaction de mémoire.")

st.set_page_config(
    page_title=SEO_TITRE, page_icon="🎓", layout="wide", initial_sidebar_state="collapsed",
    menu_items={"About": f"**Étudiants Learn** — conçu par {AUTEUR} · {EMAIL} · {TEL_AMB[0]}"})

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Playfair+Display:wght@600;700&display=swap');
:root { --navy:#12306B; --navy-d:#0B2147; --gold:#C9A227; --gold-l:#FBF5DF; --ink:#1B2433;
        --muted:#5F6B7F; --line:#DDE3EC; --bg:#EEF1F6; --accent:#12306B; --accent-d:#0B2147;
        --accent-l:#EAF0FA; --serif:'Playfair Display',Georgia,'Times New Roman',serif; }
html, body, .stApp { font-family:'Inter',-apple-system,'Segoe UI',Roboto,Arial,sans-serif; color:var(--ink); }
.stApp { background:var(--bg); }
h1, h2, h3 { font-family:var(--serif); color:var(--navy-d); letter-spacing:-.005em; }
#MainMenu, footer { visibility:hidden; }
header[data-testid="stHeader"] { display:none; }
[data-testid="stSidebar"], [data-testid="collapsedControl"], [data-testid="stSidebarCollapsedControl"] { display:none; }
iframe[height="0"] { display:none; }
.block-container { padding:0 2.2rem 0 2.2rem; max-width:1180px; background:#fff;
  border-left:1px solid var(--line); border-right:1px solid var(--line);
  box-shadow:0 0 40px rgba(11,33,71,.07); }

/* En-tête du site : barre d'information, identité, menu */
.st-key-top { gap:0 !important; margin:0 -2.2rem -1rem -2.2rem; }
.topbar { background:var(--navy-d); color:#C3D0E8; font-size:.78rem; padding:.5rem 2.2rem;
  display:flex; justify-content:space-between; flex-wrap:wrap; gap:.3rem 1.5rem; }
.topbar a { color:#fff !important; text-decoration:none; font-weight:500; }
.topbar a:hover { text-decoration:underline; }
.topbar .sep { opacity:.35; margin:0 .5rem; }
.entete-site { display:flex; align-items:center; gap:1rem; padding:1.15rem 2.2rem; background:#fff;
  border-bottom:1px solid var(--line); }
.entete-site .logo { width:54px; height:54px; border-radius:50%; background:var(--navy); color:var(--gold);
  border:3px solid var(--gold); display:flex; align-items:center; justify-content:center;
  font-family:var(--serif); font-weight:700; font-size:1.2rem; letter-spacing:.02em; flex:0 0 54px; }
.entete-site .nom { font-family:var(--serif); font-size:1.65rem; font-weight:700; color:var(--navy-d);
  line-height:1.1; }
.entete-site .devise { font-size:.78rem; letter-spacing:.14em; text-transform:uppercase; color:var(--muted);
  margin-top:.25rem; }
.st-key-navbar { background:var(--navy); padding:0 2.2rem; }
.st-key-navbar [role="radiogroup"] { gap:0; flex-wrap:nowrap; overflow-x:auto; }
.st-key-navbar [role="radiogroup"] label { padding:.85rem 1.05rem; margin:0; border-radius:0;
  border-bottom:3px solid transparent; white-space:nowrap; cursor:pointer; transition:background .15s; }
.st-key-navbar [role="radiogroup"] label > div:first-child { display:none; }
.st-key-navbar [role="radiogroup"] label p { color:#DCE5F5; font-size:.9rem; font-weight:500; }
.st-key-navbar [role="radiogroup"] label:hover { background:rgba(255,255,255,.08); }
.st-key-navbar [role="radiogroup"] label:has(input:checked) { border-bottom-color:var(--gold);
  background:rgba(255,255,255,.10); }
.st-key-navbar [role="radiogroup"] label:has(input:checked) p { color:#fff; font-weight:700; }

/* Bandeau de page (pages intérieures) */
.banniere { background:linear-gradient(120deg,var(--navy-d),var(--navy) 70%,#1D4690); color:#fff;
  padding:2rem 2.2rem 1.8rem; margin:0 -2.2rem 1.8rem -2.2rem; border-bottom:4px solid var(--gold); }
.banniere .fil { font-size:.78rem; color:#B9C7E2; margin-bottom:.55rem; }
.banniere h1 { font-family:var(--serif); color:#fff; font-size:2rem; line-height:1.15; margin:0; padding:0; }
.banniere p { color:#D5DFF1; margin:.55rem 0 0 0; max-width:780px; font-size:.98rem; line-height:1.55; }
.banniere .chip { display:inline-block; margin-top:.95rem; background:rgba(255,255,255,.12);
  border:1px solid rgba(255,255,255,.28); border-radius:999px; padding:.22rem .85rem; font-size:.78rem; }

/* Page d'accueil */
.hero { position:relative; overflow:hidden; color:#fff; padding:3.6rem 2.2rem 3.4rem;
  margin:0 -2.2rem 1.4rem -2.2rem; border-bottom:4px solid var(--gold);
  background:linear-gradient(115deg,var(--navy-d) 0%,var(--navy) 60%,#1F4C9A 100%); }
.hero::after { content:"∑  π  ∫  Δ  ∞"; position:absolute; right:1.5rem; bottom:-1.4rem;
  font-family:Georgia,serif; font-size:7rem; color:rgba(255,255,255,.06); white-space:nowrap; }
.hero .kicker { text-transform:uppercase; letter-spacing:.16em; font-size:.74rem; color:var(--gold);
  font-weight:700; margin-bottom:.9rem; }
.hero h1 { font-family:var(--serif); color:#fff; font-size:2.7rem; line-height:1.15; margin:0 0 1rem 0;
  padding:0; max-width:780px; }
.hero p { color:#D5DFF1; font-size:1.05rem; line-height:1.65; max-width:640px; margin:0; position:relative; z-index:1; }
.chiffres { display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); border:1px solid var(--line);
  border-radius:6px; margin:.6rem 0 .4rem 0; overflow:hidden; }
.chiffres div { padding:1.2rem 1rem; text-align:center; background:#fff; border-right:1px solid var(--line); }
.chiffres div:last-child { border-right:none; }
.chiffres b { display:block; font-family:var(--serif); font-size:2rem; color:var(--navy); line-height:1.1; }
.chiffres span { font-size:.8rem; color:var(--muted); }
.section-titre { font-family:var(--serif); font-size:1.5rem; font-weight:700; color:var(--navy-d);
  margin:2.4rem 0 1.1rem 0; padding-bottom:.6rem; position:relative; }
.section-titre::after { content:""; position:absolute; left:0; bottom:0; width:56px; height:3px;
  background:var(--gold); }
.cartes { display:grid; grid-template-columns:repeat(auto-fit,minmax(240px,1fr)); gap:1.1rem;
  margin:.4rem 0 1rem 0; }
.carte { border:1px solid var(--line); border-top:3px solid var(--navy); border-radius:4px;
  padding:1.2rem 1.3rem; background:#fff; transition:box-shadow .2s, transform .2s; }
.carte:hover { box-shadow:0 10px 26px rgba(11,33,71,.10); transform:translateY(-2px); }
.carte .tag { font-size:.68rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase;
  color:var(--gold); margin-bottom:.45rem; }
.carte h4 { margin:0 0 .4rem 0; font-family:var(--serif); font-size:1.15rem; color:var(--navy-d); }
.carte p { margin:0; color:#4A566A; font-size:.88rem; line-height:1.6; }
.etapes { display:flex; flex-wrap:wrap; gap:.5rem; margin:0 0 1.4rem 0; }
.etape { display:inline-flex; align-items:center; gap:.5rem; padding:.3rem .85rem .3rem .35rem;
  border:1px solid var(--line); border-radius:999px; background:#fff; font-size:.82rem;
  font-weight:500; color:var(--muted); }
.etape i { font-style:normal; width:22px; height:22px; border-radius:50%; background:#F1F5F9;
  display:inline-flex; align-items:center; justify-content:center; font-size:.72rem; font-weight:700; }
.etape.ok { color:#065F46; border-color:#A7F3D0; background:#F0FDF4; }
.etape.ok i { background:#10B981; color:#fff; }

/* Boutons et composants Streamlit */
.btn { display:inline-block; padding:.6rem 1.1rem; border-radius:4px; font-weight:600; font-size:.88rem;
  text-decoration:none !important; background:var(--navy); color:#fff !important;
  margin:.25rem .4rem .25rem 0; }
.btn:hover { background:var(--navy-d); }
.btn.wa { background:#15803D; } .btn.wa:hover { background:#166534; }
.btn.clair { background:var(--accent-l); color:var(--navy) !important; }
.stButton>button, .stDownloadButton>button { border-radius:4px; font-weight:600;
  border:1px solid var(--line); transition:all .15s; }
.stButton>button:hover, .stDownloadButton>button:hover { border-color:var(--navy); color:var(--navy); }
.stButton>button[kind="primary"], .stDownloadButton>button[kind="primary"],
button[data-testid="stBaseButton-primary"], button[data-testid="stBaseButton-primaryFormSubmit"] {
  background:var(--navy); border:1px solid var(--navy); color:#fff; }
.stButton>button[kind="primary"]:hover, .stDownloadButton>button[kind="primary"]:hover,
button[data-testid="stBaseButton-primary"]:hover, button[data-testid="stBaseButton-primaryFormSubmit"]:hover {
  background:var(--navy-d); border-color:var(--navy-d); color:#fff; }
div[data-testid="stMetric"] { background:#fff; border:1px solid var(--line); border-left:3px solid var(--gold);
  border-radius:4px; padding:.9rem 1.1rem; }
div[data-testid="stExpander"] { border:1px solid var(--line); border-radius:6px; background:#fff; }
div[data-testid="stForm"] { border:1px solid var(--line); border-radius:6px; background:#fff; }
[data-testid="stFileUploaderDropzone"] { border:1.5px dashed #B8C4D8; background:#F8FAFD; border-radius:6px; }
[data-testid="stChatMessage"] { background:#fff; border:1px solid var(--line); border-radius:8px; }
[data-testid="stVerticalBlockBorderWrapper"] { border-radius:6px; }

/* Services */
.services { display:grid; grid-template-columns:repeat(auto-fit,minmax(300px,1fr)); gap:1.2rem;
  margin:.6rem 0 1.6rem 0; }
.service { background:#fff; border:1px solid var(--line); border-top:3px solid var(--gold); border-radius:4px;
  padding:1.4rem 1.4rem 1.1rem; display:flex; flex-direction:column;
  transition:box-shadow .2s, transform .2s; }
.service:hover { box-shadow:0 12px 30px rgba(11,33,71,.10); transform:translateY(-2px); }
.service .ico { width:46px; height:46px; border-radius:50%; background:var(--accent-l); color:var(--navy);
  display:flex; align-items:center; justify-content:center; margin-bottom:.9rem; }
.service .ico svg { width:22px; height:22px; fill:none; stroke:currentColor; stroke-width:1.8;
  stroke-linecap:round; stroke-linejoin:round; }
.service .tag { font-size:.68rem; font-weight:700; letter-spacing:.1em; text-transform:uppercase;
  color:var(--gold); margin-bottom:.2rem; }
.service h4 { margin:0 0 .4rem 0; font-family:var(--serif); font-size:1.2rem; color:var(--navy-d); }
.service > p { margin:0 0 .8rem 0; color:#4A566A; font-size:.88rem; line-height:1.6; }
.service ul { margin:0 0 1rem 0; padding:0; list-style:none; display:grid; gap:.4rem; }
.service li { font-size:.86rem; color:#334155; padding-left:1.15rem; position:relative; }
.service li::before { content:""; position:absolute; left:0; top:.5em; width:6px; height:6px;
  background:var(--gold); transform:rotate(45deg); }
.service .contact { margin-top:auto; padding-top:.8rem; border-top:1px solid var(--line);
  font-size:.8rem; color:var(--muted); }
.service .contact b { color:var(--ink); }
.parcours { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:1.2rem;
  margin:.5rem 0 1.5rem 0; counter-reset:etape; }
.parcours div { border-left:3px solid var(--gold); background:#F8FAFD; padding:.9rem 1rem; border-radius:0 4px 4px 0; }
.parcours b { display:block; font-family:var(--serif); font-size:1.02rem; color:var(--navy-d); margin-bottom:.2rem; }
.parcours span { font-size:.85rem; color:#4A566A; line-height:1.55; }

/* Pied de page */
.pied-site { background:var(--navy-d); color:#B9C7E2; margin:3rem -2.2rem 0 -2.2rem;
  padding:2.5rem 2.2rem 0; border-top:4px solid var(--gold); }
.pied-site .cols { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:2rem; }
.pied-site h5 { color:#fff; font-family:var(--serif); font-size:1.05rem; margin:0 0 .8rem 0; }
.pied-site p, .pied-site li { font-size:.85rem; line-height:1.75; margin:0; }
.pied-site ul { list-style:none; padding:0; margin:0; }
.pied-site a { color:#E2E8F5 !important; text-decoration:none; }
.pied-site a:hover { text-decoration:underline; }
.pied-site .copy { border-top:1px solid rgba(255,255,255,.14); margin-top:1.8rem; padding:1rem 0;
  font-size:.76rem; text-align:center; color:#8FA2C4; }
@media (max-width: 700px) {
  .block-container, .banniere, .hero, .pied-site, .topbar, .entete-site, .st-key-navbar { }
  .hero h1 { font-size:2rem; } .banniere h1 { font-size:1.5rem; } .entete-site .nom { font-size:1.3rem; }
  .topbar { font-size:.7rem; } }
</style>
""",
    unsafe_allow_html=True,
)


def injecter_seo():
    """Ajoute les balises SEO dans le <head> de la page (meilleur effort : Streamlit est une
    application dynamique, le site vitrine statique reste indispensable pour un bon référencement)."""
    if st.session_state.get("seo_ok"):
        return
    st.session_state.seo_ok = True
    donnees = {
        "titre": SEO_TITRE, "desc": SEO_DESC, "url": APP_URL,
        "json": {
            "@context": "https://schema.org", "@type": "WebApplication", "name": "Étudiants Learn",
            "url": APP_URL, "description": SEO_DESC, "inLanguage": "fr",
            "applicationCategory": "EducationalApplication", "operatingSystem": "Web",
            "author": {"@type": "Person", "name": AUTEUR, "email": EMAIL, "telephone": TEL_AMB[0]},
        },
    }
    js = """<script>
    try {
      const D = __DATA__, d = window.parent.document;
      d.documentElement.lang = "fr";
      const meta = (a, k, v) => { let m = d.head.querySelector('meta[' + a + '="' + k + '"]');
        if (!m) { m = d.createElement("meta"); m.setAttribute(a, k); d.head.appendChild(m); }
        m.setAttribute("content", v); };
      meta("name", "description", D.desc);
      meta("name", "keywords", "réviser cours, résumé de cours IA, QCM en ligne, quiz, assistant IA étudiant, " +
        "traitement de données, analyse de données, rédaction de mémoire, aide mémoire étudiant, Togo");
      meta("name", "author", "BOKOUMBO Tibetobe Ambroise");
      meta("name", "robots", "index, follow");
      meta("property", "og:type", "website"); meta("property", "og:title", D.titre);
      meta("property", "og:description", D.desc); meta("property", "og:url", D.url);
      meta("property", "og:locale", "fr_FR");
      meta("name", "twitter:card", "summary"); meta("name", "twitter:title", D.titre);
      let c = d.head.querySelector('link[rel="canonical"]');
      if (!c) { c = d.createElement("link"); c.rel = "canonical"; d.head.appendChild(c); }
      c.href = D.url;
      if (!d.getElementById("ld-revisepro")) { const s = d.createElement("script");
        s.type = "application/ld+json"; s.id = "ld-revisepro"; s.textContent = JSON.stringify(D.json);
        d.head.appendChild(s); }
    } catch (e) {}
    </script>""".replace("__DATA__", json.dumps(donnees))
    if hasattr(st, "iframe"):          # Streamlit récent
        st.iframe(js, height=1)
    else:                              # anciennes versions
        import streamlit.components.v1 as components
        components.html(js, height=0)


injecter_seo()


# ----------------------------------------------------------------------------
# Stockage de la session (sans compte, sans base de données partagée)
# ----------------------------------------------------------------------------
# Chaque visiteur travaille dans sa propre session : ses cours, résumés et QCM ne sont visibles
# que de lui et disparaissent à la fermeture ou à l'actualisation de la page (pensez à télécharger
# vos fiches). Cela évite aussi de mélanger les données de visiteurs différents.
def maintenant() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def stock() -> dict:
    return st.session_state.setdefault(
        "stock", {"seq": 0, "cours": {}, "resumes": {}, "qcms": {}, "tentatives": [], "chats": {}})


def nouvel_id() -> int:
    s = stock()
    s["seq"] += 1
    return s["seq"]


def enregistrer_cours(c: dict) -> int:
    i = nouvel_id()
    c["id"] = i
    c["cree"] = maintenant()
    stock()["cours"][i] = c
    return i


def charger_cours(cid: int):
    return stock()["cours"].get(cid)


def supprimer_cours(cid: int):
    s = stock()
    ids_qcm = [i for i, q in s["qcms"].items() if q.get("cours_id") == cid]
    s["tentatives"] = [t for t in s["tentatives"] if t["qcm_id"] not in ids_qcm]
    for i in ids_qcm:
        s["qcms"].pop(i, None)
    for i in [i for i, r in s["resumes"].items() if r.get("cours_id") == cid]:
        s["resumes"].pop(i, None)
    s["chats"].pop(cid, None)
    s["cours"].pop(cid, None)


def charger_chat(cid):
    return list(stock()["chats"].get(cid, [])) if cid else []


def sauver_chat(cid, msgs):
    if cid:
        stock()["chats"][cid] = list(msgs)


# ----------------------------------------------------------------------------
# Lecture des fichiers
# ----------------------------------------------------------------------------
def lire_fichier(fichier) -> dict:
    nom = fichier.name
    data = fichier.getvalue()
    pdf, pages = None, 0
    if nom.lower().endswith(".pdf"):
        reader = PdfReader(io.BytesIO(data))
        pages = len(reader.pages)
        texte = "\n".join((p.extract_text() or "") for p in reader.pages)
        if len(data) <= MAX_PDF_OCTETS:
            pdf = data
    elif nom.lower().endswith(".docx"):
        doc = Document(io.BytesIO(data))
        lignes = [p.text for p in doc.paragraphs]
        for t in doc.tables:
            for row in t.rows:
                lignes.append(" | ".join(c.text for c in row.cells))
        texte = "\n".join(lignes)
    else:
        raise ValueError("Format non pris en charge (PDF ou .docx uniquement).")
    texte = re.sub(r"[ \t]+", " ", texte)
    texte = re.sub(r"\n{3,}", "\n\n", texte).strip()
    return {"nom": nom, "texte": texte, "pdf": pdf, "pages": pages, "vision": pdf is not None}


# ----------------------------------------------------------------------------
# Gemini
# ----------------------------------------------------------------------------
@st.cache_data(ttl=3600, show_spinner=False)
def modeles_disponibles(cle: str) -> list:
    """Modèle préféré d'abord, puis les modèles 'flash' du compte, du plus récent au plus ancien."""
    trouves = []
    try:
        client = genai.Client(api_key=cle)
        exclus = ("image", "tts", "live", "audio", "embedding", "robotics", "thinking")
        for m in client.models.list():
            n = m.name.replace("models/", "")
            actions = getattr(m, "supported_actions", None) or ["generateContent"]
            if "flash" in n and "generateContent" in actions and not any(x in n for x in exclus):
                trouves.append(n)
    except Exception:
        pass

    def tri(n):
        v = re.search(r"gemini-(\d+(?:\.\d+)?)", n)
        version = float(v.group(1)) if v else 0.0
        stable = not any(x in n for x in ("preview", "exp"))
        return (version, stable, "lite" not in n)

    trouves.sort(key=tri, reverse=True)
    noms = [MODEL] + [n for n in trouves if n != MODEL]
    return noms[:4]


def _temporaire(msg: str) -> bool:
    return any(m in msg for m in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED",
                                  "500", "INTERNAL", "overloaded", "DEADLINE"))


def parts_cours(cours: dict) -> list:
    if cours.get("pdf") and cours.get("vision"):
        return [types.Part.from_bytes(data=cours["pdf"], mime_type="application/pdf")]
    return [types.Part.from_text(text="COURS :\n" + cours["texte"][:MAX_CHARS])]


def requete(cours: dict, prompt: str) -> list:
    return [types.Content(role="user",
                          parts=parts_cours(cours) + [types.Part.from_text(text=prompt)])]


REGLAGES = {"sans_reflexion": True, "max_tokens": None}  # mémorise ce que le modèle accepte


def generer(contenu, cle: str, system=None, max_tokens=16000, temperature=0.4,
            reflexion=True) -> str:
    client = genai.Client(api_key=cle)

    def config():
        kw = dict(system_instruction=system, temperature=temperature,
                  max_output_tokens=min(max_tokens, REGLAGES["max_tokens"] or max_tokens))
        if not reflexion and REGLAGES["sans_reflexion"]:
            kw["thinking_config"] = types.ThinkingConfig(thinking_budget=0)
        return types.GenerateContentConfig(**kw)

    derniere = None
    for modele in modeles_disponibles(cle):
        for essai in range(2):
            try:
                rep = client.models.generate_content(model=modele, contents=contenu, config=config())
                if rep.text and rep.text.strip():
                    return rep.text
                derniere = RuntimeError("Réponse vide")
            except Exception as e:
                derniere = e
                msg = str(e)
                if "400" in msg or "INVALID_ARGUMENT" in msg:
                    # Réglage refusé par le modèle : on retire d'abord la "réflexion",
                    # puis on plafonne la longueur de sortie, sinon c'est une vraie erreur.
                    if not reflexion and REGLAGES["sans_reflexion"]:
                        REGLAGES["sans_reflexion"] = False
                        continue
                    if (REGLAGES["max_tokens"] or max_tokens) > 8192:
                        REGLAGES["max_tokens"] = 8192
                        continue
                    raise
                if "404" in msg or "NOT_FOUND" in msg:
                    break  # modèle indisponible : passer au suivant
                if _temporaire(msg):
                    time.sleep(3 * (essai + 1))
                    continue
                raise
    raise RuntimeError("Gemini est surchargé ou son quota est atteint. "
                       f"Réessaie dans quelques minutes. ({derniere})")


def flux_chat(contenu, cle: str, system: str):
    """Génère la réponse en streaming (avec repli sur d'autres modèles)."""
    client = genai.Client(api_key=cle)
    cfg = types.GenerateContentConfig(system_instruction=system, temperature=0.5,
                                      max_output_tokens=16000)
    derniere = None
    for modele in modeles_disponibles(cle):
        deja = False
        try:
            for morceau in client.models.generate_content_stream(
                    model=modele, contents=contenu, config=cfg):
                if morceau.text:
                    deja = True
                    yield morceau.text
            return
        except Exception as e:
            derniere = e
            if deja:
                raise
            msg = str(e)
            if "404" in msg or "NOT_FOUND" in msg or _temporaire(msg):
                time.sleep(2)
                continue
            raise
    raise RuntimeError(f"Gemini est indisponible pour le moment. ({derniere})")


def normaliser_math(t: str) -> str:
    """Convertit \\( \\) et \\[ \\] en $ et $$ (rendu LaTeX de Streamlit)."""
    t = re.sub(r"\\\[(.+?)\\\]", lambda m: "\n$$" + m.group(1).strip() + "$$\n", t, flags=re.S)
    t = re.sub(r"\\\((.+?)\\\)", lambda m: "$" + m.group(1).strip() + "$", t, flags=re.S)
    return t


REGLE_MATHS = ("Pour toutes les formules mathématiques, utilise LaTeX : $...$ en ligne et $$...$$ "
               "pour les formules sur leur propre ligne. N'utilise jamais \\( \\) ni \\[ \\].")


# ----------------------------------------------------------------------------
# Chapitres
# ----------------------------------------------------------------------------
def detecter_chapitres(cours: dict, cle: str) -> list:
    prompt = f"""Analyse la structure du cours fourni et liste ses chapitres (ou grandes parties) DANS L'ORDRE,
en couvrant TOUT le document du début à la fin. Si le cours n'a pas de chapitres explicites, découpe-le
en parties thématiques cohérentes (entre 3 et 12). Maximum 40 entrées.
Réponds UNIQUEMENT avec une ligne par chapitre, au format exact :
CHAPITRE | titre du chapitre | description en une phrase des notions couvertes
Aucun autre texte."""
    brut = generer(requete(cours, prompt), cle, max_tokens=6000, temperature=0.1, reflexion=False)
    chap = []
    for ligne in brut.splitlines():
        m = re.match(r"\s*CHAPITRE\s*\|\s*(.+?)\s*\|\s*(.*)", ligne)
        if m:
            chap.append({"titre": m.group(1).strip(" *#"), "desc": m.group(2).strip()})
    return chap[:40] or [{"titre": "Cours complet", "desc": ""}]


def portee_texte(c: dict) -> str:
    if c["titre"] == "Cours complet":
        return "l'ensemble du cours, en couvrant toutes ses parties de façon équilibrée"
    return f"le chapitre « {c['titre']} » ({c['desc']})" if c["desc"] else f"le chapitre « {c['titre']} »"


# ----------------------------------------------------------------------------
# Résumé approfondi
# ----------------------------------------------------------------------------
NIVEAUX = {
    "Essentiel": ("synthétique : environ une page par chapitre, uniquement l'essentiel "
                  "(définitions, formules, résultats à retenir)", 8000),
    "Détaillé": ("détaillée : environ deux à trois pages par chapitre, avec explications, "
                 "formules et exemples", 16000),
    "Approfondi": ("très approfondie et EXHAUSTIVE : reprends TOUT le contenu du chapitre sans rien omettre, "
                   "avec démonstrations, exemples résolus étape par étape (calculs intermédiaires détaillés), "
                   "remarques, cas particuliers et applications. Plus c'est complet et précis, mieux c'est",
                   32000),
}


def generer_resume(cours, chapitres, niveau, cle, progres=None) -> dict:
    consigne, max_tok = NIVEAUX[niveau]
    liste = "\n".join(f"- {c['titre']}" for c in chapitres)
    modeles_disponibles(cle)  # préchauffe le cache (avant les threads)

    def apercu_fn():
        prompt = f"""À partir du cours fourni, rédige la « Vue d'ensemble » d'une fiche de révision portant sur :
{liste}
Format strict : la première ligne est « TITRE: <titre du cours> ». Ensuite, en Markdown
(titres de niveau ### uniquement) :
### Résumé de synthèse (3 à 4 paragraphes denses)
### Plan et objectifs d'apprentissage
### Notions clés à maîtriser (liste à puces)
### Glossaire (15 à 25 termes, au format : **terme** : définition)
Rédige en français. {REGLE_MATHS}"""
        return generer(requete(cours, prompt), cle, max_tokens=12000, reflexion=False)

    def fiche_fn(c):
        prompt = f"""Rédige la fiche de révision du chapitre suivant du cours fourni : {portee_texte(c)}.
Cette fiche doit être {consigne}.
Structure en Markdown, avec des titres de niveau ### (jamais # ni ##) :
### Idée générale
### Notions et définitions   (termes importants en **gras**)
### Théorèmes, propriétés et formules   (énoncés complets avec leurs conditions d'application)
### Démonstrations et justifications importantes   (étapes détaillées)
### Méthodes et techniques
### Exemples résolus   (ceux du cours, résolus pas à pas, puis un ou deux exemples types supplémentaires)
### Pièges et erreurs fréquentes
### À retenir   (mini fiche mémo)
Omets une rubrique seulement si elle n'a aucun sens pour ce chapitre. Utilise des tableaux Markdown quand
c'est utile. Rédige en français, sans phrase d'introduction du type « Voici ». {REGLE_MATHS}"""
        return generer(requete(cours, prompt), cle, max_tokens=max_tok)

    total = len(chapitres) + 1
    brut_apercu, mds, echecs, fait = "", [None] * len(chapitres), 0, 0
    with ThreadPoolExecutor(max_workers=min(4, total)) as ex:
        futs = {ex.submit(apercu_fn): -1}
        for k, c in enumerate(chapitres):
            futs[ex.submit(fiche_fn, c)] = k
        for f in as_completed(futs):
            k = futs[f]
            try:
                res = f.result()
            except Exception as e:
                res = None
                if k >= 0:
                    echecs += 1
                    mds[k] = f"*⚠️ Cette section n'a pas pu être générée ({e}). Relance le résumé.*"
                else:
                    brut_apercu = ""
            if res is not None:
                if k < 0:
                    brut_apercu = res
                else:
                    mds[k] = normaliser_math(res.strip())
            fait += 1
            if progres:
                progres(fait / total, f"{fait}/{total} parties terminées")

    titre = cours["nom"].rsplit(".", 1)[0]
    m = re.match(r"\s*TITRE\s*:\s*(.+)", brut_apercu)
    if m:
        titre = m.group(1).strip(" *#")
        brut_apercu = brut_apercu[m.end():]
    sections = [{"titre": c["titre"], "md": mds[k]} for k, c in enumerate(chapitres)]
    return {"titre": titre, "niveau": niveau, "apercu": normaliser_math(brut_apercu.strip()),
            "sections": sections, "echecs": echecs, "cree": maintenant()}


def resume_en_markdown(r: dict) -> str:
    out = [f"# {r['titre']}", "", r["apercu"], ""]
    for s in r["sections"]:
        out += [f"## {s['titre']}", "", s["md"], ""]
    return "\n".join(out)


# ----------------------------------------------------------------------------
# Exports : HTML, Word et PDF (formules rendues en images)
# ----------------------------------------------------------------------------
def slug(t: str) -> str:
    return re.sub(r"[^a-zA-Z0-9]+", "_", t).strip("_")[:50] or "resume"


def md_vers_html(md_texte: str, titre: str) -> str:
    import markdown
    maths = []

    def garde(m):
        maths.append(m.group(0))
        return f"@@MATH{len(maths) - 1}@@"

    t = re.sub(r"\$\$.+?\$\$|\$[^$\n]+?\$", garde, md_texte, flags=re.S)
    corps = markdown.markdown(t, extensions=["tables", "fenced_code", "sane_lists"])
    corps = re.sub(r"@@MATH(\d+)@@", lambda m: html.escape(maths[int(m.group(1))], quote=False), corps)
    modele = """<!DOCTYPE html><html lang="fr"><head><meta charset="utf-8">
<title>__TITRE__</title>
<script>window.MathJax={tex:{inlineMath:[['$','$']],displayMath:[['$$','$$']]}};</script>
<script async src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>
<style>
body{font-family:Georgia,'Times New Roman',serif;max-width:820px;margin:2rem auto;padding:0 1.2rem;
line-height:1.65;color:#1F2937}
h1{color:#12306B;border-bottom:3px solid #12306B;padding-bottom:.3rem}
h2{color:#12306B;margin-top:2.2rem;border-bottom:1px solid #E5E7EB;padding-bottom:.2rem;
page-break-before:auto}
h3{color:#374151;margin-top:1.4rem}
table{border-collapse:collapse;margin:1rem 0}th,td{border:1px solid #D1D5DB;padding:.4rem .7rem}
th{background:#EAF0FA}code{background:#F3F4F6;padding:.1rem .3rem;border-radius:4px}
.note{color:#6B7280;font-size:.85rem;border-top:1px solid #E5E7EB;margin-top:3rem;padding-top:.6rem}
@media print{body{margin:0;max-width:none}.noprint{display:none}h2{page-break-before:always}}
</style></head><body>
<p class="noprint" style="background:#EAF0FA;padding:.6rem 1rem;border-radius:8px;font-family:sans-serif">
Pour obtenir un PDF : <b>Ctrl+P</b> (ou Partager → Imprimer) puis « Enregistrer au format PDF ».</p>
__CORPS__
<p class="note">Généré par Étudiants Learn · __DATE__</p></body></html>"""
    return (modele.replace("__TITRE__", html.escape(titre)).replace("__CORPS__", corps)
            .replace("__DATE__", maintenant()))


DPI_FORMULE = 300


def _nettoyer_latex(e: str) -> str:
    e = e.replace("\\dfrac", "\\frac").replace("\\tfrac", "\\frac").replace("\\displaystyle", "")
    e = re.sub(r"\\(?:text|operatorname|mathrm)\{", r"\\mathrm{", e)
    e = re.sub(r"\\(?:,|;|!|:|quad|qquad)", " ", e)
    e = e.replace("\\left.", "").replace("\\right.", "").replace("\\Bigg", "").replace("\\Big", "")
    return e.strip()


@lru_cache(maxsize=2048)
def formule_png(expr: str, taille: int):
    """Rend une formule LaTeX en PNG avec matplotlib. Retourne (octets, largeur_px, hauteur_px) ou None."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        from matplotlib.figure import Figure
        from PIL import Image
        matplotlib.rcParams["mathtext.fontset"] = "cm"
        fig = Figure(figsize=(0.1, 0.1), dpi=DPI_FORMULE)
        fig.text(0, 0, f"${_nettoyer_latex(expr)}$", fontsize=taille)
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=DPI_FORMULE, transparent=True,
                    bbox_inches="tight", pad_inches=0.03)
        data = buf.getvalue()
        largeur, hauteur = Image.open(io.BytesIO(data)).size
        return data, largeur, hauteur
    except Exception:
        return None


RE_INLINE = re.compile(r"(\$\$.+?\$\$|\$[^$\n]+?\$|\*\*.+?\*\*|(?<!\*)\*[^*\n]+?\*(?!\*)|`[^`\n]+`)", re.S)


# ---- Word ------------------------------------------------------------------
def _formule(par, expr, taille):
    rendu = formule_png(expr, taille)
    if rendu:
        par.add_run().add_picture(io.BytesIO(rendu[0]), width=Inches(rendu[1] / DPI_FORMULE))
    else:
        r = par.add_run(expr)
        r.italic = True


def _inline(par, texte):
    for m in RE_INLINE.split(texte):
        if not m:
            continue
        if m.startswith("$$"):
            _formule(par, m.strip("$").strip(), 12)
        elif m.startswith("$"):
            _formule(par, m.strip("$").strip(), 11)
        elif m.startswith("**"):
            par.add_run(m[2:-2]).bold = True
        elif m.startswith("`"):
            r = par.add_run(m[1:-1])
            r.font.name = "Consolas"
        elif m.startswith("*") and len(m) > 2:
            par.add_run(m[1:-1]).italic = True
        else:
            par.add_run(m)


def _md_dans_docx(doc, md_texte: str):
    t = re.sub(r"\$\$(.+?)\$\$", lambda m: "$$" + " ".join(m.group(1).split()) + "$$", md_texte, flags=re.S)
    lignes, i = t.splitlines(), 0
    while i < len(lignes):
        l = lignes[i].rstrip()
        s = l.strip()
        if not s:
            i += 1
            continue
        if s.startswith("|"):
            bloc = []
            while i < len(lignes) and lignes[i].strip().startswith("|"):
                bloc.append(lignes[i].strip())
                i += 1
            rows = [[c.strip() for c in b.strip("|").split("|")] for b in bloc
                    if not re.match(r"^\|?[\s:\-|]+\|?$", b)]
            if rows:
                n = max(len(r) for r in rows)
                tab = doc.add_table(rows=len(rows), cols=n)
                try:
                    tab.style = "Light Grid Accent 1"
                except Exception:
                    tab.style = "Table Grid"
                for a, r in enumerate(rows):
                    for b, c in enumerate(r):
                        cell = tab.cell(a, b)
                        cell.paragraphs[0].text = ""
                        _inline(cell.paragraphs[0], c)
                doc.add_paragraph()
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", s)
        if m:
            niveau = max(1, min(len(m.group(1)) - 1, 4))
            doc.add_heading(m.group(2).strip(" *"), level=niveau)
        elif s.startswith("$$") and s.endswith("$$") and len(s) > 4:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _formule(p, s[2:-2].strip(), 13)
        elif re.match(r"^[-*•]\s+", s):
            indent = len(l) - len(l.lstrip())
            p = doc.add_paragraph(style="List Bullet 2" if indent >= 2 else "List Bullet")
            _inline(p, re.sub(r"^[-*•]\s+", "", s))
        elif re.match(r"^\d+[.)]\s+", s):
            p = doc.add_paragraph(style="List Number")
            _inline(p, re.sub(r"^\d+[.)]\s+", "", s))
        else:
            _inline(doc.add_paragraph(), s)
        i += 1


def resume_en_docx(r: dict) -> bytes:
    doc = Document()
    for sec in doc.sections:
        sec.left_margin = sec.right_margin = Inches(0.9)
        sec.top_margin = sec.bottom_margin = Inches(0.8)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    for nom in ("Heading 1", "Heading 2", "Heading 3", "Heading 4"):
        try:
            doc.styles[nom].font.color.rgb = RGBColor(0x12, 0x30, 0x6B)
        except Exception:
            pass
    t = doc.add_heading(r["titre"], level=0)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph(f"Fiche de révision ({r.get('niveau', '')}) · {r.get('cree', '')}")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _md_dans_docx(doc, r["apercu"])
    for s in r["sections"]:
        doc.add_page_break()
        doc.add_heading(s["titre"], level=1)
        _md_dans_docx(doc, s["md"])
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ---- PDF (reportlab) ---------------------------------------------------------
_POLICES_PDF = {}


def _polices_pdf() -> dict:
    """Enregistre les polices DejaVu fournies avec matplotlib (accents, symboles mathématiques).
    Repli sur Helvetica/Courier si elles sont introuvables."""
    if _POLICES_PDF:
        return _POLICES_PDF
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    try:
        import matplotlib
        dossier = os.path.join(matplotlib.get_data_path(), "fonts", "ttf")
        fichiers = {"EL": "DejaVuSans.ttf", "EL-B": "DejaVuSans-Bold.ttf",
                    "EL-I": "DejaVuSans-Oblique.ttf", "EL-BI": "DejaVuSans-BoldOblique.ttf",
                    "EL-Mono": "DejaVuSansMono.ttf"}
        for nom, f in fichiers.items():
            pdfmetrics.registerFont(TTFont(nom, os.path.join(dossier, f)))
        pdfmetrics.registerFontFamily("EL", normal="EL", bold="EL-B", italic="EL-I", boldItalic="EL-BI")
        _POLICES_PDF.update(normal="EL", gras="EL-B", mono="EL-Mono")
    except Exception:
        _POLICES_PDF.update(normal="Helvetica", gras="Helvetica-Bold", mono="Courier")
    return _POLICES_PDF


def _pdf_texte(t: str) -> str:
    """Échappe le texte pour reportlab et retire les caractères que la police ne sait pas dessiner (emojis)."""
    t = re.sub("[\U00010000-\U0010FFFF\ufe0f\u200d\u26a0]", "", t)
    return html.escape(t, quote=False)


def _pdf_formule_tag(expr: str, taille: int, ctx: dict) -> str:
    """Balise <img> d'une formule pour un paragraphe reportlab (image PNG stockée dans un dossier temporaire)."""
    cle = (expr, taille)
    if cle in ctx["images"]:
        return ctx["images"][cle]
    rendu = formule_png(expr, taille)
    if not rendu:
        tag = f"<i>{_pdf_texte(expr)}</i>"
    else:
        data, wpx, hpx = rendu
        chemin = os.path.join(ctx["dossier"], f"f{len(ctx['images'])}.png")
        with open(chemin, "wb") as fh:
            fh.write(data)
        w, h = wpx / DPI_FORMULE * 72, hpx / DPI_FORMULE * 72
        tag = f'<img src="{chemin}" width="{w:.1f}" height="{h:.1f}" valign="middle"/>'
    ctx["images"][cle] = tag
    return tag


def _pdf_inline(texte: str, ctx: dict) -> str:
    out = []
    for m in RE_INLINE.split(texte):
        if not m:
            continue
        if m.startswith("$$"):
            out.append(_pdf_formule_tag(m.strip("$").strip(), 11, ctx))
        elif m.startswith("$"):
            out.append(_pdf_formule_tag(m.strip("$").strip(), 11, ctx))
        elif m.startswith("**") and len(m) > 4:
            out.append("<b>" + _pdf_inline(m[2:-2], ctx) + "</b>")
        elif m.startswith("`"):
            out.append(f'<font name="{_polices_pdf()["mono"]}">{_pdf_texte(m[1:-1])}</font>')
        elif m.startswith("*") and len(m) > 2:
            out.append("<i>" + _pdf_inline(m[1:-1], ctx) + "</i>")
        else:
            out.append(_pdf_texte(m))
    return "".join(out)


def _pdf_para(texte: str, style, ctx: dict, **kw):
    """Crée un paragraphe ; en cas de balisage invalide, retombe sur du texte brut."""
    from reportlab.platypus import Paragraph
    try:
        return Paragraph(_pdf_inline(texte, ctx), style, **kw)
    except Exception:
        brut = re.sub(r"[$*`]", "", texte)
        return Paragraph(_pdf_texte(brut), style, **kw)


def _pdf_styles() -> dict:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.styles import ParagraphStyle
    p = _polices_pdf()
    encre, accent = colors.HexColor("#1F2937"), colors.HexColor("#12306B")
    corps = ParagraphStyle("corps", fontName=p["normal"], fontSize=10, leading=15,
                           textColor=encre, spaceAfter=5)
    S = {
        "corps": corps,
        "puce": ParagraphStyle("puce", parent=corps, leftIndent=16, bulletIndent=4, spaceAfter=2),
        "puce2": ParagraphStyle("puce2", parent=corps, leftIndent=32, bulletIndent=20, spaceAfter=2),
        "cellule": ParagraphStyle("cellule", parent=corps, fontSize=9, leading=12, spaceAfter=0),
        "cellule_t": ParagraphStyle("cellule_t", parent=corps, fontSize=9, leading=12, spaceAfter=0,
                                    fontName=p["gras"]),
        "formule": ParagraphStyle("formule", parent=corps, alignment=TA_CENTER),
        "code": ParagraphStyle("code", fontName=p["mono"], fontSize=8, leading=10.5,
                               backColor=colors.HexColor("#F3F4F6"), leftIndent=6, rightIndent=6,
                               spaceBefore=4, spaceAfter=8),
        "titre": ParagraphStyle("titre", fontName=p["gras"], fontSize=22, leading=27, alignment=TA_CENTER,
                                textColor=accent, spaceAfter=6),
        "sous": ParagraphStyle("sous", parent=corps, alignment=TA_CENTER, textColor=colors.HexColor("#6B7280"),
                               fontSize=9, spaceAfter=10),
    }
    tailles = {1: (16, 21, 16), 2: (13, 18, 12), 3: (11.5, 16, 9), 4: (10.5, 14, 7)}
    for n, (fs, ld, avant) in tailles.items():
        S[f"h{n}"] = ParagraphStyle(f"h{n}", fontName=p["gras"], fontSize=fs, leading=ld, spaceBefore=avant,
                                    spaceAfter=5, keepWithNext=1,
                                    textColor=accent if n <= 2 else colors.HexColor("#374151"))
    return S


def _md_vers_pdf(md_texte: str, ctx: dict, S: dict, largeur: float) -> list:
    from reportlab.lib import colors
    from reportlab.platypus import HRFlowable, Image, Paragraph, Preformatted, Spacer, Table, TableStyle
    t = re.sub(r"\$\$(.+?)\$\$", lambda m: "$$" + " ".join(m.group(1).split()) + "$$", md_texte, flags=re.S)
    lignes, i, out = t.splitlines(), 0, []
    while i < len(lignes):
        l = lignes[i].rstrip()
        s = l.strip()
        if not s:
            i += 1
            continue
        # Bloc de code
        if s.startswith("```"):
            i += 1
            code = []
            while i < len(lignes) and not lignes[i].strip().startswith("```"):
                code.append(lignes[i])
                i += 1
            i += 1
            out.append(Preformatted(re.sub("[\U00010000-\U0010FFFF]", "", "\n".join(code)), S["code"]))
            continue
        # Tableau
        if s.startswith("|"):
            bloc = []
            while i < len(lignes) and lignes[i].strip().startswith("|"):
                bloc.append(lignes[i].strip())
                i += 1
            rows = [[c.strip() for c in b.strip("|").split("|")] for b in bloc
                    if not re.match(r"^\|?[\s:\-|]+\|?$", b)]
            if rows:
                n = max(len(r) for r in rows)
                data = []
                for a, r in enumerate(rows):
                    r = r + [""] * (n - len(r))
                    data.append([_pdf_para(c, S["cellule_t"] if a == 0 else S["cellule"], ctx) for c in r])
                tab = Table(data, colWidths=[largeur / n] * n, repeatRows=1, hAlign="LEFT")
                tab.setStyle(TableStyle([
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D1D5DB")),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAF0FA")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5)]))
                out += [tab, Spacer(1, 8)]
            continue
        # Titres
        m = re.match(r"^(#{1,6})\s+(.*)", s)
        if m:
            niveau = max(1, min(len(m.group(1)) - 1, 4))
            out.append(_pdf_para(m.group(2).strip(" *"), S[f"h{niveau}"], ctx))
        # Filet horizontal
        elif re.match(r"^(-{3,}|\*{3,}|_{3,})$", s):
            out.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#E5E7EB"),
                                  spaceBefore=4, spaceAfter=6))
        # Formule seule sur sa ligne
        elif s.startswith("$$") and s.endswith("$$") and len(s) > 4:
            expr = s[2:-2].strip()
            rendu = formule_png(expr, 13)
            if rendu:
                data, wpx, hpx = rendu
                w, h = wpx / DPI_FORMULE * 72, hpx / DPI_FORMULE * 72
                if w > largeur:
                    h, w = h * largeur / w, largeur
                img = Image(io.BytesIO(data), width=w, height=h)
                img.hAlign = "CENTER"
                out += [Spacer(1, 3), img, Spacer(1, 6)]
            else:
                out.append(Paragraph(f"<i>{_pdf_texte(expr)}</i>", S["formule"]))
        # Listes
        elif re.match(r"^[-*•]\s+", s):
            indent = len(l) - len(l.lstrip())
            out.append(_pdf_para(re.sub(r"^[-*•]\s+", "", s), S["puce2" if indent >= 2 else "puce"], ctx,
                                 bulletText="–" if indent >= 2 else "•"))
        elif re.match(r"^\d+[.)]\s+", s):
            num = re.match(r"^(\d+)[.)]\s+", s).group(1)
            out.append(_pdf_para(re.sub(r"^\d+[.)]\s+", "", s), S["puce"], ctx, bulletText=f"{num}."))
        else:
            out.append(_pdf_para(s, S["corps"], ctx))
        i += 1
    return out


def resume_en_pdf(r: dict) -> bytes:
    """Génère le résumé au format PDF (A4, formules en images, pagination, un chapitre par page)."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.platypus import HRFlowable, PageBreak, SimpleDocTemplate, Spacer

    p = _polices_pdf()
    S = _pdf_styles()
    marge = 2 * cm
    largeur = A4[0] - 2 * marge
    titre_court = r["titre"][:70]

    def pied_page(canvas, doc):
        canvas.saveState()
        canvas.setFont(p["normal"], 8)
        canvas.setFillColor(colors.HexColor("#9CA3AF"))
        canvas.drawString(marge, 1.1 * cm, f"Étudiants Learn · {titre_court}")
        canvas.drawRightString(A4[0] - marge, 1.1 * cm, f"Page {doc.page}")
        canvas.restoreState()

    buf = io.BytesIO()
    with tempfile.TemporaryDirectory() as dossier:
        ctx = {"dossier": dossier, "images": {}}
        doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=marge, rightMargin=marge,
                                topMargin=marge, bottomMargin=1.8 * cm,
                                title=r["titre"], author=AUTEUR, subject="Fiche de révision")
        story = [_pdf_para(r["titre"], S["titre"], ctx),
                 _pdf_para(f"Fiche de révision ({r.get('niveau', '')}) · {r.get('cree', '')}", S["sous"], ctx),
                 HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#12306B"), spaceAfter=10)]
        story += _md_vers_pdf(r["apercu"], ctx, S, largeur)
        for s in r["sections"]:
            story += [PageBreak(), _pdf_para(s["titre"], S["h1"], ctx),
                      HRFlowable(width="100%", thickness=0.6, color=colors.HexColor("#E5E7EB"), spaceAfter=6)]
            story += _md_vers_pdf(s["md"], ctx, S, largeur)
        doc.build(story, onFirstPage=pied_page, onLaterPages=pied_page)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# QCM
# ----------------------------------------------------------------------------
PATRON_QCM = re.compile(
    r"Question\s*:\s*(.+?)\n\s*A\s*[\).]\s*(.+?)\n\s*B\s*[\).]\s*(.+?)\n\s*C\s*[\).]\s*(.+?)\n"
    r"\s*D\s*[\).]\s*(.+?)\n\s*R[ée]ponse\s*:\s*\**\s*([ABCD])\b\**[ \t]*(?:\n\s*Explication\s*:\s*(.+))?",
    re.S | re.I)


def parser_qcm(brut: str, chap: str) -> list:
    out = []
    for bloc in brut.split("@@Q")[1:]:
        m = PATRON_QCM.search(bloc.strip() + "\n")
        if not m:
            continue
        opts = [normaliser_math(x.strip()) for x in m.groups()[1:5]]
        bonne = LETTRES.index(m.group(6).upper())
        ordre = list(range(4))
        random.shuffle(ordre)  # évite le biais de position de la bonne réponse
        out.append({"q": normaliser_math(m.group(1).strip()),
                    "o": [opts[k] for k in ordre], "r": ordre.index(bonne),
                    "e": normaliser_math((m.group(7) or "").strip()), "chap": chap})
    return out


DIFFICULTES = {
    "Facile": "facile (restitution des définitions et résultats de base)",
    "Moyen": "moyenne (compréhension et application directe)",
    "Difficile": "difficile (raisonnement, calculs, cas pièges)",
    "Mixte": "variée (environ 30 % faciles, 40 % moyennes, 30 % difficiles)",
}


def generer_qcm(cours, cibles, nb, difficulte, calculs, melanger, cle, progres=None) -> list:
    modeles_disponibles(cle)  # préchauffe le cache (avant les threads)
    # Répartition équilibrée des questions entre les chapitres ciblés
    base, reste = divmod(nb, len(cibles))
    bonus = set(random.sample(range(len(cibles)), reste))
    taches = [(c, base + (1 if i in bonus else 0)) for i, c in enumerate(cibles)]
    taches = [(c, n) for c, n in taches if n > 0]

    # Regroupe les petits chapitres : moins d'appels = plus rapide
    groupes, cour, somme = [], [], 0
    for c, n in taches:
        cour.append((c, n))
        somme += n
        if somme >= 6:
            groupes.append(cour)
            cour, somme = [], 0
    if cour:
        if groupes:
            groupes[-1] += cour
        else:
            groupes.append(cour)
    jobs = []
    for g in groupes:
        n = sum(x[1] for x in g)
        if len(g) == 1:
            jobs.append((g[0][0]["titre"], portee_texte(g[0][0]), n))
        else:
            noms = [x[0]["titre"] for x in g]
            jobs.append((" · ".join(noms)[:90],
                         "les chapitres suivants, en répartissant équitablement les questions : "
                         + " ; ".join(f"« {x} »" for x in noms), n))

    def produire(label, portee, n):
        recues, tentatives, derniere = [], 0, None
        while len(recues) < n and tentatives < 3:
            tentatives += 1
            lot = min(LOT_QCM, n - len(recues))
            evite = "; ".join(q["q"][:70].replace("\n", " ") for q in recues[-40:]) or "aucune pour l'instant"
            prompt = f"""Tu es un professeur expérimenté qui prépare un examen. À partir du cours fourni,
rédige {lot} questions à choix multiples portant UNIQUEMENT sur {portee}.
Règles :
- Questions fidèles au contenu du cours, précises et autonomes ; 4 options plausibles, UNE seule bonne réponse.
- Difficulté : {DIFFICULTES[difficulte]}.
- Varie les types : définition, compréhension, application, piège classique.
{"- Inclus des questions de calcul / d'application numérique avec le calcul justifié dans l'explication." if calculs else ""}
- L'explication justifie la bonne réponse (et le calcul si besoin) en 1 à 3 phrases.
- Ne reprends pas ces questions déjà posées : {evite}
- {REGLE_MATHS}
FORMAT STRICT (aucun JSON, aucun autre texte), répété pour chaque question :
@@Q
Question: ...
A) ...
B) ...
C) ...
D) ...
Reponse: une seule lettre parmi A, B, C, D
Explication: ..."""
            try:
                brut = generer(requete(cours, prompt), cle, max_tokens=12000,
                               temperature=0.7, reflexion=False)
            except Exception as e:
                derniere = e
                continue
            recues += parser_qcm(brut, label)[:lot]
        if not recues and derniere:
            raise derniere
        return recues

    par_job, erreurs, fait = {}, [], 0
    with ThreadPoolExecutor(max_workers=min(4, len(jobs))) as ex:
        futs = {ex.submit(produire, *j): k for k, j in enumerate(jobs)}
        for f in as_completed(futs):
            k = futs[f]
            try:
                par_job[k] = f.result()
            except Exception as e:
                erreurs.append(e)
                par_job[k] = []
            fait += jobs[k][2]
            if progres:
                progres(min(fait / nb, 1.0),
                        f"{sum(len(v) for v in par_job.values())} / {nb} questions générées")
    if erreurs and not any(par_job.values()):
        raise erreurs[0]

    deja, resultat = set(), []
    for k in range(len(jobs)):  # on garde l'ordre du cours
        for q in par_job.get(k, []):
            cle_q = re.sub(r"\W+", "", q["q"].lower())[:80]
            if cle_q not in deja:
                deja.add(cle_q)
                resultat.append(q)
    if melanger:
        random.shuffle(resultat)
    return resultat


def sauver_qcm(titre, questions) -> int:
    c = st.session_state.get("cours") or {}
    i = nouvel_id()
    stock()["qcms"][i] = {"id": i, "cours_id": c.get("id"), "titre": titre,
                          "questions": questions, "cree": maintenant()}
    return i


# ----------------------------------------------------------------------------
# Interface : composants communs
# ----------------------------------------------------------------------------
def entete(titre, sous=""):
    cours = st.session_state.get("cours")
    chip = (f'<div class="chip">Cours actif : {html.escape(cours["nom"])}</div>'
            if cours and titre != "Services & contact" else "")
    st.markdown(f'<div class="banniere"><div class="fil">Accueil &nbsp;›&nbsp; {titre}</div>'
                f'<h1>{titre}</h1><p>{sous}</p>{chip}</div>', unsafe_allow_html=True)


def aller(page: str):
    """Change de page (utilisé par les boutons de l'accueil)."""
    st.session_state.nav = page


def cle_api(afficher=True):
    cle = os.getenv("GEMINI_API_KEY")
    if not cle:
        try:
            cle = st.secrets.get("GEMINI_API_KEY")
        except Exception:
            cle = None
    if not cle and afficher:
        with st.expander("Configuration : clé API Gemini requise", expanded=True):
            cle = st.text_input("Clé API Gemini", type="password",
                                help="Ou définis la variable GEMINI_API_KEY (secrets Streamlit).")
    return cle


def pied():
    liens = "".join(f"<li>{p}</li>" for p in PAGES)
    st.markdown(
        f'''<div class="pied-site"><div class="cols">
        <div><h5>Étudiants Learn</h5>
          <p>Plateforme d'aide à la réussite : résumés de cours, QCM corrigés et assistant pédagogique,
          complétés par un accompagnement en traitement de données et en rédaction de mémoire.</p></div>
        <div><h5>Navigation</h5><ul>{liens}</ul></div>
        <div><h5>Contact</h5><ul>
          <li><a href="mailto:{EMAIL}">{EMAIL}</a></li>
          <li><a href="tel:{TEL_AMB[1]}">{TEL_AMB[0]}</a></li>
          <li><a href="tel:{TEL_REP[1]}">{TEL_REP[0]}</a></li></ul></div>
        </div>
        <div class="copy">© {datetime.now().year} Étudiants Learn · Conçu par {AUTEUR}</div></div>''',
        unsafe_allow_html=True)


def haut_de_page() -> str:
    """Barre d'information, identité du site et menu principal."""
    st.session_state.setdefault("nav", PAGES[0])
    with st.container(key="top"):
        st.markdown(
            f'''<div class="topbar"><span>Plateforme d'aide à la réussite universitaire</span>
            <span><a href="mailto:{EMAIL}">{EMAIL}</a><span class="sep">|</span>
            <a href="tel:{TEL_AMB[1]}">{TEL_AMB[0]}</a></span></div>
            <div class="entete-site"><span class="logo">EL</span><div>
            <div class="nom">Étudiants Learn</div>
            <div class="devise">Réviser · Comprendre · Réussir</div></div></div>''',
            unsafe_allow_html=True)
        with st.container(key="navbar"):
            st.radio("Navigation", PAGES, horizontal=True, label_visibility="collapsed", key="nav")
    return st.session_state.nav


def boutons_contact(numero, message="Bonjour, je vous contacte depuis Étudiants Learn."):
    from urllib.parse import quote
    return (f'<a class="btn" href="tel:{numero[1]}">Appeler {numero[0]}</a>'
            f'<a class="btn wa" target="_blank" href="https://wa.me/{numero[1].lstrip("+")}?text={quote(message)}">'
            f'WhatsApp {numero[0]}</a>')


ICONES = {
    "graphique": '<svg viewBox="0 0 24 24"><path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/></svg>',
    "memoire": '<svg viewBox="0 0 24 24"><path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4z"/></svg>',
}

SERVICES = [
    {"icone": "graphique", "tag": "Données", "titre": "Traitement et analyse de données",
     "texte": "Des données propres, correctement analysées et clairement présentées, "
              "pour votre mémoire, vos travaux ou votre activité.",
     "points": ["Saisie, nettoyage et codage des données", "Analyses statistiques et calculs",
                "Tableaux, graphiques et mise en forme", "Interprétation et présentation des résultats"],
     "contact": [TEL_AMB[0], TEL_REP[0]]},
    {"icone": "memoire", "tag": "Académique", "titre": "Accompagnement à la rédaction de mémoire",
     "texte": "Un accompagnement méthodologique, de l'idée de départ jusqu'à la soutenance.",
     "points": ["Choix du sujet, problématique et plan", "Revue de littérature et méthodologie",
                "Rédaction, mise en forme et références", "Préparation à la soutenance"],
     "contact": [TEL_AMB[0], TEL_REP[0]]},
]


def bloc_services():
    cartes = ""
    for s in SERVICES:
        pts = "".join(f"<li>{p}</li>" for p in s["points"])
        tels = " · ".join(s["contact"])
        cartes += (f'<div class="service"><div class="ico">{ICONES[s["icone"]]}</div>'
                   f'<div class="tag">{s["tag"]}</div>'
                   f'<h4>{s["titre"]}</h4><p>{s["texte"]}</p><ul>{pts}</ul>'
                   f'<div class="contact">Contact : <b>{tels}</b></div></div>')
    st.markdown(f'<div class="services">{cartes}</div>', unsafe_allow_html=True)


def page_services():
    entete("Services & contact",
           "Au-delà de l'application, nous vous accompagnons concrètement dans l'analyse de vos données "
           "et la rédaction de votre mémoire.")
    bloc_services()

    st.markdown('<div class="section-titre">Comment ça se passe</div>', unsafe_allow_html=True)
    st.markdown(
        '''<div class="parcours">
        <div><b>1. Vous nous contactez</b><span>Par appel, WhatsApp ou e-mail, en décrivant votre besoin.</span></div>
        <div><b>2. Nous cadrons la demande</b><span>Nous précisons ensemble les objectifs, le délai et les modalités.</span></div>
        <div><b>3. Nous vous accompagnons</b><span>Nous réalisons la mission et restons disponibles jusqu'à son terme.</span></div>
        </div>''', unsafe_allow_html=True)

    st.markdown('<div class="section-titre">Nous contacter</div>', unsafe_allow_html=True)
    st.markdown("**Traitement et analyse de données**")
    msg = "Bonjour, je souhaite être accompagné(e) pour le traitement et l'analyse de mes données."
    st.markdown(boutons_contact(TEL_AMB, msg) + boutons_contact(TEL_REP, msg), unsafe_allow_html=True)
    st.markdown("**Accompagnement à la rédaction de mémoire**")
    msg = "Bonjour, je souhaite être accompagné(e) pour la rédaction de mon mémoire."
    st.markdown(boutons_contact(TEL_AMB, msg) + boutons_contact(TEL_REP, msg), unsafe_allow_html=True)
    st.markdown("**Par e-mail**")
    st.markdown(f'<a class="btn clair" href="mailto:{EMAIL}">{EMAIL}</a>', unsafe_allow_html=True)
    st.caption("Accompagnement méthodologique : le mémoire reste votre travail personnel.")


def etapes():
    ok = [bool(st.session_state.get("cours")), bool(st.session_state.get("resume")),
          bool(st.session_state.get("qcm")), bool(st.session_state.get("chat"))]
    noms = ["Charger le cours", "Résumé", "QCM", "Assistant"]
    st.markdown('<div class="etapes">' + "".join(
        f'<span class="etape {"ok" if o else ""}"><i>{"✓" if o else k}</i>{n}</span>'
        for k, (o, n) in enumerate(zip(ok, noms), 1)) + "</div>", unsafe_allow_html=True)


def besoin_cours():
    if not st.session_state.get("cours"):
        st.info("Commence par charger un cours dans l'onglet **Mon cours**.")
        return None
    return st.session_state.cours


def reset_travail(cours):
    st.session_state.cours = cours
    st.session_state.resume = None
    st.session_state.qcm = None
    st.session_state.valide = False
    st.session_state.chat = charger_chat(cours.get("id"))


# ----------------------------------------------------------------------------
# Pages
# ----------------------------------------------------------------------------
def page_accueil():
    st.markdown(
        '''<div class="hero"><div class="kicker">Plateforme d'aide à la réussite</div>
        <h1>Réussissez vos études à partir de vos propres cours</h1>
        <p>Importez un cours au format PDF ou Word : obtenez une fiche de révision approfondie, des QCM
        corrigés et un assistant qui répond à vos questions. Sans inscription.</p></div>''',
        unsafe_allow_html=True)
    b1, b2, _ = st.columns([1.2, 1.2, 3])
    b1.button("Charger mon cours", type="primary", use_container_width=True, on_click=aller, args=("Mon cours",))
    b2.button("Nos services", use_container_width=True, on_click=aller, args=("Services & contact",))

    st.markdown(
        f'''<div class="chiffres">
        <div><b>{MAX_QCM}</b><span>questions par QCM</span></div>
        <div><b>{len(NIVEAUX)}</b><span>niveaux de résumé</span></div>
        <div><b>4</b><span>formats d'export</span></div>
        <div><b>24h/24</b><span>assistant disponible</span></div></div>''',
        unsafe_allow_html=True)

    st.markdown('<div class="section-titre">Les outils de révision</div>', unsafe_allow_html=True)
    st.markdown(
        '''<div class="cartes">
        <div class="carte"><div class="tag">Résumé</div><h4>Fiches approfondies</h4>
          <p>Une fiche par chapitre : définitions, formules, démonstrations, exemples résolus et pièges.</p></div>
        <div class="carte"><div class="tag">Évaluation</div><h4>QCM corrigés</h4>
          <p>Sur tout le cours ou par chapitre, avec correction détaillée, note sur 20 et révision des erreurs.</p></div>
        <div class="carte"><div class="tag">Tutorat</div><h4>Assistant pédagogique</h4>
          <p>Explications pas à pas, exercices corrigés et conseils pour préparer vos examens.</p></div>
        <div class="carte"><div class="tag">Export</div><h4>PDF, Word, HTML, Markdown</h4>
          <p>Vos fiches téléchargeables, formules mathématiques comprises, pour réviser hors ligne.</p></div>
        </div>''', unsafe_allow_html=True)

    st.markdown('<div class="section-titre">Comment ça marche</div>', unsafe_allow_html=True)
    st.markdown(
        '''<div class="parcours">
        <div><b>1. Déposez votre cours</b><span>PDF ou document Word : il est lu en entier et découpé en chapitres.</span></div>
        <div><b>2. Choisissez votre outil</b><span>Résumé, QCM ou questions à l'assistant, chapitre par chapitre.</span></div>
        <div><b>3. Révisez et exportez</b><span>Téléchargez vos fiches et refaites vos QCM jusqu'à maîtriser le sujet.</span></div>
        </div>''', unsafe_allow_html=True)

    st.markdown('<div class="section-titre">Accompagnement académique</div>', unsafe_allow_html=True)
    bloc_services()


def page_cours(cle):
    entete("Mon cours", "Déposez un PDF ou un document Word : le cours est lu en entier et découpé en chapitres.")
    etapes()
    fichier = st.file_uploader("Fichier du cours", type=["pdf", "docx"])
    if st.button("Analyser le cours", type="primary", disabled=not fichier):
        if not cle:
            st.error("Ajoute ta clé API Gemini.")
        else:
            try:
                with st.status("Analyse du cours…", expanded=True) as s:
                    st.write("Lecture du fichier…")
                    cours = lire_fichier(fichier)
                    if len(cours["texte"]) < 80 and not cours["pdf"]:
                        raise ValueError("Aucun texte lisible (document vide ou PDF scanné trop volumineux).")
                    if len(cours["texte"]) > MAX_CHARS:
                        st.write("Attention : cours très long, seule la partie initiale sera exploitée en mode texte.")
                    st.write("Détection des chapitres…")
                    cours["chapitres"] = detecter_chapitres(cours, cle)
                    enregistrer_cours(cours)
                    reset_travail(cours)
                    s.update(label=f"Cours prêt — {len(cours['chapitres'])} chapitre(s)", state="complete")
            except Exception as e:
                st.error(f"Erreur : {e}")

    cours = st.session_state.get("cours")
    if cours:
        st.subheader(cours["nom"])
        c1, c2, c3 = st.columns(3)
        c1.metric("Pages", cours.get("pages") or "—")
        c2.metric("Mots", f"{len(cours['texte'].split()):,}".replace(",", " "))
        c3.metric("Chapitres", len(cours["chapitres"]))
        if cours.get("pdf"):
            cours["vision"] = st.toggle(
                "Lecture visuelle du PDF (formules, schémas, scans) — recommandé",
                value=cours.get("vision", True),
                help="Gemini lit directement les pages du PDF. Désactive pour un mode texte plus rapide.")
        with st.expander("Structure détectée", expanded=True):
            for i, c in enumerate(cours["chapitres"], 1):
                st.markdown(f"**{i}. {c['titre']}**" + (f" — {c['desc']}" if c["desc"] else ""))

    lst = sorted(stock()["cours"].values(), key=lambda c: c["id"], reverse=True)
    if lst:
        st.divider()
        st.subheader("Mes cours (cette session)")
        for c in lst:
            with st.container(border=True):
                a, b, d = st.columns([5, 1.3, 1.5])
                a.markdown(f"**{html.escape(c['nom'])}**  \n<small>{c['cree']}</small>", unsafe_allow_html=True)
                if b.button("Ouvrir", key=f"o{c['id']}", use_container_width=True):
                    reset_travail(charger_cours(c["id"]))
                    st.rerun()
                if d.button("Supprimer", key=f"s{c['id']}", use_container_width=True):
                    supprimer_cours(c["id"])
                    if (st.session_state.get("cours") or {}).get("id") == c["id"]:
                        st.session_state.cours = None
                    st.rerun()


@st.cache_data(show_spinner=False, max_entries=6)
def html_export(md: str, titre: str) -> bytes:
    return md_vers_html(md, titre).encode("utf-8")


def afficher_resume(r: dict, cle_export: str):
    st.header(r["titre"])
    st.caption(f"Niveau : {r.get('niveau', '')} · généré le {r.get('cree', '')}")
    md = resume_en_markdown(r)

    # Exports lourds (PDF, Word) : générés à la demande, puis proposés au téléchargement
    cp, cw = st.columns(2)
    with cp:
        prete = st.session_state.get("pdf")
        if prete and prete[0] == cle_export:
            st.download_button("Télécharger le PDF", prete[1], file_name=f"{slug(r['titre'])}.pdf",
                               mime="application/pdf", type="primary", use_container_width=True,
                               key=f"dpdf{cle_export}")
        elif st.button("Préparer le PDF", type="primary", use_container_width=True, key=f"ppdf{cle_export}"):
            try:
                with st.spinner("Création du PDF (rendu des formules)…"):
                    st.session_state.pdf = (cle_export, resume_en_pdf(r))
                st.rerun()
            except Exception as e:
                st.error(f"Impossible de créer le PDF : {e}")
    with cw:
        prete = st.session_state.get("docx")
        if prete and prete[0] == cle_export:
            st.download_button("Télécharger le fichier Word", prete[1], file_name=f"{slug(r['titre'])}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True, key=f"dw{cle_export}")
        elif st.button("Préparer le fichier Word", use_container_width=True, key=f"pw{cle_export}"):
            try:
                with st.spinner("Création du document Word (rendu des formules)…"):
                    st.session_state.docx = (cle_export, resume_en_docx(r))
                st.rerun()
            except Exception as e:
                st.error(f"Impossible de créer le document Word : {e}")

    c3, c4 = st.columns(2)
    c3.download_button("Télécharger en HTML", html_export(md, r["titre"]),
                       file_name=f"{slug(r['titre'])}.html", mime="text/html", use_container_width=True,
                       help="Version web avec formules MathJax.")
    c4.download_button("Télécharger en Markdown", md.encode("utf-8"), file_name=f"{slug(r['titre'])}.md",
                       mime="text/markdown", use_container_width=True)
    st.divider()
    st.markdown(r["apercu"])
    st.divider()
    st.subheader("Fiches par chapitre")
    noms = [f"{i + 1}. {sec['titre']}" for i, sec in enumerate(r["sections"])]
    choix_chap = st.selectbox("Chapitre à afficher", noms, key=f"vue_{cle_export}")
    st.markdown(r["sections"][noms.index(choix_chap)]["md"])


def page_resume(cle):
    entete("Résumé", "Une fiche de révision complète, chapitre par chapitre, avec formules et exemples.")
    cours = besoin_cours()
    if not cours:
        return
    titres = [c["titre"] for c in cours["chapitres"]]
    choix = st.multiselect("Chapitres à résumer", titres, default=titres)
    niveau = st.select_slider("Niveau de détail", list(NIVEAUX), value="Approfondi")
    if st.button("Générer le résumé", type="primary", disabled=not choix):
        if not cle:
            st.error("Ajoute ta clé API Gemini.")
        else:
            cibles = [c for c in cours["chapitres"] if c["titre"] in choix]
            barre = st.progress(0.0, text="Démarrage…")
            try:
                r = generer_resume(cours, cibles, niveau, cle,
                                   lambda p, t: barre.progress(min(p, 1.0), text=t))
                r["id"] = nouvel_id()
                r["cours_id"] = cours.get("id")
                stock()["resumes"][r["id"]] = r
                st.session_state.resume = r
                barre.empty()
                if r["echecs"]:
                    st.warning(f"{r['echecs']} section(s) non générée(s) : relance pour les compléter.")
            except Exception as e:
                barre.empty()
                st.error(f"Erreur : {e}")
    r = st.session_state.get("resume")
    if r:
        st.divider()
        afficher_resume(r, str(r.get("id") or r["cree"]))


def page_qcm(cle):
    entete("QCM", "Jusqu'à 100 questions, sur tout le cours ou chapitre par chapitre, avec correction détaillée.")
    cours = besoin_cours()
    if not cours:
        return
    titres = [c["titre"] for c in cours["chapitres"]]

    with st.expander("Paramètres du QCM", expanded=not st.session_state.get("qcm")):
        mode = st.radio("Portée", ["Cours entier (tous les chapitres, équilibré)", "Par chapitre(s)"],
                        horizontal=True)
        if mode.startswith("Par"):
            sel = st.multiselect("Chapitres", titres, default=titres[:1])
        else:
            sel = titres
        c1, c2 = st.columns(2)
        nb = c1.slider("Nombre de questions", 5, MAX_QCM, 20, step=5)
        diff = c2.select_slider("Difficulté", list(DIFFICULTES), value="Mixte")
        c3, c4 = st.columns(2)
        calculs = c3.toggle("Inclure des questions de calcul / d'application", value=True)
        melange = c4.toggle("Mélanger l'ordre des questions", value=False)
        go = st.button(f"Générer {nb} questions", type="primary", disabled=not sel)

    if go:
        if not cle:
            st.error("Ajoute ta clé API Gemini.")
        else:
            cibles = [c for c in cours["chapitres"] if c["titre"] in sel]
            barre = st.progress(0.0, text="Démarrage…")
            try:
                qs = generer_qcm(cours, cibles, nb, diff, calculs, melange, cle,
                                 lambda p, t: barre.progress(p, text=t))
                barre.empty()
                if not qs:
                    st.error("Aucune question valide générée. Réessaie.")
                else:
                    if len(qs) < nb:
                        st.warning(f"{len(qs)} questions valides générées sur {nb} demandées.")
                    portee = "cours entier" if len(sel) == len(titres) else ", ".join(sel)[:80]
                    titre = f"QCM {len(qs)} q. — {portee}"
                    st.session_state.qcm = {"id": sauver_qcm(titre, qs), "titre": titre, "questions": qs}
                    st.session_state.valide = False
                    st.session_state.version = st.session_state.get("version", 0) + 1
                    st.rerun()
            except Exception as e:
                barre.empty()
                st.error(f"Erreur : {e}")

    qcm = st.session_state.get("qcm")
    if not qcm:
        st.info("Règle les paramètres ci-dessus puis clique sur « Générer ».")
        return
    qs, v = qcm["questions"], st.session_state.get("version", 0)
    st.subheader(qcm["titre"])

    if not st.session_state.get("valide"):
        st.caption(f"{len(qs)} questions · une seule bonne réponse par question.")
        with st.form(f"form_{v}"):
            for i, q in enumerate(qs):
                st.markdown(f"**{i + 1}.** {q['q']}")
                if q.get("chap") and q["chap"] != "Cours complet":
                    st.caption(f"Chapitre : {q['chap']}")
                st.radio("Réponse", range(4), format_func=lambda j, q=q: f"{LETTRES[j]}. {q['o'][j]}",
                         index=None, key=f"q{v}_{i}", label_visibility="collapsed")
                st.divider()
            ok = st.form_submit_button("Valider et voir la correction", type="primary")
        if ok:
            reps = [st.session_state.get(f"q{v}_{i}") for i in range(len(qs))]
            st.session_state.reponses = reps
            st.session_state.valide = True
            if qcm.get("id"):
                stock()["tentatives"].append(
                    {"qcm_id": qcm["id"], "score": sum(r == q["r"] for r, q in zip(reps, qs)),
                     "total": len(qs), "cree": maintenant()})
            st.rerun()
        return

    # Correction
    reps = st.session_state.get("reponses", [None] * len(qs))
    score = sum(r == q["r"] for r, q in zip(reps, qs))
    note = round(score / len(qs) * 20, 1)
    m1, m2, m3 = st.columns(3)
    m1.metric("Ta note", f"{note} / 20")
    m2.metric("Bonnes réponses", f"{score} / {len(qs)}")
    m3.metric("Sans réponse", sum(r is None for r in reps))

    par_chap = {}
    for r, q in zip(reps, qs):
        a = par_chap.setdefault(q.get("chap", ""), [0, 0])
        a[0] += r == q["r"]
        a[1] += 1
    if len(par_chap) > 1:
        st.markdown("**Résultats par chapitre**")
        for ch, (b, t) in par_chap.items():
            st.progress(b / t, text=f"{ch} — {b}/{t}")
    st.divider()

    for i, (r, q) in enumerate(zip(reps, qs)):
        st.markdown(f"**{i + 1}.** {q['q']}")
        if r == q["r"]:
            st.success(f"**Correct** : {LETTRES[r]}. {q['o'][r]}\n\n{q['e']}")
        else:
            ta = "**Sans réponse**" if r is None else f"**Ta réponse** : {LETTRES[r]}. {q['o'][r]}"
            st.error(f"{ta}\n\n**Bonne réponse : {LETTRES[q['r']]}. {q['o'][q['r']]}**\n\n{q['e']}")

    erreurs = [q for r, q in zip(reps, qs) if r != q["r"]]
    b1, b2 = st.columns(2)
    if b1.button("Recommencer", use_container_width=True):
        st.session_state.valide = False
        st.session_state.version += 1
        st.rerun()
    if erreurs and b2.button(f"Refaire mes {len(erreurs)} erreur(s)", use_container_width=True):
        st.session_state.qcm = {"id": None, "titre": "Révision de mes erreurs", "questions": erreurs}
        st.session_state.valide = False
        st.session_state.version += 1
        st.rerun()


SYSTEME_CHAT = """Tu es Étudiants Learn, un tuteur bienveillant, précis et pédagogue. L'étudiant t'a fourni son cours.
- Réponds en français, en t'appuyant d'abord sur le contenu du cours ; si tu sors du cours, dis-le clairement.
- Explique étape par étape, donne des exemples, et propose des exercices corrigés si on te le demande.
- Pour les maths, utilise LaTeX : $...$ en ligne et $$...$$ en bloc (jamais \\( \\) ni \\[ \\]).
- Sois concis quand la question est simple, détaillé quand elle est complexe."""


def page_chat(cle):
    entete("Assistant", "Posez vos questions sur votre cours : explications, exercices, méthodes, révisions.")
    cours = besoin_cours()
    if not cours:
        return
    msgs = st.session_state.setdefault("chat", [])
    if not msgs:
        st.caption("Suggestions pour démarrer")
        sug = {"Notions essentielles": "Explique-moi simplement les notions les plus importantes du cours.",
               "Exercices corrigés": "Donne-moi 3 exercices corrigés pas à pas sur ce cours.",
               "Erreurs à éviter": "Quelles sont les erreurs classiques à éviter à l'examen ?"}
        cols = st.columns(len(sug))
        for col, (libelle, question_type) in zip(cols, sug.items()):
            if col.button(libelle, use_container_width=True, key=f"sug_{libelle}"):
                st.session_state.question_en_attente = question_type
                st.rerun()
    for m in msgs:
        with st.chat_message("user" if m["role"] == "user" else "assistant"):
            st.markdown(m["content"])

    question = st.chat_input("Pose ta question sur le cours…") or st.session_state.pop("question_en_attente", None)
    if question:
        if not cle:
            st.error("Ajoute ta clé API Gemini.")
            return
        msgs.append({"role": "user", "content": question})
        with st.chat_message("user"):
            st.markdown(question)
        histo = msgs[-20:]
        contenu = [
            types.Content(role="user", parts=parts_cours(cours) + [
                types.Part.from_text(text="Voici mon cours. Sers-t'en comme base pour répondre à mes questions.")]),
            types.Content(role="model", parts=[types.Part.from_text(text="Compris, je suis prêt à t'aider.")]),
        ] + [types.Content(role="user" if m["role"] == "user" else "model",
                           parts=[types.Part.from_text(text=m["content"])]) for m in histo]
        with st.chat_message("assistant"):
            try:
                rep = st.write_stream(flux_chat(contenu, cle, SYSTEME_CHAT))
                msgs.append({"role": "assistant", "content": normaliser_math(rep)})
                sauver_chat(cours.get("id"), msgs)
                st.rerun()
            except Exception as e:
                msgs.pop()
                st.error(f"Erreur : {e}")
    if msgs and st.button("Effacer la conversation"):
        st.session_state.chat = []
        sauver_chat(cours.get("id"), [])
        st.rerun()


def page_historique():
    entete("Historique", "Retrouvez les résumés, QCM et résultats de cette session.")
    st.caption("L'historique est conservé tant que cette page reste ouverte : téléchargez vos fiches "
               "pour les garder durablement.")
    s = stock()
    t1, t2 = st.tabs(["Résumés", "QCM et notes"])
    with t1:
        lst = sorted(s["resumes"].values(), key=lambda r: r["id"], reverse=True)
        if not lst:
            st.caption("Aucun résumé pour l'instant.")
        for r in lst:
            with st.container(border=True):
                a, b, c = st.columns([5, 1.3, 1.5])
                a.markdown(f"**{html.escape(r['titre'])}**  \n<small>{r['cree']}</small>", unsafe_allow_html=True)
                if b.button("Ouvrir", key=f"hr{r['id']}", use_container_width=True):
                    st.session_state.resume = r
                    st.success("Résumé chargé : ouvre l'onglet **Résumé**.")
                if c.button("Supprimer", key=f"dr{r['id']}", use_container_width=True):
                    s["resumes"].pop(r["id"], None)
                    if (st.session_state.get("resume") or {}).get("id") == r["id"]:
                        st.session_state.resume = None
                    st.rerun()
    with t2:
        lst = sorted(s["qcms"].values(), key=lambda q: q["id"], reverse=True)
        if not lst:
            st.caption("Aucun QCM pour l'instant.")
        for q in lst:
            essais = [t for t in s["tentatives"] if t["qcm_id"] == q["id"]]
            meilleure = max((t["score"] / t["total"] * 20 for t in essais if t["total"]), default=None)
            bilan = f"meilleure note {round(meilleure, 1)}/20" if meilleure is not None else "pas encore tenté"
            with st.container(border=True):
                a, b, c = st.columns([5, 1.3, 1.5])
                a.markdown(f"**{html.escape(q['titre'])}**  \n<small>{q['cree']} · {len(essais)} essai(s) · {bilan}</small>",
                           unsafe_allow_html=True)
                if b.button("Refaire", key=f"hq{q['id']}", use_container_width=True):
                    st.session_state.qcm = {"id": q["id"], "titre": q["titre"], "questions": q["questions"]}
                    st.session_state.valide = False
                    st.session_state.version = st.session_state.get("version", 0) + 1
                    st.success("QCM chargé : ouvre l'onglet **QCM**.")
                if c.button("Supprimer", key=f"dq{q['id']}", use_container_width=True):
                    s["tentatives"] = [t for t in s["tentatives"] if t["qcm_id"] != q["id"]]
                    s["qcms"].pop(q["id"], None)
                    st.rerun()
    st.divider()
    if st.button("Nouvelle session", help="Efface les cours, résumés et QCM de cette session."):
        for k in list(st.session_state.keys()):
            if k not in ("seo_ok", "nav"):
                del st.session_state[k]
        st.rerun()


# ----------------------------------------------------------------------------
# Application (accès direct, sans connexion)
# ----------------------------------------------------------------------------
page = haut_de_page()
cle = cle_api(afficher=page in ("Mon cours", "Résumé", "QCM", "Assistant"))

if page == "Accueil":
    page_accueil()
elif page == "Mon cours":
    page_cours(cle)
elif page == "Résumé":
    page_resume(cle)
elif page == "QCM":
    page_qcm(cle)
elif page == "Assistant":
    page_chat(cle)
elif page == "Historique":
    page_historique()
else:
    page_services()

pied()
