"""
Étudiants Learn (version Gemini) – cours (PDF/Word) -> résumé approfondi, QCM par chapitre (jusqu'à 100),
assistant IA, comptes utilisateurs et exports (PDF, Word, HTML, Markdown).
Auteur : BOKOUMBO Tibetobe Ambroise
Contact : ambroisebok@gmail.com | +228 91942948
"""
import hashlib
import hmac
import html
import io
import json
import os
import random
import re
import secrets
import sqlite3
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

from style import ILLUSTRATION, appliquer_style

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
# Modèle préféré (modifiable sans toucher au code via la variable GEMINI_MODEL / secrets Streamlit)
MODEL = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite")
MAX_CHARS = 700_000                 # le cours est lu en entier (≈ 175 000 tokens)
MAX_PDF_OCTETS = 18 * 1024 * 1024   # taille max d'un PDF envoyé directement à Gemini
LOT_QCM = 20                        # questions générées par appel
MAX_QCM = 100
LETTRES = "ABCD"
DB_PATH = os.getenv("DB_PATH", "revisepro.db")
PAGES = ["Mon cours", "Résumé", "QCM", "Assistant", "Historique", "Services & contact"]

AUTEUR = "BOKOUMBO Tibetobe Ambroise"
EMAIL = "ambroisebok@gmail.com"
TEL_REP = ("+228 96286244", "+22896286244")   # répétitions + démarches
TEL_AMB = ("+228 91942948", "+22891942948")   # contact principal
APP_URL = "https://etudiantslearn.streamlit.app/"
SEO_TITRE = "Étudiants Learn – Résumés de cours, QCM et assistant IA"
SEO_DESC = ("Étudiants Learn aide les étudiants à réviser leurs cours PDF ou Word grâce à des résumés "
            "approfondis, des QCM corrigés jusqu'à 100 questions et un assistant IA pour comprendre les cours "
            "et préparer les examens.")

st.set_page_config(
    page_title=SEO_TITRE, page_icon="🎓", layout="wide", initial_sidebar_state="expanded",
    menu_items={"About": f"**Étudiants Learn** — conçu par {AUTEUR} · {EMAIL} · {TEL_AMB[0]}"})

appliquer_style()


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
        "répétiteur Togo, démarches en ligne Togo, passeport, nationalité, casier judiciaire");
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
# Base de données (SQLite) et comptes utilisateurs
# ----------------------------------------------------------------------------
def maintenant() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def sql(requete, params=(), fetch=None):
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    try:
        cur = c.execute(requete, params)
        c.commit()
        if fetch == "one":
            r = cur.fetchone()
            return dict(r) if r else None
        if fetch == "all":
            return [dict(r) for r in cur.fetchall()]
        return cur.lastrowid
    finally:
        c.close()


@st.cache_resource
def init_db():
    c = sqlite3.connect(DB_PATH)
    c.executescript(
        """
        CREATE TABLE IF NOT EXISTS users(
            id INTEGER PRIMARY KEY AUTOINCREMENT, email TEXT UNIQUE, nom TEXT,
            sel TEXT, hash TEXT, cree TEXT);
        CREATE TABLE IF NOT EXISTS cours(
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, nom TEXT,
            texte TEXT, pdf BLOB, chapitres TEXT, cree TEXT);
        CREATE TABLE IF NOT EXISTS resumes(
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, cours_id INTEGER,
            titre TEXT, contenu TEXT, cree TEXT);
        CREATE TABLE IF NOT EXISTS qcms(
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, cours_id INTEGER,
            titre TEXT, data TEXT, cree TEXT);
        CREATE TABLE IF NOT EXISTS tentatives(
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, qcm_id INTEGER,
            score INTEGER, total INTEGER, cree TEXT);
        CREATE TABLE IF NOT EXISTS chats(
            cours_id INTEGER, user_id INTEGER, messages TEXT,
            PRIMARY KEY(cours_id, user_id));
        """
    )
    c.commit()
    c.close()
    return True


def hacher(mdp: str, sel: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", mdp.encode(), bytes.fromhex(sel), 200_000).hex()


def creer_compte(nom: str, email: str, mdp: str):
    email = email.strip().lower()
    if len(nom.strip()) < 2:
        return None, "Entre ton nom."
    if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
        return None, "Adresse e-mail invalide."
    if len(mdp) < 8:
        return None, "Le mot de passe doit contenir au moins 8 caractères."
    sel = secrets.token_hex(16)
    try:
        i = sql("INSERT INTO users(email,nom,sel,hash,cree) VALUES(?,?,?,?,?)",
                (email, nom.strip(), sel, hacher(mdp, sel), maintenant()))
    except sqlite3.IntegrityError:
        return None, "Un compte existe déjà avec cet e-mail."
    return {"id": i, "nom": nom.strip(), "email": email}, None


def connecter(email: str, mdp: str):
    u = sql("SELECT * FROM users WHERE email=?", (email.strip().lower(),), fetch="one")
    if not u or not hmac.compare_digest(u["hash"], hacher(mdp, u["sel"])):
        return None, "E-mail ou mot de passe incorrect."
    return {"id": u["id"], "nom": u["nom"], "email": u["email"]}, None


def uid():
    u = st.session_state.get("user")
    return u["id"] if u else None


def enregistrer_cours(c: dict) -> int:
    return sql("INSERT INTO cours(user_id,nom,texte,pdf,chapitres,cree) VALUES(?,?,?,?,?,?)",
               (uid(), c["nom"], c["texte"], c.get("pdf"), json.dumps(c["chapitres"]), maintenant()))


def charger_cours(cid: int):
    r = sql("SELECT * FROM cours WHERE id=? AND user_id=?", (cid, uid()), fetch="one")
    if not r:
        return None
    return {"id": r["id"], "nom": r["nom"], "texte": r["texte"], "pdf": r["pdf"],
            "chapitres": json.loads(r["chapitres"]), "vision": bool(r["pdf"]),
            "pages": 0}


def supprimer_cours(cid: int):
    sql("DELETE FROM tentatives WHERE qcm_id IN (SELECT id FROM qcms WHERE cours_id=? AND user_id=?)",
        (cid, uid()))
    for t in ("resumes", "qcms", "chats"):
        sql(f"DELETE FROM {t} WHERE cours_id=? AND user_id=?", (cid, uid()))
    sql("DELETE FROM cours WHERE id=? AND user_id=?", (cid, uid()))


def charger_chat(cid):
    if not (uid() and cid):
        return []
    r = sql("SELECT messages FROM chats WHERE cours_id=? AND user_id=?", (cid, uid()), fetch="one")
    return json.loads(r["messages"]) if r else []


def sauver_chat(cid, msgs):
    if uid() and cid:
        sql("INSERT OR REPLACE INTO chats(cours_id,user_id,messages) VALUES(?,?,?)",
            (cid, uid(), json.dumps(msgs)))


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
    return noms[:8]


def _temporaire(msg: str) -> bool:
    return any(m in msg for m in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED",
                                  "500", "INTERNAL", "overloaded", "DEADLINE"))


def parts_cours(cours: dict) -> list:
    if cours.get("pdf") and cours.get("vision"):
        return [types.Part.from_bytes(data=cours["pdf"], mime_type="application/pdf")]
    return [types.Part.from_text(text="COURS :\n" + cours["texte"][:MAX_CHARS])]


def parts_chat(cours: dict, visuel: bool) -> list:
    """Contexte du chat : le texte extrait est bien plus léger qu'un PDF envoyé à chaque question."""
    if visuel or len(cours["texte"]) < 500:
        return parts_cours(cours)
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
        for essai in range(3):
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
    """Réponse en streaming. La « réflexion » du modèle est coupée (premiers mots bien plus vite) ;
    si le modèle la refuse, on la réactive. Repli automatique sur d'autres modèles."""
    client = genai.Client(api_key=cle)
    derniere = None
    for modele in modeles_disponibles(cle):
        for _ in range(2):
            kw = dict(system_instruction=system, temperature=0.5,
                      max_output_tokens=REGLAGES["max_tokens"] or 8192)
            if REGLAGES["sans_reflexion"]:
                kw["thinking_config"] = types.ThinkingConfig(thinking_budget=0)
            deja = False
            try:
                for morceau in client.models.generate_content_stream(
                        model=modele, contents=contenu, config=types.GenerateContentConfig(**kw)):
                    if morceau.text:
                        deja = True
                        yield morceau.text
                return
            except Exception as e:
                derniere = e
                if deja:
                    raise
                msg = str(e)
                if ("400" in msg or "INVALID_ARGUMENT" in msg) and REGLAGES["sans_reflexion"]:
                    REGLAGES["sans_reflexion"] = False   # réglage refusé : on réessaie sans
                    continue
                if "404" in msg or "NOT_FOUND" in msg or _temporaire(msg):
                    break                                # modèle indisponible : on passe au suivant
                raise
    raise RuntimeError(f"L'assistant est indisponible pour le moment. Réessaie dans un instant. ({derniere})")


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
h1{color:#4338CA;border-bottom:3px solid #4338CA;padding-bottom:.3rem}
h2{color:#4338CA;margin-top:2.2rem;border-bottom:1px solid #E5E7EB;padding-bottom:.2rem;
page-break-before:auto}
h3{color:#374151;margin-top:1.4rem}
table{border-collapse:collapse;margin:1rem 0}th,td{border:1px solid #D1D5DB;padding:.4rem .7rem}
th{background:#EEF2FF}code{background:#F3F4F6;padding:.1rem .3rem;border-radius:4px}
.note{color:#6B7280;font-size:.85rem;border-top:1px solid #E5E7EB;margin-top:3rem;padding-top:.6rem}
@media print{body{margin:0;max-width:none}.noprint{display:none}h2{page-break-before:always}}
</style></head><body>
<p class="noprint" style="background:#EEF2FF;padding:.6rem 1rem;border-radius:8px;font-family:sans-serif">
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
            doc.styles[nom].font.color.rgb = RGBColor(0x43, 0x38, 0xCA)
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
    t = re.sub("[\U00010000-\U0010FFFF\ufe0f\u200d\u26a0\u2705\u274c\u274e\u2728\u2b50\u2753\u2757\u26a1]", "", t)
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
    encre, accent = colors.HexColor("#1F2937"), colors.HexColor("#4338CA")
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
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EEF2FF")),
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
                 HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#4338CA"), spaceAfter=10)]
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


def sauver_qcm(titre, questions):
    c = st.session_state.get("cours") or {}
    if uid():
        return sql("INSERT INTO qcms(user_id,cours_id,titre,data,cree) VALUES(?,?,?,?,?)",
                   (uid(), c.get("id"), titre, json.dumps(questions), maintenant()))
    return None


# ----------------------------------------------------------------------------
# Interface : composants communs
# ----------------------------------------------------------------------------
def entete(titre, sous=""):
    st.markdown(f'<div class="entete"><h1>{titre}</h1><p>{sous}</p></div>', unsafe_allow_html=True)


def cle_api():
    cle = os.getenv("GEMINI_API_KEY")
    if not cle:
        try:
            cle = st.secrets.get("GEMINI_API_KEY")
        except Exception:
            cle = None
    if not cle:
        with st.expander("Réglages de l'IA", expanded=True):
            cle = st.text_input("Clé API Gemini", type="password",
                                help="Ou définis la variable GEMINI_API_KEY (secrets Streamlit).")
    return cle


def pied():
    st.markdown(
        f'''<div class="barre-pied">
        <b>Étudiants Learn</b> · Conçu par {AUTEUR}<span class="sep">|</span>
        <a href="mailto:{EMAIL}">{EMAIL}</a><span class="sep">|</span>
        <a href="tel:{TEL_AMB[1]}">{TEL_AMB[0]}</a><span class="sep">|</span>
        Répétitions : <a href="tel:{TEL_REP[1]}">{TEL_REP[0]}</a></div>''',
        unsafe_allow_html=True)


def boutons_contact(numero, message="Bonjour, je vous contacte depuis Étudiants Learn."):
    from urllib.parse import quote
    return (f'<a class="btn" href="tel:{numero[1]}">Appeler {numero[0]}</a>'
            f'<a class="btn wa" target="_blank" href="https://wa.me/{numero[1].lstrip("+")}?text={quote(message)}">'
            f'WhatsApp</a>')


ICONES = {
    "livre": '<svg viewBox="0 0 24 24"><path d="M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2z"/><path d="M4 21V5"/></svg>',
    "dossier": '<svg viewBox="0 0 24 24"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="M9 14l2 2 4-4"/></svg>',
    "graphique": '<svg viewBox="0 0 24 24"><path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/></svg>',
}

SERVICES = [
    {"icone": "livre", "tag": "Accompagnement", "titre": "Répétitions",
     "texte": "Un suivi sur mesure pour progresser et aborder vos examens avec confiance.",
     "points": ["Soutien scolaire et universitaire", "Préparation aux examens et concours",
                "Méthodes de travail et organisation"],
     "contact": [TEL_REP[0]]},
    {"icone": "dossier", "tag": "Administratif", "titre": "Démarches en ligne",
     "texte": "Nous vous guidons dans vos demandes administratives, de la préparation du dossier à son dépôt.",
     "points": ["Certificat de nationalité", "Passeport", "Casier judiciaire"],
     "contact": [TEL_REP[0], TEL_AMB[0]]},
    {"icone": "graphique", "tag": "Données", "titre": "Traitement de données",
     "texte": "Des données propres, exploitables et clairement présentées.",
     "points": ["Saisie et nettoyage", "Analyse et calculs", "Tableaux, graphiques et mise en forme"],
     "contact": [TEL_REP[0], TEL_AMB[0]]},
]


def bloc_services():
    cartes = ""
    for s in SERVICES:
        pts = "".join(f"<li>{p}</li>" for p in s["points"])
        tels = " · ".join(s["contact"])
        cartes += (f'<div class="service"><div class="ico">{ICONES[s["icone"]]}</div>'
                   f'<div class="tag" style="font-size:.68rem;font-weight:700;letter-spacing:.08em;'
                   f'text-transform:uppercase;color:var(--accent);margin-bottom:.2rem">{s["tag"]}</div>'
                   f'<h4>{s["titre"]}</h4><p>{s["texte"]}</p><ul>{pts}</ul>'
                   f'<div class="contact">Contact : <b>{tels}</b></div></div>')
    st.markdown(f'<div class="services">{cartes}</div>', unsafe_allow_html=True)


def page_services():
    entete("Services & contact",
           "Au-delà de l'application, nous vous accompagnons concrètement dans vos études et vos démarches.")
    bloc_services()

    st.markdown('<div class="section-titre">Comment ça se passe</div>', unsafe_allow_html=True)
    st.markdown(
        '''<div class="parcours">
        <div><b>1. Vous nous contactez</b><span>Par appel, WhatsApp ou e-mail, en décrivant votre besoin.</span></div>
        <div><b>2. Nous cadrons la demande</b><span>Nous précisons ensemble les pièces, le délai et les modalités.</span></div>
        <div><b>3. Nous vous accompagnons</b><span>Nous réalisons la mission et restons disponibles jusqu'à son terme.</span></div>
        </div>''', unsafe_allow_html=True)

    st.markdown('<div class="section-titre">Nous contacter</div>', unsafe_allow_html=True)
    st.markdown("**Répétitions**")
    st.markdown(boutons_contact(TEL_REP, "Bonjour, je souhaite des informations sur les répétitions."),
                unsafe_allow_html=True)
    st.markdown("**Démarches en ligne et traitement de données**")
    msg = "Bonjour, je souhaite être accompagné(e) pour une démarche ou un traitement de données."
    st.markdown(boutons_contact(TEL_REP, msg) + boutons_contact(TEL_AMB, msg), unsafe_allow_html=True)
    st.markdown("**Par e-mail**")
    st.markdown(f'<a class="btn clair" href="mailto:{EMAIL}">{EMAIL}</a>', unsafe_allow_html=True)
    st.caption("Service d'accompagnement privé, non affilié à l'administration. "
               "Les délais et les décisions relèvent des services officiels.")


def etapes():
    ok = [bool(st.session_state.get("cours")), bool(st.session_state.get("resume")),
          bool(st.session_state.get("qcm")), bool(st.session_state.get("chat"))]
    noms = ["Charger le cours", "Résumé", "QCM", "Assistant"]
    st.markdown('<div class="etapes">' + "".join(
        f'<span class="etape {"ok" if o else ""}"><i>{"✓" if o else k}</i>{n}</span>'
        for k, (o, n) in enumerate(zip(ok, noms), 1)) + "</div>", unsafe_allow_html=True)


def aller(page: str):
    """Rappel de bouton : change de page sans repasser par la barre latérale."""
    st.session_state.nav = page


def besoin_cours():
    if not st.session_state.get("cours"):
        st.markdown('<div class="vide"><b>Aucun cours chargé</b>Commence par déposer un cours (PDF ou Word) '
                    'dans l\'onglet « Mon cours » : le reste de l\'application s\'appuie dessus.</div>',
                    unsafe_allow_html=True)
        st.button("Charger un cours", type="primary", on_click=aller, args=(PAGES[0],), key="va_cours")
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
def page_connexion():
    # Pas de barre latérale avant la connexion
    st.markdown("<style>[data-testid='stSidebar'], [data-testid='collapsedControl'], "
                "[data-testid='stSidebarCollapsedControl'] { display:none; }</style>",
                unsafe_allow_html=True)
    gauche, droite = st.columns([1.12, 1], gap="large")
    with gauche:
        st.markdown(
            f'''<div class="marque-lg"><span class="logo">EL</span><b>Étudiants Learn</b></div>
<div class="titre-accueil">Révisez plus vite, <em>à partir de vos propres cours.</em></div>
<p class="sous-accueil">Importez un cours au format PDF ou Word : obtenez des fiches de révision structurées,
des QCM corrigés et un assistant qui répond à vos questions.</p>
<div class="illus">{ILLUSTRATION}</div>
<div class="chiffres"><div><b>PDF &amp; Word</b>formats acceptés</div>
<div><b>100</b>questions par QCM</div><div><b>4</b>exports : PDF, Word, HTML, Markdown</div></div>''',
            unsafe_allow_html=True)
    with droite:
        st.markdown('<div class="acces-titre">Accédez à votre espace</div>'
                    '<div class="acces-sous">Connectez-vous ou créez un compte gratuit pour retrouver vos cours, '
                    'résumés et résultats à tout moment.</div>', unsafe_allow_html=True)
        t1, t2 = st.tabs(["Connexion", "Créer un compte"])
        with t1:
            with st.form("login"):
                email = st.text_input("Adresse e-mail", placeholder="nom@exemple.com")
                mdp = st.text_input("Mot de passe", type="password")
                if st.form_submit_button("Se connecter", type="primary", use_container_width=True):
                    u, err = connecter(email, mdp)
                    if err:
                        st.error(err)
                    else:
                        st.session_state.user = u
                        st.rerun()
        with t2:
            with st.form("signup"):
                nom = st.text_input("Nom complet", placeholder="Prénom Nom")
                email = st.text_input("Adresse e-mail", key="se", placeholder="nom@exemple.com")
                mdp = st.text_input("Mot de passe (8 caractères minimum)", type="password", key="sm")
                if st.form_submit_button("Créer mon compte", type="primary", use_container_width=True):
                    u, err = creer_compte(nom, email, mdp)
                    if err:
                        st.error(err)
                    else:
                        st.session_state.user = u
                        st.rerun()
        st.markdown(
            '''<ul class="atouts">
<li><i class="pt">✓</i><div><b>Résumés approfondis</b><span>Une fiche par chapitre : définitions, formules,
démonstrations et exemples résolus.</span></div></li>
<li><i class="pt">✓</i><div><b>QCM avec correction détaillée</b><span>Sur tout le cours ou par chapitre,
avec une note sur 20 et vos points faibles.</span></div></li>
<li><i class="pt">✓</i><div><b>Assistant pédagogique</b><span>Explications pas à pas, exercices corrigés,
conseils pour l'examen.</span></div></li></ul>
<div class="confiance">Vos cours et vos résultats sont liés à votre compte.</div>''',
            unsafe_allow_html=True)
    st.markdown('<div class="section-titre">Nos services</div>', unsafe_allow_html=True)
    bloc_services()
    pied()


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
                    if uid():
                        st.write("Enregistrement dans ta bibliothèque…")
                        cours["id"] = enregistrer_cours(cours)
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
        n1, n2, n3 = st.columns(3)
        n1.button("Générer le résumé", type="primary", use_container_width=True, on_click=aller,
                  args=(PAGES[1],), key="suite_resume")
        n2.button("Passer un QCM", use_container_width=True, on_click=aller, args=(PAGES[2],), key="suite_qcm")
        n3.button("Poser une question", use_container_width=True, on_click=aller, args=(PAGES[3],), key="suite_chat")

    if uid():
        st.divider()
        st.subheader("Ma bibliothèque")
        lst = sql("SELECT id, nom, cree FROM cours WHERE user_id=? ORDER BY id DESC", (uid(),), fetch="all")
        if not lst:
            st.caption("Aucun cours enregistré pour l'instant.")
        for c in lst:
            with st.container(border=True):
                a, b, d = st.columns([5, 1.3, 1.5])
                a.markdown(f"**{c['nom']}**  \n<small>{c['cree']}</small>", unsafe_allow_html=True)
                if b.button("Ouvrir", key=f"o{c['id']}", use_container_width=True):
                    reset_travail(charger_cours(c["id"]))
                    st.rerun()
                if d.button("Supprimer", key=f"s{c['id']}", use_container_width=True):
                    supprimer_cours(c["id"])
                    if (st.session_state.get("cours") or {}).get("id") == c["id"]:
                        st.session_state.cours = None
                    st.rerun()


@st.cache_data(show_spinner=False, max_entries=6)
def pdf_en_cache(r_json: str) -> bytes:
    return resume_en_pdf(json.loads(r_json))


@st.cache_data(show_spinner=False, max_entries=6)
def docx_en_cache(r_json: str) -> bytes:
    return resume_en_docx(json.loads(r_json))


def afficher_resume(r: dict, cle_export: str):
    st.header(r["titre"])
    st.caption(f"Niveau : {r.get('niveau', '')} · généré le {r.get('cree', '')}")
    zone_export = st.container()          # rempli après l'affichage : le résumé apparaît sans attendre le PDF
    st.divider()
    st.markdown(r["apercu"])
    st.divider()
    st.subheader("Fiches par chapitre")
    for i, s in enumerate(r["sections"]):
        with st.expander(f"{i + 1}. {s['titre']}", expanded=(i == 0 and len(r["sections"]) == 1)):
            st.markdown(s["md"])

    md = resume_en_markdown(r)
    r_json = json.dumps(r, sort_keys=True, default=str)
    with zone_export:
        cp, cw = st.columns(2)
        with cp:
            try:
                with st.spinner("Préparation du PDF…"):
                    pdf = pdf_en_cache(r_json)
                st.download_button("Télécharger en PDF", pdf, file_name=f"{slug(r['titre'])}.pdf",
                                   mime="application/pdf", type="primary", use_container_width=True,
                                   key=f"dpdf{cle_export}")
            except Exception as e:
                st.error(f"Impossible de créer le PDF : {e}")
        with cw:
            try:
                with st.spinner("Préparation du fichier Word…"):
                    docx_octets = docx_en_cache(r_json)
                st.download_button("Télécharger en Word", docx_octets, file_name=f"{slug(r['titre'])}.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                   use_container_width=True, key=f"dw{cle_export}")
            except Exception as e:
                st.error(f"Impossible de créer le document Word : {e}")
        c3, c4 = st.columns(2)
        c3.download_button("Télécharger en HTML", md_vers_html(md, r["titre"]).encode("utf-8"),
                           file_name=f"{slug(r['titre'])}.html", mime="text/html", use_container_width=True,
                           help="Version web avec formules MathJax.", key=f"dh{cle_export}")
        c4.download_button("Télécharger en Markdown", md.encode("utf-8"), file_name=f"{slug(r['titre'])}.md",
                           mime="text/markdown", use_container_width=True, key=f"dm{cle_export}")


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
                if uid():
                    r["id"] = sql("INSERT INTO resumes(user_id,cours_id,titre,contenu,cree) VALUES(?,?,?,?,?)",
                                  (uid(), cours.get("id"), r["titre"], json.dumps(r), maintenant()))
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
            if uid() and qcm.get("id"):
                sql("INSERT INTO tentatives(user_id,qcm_id,score,total,cree) VALUES(?,?,?,?,?)",
                    (uid(), qcm["id"], sum(r == q["r"] for r, q in zip(reps, qs)), len(qs), maintenant()))
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
    visuel = False
    if cours.get("pdf"):
        visuel = st.toggle("Lecture visuelle du PDF (formules, schémas, scans) — plus lent", value=False,
                           key="chat_visuel", help="Par défaut, l'assistant lit le texte du cours : réponses plus rapides.")

    zone_sugg = st.empty()
    if not msgs:
        with zone_sugg.container():
            st.markdown('<div class="sugg-titre">Pour démarrer</div>', unsafe_allow_html=True)
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
            st.error("Ajoute ta clé API Gemini dans la barre latérale.")
            return
        zone_sugg.empty()
        msgs.append({"role": "user", "content": question})
        with st.chat_message("user"):
            st.markdown(question)
        histo = msgs[-20:]
        while histo and histo[0]["role"] != "user":
            histo = histo[1:]
        contenu = [
            types.Content(role="user", parts=parts_chat(cours, visuel) + [
                types.Part.from_text(text="Voici mon cours. Sers-t'en comme base pour répondre à mes questions.")]),
            types.Content(role="model", parts=[types.Part.from_text(text="Compris, je suis prêt à t'aider.")]),
        ] + [types.Content(role="user" if m["role"] == "user" else "model",
                           parts=[types.Part.from_text(text=m["content"])]) for m in histo]
        with st.chat_message("assistant"):
            zone = st.empty()
            zone.markdown("*L'assistant rédige sa réponse…*")
            acc = ""
            try:
                for morceau in flux_chat(contenu, cle, SYSTEME_CHAT):
                    acc += morceau
                    zone.markdown(normaliser_math(acc) + " ▌")
                if not acc.strip():
                    raise RuntimeError("réponse vide, réessaie.")
                final = normaliser_math(acc)
                zone.markdown(final)                      # pas de rechargement de page : la réponse reste affichée
                msgs.append({"role": "assistant", "content": final})
                sauver_chat(cours.get("id"), msgs)
            except Exception as e:
                msgs.pop()
                zone.empty()
                st.error(f"Erreur : {e}")
    if msgs and st.button("Effacer la conversation"):
        st.session_state.chat = []
        sauver_chat(cours.get("id"), [])
        st.rerun()


def page_historique():
    entete("Historique", "Retrouvez vos résumés, vos QCM et vos résultats.")
    if not uid():
        st.info("Connecte-toi pour consulter ton historique.")
        return
    t1, t2 = st.tabs(["Résumés", "QCM et notes"])
    with t1:
        lst = sql("SELECT id, titre, cree FROM resumes WHERE user_id=? ORDER BY id DESC", (uid(),), fetch="all")
        if not lst:
            st.caption("Aucun résumé enregistré.")
        for r in lst:
            with st.container(border=True):
                a, b, c = st.columns([5, 1.3, 1.5])
                a.markdown(f"**{r['titre']}**  \n<small>{r['cree']}</small>", unsafe_allow_html=True)
                if b.button("Ouvrir", key=f"hr{r['id']}", use_container_width=True):
                    row = sql("SELECT contenu FROM resumes WHERE id=? AND user_id=?", (r["id"], uid()), fetch="one")
                    d = json.loads(row["contenu"])
                    d["id"] = r["id"]
                    st.session_state.resume = d
                    st.success("Résumé chargé : ouvre l'onglet **Résumé**.")
                if c.button("Supprimer", key=f"dr{r['id']}", use_container_width=True):
                    sql("DELETE FROM resumes WHERE id=? AND user_id=?", (r["id"], uid()))
                    st.rerun()
    with t2:
        lst = sql("""SELECT q.id, q.titre, q.cree, COUNT(t.id) AS essais,
                     MAX(CAST(t.score AS REAL) / t.total * 20) AS meilleure
                     FROM qcms q LEFT JOIN tentatives t ON t.qcm_id = q.id
                     WHERE q.user_id=? GROUP BY q.id ORDER BY q.id DESC""", (uid(),), fetch="all")
        if not lst:
            st.caption("Aucun QCM enregistré.")
        for q in lst:
            with st.container(border=True):
                a, b, c = st.columns([5, 1.3, 1.5])
                meilleure = (f"meilleure note {round(q['meilleure'], 1)}/20"
                             if q["meilleure"] is not None else "pas encore tenté")
                a.markdown(f"**{q['titre']}**  \n<small>{q['cree']} · {q['essais']} essai(s) · {meilleure}</small>",
                           unsafe_allow_html=True)
                if b.button("Refaire", key=f"hq{q['id']}", use_container_width=True):
                    row = sql("SELECT data FROM qcms WHERE id=? AND user_id=?", (q["id"], uid()), fetch="one")
                    st.session_state.qcm = {"id": q["id"], "titre": q["titre"], "questions": json.loads(row["data"])}
                    st.session_state.valide = False
                    st.session_state.version = st.session_state.get("version", 0) + 1
                    st.success("QCM chargé : ouvre l'onglet **QCM**.")
                if c.button("Supprimer", key=f"dq{q['id']}", use_container_width=True):
                    sql("DELETE FROM tentatives WHERE qcm_id=? AND user_id=?", (q["id"], uid()))
                    sql("DELETE FROM qcms WHERE id=? AND user_id=?", (q["id"], uid()))
                    st.rerun()


# ----------------------------------------------------------------------------
# Application
# ----------------------------------------------------------------------------
init_db()

# Accès réservé aux comptes : pas de mode invité
if not st.session_state.get("user") or not st.session_state.user.get("id"):
    st.session_state.pop("user", None)
    page_connexion()
    st.stop()

with st.sidebar:
    utilisateur = st.session_state.user
    initiale = html.escape((utilisateur["nom"].strip()[:1] or "?").upper())
    st.markdown('<div class="marque"><span class="logo">EL</span><b>Étudiants Learn</b></div>',
                unsafe_allow_html=True)
    st.markdown(
        f'<div class="profil"><span class="av">{initiale}</span><div>'
        f'<div class="n">{html.escape(utilisateur["nom"])}</div>'
        f'<div class="e">{html.escape(utilisateur.get("email", ""))}</div></div></div>',
        unsafe_allow_html=True)
    page = st.radio("Navigation", PAGES, label_visibility="collapsed", key="nav")
    st.divider()
    if st.session_state.get("cours"):
        st.markdown(
            f'<div class="cours-actif"><span>COURS ACTIF</span>'
            f'<b>{html.escape(st.session_state.cours["nom"])}</b></div>', unsafe_allow_html=True)
    cle = cle_api()
    st.markdown(
        f'''<div class="promo"><b>Besoin d'un coup de main ?</b><br>
        Répétitions : {TEL_REP[0]}<br>
        Démarches et données : {TEL_AMB[0]}<br>
        <span style="font-size:.72rem">Détails dans l'onglet « Services & contact ».</span></div>''',
        unsafe_allow_html=True)
    st.write("")
    if st.button("Se déconnecter", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

if page == PAGES[0]:
    page_cours(cle)
elif page == PAGES[1]:
    page_resume(cle)
elif page == PAGES[2]:
    page_qcm(cle)
elif page == PAGES[3]:
    page_chat(cle)
elif page == PAGES[4]:
    page_historique()
else:
    page_services()

pied()
