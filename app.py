"""
RévisePro – Site de révision : cours (PDF/Word) -> résumé + QCM notés avec correction.
Auteur : BOKOUMBO Tibetobe Ambroise
Contact : ambroisebok@gmail.com | +228 91942948
"""
import io
import json
import os
import re

import anthropic
import streamlit as st
from docx import Document
from pypdf import PdfReader

MODEL = "claude-sonnet-5-5"
MAX_CHARS = 30000
LETTRES = "ABCD"

st.set_page_config(page_title="RévisePro", page_icon="📚", layout="centered")


# ---------- Lecture des fichiers ----------
def lire_fichier(fichier) -> str:
    nom = fichier.name.lower()
    data = fichier.getvalue()
    if nom.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(data))
        texte = "\n".join((p.extract_text() or "") for p in reader.pages)
    elif nom.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        texte = "\n".join(p.text for p in doc.paragraphs)
    else:
        raise ValueError("Format non pris en charge (PDF ou .docx uniquement).")
    return re.sub(r"\s+", " ", texte).strip()


# ---------- Génération par IA ----------
def extraire_json(brut: str) -> dict:
    brut = re.sub(r"^```(?:json)?|```$", "", brut.strip(), flags=re.M).strip()
    debut, fin = brut.find("{"), brut.rfind("}")
    return json.loads(brut[debut : fin + 1])


def generer(texte: str, nb: int, cle: str) -> dict:
    client = anthropic.Anthropic(api_key=cle)
    prompt = f"""Tu es un professeur. À partir du cours ci-dessous, réponds UNIQUEMENT
en JSON valide, en français, avec exactement cette structure :
{{
  "titre": "...",
  "resume": "résumé en 2 à 3 paragraphes",
  "points": ["10 points clés maximum"],
  "definitions": [{{"terme": "...", "def": "..."}}],
  "qcm": [{{"q": "question", "o": ["A", "B", "C", "D"], "r": 0, "e": "explication courte"}}]
}}
Le QCM contient exactement {nb} questions fidèles au cours, une seule bonne réponse
par question, "r" étant l'indice (0 à 3) de la bonne option.

COURS :
{texte[:MAX_CHARS]}"""
    rep = client.messages.create(
        model=MODEL,
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}],
    )
    return extraire_json(rep.content[0].text)


# ---------- Interface ----------
st.title("📚 RévisePro")
st.caption("Dépose ton cours : obtiens un résumé et des QCM notés avec correction.")

cle = os.getenv("ANTHROPIC_API_KEY") or st.sidebar.text_input(
    "Clé API Anthropic", type="password"
)
fichier = st.file_uploader("Ton cours (PDF ou Word)", type=["pdf", "docx"])
nb = st.selectbox("Nombre de questions", [5, 10, 15, 20], index=1)

if st.button("Générer résumé et QCM", type="primary", disabled=not fichier):
    if not cle:
        st.error("Ajoute ta clé API Anthropic (barre latérale ou variable ANTHROPIC_API_KEY).")
    else:
        try:
            with st.spinner("Lecture du cours…"):
                texte = lire_fichier(fichier)
            if len(texte) < 80:
                st.error("Aucun texte lisible (PDF scanné ou vide ?).")
            else:
                with st.spinner("Génération en cours…"):
                    st.session_state.data = generer(texte, nb, cle)
                st.session_state.valide = False
                st.session_state.version = st.session_state.get("version", 0) + 1
        except Exception as e:
            st.error(f"Erreur : {e}")

data = st.session_state.get("data")
if data:
    onglet_resume, onglet_qcm = st.tabs(["📝 Résumé", "✅ QCM"])

    with onglet_resume:
        st.header(data.get("titre", "Résumé du cours"))
        st.write(data.get("resume", ""))
        if data.get("points"):
            st.subheader("Points clés")
            for p in data["points"]:
                st.markdown(f"- {p}")
        if data.get("definitions"):
            st.subheader("Définitions")
            for d in data["definitions"]:
                st.markdown(f"- **{d['terme']}** : {d['def']}")

    with onglet_qcm:
        qcm = data.get("qcm", [])
        v = st.session_state.get("version", 0)
        valide = st.session_state.get("valide", False)
        reponses = []
        for i, q in enumerate(qcm):
            st.markdown(f"**{i + 1}. {q['q']}**")
            choix = st.radio(
                "Réponse",
                range(len(q["o"])),
                format_func=lambda j, q=q: f"{LETTRES[j]}. {q['o'][j]}",
                index=None,
                key=f"q{v}_{i}",
                disabled=valide,
                label_visibility="collapsed",
            )
            reponses.append(choix)
            if valide:
                if choix == q["r"]:
                    st.success(f"✅ Correct. {q.get('e', '')}")
                else:
                    st.error(
                        f"{'⚠️ Sans réponse' if choix is None else '❌ Incorrect'}. "
                        f"Bonne réponse : {LETTRES[q['r']]}. {q.get('e', '')}"
                    )
            st.divider()

        if not valide:
            if st.button("Valider et voir la correction", type="primary"):
                st.session_state.valide = True
                st.rerun()
        else:
            score = sum(r == q["r"] for r, q in zip(reponses, qcm))
            st.metric("Ta note", f"{round(score / len(qcm) * 20, 1)} / 20",
                      f"{score} bonne(s) réponse(s) sur {len(qcm)}")
            if st.button("Recommencer"):
                st.session_state.valide = False
                st.session_state.version += 1
                st.rerun()

st.markdown("---")
st.caption(
    "Conçu par **BOKOUMBO Tibetobe Ambroise** · ambroisebok@gmail.com · +228 91942948  \n"
    "Contenus générés par IA : vérifie toujours avec ton cours."
)
