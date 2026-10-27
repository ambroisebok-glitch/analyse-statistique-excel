"""
RévisePro (version Gemini) – cours (PDF/Word) -> résumé et QCM notés avec correction.
Auteur : BOKOUMBO Tibetobe Ambroise
Contact : ambroisebok@gmail.com | +228 91942948
"""
import io
import json
import os
import re
import time

import streamlit as st
from docx import Document
from google import genai
from pypdf import PdfReader

MODEL = "gemini-3.8-flash"
MAX_CHARS = 30000
LETTRES = "ABCD"

st.set_page_config(page_title="RévisePro", page_icon="📚", layout="centered")


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


def extraire_json(brut: str) -> dict:
    brut = re.sub(r"^```(?:json)?|```$", "", brut.strip(), flags=re.M).strip()
    debut, fin = brut.find("{"), brut.rfind("}")
    return json.loads(brut[debut : fin + 1])


def modeles_disponibles(client) -> list:
    """Modèle préféré d'abord, puis les autres modèles 'flash' de ton compte."""
    noms = [MODEL]
    try:
        for m in client.models.list():
            n = m.name.replace("models/", "")
            actions = getattr(m, "supported_actions", None) or ["generateContent"]
            exclus = ("image", "tts", "live", "audio", "embedding", "robotics")
            if ("flash" in n and "generateContent" in actions
                    and not any(x in n for x in exclus) and n not in noms):
                noms.append(n)
    except Exception:
        pass
    return noms[:5]


def appeler_gemini(prompt: str, cle: str) -> dict:
    client = genai.Client(api_key=cle)
    derniere_erreur = None
    for modele in modeles_disponibles(client):
        for essai in range(2):  # 2 essais par modèle
            try:
                rep = client.models.generate_content(
                    model=modele,
                    contents=prompt,
                    config={"response_mime_type": "application/json"},
                )
                return extraire_json(rep.text)
            except json.JSONDecodeError as e:
                derniere_erreur = e
                continue
            except Exception as e:
                derniere_erreur = e
                msg = str(e)
                if any(m in msg for m in ("503", "UNAVAILABLE", "429")):
                    time.sleep(4)
                    continue
                if "404" in msg or "NOT_FOUND" in msg:
                    break  # modèle indisponible : passer au suivant
                raise
    raise RuntimeError(
        "Gemini est surchargé sur tous les modèles essayés. "
        f"Réessaie dans quelques minutes. ({derniere_erreur})"
    )


def generer_resume(texte: str, cle: str) -> dict:
    prompt = f"""Tu es un professeur. À partir du cours ci-dessous, réponds UNIQUEMENT
en JSON valide, en français, avec exactement cette structure :
{{
  "titre": "...",
  "resume": "résumé en 2 à 3 paragraphes",
  "points": ["10 points clés maximum"],
  "definitions": [{{"terme": "...", "def": "..."}}]
}}

COURS :
{texte[:MAX_CHARS]}"""
    return appeler_gemini(prompt, cle)


def generer_qcm(texte: str, nb: int, cle: str) -> list:
    prompt = f"""Tu es un professeur. À partir du cours ci-dessous, réponds UNIQUEMENT
en JSON valide, en français, avec exactement cette structure :
{{
  "qcm": [{{"q": "question", "o": ["A", "B", "C", "D"], "r": 0, "e": "explication courte"}}]
}}
Le QCM contient exactement {nb} questions fidèles au cours, une seule bonne réponse
par question, "r" étant l'indice (0 à 3) de la bonne option.

COURS :
{texte[:MAX_CHARS]}"""
    return appeler_gemini(prompt, cle).get("qcm", [])


# ---------- Interface ----------
st.title("📚 RévisePro")
st.caption("Dépose ton cours, puis génère un résumé ou un QCM noté avec correction.")

cle = os.getenv("GEMINI_API_KEY") or st.sidebar.text_input("Clé API Gemini", type="password")
fichier = st.file_uploader("Ton cours (PDF ou Word)", type=["pdf", "docx"])

col1, col2 = st.columns(2)
with col1:
    st.subheader("📝 Résumé")
    btn_resume = st.button("Générer le résumé", disabled=not fichier, use_container_width=True)
with col2:
    st.subheader("✅ QCM")
    nb = st.selectbox("Nombre de questions", [5, 10, 15, 20], index=1)
    btn_qcm = st.button("Générer le QCM", type="primary", disabled=not fichier,
                        use_container_width=True)

if btn_resume or btn_qcm:
    if not cle:
        st.error("Ajoute ta clé API Gemini (barre latérale ou variable GEMINI_API_KEY).")
    else:
        try:
            with st.spinner("Lecture du cours…"):
                texte = lire_fichier(fichier)
            if len(texte) < 80:
                st.error("Aucun texte lisible (PDF scanné ou vide ?).")
            elif btn_resume:
                with st.spinner("Génération du résumé…"):
                    st.session_state.resume = generer_resume(texte, cle)
                st.session_state.vue = "📝 Résumé"
            else:
                with st.spinner("Génération du QCM…"):
                    st.session_state.qcm = generer_qcm(texte, nb, cle)
                st.session_state.valide = False
                st.session_state.version = st.session_state.get("version", 0) + 1
                st.session_state.vue = "✅ QCM"
        except Exception as e:
            st.error(f"Erreur : {e}")

if st.session_state.get("resume") or st.session_state.get("qcm"):
    vue = st.radio("Affichage", ["📝 Résumé", "✅ QCM"], horizontal=True, key="vue",
                   label_visibility="collapsed")
    st.divider()

    if vue == "📝 Résumé":
        data = st.session_state.get("resume")
        if not data:
            st.info("Clique sur « Générer le résumé » pour l'afficher ici.")
        else:
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
    else:
        qcm = st.session_state.get("qcm")
        if not qcm:
            st.info("Clique sur « Générer le QCM » pour l'afficher ici.")
        else:
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
    "Conçu par **BOKOUMBO Tibetobe Ambroise** · ambroisebok@gmail.com · +228 91942948"
)
