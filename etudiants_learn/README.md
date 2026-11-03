# Étudiants Learn

Application Streamlit : cours (PDF/Word) → résumé approfondi, QCM jusqu'à 100 questions, assistant IA.
Exports du résumé : PDF, Word, HTML, Markdown.

## Lancer
```
pip install -r requirements.txt
streamlit run app_gemini.py
```
Clé Gemini : variable `GEMINI_API_KEY` ou secrets Streamlit. Modèle : variable `GEMINI_MODEL`.

## Fichiers
- `app_gemini.py` : application
- `style.py` : design (CSS) et illustration de la page de connexion
- `.streamlit/config.toml` : thème de base
- Ne versionnez pas `revisepro.db` (comptes utilisateurs) : ajoutez-le à `.gitignore`.
