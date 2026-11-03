"""Habillage d'Étudiants Learn : feuille de style, illustration et blocs HTML réutilisables.
Aucun appel réseau, aucun fichier image : l'illustration est un SVG intégré (net sur tous les écrans).
Usage : from style import appliquer_style, ILLUSTRATION"""
import streamlit as st

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
:root{
  --ink:#0F172A; --muted:#64748B; --line:#E2E8F0; --bg:#F8FAFC; --surface:#FFFFFF;
  --accent:#4338CA; --accent-d:#3730A3; --accent-l:#EEF2FF; --accent-m:#C7D2FE;
  --ok:#047857; --ok-l:#ECFDF5; --r:12px;
}
html,body,.stApp{font-family:'Inter',system-ui,-apple-system,'Segoe UI',sans-serif;color:var(--ink);}
.stApp{background:var(--bg);}
#MainMenu,footer,[data-testid="stToolbar"],[data-testid="stDecoration"]{display:none!important;}
header[data-testid="stHeader"]{background:transparent;}
iframe[height="0"],iframe[height="1"]{display:none;}
.block-container{max-width:1120px;padding:2.2rem 2.2rem 2rem;}
h1,h2,h3{letter-spacing:-.02em;color:var(--ink);}
a:focus-visible,button:focus-visible,input:focus-visible{outline:2px solid var(--accent);outline-offset:2px;}
@media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important;}}

/* ---------- Barre latérale ---------- */
[data-testid="stSidebar"]{background:var(--surface);border-right:1px solid var(--line);}
.marque,.marque-lg{display:flex;align-items:center;gap:.65rem;}
.marque{margin:.2rem 0 1.1rem;}
.logo{width:36px;height:36px;border-radius:10px;background:linear-gradient(135deg,#4338CA,#6D28D9);color:#fff;
  display:flex;align-items:center;justify-content:center;font-weight:800;font-size:.9rem;letter-spacing:.02em;}
.marque-lg .logo{width:42px;height:42px;font-size:1rem;border-radius:12px;}
.marque b,.marque-lg b{font-size:1.15rem;font-weight:700;color:var(--ink);letter-spacing:-.01em;}
.profil{display:flex;align-items:center;gap:.7rem;padding:.65rem .8rem;border:1px solid var(--line);
  border-radius:var(--r);background:var(--bg);margin-bottom:1rem;}
.profil .av{flex:0 0 34px;height:34px;border-radius:50%;background:var(--accent);color:#fff;font-weight:700;
  display:flex;align-items:center;justify-content:center;font-size:.9rem;}
.profil .n{font-weight:600;font-size:.88rem;line-height:1.25;}
.profil .e{font-size:.74rem;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:150px;}
[data-testid="stSidebar"] [role="radiogroup"]{gap:.15rem;}
[data-testid="stSidebar"] [role="radiogroup"] label{padding:.55rem .8rem;border-radius:10px;width:100%;
  border-left:3px solid transparent;transition:background .12s;}
[data-testid="stSidebar"] [role="radiogroup"] label>div:first-child{display:none;}
[data-testid="stSidebar"] [role="radiogroup"] label:hover{background:#F1F5F9;}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked){background:var(--accent-l);border-left-color:var(--accent);}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) p{color:var(--accent);font-weight:600;}
.cours-actif{border:1px solid var(--line);border-left:3px solid var(--accent);border-radius:var(--r);
  padding:.6rem .8rem;background:var(--surface);margin-bottom:.8rem;}
.cours-actif span{display:block;font-size:.68rem;font-weight:700;letter-spacing:.08em;color:var(--muted);margin-bottom:.15rem;}
.cours-actif b{font-size:.85rem;font-weight:600;word-break:break-word;}
.promo{border-top:1px solid var(--line);padding-top:.8rem;font-size:.78rem;color:var(--muted);line-height:1.55;}
.promo b{color:var(--ink);}

/* ---------- En-têtes, étapes, cartes ---------- */
.entete{margin:0 0 1.2rem;padding-bottom:1rem;border-bottom:1px solid var(--line);}
.entete h1{font-size:1.7rem;font-weight:800;margin:0;padding:0;}
.entete p{margin:.35rem 0 0;color:var(--muted);font-size:.98rem;max-width:760px;line-height:1.55;}
.etapes{display:flex;flex-wrap:wrap;gap:.5rem;margin:0 0 1.3rem;}
.etape{display:inline-flex;align-items:center;gap:.5rem;padding:.3rem .85rem .3rem .35rem;border:1px solid var(--line);
  border-radius:999px;background:var(--surface);font-size:.82rem;font-weight:500;color:var(--muted);}
.etape i{font-style:normal;width:22px;height:22px;border-radius:50%;background:#F1F5F9;display:inline-flex;
  align-items:center;justify-content:center;font-size:.72rem;font-weight:700;}
.etape.ok{color:var(--ok);border-color:#A7F3D0;background:var(--ok-l);}
.etape.ok i{background:#10B981;color:#fff;}
.section-titre{font-size:1.05rem;font-weight:700;margin:1.8rem 0 .5rem;}
.vide{border:1.5px dashed var(--accent-m);background:var(--accent-l);border-radius:var(--r);padding:1.4rem 1.5rem;
  color:#312E81;font-size:.95rem;}
.vide b{display:block;font-size:1.02rem;margin-bottom:.2rem;}

/* ---------- Boutons et composants Streamlit ---------- */
.btn{display:inline-block;padding:.55rem 1rem;border-radius:10px;font-weight:600;font-size:.88rem;
  text-decoration:none!important;background:var(--accent);color:#fff!important;margin:.25rem .4rem .25rem 0;}
.btn:hover{background:var(--accent-d);} .btn.wa{background:#15803D;} .btn.wa:hover{background:#166534;}
.btn.clair{background:var(--accent-l);color:var(--accent)!important;}
.stButton>button,.stDownloadButton>button{border-radius:10px;font-weight:600;border:1px solid var(--line);transition:border-color .12s,color .12s;}
.stButton>button:hover,.stDownloadButton>button:hover{border-color:var(--accent);color:var(--accent);}
button[data-testid="stBaseButton-primary"],button[data-testid="stBaseButton-primaryFormSubmit"]{
  background:var(--accent);border:1px solid var(--accent);color:#fff;}
button[data-testid="stBaseButton-primary"]:hover,button[data-testid="stBaseButton-primaryFormSubmit"]:hover{
  background:var(--accent-d);border-color:var(--accent-d);color:#fff;}
div[data-testid="stMetric"]{background:var(--surface);border:1px solid var(--line);border-radius:var(--r);padding:.85rem 1.1rem;}
div[data-testid="stExpander"],div[data-testid="stForm"]{border:1px solid var(--line);border-radius:var(--r);background:var(--surface);}
[data-testid="stFileUploaderDropzone"]{border:1.5px dashed #CBD5E1;background:var(--surface);border-radius:var(--r);}
[data-testid="stVerticalBlockBorderWrapper"]{border-radius:var(--r);}
.stTabs [data-baseweb="tab"]{font-weight:600;}

/* ---------- Assistant (chat) ---------- */
[data-testid="stChatMessage"]{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:.9rem 1.1rem;}
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]){background:var(--accent-l);border-color:var(--accent-m);}
[data-testid="stChatInput"]{border-radius:14px;}
.sugg-titre{font-size:.8rem;font-weight:600;color:var(--muted);margin:.2rem 0 .5rem;}

/* ---------- Page de connexion ---------- */
.titre-accueil{font-size:2.45rem;font-weight:800;line-height:1.12;letter-spacing:-.03em;margin:1.4rem 0 .8rem;}
.titre-accueil em{font-style:normal;background:linear-gradient(120deg,#4338CA,#7C3AED);-webkit-background-clip:text;
  background-clip:text;color:transparent;}
.sous-accueil{color:#475569;font-size:1.04rem;line-height:1.65;max-width:540px;margin:0;}
.illus{margin:1.2rem 0 .4rem;max-width:520px;}
.illus svg{width:100%;height:auto;display:block;}
.chiffres{display:flex;flex-wrap:wrap;gap:1.6rem;margin:.6rem 0 0;}
.chiffres div{font-size:.82rem;color:var(--muted);line-height:1.3;}
.chiffres b{display:block;font-size:1.35rem;font-weight:800;color:var(--accent);letter-spacing:-.02em;}
.acces-titre{font-size:1.45rem;font-weight:800;margin:.4rem 0 .2rem;letter-spacing:-.02em;}
.acces-sous{color:var(--muted);font-size:.92rem;margin-bottom:1rem;line-height:1.5;}
.atouts{list-style:none;padding:0;margin:1.2rem 0 0;display:grid;gap:.9rem;}
.atouts li{display:flex;gap:.8rem;align-items:flex-start;}
.atouts .pt{flex:0 0 22px;height:22px;border-radius:50%;background:var(--accent-l);color:var(--accent);
  display:flex;align-items:center;justify-content:center;font-size:.72rem;font-weight:800;margin-top:.1rem;font-style:normal;}
.atouts b{display:block;font-size:.93rem;font-weight:600;}
.atouts span{color:#475569;font-size:.86rem;line-height:1.5;}
.confiance{font-size:.76rem;color:var(--muted);text-align:center;margin-top:.7rem;}

/* ---------- Services ---------- */
.services{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:1.1rem;margin:.6rem 0 1.6rem;}
.service{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:1.3rem 1.3rem 1.1rem;
  display:flex;flex-direction:column;transition:box-shadow .15s,border-color .15s;}
.service:hover{box-shadow:0 10px 28px rgba(67,56,202,.10);border-color:var(--accent-m);}
.service .ico{width:44px;height:44px;border-radius:12px;background:var(--accent-l);color:var(--accent);
  display:flex;align-items:center;justify-content:center;margin-bottom:.9rem;}
.service .ico svg{width:22px;height:22px;fill:none;stroke:currentColor;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round;}
.service h4{margin:0 0 .35rem;font-size:1.05rem;font-weight:700;}
.service>p{margin:0 0 .8rem;color:#475569;font-size:.88rem;line-height:1.6;}
.service ul{margin:0 0 1rem;padding:0;list-style:none;display:grid;gap:.35rem;}
.service li{font-size:.85rem;color:#334155;padding-left:1.15rem;position:relative;}
.service li::before{content:"";position:absolute;left:0;top:.5em;width:6px;height:6px;border-radius:50%;background:var(--accent);}
.service .contact{margin-top:auto;padding-top:.8rem;border-top:1px solid var(--line);font-size:.8rem;color:var(--muted);}
.service .contact b{color:var(--ink);}
.parcours{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:1rem;margin:.5rem 0 1.5rem;}
.parcours div{border-left:3px solid var(--accent);padding:.2rem 0 .2rem .9rem;}
.parcours b{display:block;font-size:.92rem;} .parcours span{font-size:.84rem;color:#475569;line-height:1.5;}

/* ---------- Pied de page (dans le flux : ne masque plus le contenu) ---------- */
.barre-pied{margin-top:2.5rem;padding:1.1rem 0 .3rem;border-top:1px solid var(--line);text-align:center;
  color:var(--muted);font-size:.78rem;line-height:1.8;}
.barre-pied a{color:var(--accent)!important;text-decoration:none;font-weight:500;}
.barre-pied a:hover{text-decoration:underline;} .barre-pied b{color:var(--ink);font-weight:600;}
.barre-pied .sep{opacity:.35;margin:0 .6rem;}

@media (max-width:640px){
  .block-container{padding:1.2rem 1rem 2rem;} .entete h1{font-size:1.4rem;} .titre-accueil{font-size:1.8rem;}
  .barre-pied .sep{display:none;} .barre-pied a,.barre-pied b{display:block;}}
</style>
"""

# Illustration : étudiant à son bureau, ordinateur ouvert sur ses fiches, entouré d'un résumé, d'une note de QCM
# et d'une bulle de discussion. Dessin original, sans texte propriétaire.
ILLUSTRATION = """<svg viewBox="0 0 580 440" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Un étudiant révise sur son ordinateur : résumé, QCM noté et assistant IA">
<defs>
<linearGradient id="g1" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#EEF2FF"/><stop offset="1" stop-color="#E0E7FF"/></linearGradient>
<filter id="sh" x="-20%" y="-20%" width="140%" height="150%"><feDropShadow dx="0" dy="6" stdDeviation="8" flood-color="#312E81" flood-opacity=".16"/></filter>
</defs>
<circle cx="285" cy="235" r="196" fill="url(#g1)"/>
<circle cx="470" cy="80" r="26" fill="#C7D2FE" opacity=".55"/>
<circle cx="70" cy="300" r="14" fill="#FDE68A" opacity=".8"/>
<rect x="40" y="408" width="480" height="6" rx="3" fill="#C7D2FE"/>
<!-- plante -->
<g><path d="M62 408h34l-4-34H66z" fill="#F59E0B"/><path d="M79 374c-22-6-30-30-22-50 14 8 24 26 22 50z" fill="#10B981"/><path d="M79 374c4-24 20-38 38-40-2 22-16 36-38 40z" fill="#059669"/><path d="M79 374c-2-28 6-48 18-60 6 24 0 46-18 60z" fill="#34D399"/></g>
<!-- bureau -->
<rect x="236" y="290" width="296" height="12" rx="6" fill="#312E81"/>
<rect x="256" y="302" width="10" height="106" fill="#4338CA"/><rect x="502" y="302" width="10" height="106" fill="#4338CA"/>
<!-- livres -->
<rect x="470" y="274" width="56" height="16" rx="3" fill="#F59E0B"/><rect x="474" y="260" width="48" height="14" rx="3" fill="#10B981"/><rect x="468" y="247" width="52" height="13" rx="3" fill="#6366F1"/>
<!-- ordinateur -->
<g filter="url(#sh)"><rect x="300" y="186" width="156" height="96" rx="9" fill="#1E293B"/>
<rect x="308" y="194" width="140" height="80" rx="4" fill="#fff"/>
<rect x="308" y="194" width="140" height="13" rx="4" fill="#4338CA"/><circle cx="316" cy="200.5" r="2" fill="#C7D2FE"/><circle cx="323" cy="200.5" r="2" fill="#C7D2FE"/>
<rect x="316" y="216" width="58" height="5" rx="2.5" fill="#CBD5E1"/><rect x="316" y="226" width="70" height="5" rx="2.5" fill="#E2E8F0"/><rect x="316" y="236" width="50" height="5" rx="2.5" fill="#E2E8F0"/><rect x="316" y="246" width="64" height="5" rx="2.5" fill="#E2E8F0"/>
<rect x="394" y="244" width="8" height="20" rx="2" fill="#A5B4FC"/><rect x="406" y="234" width="8" height="30" rx="2" fill="#6366F1"/><rect x="418" y="224" width="8" height="40" rx="2" fill="#4338CA"/><rect x="430" y="214" width="8" height="50" rx="2" fill="#10B981"/>
<rect x="288" y="282" width="180" height="8" rx="4" fill="#CBD5E1"/></g>
<!-- chaise -->
<rect x="112" y="226" width="12" height="112" rx="6" fill="#4338CA"/><rect x="112" y="336" width="110" height="12" rx="6" fill="#4338CA"/><rect x="162" y="348" width="10" height="60" fill="#312E81"/><rect x="132" y="404" width="70" height="8" rx="4" fill="#312E81"/>
<!-- jambes -->
<path d="M170 322h86a14 14 0 0 1 14 14v6a14 14 0 0 1-14 14h-86z" fill="#1E293B"/>
<path d="M246 340h28v62h-28z" fill="#1E293B"/><path d="M240 400h46a8 8 0 0 1 8 8v2H240z" fill="#F8FAFC" stroke="#CBD5E1"/>
<!-- buste -->
<path d="M160 322c-6-40-2-82 22-102 12-9 36-6 44 8 8 16 4 40-6 62l-4 32z" fill="#6366F1"/>
<path d="M200 232c10-8 26 0 28 12" fill="none" stroke="#4F46E5" stroke-width="2" opacity=".6"/>
<!-- bras et main -->
<path d="M196 238c26 2 44 18 52 42l50 4" fill="none" stroke="#4338CA" stroke-width="22" stroke-linecap="round" stroke-linejoin="round"/>
<circle cx="302" cy="284" r="10" fill="#8D5524"/>
<!-- cou, tête, cheveux, casque -->
<rect x="196" y="196" width="22" height="30" rx="10" fill="#8D5524"/>
<circle cx="212" cy="172" r="32" fill="#8D5524"/>
<path d="M180 168c-4-30 18-46 40-42 20 4 30 20 26 40-8-10-20-16-34-14-12 2-24 8-32 16z" fill="#1F1A17"/>
<path d="M243 176c6 2 10 6 8 10l-6 2z" fill="#7A4A1F"/>
<circle cx="234" cy="172" r="2.6" fill="#1F1A17"/>
<path d="M226 190c5 3 11 3 15-1" fill="none" stroke="#5B3412" stroke-width="2.2" stroke-linecap="round"/>
<path d="M182 164c0-26 20-42 42-36" fill="none" stroke="#0F172A" stroke-width="5" stroke-linecap="round"/><rect x="176" y="164" width="12" height="22" rx="6" fill="#0F172A"/>
<!-- carte résumé -->
<g filter="url(#sh)"><rect x="330" y="62" width="150" height="92" rx="14" fill="#fff"/>
<rect x="346" y="78" width="26" height="26" rx="7" fill="#EEF2FF"/><path d="M353 86h12M353 91h12M353 96h7" stroke="#4338CA" stroke-width="2" stroke-linecap="round"/>
<text x="382" y="89" font-size="13" font-weight="700" fill="#0F172A">Résumé</text><text x="382" y="103" font-size="10" fill="#64748B">12 chapitres</text>
<rect x="346" y="118" width="118" height="5" rx="2.5" fill="#E2E8F0"/><rect x="346" y="130" width="84" height="5" rx="2.5" fill="#E2E8F0"/></g>
<!-- note QCM -->
<g transform="translate(22,0)" filter="url(#sh)"><rect x="440" y="168" width="96" height="58" rx="14" fill="#fff"/>
<circle cx="466" cy="197" r="13" fill="#10B981"/><path d="M459.5 197l5 5 8-9" fill="none" stroke="#fff" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>
<text x="486" y="195" font-size="15" font-weight="800" fill="#0F172A">18/20</text><text x="486" y="209" font-size="9.5" fill="#64748B">Réussi</text></g>
<!-- bulle assistant -->
<g transform="translate(0,-44)" filter="url(#sh)"><path d="M70 96a16 16 0 0 1 16-16h112a16 16 0 0 1 16 16v38a16 16 0 0 1-16 16h-74l-22 18v-18h-16a16 16 0 0 1-16-16z" fill="#4338CA"/>
<circle cx="108" cy="115" r="5" fill="#fff"/><circle cx="130" cy="115" r="5" fill="#C7D2FE"/><circle cx="152" cy="115" r="5" fill="#A5B4FC"/>
<text x="176" y="119" font-size="10" font-weight="600" fill="#E0E7FF">IA</text></g>
</svg>"""


def appliquer_style():
    st.markdown(CSS, unsafe_allow_html=True)
