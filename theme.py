"""Thème Étudiants Learn : sobre, lisible, léger.
Usage dans app.py :  from theme import appliquer_theme  puis, juste après st.set_page_config(...) :  appliquer_theme()
Les noms de classes CSS sont identiques à l'ancienne version : aucune autre modification de l'interface n'est requise."""
import streamlit as st

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=Source+Serif+4:opsz,wght@8..60,600;8..60,700&display=swap');
:root{
  --ink:#16202E; --muted:#5B6678; --line:#DDE2EA; --paper:#F5F6F8; --surface:#FFFFFF;
  --accent:#1B4B7A; --accent-d:#143A60; --accent-l:#E8EFF7; --gold:#B7822E;
  --ok:#1E7B4F; --ok-l:#EAF6EF; --r:10px;
  --serif:'Source Serif 4',Georgia,serif; --sans:'IBM Plex Sans',system-ui,-apple-system,'Segoe UI',sans-serif;
}
/* Fond uni : plus de motifs SVG, de dégradés ni de flou (ils ralentissent le rendu, surtout sur mobile) */
.stApp{font-family:var(--sans);color:var(--ink);background:var(--paper);}
#MainMenu,footer{visibility:hidden;} header[data-testid="stHeader"]{background:transparent;}
iframe[height="0"],iframe[height="1"]{display:none;}
.block-container{max-width:1080px;padding:2.2rem 2.2rem 3rem;}
h1,h2,h3{font-family:var(--serif);letter-spacing:-.01em;color:var(--ink);}
h2,h3{font-weight:700;}
a:focus-visible,button:focus-visible,input:focus-visible{outline:2px solid var(--accent);outline-offset:2px;}
@media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important;}}

/* Barre latérale */
[data-testid="stSidebar"]{background:var(--surface);border-right:1px solid var(--line);}
.marque,.marque-lg{display:flex;align-items:center;gap:.65rem;}
.marque{margin:.2rem 0 1.1rem;}
.logo{width:34px;height:34px;border-radius:8px;background:var(--accent);color:#fff;display:flex;
  align-items:center;justify-content:center;font-family:var(--serif);font-weight:700;font-size:.9rem;}
.marque-lg .logo{width:40px;height:40px;font-size:1rem;}
.marque b,.marque-lg b{font-family:var(--serif);font-size:1.15rem;font-weight:700;color:var(--ink);}
.profil{display:flex;align-items:center;gap:.65rem;padding:.6rem .75rem;border:1px solid var(--line);
  border-radius:var(--r);background:var(--paper);margin-bottom:1rem;}
.profil .av{flex:0 0 32px;height:32px;border-radius:50%;background:var(--accent);color:#fff;font-weight:600;
  display:flex;align-items:center;justify-content:center;font-size:.85rem;}
.profil .n{font-weight:600;font-size:.88rem;line-height:1.25;}
.profil .e{font-size:.74rem;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:150px;}
[data-testid="stSidebar"] [role="radiogroup"]{gap:.1rem;}
[data-testid="stSidebar"] [role="radiogroup"] label{padding:.5rem .75rem;border-radius:8px;width:100%;
  border-left:3px solid transparent;}
[data-testid="stSidebar"] [role="radiogroup"] label>div:first-child{display:none;}
[data-testid="stSidebar"] [role="radiogroup"] label:hover{background:var(--paper);}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked){background:var(--accent-l);border-left-color:var(--accent);}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) p{color:var(--accent);font-weight:600;}
.cours-actif{border:1px solid var(--line);border-left:3px solid var(--gold);border-radius:var(--r);
  padding:.6rem .8rem;background:var(--surface);margin-bottom:.8rem;}
.cours-actif span{display:block;font-size:.72rem;font-weight:600;color:var(--muted);margin-bottom:.1rem;}
.cours-actif b{font-size:.85rem;font-weight:600;word-break:break-word;}
.promo{border-top:1px solid var(--line);padding-top:.8rem;font-size:.78rem;color:var(--muted);line-height:1.55;}
.promo b{color:var(--ink);}

/* En-tête de page et progression */
.entete{margin:0 0 1.3rem;padding-bottom:.9rem;border-bottom:1px solid var(--line);}
.entete h1{font-size:1.75rem;font-weight:700;margin:0;padding:0;}
.entete p{margin:.35rem 0 0;color:var(--muted);font-size:.97rem;max-width:720px;line-height:1.55;}
.etapes{display:flex;flex-wrap:wrap;gap:.45rem;margin:0 0 1.3rem;}
.etape{display:inline-flex;align-items:center;gap:.5rem;padding:.28rem .8rem .28rem .3rem;border:1px solid var(--line);
  border-radius:999px;background:var(--surface);font-size:.82rem;font-weight:500;color:var(--muted);}
.etape i{font-style:normal;width:22px;height:22px;border-radius:50%;background:var(--paper);display:inline-flex;
  align-items:center;justify-content:center;font-size:.72rem;font-weight:600;}
.etape.ok{color:var(--ok);border-color:#BFE3CF;background:var(--ok-l);}
.etape.ok i{background:var(--ok);color:#fff;}

/* Contenu */
.cartes{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:1rem;margin:.5rem 0 1.5rem;}
.carte{border:1px solid var(--line);border-radius:var(--r);padding:1.1rem 1.2rem;background:var(--surface);}
.carte .tag{font-size:.78rem;font-weight:600;color:var(--accent);margin-bottom:.4rem;}
.carte h4{margin:0 0 .35rem;font-size:1rem;font-weight:600;}
.carte p{margin:0;color:var(--muted);font-size:.88rem;line-height:1.55;}
.section-titre{font-family:var(--serif);font-size:1.2rem;font-weight:700;margin:2rem 0 .5rem;}

/* Page de connexion */
.titre-accueil{font-family:var(--serif);font-size:2.4rem;font-weight:700;line-height:1.15;letter-spacing:-.02em;margin:1.6rem 0 .8rem;}
.sous-accueil{color:var(--muted);font-size:1.02rem;line-height:1.65;max-width:520px;margin:0;}
.atouts{list-style:none;padding:0;margin:1.6rem 0 0;display:grid;gap:1rem;}
.atouts li{display:flex;gap:.85rem;align-items:flex-start;}
.atouts .pt{flex:0 0 3px;align-self:stretch;background:var(--gold);border-radius:2px;}
.atouts b{display:block;font-size:.95rem;font-weight:600;}
.atouts span{color:var(--muted);font-size:.88rem;line-height:1.5;}
.acces-titre{font-family:var(--serif);font-size:1.3rem;font-weight:700;margin:.2rem 0;}
.acces-sous{color:var(--muted);font-size:.9rem;margin-bottom:.8rem;}

/* Boutons et composants Streamlit */
.btn{display:inline-block;padding:.5rem 1rem;border-radius:8px;font-weight:600;font-size:.88rem;
  text-decoration:none!important;background:var(--accent);color:#fff!important;margin:.25rem .4rem .25rem 0;}
.btn:hover{background:var(--accent-d);} .btn.wa{background:var(--ok);} .btn.wa:hover{background:#176640;}
.btn.clair{background:var(--accent-l);color:var(--accent)!important;}
.stButton>button,.stDownloadButton>button{border-radius:8px;font-weight:600;border:1px solid var(--line);
  font-family:var(--sans);}
.stButton>button:hover,.stDownloadButton>button:hover{border-color:var(--accent);color:var(--accent);}
button[data-testid="stBaseButton-primary"],button[data-testid="stBaseButton-primaryFormSubmit"]{
  background:var(--accent);border:1px solid var(--accent);color:#fff;}
button[data-testid="stBaseButton-primary"]:hover,button[data-testid="stBaseButton-primaryFormSubmit"]:hover{
  background:var(--accent-d);border-color:var(--accent-d);color:#fff;}
div[data-testid="stMetric"]{background:var(--surface);border:1px solid var(--line);border-radius:var(--r);padding:.85rem 1.1rem;}
div[data-testid="stExpander"],div[data-testid="stForm"],[data-testid="stChatMessage"]{
  border:1px solid var(--line);border-radius:var(--r);background:var(--surface);}
[data-testid="stFileUploaderDropzone"]{border:1.5px dashed #B9C2D0;background:var(--surface);border-radius:var(--r);}
[data-testid="stVerticalBlockBorderWrapper"]{border-radius:var(--r);}
.stTabs [data-baseweb="tab"]{font-weight:600;}

/* Services */
.services{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:1rem;margin:.6rem 0 1.6rem;}
.service{background:var(--surface);border:1px solid var(--line);border-top:3px solid var(--accent);
  border-radius:var(--r);padding:1.25rem 1.3rem 1.05rem;display:flex;flex-direction:column;}
.service .ico{width:40px;height:40px;border-radius:8px;background:var(--accent-l);color:var(--accent);
  display:flex;align-items:center;justify-content:center;margin-bottom:.8rem;}
.service .ico svg{width:21px;height:21px;fill:none;stroke:currentColor;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round;}
.service h4{margin:0 0 .35rem;font-family:var(--serif);font-size:1.12rem;font-weight:700;}
.service>p{margin:0 0 .8rem;color:var(--muted);font-size:.88rem;line-height:1.6;}
.service ul{margin:0 0 1rem;padding:0;list-style:none;display:grid;gap:.35rem;}
.service li{font-size:.85rem;padding-left:1.1rem;position:relative;}
.service li::before{content:"";position:absolute;left:0;top:.55em;width:6px;height:2px;background:var(--gold);}
.service .contact{margin-top:auto;padding-top:.8rem;border-top:1px solid var(--line);font-size:.8rem;color:var(--muted);}
.service .contact b{color:var(--ink);}
.parcours{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:1rem;margin:.5rem 0 1.5rem;}
.parcours div{border-left:3px solid var(--accent);padding:.2rem 0 .2rem .9rem;}
.parcours b{display:block;font-size:.92rem;} .parcours span{font-size:.84rem;color:var(--muted);line-height:1.5;}

/* Pied de page : dans le flux (ne masque plus le contenu) */
.barre-pied{margin-top:2.5rem;padding:1rem 0 .2rem;border-top:1px solid var(--line);text-align:center;
  color:var(--muted);font-size:.78rem;line-height:1.7;}
.barre-pied a{color:var(--accent)!important;text-decoration:none;font-weight:500;}
.barre-pied a:hover{text-decoration:underline;} .barre-pied b{color:var(--ink);font-weight:600;}
.barre-pied .sep{opacity:.35;margin:0 .55rem;}
@media (max-width:640px){
  .block-container{padding:1.2rem 1rem 2rem;} .entete h1{font-size:1.4rem;} .titre-accueil{font-size:1.8rem;}
  .barre-pied .sep{display:none;} .barre-pied a,.barre-pied b{display:block;}}
</style>
"""

CONFIG_TOML = """# .streamlit/config.toml  (thème de base + démarrage plus rapide)
[theme]
base = "light"
primaryColor = "#1B4B7A"
backgroundColor = "#F5F6F8"
secondaryBackgroundColor = "#FFFFFF"
textColor = "#16202E"
font = "sans serif"

[server]
fileWatcherType = "none"
runOnSave = false

[browser]
gatherUsageStats = false
"""


def appliquer_theme():
    st.markdown(CSS, unsafe_allow_html=True)
