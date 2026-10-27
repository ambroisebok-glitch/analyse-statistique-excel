#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Carte interactive des coopératives (enquête CAT) - Golfe, Zio, Lacs (Togo)

Installation (une seule fois) :
    pip install pandas openpyxl folium

Utilisation (le fichier Excel doit être dans le même dossier) :
    python carte_cooperatives.py                    # lit coops_lome.xlsx
    python carte_cooperatives.py mon_fichier.xlsx   # ou un autre fichier

Résultats :
    carte_cooperatives.html      -> à ouvrir dans le navigateur (internet requis)
    synthese_cooperatives.xlsx   -> tableau récapitulatif par coopérative

Ce que montre la carte :
  - un BADGE rond par coopérative et par site, avec son score /100 dedans,
    coloré du rouge (faible) au vert (élevé) ; le nom s'affiche en zoomant ;
  - un POINT par enquêté (position GPS), coloré selon son propre score ;
  - un panneau de classement : cliquer sur une coopérative zoome dessus ;
  - un menu (en haut à droite) pour choisir le fond de carte et pour
    afficher/masquer chaque coopérative ;
  - au clic sur un badge : détail des scores (création, 7 principes, fonctionnement).
"""
import sys
import json
import math
import html as H

import pandas as pd
import folium
from folium.plugins import Fullscreen, MiniMap, MeasureControl
import branca.colormap as cm
from branca.element import Element

# =============================================================== PARAMÈTRES
FICHIER = sys.argv[1] if len(sys.argv) > 1 else "coops_lome.xlsx"
SORTIE_CARTE = "carte_cooperatives.html"
SORTIE_SYNTHESE = "synthese_cooperatives.xlsx"

# Noms des feuilles du classeur
FEUILLE_TEXTE = "QUESTIONNAIRE CAT MEMOIRE 20..."  # réponses en texte + GPS
FEUILLE_NOTES = "QUESTIONNAIRE CAT MEMOIRE 20"     # réponses chiffrées
FEUILLE_MOYENNES = "Score_moyen (2)"               # score_1, Score_2, Score_3

# Score final = moyenne de score_1, Score_2 et Score_3, multipliée par 10.
# ATTENTION : score_1 est noté sur 5 alors que Score_2 et Score_3 sont sur 10,
# donc le maximum possible est 83,3/100. Mettre True pour doubler score_1
# (le ramener sur 10) afin que le maximum soit 100/100.
NORMALISER_SCORE_1 = False

# Catégories (scores /100). Seuils déduits de vos résultats : à ajuster si besoin.
SEUIL_FRAGILE = 50   # en dessous -> "Fragile"
SEUIL_HAUTE = 75     # à partir de -> "Haute Performance" ; entre les deux -> "Modéré"

RAYON_SITE_KM = 3    # points d'une même coopérative à moins de 3 km = même site
ZOOM_ETIQUETTES = 13 # les noms s'affichent à partir de ce niveau de zoom

# Variantes d'écriture à regrouper (en MAJUSCULES, espaces nettoyés)
ALIAS = {
    "AGAPÉ": "AGAPE",
    "ASSILASSIMÉ": "ASSILASIME",
    "ASSILASIMÉ": "ASSILASIME",
    "MOPKOKPO": "MOKPOKPO",
    "VERT DE NATURE": "VERS DE NATURE",
}

# Noms des colonnes
C_NOM = "Nom de la coopérative"
C_LAT = "_Localisation / Enregistrez les coordonnées GPS_latitude"
C_LON = "_Localisation / Enregistrez les coordonnées GPS_longitude"
C_PREF = "Préfecture"
C_CANTON = "Canton de l’enquêté"
C_VILLAGE = "Village de l’enquêté"
C_TYPE = "Type de coopérative"
C_ANNEE = "Année de création de la coopérative"
C_MEMBRES = "Nombre de membres au sein de la coopérative"
C_FEMMES = "Nombre de femmes au sein de la coopérative"

PRINCIPES = [
    ("Adhésion volontaire et ouverte", "Adhésion volontaire"),
    ("Pouvoir démocratique dans la prise de décisions", "Pouvoir démocratique"),
    ("Participation économique des membres", "Participation économique"),
    ("Autonomie et indépendance vis-à-vis des partenaires", "Autonomie"),
    ("Éducation et formation des membres", "Éducation / formation"),
    ("Coopération entre coopératives", "Coopération inter-coop."),
    ("Engagement communautaire et développement durable", "Engagement communautaire"),
]


# =============================================================== OUTILS
def norm(nom):
    n = " ".join(str(nom).split()).upper()
    return ALIAS.get(n, n)


def categorie(score):
    if score < SEUIL_FRAGILE:
        return "Fragile"
    if score >= SEUIL_HAUTE:
        return "Haute Performance"
    return "Modéré"


def km(a, b):
    p1, p2 = math.radians(a[0]), math.radians(b[0])
    dl = math.radians(b[1] - a[1])
    h = math.sin((p2 - p1) / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * 6371 * math.asin(math.sqrt(h))


def sites(df):
    """Regroupe les points d'une coopérative en 'sites' géographiques."""
    res = []
    for _, r in df.iterrows():
        p = (r["lat"], r["lon"])
        for s in res:
            if km((s["lat"], s["lon"]), p) <= RAYON_SITE_KM:
                s["pts"].append(p)
                s["villages"].add(r["village"])
                s["lat"] = sum(q[0] for q in s["pts"]) / len(s["pts"])
                s["lon"] = sum(q[1] for q in s["pts"]) / len(s["pts"])
                break
        else:
            res.append({"lat": p[0], "lon": p[1], "pts": [p], "villages": {r["village"]}})
    return res


def hex6(c):
    return c[:7]


def couleur_texte(hexc):
    r, g, b = (int(hexc[i:i + 2], 16) for i in (1, 3, 5))
    return "#fff" if (0.299 * r + 0.587 * g + 0.114 * b) < 140 else "#111"


def e(x):
    return H.escape(str(x))


# =============================================================== LECTURE
print(f"Lecture de {FICHIER} ...")
try:
    xl = pd.read_excel(FICHIER, sheet_name=None)
except FileNotFoundError:
    sys.exit(f"Fichier introuvable : {FICHIER} (placez-le dans le même dossier que le script)")

for f in (FEUILLE_TEXTE, FEUILLE_NOTES, FEUILLE_MOYENNES):
    if f not in xl:
        sys.exit(f"Feuille '{f}' introuvable. Feuilles disponibles : {list(xl)}")

txt, num, moy = xl[FEUILLE_TEXTE], xl[FEUILLE_NOTES], xl[FEUILLE_MOYENNES]
if not (len(txt) == len(num) == len(moy)):
    sys.exit("Les feuilles n'ont pas le même nombre de lignes.")
if not (txt[C_NOM].map(norm).values == moy[C_NOM].map(norm).values).all():
    sys.exit("Les lignes de 'Score_moyen (2)' ne correspondent pas à celles du questionnaire.")

d = pd.DataFrame({
    "coop": txt[C_NOM].map(norm),
    "lat": txt[C_LAT], "lon": txt[C_LON],
    "pref": txt[C_PREF].astype(str).str.strip(),
    "canton": txt[C_CANTON].astype(str).str.strip(),
    "village": txt[C_VILLAGE].astype(str).str.strip(),
    "type": txt[C_TYPE].astype(str).str.strip(),
    "annee": txt[C_ANNEE],
    "membres": txt[C_MEMBRES], "femmes": txt[C_FEMMES],
    "s1": moy["score_1"], "s2": moy["Score_2"], "s3": moy["Score_3"],
})
for long, court in PRINCIPES:
    d[court] = num[long].values

f1 = 2 if NORMALISER_SCORE_1 else 1
d["score"] = (d["s1"] * f1 + d["s2"] + d["s3"]) / 3 * 10

avant = len(d)
d = d.dropna(subset=["lat", "lon"])
d = d[d["lat"].between(-90, 90) & d["lon"].between(-180, 180) & ~((d["lat"] == 0) & (d["lon"] == 0))]
if len(d) < avant:
    print(f"  ! {avant - len(d)} ligne(s) sans GPS valide ignorée(s)")

# =============================================================== AGRÉGATION
agg = d.groupby("coop").agg(
    n=("lat", "size"),
    score=("score", "mean"),
    s1=("s1", "mean"), s2=("s2", "mean"), s3=("s3", "mean"),
    membres=("membres", "median"), femmes=("femmes", "median"),
    annee=("annee", "median"),
    type=("type", lambda v: v.mode().iloc[0]),
    pref=("pref", lambda v: ", ".join(sorted(set(v)))),
    villages=("village", lambda v: ", ".join(sorted(set(v)))),
    **{c: (c, "mean") for _, c in PRINCIPES},
).sort_values("score", ascending=False)
agg["categorie"] = agg["score"].map(categorie)
print(f"  {len(d)} enquêtés, {len(agg)} coopératives :")
print("   " + " | ".join(agg.index))

# Synthèse Excel
syn = agg.reset_index().rename(columns={
    "coop": "Coopérative", "categorie": "Catégorie", "score": "Score /100", "n": "Nb enquêtés",
    "s1": "Score création (/5)", "s2": "Score principes (/10)", "s3": "Score fonctionnement (/10)",
    "membres": "Membres (médiane)", "femmes": "Femmes (médiane)", "annee": "Année création",
    "type": "Type", "pref": "Préfecture", "villages": "Villages",
})
cols = ["Coopérative", "Catégorie", "Score /100", "Score création (/5)", "Score principes (/10)",
        "Score fonctionnement (/10)", "Nb enquêtés", "Type", "Préfecture", "Villages",
        "Membres (médiane)", "Femmes (médiane)", "Année création"] + [c for _, c in PRINCIPES]
syn = syn[cols].round(2)
with pd.ExcelWriter(SORTIE_SYNTHESE, engine="openpyxl") as w:
    syn.to_excel(w, index=False, sheet_name="Synthèse")
    ws = w.sheets["Synthèse"]
    for i, c in enumerate(syn.columns, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = min(max(len(str(c)), 12) + 2, 45)
    ws.freeze_panes = "B2"

# =============================================================== CARTE
vmin = math.floor(d["score"].min() / 5) * 5
vmax = math.ceil(d["score"].max() / 5) * 5
cmap = cm.LinearColormap(["#d7191c", "#fdae61", "#ffffbf", "#a6d96a", "#1a9641"], vmin=vmin, vmax=vmax)
cmap.caption = "Score de gouvernance (/100)"

m = folium.Map(location=[d["lat"].mean(), d["lon"].mean()], zoom_start=10, tiles=None, control_scale=True)
folium.TileLayer("OpenStreetMap", name="OpenStreetMap").add_to(m)
folium.TileLayer(
    tiles="https://tile.opentopomap.org/{z}/{x}/{y}.png", max_native_zoom=17,
    attr="OpenTopoMap (CC-BY-SA)", name="Relief (OpenTopoMap)").add_to(m)
folium.TileLayer(
    tiles="https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
    attr="Esri, Maxar, Earthstar Geographics", name="Satellite (Esri)").add_to(m)
m.fit_bounds([[d["lat"].min(), d["lon"].min()], [d["lat"].max(), d["lon"].max()]], padding=(30, 30))

# --- 1) points GPS par enquêté, un calque par coopérative
noms_calques = []
for coop in agg.index:
    fg = folium.FeatureGroup(name=f"{coop} — {agg.loc[coop, 'score']:.0f}", show=True)
    for _, r in d[d["coop"] == coop].iterrows():
        col = hex6(cmap(r["score"]))
        pop = (f"<div style='font:12px Arial'><b>{e(coop)}</b><br>"
               f"Village : {e(r['village'])} ({e(r['canton'])})<br>"
               f"Score de cet enquêté : <b>{r['score']:.0f}/100</b></div>")
        folium.CircleMarker(
            [r["lat"], r["lon"]], radius=5, color="#333", weight=1, fill=True,
            fill_color=col, fill_opacity=0.9,
            tooltip=f"{coop} — enquêté : {r['score']:.0f}/100",
            popup=folium.Popup(pop, max_width=260)).add_to(fg)
    fg.add_to(m)
    noms_calques.append(fg.get_name())

# --- 2) badges de score par coopérative et par site
badges = folium.FeatureGroup(name="★ Badges de score (par coopérative)", show=True)
coops_js = []
for coop, a in agg.iterrows():
    sous = d[d["coop"] == coop]
    col = hex6(cmap(a["score"]))
    barres = "".join(
        f"<tr><td style='padding:1px 6px 1px 0'>{c}</td>"
        f"<td style='width:90px'><div style='background:#e5e5e5;height:8px'>"
        f"<div style='background:#2b6cb0;height:8px;width:{min(a[c], 10) * 10:.0f}%'></div></div></td>"
        f"<td style='padding-left:6px'>{a[c]:.1f}</td></tr>" for _, c in PRINCIPES)
    for s in sites(sous):
        pop = f"""<div style="font:12px Arial;min-width:260px">
          <div style="font-size:15px;font-weight:700">{e(coop)}</div>
          <div style="margin:2px 0 6px"><span style="background:{col};color:{couleur_texte(col)};
            padding:2px 8px;border-radius:10px;font-weight:700">{a['score']:.1f}/100</span>
            &nbsp;{e(a['categorie'])}</div>
          <b>Type :</b> {e(a['type'])} &nbsp; <b>Créée :</b> {a['annee']:.0f}<br>
          <b>Membres :</b> ~{a['membres']:.0f} (dont ~{a['femmes']:.0f} femmes)<br>
          <b>Préfecture :</b> {e(a['pref'])}<br>
          <b>Site :</b> {e(', '.join(sorted(s['villages'])))} ({len(s['pts'])} enquêté(s) ici, {a['n']} au total)
          <hr style="margin:6px 0">
          <b>Création</b> : {a['s1']:.1f}/5 &nbsp; <b>Fonctionnement</b> : {a['s3']:.1f}/10<br>
          <b>7 principes coopératifs</b> ({a['s2']:.1f}/10)
          <table style="font-size:11px;margin-top:3px">{barres}</table></div>"""
        ic = folium.DivIcon(
            html=(f'<div class="cat-badge" style="background:{col};color:{couleur_texte(col)}">{a["score"]:.0f}</div>'
                  f'<div class="cat-lbl">{e(coop)}</div>'),
            icon_size=(36, 36), icon_anchor=(18, 18), class_name="cat-icon")
        folium.Marker([s["lat"], s["lon"]], icon=ic,
                      tooltip=f"{coop} — {a['score']:.1f}/100 · {a['categorie']}",
                      popup=folium.Popup(pop, max_width=340)).add_to(badges)
    coops_js.append({"b": [[sous["lat"].min(), sous["lon"].min()], [sous["lat"].max(), sous["lon"].max()]]})
badges.add_to(m)

Fullscreen(position="topleft").add_to(m)
MeasureControl(primary_length_unit="kilometers", position="topleft").add_to(m)
MiniMap(toggle_display=True, position="bottomright").add_to(m)
cmap.add_to(m)
folium.LayerControl(collapsed=True, position="topright").add_to(m)

# --- 3) panneau de classement + style + scripts
rows = ""
for i, (coop, a) in enumerate(agg.iterrows()):
    col = hex6(cmap(a["score"]))
    rows += (f'<div class="rk" onclick="goCoop({i})"><span class="dot" style="background:{col}"></span>'
             f'<b>{e(coop)}</b><span class="sc">{a["score"]:.0f}</span>'
             f'<small>{e(a["categorie"])}</small></div>')

panel = f"""
<style>
 .cat-icon{{background:none;border:none}}
 .cat-badge{{width:36px;height:36px;border-radius:50%;border:2px solid #fff;
   box-shadow:0 1px 5px rgba(0,0,0,.6);font:700 13px/32px Arial;text-align:center}}
 .cat-lbl{{position:absolute;left:50%;top:38px;transform:translateX(-50%);white-space:nowrap;
   font:700 11px Arial;color:#222;text-shadow:0 0 3px #fff,0 0 3px #fff,0 0 3px #fff,0 0 3px #fff}}
 .hide-labels .cat-lbl{{display:none}}
 .leaflet-control-layers-list{{max-height:60vh;overflow:auto}}
 #cat-panel{{position:fixed;top:120px;left:10px;z-index:1000;width:250px;background:#fffffff2;
   border-radius:6px;box-shadow:0 1px 6px rgba(0,0,0,.4);font:12px Arial}}
 #cat-panel summary{{cursor:pointer;padding:7px 10px;font-weight:700;font-size:13px}}
 #cat-list{{max-height:48vh;overflow:auto;border-top:1px solid #ddd}}
 .rk{{display:flex;align-items:center;gap:6px;padding:4px 10px;cursor:pointer;border-bottom:1px solid #f0f0f0}}
 .rk:hover{{background:#eef4ff}}
 .rk b{{flex:1}} .rk small{{color:#777;width:62px;text-align:right;font-size:10px}}
 .dot{{width:12px;height:12px;border-radius:50%;border:1px solid #555;flex:none}}
 .sc{{font-weight:700;width:24px;text-align:right}}
 #cat-btns{{padding:5px 10px;border-top:1px solid #ddd}}
 #cat-btns button{{font:11px Arial;margin-right:4px;cursor:pointer}}
</style>
<details id="cat-panel" open>
  <summary>Classement des coopératives (score /100)</summary>
  <div id="cat-list">{rows}</div>
  <div id="cat-btns">Points GPS :
    <button onclick="allPts(true)">tout afficher</button><button onclick="allPts(false)">tout masquer</button>
  </div>
</details>
<script>
var COOPS = {json.dumps(coops_js)};
var PTS = {json.dumps(noms_calques)};
function goCoop(i) {{
  var b = COOPS[i].b;
  {m.get_name()}.flyToBounds(b, {{maxZoom: 17, padding: [60, 60]}});
}}
function allPts(on) {{
  PTS.forEach(function(n) {{ var l = window[n];
    if (on) {{ {m.get_name()}.addLayer(l); }} else {{ {m.get_name()}.removeLayer(l); }} }});
}}
window.addEventListener('load', function() {{
  var mp = {m.get_name()}, el = mp.getContainer();
  function maj() {{ el.classList.toggle('hide-labels', mp.getZoom() < {ZOOM_ETIQUETTES}); }}
  mp.on('zoomend', maj); maj();
  var p = document.getElementById('cat-panel');
  L.DomEvent.disableClickPropagation(p); L.DomEvent.disableScrollPropagation(p);
}});
</script>
"""
m.get_root().html.add_child(Element(panel))
m.save(SORTIE_CARTE)

print(f"\nTerminé :\n  - {SORTIE_CARTE}  (ouvrir dans Chrome / Firefox / Edge)\n  - {SORTIE_SYNTHESE}")
