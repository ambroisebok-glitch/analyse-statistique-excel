# -*- coding: utf-8 -*-
"""
Carte de zone d'étude "style article" :
  - carte principale : préfectures concernées colorées, points GPS en carrés rouges,
    graticule en degrés-minutes, rose des vents, échelle graphique en km ;
  - carte de localisation du Togo (régions, pays voisins, région d'étude) ;
  - légende, datum, auteur, cadre.

Sorties : carte_article.png (300 dpi) et carte_article.pdf (vectoriel).

Installation (PowerShell, une seule fois) :
    python -m pip install pandas openpyxl geopandas matplotlib

Lancement (depuis le dossier contenant le fichier Excel) :
    python carte_article.py
"""
import os
import urllib.request

import geopandas as gpd
import matplotlib.patheffects as pe
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.lines import Line2D
from matplotlib.patches import Circle, Patch, Polygon, Rectangle
from matplotlib.ticker import FuncFormatter, MultipleLocator
from shapely.geometry import box

# ---------------- RÉGLAGES ----------------
FICHIER = "carte_CAT_propre.xlsx"
COL_LAT, COL_LON, COL_ALT = "Latitude", "Longitude", "Altitude (m)"

TITRE = "DONNÉES DE L'OUTIL CAT : ZONE D'ÉTUDE"
AUTEUR = "BOKOUMBO Tibetobe Ambroise"                 # nom affiché en bas à droite
EMPRISE = "prefectures"                # "prefectures" (préfectures concernées) ou "points"
TAMPON_ZONE_M = 1500                   # marge de la zone d'étude autour des points
SORTIE = "carte_article"

DOSSIER_FONDS = "fonds_carte"
PAYS_MONDE = "ne_admin0_pays.geojson"      # voisins (facultatif)
PREF_LOCAL = "prefectures_togo.geojson"    # préfectures (sinon téléchargées)
BASE = ("https://media.githubusercontent.com/media/wmgeolab/geoBoundaries/"
        "main/releaseData/gbOpen/TGO/{niv}/geoBoundaries-TGO-{niv}.geojson")

PALETTE = ["#f6e3b0", "#5b9bd5", "#b5d99c", "#f4b6c2", "#cdb4db", "#ffd166", "#9ad1d4", "#e9c46a"]
C_POINT = "#d7191c"
C_ZONE = "#b2182b"
C_MER = "#e9f0f6"
# ------------------------------------------

plt.rcParams.update({
    "font.family": "sans-serif",
    "font.sans-serif": ["Arial", "DejaVu Sans"],
    "font.size": 9,
    "axes.linewidth": 0.8,
})
HALO = [pe.withStroke(linewidth=2.5, foreground="white")]


# ---------- utilitaires ----------
def trouver_col(gdf, candidats):
    bas = {str(c).lower(): c for c in gdf.columns}
    for c in candidats:
        if c.lower() in bas:
            return bas[c.lower()]
    return None


def charger_limites(niveau):
    os.makedirs(DOSSIER_FONDS, exist_ok=True)
    chemin = os.path.join(DOSSIER_FONDS, f"TGO_{niveau}.geojson")
    if not os.path.exists(chemin):
        print(f"Téléchargement des limites {niveau} ...")
        try:
            urllib.request.urlretrieve(BASE.format(niv=niveau), chemin)
        except Exception as e:
            raise SystemExit(f"Téléchargement impossible ({e}). "
                             f"Placez le fichier geoBoundaries {niveau} du Togo sous : {chemin}")
    return gpd.read_file(chemin).to_crs(epsg=4326)


def nettoyer(nom):
    for mot in (" Region", " Région", " Prefecture", " Préfecture", " District"):
        nom = str(nom).replace(mot, "")
    return nom.strip()


def attribuer(pts, unites, colonne="nom", max_m=3000):
    """Unité (polygone) la plus proche de chaque point ; distance 0 si le point est dedans."""
    u = unites[[colonne, "geometry"]].to_crs(epsg=32631)
    j = gpd.sjoin_nearest(pts.to_crs(epsg=32631), u, how="left", max_distance=max_m)
    j = j[~j.index.duplicated()]
    return j[colonne]


def dms(v, pos_hem, neg_hem):
    a = abs(v)
    d = int(a)
    m = int(round((a - d) * 60))
    if m == 60:
        d, m = d + 1, 0
    return f"{d}°{m:02d}′{pos_hem if v >= 0 else neg_hem}"


def pas_graduation(etendue_deg):
    for st in (1, 2, 5, 10, 15, 20, 30, 60):
        if etendue_deg * 60 / st <= 5:
            return st / 60
    return 1.0


def fixer_etendue(ax, xmin, xmax, ymin, ymax):
    ax.set_xlim(xmin, xmax)
    ax.set_ylim(ymin, ymax)
    ax.set_aspect(1 / np.cos(np.radians((ymin + ymax) / 2)))


def graticule(ax):
    xmin, xmax = ax.get_xlim()
    ymin, ymax = ax.get_ylim()
    ax.xaxis.set_major_locator(MultipleLocator(pas_graduation(xmax - xmin)))
    ax.yaxis.set_major_locator(MultipleLocator(pas_graduation(ymax - ymin)))
    ax.xaxis.set_major_formatter(FuncFormatter(lambda x, p: dms(x, "E", "O")))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda y, p: dms(y, "N", "S")))
    ax.tick_params(direction="out", length=3, labelsize=7.5, top=True, right=True)
    ax.grid(True, ls="--", lw=0.3, color="#8c8c8c", alpha=0.7)
    ax.set_axisbelow(False)
    ax.set_xlabel("Longitude", fontsize=8.5)
    ax.set_ylabel("Latitude", fontsize=8.5)
    for lab in ax.get_yticklabels():
        lab.set_rotation(90)
        lab.set_va("center")


def echelle(ax, km_max=None, pos=(0.05, 0.05), segments=2):
    xmin, xmax = ax.get_xlim()
    ymin, ymax = ax.get_ylim()
    lat = (ymin + ymax) / 2
    largeur_km = (xmax - xmin) * 111.32 * np.cos(np.radians(lat))
    if km_max is None:
        km_max = max(k for k in (2, 4, 10, 20, 40, 60, 100, 200) if k <= 0.3 * largeur_km)
    deg = km_max / (111.32 * np.cos(np.radians(lat)))
    x0 = xmin + pos[0] * (xmax - xmin)
    y0 = ymin + pos[1] * (ymax - ymin)
    h = 0.014 * (ymax - ymin)
    ax.add_patch(Rectangle((x0 - 0.015 * (xmax - xmin), y0 - 0.05 * (ymax - ymin)),
                           deg + 0.05 * (xmax - xmin), 0.095 * (ymax - ymin),
                           fc="white", ec="none", alpha=0.85, zorder=9))
    for i in range(segments):
        ax.add_patch(Rectangle((x0 + i * deg / segments, y0), deg / segments, h,
                               fc="black" if i % 2 == 0 else "white", ec="black", lw=0.7, zorder=10))
    for i in range(segments + 1):
        ax.text(x0 + i * deg / segments, y0 - 0.012 * (ymax - ymin), f"{km_max * i / segments:g}",
                ha="center", va="top", fontsize=7.5, zorder=10)
    ax.text(x0 + deg + 0.012 * (xmax - xmin), y0 + h / 2, "km", ha="left", va="center",
            fontsize=7.5, zorder=10)


def rose_des_vents(ax_parent, bounds=(0.02, 0.74, 0.17, 0.24)):
    """Rose des vents dessinée dans un petit axe incrusté (proportions respectées)."""
    ac = ax_parent.inset_axes(list(bounds))
    ac.set_xlim(-1.5, 1.5)
    ac.set_ylim(-1.5, 1.5)
    ac.set_aspect("equal")
    ac.axis("off")
    ac.set_zorder(12)
    ac.add_patch(Circle((0, 0), 0.62, fc="white", ec="black", lw=0.8, alpha=0.8))
    for ang in (0, 90, 180, 270):
        t = np.radians(ang)
        R = np.array([[np.cos(t), -np.sin(t)], [np.sin(t), np.cos(t)]])
        for pts_, fc in (([(0, 1.0), (0, 0), (-0.2, 0.2)], "black"),
                         ([(0, 1.0), (0, 0), (0.2, 0.2)], "white")):
            ac.add_patch(Polygon([tuple(R @ np.array(p)) for p in pts_], fc=fc, ec="black", lw=0.7))
    for txt, (x, y), fs in (("N", (0, 1.28), 10), ("E", (1.27, 0), 8), ("S", (0, -1.28), 8), ("O", (-1.27, 0), 8)):
        ac.text(x, y, txt, ha="center", va="center", fontsize=fs, fontweight="bold")


# ---------- 1) données ----------
df = pd.read_excel(FICHIER).dropna(subset=[COL_LAT, COL_LON]).reset_index(drop=True)
pts = gpd.GeoDataFrame(df, geometry=gpd.points_from_xy(df[COL_LON], df[COL_LAT]), crs="EPSG:4326")

pays = charger_limites("ADM0")
regions = charger_limites("ADM1")
col_reg = trouver_col(regions, ["shapeName", "NAME_1", "region", "name"]) or regions.columns[0]
regions["nom"] = regions[col_reg].map(nettoyer)

if os.path.exists(PREF_LOCAL):
    prefs = gpd.read_file(PREF_LOCAL).to_crs(epsg=4326)
    col_pref = trouver_col(prefs, ["prefecture", "préfecture", "pref", "name_2", "shapeName",
                                   "nom", "name"]) or prefs.columns[0]
else:
    prefs = charger_limites("ADM2")
    col_pref = trouver_col(prefs, ["shapeName", "NAME_2", "name"]) or prefs.columns[0]
prefs["nom"] = prefs[col_pref].map(nettoyer)

voisins, col_vois = None, None
if os.path.exists(PAYS_MONDE):
    monde = gpd.read_file(PAYS_MONDE).to_crs(epsg=4326)
    col_vois = trouver_col(monde, ["NAME_FR", "NAME", "ADMIN", "name"])
    txt = monde.drop(columns="geometry").astype(str).apply(lambda c: c.str.lower())
    voisins = monde[~txt.isin(["togo", "togolese republic", "tgo"]).any(axis=1)]

pts["pref"] = attribuer(pts, prefs).fillna("Hors limites")
pts["reg"] = attribuer(pts, regions).fillna("—")
regions_etude = sorted(pts.loc[pts["reg"] != "—", "reg"].unique().tolist())
effectifs = pts.loc[pts["pref"] != "Hors limites", "pref"].value_counts()
prefs_etude = effectifs.index.tolist()
couleur_pref = {p: PALETTE[i % len(PALETTE)] for i, p in enumerate(prefs_etude)}
print("Région(s) :", regions_etude)
print("Préfecture(s) :", prefs_etude)

# zone d'étude (enveloppe convexe des points + marge)
zone_m = pts.to_crs(epsg=32631).union_all().convex_hull.buffer(TAMPON_ZONE_M)
zone = gpd.GeoSeries([zone_m], crs="EPSG:32631").to_crs(epsg=4326)

# emprise de la carte principale
if EMPRISE == "prefectures" and prefs_etude:
    x0, y0, x1, y1 = prefs[prefs["nom"].isin(prefs_etude)].total_bounds
else:
    x0, y0, x1, y1 = zone.total_bounds
mx, my = 0.05 * (x1 - x0), 0.05 * (y1 - y0)
emprise = (x0 - mx, x1 + mx, y0 - my, y1 + my)

# ---------- 2) figure ----------
fig = plt.figure(figsize=(11.5, 8.8), dpi=150)
fig.add_artist(Rectangle((0.008, 0.008), 0.984, 0.984, fill=False, ec="black", lw=1.3,
                         transform=fig.transFigure))
fig.text(0.5, 0.952, TITRE, ha="center", va="center", fontsize=16, fontweight="bold")

# ---- carte principale ----
axm = fig.add_axes([0.075, 0.15, 0.60, 0.74])
axm.set_facecolor(C_MER)
fixer_etendue(axm, *emprise)
pays.plot(ax=axm, fc="white", ec="none", zorder=1)
prefs[~prefs["nom"].isin(prefs_etude)].plot(ax=axm, fc="#f2f2f2", ec="#a0a0a0", lw=0.5, zorder=2)
for nom in prefs_etude:
    prefs[prefs["nom"] == nom].plot(ax=axm, fc=couleur_pref[nom], ec="#444444", lw=0.9, zorder=3)
pays.boundary.plot(ax=axm, color="black", lw=1.0, zorder=4)
zone.boundary.plot(ax=axm, color=C_ZONE, lw=1.3, ls="--", zorder=6)
axm.scatter(pts[COL_LON], pts[COL_LAT], marker="s", s=17, c=C_POINT,
            edgecolor="black", linewidth=0.35, zorder=8)

cadre = box(*[emprise[0], emprise[2], emprise[1], emprise[3]])
for _, r in prefs.iterrows():
    g = r.geometry.intersection(cadre)
    if g.is_empty or (r["nom"] not in prefs_etude and g.area < 0.02 * cadre.area):
        continue
    p = g.representative_point()
    en_etude = r["nom"] in prefs_etude
    axm.text(p.x, p.y, r["nom"], fontsize=9 if en_etude else 7, ha="center", va="center",
             fontweight="bold" if en_etude else "normal",
             color="black" if en_etude else "#6b6b6b", path_effects=HALO, zorder=7)

graticule(axm)
echelle(axm, pos=(0.05, 0.05), segments=2)
rose_des_vents(axm)

# ---- localisation (haut droite) ----
axi = fig.add_axes([0.705, 0.62, 0.265, 0.27])
bx0, by0, bx1, by1 = pays.total_bounds
ei = (bx0 - 0.9, bx1 + 0.9, by0 - 0.30, by1 + 0.25)
axi.set_facecolor(C_MER)
fixer_etendue(axi, *ei)
if voisins is not None:
    voisins.plot(ax=axi, fc="#f2f2f2", ec="#9a9a9a", lw=0.5, zorder=1)
regions[~regions["nom"].isin(regions_etude)].plot(ax=axi, fc="#d9d9d9", ec="#555555", lw=0.6, zorder=2)
regions[regions["nom"].isin(regions_etude)].plot(ax=axi, fc="#7bc47f", ec="#2d6a2d", lw=0.9, zorder=3)
pays.boundary.plot(ax=axi, color="black", lw=0.9, zorder=4)
cadre_i = box(ei[0], ei[2], ei[1], ei[3])
if voisins is not None and col_vois:
    for _, r in voisins.iterrows():
        g = r.geometry.intersection(cadre_i)
        if not g.is_empty:
            p = g.representative_point()
            axi.text(p.x, p.y, str(r[col_vois]), fontsize=6, style="italic", color="#555555",
                     ha="center", va="center", zorder=5)
for _, r in regions.iterrows():
    p = r.geometry.representative_point()
    y = p.y
    if r["nom"] in regions_etude:
        y = max(p.y, emprise[3] + 0.25)          # étiquette au-dessus du cadre rouge
    axi.text(p.x, y, r["nom"], fontsize=6, ha="center", va="center", zorder=7, path_effects=HALO,
             fontweight="bold" if r["nom"] in regions_etude else "normal")
axi.add_patch(Rectangle((emprise[0], emprise[2]), emprise[1] - emprise[0], emprise[3] - emprise[2],
                        fill=False, ec=C_ZONE, lw=1.4, zorder=6))
axi.set_xticks([])
axi.set_yticks([])
axi.set_xlabel("")
axi.set_ylabel("")
axi.set_title("Localisation", fontsize=9, fontweight="bold", pad=3)

# ---- légende ----
axl = fig.add_axes([0.705, 0.375, 0.265, 0.235])
axl.axis("off")
handles = [Line2D([], [], marker="s", ls="", mfc=C_POINT, mec="black", ms=6, label=f"Points GPS (n = {len(pts)})"),
           Line2D([], [], color=C_ZONE, lw=1.3, ls="--", label="Zone d'étude")]
handles += [Patch(fc=couleur_pref[p], ec="#444444", label=f"Préfecture : {p}") for p in prefs_etude]
handles += [Patch(fc="#f2f2f2", ec="#a0a0a0", label="Autres préfectures"),
            Patch(fc="#7bc47f", ec="#2d6a2d", label=f"Région d'étude : {', '.join(regions_etude)}")]
axl.legend(handles=handles, loc="upper left", title="Légende", fontsize=7.5, title_fontsize=9,
           frameon=True, edgecolor="#555555", framealpha=1, alignment="left")

# ---- pied de page ----
fig.text(0.705, 0.20, "Datum : WGS 84 (EPSG:4326)\nProjection de référence : UTM zone 31N\n"
         f"Auteur : {AUTEUR}", fontsize=7.5, ha="left", va="top", fontweight="bold")
fig.text(0.075, 0.03, "Sources : relevés GPS de terrain ; limites administratives : geoBoundaries, Natural Earth.",
         fontsize=7.5, ha="left", va="center", color="#333333")

# ---------- 3) export ----------
fig.savefig(SORTIE + ".png", dpi=300, facecolor="white")
fig.savefig(SORTIE + ".pdf", facecolor="white")
print(f"Figures enregistrées : {SORTIE}.png (300 dpi) et {SORTIE}.pdf (vectoriel)")
plt.show()
