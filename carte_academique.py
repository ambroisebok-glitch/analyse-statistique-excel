# -*- coding: utf-8 -*-
"""
Carte académique : localisation de la zone d'étude (Togo).
Figure statique de qualité publication : deux panneaux (a) contexte national,
(b) zone d'étude, avec flèche du nord, échelle graphique, graticule en
coordonnées géographiques, légende, source et système de projection.

Sorties : carte_academique.png (300 dpi) et carte_academique.pdf (vectoriel).

Installation (PowerShell, une seule fois) :
    python -m pip install pandas openpyxl geopandas matplotlib

Lancement (depuis le dossier contenant le fichier Excel) :
    python carte_academique.py
"""
import os
import urllib.request

import geopandas as gpd
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.lines import Line2D
from matplotlib.patches import ConnectionPatch, Patch, Rectangle
from matplotlib.ticker import FuncFormatter, MaxNLocator
from shapely.geometry import box

# ---------------- RÉGLAGES ----------------
FICHIER = "carte_CAT_propre.xlsx"
COL_LAT, COL_LON, COL_ALT = "Latitude", "Longitude", "Altitude (m)"

TITRE = "Localisation de la zone d'étude (Région Maritime, Togo)"
SOURCE = ("Sources : relevés GPS de terrain (n = {n}) ; limites administratives : "
          "geoBoundaries, Natural Earth.\n"
          "Projection : WGS 84 (EPSG:4326). Échelle graphique calculée à la latitude moyenne de chaque panneau.")
TAMPON_ZONE_M = 1500           # marge autour des points pour la zone d'étude
SORTIE = "carte_academique"

# Fonds de carte (le premier trouvé est utilisé ; sinon téléchargement)
DOSSIER_FONDS = "fonds_carte"
PAYS_MONDE = "ne_admin0_pays.geojson"      # voisins (facultatif)
PREF_LOCAL = "prefectures_togo.geojson"    # préfectures (facultatif)
BASE = ("https://media.githubusercontent.com/media/wmgeolab/geoBoundaries/"
        "main/releaseData/gbOpen/TGO/{niv}/geoBoundaries-TGO-{niv}.geojson")

# Palette sobre, lisible aussi en noir et blanc
C_MER, C_VOISIN, C_REGION, C_ETUDE = "#e8f1f8", "#f2f2f2", "#d9d9d9", "#fdd49e"
C_ZONE = "#b2182b"
# ------------------------------------------

plt.rcParams.update({
    "font.family": "serif",
    "font.serif": ["Times New Roman", "DejaVu Serif"],
    "font.size": 9,
    "axes.linewidth": 0.8,
})


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
    return str(nom).replace(" Region", "").replace(" Région", "").strip()


def habiller_axes(ax, pas_x=5, pas_y=5, dec=2):
    ax.xaxis.set_major_locator(MaxNLocator(pas_x, steps=[1, 2, 2.5, 5, 10]))
    ax.yaxis.set_major_locator(MaxNLocator(pas_y, steps=[1, 2, 2.5, 5, 10]))
    ax.xaxis.set_major_formatter(FuncFormatter(lambda x, p: f"{abs(x):.{dec}f}°{'E' if x >= 0 else 'O'}"))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda y, p: f"{abs(y):.{dec}f}°{'N' if y >= 0 else 'S'}"))
    ax.set_xlabel("Longitude", fontsize=8.5)
    ax.set_ylabel("Latitude", fontsize=8.5)
    ax.set_anchor("N")
    ax.tick_params(direction="out", length=3, labelsize=7.5)
    ax.grid(True, ls="--", lw=0.35, color="#7f7f7f", alpha=0.7)
    ax.set_axisbelow(False)
    ax.set_facecolor(C_MER)
    for lab in ax.get_yticklabels():
        lab.set_rotation(90)
        lab.set_va("center")


def fixer_etendue(ax, xmin, xmax, ymin, ymax):
    ax.set_xlim(xmin, xmax)
    ax.set_ylim(ymin, ymax)
    lat = (ymin + ymax) / 2
    ax.set_aspect(1 / np.cos(np.radians(lat)))   # proportions correctes à cette latitude


def echelle(ax, km, pos=(0.05, 0.05), segments=2):
    """Échelle graphique (alternance noir / blanc) en km."""
    xmin, xmax = ax.get_xlim()
    ymin, ymax = ax.get_ylim()
    lat = (ymin + ymax) / 2
    deg = km / (111.32 * np.cos(np.radians(lat)))
    x0 = xmin + pos[0] * (xmax - xmin)
    y0 = ymin + pos[1] * (ymax - ymin)
    h = 0.014 * (ymax - ymin)
    ax.add_patch(Rectangle((x0 - 0.015 * (xmax - xmin), y0 - 0.045 * (ymax - ymin)),
                           deg + 0.035 * (xmax - xmin), 0.085 * (ymax - ymin),
                           fc="white", ec="none", alpha=0.85, zorder=9))
    for i in range(segments):
        ax.add_patch(Rectangle((x0 + i * deg / segments, y0), deg / segments, h,
                               fc="black" if i % 2 == 0 else "white",
                               ec="black", lw=0.7, zorder=10))
    for i in range(segments + 1):
        val = km * i / segments
        ax.text(x0 + i * deg / segments, y0 - 0.012 * (ymax - ymin),
                f"{val:g}", ha="center", va="top", fontsize=7, zorder=10)
    ax.text(x0 + deg + 0.012 * (xmax - xmin), y0 + h / 2, "km",
            ha="left", va="center", fontsize=7, zorder=10)


def fleche_nord(ax, pos=(0.92, 0.80), longueur=0.11):
    ax.annotate("", xy=(pos[0], pos[1] + longueur), xytext=pos, xycoords="axes fraction",
                arrowprops=dict(arrowstyle="-|>", lw=1.4, color="black", mutation_scale=14),
                zorder=11)
    ax.text(pos[0], pos[1] + longueur + 0.015, "N", transform=ax.transAxes,
            ha="center", va="bottom", fontsize=11, fontweight="bold", zorder=11)


# ---------- 1) données ----------
df = pd.read_excel(FICHIER).dropna(subset=[COL_LAT, COL_LON])
pts = gpd.GeoDataFrame(df, geometry=gpd.points_from_xy(df[COL_LON], df[COL_LAT]), crs="EPSG:4326")

# pays et régions
pays = charger_limites("ADM0")
regions = charger_limites("ADM1")
col_reg = trouver_col(regions, ["shapeName", "NAME_1", "region", "name"]) or regions.columns[0]
regions["nom"] = regions[col_reg].map(nettoyer)

# voisins (facultatif)
voisins = None
col_vois = None
if os.path.exists(PAYS_MONDE):
    monde = gpd.read_file(PAYS_MONDE).to_crs(epsg=4326)
    col_vois = trouver_col(monde, ["NAME_FR", "NAME", "ADMIN", "name"])
    txt = monde.drop(columns="geometry").astype(str).apply(lambda c: c.str.lower())
    est_togo = txt.isin(["togo", "togolese republic", "tgo"]).any(axis=1)
    voisins = monde[~est_togo]

# préfectures (facultatif)
prefs = None
if os.path.exists(PREF_LOCAL):
    prefs = gpd.read_file(PREF_LOCAL).to_crs(epsg=4326)
    col_pref = trouver_col(prefs, ["prefecture", "préfecture", "pref", "name_2", "shapeName",
                                   "nom", "name"]) or prefs.columns[0]

# région(s) de l'étude et zone d'étude
jr = gpd.sjoin(pts, regions[["nom", "geometry"]], how="left", predicate="within")
regions_etude = sorted(jr["nom"].dropna().unique().tolist())
regions["etude"] = regions["nom"].isin(regions_etude)
print("Région(s) de l'étude :", regions_etude)

zone_m = pts.to_crs(epsg=32631).union_all().convex_hull.buffer(TAMPON_ZONE_M)
zone = gpd.GeoSeries([zone_m], crs="EPSG:32631").to_crs(epsg=4326)
zx0, zy0, zx1, zy1 = zone.total_bounds
surface_km2 = zone_m.area / 1e6

# ---------- 2) figure ----------
fig = plt.figure(figsize=(11.5, 6.4), dpi=150)
gs = fig.add_gridspec(1, 2, width_ratios=[1, 1.35], left=0.06, right=0.97,
                      top=0.88, bottom=0.24, wspace=0.16)
axa = fig.add_subplot(gs[0])
axb = fig.add_subplot(gs[1])

# ---- (a) contexte national ----
bx0, by0, bx1, by1 = pays.total_bounds
m = 0.35
ea = (bx0 - m, bx1 + m, by0 - 0.60, by1 + 0.25)
fixer_etendue(axa, *ea)
if voisins is not None:
    voisins.plot(ax=axa, fc=C_VOISIN, ec="#9a9a9a", lw=0.6, zorder=1)
regions[~regions["etude"]].plot(ax=axa, fc=C_REGION, ec="#555555", lw=0.7, zorder=2)
regions[regions["etude"]].plot(ax=axa, fc=C_ETUDE, ec="#555555", lw=0.9, zorder=3)
pays.boundary.plot(ax=axa, color="black", lw=1.2, zorder=4)

cadre = box(ea[0], ea[2], ea[1], ea[3])
if voisins is not None and col_vois:
    for _, r in voisins.iterrows():
        g = r.geometry.intersection(cadre)
        if not g.is_empty:
            p = g.representative_point()
            axa.text(p.x, p.y, str(r[col_vois]), fontsize=7.5, style="italic",
                     color="#555555", ha="center", va="center", zorder=5)
for _, r in regions.iterrows():
    p = r.geometry.representative_point()
    axa.text(p.x, p.y, r["nom"], fontsize=7.5, ha="center", va="center", zorder=5,
             fontweight="bold" if r["etude"] else "normal")
axa.text(ea[0] + 0.72 * (ea[1] - ea[0]), by0 - 0.22, "Golfe de Guinée",
         fontsize=8, style="italic", color="#3b6a8f", ha="center", zorder=5)

axa.add_patch(Rectangle((zx0, zy0), zx1 - zx0, zy1 - zy0, fill=False, ec=C_ZONE, lw=1.6, zorder=6))
habiller_axes(axa, 4, 6, dec=1)
echelle(axa, 100, pos=(0.06, 0.045), segments=2)
fleche_nord(axa, pos=(0.90, 0.80))
axa.set_title("(a) Contexte national", loc="left", fontsize=10, fontweight="bold")
fig.legend(handles=[
    Patch(fc=C_ETUDE, ec="#555555", label=f"Région d'étude ({', '.join(regions_etude)})"),
    Patch(fc=C_REGION, ec="#555555", label="Autres régions"),
    Patch(fc="none", ec=C_ZONE, lw=1.6, label="Emprise du panneau (b)"),
], loc="lower center", bbox_to_anchor=(0.5, 0.085), ncol=3, fontsize=8, frameon=True,
    framealpha=1, edgecolor="#555555")

# ---- (b) zone d'étude ----
mx, my = 0.10 * (zx1 - zx0), 0.12 * (zy1 - zy0)
eb = (zx0 - mx, zx1 + mx, zy0 - my, zy1 + my)
fixer_etendue(axb, *eb)
regions.plot(ax=axb, fc=C_REGION, ec="#555555", lw=0.6, zorder=1)
regions[regions["etude"]].plot(ax=axb, fc=C_ETUDE, ec="#555555", lw=0.8, zorder=2)
if prefs is not None:
    prefs.boundary.plot(ax=axb, color="#6b6b6b", lw=0.7, ls="--", zorder=3)
    cadre_b = box(eb[0], eb[2], eb[1], eb[3])
    for _, r in prefs.iterrows():
        g = r.geometry.intersection(cadre_b)
        if not g.is_empty and g.area > 0.04 * cadre_b.area:
            p = g.representative_point()
            axb.text(p.x, p.y, str(r[col_pref]), fontsize=7, style="italic",
                     color="#444444", ha="center", va="center", zorder=4)
pays.boundary.plot(ax=axb, color="black", lw=1.0, zorder=4)
zone.boundary.plot(ax=axb, color=C_ZONE, lw=1.6, ls="--", zorder=6)

sc = axb.scatter(df[COL_LON], df[COL_LAT], c=df[COL_ALT], cmap="viridis", s=24,
                 edgecolor="black", linewidth=0.4, zorder=8)
habiller_axes(axb, 5, 5, dec=2)
echelle(axb, 5 if (zx1 - zx0) * 105 < 60 else 10, pos=(0.05, 0.06), segments=2)
fleche_nord(axb, pos=(0.93, 0.80))
axb.set_title("(b) Zone d'étude et points GPS", loc="left", fontsize=10, fontweight="bold")

cax = axb.inset_axes([0.0, -0.20, 1.0, 0.035])
cb = fig.colorbar(sc, cax=cax, orientation="horizontal")
cb.set_label("Altitude (m)", fontsize=8)
cb.ax.tick_params(labelsize=7.5, length=2)

axb.legend(handles=[
    Line2D([], [], marker="o", ls="", mfc="#21918c", mec="black", ms=5, label=f"Point GPS (n = {len(df)})"),
    Line2D([], [], color=C_ZONE, lw=1.6, ls="--", label=f"Zone d'étude (≈ {surface_km2:,.0f} km²)".replace(",", " ")),
] + ([Line2D([], [], color="#6b6b6b", lw=0.8, ls="--", label="Limite de préfecture")] if prefs is not None else []),
    loc="best", fontsize=7, frameon=True, framealpha=0.95, edgecolor="#555555").set_zorder(12)

# lignes de rappel entre (a) et (b)
for (xa, ya), (xb_, yb_) in (((zx1, zy1), (0, 1)), ((zx1, zy0), (0, 0))):
    fig.add_artist(ConnectionPatch(xyA=(xa, ya), coordsA=axa.transData,
                                   xyB=(xb_, yb_), coordsB=axb.transAxes,
                                   color=C_ZONE, lw=0.8, ls=":", zorder=1))

# titre et source
fig.suptitle(TITRE, fontsize=13, fontweight="bold", y=0.965)
fig.text(0.06, 0.03, SOURCE.format(n=len(df)), fontsize=7.5, ha="left", va="bottom", color="#333333")

# ---------- 3) export ----------
fig.savefig(SORTIE + ".png", dpi=300, bbox_inches="tight", facecolor="white")
fig.savefig(SORTIE + ".pdf", bbox_inches="tight", facecolor="white")
print(f"Figures enregistrées : {SORTIE}.png (300 dpi) et {SORTIE}.pdf (vectoriel)")
plt.show()
