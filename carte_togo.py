# -*- coding: utf-8 -*-
"""
Carte interactive : Togo, région(s) concernée(s), zone d'étude et points GPS.
Résultat : carte_togo.html (à ouvrir dans le navigateur).

Installation (PowerShell, une seule fois) :
    python -m pip install pandas openpyxl folium geopandas

Lancement (depuis le dossier qui contient le fichier Excel) :
    python carte_togo.py
"""
import json
import os
import urllib.request

import folium
import geopandas as gpd
import pandas as pd
from folium import plugins
from branca.colormap import LinearColormap

# ---------------- RÉGLAGES ----------------
FICHIER = "carte_CAT_propre.xlsx"
COL_LAT, COL_LON = "Latitude", "Longitude"
COL_ID, COL_ALT, COL_PREC = "ID", "Altitude (m)", "Précision GPS (m)"
SORTIE = "carte_togo.html"
TAMPON_ZONE_M = 1500        # marge autour des points pour dessiner la zone d'étude
DOSSIER_FONDS = "fonds_carte"

# Limites administratives (geoBoundaries, licence ouverte). Le script les
# télécharge une fois, puis les garde dans le dossier "fonds_carte".
BASE = ("https://media.githubusercontent.com/media/wmgeolab/geoBoundaries/"
        "main/releaseData/gbOpen/TGO/{niv}/geoBoundaries-TGO-{niv}.geojson")
# ------------------------------------------


def charger_limites(niveau):
    """ADM0 = pays, ADM1 = régions. Téléchargement avec mise en cache."""
    os.makedirs(DOSSIER_FONDS, exist_ok=True)
    chemin = os.path.join(DOSSIER_FONDS, f"TGO_{niveau}.geojson")
    if not os.path.exists(chemin):
        print(f"Téléchargement des limites {niveau} ...")
        try:
            urllib.request.urlretrieve(BASE.format(niv=niveau), chemin)
        except Exception as e:
            raise SystemExit(
                f"Téléchargement impossible ({e}).\n"
                f"Téléchargez à la main le fichier geoBoundaries TGO {niveau} "
                f"(GeoJSON) et enregistrez-le sous : {chemin}"
            )
    return gpd.read_file(chemin).to_crs(epsg=4326)


def nom_colonne(gdf):
    for c in ("shapeName", "NAME_1", "name", "region"):
        if c in gdf.columns:
            return c
    return gdf.columns[0]


# 1) Données GPS ------------------------------------------------------------
df = pd.read_excel(FICHIER).dropna(subset=[COL_LAT, COL_LON])
pts = gpd.GeoDataFrame(
    df, geometry=gpd.points_from_xy(df[COL_LON], df[COL_LAT]), crs="EPSG:4326"
)

# 2) Fonds de carte administratifs -----------------------------------------
pays = charger_limites("ADM0")
regions = charger_limites("ADM1")
col_reg = nom_colonne(regions)

# Région(s) qui contiennent les points
jointure = gpd.sjoin(pts, regions[[col_reg, "geometry"]], how="left", predicate="within")
regions_etude = sorted(jointure[col_reg].dropna().unique().tolist())
print("Région(s) de la zone d'étude :", regions_etude)
regions["en_etude"] = regions[col_reg].isin(regions_etude)

# 3) Zone d'étude = enveloppe convexe des points + marge --------------------
zone = pts.to_crs(epsg=32631).union_all().convex_hull.buffer(TAMPON_ZONE_M)
zone = gpd.GeoDataFrame(geometry=[zone], crs="EPSG:32631").to_crs(epsg=4326)
surface_km2 = gpd.GeoSeries(zone.to_crs(epsg=32631).geometry).area.iloc[0] / 1e6

# 4) Carte -------------------------------------------------------------------
minx, miny, maxx, maxy = pays.total_bounds
m = folium.Map(location=[(miny + maxy) / 2, (minx + maxx) / 2],
               zoom_start=7, tiles=None, control_scale=True)
folium.TileLayer("OpenStreetMap", name="Fond : OpenStreetMap").add_to(m)
folium.TileLayer(
    "https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}.png",
    attr="&copy; OpenStreetMap, &copy; CARTO", name="Fond : clair", show=False,
).add_to(m)
folium.TileLayer(
    "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
    attr="Esri World Imagery", name="Fond : satellite", show=False,
).add_to(m)

# Palette des régions (la région d'étude ressort)
couleurs = {"Maritime": "#e67e22", "Plateaux": "#27ae60", "Centrale": "#8e44ad",
            "Kara": "#2980b9", "Savanes": "#f1c40f"}


def style_region(f):
    nom = f["properties"][col_reg]
    etude = f["properties"]["en_etude"]
    base = next((c for k, c in couleurs.items() if k.lower() in str(nom).lower()), "#95a5a6")
    return {"fillColor": base, "color": "#34495e",
            "weight": 2.5 if etude else 1, "fillOpacity": 0.45 if etude else 0.22}


folium.GeoJson(
    pays, name="Togo (frontière)",
    style_function=lambda f: {"fillOpacity": 0, "color": "#111", "weight": 3},
).add_to(m)

folium.GeoJson(
    json.loads(regions.to_json()), name="Régions",
    style_function=style_region,
    highlight_function=lambda f: {"weight": 3, "fillOpacity": 0.6},
    tooltip=folium.GeoJsonTooltip(fields=[col_reg], aliases=["Région :"]),
).add_to(m)

folium.GeoJson(
    json.loads(zone.to_json()), name="Zone d'étude",
    style_function=lambda f: {"fillColor": "#e74c3c", "color": "#c0392b",
                              "weight": 3, "dashArray": "8 5", "fillOpacity": 0.15},
    tooltip=f"Zone d'étude (~{surface_km2:,.0f} km²)".replace(",", " "),
).add_to(m)

# Points colorés selon l'altitude
cmap = LinearColormap(["#2c7bb6", "#abd9e9", "#ffffbf", "#fdae61", "#d7191c"],
                      vmin=df[COL_ALT].min(), vmax=df[COL_ALT].max(),
                      caption="Altitude (m)")
groupe = folium.FeatureGroup(name=f"Points GPS ({len(df)})")
for _, r in df.iterrows():
    popup = (f"<b>ID :</b> {r[COL_ID]}<br><b>Altitude :</b> {r[COL_ALT]} m<br>"
             f"<b>Précision GPS :</b> {r[COL_PREC]} m")
    folium.CircleMarker(
        [r[COL_LAT], r[COL_LON]], radius=5, color="#222", weight=0.8,
        fill=True, fill_color=cmap(r[COL_ALT]), fill_opacity=0.9,
        popup=folium.Popup(popup, max_width=220), tooltip=str(r[COL_ID]),
    ).add_to(groupe)
groupe.add_to(m)
cmap.add_to(m)

# Habillage
titre = f"""
<div style="position:fixed;top:10px;left:50px;z-index:9999;background:white;
 padding:8px 14px;border-radius:6px;box-shadow:0 1px 6px rgba(0,0,0,.35);
 font-family:Arial;font-size:14px;">
 <b>Togo – zone d'étude</b><br>
 Région : {', '.join(regions_etude) or '—'}<br>
 {len(df)} points GPS · ~{surface_km2:,.0f} km²
</div>""".replace(",", " ")
m.get_root().html.add_child(folium.Element(titre))

plugins.MiniMap(toggle_display=True, position="bottomleft").add_to(m)
plugins.Fullscreen().add_to(m)
plugins.MeasureControl(position="topleft", primary_length_unit="kilometers").add_to(m)
folium.LayerControl(collapsed=False).add_to(m)

m.fit_bounds([[miny, minx], [maxy, maxx]])
m.save(SORTIE)
print(f"Carte enregistrée : {SORTIE}")
