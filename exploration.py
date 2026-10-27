# -*- coding: utf-8 -*-
"""
Exploration automatique de StatExcel : détection des variables socio-démographiques,
suggestion des hypothèses testables et traitement des coordonnées géographiques.

Ce module ne contient aucun appel Streamlit (voir ui_exploration.py pour l'interface).

Auteur : BOKOUMBO Tibetobe Ambroise
"""
import re
import unicodedata

import altair as alt
import numpy as np
import pandas as pd
from scipy import stats
from statsmodels.stats.multitest import multipletests

import stats_utils as su

BLEU = "#1f77b4"
PALETTE = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
           "#8c564b", "#e377c2", "#17becf", "#bcbd22", "#393b79"]
GRIS = "#9aa0a6"


def _norm(nom):
    """'Âge (ans)' -> 'age ans' : minuscules, sans accents ni ponctuation."""
    t = unicodedata.normalize("NFKD", str(nom)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", " ", t).strip()


# ======================================================================
# 1. Variables socio-démographiques
# ======================================================================
CATEGORIES = [
    ("Sexe / genre", ["sexe", "genre", "gender", "sex"]),
    ("Âge", ["age", "ages", "ans"]),
    ("Situation matrimoniale", ["matrimonial", "matrimoniale", "marital", "conjugal", "etat civil",
                                "situation familiale"]),
    ("Niveau d'instruction", ["instruction", "etude", "etudes", "education", "scolaire", "scolarite",
                              "diplome", "alphabetisation"]),
    ("Profession / activité", ["profession", "occupation", "metier", "emploi", "activite", "travail",
                               "fonction"]),
    ("Revenu / ressources", ["revenu", "revenus", "income", "salaire", "depense", "depenses",
                             "ressource", "ressources"]),
    ("Lieu de résidence", ["region", "prefecture", "commune", "ville", "localite", "village", "quartier",
                           "canton", "district", "departement", "province", "milieu", "residence",
                           "urbain", "rural", "zone", "arrondissement", "pays"]),
    ("Religion", ["religion", "confession"]),
    ("Ethnie / nationalité / langue", ["ethnie", "nationalite", "origine", "langue"]),
    ("Ménage", ["menage", "household", "enfants", "famille"]),
]
_VALEURS_SEXE = {"m", "f", "h", "masculin", "feminin", "homme", "femme", "male", "female",
                 "garcon", "fille", "man", "woman"}


def _correspond(nom_norm, mots):
    jetons = nom_norm.split()
    for m in mots:
        if " " in m:
            if m in nom_norm:
                return True
        elif m in jetons or (len(m) >= 6 and m in nom_norm):
            return True
    return False


def _est_identifiant(s, nom):
    if set(_norm(nom).split()) & {"id", "identifiant", "uuid", "index", "instance", "submission"}:
        return True
    v = s.dropna()
    if len(v) < 20 or pd.api.types.is_datetime64_any_dtype(s):
        return False
    if not pd.api.types.is_numeric_dtype(s):
        return v.nunique() == len(v)
    if (v % 1 == 0).all():
        return bool(v.is_monotonic_increasing and v.nunique() == len(v) and (v.diff().dropna() == 1).all())
    return False


def _ressemble_sexe(s):
    if pd.api.types.is_numeric_dtype(s):
        return False
    vals = {_norm(v) for v in s.dropna().unique()}
    return 2 <= len(vals) <= 3 and vals <= _VALEURS_SEXE


def detecter_sociodemographiques(df, exclure=()):
    """Retourne {catégorie: [colonnes]} dans l'ordre de CATEGORIES."""
    trouve = {}
    for c in df.columns:
        if c in exclure or _est_identifiant(df[c], c) or df[c].nunique(dropna=True) < 2:
            continue
        n = _norm(c)
        for cat, mots in CATEGORIES:
            if _correspond(n, mots):
                trouve.setdefault(cat, []).append(c)
                break
        else:
            if "Sexe / genre" not in trouve and _ressemble_sexe(df[c]):
                trouve.setdefault("Sexe / genre", []).append(c)
    return {cat: trouve[cat] for cat, _ in CATEGORIES if cat in trouve}


def est_quantitative(s):
    return pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s) and s.nunique(dropna=True) > 12


def resume_variable(s):
    """Phrase courte de lecture d'une variable."""
    if est_quantitative(s):
        v = pd.to_numeric(s, errors="coerce").dropna()
        return (f"N = {len(v)} · moyenne {v.mean():.4g} · médiane {v.median():.4g} · "
                f"min–max {v.min():.4g}–{v.max():.4g}")
    f = su.frequences(s).drop(index="Manquant", errors="ignore")
    if f.empty:
        return "Aucune valeur"
    top = f["Effectif"].idxmax()
    return f"{len(f)} modalités · la plus fréquente : {top} ({f.loc[top, '% valides']:.1f} %)".replace(".", ",")


# ---------------------------------------------------------------- graphiques du profil
def graphique_variable(df, col, hauteur=270):
    """Histogramme (variable quantitative) ou barres/secteurs (variable qualitative)."""
    import graphiques as gr
    s = df[col]
    if est_quantitative(s):
        v = pd.to_numeric(s, errors="coerce").dropna()
        classes = int(np.clip(np.ceil(np.log2(max(len(v), 2)) + 1), 8, 30))
        return gr.histogramme(s, col, classes, "Aucune").properties(height=hauteur)
    freq = su.frequences(s)
    nb = len(freq.drop(index="Manquant", errors="ignore"))
    return gr.frequences_graphique(freq, col, "Secteurs" if nb <= 5 else "Barres").properties(height=hauteur)


def pyramide_ages(df, col_age, col_sexe):
    """Pyramide des âges : effectifs par classe d'âge, une modalité de sexe de chaque côté."""
    mods = su.modalites(df[col_sexe])
    if len(mods) != 2:
        raise ValueError("la pyramide exige exactement deux modalités de sexe")
    sexe = df[col_sexe].map(su.texte_modalite)
    age = df[col_age]
    if pd.api.types.is_numeric_dtype(age):
        a = pd.to_numeric(age, errors="coerce")
        if a.notna().sum() < 5:
            raise ValueError("effectif insuffisant")
        pas = 5 if a.max() - a.min() <= 90 else 10
        debut = np.floor(a.min() / pas) * pas
        bornes = np.arange(debut, a.max() + pas + 1e-9, pas)
        ordre = [f"{int(b)}–{int(b + pas - 1)}" for b in bornes[:-1]]
        classe = pd.cut(a, bins=bornes, labels=ordre, right=False).astype(object)
    else:
        classe = age.map(su.texte_modalite)
        ordre = su.modalites(age)
    d = pd.DataFrame({"classe": classe, "sexe": sexe}).dropna()
    d = d[d["sexe"].isin(mods) & d["classe"].isin(ordre)]
    tab = d.groupby(["classe", "sexe"]).size().unstack(fill_value=0).reindex(index=ordre, columns=mods, fill_value=0)
    t = tab.stack().rename("effectif").reset_index()
    t.columns = ["classe", "sexe", "effectif"]
    t["valeur"] = np.where(t["sexe"] == mods[0], -t["effectif"], t["effectif"])
    return alt.Chart(t).mark_bar().encode(
        y=alt.Y("classe:N", sort=list(reversed(ordre)), title=str(col_age)),
        x=alt.X("valeur:Q", title="Effectif", axis=alt.Axis(labelExpr="abs(datum.value)")),
        color=alt.Color("sexe:N", sort=mods, title=str(col_sexe), scale=alt.Scale(range=[PALETTE[0], PALETTE[1]])),
        tooltip=[alt.Tooltip("classe:N", title="Classe d'âge"), alt.Tooltip("sexe:N", title="Sexe"),
                 alt.Tooltip("effectif:Q", title="Effectif")]
    ).properties(height=max(240, 24 * len(ordre)), title="Pyramide des âges")


def croisement_empile(df, a, b, hauteur=260):
    """Barres empilées à 100 % : répartition de b dans chaque modalité de a."""
    d = df[[a, b]].dropna()
    t = pd.crosstab(d[a].map(su.texte_modalite), d[b].map(su.texte_modalite))
    t = t.stack().rename("effectif").reset_index()
    t.columns = ["ligne", "colonne", "effectif"]
    ordre_l, ordre_c = su.modalites(d[a]), su.modalites(d[b])
    return alt.Chart(t).mark_bar().encode(
        y=alt.Y("ligne:N", sort=ordre_l, title=str(a)),
        x=alt.X("effectif:Q", stack="normalize", axis=alt.Axis(format="%"), title="Répartition"),
        color=alt.Color("colonne:N", sort=ordre_c, title=str(b)),
        tooltip=[alt.Tooltip("ligne:N", title=str(a)), alt.Tooltip("colonne:N", title=str(b)),
                 alt.Tooltip("effectif:Q", title="Effectif")]
    ).properties(height=hauteur, title=f"{b} selon {a}")


def graphique_manquants(df):
    """Pourcentage de valeurs manquantes par variable (None s'il n'y en a pas)."""
    m = (100 * df.isna().mean()).rename("pct").reset_index()
    m.columns = ["variable", "pct"]
    m = m[m["pct"] > 0].sort_values("pct", ascending=False)
    if m.empty:
        return None
    return alt.Chart(m).mark_bar(color="#d62728", opacity=0.8).encode(
        y=alt.Y("variable:N", sort=list(m["variable"]), title=None),
        x=alt.X("pct:Q", title="% de valeurs manquantes"),
        tooltip=[alt.Tooltip("variable:N", title="Variable"), alt.Tooltip("pct:Q", title="% manquants", format=".1f")]
    ).properties(height=max(80, 24 * len(m)), title="Qualité des données : valeurs manquantes")


# ======================================================================
# 2. Hypothèses testables
# ======================================================================
def _derive(a, b):
    """Une variable recodée (« Age_classes ») est dérivée de sa source (« Age ») : test sans intérêt."""
    return a.startswith(b + "_") or b.startswith(a + "_")


def _decimales(txt):
    return str(txt).replace(".", ",")


def _test_groupes(df, y, g, alpha, n_min):
    groupes = {k: v for k, v in su.groupes_numeriques(df, y, g).items() if len(v) > 0}
    if len(groupes) < 2:
        return None
    if any(len(v) < n_min for v in groupes.values()):
        return "petits"
    n = sum(len(v) for v in groupes.values())
    petit = min(len(v) for v in groupes.values()) < 10
    if len(groupes) == 2:
        a, b = list(groupes)
        r = su.comparer_deux(groupes[a], groupes[b], alpha=alpha)
        ligne = {"t de Welch": "t de Welch (variances inégales)",
                 "Mann-Whitney U": "Mann-Whitney U (non paramétrique)"}[r["reco"]]
        effet = r["tests"].loc[ligne, "Taille d'effet"]
    else:
        r = su.comparer_k(groupes, alpha=alpha)
        ligne = {"ANOVA à un facteur": "ANOVA à un facteur", "ANOVA de Welch": "ANOVA de Welch",
                 "Kruskal-Wallis": "Kruskal-Wallis (non paramétrique)"}[r["reco"]]
        effet = r["tests"].loc[ligne, "Taille d'effet"]
        if effet == "—":
            effet = r["tests"].loc["ANOVA à un facteur", "Taille d'effet"]
    return {"p": float(r["p_reco"]), "test": r["reco"], "effet": _decimales(effet), "n": n,
            "prudence": "Petits groupes (n < 10)" if petit else "", "nb_groupes": len(groupes)}


def _test_khi2(df, a, b):
    r = su.tester_khi2(df, a, b)
    if r["n"] < 20:
        return None
    if r["table_2x2"] and float(r["attendus"].values.min()) < 5:
        test, p = "Test exact de Fisher", float(r["p_fisher"])
    else:
        test, p = "Khi² de Pearson", float(r["p"])
    return {"p": p, "test": test, "effet": _decimales(f"V de Cramér = {r['cramer_v']:.2f}"), "n": r["n"],
            "prudence": "Effectifs théoriques faibles" if r["pct_attendus_faibles"] > 20 else ""}


def suggerer_tests(df, alpha=0.05, prioritaires=(), exclure=(), max_quant=12, max_qual=10, n_min=5):
    """Teste automatiquement les hypothèses plausibles entre les variables de la base.

    Retourne (tableau, specs, infos, normalite) :
      - tableau   : un test par ligne, trié par p-value, avec p ajustée (Benjamini-Hochberg) ;
      - specs     : paramètres de chaque ligne (pour l'ajouter au rapport), alignés sur le tableau ;
      - infos     : compteurs (tests réalisés, ignorés…) ;
      - normalite : diagnostic de normalité des variables quantitatives.
    """
    prioritaires, exclure = list(prioritaires), set(exclure)
    quant, qual = [], []
    for c in df.columns:
        s = df[c]
        if c in exclure or _est_identifiant(s, c) or pd.api.types.is_datetime64_any_dtype(s):
            continue
        k = s.nunique(dropna=True)
        if pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s):
            if k > 10 and s.notna().sum() >= 10:
                quant.append(c)
            elif 2 <= k <= 8:
                qual.append(c)
        elif 2 <= k <= 8:
            qual.append(c)
    ordre = {c: i for i, c in enumerate(df.columns)}
    tri = lambda L: sorted(L, key=lambda c: (c not in prioritaires, ordre[c]))
    quant, qual = tri(quant)[:max_quant], tri(qual)[:max_qual]

    # normalité / symétrie
    lignes_norm, sym = [], {}
    for c in quant:
        v = pd.to_numeric(df[c], errors="coerce").dropna()
        asym = float(stats.skew(v, bias=False)) if len(v) > 2 else np.nan
        sym[c] = bool(abs(asym) <= 1) if not np.isnan(asym) else False
        ps = float(stats.shapiro(v.sample(5000, random_state=1) if len(v) > 5000 else v)[1]) if len(v) >= 3 and v.nunique() > 1 else np.nan
        normal = (not np.isnan(ps)) and ps >= alpha
        if len(v) < 30:
            conseil = ("Paramétriques (t, ANOVA, Pearson)" if su.est_normal(v, alpha)
                       else "Non paramétriques (Mann-Whitney, Kruskal-Wallis, Spearman)")
        elif sym[c]:
            conseil = "Paramétriques (N ≥ 30, distribution peu asymétrique)"
        else:
            conseil = "Non paramétriques plus prudents (N ≥ 30 mais asymétrie forte)"
        lignes_norm.append({"Variable": c, "N": len(v), "Asymétrie": round(asym, 2) if not np.isnan(asym) else np.nan,
                            "Shapiro-Wilk (p)": ps,
                            "Distribution": "Compatible avec la normale" if normal else "S'écarte de la normale",
                            "Tests conseillés": conseil})
    normalite = pd.DataFrame(lignes_norm).set_index("Variable") if lignes_norm else pd.DataFrame()

    lignes, specs = [], []
    ignores = 0

    for y in quant:                                   # comparaisons de groupes
        for g in qual:
            if _derive(y, g):
                continue
            try:
                res = _test_groupes(df, y, g, alpha, n_min)
            except Exception:
                res = None
            if res == "petits":
                ignores += 1
                continue
            if res is None:
                continue
            lignes.append({"Type": "Comparaison de groupes", "Analyse": f"{y} selon {g}", **{k: res[k] for k in ("test", "effet", "n", "p", "prudence")},
                           "_sig": f"{y} diffère selon {g}", "_ns": f"Pas de différence détectée de {y} selon {g}"})
            specs.append({"type": "groupes", "params": {"y": y, "g": g}, "nb_groupes": res["nb_groupes"]})

    for i, a in enumerate(qual):                      # associations entre qualitatives
        for b in qual[i + 1:]:
            if _derive(a, b):
                continue
            try:
                res = _test_khi2(df, a, b)
            except Exception:
                res = None
            if res is None:
                continue
            lignes.append({"Type": "Association (khi-deux)", "Analyse": f"{a} et {b}", **{k: res[k] for k in ("test", "effet", "n", "p", "prudence")},
                           "_sig": f"{a} et {b} sont associées", "_ns": f"Pas d'association détectée entre {a} et {b}"})
            specs.append({"type": "khi2", "params": {"a": a, "b": b}, "nb_groupes": 2})

    if len(quant) >= 2:                               # corrélations
        _, _, _, lp = su.correlations(df, quant, "pearson")
        _, _, _, ls = su.correlations(df, quant, "spearman")
        idx = {(r["Variable 1"], r["Variable 2"]): r for _, r in lp.iterrows()}
        idx_s = {(r["Variable 1"], r["Variable 2"]): r for _, r in ls.iterrows()}
        for (a, b), rp_ in idx.items():
            if _derive(a, b):
                continue
            pearson = sym.get(a) and sym.get(b)
            ligne = rp_ if pearson else idx_s[(a, b)]
            if pd.isna(ligne["p-value"]) or ligne["N"] < 10:
                continue
            symb = "r" if pearson else "ρ"
            signe = "positivement" if ligne["r"] > 0 else "négativement"
            lignes.append({"Type": "Corrélation", "Analyse": f"{a} et {b}",
                           "test": "Pearson" if pearson else "Spearman",
                           "effet": _decimales(f"{symb} = {ligne['r']:.2f} ({ligne['Force']})"),
                           "n": int(ligne["N"]), "p": float(ligne["p-value"]),
                           "prudence": "Petit effectif (N < 30)" if ligne["N"] < 30 else "",
                           "_sig": f"{a} et {b} sont corrélées {signe} ({ligne['Force']})",
                           "_ns": f"Pas de corrélation détectée entre {a} et {b}"})
            specs.append({"type": "correlation", "params": {"x": a, "y": b, "methode": "pearson" if pearson else "spearman"},
                          "nb_groupes": 2})

    if not lignes:
        return pd.DataFrame(), [], {"realises": 0, "ignores": ignores, "quant": quant, "qual": qual}, normalite

    p = np.array([l["p"] for l in lignes], float)
    padj = multipletests(p, method="fdr_bh")[1]
    ordre_p = np.argsort(p, kind="stable")
    sortie, specs_tries = [], []
    for rang, i in enumerate(ordre_p):
        l = lignes[i]
        if padj[i] < alpha:
            verdict, lecture = "Significatif (corrigé)", l["_sig"]
        elif p[i] < alpha:
            verdict, lecture = "Significatif avant correction seulement", l["_sig"] + " (à confirmer)"
        else:
            verdict, lecture = "Non significatif", l["_ns"]
        sortie.append({"Type": l["Type"], "Analyse": l["Analyse"], "Test utilisé": l["test"],
                       "Effet / intensité": l["effet"], "N": l["n"], "p-value": su.formater_p(p[i]),
                       "p ajustée (FDR)": su.formater_p(padj[i]), "Verdict": verdict, "Lecture": lecture,
                       "Prudence": l["prudence"], "_p": p[i], "_padj": padj[i]})
        specs_tries.append(specs[i])
    tableau = pd.DataFrame(sortie)
    tableau.index = pd.RangeIndex(1, len(tableau) + 1, name="N°")
    infos = {"realises": len(lignes), "ignores": ignores, "quant": quant, "qual": qual,
             "sig_brut": int((p < alpha).sum()), "sig_corrige": int((padj < alpha).sum())}
    return tableau, specs_tries, infos, normalite


# ======================================================================
# 3. Coordonnées géographiques
# ======================================================================
_LAT = {"latitude", "lat", "latitud", "latitudes"}
_LON = {"longitude", "lon", "lng", "long", "longitud", "longitudes"}
_GPS = {"gps", "geopoint", "geolocation", "coordonnees", "coordinates", "geo", "localisation", "position"}
_RAYON_TERRE_KM = 6371.0088


def _nombres(s):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").astype(float)
    return pd.to_numeric(s.astype("string").str.strip().str.replace(",", ".", regex=False), errors="coerce").astype(float)


def _plage_ok(s, limite):
    v = _nombres(s).dropna()
    return len(v) >= 3 and (v.abs() <= limite).mean() >= 0.8


def detecter_coordonnees(df):
    """(colonne_lat, colonne_lon) ou (colonne_combinée, None) ; None si rien n'est trouvé."""
    lat = lon = None
    for c in df.columns:
        t = set(_norm(c).split())
        if lat is None and t & _LAT and _plage_ok(df[c], 180):
            lat = c
        elif lon is None and t & _LON and _plage_ok(df[c], 180):
            lon = c
    if lat and lon:
        return lat, lon
    noms = {_norm(c): c for c in df.columns}                      # colonnes « x » et « y »
    cx = next((noms[n] for n in ("x", "coord x", "coordonnee x", "coordx") if n in noms), None)
    cy = next((noms[n] for n in ("y", "coord y", "coordonnee y", "coordy") if n in noms), None)
    if cx and cy and _plage_ok(df[cx], 180) and _plage_ok(df[cy], 90):
        return cy, cx
    for c in df.columns:                                           # une seule colonne « lat, lon »
        if set(_norm(c).split()) & _GPS and not pd.api.types.is_numeric_dtype(df[c]):
            pts, _ = extraire_points(df, c)
            n_val = df[c].notna().sum()
            if n_val and pts["valide"].sum() >= 0.5 * n_val:
                return c, None
    return None


def _haversine(lat1, lon1, lat2, lon2):
    p1, p2 = np.radians(lat1), np.radians(lat2)
    a = np.sin((p2 - p1) / 2) ** 2 + np.cos(p1) * np.cos(p2) * np.sin(np.radians(lon2 - lon1) / 2) ** 2
    return 2 * _RAYON_TERRE_KM * np.arcsin(np.sqrt(a))


def extraire_points(df, col_lat, col_lon=None):
    """Retourne (points, notes). `points` : lat, lon, valide, aberrant (index = celui de df)."""
    notes = []
    if col_lon is None:
        ext = df[col_lat].astype("string").str.extract(r"(-?\d+(?:[.,]\d+)?)\s*[,;\s]\s*(-?\d+(?:[.,]\d+)?)")
        lat, lon = _nombres(ext[0]), _nombres(ext[1])
    else:
        lat, lon = _nombres(df[col_lat]), _nombres(df[col_lon])
    la, lo = lat.dropna(), lon.dropna()
    if len(la) and len(lo) and (la.abs() > 90).mean() > 0.5 and (lo.abs() <= 90).mean() > 0.9:
        lat, lon = lon, lat
        notes.append("La latitude et la longitude semblaient inversées : elles ont été échangées pour la carte.")
    valide = lat.between(-90, 90) & lon.between(-180, 180) & ~((lat == 0) & (lon == 0))
    pts = pd.DataFrame({"lat": lat, "lon": lon, "valide": valide})
    pts["aberrant"] = False
    v = pts[pts["valide"]]
    if len(v) >= 10:
        d = _haversine(v["lat"], v["lon"], v["lat"].median(), v["lon"].median())
        q1, q3 = d.quantile(0.25), d.quantile(0.75)
        pts.loc[v.index, "aberrant"] = d > max(10.0, q3 + 3 * (q3 - q1))
    invalides = int((lat.notna() & lon.notna() & ~valide).sum())
    manquants = int((lat.isna() | lon.isna()).sum())
    if invalides:
        notes.append(f"{invalides} coordonnée(s) invalide(s) (hors des limites ou égales à 0, 0) : ignorée(s).")
    if manquants:
        notes.append(f"{manquants} observation(s) sans coordonnées complètes.")
    if int(pts["aberrant"].sum()):
        notes.append(f"{int(pts['aberrant'].sum())} point(s) très éloigné(s) du reste de l'échantillon "
                     "(erreur de saisie possible) : vous pouvez les masquer.")
    return pts, notes


def resume_zone(pts):
    """Emprise de la zone d'étude à partir de points valides (colonnes lat, lon)."""
    from scipy.spatial import ConvexHull
    lat0, lon0 = float(pts["lat"].median()), float(pts["lon"].median())
    x = np.radians(pts["lon"] - lon0) * _RAYON_TERRE_KM * np.cos(np.radians(lat0))
    y = np.radians(pts["lat"] - lat0) * _RAYON_TERRE_KM
    res = {"n": len(pts), "centre": (lat0, lon0),
           "lat_min": float(pts["lat"].min()), "lat_max": float(pts["lat"].max()),
           "lon_min": float(pts["lon"].min()), "lon_max": float(pts["lon"].max()),
           "ns_km": float(y.max() - y.min()), "eo_km": float(x.max() - x.min()),
           "aire_km2": np.nan, "hull": None, "x": x.to_numpy(), "y": y.to_numpy()}
    if len(pts) >= 3:
        try:
            h = ConvexHull(np.c_[x, y])
            res["aire_km2"] = float(h.volume)
            res["hull"] = np.c_[x.to_numpy()[h.vertices], y.to_numpy()[h.vertices]]
        except Exception:
            pass
    return res


def schema_emprise(res, couleurs=None, table_couleurs=None):
    """Schéma de l'emprise en kilomètres (repère orthonormé) : points, enveloppe convexe, centre."""
    pts = pd.DataFrame({"x": res["x"], "y": res["y"]})
    if couleurs is not None:
        pts["groupe"] = list(couleurs)
    demi = max(res["eo_km"], res["ns_km"], 0.2) / 2 * 1.15
    ech = lambda: alt.Scale(domain=[-demi, demi])
    p = alt.Chart(pts).mark_circle(size=40, opacity=0.65).encode(
        x=alt.X("x:Q", title="Distance est–ouest au centre (km)", scale=ech()),
        y=alt.Y("y:Q", title="Distance nord–sud au centre (km)", scale=ech()),
        **({"color": alt.Color("groupe:N", title="Groupe",
                               scale=alt.Scale(domain=list(table_couleurs), range=list(table_couleurs.values()))
                               if table_couleurs else alt.Undefined)}
           if couleurs is not None else {"color": alt.value(BLEU)}))
    couches = [p]
    if res["hull"] is not None:
        h = pd.DataFrame(np.vstack([res["hull"], res["hull"][:1]]), columns=["x", "y"])
        h["ordre"] = range(len(h))
        couches.append(alt.Chart(h).mark_line(color="#d62728", strokeWidth=2).encode(x="x:Q", y="y:Q", order="ordre:Q"))
    couches.append(alt.Chart(pd.DataFrame({"x": [0.0], "y": [0.0]})).mark_point(
        shape="cross", size=160, color="black", strokeWidth=2).encode(x="x:Q", y="y:Q"))
    return alt.layer(*couches).properties(width=420, height=420, title="Emprise de la zone (enveloppe convexe en rouge)")


def couleurs_modalites(serie, maximum=10):
    """Associe une couleur hexadécimale à chaque modalité (les moins fréquentes sont regroupées en « Autres »)."""
    txt = serie.map(su.texte_modalite)
    ordre = list(txt.value_counts().index)
    garde = ordre[:maximum]
    table = {m: PALETTE[i % len(PALETTE)] for i, m in enumerate(garde)}
    if len(ordre) > maximum:
        table["Autres"] = GRIS
        txt = txt.where(txt.isin(garde), "Autres")
    return txt.map(table), table, txt
