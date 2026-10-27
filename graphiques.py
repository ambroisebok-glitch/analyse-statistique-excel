# -*- coding: utf-8 -*-
"""
Graphiques de StatExcel (Altair), utilisés par app.py.

Auteur : BOKOUMBO Tibetobe Ambroise
"""
import altair as alt
import numpy as np
import pandas as pd
from scipy import stats

from stats_utils import modalites, texte_modalite

alt.data_transformers.disable_max_rows()

BLEU, ROUGE = "#1f77b4", "#d62728"


def _num(s):
    return pd.to_numeric(pd.Series(s), errors="coerce").dropna()


def histogramme(s, nom, classes=20, courbe="Aucune"):
    """Histogramme avec repères moyenne/médiane et courbe de densité ou de loi normale en option."""
    v = _num(s)
    if len(v) < 2 or v.nunique() < 2:
        raise ValueError("Effectif insuffisant ou variable constante.")
    comptes, bords = np.histogram(v, bins=int(classes))
    larg = bords[1] - bords[0]
    d = pd.DataFrame({"debut": bords[:-1], "fin": bords[1:], "effectif": comptes})
    axe_x = alt.X("debut:Q", title=nom)
    barres = alt.Chart(d).mark_bar(color=BLEU, opacity=0.75, stroke="white").encode(
        x=axe_x, x2="fin:Q", y=alt.Y("effectif:Q", title="Effectif"), y2=alt.datum(0),
        tooltip=[alt.Tooltip("debut:Q", title="De", format=".4g"),
                 alt.Tooltip("fin:Q", title="À", format=".4g"),
                 alt.Tooltip("effectif:Q", title="Effectif")])
    couches = [barres]
    if courbe != "Aucune":
        xs = np.linspace(bords[0], bords[-1], 200)
        ys = stats.gaussian_kde(v)(xs) if courbe.startswith("Densité") else stats.norm.pdf(xs, v.mean(), v.std(ddof=1))
        c = pd.DataFrame({"x": xs, "y": ys * len(v) * larg})
        couches.append(alt.Chart(c).mark_line(color=ROUGE, strokeWidth=2).encode(
            x=alt.X("x:Q", title=nom), y=alt.Y("y:Q", title="Effectif")))
    rep = pd.DataFrame({"valeur": [v.mean(), v.median()], "Repère": ["Moyenne", "Médiane"]})
    couches.append(alt.Chart(rep).mark_rule(strokeDash=[5, 3], strokeWidth=2).encode(
        x=alt.X("valeur:Q", title=nom),
        color=alt.Color("Repère:N", scale=alt.Scale(domain=["Moyenne", "Médiane"], range=["#2ca02c", "#ff7f0e"])),
        tooltip=[alt.Tooltip("Repère:N"), alt.Tooltip("valeur:Q", title="Valeur", format=".4g")]))
    return alt.layer(*couches).properties(height=320, title=f"Distribution de {nom}")


def boite(df, col, groupe=None):
    """Boîte à moustaches (losange = moyenne), éventuellement par groupe."""
    d = pd.DataFrame({"valeur": pd.to_numeric(df[col], errors="coerce")})
    if groupe:
        d["groupe"] = df[groupe].map(texte_modalite)
        ordre = modalites(df[groupe])
    d = d.dropna()
    if len(d) < 2:
        raise ValueError("Effectif insuffisant.")
    y = alt.Y("valeur:Q", title=col, scale=alt.Scale(zero=False))
    ymoy = alt.Y("mean(valeur):Q", title=col, scale=alt.Scale(zero=False))
    if groupe:
        x = alt.X("groupe:N", sort=ordre, title=groupe, axis=alt.Axis(labelAngle=-30))
        boites = alt.Chart(d).mark_boxplot(extent=1.5, size=40).encode(
            x=x, y=y, color=alt.Color("groupe:N", sort=ordre, legend=None))
        moy = alt.Chart(d).mark_point(shape="diamond", filled=True, color="black", size=70).encode(x=x, y=ymoy)
        titre = f"{col} selon {groupe}"
    else:
        boites = alt.Chart(d).mark_boxplot(extent=1.5, size=60, color=BLEU).encode(y=y)
        moy = alt.Chart(d).mark_point(shape="diamond", filled=True, color="black", size=70).encode(y=ymoy)
        titre = f"Boîte à moustaches de {col}"
    return alt.layer(boites, moy).properties(height=340, title=titre)


def qq(s, nom):
    """Diagramme Q-Q contre la loi normale."""
    v = _num(s)
    if len(v) < 3 or v.nunique() < 2:
        raise ValueError("Effectif insuffisant ou variable constante.")
    (osm, osr), (pente, ordonnee, _) = stats.probplot(v, dist="norm")
    pts = pd.DataFrame({"theorique": osm, "observe": osr})
    droite = pd.DataFrame({"theorique": [osm.min(), osm.max()],
                           "observe": [ordonnee + pente * osm.min(), ordonnee + pente * osm.max()]})
    ech_x, ech_y = alt.Scale(zero=False), alt.Scale(zero=False)
    pc = alt.Chart(pts).mark_point(filled=True, opacity=0.6, color=BLEU).encode(
        x=alt.X("theorique:Q", title="Quantiles théoriques (loi normale)", scale=ech_x),
        y=alt.Y("observe:Q", title=f"Quantiles observés — {nom}", scale=ech_y))
    lc = alt.Chart(droite).mark_line(color=ROUGE, strokeWidth=2).encode(
        x=alt.X("theorique:Q", scale=ech_x), y=alt.Y("observe:Q", scale=ech_y))
    return (pc + lc).properties(height=340, title=f"Diagramme Q-Q de {nom}")


def frequences_graphique(freq, nom, type_graphique="Barres", max_modalites=20):
    """Barres ou secteurs à partir du tableau produit par stats_utils.frequences()."""
    t = freq.drop(index="Manquant", errors="ignore").reset_index().head(max_modalites)
    if t.empty:
        raise ValueError("Aucune valeur à représenter.")
    t = t.rename(columns={"Modalité": "modalite", "Effectif": "effectif", "% valides": "pct"})
    t["ordre"] = range(len(t))
    t["etiquette"] = t["pct"].map(lambda p: f"{p:.1f} %".replace(".", ","))
    ordre = list(t["modalite"])
    info = [alt.Tooltip("modalite:N", title="Modalité"), alt.Tooltip("effectif:Q", title="Effectif"),
            alt.Tooltip("pct:Q", title="% valides", format=".1f")]
    if type_graphique == "Secteurs":
        return alt.Chart(t).mark_arc(innerRadius=60, outerRadius=130).encode(
            theta=alt.Theta("effectif:Q", stack=True), order=alt.Order("ordre:Q"),
            color=alt.Color("modalite:N", sort=ordre, title=nom), tooltip=info
        ).properties(height=340, title=f"Répartition de {nom}")
    x = alt.X("modalite:N", sort=ordre, title=nom, axis=alt.Axis(labelAngle=-30))
    barres = alt.Chart(t).mark_bar(color=BLEU).encode(x=x, y=alt.Y("effectif:Q", title="Effectif"), tooltip=info)
    texte = alt.Chart(t).mark_text(dy=-7).encode(x=x, y="effectif:Q", text="etiquette:N")
    return (barres + texte).properties(height=340, title=f"Répartition de {nom}")


def nuage(df, x, y, groupe=None, droite=True):
    """Nuage de points avec droite de régression (une par groupe si un groupe est choisi)."""
    d = pd.DataFrame({"x": pd.to_numeric(df[x], errors="coerce"), "y": pd.to_numeric(df[y], errors="coerce")})
    if groupe:
        d["groupe"] = df[groupe].map(texte_modalite)
    d = d.dropna()
    if len(d) < 3:
        raise ValueError("Effectif insuffisant.")
    pts = alt.Chart(d).mark_circle(size=55, opacity=0.6, color=BLEU).encode(
        x=alt.X("x:Q", title=x, scale=alt.Scale(zero=False)),
        y=alt.Y("y:Q", title=y, scale=alt.Scale(zero=False)),
        tooltip=[alt.Tooltip("x:Q", title=x), alt.Tooltip("y:Q", title=y)])
    if groupe:
        pts = pts.encode(color=alt.Color("groupe:N", title=groupe, sort=modalites(df[groupe])))
    if not droite:
        return pts.properties(height=380, title=f"{y} selon {x}")
    reg = pts.transform_regression("x", "y", **({"groupby": ["groupe"]} if groupe else {})).mark_line(size=3)
    if not groupe:
        reg = reg.mark_line(size=3, color=ROUGE)
    return (pts + reg).properties(height=380, title=f"{y} selon {x}")


def carte_correlations(r):
    """Carte de chaleur à partir d'une matrice de corrélations."""
    d = pd.DataFrame([(a, b, r.loc[a, b]) for a in r.index for b in r.columns], columns=["v1", "v2", "r"])
    d["etiquette"] = d["r"].map(lambda v: "" if pd.isna(v) else f"{v:.2f}".replace(".", ","))
    base = alt.Chart(d).encode(
        x=alt.X("v1:N", sort=list(r.columns), title=None, axis=alt.Axis(labelAngle=-40)),
        y=alt.Y("v2:N", sort=list(r.index), title=None))
    rect = base.mark_rect().encode(
        color=alt.Color("r:Q", title="r", scale=alt.Scale(domain=[-1, 1], scheme="redblue")),
        tooltip=[alt.Tooltip("v1:N", title="Variable 1"), alt.Tooltip("v2:N", title="Variable 2"),
                 alt.Tooltip("r:Q", title="r", format=".3f")])
    texte = base.mark_text(fontSize=12).encode(
        text="etiquette:N", color=alt.condition("abs(datum.r) > 0.6", alt.value("white"), alt.value("black")))
    taille = max(260, 56 * len(r))
    return (rect + texte).properties(width=taille, height=taille, title="Matrice de corrélation")
