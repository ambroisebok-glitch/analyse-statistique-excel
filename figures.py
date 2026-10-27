# -*- coding: utf-8 -*-
"""Figures PNG destinées au rapport Word de StatExcel."""
import io
import math
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy import stats

def _png(fig):
    buf = io.BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()

def _serie(s):
    return pd.to_numeric(pd.Series(s), errors="coerce").dropna()

def normalite(s, nom):
    v = _serie(s)
    fig, axes = plt.subplots(1, 2, figsize=(10, 4))
    axes[0].hist(v, bins="auto", edgecolor="white")
    axes[0].set_title(f"Histogramme — {nom}")
    axes[0].set_xlabel(nom); axes[0].set_ylabel("Effectif")
    if len(v) >= 3 and v.nunique() > 1:
        stats.probplot(v, dist="norm", plot=axes[1])
        axes[1].set_title("Diagramme Q-Q")
    else:
        axes[1].text(.5, .5, "Effectif insuffisant", ha="center", va="center")
        axes[1].set_axis_off()
    return _png(fig)

def grille_histogrammes(df, cols):
    cols = list(cols)
    n = len(cols)
    ncols = min(3, max(1, n)); nrows = math.ceil(n / ncols)
    fig, axes = plt.subplots(nrows, ncols, figsize=(5*ncols, 3.2*nrows), squeeze=False)
    for ax, col in zip(axes.flat, cols):
        v = _serie(df[col])
        ax.hist(v, bins="auto", edgecolor="white")
        ax.set_title(str(col)); ax.set_ylabel("Effectif")
    for ax in list(axes.flat)[n:]: ax.set_visible(False)
    return _png(fig)

def grille_barres(df, cols):
    cols = list(cols)
    n = len(cols); ncols = min(2, max(1, n)); nrows = math.ceil(n/ncols)
    fig, axes = plt.subplots(nrows, ncols, figsize=(6*ncols, 3.5*nrows), squeeze=False)
    for ax, col in zip(axes.flat, cols):
        vc = df[col].fillna("Manquant").astype(str).value_counts().head(12)
        ax.bar(vc.index.astype(str), vc.values)
        ax.set_title(str(col)); ax.set_ylabel("Effectif")
        ax.tick_params(axis="x", rotation=35)
    for ax in list(axes.flat)[n:]: ax.set_visible(False)
    return _png(fig)

def boite_groupes(groupes, y, g):
    fig, ax = plt.subplots(figsize=(7, 4))
    names, vals = list(groupes.keys()), [np.asarray(v) for v in groupes.values()]
    ax.boxplot(vals, labels=[str(n) for n in names], showmeans=True)
    ax.set_title(f"{y} selon {g}"); ax.set_ylabel(y)
    ax.tick_params(axis="x", rotation=25)
    return _png(fig)

def barres_empilees(pct, a, b):
    fig, ax = plt.subplots(figsize=(8, 4.5))
    pct.plot(kind="bar", stacked=True, ax=ax)
    ax.set_title(f"Répartition de {b} selon {a}")
    ax.set_xlabel(a); ax.set_ylabel("Pourcentage en ligne (%)")
    ax.legend(title=b, bbox_to_anchor=(1.02, 1), loc="upper left")
    ax.tick_params(axis="x", rotation=25)
    return _png(fig)

def nuage_reg(x, y, nom_x, nom_y):
    dx, dy = pd.to_numeric(pd.Series(x), errors="coerce"), pd.to_numeric(pd.Series(y), errors="coerce")
    d = pd.DataFrame({"x": dx, "y": dy}).dropna()
    fig, ax = plt.subplots(figsize=(6.5, 4.5))
    ax.scatter(d.x, d.y, alpha=.7)
    if len(d) >= 2 and d.x.nunique() > 1:
        pente, intercept = np.polyfit(d.x, d.y, 1)
        xs = np.linspace(d.x.min(), d.x.max(), 100)
        ax.plot(xs, intercept + pente*xs)
    ax.set_xlabel(nom_x); ax.set_ylabel(nom_y); ax.set_title(f"{nom_y} selon {nom_x}")
    return _png(fig)

def boite_appariees(a, b, nom_a, nom_b):
    d = pd.DataFrame({"a": pd.to_numeric(pd.Series(a), errors="coerce"),
                      "b": pd.to_numeric(pd.Series(b), errors="coerce")}).dropna()
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.boxplot([d.a, d.b], labels=[nom_a, nom_b], showmeans=True)
    ax.set_title("Comparaison des mesures appariées")
    return _png(fig)

def diag_regression(modele):
    resid = pd.Series(modele.resid).dropna()
    fitted = pd.Series(modele.fittedvalues).reindex(resid.index)
    fig, axes = plt.subplots(1, 2, figsize=(10, 4))
    axes[0].scatter(fitted, resid, alpha=.65)
    axes[0].axhline(0, linestyle="--")
    axes[0].set_xlabel("Valeurs ajustées"); axes[0].set_ylabel("Résidus")
    axes[0].set_title("Résidus vs valeurs ajustées")
    if len(resid) >= 3 and resid.nunique() > 1:
        stats.probplot(resid, dist="norm", plot=axes[1])
        axes[1].set_title("Q-Q des résidus")
    else:
        axes[1].text(.5, .5, "Effectif insuffisant", ha="center", va="center")
        axes[1].set_axis_off()
    return _png(fig)

def barres_items(detail):
    col = "Corrélation item-total corrigée"
    vals = pd.to_numeric(detail[col], errors="coerce")
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.bar([str(x) for x in detail.index], vals.values)
    ax.axhline(.3, linestyle="--")
    ax.set_title("Corrélation item-total corrigée")
    ax.set_ylabel("Corrélation")
    ax.tick_params(axis="x", rotation=35)
    return _png(fig)
