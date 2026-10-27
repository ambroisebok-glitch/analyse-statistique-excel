# -*- coding: utf-8 -*-
"""
Fonctions statistiques de StatExcel (sans interface), utilisées par app.py.
"""
import itertools

AUTEUR = "BOKOUMBO Tibetobe Ambroise"

import numpy as np
import pandas as pd
import statsmodels.formula.api as smf
from scipy import stats
from statsmodels.stats.diagnostic import het_breuschpagan, lilliefors
from statsmodels.stats.multicomp import pairwise_tukeyhsd
from statsmodels.stats.multitest import multipletests
from statsmodels.stats.oneway import anova_oneway


# ---------------------------------------------------------------- mise en forme
def formater_p(p):
    if p is None or (isinstance(p, float) and np.isnan(p)):
        return "—"
    return "< 0,001" if p < 0.001 else f"{p:.3f}".replace(".", ",")


def texte_p(p):
    if p is None or (isinstance(p, float) and np.isnan(p)):
        return "p non calculable"
    return "p < 0,001" if p < 0.001 else "p = " + f"{p:.3f}".replace(".", ",")


def conclusion(p, alpha=0.05, sujet="La différence"):
    if p is None or (isinstance(p, float) and np.isnan(p)):
        return "Test non calculable."
    if p < alpha:
        return f"{sujet} est statistiquement significative ({texte_p(p)}, seuil α = {alpha})."
    return f"{sujet} n'est pas statistiquement significative ({texte_p(p)}, seuil α = {alpha})."


def force_effet(valeur, seuils, libelles):
    v = abs(valeur)
    for s, l in zip(seuils, libelles):
        if v < s:
            return l
    return libelles[-1]


# ---------------------------------------------------------------- préparation
def nettoyer(df):
    """Noms de colonnes propres, lignes/colonnes vides retirées, nombres écrits en texte convertis."""
    df = df.copy()
    noms, vus = [], {}
    for c in df.columns:
        c = str(c).strip()
        if c in vus:
            vus[c] += 1
            c = f"{c}_{vus[c]}"
        else:
            vus[c] = 0
        noms.append(c)
    df.columns = noms
    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_numeric_dtype(s) or pd.api.types.is_datetime64_any_dtype(s):
            continue
        txt = (s.astype("string").str.strip().str.replace("\u00a0", "", regex=False)
               .str.replace(" ", "", regex=False).str.replace(",", ".", regex=False))
        conv = pd.to_numeric(txt, errors="coerce")
        valides = s.notna().sum()
        if valides > 0 and conv.notna().sum() >= 0.9 * valides:
            df[c] = conv
    return df


def colonnes_numeriques(df):
    return [c for c in df.columns
            if pd.api.types.is_numeric_dtype(df[c]) and not pd.api.types.is_bool_dtype(df[c])]


def colonnes_groupes(df, max_modalites=30):
    return [c for c in df.columns if 2 <= df[c].nunique(dropna=True) <= max_modalites]


def resume_variables(df):
    lignes = []
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_bool_dtype(s):
            t = "Qualitative"
        elif pd.api.types.is_numeric_dtype(s):
            t = "Numérique"
        elif pd.api.types.is_datetime64_any_dtype(s):
            t = "Date"
        else:
            t = "Qualitative"
        exemple = s.dropna().iloc[0] if s.notna().any() else ""
        lignes.append({
            "Variable": c, "Type détecté": t, "Valides": int(s.notna().sum()),
            "Manquants": int(s.isna().sum()), "% manquants": round(100 * s.isna().mean(), 1),
            "Modalités distinctes": int(s.nunique(dropna=True)), "Exemple": str(exemple),
        })
    return pd.DataFrame(lignes).set_index("Variable")


# ---------------------------------------------------------------- descriptives
def decrire_numeriques(df, cols):
    lignes = []
    for c in cols:
        s = pd.to_numeric(df[c], errors="coerce")
        v = s.dropna()
        n = len(v)
        ligne = {"Variable": c, "N valides": n, "Manquants": int(s.isna().sum()),
                 "% manquants": 100 * s.isna().mean()}
        if n > 0:
            sd = v.std(ddof=1) if n > 1 else np.nan
            se = sd / np.sqrt(n) if n > 1 else np.nan
            marge = stats.t.ppf(0.975, n - 1) * se if n > 1 else np.nan
            m = v.mean()
            ligne.update({
                "Moyenne": m, "Écart-type": sd, "Erreur std": se,
                "IC95% inf": m - marge, "IC95% sup": m + marge,
                "Min": v.min(), "Q1": v.quantile(0.25), "Médiane": v.median(),
                "Q3": v.quantile(0.75), "Max": v.max(),
                "CV (%)": 100 * sd / m if (m != 0 and n > 1) else np.nan,
                "Asymétrie": stats.skew(v, bias=False) if n > 2 else np.nan,
                "Aplatissement": stats.kurtosis(v, bias=False) if n > 3 else np.nan,
            })
        lignes.append(ligne)
    return pd.DataFrame(lignes).set_index("Variable")


def decrire_par_groupe(df, cols, g):
    blocs = []
    for c in cols:
        d = df[[c, g]].copy()
        d[c] = pd.to_numeric(d[c], errors="coerce")
        t = d.groupby(g, observed=True)[c].agg(
            N="count", Moyenne="mean", ET="std", Médiane="median", Min="min", Max="max")
        t = t.rename(columns={"ET": "Écart-type"})
        t.insert(0, "Variable", c)
        blocs.append(t.reset_index())
    return pd.concat(blocs, ignore_index=True)


def frequences(s):
    vc = s.value_counts(dropna=False)
    t = vc.rename("Effectif").to_frame()
    t["%"] = 100 * t["Effectif"] / t["Effectif"].sum()
    valides = s.notna().sum()
    manquant = pd.isna(t.index)
    t["% valides"] = np.where(manquant, np.nan, 100 * t["Effectif"] / max(valides, 1))
    t["% cumulé (valides)"] = pd.Series(t["% valides"]).fillna(0).cumsum().to_numpy()
    t.index = ["Manquant" if pd.isna(i) else str(i) for i in t.index]
    t.index.name = "Modalité"
    return t


# ---------------------------------------------------------------- normalité
def tests_normalite(v, alpha=0.05):
    v = pd.to_numeric(pd.Series(v), errors="coerce").dropna()
    n = len(v)
    lignes = []
    if n >= 3 and v.nunique() > 1:
        if n <= 5000:
            w, p = stats.shapiro(v)
            lignes.append(("Shapiro-Wilk", w, p))
        if n >= 8:
            k2, p = stats.normaltest(v)
            lignes.append(("D'Agostino-Pearson", k2, p))
        if n >= 5:
            ks, p = lilliefors(v, dist="norm")
            lignes.append(("Kolmogorov-Smirnov (Lilliefors)", ks, p))
    t = pd.DataFrame(lignes, columns=["Test", "Statistique", "p-value"])
    if not t.empty:
        t["Conclusion"] = np.where(t["p-value"] >= alpha, "Compatible avec la normalité",
                                   "Écart à la normalité")
    return t


def est_normal(v, alpha=0.05):
    v = pd.to_numeric(pd.Series(v), errors="coerce").dropna()
    if len(v) < 3 or v.nunique() < 2:
        return False
    if len(v) >= 30:          # grands échantillons : théorème central limite
        return True
    return stats.shapiro(v)[1] >= alpha


# ---------------------------------------------------------------- groupes
def _cle_tri(x):
    try:
        return (0, float(x))
    except ValueError:
        return (1, x)


def texte_modalite(x):
    """Convertit une modalité en texte lisible pour les graphiques."""
    if pd.isna(x):
        return "Manquant"
    return str(x).strip()


def modalites(s):
    """Retourne les modalités distinctes, triées pour les graphiques."""
    valeurs = [texte_modalite(x) for x in pd.Series(s).dropna().unique()]
    try:
        return sorted(valeurs, key=_cle_tri)
    except (TypeError, ValueError):
        return sorted(valeurs, key=str)


def groupes_numeriques(df, y, g):
    d = df[[y, g]].copy()
    d[y] = pd.to_numeric(d[y], errors="coerce")
    d = d.dropna()
    d[g] = d[g].astype(str)
    cles = sorted(d[g].unique(), key=_cle_tri)
    return {k: d.loc[d[g] == k, y].to_numpy(float) for k in cles}


def descriptifs_groupes(groupes):
    lignes = []
    for k, v in groupes.items():
        lignes.append({
            "Groupe": k, "N": len(v), "Moyenne": v.mean(),
            "Écart-type": v.std(ddof=1) if len(v) > 1 else np.nan,
            "Médiane": np.median(v), "Min": v.min(), "Max": v.max(),
        })
    return pd.DataFrame(lignes).set_index("Groupe")


def comparer_deux(x1, x2, alpha=0.05):
    x1, x2 = np.asarray(x1, float), np.asarray(x2, float)
    n1, n2 = len(x1), len(x2)
    if n1 < 2 or n2 < 2:
        raise ValueError("Chaque groupe doit contenir au moins 2 observations.")
    m1, m2 = x1.mean(), x2.mean()
    s1, s2 = x1.std(ddof=1), x2.std(ddof=1)

    lev = stats.levene(x1, x2, center="median")
    welch = stats.ttest_ind(x1, x2, equal_var=False)
    stud = stats.ttest_ind(x1, x2, equal_var=True)
    mw = stats.mannwhitneyu(x1, x2, alternative="two-sided")

    v1, v2 = s1 ** 2 / n1, s2 ** 2 / n2
    se = np.sqrt(v1 + v2)
    ddl_w = (v1 + v2) ** 2 / (v1 ** 2 / (n1 - 1) + v2 ** 2 / (n2 - 1)) if se > 0 else np.nan
    marge = stats.t.ppf(1 - alpha / 2, ddl_w) * se
    diff = m1 - m2
    sp = np.sqrt(((n1 - 1) * s1 ** 2 + (n2 - 1) * s2 ** 2) / (n1 + n2 - 2))
    d = diff / sp if sp > 0 else np.nan
    j = 1 - 3 / (4 * (n1 + n2) - 9)
    g = d * j
    prob_sup = mw.statistic / (n1 * n2)

    tests = pd.DataFrame([
        {"Test": "Levene (égalité des variances)", "Statistique": lev.statistic,
         "ddl": f"1 ; {n1 + n2 - 2}", "p-value": lev.pvalue, "Taille d'effet": "—"},
        {"Test": "t de Welch (variances inégales)", "Statistique": welch.statistic,
         "ddl": f"{ddl_w:.1f}", "p-value": welch.pvalue,
         "Taille d'effet": f"g de Hedges = {g:.2f} ({force_effet(g, (0.2, 0.5, 0.8), ('négligeable', 'petit', 'moyen', 'grand'))})"},
        {"Test": "t de Student (variances égales)", "Statistique": stud.statistic,
         "ddl": f"{n1 + n2 - 2}", "p-value": stud.pvalue,
         "Taille d'effet": f"d de Cohen = {d:.2f}"},
        {"Test": "Mann-Whitney U (non paramétrique)", "Statistique": float(mw.statistic),
         "ddl": "—", "p-value": mw.pvalue,
         "Taille d'effet": f"P(X1 > X2) = {prob_sup:.2f}"},
    ]).set_index("Test")

    normaux = est_normal(x1, alpha) and est_normal(x2, alpha)
    if normaux:
        reco, p_reco = "t de Welch", welch.pvalue
        motif = ("les deux groupes sont compatibles avec la normalité (ou assez grands), "
                 "le test de Welch ne suppose pas l'égalité des variances")
    else:
        reco, p_reco = "Mann-Whitney U", mw.pvalue
        motif = "au moins un groupe s'écarte de la normalité avec un petit effectif"
    return {"tests": tests, "reco": reco, "p_reco": p_reco, "motif": motif,
            "diff": diff, "ic": (diff - marge, diff + marge), "hedges_g": g}


def comparer_k(groupes, alpha=0.05):
    tableaux = [np.asarray(v, float) for v in groupes.values()]
    if len(tableaux) < 2 or any(len(a) < 2 for a in tableaux):
        raise ValueError("Il faut au moins 2 groupes, avec au moins 2 observations chacun.")
    n = sum(len(a) for a in tableaux)
    k = len(tableaux)
    lev = stats.levene(*tableaux, center="median")
    f = stats.f_oneway(*tableaux)
    moy_g = np.concatenate(tableaux).mean()
    ssb = sum(len(a) * (a.mean() - moy_g) ** 2 for a in tableaux)
    sst = sum(((a - moy_g) ** 2).sum() for a in tableaux)
    eta2 = ssb / sst if sst > 0 else np.nan
    w = anova_oneway(tableaux, use_var="unequal", welch_correction=True)
    kw = stats.kruskal(*tableaux)
    eps2 = kw.statistic * (n + 1) / (n ** 2 - 1)

    tests = pd.DataFrame([
        {"Test": "Levene (égalité des variances)", "Statistique": lev.statistic,
         "ddl": f"{k - 1} ; {n - k}", "p-value": lev.pvalue, "Taille d'effet": "—"},
        {"Test": "ANOVA à un facteur", "Statistique": f.statistic,
         "ddl": f"{k - 1} ; {n - k}", "p-value": f.pvalue,
         "Taille d'effet": f"η² = {eta2:.3f} ({force_effet(eta2, (0.01, 0.06, 0.14), ('négligeable', 'petit', 'moyen', 'grand'))})"},
        {"Test": "ANOVA de Welch", "Statistique": float(w.statistic),
         "ddl": f"{w.df_num:.0f} ; {w.df_denom:.1f}", "p-value": float(w.pvalue), "Taille d'effet": "—"},
        {"Test": "Kruskal-Wallis (non paramétrique)", "Statistique": float(kw.statistic),
         "ddl": f"{k - 1}", "p-value": float(kw.pvalue), "Taille d'effet": f"ε² = {eps2:.3f}"},
    ]).set_index("Test")

    normaux = all(est_normal(a, alpha) for a in tableaux)
    homogenes = lev.pvalue >= alpha
    if normaux and homogenes:
        reco, p_reco, posthoc = "ANOVA à un facteur", f.pvalue, "Tukey HSD"
        motif = "groupes compatibles avec la normalité et variances homogènes"
    elif normaux:
        reco, p_reco, posthoc = "ANOVA de Welch", float(w.pvalue), "t de Welch (Holm)"
        motif = "groupes compatibles avec la normalité mais variances inégales"
    else:
        reco, p_reco, posthoc = "Kruskal-Wallis", float(kw.pvalue), "Mann-Whitney (Holm)"
        motif = "au moins un groupe s'écarte de la normalité avec un petit effectif"
    return {"tests": tests, "reco": reco, "p_reco": p_reco, "posthoc": posthoc, "motif": motif}


def posthoc_tukey(groupes, alpha=0.05):
    valeurs = np.concatenate([np.asarray(v, float) for v in groupes.values()])
    etiquettes = np.concatenate([[k] * len(v) for k, v in groupes.items()])
    res = pairwise_tukeyhsd(valeurs, etiquettes, alpha=alpha)
    data = res.summary().data
    t = pd.DataFrame(data[1:], columns=data[0])
    for c in ("meandiff", "p-adj", "lower", "upper"):
        t[c] = pd.to_numeric(t[c])
    t = t.rename(columns={"group1": "Groupe 1", "group2": "Groupe 2", "meandiff": "Écart de moyennes (2 − 1)",
                          "p-adj": "p ajustée", "lower": "IC95% inf", "upper": "IC95% sup", "reject": "Significatif"})
    t["Significatif"] = np.where(t["p ajustée"] < alpha, "Oui", "Non")
    return t


def posthoc_paires(groupes, test="welch", methode="holm", alpha=0.05):
    lignes = []
    for (a, xa), (b, xb) in itertools.combinations(groupes.items(), 2):
        if test == "welch":
            r = stats.ttest_ind(xa, xb, equal_var=False)
            stat, p = r.statistic, r.pvalue
        else:
            r = stats.mannwhitneyu(xa, xb, alternative="two-sided")
            stat, p = float(r.statistic), r.pvalue
        lignes.append({"Groupe 1": a, "Groupe 2": b, "Statistique": stat, "p brute": p,
                       "Écart de moyennes (1 − 2)": np.mean(xa) - np.mean(xb)})
    t = pd.DataFrame(lignes)
    t["p ajustée"] = multipletests(t["p brute"], method=methode)[1]
    t["Significatif"] = np.where(t["p ajustée"] < alpha, "Oui", "Non")
    return t


def comparer_apparies(a, b, alpha=0.05):
    d = pd.DataFrame({"a": pd.to_numeric(a, errors="coerce"), "b": pd.to_numeric(b, errors="coerce")}).dropna()
    n = len(d)
    if n < 3:
        raise ValueError("Il faut au moins 3 paires complètes.")
    diff = d["a"] - d["b"]
    tt = stats.ttest_rel(d["a"], d["b"])
    try:
        wil = stats.wilcoxon(d["a"], d["b"])
        w_stat, w_p = float(wil.statistic), wil.pvalue
    except ValueError:
        w_stat, w_p = np.nan, np.nan
    sd = diff.std(ddof=1)
    dz = diff.mean() / sd if sd > 0 else np.nan
    se = sd / np.sqrt(n)
    marge = stats.t.ppf(1 - alpha / 2, n - 1) * se
    tests = pd.DataFrame([
        {"Test": "t apparié", "Statistique": tt.statistic, "ddl": f"{n - 1}", "p-value": tt.pvalue,
         "Taille d'effet": f"d_z = {dz:.2f}"},
        {"Test": "Wilcoxon (rangs signés)", "Statistique": w_stat, "ddl": "—", "p-value": w_p,
         "Taille d'effet": "—"},
    ]).set_index("Test")
    normal = est_normal(diff, alpha)
    reco, p_reco = ("t apparié", tt.pvalue) if normal else ("Wilcoxon", w_p)
    return {"n": n, "tests": tests, "diff_moy": diff.mean(), "ic": (diff.mean() - marge, diff.mean() + marge),
            "reco": reco, "p_reco": p_reco,
            "motif": "les différences sont compatibles avec la normalité" if normal
            else "les différences s'écartent de la normalité"}


# ---------------------------------------------------------------- associations
def tester_khi2(df, a, b):
    d = df[[a, b]].dropna().astype(str)
    ct = pd.crosstab(d[a], d[b])
    if ct.shape[0] < 2 or ct.shape[1] < 2:
        raise ValueError("Chaque variable doit avoir au moins 2 modalités.")
    chi2, p, ddl, attendus = stats.chi2_contingency(ct, correction=False)
    n = ct.values.sum()
    v = np.sqrt(chi2 / (n * min(ct.shape[0] - 1, ct.shape[1] - 1)))
    attendus = pd.DataFrame(attendus, index=ct.index, columns=ct.columns)
    residus = (ct - attendus) / np.sqrt(attendus)
    lignes = [{"Test": "Khi² de Pearson", "Statistique": chi2, "ddl": str(ddl), "p-value": p}]
    if ct.shape == (2, 2):
        c2y, py, _, _ = stats.chi2_contingency(ct, correction=True)
        lignes.append({"Test": "Khi² corrigé (Yates)", "Statistique": c2y, "ddl": "1", "p-value": py})
        _, pf = stats.fisher_exact(ct.values)
        lignes.append({"Test": "Test exact de Fisher", "Statistique": np.nan, "ddl": "—", "p-value": pf})
    return {
        "effectifs": ct, "attendus": attendus, "residus": residus,
        "pct_ligne": ct.div(ct.sum(axis=1), axis=0) * 100,
        "pct_colonne": ct.div(ct.sum(axis=0), axis=1) * 100,
        "tests": pd.DataFrame(lignes).set_index("Test"),
        "cramer_v": v, "pct_attendus_faibles": 100 * (attendus.values < 5).mean(),
        "p": p, "n": int(n), "table_2x2": ct.shape == (2, 2),
        "p_fisher": stats.fisher_exact(ct.values)[1] if ct.shape == (2, 2) else np.nan,
    }


def correlations(df, cols, methode="pearson"):
    fonction = {"pearson": stats.pearsonr, "spearman": stats.spearmanr, "kendall": stats.kendalltau}[methode]
    num = df[cols].apply(pd.to_numeric, errors="coerce")
    r = pd.DataFrame(np.nan, index=cols, columns=cols)
    p = r.copy()
    n = r.copy()
    paires = []
    for i, a in enumerate(cols):
        r.loc[a, a] = 1.0
        n.loc[a, a] = num[a].notna().sum()
        for b in cols[i + 1:]:
            d = num[[a, b]].dropna()
            if len(d) >= 3 and d[a].nunique() > 1 and d[b].nunique() > 1:
                res = fonction(d[a], d[b])
                rr, pp = float(res[0]), float(res[1])
            else:
                rr, pp = np.nan, np.nan
            r.loc[a, b] = r.loc[b, a] = rr
            p.loc[a, b] = p.loc[b, a] = pp
            n.loc[a, b] = n.loc[b, a] = len(d)
            paires.append({"Variable 1": a, "Variable 2": b, "r": rr, "p-value": pp, "N": len(d)})
    longue = pd.DataFrame(paires)
    if not longue.empty:
        longue["|r|"] = longue["r"].abs()
        longue = longue.sort_values("|r|", ascending=False).drop(columns="|r|")
        longue["Force"] = longue["r"].apply(
            lambda x: "—" if pd.isna(x) else force_effet(x, (0.1, 0.3, 0.5, 0.7), ("nulle", "faible", "modérée", "forte", "très forte")))
    return r, p, n, longue


# ---------------------------------------------------------------- régression
def regression_ols(df, y, xs, categoriques=()):
    def q(c):
        return f'Q("{c}")'
    termes = [f"C({q(c)})" if c in categoriques else q(c) for c in xs]
    formule = f"{q(y)} ~ " + " + ".join(termes)
    d = df[[y] + list(xs)].copy()
    d[y] = pd.to_numeric(d[y], errors="coerce")
    for c in xs:
        if c not in categoriques:
            d[c] = pd.to_numeric(d[c], errors="coerce")
        else:
            d[c] = d[c].astype(str)
    d = d.dropna()
    modele = smf.ols(formule, data=d).fit()
    ic = modele.conf_int()
    coef = pd.DataFrame({
        "Coefficient": modele.params, "Erreur std": modele.bse, "t": modele.tvalues,
        "p-value": modele.pvalues, "IC95% inf": ic[0], "IC95% sup": ic[1],
    })
    coef.index = [i.replace('Q("', "").replace('")', "").replace("C(", "").replace(")", "")
                  .replace("[T.", " = ").replace("]", "") for i in coef.index]
    resume = {
        "n": int(modele.nobs), "r2": modele.rsquared, "r2_aj": modele.rsquared_adj,
        "f": modele.fvalue, "p_f": modele.f_pvalue, "aic": modele.aic,
        "ddl": f"{int(modele.df_model)} ; {int(modele.df_resid)}",
    }
    diag = {}
    res = modele.resid
    if 3 <= len(res) <= 5000:
        diag["shapiro_p"] = stats.shapiro(res)[1]
    try:
        diag["bp_p"] = het_breuschpagan(res, modele.model.exog)[1]
    except Exception:
        diag["bp_p"] = np.nan
    return modele, coef, resume, diag


# ---------------------------------------------------------------- fiabilité
def cronbach(df, cols):
    d = df[cols].apply(pd.to_numeric, errors="coerce").dropna()
    k, n = len(cols), len(d)
    if k < 2 or n < 3:
        raise ValueError("Sélectionnez au moins 2 items et 3 répondants complets.")
    var_tot = d.sum(axis=1).var(ddof=1)
    alpha = k / (k - 1) * (1 - d.var(ddof=1).sum() / var_tot)
    lignes = []
    for c in cols:
        reste = d.drop(columns=c)
        k2 = k - 1
        if k2 > 1:
            a2 = k2 / (k2 - 1) * (1 - reste.var(ddof=1).sum() / reste.sum(axis=1).var(ddof=1))
        else:
            a2 = np.nan
        lignes.append({"Item": c, "Moyenne": d[c].mean(), "Écart-type": d[c].std(ddof=1),
                       "Corrélation item-total corrigée": d[c].corr(reste.sum(axis=1)),
                       "Alpha si item supprimé": a2})
    return alpha, n, pd.DataFrame(lignes).set_index("Item")


def lire_alpha(a):
    if a >= 0.9:
        return "excellente"
    if a >= 0.8:
        return "bonne"
    if a >= 0.7:
        return "acceptable"
    if a >= 0.6:
        return "discutable"
    return "faible"




# ---------------------------------------------------------------- recodage des variables
OBJECTIFS = {
    "Préparer une analyse descriptive": {
        "aide": "Regroupez ou transformez les modalités afin de faciliter la lecture des tableaux et graphiques.",
        "operations": ["modalites", "codes", "classes_egales", "quantiles", "bornes", "seuil", "inverser", "transformer", "indicatrices"],
    },
    "Préparer une régression": {
        "aide": "Créez des variables codées, des classes ou des indicatrices adaptées au modèle.",
        "operations": ["codes", "modalites", "classes_egales", "quantiles", "bornes", "seuil", "inverser", "transformer", "indicatrices"],
    },
    "Construire une échelle / un score": {
        "aide": "Harmonisez le sens des items et créez des catégories si nécessaire.",
        "operations": ["inverser", "seuil", "classes_egales", "quantiles", "bornes", "transformer", "codes", "modalites"],
    },
    "Créer des classes d'âge ou de revenu": {
        "aide": "Transformez une variable numérique en classes interprétables.",
        "operations": ["classes_egales", "quantiles", "bornes", "seuil"],
    },
}
OPERATIONS = {
    "classes_egales": "Créer des classes de même largeur",
    "quantiles": "Créer des classes par quantiles",
    "bornes": "Créer des classes avec bornes personnalisées",
    "seuil": "Séparer selon un seuil",
    "modalites": "Regrouper des modalités",
    "codes": "Coder les modalités en nombres",
    "inverser": "Inverser une échelle numérique",
    "transformer": "Transformer une variable numérique",
    "indicatrices": "Créer des variables indicatrices (0/1)",
}
OPERATIONS_NUMERIQUES = {"classes_egales", "quantiles", "bornes", "seuil", "inverser", "transformer"}
TRANSFORMATIONS = {
    "log": "Logarithme naturel (ln)",
    "log1p": "Logarithme ln(1+x)",
    "racine": "Racine carrée",
    "carre": "Carré",
    "standardiser": "Standardisation (z-score)",
    "normaliser": "Normaliser entre 0 et 1",
}

def colonnes_qualitatives(df):
    return [c for c in df.columns if not pd.api.types.is_numeric_dtype(df[c])
            and not pd.api.types.is_bool_dtype(df[c])
            and not pd.api.types.is_datetime64_any_dtype(df[c])]

def lire_nombres(texte):
    """Lit une liste de nombres séparés par des points-virgules."""
    if not str(texte).strip():
        return []
    try:
        vals = [float(x.strip().replace(",", ".")) for x in str(texte).split(";") if x.strip()]
    except ValueError as e:
        raise ValueError("Utilisez uniquement des nombres séparés par des points-virgules.") from e
    if any(not np.isfinite(x) for x in vals):
        raise ValueError("Les bornes doivent être des nombres finis.")
    if vals != sorted(vals) or len(set(vals)) != len(vals):
        raise ValueError("Les bornes doivent être strictement croissantes.")
    return vals

def nom_propose(df, source, operation, params=None):
    base = f"{source}_{operation}"
    base = "".join(ch if ch.isalnum() or ch == "_" else "_" for ch in base)
    nom = base
    i = 2
    while nom in df.columns:
        nom = f"{base}_{i}"
        i += 1
    return nom

def decrire_recodage(spec):
    operation = OPERATIONS.get(spec.get("operation", ""), spec.get("operation", ""))
    params = spec.get("params", {})
    if spec.get("operation") in ("classes_egales", "quantiles"):
        return f"{operation} (k={params.get('k', '?')})"
    if spec.get("operation") == "bornes":
        return f"{operation} : {params.get('bornes', [])}"
    if spec.get("operation") == "seuil":
        return f"{operation} (seuil={params.get('seuil', '?')})"
    if spec.get("operation") == "inverser":
        return f"{operation} ({params.get('min', '?')} à {params.get('max', '?')})"
    if spec.get("operation") == "transformer":
        return f"{operation} : {TRANSFORMATIONS.get(params.get('methode'), params.get('methode', ''))}"
    return operation

def appliquer_recodage(df, spec):
    """Applique un recodage défini par l'interface et renvoie ses nouvelles colonnes."""
    source = spec.get("source")
    operation = spec.get("operation")
    params = spec.get("params", {})
    nom = str(spec.get("nom", "")).strip()
    if source not in df.columns:
        raise ValueError(f"Variable source introuvable : {source}")
    if not nom:
        raise ValueError("Le nom de la nouvelle variable est vide.")
    s = df[source]
    if operation in ("classes_egales", "quantiles", "bornes", "seuil", "inverser", "transformer"):
        x = pd.to_numeric(s, errors="coerce")
        if x.notna().sum() == 0:
            raise ValueError("La variable ne contient aucune valeur numérique exploitable.")
        if operation == "classes_egales":
            k = int(params.get("k", 3))
            if x.nunique() < 2: raise ValueError("Il faut au moins deux valeurs distinctes.")
            edges = np.linspace(x.min(), x.max(), k + 1)
            edges = np.unique(edges)
            if len(edges) < 3: raise ValueError("Impossible de créer des classes avec cette variable.")
            out = pd.cut(x, bins=edges, include_lowest=True, duplicates="drop").astype("string")
        elif operation == "quantiles":
            k = int(params.get("k", 4))
            try:
                out = pd.qcut(x, q=k, duplicates="drop").astype("string")
            except ValueError as e:
                raise ValueError("Impossible de créer des classes par quantiles.") from e
        elif operation == "bornes":
            edges = params.get("bornes", [])
            if not edges: raise ValueError("Saisissez au moins une borne.")
            labels = params.get("etiquettes", [])
            bins = [-np.inf] + list(edges) + [np.inf]
            if labels and len(labels) == len(bins)-1:
                labs = labels
            else:
                labs = [f"< {edges[0]}"] + [f"[{edges[i-1]} ; {edges[i]}[" for i in range(1, len(edges))] + [f"≥ {edges[-1]}"]
            out = pd.cut(x, bins=bins, labels=labs, right=False, duplicates="drop").astype("string")
        elif operation == "seuil":
            seuil = float(params.get("seuil", x.median()))
            if params.get("en_01", False):
                out = pd.Series(np.where(x.isna(), np.nan, (x > seuil).astype(int)), index=x.index, name=nom)
            else:
                out = pd.Series(np.where(x.isna(), None, np.where(x <= seuil, f"≤ {seuil:g}", f"> {seuil:g}")), index=x.index)
        elif operation == "inverser":
            mn, mx = float(params.get("min", x.min())), float(params.get("max", x.max()))
            out = mn + mx - x
        else:
            methode = params.get("methode", "standardiser")
            if methode == "log":
                out = np.log(x.where(x > 0))
            elif methode == "log1p":
                out = np.log1p(x.where(x > -1))
            elif methode == "racine":
                out = np.sqrt(x.where(x >= 0))
            elif methode == "carre":
                out = x ** 2
            elif methode == "standardiser":
                sd = x.std(ddof=1)
                if not sd or pd.isna(sd): raise ValueError("Écart-type nul : standardisation impossible.")
                out = (x - x.mean()) / sd
            elif methode == "normaliser":
                den = x.max() - x.min()
                if not den: raise ValueError("Étendue nulle : normalisation impossible.")
                out = (x - x.min()) / den
            else:
                raise ValueError(f"Transformation inconnue : {methode}")
    elif operation in ("modalites", "codes"):
        txt = s.map(texte_modalite)
        corr = {str(k).strip(): v for k, v in params.get("correspondance", {}).items()}
        mapped = txt.map(lambda v: corr.get(v, v))
        if operation == "codes":
            out = pd.to_numeric(mapped, errors="coerce")
        else:
            out = mapped.where(s.notna(), np.nan)
    elif operation == "indicatrices":
        txt = s.map(texte_modalite)
        reference = str(params.get("reference", ""))
        mods = modalites(s)
        colonnes = {}
        for mod in mods:
            if mod == reference:
                continue
            suffix = "".join(ch if ch.isalnum() else "_" for ch in mod).strip("_") or "modalite"
            col = f"{nom}_{suffix}"
            colonnes[col] = (txt == mod).astype("Int64").where(s.notna(), pd.NA)
        if not colonnes:
            raise ValueError("Aucune modalité indicatrice à créer après choix de la référence.")
        return pd.DataFrame(colonnes, index=df.index)
    else:
        raise ValueError(f"Type de recodage inconnu : {operation}")
    return pd.DataFrame({nom: out}, index=df.index)

def appliquer_recodages(df, recodages):
    """Applique successivement les recodages sauvegardés ; retourne (données, erreurs)."""
    resultat = df.copy()
    erreurs = []
    for spec in recodages or []:
        try:
            nouvelles = appliquer_recodage(resultat, spec)
            collisions = [c for c in nouvelles.columns if c in resultat.columns]
            if collisions:
                raise ValueError(f"Le nom de variable existe déjà : {collisions[0]}")
            resultat = pd.concat([resultat, nouvelles], axis=1)
        except Exception as e:
            erreurs.append(f"{spec.get('nom', spec.get('source', 'recodage'))} : {e}")
    return resultat, erreurs


# ---------------------------------------------------------------- données d'exemple
def donnees_exemple(n=240, graine=7):
    rng = np.random.default_rng(graine)
    region = rng.choice(["Maritime", "Plateaux", "Kara"], n, p=[0.45, 0.30, 0.25])
    sexe = rng.choice(["Femme", "Homme"], n, p=[0.55, 0.45])
    age = np.clip(rng.normal(41, 11, n), 18, 75).round()
    effet_region = pd.Series(region).map({"Maritime": 0, "Plateaux": 4, "Kara": 9}).to_numpy()
    score = rng.normal(60, 12, n) + effet_region + 0.15 * (age - 40)
    revenu = np.round(rng.lognormal(10.9, 0.55, n) * (1 + 0.1 * effet_region / 9), 0)
    latent = rng.normal(0, 1, n) + 0.3 * effet_region / 9
    items = {f"Q{i}": np.clip(np.round(3 + 0.9 * latent + rng.normal(0, 0.8, n)), 1, 5) for i in range(1, 6)}
    formation = np.where(rng.random(n) < np.where(sexe == "Homme", 0.6, 0.4), "Oui", "Non")
    df = pd.DataFrame({"Region": region, "Sexe": sexe, "Age": age, "Revenu": revenu,
                       "Score": score.round(1), "Formation": formation, **items})
    df.loc[rng.choice(n, 8, replace=False), "Revenu"] = np.nan
    return df
