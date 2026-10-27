import streamlit as st
import pandas as pd
import numpy as np
import stats_utils as su
import graphiques as gr
import rapport as rp
import ui_exploration as ue

st.set_page_config(
    page_title="StatExcel", page_icon="📊", layout="wide",
    menu_items={"About": f"StatExcel — analyse statistique. Auteur : {su.AUTEUR}"},
)
st.title("📊 StatExcel — Analyse statistique")
st.caption("Importez un classeur Excel, choisissez les variables et lancez les analyses.")
st.caption(f"Auteur : **{su.AUTEUR}**")

with st.sidebar:
    st.header("Données")
    fichier = st.file_uploader("Importer un fichier Excel", type=["xlsx", "xls"])
    alpha = st.selectbox("Seuil de significativité α", [0.01, 0.05, 0.10], index=1)
    st.divider()
    st.caption(f"StatExcel · Auteur : {su.AUTEUR}")

if fichier is None:
    st.info("Commencez par importer un fichier Excel dans le panneau de gauche.")
    st.markdown("""
    **Analyses disponibles dans cette première interface :**
    - Exploration automatique : profil socio-démographique, hypothèses testables, carte de la zone d'étude
    - Recodage des variables selon l'analyse visée
    - Statistiques descriptives, fréquences et graphiques
    - Tests de normalité
    - Comparaison de deux groupes ou de plusieurs groupes
    - Test du khi-deux
    - Corrélations
    - Régression linéaire (OLS)
    - Alpha de Cronbach
    """)
    st.stop()

try:
    xls = pd.ExcelFile(fichier)
    feuille = st.selectbox("Feuille du classeur", xls.sheet_names)
    df_brut = pd.read_excel(fichier, sheet_name=feuille)
    df_base = su.nettoyer(df_brut)
except Exception as e:
    st.error(f"Impossible de lire ce fichier : {e}")
    st.stop()

# Recodages : conservés pendant la session, réinitialisés quand on change de fichier ou de feuille
cle_fichier = f"{fichier.name}|{feuille}"
if st.session_state.get("cle_fichier") != cle_fichier:
    st.session_state["cle_fichier"] = cle_fichier
    st.session_state["recodages"] = []
    st.session_state["rapport_hypotheses"] = []
    st.session_state["rapport_blocs"] = None
df, erreurs_recodage = su.appliquer_recodages(df_base, st.session_state["recodages"])

if df.empty:
    st.error("La feuille sélectionnée ne contient pas de données exploitables.")
    st.stop()

st.success(f"Fichier chargé : {df.shape[0]} observations et {df.shape[1]} variables.")
c1, c2, c3 = st.columns(3)
c1.metric("Observations", df.shape[0])
c2.metric("Variables", df.shape[1])
c3.metric("Valeurs manquantes", int(df.isna().sum().sum()))

ctx_exploration = ue.analyser(df)
ue.bandeau(ctx_exploration)

tab_guide, tab_auto, tab_carte, tab_explo, tab_data, tab_recode, tab_desc, tab_tests, tab_assoc, tab_reg, tab_alpha, tab_rapport = st.tabs([
    "Guide débutant", "Assistant automatique", "Carte GPS", "Exploration auto", "Données", "Recodage", "Descriptives",
    "Tests", "Associations", "Régression", "Fiabilité", "Rapport automatique"
])

with tab_guide:
    ue.onglet_debutant(df, alpha)

with tab_auto:
    ue.onglet_assistant_auto(df, alpha, ctx_exploration)

with tab_carte:
    ue.onglet_carte(df, ctx_exploration)

with tab_explo:
    ue.onglet_exploration(df, alpha, ctx_exploration)

with tab_data:
    st.subheader("Aperçu des données nettoyées (recodages inclus)")
    st.dataframe(df, width="stretch")
    st.subheader("Résumé des variables")
    st.dataframe(su.resume_variables(df), width="stretch")
    st.download_button(
        "Télécharger les données nettoyées (CSV)",
        df.to_csv(index=False).encode("utf-8-sig"),
        file_name="donnees_nettoyees.csv",
        mime="text/csv"
    )

with tab_recode:
    st.subheader("Recodage des variables")
    st.caption("Chaque recodage crée une nouvelle variable : la colonne d'origine n'est jamais modifiée. "
               "Les variables créées sont ensuite proposées dans tous les onglets.")
    objectif = st.selectbox("Analyse visée", list(su.OBJECTIFS), key="rec_obj")
    info = su.OBJECTIFS[objectif]
    st.info(info["aide"])
    operation = st.selectbox("Type de recodage", info["operations"], format_func=su.OPERATIONS.get, key="rec_op")
    sources = su.colonnes_numeriques(df) if operation in su.OPERATIONS_NUMERIQUES else su.colonnes_groupes(df, max_modalites=50)

    if not sources:
        st.warning("Aucune variable compatible avec ce type de recodage.")
    else:
        source = st.selectbox("Variable à recoder", sources, key="rec_src")
        s = df[source]
        cle = f"{operation}_{source}"
        params = {}
        if operation in su.OPERATIONS_NUMERIQUES:
            x = pd.to_numeric(s, errors="coerce")
            st.caption(f"{x.notna().sum()} valeurs · min {x.min():.4g} · médiane {x.median():.4g} · max {x.max():.4g}")

        if operation in ("classes_egales", "quantiles"):
            params["k"] = int(st.number_input("Nombre de classes", 2, 12, 3 if operation == "classes_egales" else 4,
                                              key=f"k_{cle}"))
        elif operation == "bornes":
            texte = st.text_input("Bornes (séparées par des points-virgules)", placeholder="ex. 30 ; 45 ; 60",
                                  key=f"b_{cle}", help="Classes [a ; b[ : la borne basse est incluse, la haute exclue.")
            etiq = st.text_input("Étiquettes des classes (facultatif, séparées par des points-virgules)",
                                 key=f"e_{cle}", help="Une de plus que le nombre de bornes, dans l'ordre croissant.")
            try:
                params["bornes"] = su.lire_nombres(texte)
            except ValueError as e:
                params["bornes"] = []
                st.warning(str(e))
            params["etiquettes"] = [e.strip() for e in etiq.split(";") if e.strip()]
        elif operation == "seuil":
            params["seuil"] = st.number_input("Seuil (≤ seuil / > seuil)", value=float(x.median()), key=f"s_{cle}")
            params["en_01"] = st.checkbox("Coder en 0 / 1 (valeur numérique) plutôt qu'en classes ≤ / >", key=f"o1_{cle}")
        elif operation in ("modalites", "codes"):
            mods = su.modalites(s)
            defaut = [str(i + 1) for i in range(len(mods))] if operation == "codes" else mods
            if operation == "codes":
                st.caption("Les codes proposés suivent l'ordre alphabétique : corrigez-les si l'échelle a un autre ordre "
                           "(ex. Pas d'accord = 1 … D'accord = 5).")
            else:
                st.caption("Donnez la même valeur à plusieurs modalités pour les regrouper ; laissez vide pour ne rien changer.")
            ed = st.data_editor(pd.DataFrame({"Modalité": mods, "Nouvelle valeur": defaut}), hide_index=True,
                                disabled=["Modalité"], key=f"ed_{cle}", width="stretch")
            params["correspondance"] = dict(zip(ed["Modalité"], ed["Nouvelle valeur"]))
            params["ordre"] = list(ed["Modalité"])
        elif operation == "inverser":
            c1, c2 = st.columns(2)
            params["min"] = c1.number_input("Minimum de l'échelle", value=float(x.min()), key=f"mn_{cle}")
            params["max"] = c2.number_input("Maximum de l'échelle", value=float(x.max()), key=f"mx_{cle}")
            st.caption("Nouvelle valeur = minimum + maximum − valeur (sur une échelle de 1 à 5 : 1↔5, 2↔4).")
        elif operation == "transformer":
            params["methode"] = st.selectbox("Transformation", list(su.TRANSFORMATIONS),
                                             format_func=su.TRANSFORMATIONS.get, key=f"t_{cle}")
        elif operation == "indicatrices":
            mods = su.modalites(s)
            frequente = s.map(su.texte_modalite).value_counts().idxmax()
            params["reference"] = st.selectbox("Modalité de référence", mods, index=mods.index(frequente),
                                               key=f"ref_{cle}",
                                               help="Les coefficients de la régression se lisent par rapport à cette modalité.")

        nom = st.text_input("Nom de la nouvelle variable", value=su.nom_propose(df, source, operation, params),
                            key=f"nom_{cle}").strip()
        spec = {"nom": nom, "source": source, "operation": operation, "params": params}
        try:
            if not nom:
                raise ValueError("Donnez un nom à la variable.")
            nouveau = su.appliquer_recodage(df, spec)
            pris = [c for c in nouveau.columns if c in df.columns]
            if pris:
                raise ValueError(f"Le nom « {pris[0]} » existe déjà : choisissez-en un autre.")
            st.markdown("**Aperçu**")
            st.dataframe(pd.concat([df[[source]].rename(columns={source: f"{source} (avant)"}), nouveau], axis=1).head(10),
                         width="stretch")
            if all(pd.api.types.is_numeric_dtype(nouveau[c]) for c in nouveau.columns):
                st.dataframe(su.decrire_numeriques(nouveau, list(nouveau.columns)), width="stretch")
            else:
                st.dataframe(su.frequences(nouveau.iloc[:, 0]).round(1), width="stretch")
            apercu_ok = True
        except Exception as e:
            st.warning(f"Aperçu impossible : {e}")
            apercu_ok = False
        if st.button("Ajouter cette variable", type="primary", disabled=not apercu_ok):
            st.session_state["recodages"].append(spec)
            st.session_state.pop(f"nom_{cle}", None)
            st.rerun()

    st.divider()
    st.subheader("Recodages enregistrés")
    for msg in erreurs_recodage:
        st.warning(f"Recodage ignoré — {msg}")
    if st.session_state["recodages"]:
        for i, sp in enumerate(st.session_state["recodages"]):
            ca, cb = st.columns([6, 1])
            ca.write(f"**{sp['nom']}** ← {sp['source']} · {su.decrire_recodage(sp)}")
            if cb.button("Supprimer", key=f"sup_{i}"):
                st.session_state["recodages"].pop(i)
                st.rerun()
        if st.button("Tout supprimer", key="sup_tout"):
            st.session_state["recodages"] = []
            st.rerun()
    else:
        st.caption("Aucun recodage pour l'instant.")

with tab_desc:
    nums = su.colonnes_numeriques(df)
    groupes_candidats = su.colonnes_groupes(df)
    t_quant, t_qual, t_rel = st.tabs(["Variables quantitatives", "Variables qualitatives", "Relations entre variables"])

    with t_quant:
        if nums:
            choix_nums = st.multiselect("Variables quantitatives", nums, default=nums)
            if choix_nums:
                desc = su.decrire_numeriques(df, choix_nums)
                st.dataframe(desc, width="stretch")
                st.download_button(
                    "Télécharger les descriptives (CSV)",
                    desc.to_csv().encode("utf-8-sig"),
                    file_name="statistiques_descriptives.csv",
                    mime="text/csv"
                )
                if groupes_candidats:
                    g = st.selectbox("Optionnel : décrire par groupe", ["— Aucun —"] + groupes_candidats)
                    if g != "— Aucun —":
                        st.dataframe(su.decrire_par_groupe(df, choix_nums, g), width="stretch")

            st.divider()
            st.subheader("Graphiques")
            vg = st.selectbox("Variable à représenter", nums, key="g_var")
            type_g = st.radio("Type de graphique", ["Histogramme", "Boîte à moustaches", "Diagramme Q-Q"],
                              horizontal=True, key="g_type")
            try:
                if type_g == "Histogramme":
                    n_val = int(df[vg].notna().sum())
                    defaut = int(np.clip(np.ceil(np.log2(max(n_val, 2)) + 1), 8, 30))
                    c1, c2 = st.columns(2)
                    classes = c1.slider("Nombre de classes", 5, 60, defaut, key="g_classes")
                    courbe = c2.selectbox("Courbe superposée", ["Aucune", "Densité (noyau)", "Loi normale ajustée"],
                                          key="g_courbe")
                    st.altair_chart(gr.histogramme(df[vg], vg, classes, courbe), width="stretch")
                elif type_g == "Boîte à moustaches":
                    gb = st.selectbox("Séparer par groupe (facultatif)",
                                      ["— Aucun —"] + [c for c in groupes_candidats if c != vg], key="g_box_grp")
                    st.altair_chart(gr.boite(df, vg, None if gb == "— Aucun —" else gb), width="stretch")
                    st.caption("Losange noir = moyenne ; les points isolés sont au-delà de 1,5 fois l'écart interquartile.")
                else:
                    st.altair_chart(gr.qq(df[vg], vg), width="stretch")
                    st.caption("Plus les points suivent la droite rouge, plus la distribution est proche de la normale.")
            except ValueError as e:
                st.warning(f"Graphique impossible : {e}")
            st.caption("Le menu ⋯ en haut à droite de chaque graphique permet de l'exporter en PNG ou SVG.")
        else:
            st.warning("Aucune variable numérique détectée.")

    with t_qual:
        quals = su.colonnes_qualitatives(df)
        if quals:
            vq = st.selectbox("Variable qualitative", quals, key="q_var")
            freq = su.frequences(df[vq])
            st.dataframe(freq.round(2), width="stretch")
            st.download_button(
                "Télécharger les fréquences (CSV)",
                freq.to_csv().encode("utf-8-sig"),
                file_name=f"frequences_{vq}.csv",
                mime="text/csv"
            )
            type_q = st.radio("Type de graphique", ["Barres", "Secteurs"], horizontal=True, key="q_type")
            try:
                st.altair_chart(gr.frequences_graphique(freq, vq, type_q), width="stretch")
                if len(freq.drop(index="Manquant", errors="ignore")) > 20:
                    st.caption("Seules les 20 modalités les plus fréquentes sont représentées.")
            except ValueError as e:
                st.warning(f"Graphique impossible : {e}")
        else:
            st.info("Aucune variable qualitative détectée.")

    with t_rel:
        if len(nums) >= 2:
            st.subheader("Nuage de points")
            c1, c2, c3 = st.columns(3)
            nx = c1.selectbox("Axe X", nums, key="n_x")
            ny = c2.selectbox("Axe Y", [c for c in nums if c != nx], key="n_y")
            ng = c3.selectbox("Couleur par groupe (facultatif)",
                              ["— Aucun —"] + [c for c in groupes_candidats if c not in (nx, ny)], key="n_g")
            droite = st.checkbox("Afficher la droite de régression", value=True, key="n_droite")
            try:
                st.altair_chart(gr.nuage(df, nx, ny, None if ng == "— Aucun —" else ng, droite), width="stretch")
                _, _, _, lg = su.correlations(df, [nx, ny])
                if not lg.empty and pd.notna(lg.iloc[0]["r"]):
                    st.caption(f"Corrélation de Pearson : r = {lg.iloc[0]['r']:.3f} "
                               f"({su.texte_p(lg.iloc[0]['p-value'])}, N = {int(lg.iloc[0]['N'])}).")
            except ValueError as e:
                st.warning(f"Graphique impossible : {e}")

            st.divider()
            st.subheader("Matrice de corrélation")
            cols_h = st.multiselect("Variables", nums, default=nums[:min(8, len(nums))], key="h_cols")
            meth_h = st.selectbox("Méthode", ["pearson", "spearman", "kendall"], key="h_meth")
            if len(cols_h) >= 2:
                rmat, _, _, _ = su.correlations(df, cols_h, methode=meth_h)
                st.altair_chart(gr.carte_correlations(rmat))
        else:
            st.info("Il faut au moins deux variables numériques.")

with tab_tests:
    st.subheader("Tests de normalité")
    nums = su.colonnes_numeriques(df)
    if nums:
        v = st.selectbox("Variable à tester", nums, key="normalite_var")
        res_norm = su.tests_normalite(df[v], alpha=alpha)
        if res_norm.empty:
            st.warning("Effectif insuffisant ou variable constante : test non calculable.")
        else:
            st.dataframe(res_norm, width="stretch")
    else:
        st.warning("Aucune variable numérique détectée.")

    st.divider()
    st.subheader("Comparer deux groupes ou plus")
    groupes_candidats = su.colonnes_groupes(df)
    if nums and groupes_candidats:
        y = st.selectbox("Variable quantitative", nums, key="compar_y")
        g_options = [c for c in groupes_candidats if c != y]
        if not g_options:
            st.info("Aucune variable de regroupement différente de la variable quantitative.")
            st.stop()
        g = st.selectbox("Variable de regroupement", g_options, key="compar_g")
        groupes = su.groupes_numeriques(df, y, g)
        groupes = {k: v for k, v in groupes.items() if len(v) > 0}
        if len(groupes) >= 2:
            st.write("Descriptifs par groupe")
            st.dataframe(su.descriptifs_groupes(groupes), width="stretch")
            if len(groupes) == 2:
                noms = list(groupes)
                try:
                    r = su.comparer_deux(groupes[noms[0]], groupes[noms[1]], alpha=alpha)
                    st.write(f"**Test recommandé : {r['reco']}** — {r['motif']}")
                    st.dataframe(r["tests"], width="stretch")
                    st.write(f"Différence des moyennes : {r['diff']:.4g}")
                    st.write(f"IC à {int((1-alpha)*100)} % : [{r['ic'][0]:.4g} ; {r['ic'][1]:.4g}]")
                    st.write(su.conclusion(r["p_reco"], alpha, f"La comparaison selon {y}"))
                except Exception as e:
                    st.error(f"Comparaison impossible : {e}")
            else:
                try:
                    r = su.comparer_k(groupes, alpha=alpha)
                    st.write(f"**Test recommandé : {r['reco']}** — {r['motif']}")
                    st.dataframe(r["tests"], width="stretch")
                    st.write(su.conclusion(r["p_reco"], alpha, f"La différence de {y} entre les groupes"))
                    if r.get("posthoc") == "Tukey HSD" and r["p_reco"] < alpha:
                        st.subheader("Comparaisons post-hoc de Tukey")
                        st.dataframe(su.posthoc_tukey(groupes, alpha=alpha), width="stretch")
                    elif r["p_reco"] < alpha:
                        st.subheader("Comparaisons post-hoc par paires (correction de Holm)")
                        st.dataframe(su.posthoc_paires(groupes, test="welch" if r["reco"] == "ANOVA de Welch" else "mannwhitney", methode="holm", alpha=alpha), width="stretch")
                except Exception as e:
                    st.error(f"Comparaison impossible : {e}")
        else:
            st.info("Il faut au moins deux groupes avec des observations.")
    else:
        st.info("Il faut une variable numérique et une variable de regroupement.")

with tab_assoc:
    st.subheader("Association entre deux variables qualitatives (khi-deux)")
    candidats = su.colonnes_groupes(df)
    if len(candidats) >= 2:
        a = st.selectbox("Première variable qualitative", candidats, key="chi_a")
        b = st.selectbox("Deuxième variable qualitative", [x for x in candidats if x != a], key="chi_b")
        if st.button("Exécuter le test du khi-deux"):
            try:
                r = su.tester_khi2(df, a, b)
                st.write("Tableau des effectifs")
                st.dataframe(r["effectifs"], width="stretch")
                st.write("Pourcentages par ligne (%)")
                st.dataframe(r["pct_ligne"].round(2), width="stretch")
                st.write("Effectifs théoriques")
                st.dataframe(r["attendus"].round(2), width="stretch")
                st.dataframe(r["tests"], width="stretch")
                st.metric("V de Cramér", f"{r['cramer_v']:.3f}")
                st.write(f"Cellules avec effectif théorique inférieur à 5 : {r['pct_attendus_faibles']:.1f} %")
                st.write(su.conclusion(r["p"], alpha, f"L'association entre {a} et {b}"))
            except Exception as e:
                st.error(f"Test impossible : {e}")
    else:
        st.info("Il faut au moins deux variables avec plusieurs modalités.")

    st.divider()
    st.subheader("Corrélations")
    nums = su.colonnes_numeriques(df)
    if len(nums) >= 2:
        cols_corr = st.multiselect("Variables à corréler", nums, default=nums[:min(5, len(nums))])
        methode = st.selectbox("Méthode", ["pearson", "spearman", "kendall"])
        if len(cols_corr) >= 2:
            if st.button("Calculer les corrélations"):
                rmat, pmat, nmat, longue = su.correlations(df, cols_corr, methode=methode)
                st.write("Coefficients de corrélation")
                st.dataframe(rmat.round(3), width="stretch")
                st.write("p-values")
                st.dataframe(pmat.round(4), width="stretch")
                st.write("Paires classées par force de corrélation")
                st.dataframe(longue, width="stretch")
    else:
        st.info("Il faut au moins deux variables numériques.")

with tab_reg:
    st.subheader("Régression linéaire (OLS)")
    nums = su.colonnes_numeriques(df)
    if len(nums) >= 2:
        y = st.selectbox("Variable dépendante (Y)", nums, key="ols_y")
        xs = st.multiselect("Variables explicatives (X)", [c for c in df.columns if c != y], default=[c for c in nums if c != y][:2])
        cat = st.multiselect("Parmi les X, variables à traiter comme qualitatives", [c for c in xs if not pd.api.types.is_numeric_dtype(df[c])])
        if st.button("Estimer la régression"):
            if not xs:
                st.warning("Sélectionnez au moins une variable explicative.")
            else:
                try:
                    modele, coef, resume, diag = su.regression_ols(df, y, xs, categoriques=cat)
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("N utilisé", resume["n"])
                    m2.metric("R²", f"{resume['r2']:.3f}")
                    m3.metric("R² ajusté", f"{resume['r2_aj']:.3f}")
                    m4.metric("p-value globale", su.formater_p(resume["p_f"]))
                    st.write("Coefficients et intervalles de confiance")
                    st.dataframe(coef, width="stretch")
                    st.write(f"Test global : {su.conclusion(resume['p_f'], alpha, 'Le modèle')}")
                    st.write("Diagnostics")
                    st.json({k: (None if pd.isna(v) else float(v)) for k, v in diag.items()})
                    with st.expander("Résumé statistique complet"):
                        st.text(modele.summary().as_text())
                except Exception as e:
                    st.error(f"Régression impossible : {e}")
    else:
        st.info("Il faut au moins deux variables numériques pour commencer.")

with tab_alpha:
    st.subheader("Fiabilité interne — alpha de Cronbach")
    nums = su.colonnes_numeriques(df)
    cols_alpha = st.multiselect("Sélectionnez les items du questionnaire", nums)
    if st.button("Calculer l'alpha de Cronbach"):
        try:
            alpha_c, n, detail = su.cronbach(df, cols_alpha)
            st.metric("Alpha de Cronbach", f"{alpha_c:.3f}")
            st.write(f"Répondants complets utilisés : {n}")
            st.write(f"Interprétation indicative : fiabilité {su.lire_alpha(alpha_c)}")
            st.dataframe(detail, width="stretch")
        except Exception as e:
            st.error(f"Calcul impossible : {e}")


with tab_rapport:
    st.subheader("Rapport statistique automatique")
    st.write("Définissez vos hypothèses de recherche, sélectionnez les analyses à inclure, puis générez un rapport interprété et téléchargeable en Word.")

    nums_r = su.colonnes_numeriques(df)
    qual_r = su.colonnes_groupes(df, max_modalites=50)
    all_cols = list(df.columns)
    if "rapport_hypotheses" not in st.session_state:
        st.session_state["rapport_hypotheses"] = []
    if "rapport_blocs" not in st.session_state:
        st.session_state["rapport_blocs"] = None
    if "rapport_meta" not in st.session_state:
        st.session_state["rapport_meta"] = {}

    with st.expander("1. Informations du rapport", expanded=True):
        titre_r = st.text_input("Titre du rapport", value="Rapport d'analyse statistique", key="rp_titre")
        org_r = st.text_input("Organisation / établissement (facultatif)", key="rp_org")
        contexte_r = st.text_area("Contexte de l'étude (facultatif)", key="rp_contexte",
                                  placeholder="Présentez brièvement le sujet et le contexte de l'étude.")
        objectifs_r = st.text_area("Objectifs de l'étude (facultatif)", key="rp_objectifs",
                                  placeholder="Indiquez l'objectif général et les objectifs spécifiques.")

    with st.expander("2. Ajouter une hypothèse statistique", expanded=True):
        type_labels = {
            "groupes": "Comparer une variable quantitative entre groupes",
            "khi2": "Association entre deux variables qualitatives (khi-deux)",
            "correlation": "Corrélation entre deux variables quantitatives",
            "apparie": "Comparer deux mesures appariées",
            "normalite": "Tester la normalité d'une variable",
            "regression": "Régression linéaire multiple",
            "fiabilite": "Fiabilité d'une échelle (alpha de Cronbach)"
        }
        type_r = st.selectbox("Type d'analyse", list(type_labels), format_func=lambda x: type_labels[x], key="rp_type")
        params_r = {}
        if type_r == "groupes":
            if nums_r and qual_r:
                params_r["y"] = st.selectbox("Variable quantitative à comparer", nums_r, key="rp_grp_y")
                params_r["g"] = st.selectbox("Variable de groupe", qual_r, key="rp_grp_g")
                nbgr = df[params_r["g"]].nunique(dropna=True)
            else:
                st.warning("Cette analyse nécessite au moins une variable quantitative et une variable de groupe.")
                nbgr = 2
        elif type_r == "khi2":
            if len(qual_r) >= 2:
                params_r["a"] = st.selectbox("Première variable qualitative", qual_r, key="rp_chi_a")
                params_r["b"] = st.selectbox("Deuxième variable qualitative", [c for c in qual_r if c != params_r["a"]], key="rp_chi_b")
            else:
                st.warning("Il faut au moins deux variables qualitatives.")
        elif type_r == "correlation":
            if len(nums_r) >= 2:
                params_r["x"] = st.selectbox("Variable X", nums_r, key="rp_corr_x")
                params_r["y"] = st.selectbox("Variable Y", [c for c in nums_r if c != params_r["x"]], key="rp_corr_y")
                params_r["methode"] = st.selectbox("Coefficient", ["pearson", "spearman", "kendall"], key="rp_corr_meth")
            else:
                st.warning("Il faut au moins deux variables quantitatives.")
        elif type_r == "apparie":
            if len(nums_r) >= 2:
                params_r["a"] = st.selectbox("Première mesure", nums_r, key="rp_app_a")
                params_r["b"] = st.selectbox("Deuxième mesure", [c for c in nums_r if c != params_r["a"]], key="rp_app_b")
            else:
                st.warning("Il faut au moins deux variables quantitatives.")
        elif type_r == "normalite":
            if nums_r:
                params_r["v"] = st.selectbox("Variable à tester", nums_r, key="rp_norm_v")
            else:
                st.warning("Il faut une variable quantitative.")
        elif type_r == "regression":
            if nums_r:
                params_r["y"] = st.selectbox("Variable dépendante (Y)", nums_r, key="rp_reg_y")
                params_r["xs"] = st.multiselect("Variables explicatives (X)", [c for c in all_cols if c != params_r["y"]], key="rp_reg_xs")
            else:
                st.warning("Il faut une variable dépendante quantitative.")
        elif type_r == "fiabilite":
            params_r["items"] = st.multiselect("Items de l'échelle", nums_r, key="rp_alpha_items")
        enonce_r = st.text_input("Énoncé de l'hypothèse de recherche (facultatif)", key="rp_enonce")
        h0_default, h1_default = rp.hypotheses_par_defaut(type_r, params_r, int(nbgr) if type_r == "groupes" else 2) if params_r else ("", "")
        h0_r = st.text_area("Hypothèse nulle H0", value=h0_default, key="rp_h0")
        h1_r = st.text_area("Hypothèse alternative H1", value=h1_default, key="rp_h1")
        valide_params = (
            (type_r == "groupes" and "y" in params_r and "g" in params_r) or
            (type_r == "khi2" and "a" in params_r and "b" in params_r) or
            (type_r == "correlation" and "x" in params_r and "y" in params_r) or
            (type_r == "apparie" and "a" in params_r and "b" in params_r) or
            (type_r == "normalite" and "v" in params_r) or
            (type_r == "regression" and "y" in params_r and len(params_r.get("xs", [])) > 0) or
            (type_r == "fiabilite" and len(params_r.get("items", [])) >= 2)
        )
        if st.button("Ajouter cette hypothèse", type="primary", disabled=not valide_params, key="rp_add"):
            st.session_state["rapport_hypotheses"].append({
                "type": type_r, "params": params_r.copy(), "h0": h0_r, "h1": h1_r, "enonce": enonce_r
            })
            st.session_state["rapport_blocs"] = None
            st.rerun()

    st.markdown("**Hypothèses enregistrées**")
    if st.session_state["rapport_hypotheses"]:
        for i, hyp in enumerate(st.session_state["rapport_hypotheses"]):
            ca, cb = st.columns([8, 1])
            ca.write(f"**H{i+1} — {rp.titre_hypothese(hyp)}**")
            ca.caption(f"H0 : {hyp['h0']}  |  H1 : {hyp['h1']}")
            if cb.button("Supprimer", key=f"rp_remove_{i}"):
                st.session_state["rapport_hypotheses"].pop(i)
                st.session_state["rapport_blocs"] = None
                st.rerun()
        if st.button("Effacer toutes les hypothèses", key="rp_clear"):
            st.session_state["rapport_hypotheses"] = []
            st.session_state["rapport_blocs"] = None
            st.rerun()
    else:
        st.info("Aucune hypothèse ajoutée. Vous pouvez tout de même générer un rapport descriptif.")

    with st.expander("3. Contenu du rapport", expanded=True):
        descriptives_r = st.checkbox("Inclure les statistiques descriptives", value=True, key="rp_desc")
        nums_sel_r = st.multiselect("Variables quantitatives à décrire", nums_r, default=nums_r[:min(10, len(nums_r))], key="rp_nums")
        qual_sel_r = st.multiselect("Variables qualitatives à décrire", qual_r, default=qual_r[:min(8, len(qual_r))], key="rp_quals")
        figures_r = st.checkbox("Inclure les graphiques", value=True, key="rp_figures")
        recodages_r = st.checkbox("Décrire les recodages appliqués", value=True, key="rp_recodages")

    if st.button("Générer le rapport", type="primary", key="rp_generate"):
        meta_r = {
            "titre": titre_r.strip() or "Rapport d'analyse statistique",
            "auteur": "BOKOUMBO Tibetobe Ambroise",
            "organisation": org_r.strip(),
            "contexte": contexte_r.strip(),
            "objectifs": objectifs_r.strip(),
            "fichier": fichier.name,
            "feuille": feuille
        }
        options_r = {
            "graphiques": figures_r,
            "descriptives": descriptives_r,
            "cols_num": nums_sel_r,
            "cols_qual": qual_sel_r,
            "inclure_recodages": recodages_r
        }
        with st.spinner("Calcul des analyses et rédaction du rapport…"):
            blocs_r = rp.construire_rapport(df, meta_r, st.session_state["rapport_hypotheses"],
                                            options_r, alpha, st.session_state.get("recodages", []))
        st.session_state["rapport_blocs"] = blocs_r
        st.session_state["rapport_meta"] = meta_r
        st.success("Rapport généré. Vérifiez l'aperçu ci-dessous avant le téléchargement.")

    if st.session_state.get("rapport_blocs") is not None:
        blocs_r = st.session_state["rapport_blocs"]
        meta_r = st.session_state["rapport_meta"]
        st.divider()
        st.subheader("Aperçu du rapport")
        st.caption(f"{meta_r.get('titre')} · {df.shape[0]} observations · Auteur : {meta_r.get('auteur')}")
        for bloc in blocs_r:
            typ = bloc["type"]
            if typ in ("h1", "h2", "h3"):
                st.markdown("#" * int(typ[1]) + " " + bloc["texte"])
            elif typ == "p":
                if bloc.get("gras"):
                    st.markdown(f"**{bloc['gras']}**{bloc['texte']}")
                else:
                    st.write(bloc["texte"])
            elif typ == "liste":
                for item in bloc["items"]:
                    st.markdown(f"- {item}")
            elif typ == "table":
                st.caption(bloc["legende"])
                st.dataframe(bloc["df"], width="stretch")
            elif typ == "fig":
                st.image(bloc["png"], caption=bloc["legende"], width="stretch")
        try:
            docx_bytes = rp.exporter_docx(blocs_r, meta_r)
            st.download_button(
                "Télécharger le rapport Word (.docx)",
                data=docx_bytes,
                file_name="rapport_statistique.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                key="rp_download"
            )
        except Exception as e:
            st.error(f"Impossible de préparer le fichier Word : {e}")
