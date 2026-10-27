# -*- coding: utf-8 -*-
"""
Interface Streamlit de l'exploration automatique (profil socio-démographique,
hypothèses suggérées, carte de la zone d'étude).

Auteur : BOKOUMBO Tibetobe Ambroise
"""
import html

import numpy as np
import pandas as pd
import streamlit as st

import exploration as ex
import rapport as rp
import stats_utils as su


def analyser(df):
    """Détections réalisées une fois par exécution : coordonnées et variables socio-démographiques."""
    coords = ex.detecter_coordonnees(df)
    cols_coord = [c for c in (coords or ()) if c]
    profil = ex.detecter_sociodemographiques(df, exclure=cols_coord)
    return {"coords": coords, "cols_coord": cols_coord, "profil": profil}


def bandeau(ctx):
    """Résumé affiché dès le chargement de la base."""
    profil, coords = ctx["profil"], ctx["coords"]
    if profil:
        st.info("**Variables socio-démographiques reconnues :** "
                + " · ".join(f"{cat} ({', '.join(cols)})" for cat, cols in profil.items())
                + "  \nGraphiques, hypothèses testables et carte : onglet **Exploration auto**.")
    else:
        st.info("Aucune variable socio-démographique reconnue automatiquement par son nom : "
                "l'onglet **Exploration auto** vous permet de choisir les variables à représenter.")
    if coords:
        cols = " et ".join(f"« {c} »" for c in coords if c)
        st.success(f"Coordonnées géographiques détectées ({cols}) : la carte de la zone d'étude est disponible "
                   "dans l'onglet **Exploration auto**.")


def _md(texte):
    for c in ("\\", "*", "_", "$", "~", "`"):
        texte = texte.replace(c, "\\" + c)
    return texte


def afficher_blocs(blocs):
    """Affiche les blocs d'un rapport (titres, paragraphes, tableaux, figures)."""
    n_tab = n_fig = 0
    for b in blocs:
        t = b["type"]
        if t in ("h1", "h2", "h3"):
            st.markdown("#" * (int(t[1]) + 2) + " " + _md(b["texte"]))
        elif t == "p":
            st.markdown((f"**{_md(b['gras'])}**" if b.get("gras") else "") + _md(b["texte"]))
        elif t == "liste":
            st.markdown("\n".join(f"- {_md(i)}" for i in b["items"]))
        elif t == "table":
            n_tab += 1
            st.markdown(f"**Tableau {n_tab}** : *{_md(b['legende'])}*")
            st.dataframe(rp.formater_table(b["df"]), hide_index=True, width="stretch")
        elif t == "fig":
            n_fig += 1
            st.image(b["png"], caption=f"Figure {n_fig} : {b['legende']}")


def ajouter_au_rapport(numeros, specs):
    """Ajoute les hypothèses suggérées (numéros 1, 2, … du tableau) au rapport ; retourne le nombre ajouté."""
    hyps = st.session_state.setdefault("rapport_hypotheses", [])
    existantes = {(h["type"], repr(sorted(h["params"].items()))) for h in hyps}
    ajoutees = 0
    for i in numeros:
        sp = specs[i - 1]
        cle = (sp["type"], repr(sorted(sp["params"].items())))
        if cle in existantes:
            continue
        h0, h1 = rp.hypotheses_par_defaut(sp["type"], sp["params"], sp["nb_groupes"])
        hyps.append({"type": sp["type"], "params": dict(sp["params"]), "h0": h0, "h1": h1, "enonce": ""})
        existantes.add(cle)
        ajoutees += 1
    st.session_state["rapport_blocs"] = None
    return ajoutees


@st.cache_data(show_spinner=False)
def _suggestions(df, alpha, prioritaires, exclure):
    return ex.suggerer_tests(df, alpha, prioritaires, exclure)


# ----------------------------------------------------------------------
def onglet_debutant(df, alpha):
    """Guide d'utilisation en langage simple pour les utilisateurs débutants."""
    st.subheader("Bienvenue dans StatExcel : votre guide pas à pas")
    st.write("Vous n'avez pas besoin de maîtriser les statistiques pour commencer. Suivez les étapes ci-dessous, puis consultez l'explication affichée dans chaque rubrique d'analyse.")
    etapes = [
        ("1. Vérifier les données", "Ouvrez l'onglet Données pour vérifier les colonnes, les premières lignes et les valeurs manquantes. Une ligne représente généralement une personne, un ménage, une exploitation ou une observation."),
        ("2. Décrire votre échantillon", "Dans Descriptives, choisissez les variables à résumer. Une variable qualitative décrit une catégorie (sexe, région, type de culture). Une variable quantitative exprime une quantité (âge, revenu, superficie)."),
        ("3. Choisir une analyse", "Pour comparer des groupes, utilisez les Tests. Pour examiner un lien entre deux variables, utilisez Associations. Pour expliquer ou prédire une variable, consultez Régression. Si vous hésitez, commencez par Exploration auto."),
        ("4. Lire la p-value avec prudence", f"Le seuil choisi est α = {alpha}. En général, si la p-value est inférieure à ce seuil, le résultat est dit statistiquement significatif. Cela ne mesure pas à lui seul l'importance pratique de l'effet et ne prouve pas une causalité."),
        ("5. Préparer un rapport", "Dans Rapport automatique, sélectionnez les résultats utiles, générez l'aperçu, relisez les conclusions et téléchargez le document Word. Vérifiez toujours que les interprétations correspondent à votre question de recherche."),
    ]
    for titre, texte in etapes:
        with st.expander(titre, expanded=(titre.startswith("1."))):
            st.write(texte)
    st.info("Conseil : ne choisissez pas un test uniquement parce que son nom est familier. Le choix dépend du type de variable, du nombre de groupes, du plan d'étude et des hypothèses de recherche.")
    st.markdown("#### Repères rapides")
    st.dataframe(pd.DataFrame([
        {"Votre question": "Combien ? Quelle proportion ? Quelle moyenne ?", "Commencez par": "Descriptives"},
        {"Votre question": "La distribution est-elle compatible avec la normalité ?", "Commencez par": "Tests → normalité"},
        {"Votre question": "Deux variables qualitatives sont-elles associées ?", "Commencez par": "Associations → khi-deux"},
        {"Votre question": "Deux variables numériques évoluent-elles ensemble ?", "Commencez par": "Associations → corrélation"},
        {"Votre question": "Quels facteurs sont liés à un résultat ?", "Commencez par": "Régression"},
        {"Votre question": "Où se situent mes observations GPS ?", "Commencez par": "Carte GPS"},
    ]), hide_index=True, width="stretch")


def onglet_carte(df, ctx):
    """Carte GPS accessible directement depuis la navigation principale."""
    st.subheader("Carte GPS des observations")
    st.caption("Cette carte localise les lignes de votre fichier qui possèdent des coordonnées valides. Elle ne remplace pas une carte administrative officielle.")
    if ctx.get("coords"):
        st.success("Coordonnées détectées automatiquement. Vérifiez les colonnes sélectionnées avant d'interpréter la carte.")
    else:
        st.warning("Aucune coordonnée n'a été reconnue automatiquement. Ouvrez le choix manuel ci-dessous et sélectionnez les colonnes latitude et longitude.")
    _carte(df, ctx, key_prefix="carte_principale")


# ----------------------------------------------------------------------
def onglet_exploration(df, alpha, ctx):
    st.subheader("Exploration automatique de la base")
    st.caption("Ces analyses sont produites automatiquement à partir des variables détectées. "
               "Elles servent à explorer : elles ne remplacent pas la formulation de vos hypothèses de recherche.")
    t_profil, t_hyp, t_carte = st.tabs(["Profil socio-démographique", "Hypothèses suggérées", "Carte de la zone d'étude"])
    with t_profil:
        _profil(df, ctx)
    with t_hyp:
        _hypotheses(df, alpha, ctx)
    with t_carte:
        _carte(df, ctx, key_prefix="exploration_carte")


# ---------------------------------------------------------------- 1. profil
def _profil(df, ctx):
    profil = ctx["profil"]
    retenues = [(cat, c) for cat, cols in profil.items() for c in cols]
    deja = {c for _, c in retenues}
    if not retenues:
        st.warning("Aucune variable socio-démographique reconnue par son nom. Voici les variables qualitatives "
                   "de la base ; vous pouvez en choisir d'autres ci-dessous.")
        retenues = [("Variable", c) for c in su.colonnes_qualitatives(df)[:6]]
        deja = {c for _, c in retenues}
    else:
        with st.expander("Variables reconnues", expanded=False):
            st.dataframe(pd.DataFrame([{"Catégorie": cat, "Colonne(s)": ", ".join(cols)} for cat, cols in profil.items()]),
                         hide_index=True, width="stretch")
    extra = st.multiselect("Ajouter d'autres variables au profil",
                           [c for c in df.columns if c not in deja and c not in ctx["cols_coord"]], key="exp_extra")
    retenues += [("Variable ajoutée", c) for c in extra]

    c_age = next((c for c in profil.get("Âge", [])), None)
    c_sexe = next((c for c in profil.get("Sexe / genre", [])), None)
    if c_age and c_sexe:
        try:
            st.altair_chart(ex.pyramide_ages(df, c_age, c_sexe), width="stretch")
        except ValueError as e:
            st.caption(f"Pyramide des âges non disponible : {e}.")

    colonnes = st.columns(2)
    for i, (cat, c) in enumerate(retenues):
        with colonnes[i % 2]:
            st.markdown(f"**{cat}** — `{c}`")
            try:
                st.altair_chart(ex.graphique_variable(df, c), width="stretch")
                st.caption(ex.resume_variable(df[c]))
            except Exception as e:
                st.caption(f"Graphique impossible : {e}")

    if c_sexe:                                                        # croisements avec le sexe
        autres = [c for cat, c in retenues if c != c_sexe and not ex.est_quantitative(df[c])
                  and 2 <= df[c].nunique() <= 8][:4]
        if autres:
            st.markdown("#### Répartition selon le sexe")
            colonnes = st.columns(2)
            for i, c in enumerate(autres):
                with colonnes[i % 2]:
                    st.altair_chart(ex.croisement_empile(df, c_sexe, c), width="stretch")

    manquants = ex.graphique_manquants(df)
    if manquants is not None:
        st.markdown("#### Qualité des données")
        st.altair_chart(manquants, width="stretch")


# ---------------------------------------------------------------- 2. hypothèses
def _hypotheses(df, alpha, ctx):
    prioritaires = tuple(c for cols in ctx["profil"].values() for c in cols)
    with st.spinner("Recherche et test des hypothèses plausibles…"):
        tableau, specs, infos, normalite = _suggestions(df, alpha, prioritaires, tuple(ctx["cols_coord"]))

    if tableau.empty:
        st.info("Aucune hypothèse testable trouvée : il faut au moins une variable quantitative et une variable "
                "qualitative (2 à 8 modalités), ou deux variables de l'un des deux types.")
        return

    c1, c2, c3 = st.columns(3)
    c1.metric("Tests réalisés", infos["realises"])
    c2.metric("Significatifs avant correction", infos["sig_brut"])
    c3.metric("Significatifs après correction (FDR)", infos["sig_corrige"])
    st.caption(f"Variables quantitatives testées : {', '.join(infos['quant']) or '—'}.  \n"
               f"Variables qualitatives testées : {', '.join(infos['qual']) or '—'}.")
    st.warning("**Lecture prudente.** Quand on teste beaucoup d'hypothèses, certaines ressortent par hasard. "
               "La colonne « p ajustée (FDR) » corrige ce risque (méthode de Benjamini-Hochberg) : privilégiez les "
               "résultats « significatifs (corrigé) », et fondez vos conclusions sur des hypothèses formulées à l'avance. "
               "Le test utilisé est choisi automatiquement selon la normalité et les effectifs."
               + (f" {infos['ignores']} comparaison(s) n'ont pas été réalisées car un groupe compte moins de 5 observations."
                  if infos["ignores"] else ""))

    f1, f2 = st.columns([2, 3])
    filtre = f1.radio("Afficher", ["Tous les tests", "Significatifs (corrigés)", "Significatifs (avant correction)"],
                      key="exp_filtre")
    types = f2.multiselect("Types d'analyse", sorted(tableau["Type"].unique()),
                           default=sorted(tableau["Type"].unique()), key="exp_types")
    masque = tableau["Type"].isin(types)
    if filtre.startswith("Significatifs (corrigés"):
        masque &= tableau["_padj"] < alpha
    elif filtre.startswith("Significatifs (avant"):
        masque &= tableau["_p"] < alpha
    vue = tableau[masque].drop(columns=["_p", "_padj"])
    if vue.empty:
        st.info("Aucun test ne correspond à ce filtre.")
    else:
        vue = vue.copy()
        vue.insert(0, "Ajouter", False)
        edite = st.data_editor(
            vue, key=f"exp_ed_{filtre}_{'|'.join(types)}", width="stretch",
            disabled=[c for c in vue.columns if c != "Ajouter"],
            column_config={"Ajouter": st.column_config.CheckboxColumn("Ajouter", help="Cochez pour l'ajouter au rapport"),
                           "Lecture": st.column_config.TextColumn(width="large")})
        choisis = list(edite.index[edite["Ajouter"]])
        b1, b2 = st.columns([1, 3])
        if b1.button("Ajouter au rapport", type="primary", disabled=not choisis, key="exp_ajout"):
            ajoutees = ajouter_au_rapport(choisis, specs)
            b2.success(f"{ajoutees} hypothèse(s) ajoutée(s) au rapport (onglet « Rapport automatique »)." if ajoutees
                       else "Ces hypothèses figurent déjà dans le rapport.")

        st.markdown("#### Détail d'un test")
        ids = list(vue.index)
        i = st.selectbox("Choisissez un test pour voir son interprétation rédigée", ids, index=None,
                         format_func=lambda k: f"{k}. {tableau.loc[k, 'Analyse']} — {tableau.loc[k, 'Type']}",
                         placeholder="Sélectionnez un test…", key="exp_detail")
        if i is not None:
            sp = specs[i - 1]
            h0, h1 = rp.hypotheses_par_defaut(sp["type"], sp["params"], sp["nb_groupes"])
            st.markdown(f"**H0 :** {_md(h0)}  \n**H1 :** {_md(h1)}")
            try:
                res = rp.EXECUTION[sp["type"]](df, sp["params"], alpha, True)
                afficher_blocs(res["blocs"])
            except Exception as e:
                st.error(f"Détail indisponible : {e}")

    if not normalite.empty:
        with st.expander("Diagnostic de normalité des variables quantitatives"):
            aff = normalite.copy()
            aff["Shapiro-Wilk (p)"] = aff["Shapiro-Wilk (p)"].map(lambda p: "—" if pd.isna(p) else su.formater_p(p))
            st.dataframe(aff, width="stretch")


# ---------------------------------------------------------------- 3. assistant d'analyse automatique

def onglet_assistant_auto(df, alpha, ctx):
    """Assistant guidé : question libre, suggestions et exécution des tests recommandés."""
    st.subheader("Assistant d'analyse statistique")
    st.write("Décrivez votre objectif en langage courant. StatExcel explore les relations possibles dans vos données, "
             "propose des hypothèses et exécute les tests compatibles avec les variables disponibles.")
    question = st.text_input(
        "Que voulez-vous vérifier ?",
        placeholder="Ex. : Le revenu varie-t-il selon le sexe ? La superficie est-elle liée au rendement ?",
        key="assistant_question",
        help="Vous pouvez écrire une question de recherche ou citer les noms des variables de votre fichier."
    )
    st.caption("Exemples : comparer le revenu entre les groupes ; vérifier le lien entre deux variables ; "
               "étudier les facteurs associés au rendement.")
    prioritaires = tuple(c for cols in ctx["profil"].values() for c in cols)
    with st.spinner("Analyse automatique des variables et exécution des tests compatibles…"):
        tableau, specs, infos, normalite = _suggestions(df, alpha, prioritaires, tuple(ctx["cols_coord"]))
    if tableau.empty:
        st.warning("Je n'ai pas trouvé de test automatique exploitable. Vérifiez les types de variables, les valeurs manquantes et le nombre de modalités.")
        st.info("Pour comparer une mesure entre groupes, il faut une variable quantitative et une variable de groupe. Pour un khi-deux, il faut deux variables qualitatives. Pour une corrélation, deux variables quantitatives.")
        return

    # Filtrage simple et transparent par noms de variables / intention exprimée.
    vue = tableau.copy()
    if question.strip():
        q = question.casefold()
        noms = [str(c) for c in df.columns if len(str(c)) >= 3 and str(c).casefold() in q]
        mots_groupe = ("groupe", "selon", "différence", "comparer", "comparaison", "varie", "variation", "sexe", "région", "categorie", "catégorie")
        mots_lien = ("lien", "relation", "assoc", "association", "corrél", "lié", "liee", "liée", "influence", "facteur", "explique")
        types = []
        if any(m in q for m in mots_groupe): types.append("Comparaison de groupes")
        if any(m in q for m in mots_lien): types.extend(["Corrélation", "Association (khi-deux)"])
        mask = pd.Series(False, index=vue.index)
        if noms:
            mask |= vue["Analyse"].apply(lambda x: any(n.casefold() in str(x).casefold() for n in noms))
        if types:
            mask |= vue["Type"].isin(types)
        if mask.any():
            vue = vue[mask]
            st.info(f"Suggestions filtrées à partir de votre question ({len(vue)} test(s)). Vérifiez que les variables retenues correspondent bien à votre objectif.")
        else:
            st.info("Aucune correspondance directe détectée dans la question : voici les tests exploratoires disponibles. Vous pouvez choisir le test pertinent.")

    st.markdown("#### Résumé automatique")
    c1, c2, c3 = st.columns(3)
    c1.metric("Tests exécutés", infos.get("realises", len(tableau)))
    c2.metric("Significatifs avant correction", infos.get("sig_brut", "—"))
    c3.metric("Significatifs après correction", infos.get("sig_corrige", "—"))
    st.caption(f"Seuil de décision : α = {alpha}. Les p-values sont corrigées pour tests multiples (FDR) dans l'exploration. "
               "Une association statistique ne prouve pas à elle seule une causalité.")
    colonnes = [c for c in ["Type", "Analyse", "Test utilisé", "Effet / intensité", "N", "p-value", "p ajustée (FDR)", "Verdict", "Lecture", "Prudence"] if c in vue.columns]
    st.dataframe(vue[colonnes], hide_index=True, width="stretch")

    # Garder l'index original pour aligner chaque résultat sur ses paramètres d'exécution.
    ids = [i for i in vue.index if isinstance(i, (int, np.integer)) and 1 <= i <= len(specs)]
    if ids:
        choix = st.selectbox("Choisir un résultat pour voir les hypothèses et l'interprétation", ids,
                             format_func=lambda i: f"{tableau.loc[i, 'Analyse']} — {tableau.loc[i, 'Test utilisé']}",
                             key="assistant_detail")
        sp = specs[choix - 1]
        h0, h1 = rp.hypotheses_par_defaut(sp["type"], sp["params"], sp["nb_groupes"])
        with st.expander("Hypothèses statistiques", expanded=True):
            st.markdown(f"**H₀ (hypothèse nulle) :** {h0}")
            st.markdown(f"**H₁ (hypothèse alternative) :** {h1}")
        try:
            res = rp.EXECUTION[sp["type"]](df, sp["params"], alpha, True)
            afficher_blocs(res["blocs"])
            if st.button("Ajouter cette hypothèse au rapport automatique", key="assistant_add_report"):
                n = ajouter_au_rapport([choix], specs)
                st.success("Hypothèse ajoutée au rapport." if n else "Cette hypothèse est déjà dans le rapport.")
        except Exception as e:
            st.warning(f"Le détail de ce test n'a pas pu être généré : {e}")

    with st.expander("Comment StatExcel choisit-il les tests ?"):
        st.markdown("- **Quantitative + groupe** : comparaison de groupes (test paramétrique ou non paramétrique selon les données).\n"
                    "- **Deux variables qualitatives** : test du khi-deux d'indépendance.\n"
                    "- **Deux variables quantitatives** : corrélation de Pearson ou de Spearman selon les caractéristiques observées.\n"
                    "- **Plusieurs tests à la fois** : une correction FDR aide à limiter les découvertes dues au hasard.")
        st.warning("L'assistant propose des pistes exploratoires à valider selon votre protocole, votre échantillonnage et vos hypothèses formulées avant l'analyse. Il ne déduit pas automatiquement une causalité du texte saisi.")


# ---------------------------------------------------------------- 4. carte
def _carte(df, ctx, key_prefix="geo"):

    coords = ctx["coords"]
    numeriques = su.colonnes_numeriques(df)
    with st.expander("Choisir les colonnes de coordonnées manuellement", expanded=not coords):
        manuel = st.checkbox("Utiliser mon choix plutôt que la détection automatique", value=False, key=f"{key_prefix}_manuel")
        c1, c2 = st.columns(2)
        m_lat = c1.selectbox("Colonne de latitude", numeriques or ["—"], key=f"{key_prefix}_lat")
        m_lon = c2.selectbox("Colonne de longitude", numeriques or ["—"], key=f"{key_prefix}_lon")
    if manuel and numeriques:
        coords = (m_lat, m_lon)
    if not coords:
        st.info("Aucune coordonnée géographique détectée. Pour obtenir la carte, ajoutez à votre fichier deux colonnes "
                "**Latitude** et **Longitude** (degrés décimaux, par exemple 6,17 et 1,23), ou une colonne GPS "
                "contenant « latitude, longitude ».")
        return

    pts, notes = ex.extraire_points(df, coords[0], coords[1])
    for n in notes:
        st.warning(n)
    masquer = False
    if pts["aberrant"].any():
        masquer = st.checkbox("Masquer les points très éloignés (probables erreurs de saisie)", value=True,
                              key=f"{key_prefix}_masquer")
    utiles = pts[pts["valide"] & ~(pts["aberrant"] & masquer)]
    if utiles.empty:
        st.error("Aucun point valide : vérifiez les colonnes de coordonnées choisies.")
        return
    res = ex.resume_zone(utiles)

    st.caption(f"Colonnes utilisées : « {coords[0]} »" + (f" et « {coords[1]} »." if coords[1] else " (colonne combinée)."))
    m = st.columns(4)
    m[0].metric("Points localisés", f"{res['n']} / {len(df)}")
    m[1].metric("Centre (lat, lon)", f"{res['centre'][0]:.4f}, {res['centre'][1]:.4f}")
    m[2].metric("Étendue N–S × E–O", f"{res['ns_km']:.1f} × {res['eo_km']:.1f} km")
    m[3].metric("Superficie de la zone", "—" if pd.isna(res["aire_km2"]) else f"{res['aire_km2']:,.1f} km²".replace(",", " "))

    candidats = [c for c in su.colonnes_groupes(df, max_modalites=30) if c not in ctx["cols_coord"]]
    prefere = (ctx["profil"].get("Lieu de résidence") or [None])[0]
    options = ["— Aucune —"] + candidats
    colorier = st.selectbox("Colorer les points selon", options,
                            index=options.index(prefere) if prefere in options else 0, key=f"{key_prefix}_couleur")

    carte = pd.DataFrame({"latitude": utiles["lat"], "longitude": utiles["lon"]})
    taille = float(min(max(max(res["ns_km"], res["eo_km"]) * 6, 15), 2500))
    couleurs = None
    if colorier != "— Aucune —":
        serie = df.loc[utiles.index, colorier]
        valides = serie.notna()
        carte, serie = carte[valides], serie[valides]
        hexes, table, txt = ex.couleurs_modalites(serie)
        carte["couleur"] = hexes
        couleurs = txt
        res = ex.resume_zone(utiles.loc[carte.index])
    left, right = st.columns([3, 2])
    with left:
        st.markdown("#### Carte GPS détaillée")
        st.caption("Fond cartographique, points GPS et itinéraire indicatif entre les points. Le tracé en pointillés relie les observations dans l'ordre du fichier ; il ne représente pas une route réelle.")
        try:
            import folium
            from streamlit.components.v1 import html as st_html
            centre = [float(carte["latitude"].median()), float(carte["longitude"].median())]
            fmap = folium.Map(location=centre, zoom_start=11, control_scale=True, tiles="OpenStreetMap")
            # Ajouter un contour administratif officiel fourni par l'utilisateur, si disponible.
            geojson_upload = st.file_uploader(
                "Ajouter les limites administratives (GeoJSON, facultatif)",
                type=["geojson", "json"], key=f"{key_prefix}_geojson",
                help="Pour obtenir une carte comme l'exemple avec les préfectures colorées et leurs noms, importez un fichier GeoJSON des limites administratives de la zone étudiée."
            )
            if geojson_upload is not None:
                import json
                try:
                    geojson_data = json.load(geojson_upload)
                    folium.GeoJson(geojson_data, name="Limites administratives",
                                   style_function=lambda feat: {"color": "#555555", "weight": 1.4, "fillColor": "#f2d88a", "fillOpacity": 0.18},
                                   tooltip=folium.GeoJsonTooltip(fields=[k for k in (geojson_data.get("features", [{}])[0].get("properties", {}).keys()) if k][:2]) if geojson_data.get("features") and geojson_data["features"][0].get("properties") else None).add_to(fmap)
                except Exception as e:
                    st.warning(f"Impossible de lire ce GeoJSON : {e}")
            # Tracé rouge discontinu rappelant la carte de référence.
            coords_ligne = carte[["latitude", "longitude"]].dropna().values.tolist()
            if len(coords_ligne) > 1:
                folium.PolyLine(coords_ligne, color="#c62828", weight=2.5, opacity=0.85, dash_array="7, 7", tooltip="Liaison indicative entre observations").add_to(fmap)
            palette = ["#c62828", "#1565c0", "#2e7d32", "#6a1b9a", "#ef6c00", "#00838f"]
            for idx, row in carte.iterrows():
                cat = str(df.loc[idx, colorier]) if colorier != "— Aucune —" and idx in df.index else "Observation GPS"
                couleur = str(row.get("couleur", "#c62828"))
                popup = f"Observation : {html.escape(str(idx + 1 if isinstance(idx, (int, np.integer)) else idx))}<br>Latitude : {row['latitude']:.6f}<br>Longitude : {row['longitude']:.6f}"
                if colorier != "— Aucune —": popup += f"<br>{html.escape(str(colorier))} : {html.escape(cat)}"
                folium.Marker([float(row["latitude"]), float(row["longitude"])],
                              tooltip=f"Point GPS · {html.escape(cat)}",
                              popup=folium.Popup(popup, max_width=300),
                              icon=folium.DivIcon(html=f'<div style="font-size:14px;color:{couleur};font-weight:bold;text-shadow:1px 1px white">■</div>')).add_to(fmap)
            folium.LayerControl(collapsed=True).add_to(fmap)
            # Rose des vents discrète, similaire à la carte de référence.
            fmap.get_root().html.add_child(folium.Element('<div style="position:fixed;top:90px;left:25px;z-index:9999;background:#ffffffcc;border:1px solid #777;border-radius:50%;width:42px;height:42px;text-align:center;line-height:40px;font-weight:bold;color:#222;pointer-events:none">N<br><span style="font-size:10px">✥</span></div>'))
            st_html(fmap.get_root().render(), height=520, scrolling=False)
        except ImportError:
            st.warning("Pour afficher la carte détaillée avec fond géographique, installez Folium : `python -m pip install folium`. Carte simplifiée affichée en attendant.")
            st.map(carte, latitude="latitude", longitude="longitude", size=taille, height=480,
                   **({"color": "couleur"} if couleurs is not None else {}))
        except Exception as e:
            st.warning(f"La carte détaillée n'a pas pu être construite ({e}). Carte simplifiée affichée.")
            st.map(carte, latitude="latitude", longitude="longitude", size=taille, height=480,
                   **({"color": "couleur"} if couleurs is not None else {}))
        if couleurs is not None:
            effectifs = couleurs.value_counts()
            st.markdown(" &nbsp; ".join(
                f"<span style='color:{col}'>●</span> {html.escape(str(mod))} ({int(effectifs.get(mod, 0))})"
                for mod, col in table.items()), unsafe_allow_html=True)
        st.caption("Le fond de carte nécessite une connexion internet. Les limites administratives sont affichées seulement si vous fournissez un GeoJSON correspondant à la zone étudiée.")
    with right:
        st.markdown("#### Emprise")
        st.altair_chart(ex.schema_emprise(res, couleurs, table if couleurs is not None else None), width="content")
        st.caption("Repère en kilomètres centré sur le point médian. La superficie est celle de l'enveloppe convexe "
                   "des points : c'est une estimation de la zone couverte, pas une limite administrative.")

    lieu = prefere or (colorier if colorier != "— Aucune —" else None)
    if lieu:
        st.markdown(f"#### Points localisés par « {lieu} »")
        t = df.loc[utiles.index, lieu].map(su.texte_modalite).value_counts().rename("Points").to_frame()
        t["%"] = (100 * t["Points"] / t["Points"].sum()).round(1)
        st.dataframe(t, width="stretch")
