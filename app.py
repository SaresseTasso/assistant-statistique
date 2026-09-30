# ============================================================
# ASSISTANT STATISTIQUE — application Streamlit (version complète)
# Lancer avec :  streamlit run assistant_statistique.py
# Dépendances :  streamlit pandas numpy scipy scikit-learn openpyxl
# ============================================================

import io
import re
import unicodedata

import numpy as np
import pandas as pd
import streamlit as st

from scipy.stats import (
    chi2_contingency,
    fisher_exact,
    pearsonr,
    spearmanr,
    ttest_ind,
    mannwhitneyu,
    f_oneway,
    kruskal,
)

# scikit-learn n'est nécessaire que pour la section 11.
# Import protégé : si le paquet manque, le reste de l'application fonctionne.
try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.cluster import KMeans

    SKLEARN_DISPONIBLE = True
except ImportError:
    SKLEARN_DISPONIBLE = False

# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Assistant statistique",
    page_icon="📊",
    layout="wide",
)

st.title("Assistant statistique")

st.write(
    "Analysez vos fichiers Excel ou CSV : diagnostic, "
    "dictionnaire, nettoyage, statistiques descriptives "
    "et analyses bivariées."
)

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

MIME_XLSX = (
    "application/vnd.openxmlformats-officedocument."
    "spreadsheetml.sheet"
)


def afficher_tableau(donnees, hide_index=True):
    """st.dataframe pleine largeur, compatible anciennes et nouvelles versions."""
    try:
        st.dataframe(donnees, width="stretch", hide_index=hide_index)
    except Exception:
        st.dataframe(
            donnees, use_container_width=True, hide_index=hide_index
        )


def editeur_tableau(donnees, **kwargs):
    """st.data_editor pleine largeur, compatible anciennes et nouvelles versions."""
    try:
        return st.data_editor(
            donnees, width="stretch", hide_index=True, **kwargs
        )
    except Exception:
        return st.data_editor(
            donnees, use_container_width=True, hide_index=True, **kwargs
        )


def table_indicateurs(paires, decimales=4):
    """Tableau Indicateur / Valeur. Les valeurs sont converties en texte
    pour éviter les colonnes de types mélangés (texte + nombres)."""
    valeurs = []
    for _, valeur in paires:
        if isinstance(valeur, (float, np.floating)):
            valeurs.append(
                "—" if np.isnan(valeur) else f"{valeur:.{decimales}f}"
            )
        else:
            valeurs.append(str(valeur))
    return pd.DataFrame(
        {
            "Indicateur": [p[0] for p in paires],
            "Valeur": valeurs,
        }
    )


def format_p(p_value):
    if p_value is None or np.isnan(p_value):
        return "—"
    if p_value < 0.0001:
        return "< 0.0001"
    return f"{p_value:.4f}"


def normaliser_cle(texte):
    """Minuscules, sans accents, sans espaces superflus (comparaisons)."""
    texte = str(texte).replace("’", "'")
    texte = unicodedata.normalize("NFKD", texte)
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", texte.strip().lower())


def lire_csv(fichier_csv):
    """Lecture CSV avec détection du séparateur et de l'encodage."""
    derniere_erreur = None
    for encodage in ("utf-8", "utf-8-sig", "latin-1"):
        try:
            fichier_csv.seek(0)
            return pd.read_csv(
                fichier_csv, sep=None, engine="python", encoding=encodage
            )
        except UnicodeDecodeError as e:
            derniere_erreur = e
    raise derniere_erreur


# ============================================================
# IMPORTATION DU FICHIER
# ============================================================

fichier = st.file_uploader("Choisissez votre fichier de données", type=["xlsx", "csv"])

if fichier is None:
    st.info("Veuillez importer un fichier Excel ou CSV.")
    st.stop()

try:
    if fichier.name.lower().endswith(".xlsx"):
        df = pd.read_excel(fichier)
    else:
        df = lire_csv(fichier)

except Exception as e:
    st.error(f"Erreur lors de la lecture du fichier : {e}")
    st.stop()

df.columns = [str(c) for c in df.columns]

st.success(f"Fichier chargé : {fichier.name}")

# ============================================================
# INITIALISATION DES DONNEES NETTOYEES
# ============================================================
# On conserve toujours une copie des données originales (df).
# Toutes les corrections sont appliquées sur df_nettoye.

cle_fichier = f"{fichier.name}|{fichier.size}"

if st.session_state.get("cle_fichier_actuel") != cle_fichier:
    st.session_state["cle_fichier_actuel"] = cle_fichier
    st.session_state["df_nettoye"] = df.copy()
    st.session_state["dictionnaire_modifie"] = None
    st.session_state["donnees_nettoyage_valide"] = False

df_nettoye = st.session_state["df_nettoye"]

# ============================================================
# 1. INFORMATIONS GENERALES
# ============================================================

st.subheader("1. Informations générales")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("Lignes", df.shape[0])

with col2:
    st.metric("Variables", df.shape[1])

with col3:
    st.metric("Doublons", int(df.duplicated().sum()))

with col4:
    st.metric("Cellules manquantes", int(df.isna().sum().sum()))

# ============================================================
# 2. APERCU DES DONNEES
# ============================================================

st.subheader("2. Aperçu des données")

afficher_tableau(df.head(20))

# ============================================================
# 3. DIAGNOSTIC DES VARIABLES
# ============================================================

st.subheader("3. Diagnostic des variables")

diagnostic = pd.DataFrame(
    {
        "Variable": df.columns,
        "Type Python": [str(df[c].dtype) for c in df.columns],
        "Valeurs manquantes": [int(df[c].isna().sum()) for c in df.columns],
        "% manquant": [
            round(df[c].isna().mean() * 100, 2) for c in df.columns
        ],
        "Valeurs uniques": [
            int(df[c].nunique(dropna=True)) for c in df.columns
        ],
    }
)

afficher_tableau(diagnostic)

if df.duplicated().sum() > 0:
    st.warning(f"{df.duplicated().sum()} doublon(s) détecté(s).")
else:
    st.success("Aucun doublon détecté.")

# ============================================================
# PROPOSITION DU TYPE DE VARIABLE
# ============================================================


def proposer_type(colonne):
    serie = df[colonne]

    if pd.api.types.is_datetime64_any_dtype(serie):
        return "Date"

    if pd.api.types.is_numeric_dtype(serie):
        return "Quantitative"

    return "Qualitative"


types_analyse = [
    "Qualitative",
    "Qualitative codée",
    "Quantitative",
    "Date",
    "Identifiant",
    "À vérifier",
]

types_questions = [
    "Question fermée",
    "Question ouverte",
    "Réponses multiples",
    "Non applicable",
]

types_qualitatifs = ["Qualitative", "Qualitative codée"]

# ============================================================
# 4. DICTIONNAIRE DES VARIABLES
# ============================================================

dictionnaire = pd.DataFrame(
    {
        "Variable": df.columns,
        "Type Python": [str(df[c].dtype) for c in df.columns],
        "Nombre de modalités": [
            int(df[c].nunique(dropna=True)) for c in df.columns
        ],
        "Valeurs manquantes": [int(df[c].isna().sum()) for c in df.columns],
    }
)

dictionnaire["Type d'analyse"] = [proposer_type(c) for c in df.columns]

dictionnaire["Type de question"] = [
    "Non applicable"
    if proposer_type(c) in ["Quantitative", "Date", "Identifiant"]
    else "Question fermée"
    for c in df.columns
]

colonnes_requises = [
    "Variable",
    "Type Python",
    "Nombre de modalités",
    "Valeurs manquantes",
    "Type d'analyse",
    "Type de question",
]

# Le dictionnaire est recréé s'il n'existe pas, s'il est incomplet
# ou si les variables du fichier ont changé.
dico_actuel = st.session_state["dictionnaire_modifie"]

if (
    dico_actuel is None
    or not all(c in dico_actuel.columns for c in colonnes_requises)
    or list(dico_actuel["Variable"]) != list(df.columns)
):
    st.session_state["dictionnaire_modifie"] = dictionnaire.copy()

# Les colonnes descriptives (non modifiables) sont rafraîchies à partir
# des données nettoyées ; les choix de l'utilisateur sont conservés.
dico_base = st.session_state["dictionnaire_modifie"].copy()
dico_base["Type Python"] = [str(df_nettoye[c].dtype) for c in dico_base["Variable"]]
dico_base["Nombre de modalités"] = [
    int(df_nettoye[c].nunique(dropna=True)) for c in dico_base["Variable"]
]
dico_base["Valeurs manquantes"] = [
    int(df_nettoye[c].isna().sum()) for c in dico_base["Variable"]
]

st.subheader("4. Dictionnaire des variables")

st.caption(
    "Vérifiez le type de chaque variable. Pour activer la section 11, "
    "passez « Type de question » à « Question ouverte » pour les "
    "variables concernées."
)

dictionnaire_modifie = editeur_tableau(
    dico_base,
    column_config={
        "Type d'analyse": st.column_config.SelectboxColumn(
            "Type d'analyse", options=types_analyse
        ),
        "Type de question": st.column_config.SelectboxColumn(
            "Type de question", options=types_questions
        ),
    },
    disabled=[
        "Variable",
        "Type Python",
        "Nombre de modalités",
        "Valeurs manquantes",
    ],
)

st.session_state["dictionnaire_modifie"] = dictionnaire_modifie

# ============================================================
# 5. ANALYSE DES VARIABLES QUALITATIVES
# ============================================================

st.subheader("5. Analyse des variables qualitatives")

for _, ligne in dictionnaire_modifie.iterrows():
    variable = ligne["Variable"]
    type_analyse = ligne["Type d'analyse"]
    type_question = ligne["Type de question"]

    if variable not in df_nettoye.columns:
        continue

    # --------------------------------------------------------
    # QUESTION FERMEE
    # --------------------------------------------------------
    if type_analyse in types_qualitatifs and type_question == "Question fermée":
        st.markdown(f"### {variable}")

        serie = df_nettoye[variable]
        valide = serie.dropna()

        if len(valide) == 0:
            st.warning("Aucune réponse exploitable.")
            continue

        effectifs = valide.value_counts()
        pourcentages = valide.value_counts(normalize=True) * 100

        resultat = pd.DataFrame(
            {
                "Modalité": effectifs.index.astype(str),
                "Effectif": effectifs.values,
                "Pourcentage": pourcentages.values.round(2),
            }
        )

        afficher_tableau(resultat)

        st.caption(
            f"Réponses valides : {len(valide)} | "
            f"Manquantes : {serie.isna().sum()} | "
            f"Total : {len(serie)}"
        )

        st.bar_chart(resultat.set_index("Modalité")["Effectif"])

    # --------------------------------------------------------
    # QUESTION OUVERTE
    # --------------------------------------------------------
    elif type_analyse in types_qualitatifs and type_question == "Question ouverte":
        st.markdown(f"### {variable}")

        serie = df_nettoye[variable].dropna().astype(str)

        st.write(f"**Nombre de réponses : {len(serie)}**")

        afficher_tableau(pd.DataFrame({"Réponses": serie.head(20).values}))

        st.info(
            "Cette variable est ouverte. "
            "Elle sera traitée dans le module de codification thématique "
            "(section 11)."
        )

    # --------------------------------------------------------
    # REPONSES MULTIPLES
    # --------------------------------------------------------
    elif type_analyse in types_qualitatifs and type_question == "Réponses multiples":
        st.markdown(f"### {variable}")

        serie = df_nettoye[variable].dropna().astype(str)

        reponses = []
        for valeur in serie:
            for morceau in re.split(r"[,;|]", valeur):
                morceau = morceau.strip()
                if morceau:
                    reponses.append(morceau)

        if reponses:
            freq = pd.Series(reponses).value_counts()
            pourcentage = (freq / len(serie)) * 100

            resultat_multiple = pd.DataFrame(
                {
                    "Réponse": freq.index,
                    "Nombre de citations": freq.values,
                    "% des répondants": pourcentage.round(2).values,
                }
            )

            afficher_tableau(resultat_multiple)

            st.caption(
                "Les pourcentages peuvent dépasser 100 % au total car un "
                "répondant peut avoir plusieurs réponses."
            )
        else:
            st.warning("Aucune réponse multiple exploitable détectée.")

# ============================================================
# 6. ANALYSE DES VARIABLES QUANTITATIVES
# ============================================================

st.subheader("6. Analyse des variables quantitatives")

for _, ligne in dictionnaire_modifie.iterrows():
    variable = ligne["Variable"]

    if ligne["Type d'analyse"] == "Quantitative" and variable in df_nettoye.columns:
        st.markdown(f"### {variable}")

        serie = pd.to_numeric(df_nettoye[variable], errors="coerce").dropna()

        if len(serie) == 0:
            st.warning("Aucune valeur numérique exploitable.")
            continue

        statistiques = table_indicateurs(
            [
                ("Effectif valide", len(serie)),
                ("Valeurs manquantes", int(df_nettoye[variable].isna().sum())),
                ("Moyenne", serie.mean()),
                ("Médiane", serie.median()),
                ("Écart-type", serie.std()),
                ("Minimum", serie.min()),
                ("Q1", serie.quantile(0.25)),
                ("Q3", serie.quantile(0.75)),
                ("Maximum", serie.max()),
            ],
            decimales=2,
        )

        afficher_tableau(statistiques)

# ============================================================
# 7. RESUME DU DIAGNOSTIC
# ============================================================

st.subheader("7. Résumé du diagnostic")

type_col = dictionnaire_modifie["Type d'analyse"]

nb_qualitatives = int(type_col.isin(types_qualitatifs).sum())
nb_quantitatives = int((type_col == "Quantitative").sum())
nb_dates = int((type_col == "Date").sum())
nb_identifiants = int((type_col == "Identifiant").sum())

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric("Variables qualitatives", nb_qualitatives)

with c2:
    st.metric("Variables quantitatives", nb_quantitatives)

with c3:
    st.metric("Variables de date", nb_dates)

with c4:
    st.metric("Identifiants", nb_identifiants)

# ============================================================
# 8. NETTOYAGE DES DONNEES
# ============================================================

st.subheader("8. Nettoyage des données")

st.write(
    "Les données originales restent conservées. "
    "Les corrections sont appliquées uniquement sur une copie de travail."
)

# ------------------------------------------------------------
# 8.1 ETAT ACTUEL
# ------------------------------------------------------------

st.markdown("### 8.1 État actuel des données")

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric("Lignes", df_nettoye.shape[0])

with c2:
    st.metric("Variables", df_nettoye.shape[1])

with c3:
    st.metric("Cellules manquantes", int(df_nettoye.isna().sum().sum()))

with c4:
    st.metric("Doublons", int(df_nettoye.duplicated().sum()))

# ------------------------------------------------------------
# 8.2 VALEURS MANQUANTES
# ------------------------------------------------------------

st.markdown("### 8.2 Valeurs manquantes")

manquants = pd.DataFrame(
    {
        "Variable": df_nettoye.columns,
        "Valeurs manquantes": [
            int(df_nettoye[c].isna().sum()) for c in df_nettoye.columns
        ],
        "% manquant": [
            round(df_nettoye[c].isna().mean() * 100, 2)
            for c in df_nettoye.columns
        ],
    }
)

manquants = manquants[manquants["Valeurs manquantes"] > 0]

if len(manquants) > 0:
    afficher_tableau(manquants)
else:
    st.success("Aucune valeur manquante détectée.")

# ------------------------------------------------------------
# 8.3 VALEURS TEXTUELLES REPRESENTANT DES MANQUANTS
# ------------------------------------------------------------

st.markdown("### 8.3 Valeurs utilisées comme absence de réponse")

valeurs_manquantes_textuelles = [
    "",
    " ",
    "na",
    "n/a",
    "nan",
    "nd",
    "n.d.",
    "non disponible",
    "non renseigné",
    "non renseigne",
    "aucun",
    "aucune",
    "néant",
    "neant",
    "ras",
    "r.a.s.",
    "-",
]

colonnes_textuelles = df_nettoye.select_dtypes(include=["object", "string"]).columns

resultats_faux_manquants = []

for col in colonnes_textuelles:
    serie = df_nettoye[col].astype("string").str.strip().str.lower()
    nombre = int(serie.isin(valeurs_manquantes_textuelles).sum())

    if nombre > 0:
        resultats_faux_manquants.append(
            {"Variable": col, "Valeurs détectées": nombre}
        )

if resultats_faux_manquants:
    afficher_tableau(pd.DataFrame(resultats_faux_manquants))

    st.caption(
        "Attention : dans une question ouverte, « aucun » ou « RAS » peut "
        "être une vraie réponse. Vérifiez avant de convertir."
    )

    if st.button("Convertir en valeurs manquantes", key="convertir_faux_manquants"):
        df_nettoye = df_nettoye.copy()

        for col in colonnes_textuelles:
            serie = df_nettoye[col].astype("string").str.strip().str.lower()
            masque = serie.isin(valeurs_manquantes_textuelles)
            df_nettoye.loc[masque, col] = np.nan

        st.session_state["df_nettoye"] = df_nettoye
        st.success("Conversion effectuée.")
        st.rerun()
else:
    st.info(
        "Aucune valeur textuelle évidente représentant une absence de "
        "réponse n'a été détectée."
    )

# ------------------------------------------------------------
# 8.4 ESPACES INUTILES
# ------------------------------------------------------------

st.markdown("### 8.4 Espaces inutiles")

espaces_detectes = []

for col in colonnes_textuelles:
    serie = df_nettoye[col].astype("string")
    masque = (serie.notna() & (serie != serie.str.strip())).fillna(False)
    nombre = int(masque.sum())

    if nombre > 0:
        espaces_detectes.append({"Variable": col, "Réponses concernées": nombre})

if espaces_detectes:
    afficher_tableau(pd.DataFrame(espaces_detectes))

    if st.button("Supprimer les espaces inutiles", key="supprimer_espaces"):
        df_nettoye = df_nettoye.copy()

        for col in colonnes_textuelles:
            df_nettoye[col] = df_nettoye[col].astype("string").str.strip()

        st.session_state["df_nettoye"] = df_nettoye
        st.success("Les espaces inutiles ont été supprimés.")
        st.rerun()
else:
    st.success("Aucun espace inutile détecté.")

# ------------------------------------------------------------
# 8.5 DOUBLONS
# ------------------------------------------------------------

st.markdown("### 8.5 Doublons")

nombre_doublons = int(df_nettoye.duplicated().sum())

if nombre_doublons > 0:
    st.warning(f"{nombre_doublons} doublon(s) exact(s) détecté(s).")

    afficher_tableau(df_nettoye[df_nettoye.duplicated(keep=False)].head(50))

    if st.button("Supprimer les doublons exacts", key="supprimer_doublons"):
        avant = len(df_nettoye)
        df_nettoye = df_nettoye.drop_duplicates().reset_index(drop=True)
        apres = len(df_nettoye)

        st.session_state["df_nettoye"] = df_nettoye
        st.success(f"{avant - apres} doublon(s) supprimé(s).")
        st.rerun()
else:
    st.success("Aucun doublon exact détecté.")

# ------------------------------------------------------------
# 8.6 UNIFORMISATION DES TEXTES
# ------------------------------------------------------------

st.markdown("### 8.6 Uniformisation des réponses textuelles")

if len(colonnes_textuelles) > 0:
    variable_uniformisation = st.selectbox(
        "Variable textuelle",
        list(colonnes_textuelles),
        key="variable_uniformisation",
    )

    col_a, col_b = st.columns(2)

    with col_a:
        if st.button("Supprimer les espaces", key="uniformiser_espaces"):
            df_nettoye = df_nettoye.copy()
            df_nettoye[variable_uniformisation] = (
                df_nettoye[variable_uniformisation].astype("string").str.strip()
            )
            st.session_state["df_nettoye"] = df_nettoye
            st.success("Espaces supprimés.")
            st.rerun()

    with col_b:
        if st.button("Mettre en minuscules", key="mettre_minuscules"):
            df_nettoye = df_nettoye.copy()
            df_nettoye[variable_uniformisation] = (
                df_nettoye[variable_uniformisation].astype("string").str.lower()
            )
            st.session_state["df_nettoye"] = df_nettoye
            st.success("Réponses mises en minuscules.")
            st.rerun()
else:
    st.info("Aucune variable textuelle à uniformiser.")

# ------------------------------------------------------------
# 8.7 VERIFICATION NUMERIQUE
# ------------------------------------------------------------

st.markdown("### 8.7 Vérification des variables quantitatives")

variables_quantitatives = dictionnaire_modifie.loc[
    dictionnaire_modifie["Type d'analyse"] == "Quantitative", "Variable"
].tolist()

if variables_quantitatives:
    variable_numerique = st.selectbox(
        "Variable quantitative",
        variables_quantitatives,
        key="variable_numerique_nettoyage",
    )

    serie_originale = df_nettoye[variable_numerique]
    serie_convertie = pd.to_numeric(serie_originale, errors="coerce")

    masque_non_convertible = serie_originale.notna() & serie_convertie.isna()
    valeurs_non_convertibles = int(masque_non_convertible.sum())

    st.write(f"Valeurs non numériques détectées : **{valeurs_non_convertibles}**")

    if valeurs_non_convertibles > 0:
        valeurs_problematiques = (
            serie_originale[masque_non_convertible]
            .astype(str)
            .value_counts()
            .reset_index()
        )
        valeurs_problematiques.columns = ["Valeur", "Effectif"]

        afficher_tableau(valeurs_problematiques)

        st.warning(
            "Ces valeurs seront transformées en valeurs manquantes "
            "lors de la conversion."
        )

    if st.button("Convertir en numérique", key="convertir_numerique"):
        df_nettoye = df_nettoye.copy()
        df_nettoye[variable_numerique] = pd.to_numeric(
            df_nettoye[variable_numerique], errors="coerce"
        )
        st.session_state["df_nettoye"] = df_nettoye
        st.success("Conversion numérique effectuée.")
        st.rerun()
else:
    st.info("Aucune variable quantitative n'est actuellement identifiée.")

# ------------------------------------------------------------
# 8.8 VERIFICATION DES DATES
# ------------------------------------------------------------

st.markdown("### 8.8 Vérification des dates")

variables_dates = dictionnaire_modifie.loc[
    dictionnaire_modifie["Type d'analyse"] == "Date", "Variable"
].tolist()

if variables_dates:
    variable_date = st.selectbox(
        "Variable de date",
        variables_dates,
        key="variable_date_nettoyage",
    )

    serie_date = pd.to_datetime(df_nettoye[variable_date], errors="coerce")

    masque_date_invalide = df_nettoye[variable_date].notna() & serie_date.isna()
    valeurs_date_invalides = int(masque_date_invalide.sum())

    st.write(f"Dates non reconnues : **{valeurs_date_invalides}**")

    if valeurs_date_invalides > 0:
        dates_problematiques = (
            df_nettoye.loc[masque_date_invalide, variable_date]
            .astype(str)
            .value_counts()
            .reset_index()
        )
        dates_problematiques.columns = ["Valeur", "Effectif"]

        afficher_tableau(dates_problematiques)

    if st.button("Convertir en date", key="convertir_date"):
        df_nettoye = df_nettoye.copy()
        df_nettoye[variable_date] = pd.to_datetime(
            df_nettoye[variable_date], errors="coerce"
        )
        st.session_state["df_nettoye"] = df_nettoye
        st.success("Conversion des dates effectuée.")
        st.rerun()
else:
    st.info("Aucune variable de date n'est actuellement identifiée.")

# ------------------------------------------------------------
# 8.9 APERCU APRES NETTOYAGE
# ------------------------------------------------------------

st.markdown("### 8.9 Aperçu des données après nettoyage")

afficher_tableau(df_nettoye.head(20))

# ------------------------------------------------------------
# 8.10 COMPARAISON AVANT / APRES
# ------------------------------------------------------------

st.markdown("### 8.10 Comparaison avant / après")

comparaison = pd.DataFrame(
    {
        "Indicateur": [
            "Nombre de lignes",
            "Nombre de variables",
            "Cellules manquantes",
            "Doublons",
        ],
        "Avant nettoyage": [
            df.shape[0],
            df.shape[1],
            int(df.isna().sum().sum()),
            int(df.duplicated().sum()),
        ],
        "Après nettoyage": [
            df_nettoye.shape[0],
            df_nettoye.shape[1],
            int(df_nettoye.isna().sum().sum()),
            int(df_nettoye.duplicated().sum()),
        ],
    }
)

afficher_tableau(comparaison)

# ------------------------------------------------------------
# 8.11 VALIDATION
# ------------------------------------------------------------

st.markdown("### 8.11 Validation du nettoyage")

if st.button("Valider les données nettoyées", key="valider_nettoyage"):
    st.session_state["donnees_nettoyage_valide"] = True

if st.session_state.get("donnees_nettoyage_valide"):
    st.success("Les données nettoyées sont validées pour la suite de l'analyse.")

# ------------------------------------------------------------
# REINITIALISATION
# ------------------------------------------------------------

if st.button("Réinitialiser le nettoyage", key="reset_nettoyage"):
    st.session_state["df_nettoye"] = df.copy()
    st.session_state["donnees_nettoyage_valide"] = False
    st.success("Le nettoyage a été réinitialisé.")
    st.rerun()

# ============================================================
# 9. ANALYSE BIVARIEE
# ============================================================

st.subheader("9. Analyse bivariée")

st.write(
    "Sélectionnez deux variables pour étudier leur relation ou leur différence."
)


def comparer_groupes(var_quali, var_quant, titre):
    """Compare une variable quantitative entre les modalités d'une variable
    qualitative (2 groupes : Welch / Mann-Whitney ; 3+ : ANOVA / Kruskal)."""
    st.markdown(titre)

    donnees = df_nettoye[[var_quali, var_quant]].copy()
    donnees[var_quant] = pd.to_numeric(donnees[var_quant], errors="coerce")
    donnees = donnees.dropna()

    groupes = {
        nom: groupe[var_quant].values
        for nom, groupe in donnees.groupby(var_quali)
    }

    if len(groupes) < 2:
        st.warning("Il faut au moins deux groupes pour réaliser une comparaison.")
        return

    statistiques_groupes = (
        donnees.groupby(var_quali)[var_quant]
        .agg(
            Effectif="count",
            Moyenne="mean",
            Médiane="median",
            Écart_type="std",
            Minimum="min",
            Maximum="max",
        )
        .reset_index()
    )

    colonnes_arrondir = ["Moyenne", "Médiane", "Écart_type", "Minimum", "Maximum"]
    statistiques_groupes[colonnes_arrondir] = statistiques_groupes[
        colonnes_arrondir
    ].round(2)

    st.markdown("#### Statistiques par groupe")
    afficher_tableau(statistiques_groupes)

    # Les tests exigent au moins 2 observations par groupe.
    valides = {nom: vals for nom, vals in groupes.items() if len(vals) >= 2}

    if len(valides) < len(groupes):
        st.caption(
            "Les groupes de moins de 2 observations sont affichés "
            "mais exclus du test."
        )

    if len(valides) < 2:
        st.warning(
            "Pas assez de groupes avec au moins 2 observations "
            "pour réaliser un test."
        )
    else:
        listes = list(valides.values())
        effectif_test = int(sum(len(v) for v in listes))
        statistique, p_value = None, None

        if len(listes) == 2:
            methode = st.selectbox(
                "Test de comparaison",
                ["t-test de Welch", "Mann-Whitney"],
                key="test_deux_groupes",
            )

            try:
                if methode == "t-test de Welch":
                    statistique, p_value = ttest_ind(
                        listes[0], listes[1], equal_var=False
                    )
                else:
                    statistique, p_value = mannwhitneyu(
                        listes[0], listes[1], alternative="two-sided"
                    )
            except ValueError as e:
                st.warning(f"Le test n'a pas pu être calculé : {e}")

            libelle_signif = "La différence entre les deux groupes est"
            libelle_non_signif = "La différence entre les deux groupes n'est pas"
            precision = ""
        else:
            methode = st.selectbox(
                "Test de comparaison",
                ["ANOVA à un facteur", "Kruskal-Wallis"],
                key="test_plusieurs_groupes",
            )

            try:
                if methode == "ANOVA à un facteur":
                    statistique, p_value = f_oneway(*listes)
                else:
                    statistique, p_value = kruskal(*listes)
            except ValueError as e:
                st.warning(f"Le test n'a pas pu être calculé : {e}")

            libelle_signif = "Le test détecte une différence"
            libelle_non_signif = "Le test ne détecte pas de différence"
            precision = (
                "Ce résultat ne précise pas à lui seul quels groupes "
                "diffèrent entre eux. Des comparaisons post-hoc seraient "
                "nécessaires."
            )

        if p_value is not None and not np.isnan(p_value):
            paires = [
                ("Test", methode),
                ("Statistique", float(statistique)),
                ("p-value", format_p(p_value)),
                ("Nombre de groupes", len(listes)),
                ("Effectif du test", effectif_test),
            ]
            afficher_tableau(table_indicateurs(paires))

            if p_value < 0.05:
                if len(listes) == 2:
                    st.success(
                        f"{libelle_signif} statistiquement significative "
                        "au seuil de 5 %."
                    )
                else:
                    st.success(
                        f"{libelle_signif} statistiquement significative "
                        "entre au moins deux groupes au seuil de 5 %."
                    )
                if precision:
                    st.caption(precision)
            else:
                if len(listes) == 2:
                    st.info(
                        f"{libelle_non_signif} statistiquement significative "
                        "au seuil de 5 %."
                    )
                else:
                    st.info(
                        f"{libelle_non_signif} statistiquement significative "
                        "entre les groupes au seuil de 5 %."
                    )
        elif p_value is not None:
            st.warning(
                "Le test ne peut pas être interprété (valeurs identiques "
                "ou variance nulle dans les groupes)."
            )

    st.bar_chart(statistiques_groupes.set_index(var_quali)["Moyenne"])


variables_disponibles = list(df_nettoye.columns)

colonne1, colonne2 = st.columns(2)

with colonne1:
    variable1 = st.selectbox(
        "Variable 1", variables_disponibles, key="variable_bivariee_1"
    )

with colonne2:
    variable2 = st.selectbox(
        "Variable 2",
        variables_disponibles,
        index=1 if len(variables_disponibles) > 1 else 0,
        key="variable_bivariee_2",
    )

if variable1 == variable2:
    st.warning("Veuillez sélectionner deux variables différentes.")

else:
    type1 = dictionnaire_modifie.loc[
        dictionnaire_modifie["Variable"] == variable1, "Type d'analyse"
    ].iloc[0]

    type2 = dictionnaire_modifie.loc[
        dictionnaire_modifie["Variable"] == variable2, "Type d'analyse"
    ].iloc[0]

    st.info(f"**{variable1}** ({type1}) × **{variable2}** ({type2})")

    # ========================================================
    # QUALITATIVE × QUALITATIVE
    # ========================================================
    if type1 in types_qualitatifs and type2 in types_qualitatifs:
        st.markdown("### 9.1 Qualitative × Qualitative")

        donnees = df_nettoye[[variable1, variable2]].dropna()

        if len(donnees) == 0:
            st.warning("Aucune donnée exploitable.")

        else:
            tableau = pd.crosstab(
                donnees[variable1], donnees[variable2], margins=True
            )

            st.markdown("#### Tableau croisé")
            afficher_tableau(tableau, hide_index=False)

            tableau_pourcentage = (
                pd.crosstab(
                    donnees[variable1], donnees[variable2], normalize="index"
                )
                * 100
            )

            st.markdown("#### Pourcentages par ligne")
            afficher_tableau(tableau_pourcentage.round(2), hide_index=False)

            table_test = pd.crosstab(donnees[variable1], donnees[variable2])

            if table_test.shape[0] >= 2 and table_test.shape[1] >= 2:
                chi2, p_value, ddl, attendus = chi2_contingency(table_test)

                proportion_faible = (attendus < 5).sum() / attendus.size

                st.markdown("#### Test d'association")

                afficher_tableau(
                    table_indicateurs(
                        [
                            ("Khi²", float(chi2)),
                            ("Degrés de liberté", int(ddl)),
                            ("p-value", format_p(p_value)),
                            (
                                "% d'effectifs attendus < 5",
                                float(proportion_faible * 100),
                            ),
                        ],
                        decimales=4,
                    )
                )

                if table_test.shape == (2, 2) and proportion_faible > 0:
                    odds_ratio, fisher_p = fisher_exact(table_test)

                    st.markdown("#### Test exact de Fisher")

                    afficher_tableau(
                        table_indicateurs(
                            [
                                ("Odds ratio", float(odds_ratio)),
                                ("p-value", format_p(fisher_p)),
                            ]
                        )
                    )

                    if fisher_p < 0.05:
                        st.success(
                            "Le test exact de Fisher détecte une association "
                            "statistiquement significative au seuil de 5 %."
                        )
                    else:
                        st.info(
                            "Le test exact de Fisher ne détecte pas "
                            "d'association statistiquement significative "
                            "au seuil de 5 %."
                        )

                else:
                    if proportion_faible > 0.20:
                        st.warning(
                            "Plus de 20 % des effectifs attendus sont "
                            "inférieurs à 5. Le résultat du Khi² doit être "
                            "interprété avec prudence."
                        )
                    elif p_value < 0.05:
                        st.success(
                            "Le test du Khi² détecte une association "
                            "statistiquement significative au seuil de 5 %."
                        )
                    else:
                        st.info(
                            "Le test du Khi² ne détecte pas d'association "
                            "statistiquement significative au seuil de 5 %."
                        )
            else:
                st.warning(
                    "Le tableau croisé doit comporter au moins 2 lignes "
                    "et 2 colonnes pour réaliser un test."
                )

    # ========================================================
    # QUANTITATIVE × QUANTITATIVE
    # ========================================================
    elif type1 == "Quantitative" and type2 == "Quantitative":
        st.markdown("### 9.2 Quantitative × Quantitative")

        donnees = df_nettoye[[variable1, variable2]].copy()
        donnees[variable1] = pd.to_numeric(donnees[variable1], errors="coerce")
        donnees[variable2] = pd.to_numeric(donnees[variable2], errors="coerce")
        donnees = donnees.dropna()

        if len(donnees) < 3:
            st.warning("Pas suffisamment de données pour réaliser une corrélation.")

        elif donnees[variable1].nunique() < 2 or donnees[variable2].nunique() < 2:
            st.warning(
                "Une des deux variables est constante : la corrélation "
                "ne peut pas être calculée."
            )

        else:
            methode = st.selectbox(
                "Méthode de corrélation",
                ["Pearson", "Spearman"],
                key="methode_correlation",
            )

            x = donnees[variable1]
            y = donnees[variable2]

            if methode == "Pearson":
                coefficient, p_value = pearsonr(x, y)
            else:
                coefficient, p_value = spearmanr(x, y)

            afficher_tableau(
                table_indicateurs(
                    [
                        ("Coefficient", float(coefficient)),
                        ("p-value", format_p(p_value)),
                        ("Effectif", len(donnees)),
                    ]
                )
            )

            st.scatter_chart(donnees, x=variable1, y=variable2)

            if p_value < 0.05:
                st.success(
                    f"La corrélation de {methode} est statistiquement "
                    "significative au seuil de 5 %."
                )
            else:
                st.info(
                    f"La corrélation de {methode} n'est pas statistiquement "
                    "significative au seuil de 5 %."
                )

    # ========================================================
    # QUALITATIVE × QUANTITATIVE
    # ========================================================
    elif type1 in types_qualitatifs and type2 == "Quantitative":
        comparer_groupes(variable1, variable2, "### 9.3 Qualitative × Quantitative")

    # ========================================================
    # QUANTITATIVE × QUALITATIVE
    # ========================================================
    elif type1 == "Quantitative" and type2 in types_qualitatifs:
        comparer_groupes(variable2, variable1, "### 9.4 Quantitative × Qualitative")

    else:
        st.warning(
            "Cette combinaison de types de variables n'est pas encore "
            "prise en charge."
        )

st.success("Analyse descriptive, nettoyage et analyse bivariée disponibles.")

st.info(
    "Une association ou une différence statistiquement significative ne "
    "constitue pas à elle seule une preuve de causalité."
)

# ============================================================
# 10. CONTRÔLE QUALITÉ DES DONNÉES
# ============================================================

st.subheader("10. Contrôle qualité des données")

st.write(
    "Ce module recherche des valeurs potentiellement problématiques "
    "avant l'analyse statistique. Les réponses comme « aucune », "
    "« RAS », « néant » ou « rien » sont signalées pour vérification "
    "et ne sont pas supprimées automatiquement."
)

# ------------------------------------------------------------
# 10.1 Résumé général
# ------------------------------------------------------------

st.markdown("### 10.1 Résumé du contrôle")

k1, k2, k3, k4 = st.columns(4)

with k1:
    st.metric("Lignes", len(df_nettoye))

with k2:
    st.metric("Variables", len(df_nettoye.columns))

with k3:
    st.metric("Cellules manquantes", int(df_nettoye.isna().sum().sum()))

with k4:
    st.metric("Doublons", int(df_nettoye.duplicated().sum()))

# ------------------------------------------------------------
# 10.2 Réponses à vérifier
# ------------------------------------------------------------

st.markdown("### 10.2 Réponses textuelles à vérifier")

valeurs_a_verifier = {
    normaliser_cle(v)
    for v in [
        "aucun",
        "aucune",
        "néant",
        "neant",
        "ras",
        "r.a.s.",
        "rien",
        "n'importe quel",
        "ouvert à tous niveaux",
    ]
}

colonnes_textuelles_controle = df_nettoye.select_dtypes(
    include=["object", "string"]
).columns

reponses_suspectes = []

for col in colonnes_textuelles_controle:
    serie = df_nettoye[col].dropna().astype(str).map(normaliser_cle)
    suspectes = serie[serie.isin(valeurs_a_verifier)]

    for reponse, effectif in suspectes.value_counts().items():
        reponses_suspectes.append(
            {"Variable": col, "Réponse": reponse, "Effectif": int(effectif)}
        )

if reponses_suspectes:
    st.warning(
        "Certaines réponses nécessitent une vérification humaine. "
        "Elles n'ont pas été supprimées."
    )
    afficher_tableau(pd.DataFrame(reponses_suspectes))
else:
    st.success("Aucune réponse textuelle suspecte détectée.")

# ------------------------------------------------------------
# 10.3 Modalités très proches
# ------------------------------------------------------------

st.markdown("### 10.3 Recherche de modalités potentiellement différentes")

st.write(
    "Cette vérification recherche des différences de casse, d'accents "
    "ou d'espaces pouvant créer artificiellement plusieurs modalités."
)

modalites_proches = []

for col in colonnes_textuelles_controle:
    valeurs_originales = df_nettoye[col].dropna().astype(str).unique()

    groupes_formes = {}
    for valeur in valeurs_originales:
        groupes_formes.setdefault(normaliser_cle(valeur), []).append(valeur)

    for _, formes in groupes_formes.items():
        formes_uniques = list(dict.fromkeys(formes))

        if len(formes_uniques) > 1:
            modalites_proches.append(
                {
                    "Variable": col,
                    "Formes détectées": " | ".join(formes_uniques),
                }
            )

if modalites_proches:
    st.warning(
        "Des modalités semblent différentes uniquement à cause de la "
        "casse, des accents ou des espaces."
    )
    afficher_tableau(pd.DataFrame(modalites_proches))
else:
    st.success("Aucune modalité manifestement similaire détectée.")

# ------------------------------------------------------------
# 10.4 Valeurs quantitatives extrêmes
# ------------------------------------------------------------

st.markdown("### 10.4 Valeurs quantitatives extrêmes")

st.write(
    "Une valeur extrême n'est pas automatiquement une erreur. "
    "Elle est simplement signalée pour vérification."
)

variables_quant_controle = dictionnaire_modifie.loc[
    dictionnaire_modifie["Type d'analyse"] == "Quantitative", "Variable"
].tolist()

valeurs_extremes = []

for col in variables_quant_controle:
    serie = pd.to_numeric(df_nettoye[col], errors="coerce").dropna()

    if len(serie) < 4:
        continue

    quartile1 = serie.quantile(0.25)
    quartile3 = serie.quantile(0.75)
    iqr = quartile3 - quartile1

    borne_inf = quartile1 - 1.5 * iqr
    borne_sup = quartile3 + 1.5 * iqr

    nombre_extremes = int(((serie < borne_inf) | (serie > borne_sup)).sum())

    if nombre_extremes > 0:
        valeurs_extremes.append(
            {
                "Variable": col,
                "Valeurs extrêmes": nombre_extremes,
                "Borne inférieure": round(borne_inf, 2),
                "Borne supérieure": round(borne_sup, 2),
            }
        )

if valeurs_extremes:
    st.warning(
        "Certaines valeurs sont statistiquement extrêmes selon la règle de l'IQR."
    )
    afficher_tableau(pd.DataFrame(valeurs_extremes))
    st.caption(
        "Attention : une valeur extrême n'est pas nécessairement une erreur de saisie."
    )
else:
    st.success("Aucune valeur extrême détectée selon la règle de l'IQR.")

# ------------------------------------------------------------
# 10.5 Contrôle des variables quantitatives
# (section reconstituée : le fichier d'origine était coupé ici)
# ------------------------------------------------------------

st.markdown("### 10.5 Contrôle des variables quantitatives")

controles_quant = []

for col in variables_quant_controle:
    brute = df_nettoye[col]
    numerique = pd.to_numeric(brute, errors="coerce")
    valides = numerique.dropna()

    non_numeriques = int((brute.notna() & numerique.isna()).sum())
    negatives = int((valides < 0).sum())
    constante = len(valides) > 0 and valides.nunique() == 1

    if non_numeriques > 0 or negatives > 0 or constante:
        controles_quant.append(
            {
                "Variable": col,
                "Valeurs non numériques": non_numeriques,
                "Valeurs négatives": negatives,
                "Variable constante": "Oui" if constante else "Non",
            }
        )

if controles_quant:
    st.warning(
        "Certaines variables quantitatives méritent une vérification "
        "(valeurs négatives éventuellement légitimes selon la variable)."
    )
    afficher_tableau(pd.DataFrame(controles_quant))
else:
    st.success("Aucun problème évident détecté sur les variables quantitatives.")

# ============================================================
# 11. CODIFICATION DES QUESTIONS OUVERTES
# ============================================================

st.subheader("11. Codification des questions ouvertes")

st.write(
    "L'application analyse les réponses ouvertes, recherche "
    "les réponses lexicalement similaires et propose des "
    "regroupements thématiques. Les propositions doivent "
    "être vérifiées et validées par l'utilisateur."
)

REPONSES_NON_INFORMATIVES = {
    normaliser_cle(v)
    for v in [
        "aucun",
        "aucune",
        "néant",
        "neant",
        "ras",
        "r.a.s.",
        "rien",
        "non",
        "aucune idée",
        "je ne sais pas",
        "ne sait pas",
    ]
}


def normaliser_texte_auto(texte):
    return re.sub(r"\s+", " ", str(texte).strip().lower())


def section_codification(variables_ouvertes):
    variable_ouverte = st.selectbox(
        "Sélectionnez une question ouverte",
        variables_ouvertes,
        key="variable_question_ouverte_auto",
    )

    serie_ouverte = df_nettoye[variable_ouverte].dropna().astype(str).str.strip()
    serie_ouverte = serie_ouverte[serie_ouverte != ""]

    st.write(f"**Nombre de réponses exploitables : {len(serie_ouverte)}**")

    # ---------------- 11.1 Réponses originales ----------------
    st.markdown("### 11.1 Réponses originales")

    afficher_tableau(pd.DataFrame({"Réponse originale": serie_ouverte.values}))

    # ---------------- 11.2 Préparation ----------------
    st.markdown("### 11.2 Préparation automatique du texte")

    codification = pd.DataFrame({"Réponse originale": serie_ouverte.values})
    codification["Réponse normalisée"] = codification["Réponse originale"].apply(
        normaliser_texte_auto
    )

    # ---------------- 11.3 Réponses à vérifier ----------------
    st.markdown("### 11.3 Réponses à vérifier")

    codification["À vérifier"] = codification["Réponse normalisée"].apply(
        lambda t: normaliser_cle(t) in REPONSES_NON_INFORMATIVES
    )

    nombre_non_informatives = int(codification["À vérifier"].sum())

    if nombre_non_informatives > 0:
        st.warning(
            f"{nombre_non_informatives} réponse(s) ont été identifiées "
            "comme potentiellement non informatives. Elles sont exclues "
            "de la proposition automatique de thèmes."
        )
        afficher_tableau(codification[codification["À vérifier"]])
    else:
        st.success(
            "Aucune réponse manifestement non informative n'a été détectée."
        )

    # ---------------- 11.4 Proposition de thèmes ----------------
    st.markdown("### 11.4 Proposition automatique de thèmes")

    if len(codification) < 4:
        st.warning(
            "Il faut au moins 4 réponses pour proposer automatiquement "
            "des regroupements."
        )
        return

    donnees_clustering = codification[~codification["À vérifier"]].copy()

    if len(donnees_clustering) < 3:
        st.warning(
            "Il ne reste pas suffisamment de réponses informatives pour "
            "effectuer une proposition automatique."
        )
        return

    textes = donnees_clustering["Réponse normalisée"].tolist()
    nombre_distincts = len(set(textes))

    if nombre_distincts < 3:
        st.warning(
            "Les réponses informatives sont trop peu variées "
            "pour proposer des regroupements."
        )
        return

    nombre_max_themes = min(8, nombre_distincts)

    nombre_themes = st.slider(
        "Nombre de thèmes à proposer",
        min_value=2,
        max_value=nombre_max_themes,
        value=min(4, nombre_max_themes),
        key="nombre_themes_auto",
    )

    try:
        try:
            vectoriseur = TfidfVectorizer(
                lowercase=True,
                strip_accents="unicode",
                min_df=1,
                max_df=0.95,
                ngram_range=(1, 2),
            )
            matrice_tfidf = vectoriseur.fit_transform(textes)
        except ValueError:
            # max_df trop strict pour un petit corpus : on le relâche.
            vectoriseur = TfidfVectorizer(
                lowercase=True,
                strip_accents="unicode",
                min_df=1,
                max_df=1.0,
                ngram_range=(1, 2),
            )
            matrice_tfidf = vectoriseur.fit_transform(textes)

        if matrice_tfidf.shape[1] < 2:
            st.warning(
                "Les réponses sont trop similaires ou trop courtes pour "
                "effectuer un regroupement automatique."
            )
            return

        nombre_clusters = min(nombre_themes, nombre_distincts)

        modele = KMeans(n_clusters=nombre_clusters, random_state=42, n_init=10)
        labels = modele.fit_predict(matrice_tfidf)

        donnees_clustering["Groupe automatique"] = labels + 1

        # ---------- Mots représentatifs de chaque groupe ----------
        termes = np.array(vectoriseur.get_feature_names_out())
        centres = modele.cluster_centers_

        noms_themes = {}

        for numero in range(nombre_clusters):
            indices = centres[numero].argsort()[::-1]

            mots = []
            for indice in indices:
                if centres[numero][indice] <= 0:
                    break
                mot = termes[indice]
                if mot not in mots:
                    mots.append(mot)
                if len(mots) >= 3:
                    break

            # Le numéro de groupe garantit que deux thèmes n'ont
            # jamais le même nom proposé.
            noms_themes[numero + 1] = (
                f"Thème {numero + 1} : " + " / ".join(mots)
                if mots
                else f"Thème {numero + 1}"
            )

        donnees_clustering["Thème proposé"] = donnees_clustering[
            "Groupe automatique"
        ].map(noms_themes)

        # ---------------- 11.5 Thèmes proposés ----------------
        st.markdown("### 11.5 Thèmes proposés")

        total_informatif = len(donnees_clustering)

        resume_themes = (
            donnees_clustering["Thème proposé"].value_counts().reset_index()
        )
        resume_themes.columns = ["Thème proposé", "Effectif"]
        resume_themes["Pourcentage"] = (
            resume_themes["Effectif"] / total_informatif * 100
        ).round(2)

        afficher_tableau(resume_themes)

        st.caption(
            f"Pourcentages calculés sur {total_informatif} réponse(s) "
            f"informative(s) ({nombre_non_informatives} réponse(s) "
            "non informative(s) exclue(s))."
        )

        st.info(
            "Les thèmes proposés sont basés sur la similarité lexicale des "
            "réponses. Ils ne constituent pas une interprétation "
            "automatique définitive du sens des réponses."
        )

        # ---------------- 11.6 Réponses regroupées ----------------
        st.markdown("### 11.6 Réponses regroupées")

        for groupe, nom_propose in noms_themes.items():
            reponses_theme = donnees_clustering.loc[
                donnees_clustering["Groupe automatique"] == groupe,
                ["Réponse originale"],
            ]
            st.markdown(f"#### {nom_propose} ({len(reponses_theme)} réponse(s))")
            afficher_tableau(reponses_theme)

        # ---------------- 11.7 Validation des thèmes ----------------
        st.markdown("### 11.7 Validation des thèmes proposés")

        st.write(
            "Vous pouvez remplacer les noms proposés par des intitulés "
            "plus pertinents pour votre étude. Donner le même nom à deux "
            "thèmes les fusionne dans les résultats."
        )

        themes_valides = {}

        for groupe, nom_propose in noms_themes.items():
            themes_valides[groupe] = st.text_input(
                f"Nom du thème {groupe}",
                value=nom_propose,
                key=(
                    f"nom_theme_valide_{variable_ouverte}_"
                    f"{nombre_clusters}_{groupe}"
                ),
            )

        donnees_clustering["Thème validé"] = donnees_clustering[
            "Groupe automatique"
        ].map(themes_valides)

        # ---------------- 11.8 Codification proposée ----------------
        st.markdown("### 11.8 Codification proposée")

        resultat_final = donnees_clustering[
            [
                "Réponse originale",
                "Réponse normalisée",
                "Thème proposé",
                "Thème validé",
            ]
        ]

        afficher_tableau(resultat_final)

        # ---------------- 11.9 Résultats statistiques ----------------
        st.markdown("### 11.9 Résultats statistiques")

        statistiques_themes = (
            donnees_clustering["Thème validé"].value_counts().reset_index()
        )
        statistiques_themes.columns = ["Thème", "Effectif"]
        statistiques_themes["Pourcentage"] = (
            statistiques_themes["Effectif"] / total_informatif * 100
        ).round(2)

        afficher_tableau(statistiques_themes)

        st.bar_chart(statistiques_themes.set_index("Thème")["Effectif"])

        # ---------------- 11.10 Export Excel ----------------
        st.markdown("### 11.10 Export de la codification")

        try:
            buffer_auto = io.BytesIO()

            with pd.ExcelWriter(buffer_auto, engine="openpyxl") as writer:
                resultat_final.to_excel(
                    writer, index=False, sheet_name="Codification"
                )
                statistiques_themes.to_excel(
                    writer, index=False, sheet_name="Résultats"
                )
                resume_themes.to_excel(
                    writer, index=False, sheet_name="Propositions"
                )
                if nombre_non_informatives > 0:
                    codification[codification["À vérifier"]].to_excel(
                        writer, index=False, sheet_name="À vérifier"
                    )

            st.download_button(
                label="Télécharger la codification automatique Excel",
                data=buffer_auto.getvalue(),
                file_name="codification_automatique.xlsx",
                mime=MIME_XLSX,
                key="telecharger_codification_auto",
            )

        except Exception as e:
            st.error(f"Erreur lors de la préparation du fichier : {e}")

    except Exception as e:
        st.error(
            f"La proposition automatique des thèmes a rencontré une erreur : {e}"
        )


variables_ouvertes = dictionnaire_modifie.loc[
    dictionnaire_modifie["Type de question"] == "Question ouverte", "Variable"
].tolist()
variables_ouvertes = [v for v in variables_ouvertes if v in df_nettoye.columns]

if not variables_ouvertes:
    st.info(
        "Aucune question ouverte n'est actuellement identifiée dans le "
        "dictionnaire. Pour activer ce module, passez « Type de question » "
        "à « Question ouverte » pour une variable, dans le tableau de la "
        "section 4."
    )
elif not SKLEARN_DISPONIBLE:
    st.error(
        "Le module scikit-learn n'est pas installé : la codification "
        "automatique est indisponible. Installez-le avec "
        "« pip install scikit-learn »."
    )
else:
    section_codification(variables_ouvertes)

# ============================================================
# 12. GÉNÉRATION AUTOMATIQUE DE CONSTATS
# ============================================================

st.subheader("12. Génération automatique de constats")

st.write(
    "Ce module transforme certains résultats statistiques "
    "en constats descriptifs simples. Les constats générés "
    "doivent être relus avant leur utilisation dans un rapport."
)


def phrase_modalite_principale(variable, frequences, total, unite):
    """Constat sur la modalité la plus fréquente, avec gestion des égalités."""
    effectif_max = int(frequences.iloc[0])
    en_tete = [str(m) for m in frequences[frequences == effectif_max].index]
    pourcentage = effectif_max / total * 100

    if len(en_tete) == 1:
        return (
            f"Pour la variable « {variable} », la modalité « {en_tete[0]} » "
            f"est la plus fréquente, avec {effectif_max} réponse(s), soit "
            f"{pourcentage:.1f} % {unite} (n = {total})."
        )

    liste = " », « ".join(en_tete)
    return (
        f"Pour la variable « {variable} », les modalités « {liste} » sont "
        f"à égalité en tête, avec {effectif_max} réponse(s) chacune, soit "
        f"{pourcentage:.1f} % {unite} (n = {total})."
    )


def afficher_liste_constats(liste):
    if liste:
        afficher_tableau(pd.DataFrame(liste)[["Variable", "Constat"]])
    else:
        st.info("Aucun constat pour cette catégorie.")


# ------------------------------------------------------------
# 12.1 Variables qualitatives
# ------------------------------------------------------------

st.markdown("### 12.1 Constats descriptifs — variables qualitatives")

constats_quali = []

for _, ligne in dictionnaire_modifie.iterrows():
    variable = ligne["Variable"]
    type_analyse = ligne["Type d'analyse"]
    type_question = ligne["Type de question"]

    if variable not in df_nettoye.columns or type_analyse not in types_qualitatifs:
        continue

    if type_question == "Question fermée":
        serie = df_nettoye[variable].dropna()

        if len(serie) == 0:
            continue

        constat = phrase_modalite_principale(
            variable, serie.value_counts(), len(serie), "des réponses valides"
        )

    elif type_question == "Réponses multiples":
        serie = df_nettoye[variable].dropna().astype(str)

        citations = [
            morceau.strip()
            for valeur in serie
            for morceau in re.split(r"[,;|]", valeur)
            if morceau.strip()
        ]

        if not citations or len(serie) == 0:
            continue

        constat = phrase_modalite_principale(
            variable,
            pd.Series(citations).value_counts(),
            len(serie),
            "des répondants (plusieurs réponses possibles)",
        )

    else:
        continue

    constats_quali.append(
        {
            "Sous-section": "12.1",
            "Type": "Descriptif",
            "Variable": variable,
            "Constat": constat,
        }
    )

afficher_liste_constats(constats_quali)

# ------------------------------------------------------------
# 12.2 Variables quantitatives
# ------------------------------------------------------------

st.markdown("### 12.2 Constats descriptifs — variables quantitatives")

constats_quanti = []

for _, ligne in dictionnaire_modifie.iterrows():
    variable = ligne["Variable"]

    if ligne["Type d'analyse"] != "Quantitative" or variable not in df_nettoye.columns:
        continue

    serie = pd.to_numeric(df_nettoye[variable], errors="coerce").dropna()

    if len(serie) == 0:
        continue

    constat = (
        f"Pour « {variable} », la moyenne est de {serie.mean():.2f}, "
        f"la médiane de {serie.median():.2f}, avec des valeurs comprises "
        f"entre {serie.min():.2f} et {serie.max():.2f} (n = {len(serie)})."
    )

    constats_quanti.append(
        {
            "Sous-section": "12.2",
            "Type": "Descriptif",
            "Variable": variable,
            "Constat": constat,
        }
    )

afficher_liste_constats(constats_quanti)

# ------------------------------------------------------------
# 12.3 Règles d'interprétation
# ------------------------------------------------------------

st.markdown("### 12.3 Règles d'interprétation statistique")

st.info(
    "Une p-value inférieure à 0,05 indique que le résultat "
    "est statistiquement significatif selon le seuil retenu. "
    "Elle ne démontre pas une relation causale."
)

# ------------------------------------------------------------
# 12.4 Affichage de l'ensemble des constats
# ------------------------------------------------------------

st.markdown("### 12.4 Constats générés")

constats = constats_quali + constats_quanti

if constats:
    tableau_constats = pd.DataFrame(constats)
    afficher_tableau(tableau_constats)
else:
    st.info("Aucun constat automatique n'a pu être généré.")

# ------------------------------------------------------------
# 12.5 Téléchargement
# ------------------------------------------------------------

st.markdown("### 12.5 Téléchargement des constats")

if constats:
    try:
        buffer_constats = io.BytesIO()

        with pd.ExcelWriter(buffer_constats, engine="openpyxl") as writer:
            tableau_constats.to_excel(writer, index=False, sheet_name="Constats")

        st.download_button(
            label="Télécharger les constats Excel",
            data=buffer_constats.getvalue(),
            file_name="constats_automatiques.xlsx",
            mime=MIME_XLSX,
            key="telecharger_constats",
        )

    except Exception as e:
        st.error(f"Erreur lors de la préparation du fichier : {e}")
else:
    st.caption("Aucun fichier à télécharger.")

# ------------------------------------------------------------
# 12.6 Rappel méthodologique
# ------------------------------------------------------------

st.markdown("### 12.6 Rappel méthodologique")

st.warning(
    "Les constats automatiques sont des formulations "
    "descriptives basées sur les résultats calculés. "
    "Ils ne remplacent pas l'interprétation scientifique "
    "et ne doivent pas être utilisés pour affirmer une causalité."
)
