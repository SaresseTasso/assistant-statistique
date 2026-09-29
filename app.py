
import streamlit as st
import pandas as pd
import numpy as np
import re

from scipy.stats import (
    chi2_contingency,
    fisher_exact,
    pearsonr,
    spearmanr,
    ttest_ind,
    mannwhitneyu,
    f_oneway,
    kruskal
)

# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Assistant statistique",
    page_icon="📊",
    layout="wide"
)

st.title("Assistant statistique")

st.write(
    "Analysez vos fichiers Excel ou CSV : diagnostic, "
    "dictionnaire, nettoyage, statistiques descriptives "
    "et analyses bivariées."
)

# ============================================================
# IMPORTATION DU FICHIER
# ============================================================

fichier = st.file_uploader(
    "Choisissez votre fichier de données",
    type=["xlsx", "csv"]
)

if fichier is None:
    st.info("Veuillez importer un fichier Excel ou CSV.")
    st.stop()

try:
    if fichier.name.lower().endswith(".xlsx"):
        df = pd.read_excel(fichier)
    else:
        df = pd.read_csv(fichier)

except Exception as e:
    st.error(f"Erreur lors de la lecture du fichier : {e}")
    st.stop()

st.success(f"Fichier chargé : {fichier.name}")

# ============================================================
# INITIALISATION DES DONNEES NETTOYEES
# ============================================================

# On conserve toujours une copie des données originales.
# Toutes les corrections sont appliquées sur df_nettoye.

if (
    "nom_fichier_actuel" not in st.session_state
    or
    st.session_state["nom_fichier_actuel"] != fichier.name
):

    st.session_state["nom_fichier_actuel"] = fichier.name
    st.session_state["df_nettoye"] = df.copy()
    st.session_state["dictionnaire_modifie"] = None

df_nettoye = st.session_state["df_nettoye"]

# ============================================================
# INFORMATIONS GENERALES
# ============================================================

st.subheader("1. Informations générales")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("Lignes", df.shape[0])

with col2:
    st.metric("Variables", df.shape[1])

with col3:
    st.metric(
        "Doublons",
        int(df.duplicated().sum())
    )

with col4:
    st.metric(
        "Cellules manquantes",
        int(df.isna().sum().sum())
    )

# ============================================================
# APERCU DES DONNEES
# ============================================================

st.subheader("2. Aperçu des données")

st.dataframe(
    df.head(20),
    use_container_width=True,
    hide_index=True
)

# ============================================================
# DIAGNOSTIC DES VARIABLES
# ============================================================

st.subheader("3. Diagnostic des variables")

diagnostic = pd.DataFrame({
    "Variable": df.columns,

    "Type Python": [
        str(df[col].dtype)
        for col in df.columns
    ],

    "Valeurs manquantes": [
        int(df[col].isna().sum())
        for col in df.columns
    ],

    "% manquant": [
        round(
            df[col].isna().mean() * 100,
            2
        )
        for col in df.columns
    ],

    "Valeurs uniques": [
        int(df[col].nunique(dropna=True))
        for col in df.columns
    ]
})

st.dataframe(
    diagnostic,
    use_container_width=True,
    hide_index=True
)

if df.duplicated().sum() > 0:
    st.warning(
        f"{df.duplicated().sum()} doublon(s) détecté(s)."
    )
else:
    st.success("Aucun doublon détecté.")

# ============================================================
# FONCTION DE PROPOSITION DU TYPE
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
    "À vérifier"
]

types_questions = [
    "Question fermée",
    "Question ouverte",
    "Réponses multiples",
    "Non applicable"
]

# ============================================================
# DICTIONNAIRE DES VARIABLES
# ============================================================

dictionnaire = pd.DataFrame({
    "Variable": df.columns,

    "Type Python": [
        str(df[col].dtype)
        for col in df.columns
    ],

    "Nombre de modalités": [
        int(df[col].nunique(dropna=True))
        for col in df.columns
    ],

    "Valeurs manquantes": [
        int(df[col].isna().sum())
        for col in df.columns
    ]
})

dictionnaire["Type d'analyse"] = [
    proposer_type(col)
    for col in df.columns
]

dictionnaire["Type de question"] = [
    "Non applicable"
    if proposer_type(col) in [
        "Quantitative",
        "Date",
        "Identifiant"
    ]
    else "Question fermée"
    for col in df.columns
]

colonnes_requises = [
    "Variable",
    "Type Python",
    "Nombre de modalités",
    "Valeurs manquantes",
    "Type d'analyse",
    "Type de question"
]

# Si aucun dictionnaire n'existe ou si le fichier a changé,
# on recrée le dictionnaire.

if (
    st.session_state["dictionnaire_modifie"] is None
    or
    not all(
        col in st.session_state["dictionnaire_modifie"].columns
        for col in colonnes_requises
    )
    or
    len(st.session_state["dictionnaire_modifie"])
    != len(df.columns)
):

    st.session_state["dictionnaire_modifie"] = (
        dictionnaire.copy()
    )

# ============================================================
# EDITION DU DICTIONNAIRE
# ============================================================

st.subheader("4. Dictionnaire des variables")

dictionnaire_modifie = st.data_editor(
    st.session_state["dictionnaire_modifie"],

    column_config={

        "Type d'analyse":
            st.column_config.SelectboxColumn(
                "Type d'analyse",
                options=types_analyse
            ),

        "Type de question":
            st.column_config.SelectboxColumn(
                "Type de question",
                options=types_questions
            )
    },

    disabled=[
        "Variable",
        "Type Python",
        "Nombre de modalités",
        "Valeurs manquantes"
    ],

    use_container_width=True,
    hide_index=True
)

st.session_state[
    "dictionnaire_modifie"
] = dictionnaire_modifie

# ============================================================
# ANALYSE DES VARIABLES QUALITATIVES
# ============================================================

st.subheader(
    "5. Analyse des variables qualitatives"
)

types_qualitatifs = [
    "Qualitative",
    "Qualitative codée"
]

for _, ligne in dictionnaire_modifie.iterrows():

    variable = ligne["Variable"]
    type_analyse = ligne["Type d'analyse"]
    type_question = ligne["Type de question"]

    # --------------------------------------------------------
    # QUESTION FERMEE
    # --------------------------------------------------------

    if (
        type_analyse in types_qualitatifs
        and
        type_question == "Question fermée"
    ):

        st.markdown(
            f"### {variable}"
        )

        serie = df_nettoye[variable]

        valide = serie.dropna()

        if len(valide) == 0:

            st.warning(
                "Aucune réponse exploitable."
            )

            continue

        effectifs = valide.value_counts()

        pourcentages = (
            valide.value_counts(
                normalize=True
            ) * 100
        )

        resultat = pd.DataFrame({

            "Modalité":
                effectifs.index.astype(str),

            "Effectif":
                effectifs.values,

            "Pourcentage":
                pourcentages.values.round(2)
        })

        st.dataframe(
            resultat,
            use_container_width=True,
            hide_index=True
        )

        st.caption(
            f"Réponses valides : {len(valide)} | "
            f"Manquantes : {serie.isna().sum()} | "
            f"Total : {len(serie)}"
        )

        st.bar_chart(
            resultat.set_index(
                "Modalité"
            )["Effectif"]
        )

    # --------------------------------------------------------
    # QUESTION OUVERTE
    # --------------------------------------------------------

    elif (
        type_analyse in types_qualitatifs
        and
        type_question == "Question ouverte"
    ):

        st.markdown(
            f"### {variable}"
        )

        serie = (
            df_nettoye[variable]
            .dropna()
            .astype(str)
        )

        st.write(
            f"**Nombre de réponses : {len(serie)}**"
        )

        apercu = pd.DataFrame({
            "Réponses":
                serie.head(20).values
        })

        st.dataframe(
            apercu,
            use_container_width=True,
            hide_index=True
        )

        st.info(
            "Cette variable est ouverte. "
            "Elle sera traitée dans le module "
            "de codification thématique."
        )

    # --------------------------------------------------------
    # REPONSES MULTIPLES
    # --------------------------------------------------------

    elif (
        type_analyse in types_qualitatifs
        and
        type_question == "Réponses multiples"
    ):

        st.markdown(
            f"### {variable}"
        )

        serie = (
            df_nettoye[variable]
            .dropna()
            .astype(str)
        )

        reponses = []

        for valeur in serie:

            morceaux = re.split(
                r"[,;|]",
                valeur
            )

            for morceau in morceaux:

                morceau = morceau.strip()

                if morceau:
                    reponses.append(
                        morceau
                    )

        if reponses:

            freq = (
                pd.Series(reponses)
                .value_counts()
            )

            pourcentage = (
                freq / len(serie)
            ) * 100

            resultat_multiple = pd.DataFrame({

                "Réponse":
                    freq.index,

                "Nombre de citations":
                    freq.values,

                "% des répondants":
                    pourcentage.round(2).values
            })

            st.dataframe(
                resultat_multiple,
                use_container_width=True,
                hide_index=True
            )

            st.caption(
                "Les pourcentages peuvent dépasser "
                "100 % au total car un répondant peut "
                "avoir plusieurs réponses."
            )

        else:

            st.warning(
                "Aucune réponse multiple exploitable détectée."
            )

# ============================================================
# ANALYSE DES VARIABLES QUANTITATIVES
# ============================================================

st.subheader(
    "6. Analyse des variables quantitatives"
)

for _, ligne in dictionnaire_modifie.iterrows():

    variable = ligne["Variable"]

    if ligne["Type d'analyse"] == "Quantitative":

        st.markdown(
            f"### {variable}"
        )

        serie = pd.to_numeric(
            df_nettoye[variable],
            errors="coerce"
        ).dropna()

        if len(serie) == 0:

            st.warning(
                "Aucune valeur numérique exploitable."
            )

            continue

        statistiques = pd.DataFrame({

            "Indicateur": [

                "Effectif valide",
                "Valeurs manquantes",
                "Moyenne",
                "Médiane",
                "Écart-type",
                "Minimum",
                "Q1",
                "Q3",
                "Maximum"
            ],

            "Valeur": [

                len(serie),

                int(
                    df_nettoye[
                        variable
                    ].isna().sum()
                ),

                round(
                    serie.mean(),
                    2
                ),

                round(
                    serie.median(),
                    2
                ),

                round(
                    serie.std(),
                    2
                ),

                round(
                    serie.min(),
                    2
                ),

                round(
                    serie.quantile(0.25),
                    2
                ),

                round(
                    serie.quantile(0.75),
                    2
                ),

                round(
                    serie.max(),
                    2
                )
            ]
        })

        st.dataframe(
            statistiques,
            use_container_width=True,
            hide_index=True
        )

# ============================================================
# RESUME DU DIAGNOSTIC
# ============================================================

st.subheader(
    "7. Résumé du diagnostic"
)

nb_qualitatives = len(
    dictionnaire_modifie[
        dictionnaire_modifie[
            "Type d'analyse"
        ].isin(types_qualitatifs)
    ]
)

nb_quantitatives = len(
    dictionnaire_modifie[
        dictionnaire_modifie[
            "Type d'analyse"
        ] == "Quantitative"
    ]
)

nb_dates = len(
    dictionnaire_modifie[
        dictionnaire_modifie[
            "Type d'analyse"
        ] == "Date"
    ]
)

nb_identifiants = len(
    dictionnaire_modifie[
        dictionnaire_modifie[
            "Type d'analyse"
        ] == "Identifiant"
    ]
)

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric(
        "Variables qualitatives",
        nb_qualitatives
    )

with c2:
    st.metric(
        "Variables quantitatives",
        nb_quantitatives
    )

with c3:
    st.metric(
        "Variables de date",
        nb_dates
    )

with c4:
    st.metric(
        "Identifiants",
        nb_identifiants
    )

# ============================================================
# NETTOYAGE DES DONNEES
# ============================================================

st.subheader(
    "8. Nettoyage des données"
)

st.write(
    "Les données originales restent conservées. "
    "Les corrections sont appliquées uniquement "
    "sur une copie de travail."
)

# ------------------------------------------------------------
# ETAT ACTUEL
# ------------------------------------------------------------

st.markdown(
    "### 8.1 État actuel des données"
)

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric(
        "Lignes",
        df_nettoye.shape[0]
    )

with c2:
    st.metric(
        "Variables",
        df_nettoye.shape[1]
    )

with c3:
    st.metric(
        "Cellules manquantes",
        int(
            df_nettoye.isna()
            .sum()
            .sum()
        )
    )

with c4:
    st.metric(
        "Doublons",
        int(
            df_nettoye.duplicated()
            .sum()
        )
    )

# ------------------------------------------------------------
# VALEURS MANQUANTES
# ------------------------------------------------------------

st.markdown(
    "### 8.2 Valeurs manquantes"
)

manquants = pd.DataFrame({

    "Variable":
        df_nettoye.columns,

    "Valeurs manquantes": [
        int(
            df_nettoye[col]
            .isna()
            .sum()
        )
        for col in df_nettoye.columns
    ],

    "% manquant": [
        round(
            df_nettoye[col]
            .isna()
            .mean() * 100,
            2
        )
        for col in df_nettoye.columns
    ]
})

manquants = manquants[
    manquants[
        "Valeurs manquantes"
    ] > 0
]

if len(manquants) > 0:

    st.dataframe(
        manquants,
        use_container_width=True,
        hide_index=True
    )

else:

    st.success(
        "Aucune valeur manquante détectée."
    )

# ------------------------------------------------------------
# VALEURS TEXTUELLES REPRESENTANT DES MANQUANTS
# ------------------------------------------------------------

st.markdown(
    "### 8.3 Valeurs utilisées comme absence de réponse"
)

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
    "-"
]

colonnes_textuelles = (
    df_nettoye
    .select_dtypes(
        include=["object", "string"]
    )
    .columns
)

resultats_faux_manquants = []

for col in colonnes_textuelles:

    serie = (
        df_nettoye[col]
        .astype("string")
        .str.strip()
        .str.lower()
    )

    nombre = serie.isin(
        valeurs_manquantes_textuelles
    ).sum()

    if nombre > 0:

        resultats_faux_manquants.append({

            "Variable":
                col,

            "Valeurs détectées":
                int(nombre)
        })

if resultats_faux_manquants:

    faux_manquants = pd.DataFrame(
        resultats_faux_manquants
    )

    st.dataframe(
        faux_manquants,
        use_container_width=True,
        hide_index=True
    )

    if st.button(
        "Convertir en valeurs manquantes",
        key="convertir_faux_manquants"
    ):

        df_nettoye = df_nettoye.copy()

        for col in colonnes_textuelles:

            serie = (
                df_nettoye[col]
                .astype("string")
                .str.strip()
                .str.lower()
            )

            masque = serie.isin(
                valeurs_manquantes_textuelles
            )

            df_nettoye.loc[
                masque,
                col
            ] = np.nan

        st.session_state[
            "df_nettoye"
        ] = df_nettoye

        st.success(
            "Conversion effectuée."
        )

        st.rerun()

else:

    st.info(
        "Aucune valeur textuelle évidente "
        "représentant une absence de réponse n'a été détectée."
    )

# ------------------------------------------------------------
# ESPACES INUTILES
# ------------------------------------------------------------

st.markdown(
    "### 8.4 Espaces inutiles"
)

espaces_detectes = []

for col in colonnes_textuelles:

    serie = df_nettoye[
        col
    ].astype("string")

    masque = (
        serie.notna()
        &
        (serie != serie.str.strip())
    )

    nombre = masque.sum()

    if nombre > 0:

        espaces_detectes.append({

            "Variable":
                col,

            "Réponses concernées":
                int(nombre)
        })

if espaces_detectes:

    espaces = pd.DataFrame(
        espaces_detectes
    )

    st.dataframe(
        espaces,
        use_container_width=True,
        hide_index=True
    )

    if st.button(
        "Supprimer les espaces inutiles",
        key="supprimer_espaces"
    ):

        df_nettoye = df_nettoye.copy()

        for col in colonnes_textuelles:

            df_nettoye[col] = (
                df_nettoye[col]
                .astype("string")
                .str.strip()
            )

        st.session_state[
            "df_nettoye"
        ] = df_nettoye

        st.success(
            "Les espaces inutiles ont été supprimés."
        )

        st.rerun()

else:

    st.success(
        "Aucun espace inutile détecté."
    )

# ------------------------------------------------------------
# DOUBLONS
# ------------------------------------------------------------

st.markdown(
    "### 8.5 Doublons"
)

nombre_doublons = int(
    df_nettoye
    .duplicated()
    .sum()
)

if nombre_doublons > 0:

    st.warning(
        f"{nombre_doublons} doublon(s) exact(s) détecté(s)."
    )

    apercu_doublons = df_nettoye[
        df_nettoye.duplicated(
            keep=False
        )
    ]

    st.dataframe(
        apercu_doublons.head(50),
        use_container_width=True,
        hide_index=True
    )

    if st.button(
        "Supprimer les doublons exacts",
        key="supprimer_doublons"
    ):

        avant = len(df_nettoye)

        df_nettoye = (
            df_nettoye
            .drop_duplicates()
            .reset_index(drop=True)
        )

        apres = len(df_nettoye)

        st.session_state[
            "df_nettoye"
        ] = df_nettoye

        st.success(
            f"{avant - apres} doublon(s) supprimé(s)."
        )

        st.rerun()

else:

    st.success(
        "Aucun doublon exact détecté."
    )

# ------------------------------------------------------------
# UNIFORMISATION DES TEXTES
# ------------------------------------------------------------

st.markdown(
    "### 8.6 Uniformisation des réponses textuelles"
)

if len(colonnes_textuelles) > 0:

    variable_uniformisation = st.selectbox(
        "Variable textuelle",
        list(colonnes_textuelles),
        key="variable_uniformisation"
    )

    col_a, col_b = st.columns(2)

    with col_a:

        if st.button(
            "Supprimer les espaces",
            key="uniformiser_espaces"
        ):

            df_nettoye[
                variable_uniformisation
            ] = (
                df_nettoye[
                    variable_uniformisation
                ]
                .astype("string")
                .str.strip()
            )

            st.session_state[
                "df_nettoye"
            ] = df_nettoye

            st.success(
                "Espaces supprimés."
            )

            st.rerun()

    with col_b:

        if st.button(
            "Mettre en minuscules",
            key="mettre_minuscules"
        ):

            df_nettoye[
                variable_uniformisation
            ] = (
                df_nettoye[
                    variable_uniformisation
                ]
                .astype("string")
                .str.lower()
            )

            st.session_state[
                "df_nettoye"
            ] = df_nettoye

            st.success(
                "Réponses mises en minuscules."
            )

            st.rerun()

# ------------------------------------------------------------
# VERIFICATION NUMERIQUE
# ------------------------------------------------------------

st.markdown(
    "### 8.7 Vérification des variables quantitatives"
)

variables_quantitatives = (
    dictionnaire_modifie[
        dictionnaire_modifie[
            "Type d'analyse"
        ] == "Quantitative"
    ]["Variable"].tolist()
)

if variables_quantitatives:

    variable_numerique = st.selectbox(
        "Variable quantitative",
        variables_quantitatives,
        key="variable_numerique_nettoyage"
    )

    serie_originale = df_nettoye[
        variable_numerique
    ]

    serie_convertie = pd.to_numeric(
        serie_originale,
        errors="coerce"
    )

    valeurs_non_convertibles = (
        serie_originale.notna()
        &
        serie_convertie.isna()
    ).sum()

    st.write(
        f"Valeurs non numériques détectées : "
        f"**{valeurs_non_convertibles}**"
    )

    if valeurs_non_convertibles > 0:

        valeurs_problematiques = (
            serie_originale[
                serie_originale.notna()
                &
                serie_convertie.isna()
            ]
            .astype(str)
            .value_counts()
            .reset_index()
        )

        valeurs_problematiques.columns = [
            "Valeur",
            "Effectif"
        ]

        st.dataframe(
            valeurs_problematiques,
            use_container_width=True,
            hide_index=True
        )

        st.warning(
            "Ces valeurs seront transformées en "
            "valeurs manquantes lors de la conversion."
        )

    if st.button(
        "Convertir en numérique",
        key="convertir_numerique"
    ):

        df_nettoye[
            variable_numerique
        ] = pd.to_numeric(
            df_nettoye[
                variable_numerique
            ],
            errors="coerce"
        )

        st.session_state[
            "df_nettoye"
        ] = df_nettoye

        st.success(
            "Conversion numérique effectuée."
        )

        st.rerun()

else:

    st.info(
        "Aucune variable quantitative n'est actuellement identifiée."
    )

# ------------------------------------------------------------
# VERIFICATION DES DATES
# ------------------------------------------------------------

st.markdown(
    "### 8.8 Vérification des dates"
)

variables_dates = (
    dictionnaire_modifie[
        dictionnaire_modifie[
            "Type d'analyse"
        ] == "Date"
    ]["Variable"].tolist()
)

if variables_dates:

    variable_date = st.selectbox(
        "Variable de date",
        variables_dates,
        key="variable_date_nettoyage"
    )

    serie_date = pd.to_datetime(
        df_nettoye[
            variable_date
        ],
        errors="coerce"
    )

    valeurs_date_invalides = (
        df_nettoye[
            variable_date
        ].notna()
        &
        serie_date.isna()
    ).sum()

    st.write(
        f"Dates non reconnues : "
        f"**{valeurs_date_invalides}**"
    )

    if valeurs_date_invalides > 0:

        dates_problematiques = (
            df_nettoye[
                variable_date
            ][
                df_nettoye[
                    variable_date
                ].notna()
                &
                serie_date.isna()
            ]
            .astype(str)
            .value_counts()
            .reset_index()
        )

        dates_problematiques.columns = [
            "Valeur",
            "Effectif"
        ]

        st.dataframe(
            dates_problematiques,
            use_container_width=True,
            hide_index=True
        )

    if st.button(
        "Convertir en date",
        key="convertir_date"
    ):

        df_nettoye[
            variable_date
        ] = pd.to_datetime(
            df_nettoye[
                variable_date
            ],
            errors="coerce"
        )

        st.session_state[
            "df_nettoye"
        ] = df_nettoye

        st.success(
            "Conversion des dates effectuée."
        )

        st.rerun()

else:

    st.info(
        "Aucune variable de date n'est actuellement identifiée."
    )

# ------------------------------------------------------------
# APERCU APRES NETTOYAGE
# ------------------------------------------------------------

st.markdown(
    "### 8.9 Aperçu des données après nettoyage"
)

st.dataframe(
    df_nettoye.head(20),
    use_container_width=True,
    hide_index=True
)

# ------------------------------------------------------------
# COMPARAISON AVANT / APRES
# ------------------------------------------------------------

st.markdown(
    "### 8.10 Comparaison avant / après"
)

comparaison = pd.DataFrame({

    "Indicateur": [
        "Nombre de lignes",
        "Nombre de variables",
        "Cellules manquantes",
        "Doublons"
    ],

    "Avant nettoyage": [

        df.shape[0],

        df.shape[1],

        int(
            df.isna()
            .sum()
            .sum()
        ),

        int(
            df.duplicated()
            .sum()
        )
    ],

    "Après nettoyage": [

        df_nettoye.shape[0],

        df_nettoye.shape[1],

        int(
            df_nettoye.isna()
            .sum()
            .sum()
        ),

        int(
            df_nettoye.duplicated()
            .sum()
        )
    ]
})

st.dataframe(
    comparaison,
    use_container_width=True,
    hide_index=True
)

# ------------------------------------------------------------
# VALIDATION
# ------------------------------------------------------------

st.markdown(
    "### 8.11 Validation du nettoyage"
)

if st.button(
    "Valider les données nettoyées",
    key="valider_nettoyage"
):

    st.session_state[
        "donnees_nettoyage_valide"
    ] = True

    st.success(
        "Les données nettoyées sont validées "
        "pour la suite de l'analyse."
    )

# ------------------------------------------------------------
# REINITIALISATION
# ------------------------------------------------------------

if st.button(
    "Réinitialiser le nettoyage",
    key="reset_nettoyage"
):

    st.session_state[
        "df_nettoye"
    ] = df.copy()

    st.session_state[
        "donnees_nettoyage_valide"
    ] = False

    st.success(
        "Le nettoyage a été réinitialisé."
    )

    st.rerun()

# ============================================================
# ANALYSE BIVARIEE
# ============================================================

st.subheader(
    "9. Analyse bivariée"
)

st.write(
    "Sélectionnez deux variables pour étudier "
    "leur relation ou leur différence."
)

variables_disponibles = list(
    df_nettoye.columns
)

colonne1, colonne2 = st.columns(2)

with colonne1:

    variable1 = st.selectbox(
        "Variable 1",
        variables_disponibles,
        key="variable_bivariee_1"
    )

with colonne2:

    variable2 = st.selectbox(
        "Variable 2",
        variables_disponibles,
        index=(
            1
            if len(variables_disponibles) > 1
            else 0
        ),
        key="variable_bivariee_2"
    )

if variable1 == variable2:

    st.warning(
        "Veuillez sélectionner deux variables différentes."
    )

else:

    type1 = dictionnaire_modifie.loc[
        dictionnaire_modifie[
            "Variable"
        ] == variable1,
        "Type d'analyse"
    ].iloc[0]

    type2 = dictionnaire_modifie.loc[
        dictionnaire_modifie[
            "Variable"
        ] == variable2,
        "Type d'analyse"
    ].iloc[0]

    st.info(
        f"**{variable1}** ({type1}) × "
        f"**{variable2}** ({type2})"
    )

    # ========================================================
    # QUALITATIVE × QUALITATIVE
    # ========================================================

    if (
        type1 in types_qualitatifs
        and
        type2 in types_qualitatifs
    ):

        st.markdown(
            "### 9.1 Qualitative × Qualitative"
        )

        donnees = df_nettoye[
            [variable1, variable2]
        ].dropna()

        if len(donnees) == 0:

            st.warning(
                "Aucune donnée exploitable."
            )

        else:

            tableau = pd.crosstab(
                donnees[variable1],
                donnees[variable2],
                margins=True
            )

            st.markdown(
                "#### Tableau croisé"
            )

            st.dataframe(
                tableau,
                use_container_width=True
            )

            tableau_pourcentage = (
                pd.crosstab(
                    donnees[variable1],
                    donnees[variable2],
                    normalize="index"
                ) * 100
            )

            st.markdown(
                "#### Pourcentages par ligne"
            )

            st.dataframe(
                tableau_pourcentage.round(2),
                use_container_width=True
            )

            table_test = pd.crosstab(
                donnees[variable1],
                donnees[variable2]
            )

            if (
                table_test.shape[0] >= 2
                and
                table_test.shape[1] >= 2
            ):

                chi2, p_value, ddl, attendus = (
                    chi2_contingency(
                        table_test
                    )
                )

                proportion_faible = (
                    (attendus < 5).sum()
                    / attendus.size
                )

                st.markdown(
                    "#### Test d'association"
                )

                resultat_test = pd.DataFrame({

                    "Indicateur": [
                        "Khi²",
                        "Degrés de liberté",
                        "p-value",
                        "% d'effectifs attendus < 5"
                    ],

                    "Valeur": [

                        round(
                            chi2,
                            4
                        ),

                        ddl,

                        round(
                            p_value,
                            4
                        ),

                        round(
                            proportion_faible * 100,
                            2
                        )
                    ]
                })

                st.dataframe(
                    resultat_test,
                    use_container_width=True,
                    hide_index=True
                )

                if (
                    table_test.shape == (2, 2)
                    and
                    proportion_faible > 0
                ):

                    odds_ratio, fisher_p = (
                        fisher_exact(
                            table_test
                        )
                    )

                    st.markdown(
                        "#### Test exact de Fisher"
                    )

                    fisher_resultat = pd.DataFrame({

                        "Indicateur": [
                            "Odds ratio",
                            "p-value"
                        ],

                        "Valeur": [
                            round(
                                odds_ratio,
                                4
                            ),

                            round(
                                fisher_p,
                                4
                            )
                        ]
                    })

                    st.dataframe(
                        fisher_resultat,
                        use_container_width=True,
                        hide_index=True
                    )

                    if fisher_p < 0.05:

                        st.success(
                            "Le test exact de Fisher détecte "
                            "une association statistiquement "
                            "significative au seuil de 5 %."
                        )

                    else:

                        st.info(
                            "Le test exact de Fisher ne détecte "
                            "pas d'association statistiquement "
                            "significative au seuil de 5 %."
                        )

                else:

                    if proportion_faible > 0.20:

                        st.warning(
                            "Plus de 20 % des effectifs attendus "
                            "sont inférieurs à 5. Le résultat "
                            "du Khi² doit être interprété avec prudence."
                        )

                    elif p_value < 0.05:

                        st.success(
                            "Le test du Khi² détecte une association "
                            "statistiquement significative au seuil de 5 %."
                        )

                    else:

                        st.info(
                            "Le test du Khi² ne détecte pas "
                            "d'association statistiquement significative "
                            "au seuil de 5 %."
                        )

    # ========================================================
    # QUANTITATIVE × QUANTITATIVE
    # ========================================================

    elif (
        type1 == "Quantitative"
        and
        type2 == "Quantitative"
    ):

        st.markdown(
            "### 9.2 Quantitative × Quantitative"
        )

        donnees = df_nettoye[
            [variable1, variable2]
        ].copy()

        donnees[variable1] = pd.to_numeric(
            donnees[variable1],
            errors="coerce"
        )

        donnees[variable2] = pd.to_numeric(
            donnees[variable2],
            errors="coerce"
        )

        donnees = donnees.dropna()

        if len(donnees) < 3:

            st.warning(
                "Pas suffisamment de données "
                "pour réaliser une corrélation."
            )

        else:

            methode = st.selectbox(
                "Méthode de corrélation",
                [
                    "Pearson",
                    "Spearman"
                ],
                key="methode_correlation"
            )

            x = donnees[variable1]
            y = donnees[variable2]

            if methode == "Pearson":

                coefficient, p_value = pearsonr(
                    x,
                    y
                )

            else:

                coefficient, p_value = spearmanr(
                    x,
                    y
                )

            resultat_corr = pd.DataFrame({

                "Indicateur": [
                    "Coefficient",
                    "p-value",
                    "Effectif"
                ],

                "Valeur": [
                    round(
                        coefficient,
                        4
                    ),

                    round(
                        p_value,
                        4
                    ),

                    len(donnees)
                ]
            })

            st.dataframe(
                resultat_corr,
                use_container_width=True,
                hide_index=True
            )

            st.scatter_chart(
                donnees,
                x=variable1,
                y=variable2
            )

            if p_value < 0.05:

                st.success(
                    f"La corrélation de {methode} est "
                    "statistiquement significative au seuil de 5 %."
                )

            else:

                st.info(
                    f"La corrélation de {methode} n'est pas "
                    "statistiquement significative au seuil de 5 %."
                )

    # ========================================================
    # QUALITATIVE × QUANTITATIVE
    # ========================================================

    elif (
        type1 in types_qualitatifs
        and
        type2 == "Quantitative"
    ):

        variable_qualitative = variable1
        variable_quantitative = variable2

        st.markdown(
            "### 9.3 Qualitative × Quantitative"
        )

        donnees = df_nettoye[
            [
                variable_qualitative,
                variable_quantitative
            ]
        ].copy()

        donnees[
            variable_quantitative
        ] = pd.to_numeric(
            donnees[
                variable_quantitative
            ],
            errors="coerce"
        )

        donnees = donnees.dropna()

        groupes = [
            groupe[
                variable_quantitative
            ].values
            for _, groupe
            in donnees.groupby(
                variable_qualitative
            )
        ]

        if len(groupes) < 2:

            st.warning(
                "Il faut au moins deux groupes "
                "pour réaliser une comparaison."
            )

        else:

            statistiques_groupes = (
                donnees
                .groupby(
                    variable_qualitative
                )[
                    variable_quantitative
                ]
                .agg(
                    Effectif="count",
                    Moyenne="mean",
                    Médiane="median",
                    Écart_type="std",
                    Minimum="min",
                    Maximum="max"
                )
                .reset_index()
            )

            colonnes_arrondir = [
                "Moyenne",
                "Médiane",
                "Écart_type",
                "Minimum",
                "Maximum"
            ]

            statistiques_groupes[
                colonnes_arrondir
            ] = statistiques_groupes[
                colonnes_arrondir
            ].round(2)

            st.markdown(
                "#### Statistiques par groupe"
            )

            st.dataframe(
                statistiques_groupes,
                use_container_width=True,
                hide_index=True
            )

            if len(groupes) == 2:

                methode = st.selectbox(
                    "Test de comparaison",
                    [
                        "t-test de Welch",
                        "Mann-Whitney"
                    ],
                    key="test_deux_groupes"
                )

                if methode == "t-test de Welch":

                    statistique, p_value = ttest_ind(
                        groupes[0],
                        groupes[1],
                        equal_var=False
                    )

                else:

                    statistique, p_value = mannwhitneyu(
                        groupes[0],
                        groupes[1],
                        alternative="two-sided"
                    )

                resultat = pd.DataFrame({

                    "Indicateur": [
                        "Test",
                        "Statistique",
                        "p-value",
                        "Effectif total"
                    ],

                    "Valeur": [
                        methode,
                        round(
                            statistique,
                            4
                        ),
                        round(
                            p_value,
                            4
                        ),
                        len(donnees)
                    ]
                })

                st.dataframe(
                    resultat,
                    use_container_width=True,
                    hide_index=True
                )

                if p_value < 0.05:

                    st.success(
                        "La différence entre les deux groupes "
                        "est statistiquement significative "
                        "au seuil de 5 %."
                    )

                else:

                    st.info(
                        "La différence entre les deux groupes "
                        "n'est pas statistiquement significative "
                        "au seuil de 5 %."
                    )

            else:

                methode = st.selectbox(
                    "Test de comparaison",
                    [
                        "ANOVA à un facteur",
                        "Kruskal-Wallis"
                    ],
                    key="test_plusieurs_groupes"
                )

                if methode == "ANOVA à un facteur":

                    statistique, p_value = f_oneway(
                        *groupes
                    )

                else:

                    statistique, p_value = kruskal(
                        *groupes
                    )

                resultat = pd.DataFrame({

                    "Indicateur": [
                        "Test",
                        "Statistique",
                        "p-value",
                        "Nombre de groupes",
                        "Effectif total"
                    ],

                    "Valeur": [
                        methode,
                        round(
                            statistique,
                            4
                        ),
                        round(
                            p_value,
                            4
                        ),
                        len(groupes),
                        len(donnees)
                    ]
                })

                st.dataframe(
                    resultat,
                    use_container_width=True,
                    hide_index=True
                )

                if p_value < 0.05:

                    st.success(
                        "Le test détecte une différence "
                        "statistiquement significative entre "
                        "au moins deux groupes au seuil de 5 %."
                    )

                    st.caption(
                        "Ce résultat ne précise pas à lui seul "
                        "quels groupes diffèrent entre eux. "
                        "Des comparaisons post-hoc seraient nécessaires."
                    )

                else:

                    st.info(
                        "Le test ne détecte pas de différence "
                        "statistiquement significative entre les groupes."
                    )

            st.bar_chart(
                statistiques_groupes.set_index(
                    variable_qualitative
                )["Moyenne"]
            )

    # ========================================================
    # QUANTITATIVE × QUALITATIVE
    # ========================================================

    elif (
        type1 == "Quantitative"
        and
        type2 in types_qualitatifs
    ):

        variable_quantitative = variable1
        variable_qualitative = variable2

        st.markdown(
            "### 9.4 Quantitative × Qualitative"
        )

        donnees = df_nettoye[
            [
                variable_quantitative,
                variable_qualitative
            ]
        ].copy()

        donnees[
            variable_quantitative
        ] = pd.to_numeric(
            donnees[
                variable_quantitative
            ],
            errors="coerce"
        )

        donnees = donnees.dropna()

        groupes = [
            groupe[
                variable_quantitative
            ].values
            for _, groupe
            in donnees.groupby(
                variable_qualitative
            )
        ]

        if len(groupes) < 2:

            st.warning(
                "Il faut au moins deux groupes "
                "pour réaliser une comparaison."
            )

        else:

            statistiques_groupes = (
                donnees
                .groupby(
                    variable_qualitative
                )[
                    variable_quantitative
                ]
                .agg(
                    Effectif="count",
                    Moyenne="mean",
                    Médiane="median",
                    Écart_type="std",
                    Minimum="min",
                    Maximum="max"
                )
                .reset_index()
            )

            colonnes_arrondir = [
                "Moyenne",
                "Médiane",
                "Écart_type",
                "Minimum",
                "Maximum"
            ]

            statistiques_groupes[
                colonnes_arrondir
            ] = statistiques_groupes[
                colonnes_arrondir
            ].round(2)

            st.dataframe(
                statistiques_groupes,
                use_container_width=True,
                hide_index=True
            )

            if len(groupes) == 2:

                methode = st.selectbox(
                    "Test de comparaison",
                    [
                        "t-test de Welch",
                        "Mann-Whitney"
                    ],
                    key="test_quant_qual_2"
                )

                if methode == "t-test de Welch":

                    statistique, p_value = ttest_ind(
                        groupes[0],
                        groupes[1],
                        equal_var=False
                    )

                else:

                    statistique, p_value = mannwhitneyu(
                        groupes[0],
                        groupes[1],
                        alternative="two-sided"
                    )

                resultat = pd.DataFrame({

                    "Indicateur": [
                        "Test",
                        "Statistique",
                        "p-value"
                    ],

                    "Valeur": [
                        methode,
                        round(
                            statistique,
                            4
                        ),
                        round(
                            p_value,
                            4
                        )
                    ]
                })

                st.dataframe(
                    resultat,
                    use_container_width=True,
                    hide_index=True
                )

                if p_value < 0.05:

                    st.success(
                        "La différence entre les deux groupes "
                        "est statistiquement significative "
                        "au seuil de 5 %."
                    )

                else:

                    st.info(
                        "La différence entre les deux groupes "
                        "n'est pas statistiquement significative "
                        "au seuil de 5 %."
                    )

            else:

                methode = st.selectbox(
                    "Test de comparaison",
                    [
                        "ANOVA à un facteur",
                        "Kruskal-Wallis"
                    ],
                    key="test_quant_qual_multi"
                )

                if methode == "ANOVA à un facteur":

                    statistique, p_value = f_oneway(
                        *groupes
                    )

                else:

                    statistique, p_value = kruskal(
                        *groupes
                    )

                resultat = pd.DataFrame({

                    "Indicateur": [
                        "Test",
                        "Statistique",
                        "p-value",
                        "Nombre de groupes"
                    ],

                    "Valeur": [
                        methode,
                        round(
                            statistique,
                            4
                        ),
                        round(
                            p_value,
                            4
                        ),
                        len(groupes)
                    ]
                })

                st.dataframe(
                    resultat,
                    use_container_width=True,
                    hide_index=True
                )

                if p_value < 0.05:

                    st.success(
                        "Le test détecte une différence "
                        "statistiquement significative entre "
                        "au moins deux groupes."
                    )

                    st.caption(
                        "Des analyses post-hoc seraient nécessaires "
                        "pour identifier précisément les groupes concernés."
                    )

                else:

                    st.info(
                        "Le test ne détecte pas de différence "
                        "statistiquement significative entre les groupes."
                    )

            st.bar_chart(
                statistiques_groupes.set_index(
                    variable_qualitative
                )["Moyenne"]
            )

    else:

        st.warning(
            "Cette combinaison de types de variables "
            "n'est pas encore prise en charge."
        )

# ============================================================
# FIN
# ============================================================

st.success(
    "Analyse descriptive, nettoyage et analyse bivariée disponibles."
)

st.info(
    "Une association ou une différence statistiquement "
    "significative ne constitue pas à elle seule une preuve "
    "de causalité."
)
