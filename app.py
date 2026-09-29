
import streamlit as st
import pandas as pd
import numpy as np
import re
import io

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

nb_lignes = len(df_nettoye)
nb_variables = len(df_nettoye.columns)
nb_manquants = int(df_nettoye.isna().sum().sum())
nb_doublons = int(df_nettoye.duplicated().sum())

q1, q2, q3, q4 = st.columns(4)

with q1:
    st.metric("Lignes", nb_lignes)

with q2:
    st.metric("Variables", nb_variables)

with q3:
    st.metric("Cellules manquantes", nb_manquants)

with q4:
    st.metric("Doublons", nb_doublons)


# ------------------------------------------------------------
# 10.2 Réponses à vérifier
# ------------------------------------------------------------

st.markdown("### 10.2 Réponses textuelles à vérifier")

valeurs_a_verifier = [
    "aucun",
    "aucune",
    "néant",
    "neant",
    "ras",
    "r.a.s.",
    "rien",
    "n'importe quel",
    "n’importe quel",
    "ouvert à tous niveaux",
    "ouvert a tous niveaux"
]

colonnes_textuelles_controle = (
    df_nettoye
    .select_dtypes(include=["object", "string"])
    .columns
)

reponses_suspectes = []

for col in colonnes_textuelles_controle:

    serie = (
        df_nettoye[col]
        .astype("string")
        .str.strip()
        .str.lower()
    )

    masque = serie.isin(valeurs_a_verifier)

    if masque.any():

        valeurs = (
            serie[masque]
            .value_counts()
            .reset_index()
        )

        valeurs.columns = [
            "Réponse",
            "Effectif"
        ]

        for _, ligne in valeurs.iterrows():

            reponses_suspectes.append({
                "Variable": col,
                "Réponse": ligne["Réponse"],
                "Effectif": int(ligne["Effectif"])
            })

if reponses_suspectes:

    tableau_suspect = pd.DataFrame(
        reponses_suspectes
    )

    st.warning(
        "Certaines réponses nécessitent une vérification "
        "humaine. Elles n'ont pas été supprimées."
    )

    st.dataframe(
        tableau_suspect,
        use_container_width=True,
        hide_index=True
    )

else:

    st.success(
        "Aucune réponse textuelle suspecte détectée."
    )


# ------------------------------------------------------------
# 10.3 Modalités très proches
# ------------------------------------------------------------

st.markdown("### 10.3 Recherche de modalités potentiellement différentes")

st.write(
    "Cette vérification recherche notamment des différences "
    "de casse ou d'espaces pouvant créer artificiellement "
    "plusieurs modalités."
)

modalites_proches = []

for col in colonnes_textuelles_controle:

    serie = (
        df_nettoye[col]
        .dropna()
        .astype(str)
    )

    valeurs_originales = serie.unique()

    groupes = {}

    for valeur in valeurs_originales:

        valeur_normalisee = (
            valeur
            .strip()
            .lower()
        )

        groupes.setdefault(
            valeur_normalisee,
            []
        ).append(valeur)

    for normalisee, valeurs in groupes.items():

        valeurs_uniques = list(
            dict.fromkeys(valeurs)
        )

        if len(valeurs_uniques) > 1:

            modalites_proches.append({
                "Variable": col,
                "Formes détectées": " | ".join(
                    valeurs_uniques
                )
            })

if modalites_proches:

    st.warning(
        "Des modalités semblent différentes uniquement "
        "à cause de la casse ou des espaces."
    )

    st.dataframe(
        pd.DataFrame(modalites_proches),
        use_container_width=True,
        hide_index=True
    )

else:

    st.success(
        "Aucune modalité manifestement similaire détectée."
    )


# ------------------------------------------------------------
# 10.4 Valeurs quantitatives extrêmes
# ------------------------------------------------------------

st.markdown("### 10.4 Valeurs quantitatives extrêmes")

st.write(
    "Une valeur extrême n'est pas automatiquement une erreur. "
    "Elle est simplement signalée pour vérification."
)

variables_quant_controle = (
    dictionnaire_modifie[
        dictionnaire_modifie[
            "Type d'analyse"
        ] == "Quantitative"
    ]["Variable"].tolist()
)

valeurs_extremes = []

for col in variables_quant_controle:

    serie = pd.to_numeric(
        df_nettoye[col],
        errors="coerce"
    ).dropna()

    if len(serie) < 4:
        continue

    q1 = serie.quantile(0.25)
    q3 = serie.quantile(0.75)

    iqr = q3 - q1

    borne_inf = q1 - 1.5 * iqr
    borne_sup = q3 + 1.5 * iqr

    masque = (
        (serie < borne_inf)
        |
        (serie > borne_sup)
    )

    nombre_extremes = int(masque.sum())

    if nombre_extremes > 0:

        valeurs_extremes.append({
            "Variable": col,
            "Valeurs extrêmes": nombre_extremes,
            "Borne inférieure": round(
                borne_inf,
                2
            ),
            "Borne supérieure": round(
                borne_sup,
                2
            )
        })

if valeurs_extremes:

    st.warning(
        "Certaines valeurs sont statistiquement extrêmes "
        "selon la règle de l'IQR."
    )

    st.dataframe(
        pd.DataFrame(valeurs_extremes),
        use_container_width=True,
        hide_index=True
    )

    st.caption(
        "Attention : une valeur extrême n'est pas nécessairement "
        "une erreur de saisie."
    )

else:

    st.success(
        "Aucune valeur extrême détectée selon la règle de l'IQR."
    )


# ------------------------------------------------------------
# 10.5 Contrôle des variables quantitatives
# ---------------------


# ============================================================
# 11. CODIFICATION AUTOMATIQUE DES QUESTIONS OUVERTES
# ============================================================

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans


st.subheader("11. Codification des questions ouvertes")

st.write(
    "L'application analyse les réponses ouvertes, recherche "
    "les réponses lexicalement similaires et propose des "
    "regroupements thématiques. Les propositions doivent "
    "être vérifiées et validées par l'utilisateur."
)


# ------------------------------------------------------------
# 11.1 Recherche des questions ouvertes
# ------------------------------------------------------------

variables_ouvertes = (
    dictionnaire_modifie[
        dictionnaire_modifie[
            "Type de question"
        ] == "Question ouverte"
    ]["Variable"].tolist()
)


if not variables_ouvertes:

    st.info(
        "Aucune question ouverte n'est actuellement identifiée "
        "dans le dictionnaire."
    )

else:

    variable_ouverte = st.selectbox(
        "Sélectionnez une question ouverte",
        variables_ouvertes,
        key="variable_question_ouverte_auto"
    )

    serie_ouverte = (
        df_nettoye[
            variable_ouverte
        ]
        .dropna()
        .astype(str)
        .str.strip()
    )

    serie_ouverte = serie_ouverte[
        serie_ouverte != ""
    ]

    st.write(
        f"**Nombre de réponses exploitables : "
        f"{len(serie_ouverte)}**"
    )


    # --------------------------------------------------------
    # 11.2 Affichage des réponses
    # --------------------------------------------------------

    st.markdown(
        "### 11.1 Réponses originales"
    )

    reponses_originales = pd.DataFrame({
        "Réponse originale":
            serie_ouverte.values
    })

    st.dataframe(
        reponses_originales,
        use_container_width=True,
        hide_index=True
    )


    # --------------------------------------------------------
    # 11.3 Normalisation
    # --------------------------------------------------------

    st.markdown(
        "### 11.2 Préparation automatique du texte"
    )

    def normaliser_texte_auto(texte):

        texte = str(texte)

        texte = (
            texte
            .strip()
            .lower()
        )

        texte = re.sub(
            r"\s+",
            " ",
            texte
        )

        return texte


    codification = pd.DataFrame({

        "Réponse originale":
            serie_ouverte.values

    })


    codification[
        "Réponse normalisée"
    ] = (
        codification[
            "Réponse originale"
        ]
        .apply(normaliser_texte_auto)
    )


    # --------------------------------------------------------
    # 11.4 Détection des réponses non informatives
    # --------------------------------------------------------

    st.markdown(
        "### 11.3 Réponses à vérifier"
    )

    reponses_non_informatives = [
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
        "ne sait pas"
    ]

    codification[
        "À vérifier"
    ] = (
        codification[
            "Réponse normalisée"
        ].isin(
            reponses_non_informatives
        )
    )

    nombre_non_informatives = int(
        codification[
            "À vérifier"
        ].sum()
    )

    if nombre_non_informatives > 0:

        st.warning(
            f"{nombre_non_informatives} réponse(s) "
            "ont été identifiées comme potentiellement "
            "non informatives."
        )

        st.dataframe(
            codification[
                codification["À vérifier"]
            ],
            use_container_width=True,
            hide_index=True
        )

    else:

        st.success(
            "Aucune réponse manifestement non informative "
            "n'a été détectée."
        )


    # --------------------------------------------------------
    # 11.5 Paramètres de l'analyse automatique
    # --------------------------------------------------------

    st.markdown(
        "### 11.4 Proposition automatique de thèmes"
    )

    nombre_reponses = len(codification)

    if nombre_reponses < 4:

        st.warning(
            "Il faut au moins 4 réponses pour proposer "
            "automatiquement des regroupements."
        )

    else:

        nombre_max_themes = min(
            8,
            nombre_reponses
        )

        nombre_themes = st.slider(
            "Nombre de thèmes à proposer",
            min_value=2,
            max_value=nombre_max_themes,
            value=min(
                4,
                nombre_max_themes
            ),
            key="nombre_themes_auto"
        )


        # ----------------------------------------------------
        # 11.6 Exclusion des réponses non informatives
        # ----------------------------------------------------

        donnees_clustering = codification[
            ~codification["À vérifier"]
        ].copy()

        if len(donnees_clustering) < 3:

            st.warning(
                "Il ne reste pas suffisamment de réponses "
                "informatives pour effectuer une proposition "
                "automatique."
            )

        else:

            textes = donnees_clustering[
                "Réponse normalisée"
            ].tolist()


            # ------------------------------------------------
            # TF-IDF
            # ------------------------------------------------

            try:

                vectoriseur = TfidfVectorizer(
                    lowercase=True,
                    strip_accents="unicode",
                    stop_words=None,
                    min_df=1,
                    max_df=0.95,
                    ngram_range=(1, 2)
                )

                matrice_tfidf = (
                    vectoriseur
                    .fit_transform(textes)
                )


                # --------------------------------------------
                # Vérification du nombre de caractéristiques
                # --------------------------------------------

                if matrice_tfidf.shape[1] < 2:

                    st.warning(
                        "Les réponses sont trop similaires "
                        "ou trop courtes pour effectuer "
                        "un regroupement automatique."
                    )

                else:

                    nombre_clusters = min(
                        nombre_themes,
                        len(textes)
                    )


                    # ----------------------------------------
                    # K-Means
                    # ----------------------------------------

                    modele = KMeans(
                        n_clusters=nombre_clusters,
                        random_state=42,
                        n_init=10
                    )

                    labels = modele.fit_predict(
                        matrice_tfidf
                    )


                    donnees_clustering[
                        "Groupe automatique"
                    ] = labels + 1


                    # ----------------------------------------
                    # Recherche des mots représentatifs
                    # ----------------------------------------

                    noms_themes = {}

                    termes = np.array(
                        vectoriseur
                        .get_feature_names_out()
                    )

                    centres = modele.cluster_centers_


                    for cluster_num in range(
                        nombre_clusters
                    ):

                        indices = (
                            centres[
                                cluster_num
                            ]
                            .argsort()[::-1]
                        )

                        mots = []

                        for indice in indices:

                            mot = termes[indice]

                            if mot not in mots:

                                mots.append(
                                    mot
                                )

                            if len(mots) >= 3:
                                break


                        nom_propose = (
                            " / ".join(mots)
                            if mots
                            else
                            f"Thème {cluster_num + 1}"
                        )

                        noms_themes[
                            cluster_num + 1
                        ] = nom_propose


                    # ----------------------------------------
                    # Attribution des noms
                    # ----------------------------------------

                    donnees_clustering[
                        "Thème proposé"
                    ] = (
                        donnees_clustering[
                            "Groupe automatique"
                        ]
                        .map(noms_themes)
                    )


                    # ----------------------------------------
                    # Résultats
                    # ----------------------------------------

                    st.markdown(
                        "### 11.5 Thèmes proposés"
                    )

                    resume_themes = (
                        donnees_clustering[
                            "Thème proposé"
                        ]
                        .value_counts()
                        .reset_index()
                    )

                    resume_themes.columns = [
                        "Thème proposé",
                        "Effectif"
                    ]

                    resume_themes[
                        "Pourcentage"
                    ] = (
                        resume_themes[
                            "Effectif"
                        ]
                        / len(codification)
                        * 100
                    ).round(2)


                    st.dataframe(
                        resume_themes,
                        use_container_width=True,
                        hide_index=True
                    )


                    st.info(
                        "Les thèmes proposés sont basés sur la "
                        "similarité lexicale des réponses. Ils ne "
                        "constituent pas une interprétation automatique "
                        "définitive du sens des réponses."
                    )


                    # ----------------------------------------
                    # Réponses par thème
                    # ----------------------------------------

                    st.markdown(
                        "### 11.6 Réponses regroupées"
                    )

                    for theme in noms_themes.values():

                        st.markdown(
                            f"#### {theme}"
                        )

                        reponses_theme = (
                            donnees_clustering[
                                donnees_clustering[
                                    "Thème proposé"
                                ] == theme
                            ][
                                [
                                    "Réponse originale"
                                ]
                            ]
                        )

                        st.dataframe(
                            reponses_theme,
                            use_container_width=True,
                            hide_index=True
                        )


                    # ----------------------------------------
                    # 11.7 Validation des thèmes
                    # ----------------------------------------

                    st.markdown(
                        "### 11.7 Validation des thèmes proposés"
                    )

                    st.write(
                        "Vous pouvez remplacer les noms proposés "
                        "par des intitulés plus pertinents pour "
                        "votre étude."
                    )

                    themes_valides = {}

                    for groupe, nom_propose in (
                        noms_themes.items()
                    ):

                        nom_valide = st.text_input(
                            f"Nom du thème {groupe}",
                            value=nom_propose,
                            key=(
                                f"nom_theme_valide_"
                                f"{variable_ouverte}_"
                                f"{groupe}"
                            )
                        )

                        themes_valides[
                            groupe
                        ] = nom_valide


                    # ----------------------------------------
                    # 11.8 Application des noms validés
                    # ----------------------------------------

                    donnees_clustering[
                        "Thème validé"
                    ] = (
                        donnees_clustering[
                            "Groupe automatique"
                        ]
                        .map(themes_valides)
                    )


                    # ----------------------------------------
                    # 11.9 Tableau final
                    # ----------------------------------------

                    st.markdown(
                        "### 11.8 Codification proposée"
                    )

                    resultat_final = (
                        donnees_clustering[
                            [
                                "Réponse originale",
                                "Réponse normalisée",
                                "Thème proposé",
                                "Thème validé"
                            ]
                        ]
                    )

                    st.dataframe(
                        resultat_final,
                        use_container_width=True,
                        hide_index=True
                    )


                    # ----------------------------------------
                    # 11.10 Résultats des thèmes validés
                    # ----------------------------------------

                    st.markdown(
                        "### 11.9 Résultats statistiques"
                    )

                    statistiques_themes = (
                        donnees_clustering[
                            "Thème validé"
                        ]
                        .value_counts()
                        .reset_index()
                    )

                    statistiques_themes.columns = [
                        "Thème",
                        "Effectif"
                    ]

                    statistiques_themes[
                        "Pourcentage"
                    ] = (
                        statistiques_themes[
                            "Effectif"
                        ]
                        / len(codification)
                        * 100
                    ).round(2)


                    st.dataframe(
                        statistiques_themes,
                        use_container_width=True,
                        hide_index=True
                    )


                    st.bar_chart(
                        statistiques_themes.set_index(
                            "Thème"
                        )["Effectif"]
                    )


                    # ----------------------------------------
                    # 11.11 Export Excel
                    # ----------------------------------------

                    st.markdown(
                        "### 11.10 Export de la codification"
                    )

                    try:

                        buffer_auto = io.BytesIO()

                        with pd.ExcelWriter(
                            buffer_auto,
                            engine="openpyxl"
                        ) as writer:

                            resultat_final.to_excel(
                                writer,
                                index=False,
                                sheet_name="Codification"
                            )

                            statistiques_themes.to_excel(
                                writer,
                                index=False,
                                sheet_name="Résultats"
                            )

                            resume_themes.to_excel(
                                writer,
                                index=False,
                                sheet_name="Propositions"
                            )

                        st.download_button(
                            label=(
                                "Télécharger la codification "
                                "automatique Excel"
                            ),
                            data=buffer_auto.getvalue(),
                            file_name=(
                                "codification_automatique.xlsx"
                            ),
                            mime=(
                                "application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet"
                            ),
                            key=(
                                "telecharger_codification_auto"
                            )
                        )

                    except Exception as e:

                        st.error(
                            "Erreur lors de la préparation "
                            f"du fichier : {e}"
                        )

            except Exception as e:

                st.error(
                    "La proposition automatique des thèmes "
                    f"a rencontré une erreur : {e}"
                )
