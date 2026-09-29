import streamlit as st
import pandas as pd
import numpy as np
import re

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
    "dictionnaire, fréquences et statistiques descriptives."
)

# ============================================================
# IMPORTATION
# ============================================================

fichier = st.file_uploader(
    "Choisissez votre fichier de données",
    type=["xlsx", "csv"]
)

if fichier is None:
    st.info("Veuillez importer un fichier Excel ou CSV.")
    st.stop()

# ============================================================
# LECTURE
# ============================================================

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
# INFORMATIONS GÉNÉRALES
# ============================================================

st.subheader("1. Informations générales")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("Lignes", df.shape[0])

with col2:
    st.metric("Variables", df.shape[1])

with col3:
    st.metric("Doublons", df.duplicated().sum())

with col4:
    st.metric(
        "Cellules manquantes",
        int(df.isna().sum().sum())
    )

# ============================================================
# APERÇU
# ============================================================

st.subheader("2. Aperçu des données")

st.dataframe(
    df.head(20),
    use_container_width=True,
    hide_index=True
)

# ============================================================
# DIAGNOSTIC
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
        round(df[col].isna().mean() * 100, 2)
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

# ============================================================
# DOUBLONS
# ============================================================

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


# ============================================================
# TYPES
# ============================================================

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
# DICTIONNAIRE
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
        "Date"
    ]
    else "Question fermée"
    for col in df.columns
]

# ============================================================
# INITIALISATION SESSION
# ============================================================

colonnes_requises = [
    "Variable",
    "Type Python",
    "Nombre de modalités",
    "Valeurs manquantes",
    "Type d'analyse",
    "Type de question"
]

if (
    "dictionnaire_modifie" not in st.session_state
    or
    not all(
        col in st.session_state[
            "dictionnaire_modifie"
        ].columns
        for col in colonnes_requises
    )
    or
    len(
        st.session_state[
            "dictionnaire_modifie"
        ]
    ) != len(df.columns)
):

    st.session_state[
        "dictionnaire_modifie"
    ] = dictionnaire.copy()

# ============================================================
# ÉDITION DU DICTIONNAIRE
# ============================================================

st.subheader("4. Dictionnaire des variables")

dictionnaire_modifie = st.data_editor(
    st.session_state[
        "dictionnaire_modifie"
    ],

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
# ANALYSE QUALITATIVE
# ============================================================

st.subheader("5. Analyse des variables qualitatives")

types_qualitatifs = [
    "Qualitative",
    "Qualitative codée"
]

for _, ligne in dictionnaire_modifie.iterrows():

    variable = ligne["Variable"]
    type_analyse = ligne["Type d'analyse"]
    type_question = ligne["Type de question"]

    # ========================================================
    # QUESTIONS FERMÉES
    # ========================================================

    if (
        type_analyse in types_qualitatifs
        and type_question == "Question fermée"
    ):

        st.markdown(f"### {variable}")

        serie = df[variable]

        valide = serie.dropna()

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
            resultat.set_index("Modalité")["Effectif"]
        )

    # ========================================================
    # QUESTIONS OUVERTES
    # ========================================================

    elif (
        type_analyse in types_qualitatifs
        and type_question == "Question ouverte"
    ):

        st.markdown(f"### {variable}")

        serie = df[variable].dropna().astype(str)

        st.write(
            f"**Nombre de réponses : {len(serie)}**"
        )

        # Affichage des réponses
        apercu = pd.DataFrame({
            "Réponses": serie.head(20).values
        })

        st.dataframe(
            apercu,
            use_container_width=True,
            hide_index=True
        )

        st.info(
            "Cette variable est ouverte. "
            "Elle sera soumise au module de codification."
        )

    # ========================================================
    # RÉPONSES MULTIPLES
    # ========================================================

    elif (
        type_analyse in types_qualitatifs
        and type_question == "Réponses multiples"
    ):

        st.markdown(f"### {variable}")

        serie = df[variable].dropna().astype(str)

        # ----------------------------------------------------
        # Séparation des réponses
        # ----------------------------------------------------

        reponses = []

        for valeur in serie:

            morceaux = re.split(
                r"[,;|]",
                valeur
            )

            for morceau in morceaux:

                morceau = morceau.strip()

                if morceau:
                    reponses.append(morceau)

        if reponses:

            freq = pd.Series(
                reponses
            ).value_counts()

            pourcentage = (
                freq / len(serie) * 100
            )

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
                "Les pourcentages peuvent dépasser 100 % au total "
                "car un répondant peut avoir plusieurs réponses."
            )

        else:

            st.warning(
                "Aucune réponse multiple exploitable détectée."
            )

# ============================================================
# ANALYSE QUANTITATIVE
# ============================================================

st.subheader("6. Analyse des variables quantitatives")

for _, ligne in dictionnaire_modifie.iterrows():

    variable = ligne["Variable"]

    type_analyse = ligne["Type d'analyse"]

    if type_analyse == "Quantitative":

        st.markdown(f"### {variable}")

        serie = pd.to_numeric(
            df[variable],
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

                df[variable].isna().sum(),

                round(serie.mean(), 2),

                round(serie.median(), 2),

                round(serie.std(), 2),

                round(serie.min(), 2),

                round(serie.quantile(0.25), 2),

                round(serie.quantile(0.75), 2),

                round(serie.max(), 2)
            ]
        })

        st.dataframe(
            statistiques,
            use_container_width=True,
            hide_index=True
        )

# ============================================================
# RÉSUMÉ FINAL DU DIAGNOSTIC
# ============================================================

st.subheader("7. Résumé du diagnostic")

nb_qualitatives = len(
    dictionnaire_modifie[
        dictionnaire_modifie["Type d'analyse"].isin(
            types_qualitatifs
        )
    ]
)

nb_quantitatives = len(
    dictionnaire_modifie[
        dictionnaire_modifie["Type d'analyse"]
        == "Quantitative"
    ]
)

nb_dates = len(
    dictionnaire_modifie[
        dictionnaire_modifie["Type d'analyse"]
        == "Date"
    ]
)

nb_identifiants = len(
    dictionnaire_modifie[
        dictionnaire_modifie["Type d'analyse"]
        == "Identifiant"
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

st.success(
    "Analyse descriptive terminée."
)
