import streamlit as st
import pandas as pd

# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Assistant statistique",
    page_icon="📊",
    layout="wide"
)

# ============================================================
# TITRE
# ============================================================

st.title("Assistant statistique")

st.write(
    "Importez votre fichier de données pour commencer l'analyse."
)

# ============================================================
# IMPORTATION DU FICHIER
# ============================================================

fichier = st.file_uploader(
    "Choisissez votre fichier de données",
    type=["xlsx", "csv"]
)

if fichier is not None:

    # ========================================================
    # LECTURE DU FICHIER
    # ========================================================

    try:

        if fichier.name.endswith(".xlsx"):
            df = pd.read_excel(fichier)

        else:
            df = pd.read_csv(fichier)

        st.success(
            f"Fichier chargé avec succès : {fichier.name}"
        )

    except Exception as e:

        st.error(
            f"Une erreur est survenue lors de la lecture du fichier : {e}"
        )

        st.stop()

    # ========================================================
    # INFORMATIONS GÉNÉRALES
    # ========================================================

    st.subheader("Informations générales")

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Nombre de lignes",
            df.shape[0]
        )

    with col2:

        st.metric(
            "Nombre de variables",
            df.shape[1]
        )

    with col3:

        st.metric(
            "Nombre de doublons",
            df.duplicated().sum()
        )

    # ========================================================
    # APERÇU
    # ========================================================

    st.subheader("Aperçu des données")

    st.dataframe(
        df.head(20),
        use_container_width=True
    )

    # ========================================================
    # DIAGNOSTIC
    # ========================================================

    st.subheader("Diagnostic des variables")

    diagnostic = pd.DataFrame({

        "Variable": df.columns,

        "Type Python": [
            str(df[col].dtype)
            for col in df.columns
        ],

        "Valeurs manquantes": [
            df[col].isna().sum()
            for col in df.columns
        ],

        "Valeurs uniques": [
            df[col].nunique(dropna=True)
            for col in df.columns
        ]
    })

    st.dataframe(
        diagnostic,
        use_container_width=True,
        hide_index=True
    )

    # ========================================================
    # VALEURS MANQUANTES
    # ========================================================

    st.subheader("Valeurs manquantes")

    manquants = pd.DataFrame({

        "Variable": df.columns,

        "Effectif manquant": [
            df[col].isna().sum()
            for col in df.columns
        ],

        "Pourcentage manquant": [
            round(
                df[col].isna().mean() * 100,
                2
            )
            for col in df.columns
        ]
    })

    st.dataframe(
        manquants,
        use_container_width=True,
        hide_index=True
    )

    # ========================================================
    # DOUBLONS
    # ========================================================

    st.subheader("Doublons")

    nombre_doublons = df.duplicated().sum()

    if nombre_doublons == 0:

        st.success(
            "Aucun doublon détecté."
        )

    else:

        st.warning(
            f"{nombre_doublons} doublon(s) détecté(s)."
        )

    # ========================================================
    # FONCTION DE PROPOSITION DU TYPE
    # ========================================================

    def proposer_type(colonne):

        serie = df[colonne]

        # Date
        if pd.api.types.is_datetime64_any_dtype(serie):

            return "Date"

        # Numérique
        if pd.api.types.is_numeric_dtype(serie):

            return "Quantitative"

        # Texte
        return "Qualitative"

    # ========================================================
    # TYPES DISPONIBLES
    # ========================================================

    types_possibles = [
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

    # ========================================================
    # CONSTRUCTION DU DICTIONNAIRE
    # ========================================================

    dictionnaire = pd.DataFrame({

        "Variable": df.columns,

        "Type Python": [
            str(df[col].dtype)
            for col in df.columns
        ],

        "Nombre de modalités": [
            df[col].nunique(dropna=True)
            for col in df.columns
        ],

        "Valeurs manquantes": [
            df[col].isna().sum()
            for col in df.columns
        ]
    })

    dictionnaire["Type d'analyse"] = [
        proposer_type(col)
        for col in df.columns
    ]

    # Par défaut :
    # - numérique/date = non applicable
    # - qualitative = à vérifier

    dictionnaire["Type de question"] = [

        "Non applicable"
        if proposer_type(col) in [
            "Quantitative",
            "Date"
        ]
        else "Question fermée"

        for col in df.columns
    ]

    # ========================================================
    # INITIALISATION / ACTUALISATION DU DICTIONNAIRE
    # ========================================================

    colonnes_requises = [
        "Variable",
        "Type Python",
        "Nombre de modalités",
        "Valeurs manquantes",
        "Type d'analyse",
        "Type de question"
    ]

    # Si le dictionnaire n'existe pas
    # OU si son ancienne structure est différente,
    # on le reconstruit.

    if (
        "dictionnaire_modifie" not in st.session_state
        or
        not all(
            colonne in st.session_state[
                "dictionnaire_modifie"
            ].columns
            for colonne in colonnes_requises
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

    # ========================================================
    # DICTIONNAIRE DES VARIABLES
    # ========================================================

    st.subheader(
        "Dictionnaire des variables"
    )

    dictionnaire_modifie = st.data_editor(

        st.session_state[
            "dictionnaire_modifie"
        ],

        column_config={

            "Type d'analyse":
                st.column_config.SelectboxColumn(
                    "Type d'analyse",
                    options=types_possibles
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

    # Sauvegarde des modifications
    st.session_state[
        "dictionnaire_modifie"
    ] = dictionnaire_modifie

    # ========================================================
    # INFORMATION
    # ========================================================

    st.info(
        "Vérifiez le type d'analyse et le type de question "
        "de chaque variable avant de poursuivre."
    )

    # ========================================================
    # ANALYSE DES VARIABLES QUALITATIVES
    # ========================================================

    st.subheader(
        "Analyse des variables qualitatives"
    )

    types_qualitatifs = [
        "Qualitative",
        "Qualitative codée"
    ]

    # --------------------------------------------------------
    # PARCOURS DES VARIABLES
    # --------------------------------------------------------

    for _, ligne in dictionnaire_modifie.iterrows():

        variable = ligne["Variable"]

        type_analyse = ligne["Type d'analyse"]

        type_question = ligne["Type de question"]

        # ====================================================
        # QUESTION FERMÉE
        # ====================================================

        if (
            type_analyse in types_qualitatifs
            and
            type_question == "Question fermée"
        ):

            st.markdown(
                f"### {variable}"
            )

            serie = df[variable]

            # Effectifs
            effectifs = serie.value_counts(
                dropna=False
            )

            # Pourcentages
            pourcentages = (
                serie.value_counts(
                    normalize=True,
                    dropna=False
                ) * 100
            )

            # Tableau
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

            # Graphique
            graphique = resultat.set_index(
                "Modalité"
            )["Effectif"]

            st.bar_chart(
                graphique
            )

        # ====================================================
        # QUESTION OUVERTE
        # ====================================================

        elif (
            type_analyse in types_qualitatifs
            and
            type_question == "Question ouverte"
        ):

            st.markdown(
                f"### {variable}"
            )

            st.info(
                "Cette variable est identifiée comme "
                "question ouverte. Elle sera traitée "
                "dans le module de codification des "
                "réponses ouvertes."
            )

        # ====================================================
        # RÉPONSES MULTIPLES
        # ====================================================

        elif (
            type_analyse in types_qualitatifs
            and
            type_question == "Réponses multiples"
        ):

            st.markdown(
                f"### {variable}"
            )

            st.info(
                "Cette variable est identifiée comme "
                "question à réponses multiples. "
                "Elle sera traitée dans le module "
                "d'analyse des réponses multiples."
            )
