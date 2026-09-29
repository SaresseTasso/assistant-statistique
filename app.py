import streamlit as st
import pandas as pd

# ============================================================
# CONFIGURATION DE LA PAGE
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

# ============================================================
# ANALYSE DU FICHIER
# ============================================================

if fichier is not None:

    # Lecture Excel
    if fichier.name.endswith(".xlsx"):
        df = pd.read_excel(fichier)

    # Lecture CSV
    else:
        df = pd.read_csv(fichier)

    st.success(
        f"Fichier chargé avec succès : {fichier.name}"
    )

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
    # APERÇU DES DONNÉES
    # ========================================================

    st.subheader("Aperçu des données")

    st.dataframe(
        df.head(20),
        use_container_width=True
    )

    # ========================================================
    # DIAGNOSTIC DES VARIABLES
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
        use_container_width=True
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
        use_container_width=True
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
    # ============================================================
# DICTIONNAIRE DES VARIABLES
# ============================================================

st.subheader("Dictionnaire des variables")

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

# Proposition automatique du type d'analyse
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


# ============================================================
# VALIDATION DU TYPE DES VARIABLES
# ============================================================

types_possibles = [
    "Qualitative",
    "Qualitative codée",
    "Quantitative",
    "Date",
    "Identifiant",
    "À vérifier"
]

dictionnaire["Type d'analyse"] = [
    proposer_type(col)
    for col in df.columns
]

if "dictionnaire_modifie" not in st.session_state:
    st.session_state["dictionnaire_modifie"] = dictionnaire.copy()

dictionnaire_modifie = st.data_editor(
    st.session_state["dictionnaire_modifie"],
    column_config={
        "Type d'analyse": st.column_config.SelectboxColumn(
            "Type d'analyse",
            options=types_possibles
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

st.info(
    "Vérifiez le type d'analyse de chaque variable avant "
    "de poursuivre l'analyse."
)
# ============================================================
# ANALYSE DES VARIABLES QUALITATIVES
# ============================================================

st.subheader("Analyse des variables qualitatives")

types_qualitatifs = [
    "Qualitative",
    "Qualitative codée"
]

for _, ligne in dictionnaire_modifie.iterrows():

    variable = ligne["Variable"]
    type_analyse = ligne["Type d'analyse"]

    if type_analyse in types_qualitatifs:

        st.markdown(f"### {variable}")

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

        # Tableau de résultats
        resultat = pd.DataFrame({
            "Modalité": effectifs.index.astype(str),
            "Effectif": effectifs.values,
            "Pourcentage": pourcentages.values.round(2)
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

        st.bar_chart(graphique)
