import streamlit as st
import pandas as pd

# Configuration de la page
st.set_page_config(
    page_title="Assistant statistique",
    page_icon="📊",
    layout="wide"
)

# Titre principal
st.title("Assistant statistique")

st.write(
    "Importez votre fichier de données pour commencer l'analyse."
)

# Importation du fichier
fichier = st.file_uploader(
    "Choisissez votre fichier de données",
    type=["xlsx", "csv"]
)

# Si un fichier est importé
if fichier is not None:

    # Lecture du fichier Excel
    if fichier.name.endswith(".xlsx"):
        df = pd.read_excel(fichier)

    # Lecture du fichier CSV
    else:
        df = pd.read_csv(fichier)

    st.success(
        f"Fichier chargé avec succès : {fichier.name}"
    )

    # Informations générales
    st.subheader("Informations sur les données")

    col1, col2 = st.columns(2)

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

    # Aperçu
    st.subheader("Aperçu des données")

    st.dataframe(
        df,
        use_container_width=True
    )
