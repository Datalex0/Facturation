import streamlit as st
import pandas as pd
from io import BytesIO
from PIL import Image

st.set_page_config(page_title="Excel Transformer", layout="centered")

# image = Image.open("SRC/logo.png")

# Titre
st.title("📊 Traitement de fichier Excel")

# Chargement du fichier Excel
st.header("Charger un fichier Excel")
uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=["xlsx", "xls", "csv"])

if uploaded_file:
    # Lecture du fichier Excel
    try:
        df = pd.read_excel(uploaded_file)
        st.success("Fichier chargé avec succès !")
        st.subheader("Aperçu des données")
        st.dataframe(df)

        cleaned_df = df.dropna()

        st.write("Aperçu après nettoyage :")
        st.dataframe(cleaned_df)


        # Export du fichier modifié
        st.header("Exporter le fichier modifié en Excel")

        # Fonction transformation en excel
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
            processed_data = output.getvalue()
            return processed_data

        # Téléchargement
        st.subheader("📥 Télécharger le fichier transformé")
        excel_data = to_excel(cleaned_df)
        st.download_button(
            label="📁 Télécharger Excel",
            data=excel_data,
            file_name="fichier_modifié.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier : {e}")



