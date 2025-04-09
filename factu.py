import streamlit as st
import pandas as pd
from io import BytesIO
from PIL import Image
import numpy as np

st.set_page_config(page_title="Excel Transformer", page_icon=":bar_chart:", layout="wide")

# image = Image.open("SRC/logo.png")

# Titre
st.title("📊 Traitement de fichier Excel")

# Chargement du fichier Excel
st.header("Charger un fichier Excel")
uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=["xlsx", "xls", "csv"])

if uploaded_file:

    ligne_noms_colonnes = st.number_input(
    "Sur quelle ligne du fichier se trouvent les noms de colonnes ?",
    min_value=1, value=1, step=1
)
    # Lecture du fichier Excel
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl', header=ligne_noms_colonnes - 1) # Header= numéro de la ligne sur laquelle se trouvent les noms de colonne
        st.success("Fichier chargé avec succès !")
        st.subheader("Aperçu des données avant Transformation")
        st.dataframe(df)

    except Exception as e:
        st.error(f"Erreur lors du traitement du fichier : {e}")

    st.markdown("***")
    st.markdown("***")

    ### TRAITEMENT ###
    # cleaned_df = df.dropna() # Suppresion des doublons
    cleaned_df = df.iloc[:, 1:] # Suppression de la 1ere colonne (si vide)
    # cleaned_df = df.drop(columns=["Unnamed: 0"]) # suppression de la colonne nommée "Colonne 0"
    
    # Suppression de colonnes
    colonnes_a_supprimer = st.multiselect(
        "🗑️ Sélectionnez les colonnes à supprimer",
        options=df.columns.tolist()
    )
    # Supprimer les colonnes sélectionnées
    df_affiche = df.drop(columns=colonnes_a_supprimer) if colonnes_a_supprimer else df
    st.info(f"Colonnes supprimées : {', '.join(colonnes_a_supprimer)}")

    st.subheader("Aperçu après Transformation :")
    st.dataframe(df_affiche)

    st.markdown("***")
    st.markdown("***")

    # Export du fichier modifié
    st.header("Exporter le fichier modifié au format Excel")

    # # Fonction transformation en excel
    # def to_excel(df):
    #     output = BytesIO()
    #     with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    #         df.to_excel(writer, index=False, sheet_name='Feuil1')
    #     processed_data = output.getvalue()
    #     return processed_data

    # Fonction de mise en forme et transformation en excel
    def to_excel_with_format(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Feuil1')

            # Accéder à l'objet worksheet et workbook
            workbook = writer.book
            worksheet = writer.sheets['Feuil1']

            # Définir le format pour les cellules
            header_format = workbook.add_format({
                'bold': True,
                'font_color': 'white',
                'bg_color': '#4F81BD',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            cell_format = workbook.add_format({
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            # Nettoyer les NaN/inf
            df = df.replace([np.nan, np.inf, -np.inf], None)

            # Appliquer le format aux en-têtes
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # Appliquer le format aux autres cellules
            for row_num, row in enumerate(df.values, start=1):
                for col_num, value in enumerate(row):
                    worksheet.write(row_num, col_num, value, cell_format)

            # Ajuster la largeur des colonnes
            for col_num, value in enumerate(df.columns.values):
                max_length = max(df[value].astype(str).map(len).max(), len(value))
                worksheet.set_column(col_num, col_num, max_length + 2)

            # ✅ Ajout du tableau structuré Excel
            worksheet.add_table(0, 0, df.shape[0], df.shape[1] - 1, {
                'columns': [{'header': col} for col in df.columns],
                'name': 'MonTableau',
                'style': 'Table Style Medium 9'
            })

        processed_data = output.getvalue()
        return processed_data

    # Téléchargement
    st.subheader("📥 Télécharger le fichier transformé")
    excel_data = to_excel_with_format(cleaned_df)

    # Nom de fichier de sortie (sans extension)
    file_name = st.text_input('📝 Choisir le Nom du fichier de sortie (sans extension) et appuyer sur "Entrée"', value="fichier_modifié")
    file_name=f"{file_name}.xlsx"

    st.download_button(
        label = "📁 Télécharger le fichier",
        data = excel_data,
        file_name = file_name,
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
