import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Convertisseur XLS → CSV", page_icon="📄")
st.title("📥 Convertisseur .xls vers .csv (ANSI / ; / qte)")
st.write("Téléverse un fichier `.xls`, et télécharge un `.csv` propre.")

uploaded_file = st.file_uploader("Dépose ton fichier .xls ici", type=["xls"])

def clean_column_name(name):
    if str(name).strip().lower() == "quantité":
        return "qte"
    return name

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, dtype=object, engine="xlrd")
        df.columns = [clean_column_name(col) for col in df.columns]

        for col in df.columns:
            df[col] = df[col].apply(lambda x: int(x) if isinstance(x, float) and x.is_integer() else x)

        output = io.StringIO()
        df.to_csv(output, index=False, sep=';', encoding='cp1252')
        output.seek(0)

        st.success("✅ Fichier prêt à être téléchargé.")
        st.download_button(
            label="📤 Télécharger le CSV",
            data=output.getvalue(),
            file_name="fichier_converti.csv",
            mime="text/csv",
        )

    except Exception as e:
        st.error(f"❌ Erreur : {e}")
