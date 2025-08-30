import streamlit as st
from CuadroFacturacionGenerator import CuadroFacturacionGenerator
import os
import zipfile

st.set_page_config(page_title="Cuadro de Facturación", layout="wide")

st.title("📊 Generador de Cuadro de Facturación")

uploaded_file = st.file_uploader("Sube el archivo del conglomerado (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        generator = CuadroFacturacionGenerator(uploaded_file)
        generator.load_data()

        st.success("✅ Archivo cargado correctamente")

        if st.button("Generar Cuadros de Facturación"):
            with st.spinner("Procesando archivos..."):
                archivos_individuales = generator.generar_por_profesional()
                archivo_consolidado = generator.generar_consolidado()

            st.success("✅ Archivos generados")

            # ZIP con todos los archivos individuales
            zip_path = "output/Facturacion_Individual.zip"
            with zipfile.ZipFile(zip_path, "w") as zf:
                for archivo in archivos_individuales:
                    zf.write(archivo, os.path.basename(archivo))

            # Descarga de ZIP
            with open(zip_path, "rb") as f:
                st.download_button(
                    "⬇️ Descargar archivos individuales (ZIP)",
                    f,
                    file_name="Facturacion_Individual.zip"
                )

            # Descarga de consolidado
            with open(archivo_consolidado, "rb") as f:
                st.download_button(
                    "⬇️ Descargar consolidado",
                    f,
                    file_name="Consolidado_Facturacion.xlsx"
                )

    except Exception as e:
        st.error(f"❌ Error: {e}")
