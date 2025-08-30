import streamlit as st
import pandas as pd
import tempfile
import os
from CuadroFacturacionGenerator import CuadroFacturacionGenerator

st.set_page_config(page_title="Generador de Cuadro de Facturación", layout="centered")

st.title("🧾 Generador de Cuadro de Facturación")
st.markdown("Sube el archivo de Excel con el conglomerado, selecciona un profesional o descarga el consolidado.")

uploaded_file = st.file_uploader("📤 Cargar archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_input:
        temp_input.write(uploaded_file.read())
        temp_input_path = temp_input.name

    try:
        df_preview = pd.read_excel(temp_input_path, sheet_name="CONGLOMERADO", engine="openpyxl")
        nombres_profesionales = sorted(df_preview["NOMBRE DEL PROFESIONAL"].dropna().unique())

        nombre_seleccionado = st.selectbox("👤 Selecciona el profesional:", nombres_profesionales)

        col1, col2 = st.columns(2)

        with col1:
            if nombre_seleccionado and st.button("🚀 Generar archivo por profesional"):
                generador = CuadroFacturacionGenerator()

                with st.spinner("⏳ Generando archivo del profesional..."):
                    temp_output = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{nombre_seleccionado.replace(' ', '_')}.xlsx")
                    temp_output_path = temp_output.name
                    temp_output.close()

                    generador.generar_filtrado_por_profesional(temp_input_path, temp_output_path, nombre_seleccionado)

                st.success("✅ Archivo generado. Descárgalo a continuación:")

                with open(temp_output_path, "rb") as f:
                    st.download_button(
                        label=f"📥 Descargar {nombre_seleccionado}",
                        data=f,
                        file_name=f"CUADRO_{nombre_seleccionado.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{nombre_seleccionado}"
                    )

        with col2:
            if st.button("📊 Generar consolidado de todos"):
                generador = CuadroFacturacionGenerator()

                with st.spinner("⏳ Generando consolidado..."):
                    temp_output = tempfile.NamedTemporaryFile(delete=False, suffix="_Consolidado.xlsx")
                    temp_output_path = temp_output.name
                    temp_output.close()

                    generador.generar_todos(temp_input_path, temp_output_path)

                st.success("✅ Consolidado generado.")

                with open(temp_output_path, "rb") as f:
                    st.download_button(
                        label="📥 Descargar consolidado",
                        data=f,
                        file_name="CUADRO_CONSOLIDADO.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_consolidado"
                    )

    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {e}")
    finally:
        os.remove(temp_input_path)
