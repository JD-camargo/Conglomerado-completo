import streamlit as st
import pandas as pd
import tempfile
import os
from CuadroFacturacionGenerator import CuadroFacturacionGenerator

st.set_page_config(page_title="Generador de Cuadro de Facturación", layout="centered")

st.title("🧾 Generador de Cuadro de Facturación")
st.markdown("Sube el archivo de Excel con el conglomerado, selecciona un profesional o genera el cuadro de todos y descarga el archivo.")

uploaded_file = st.file_uploader("📤 Cargar archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_input:
        temp_input.write(uploaded_file.read())
        temp_input_path = temp_input.name

    try:
        # Leer hoja "CONGLOMERADO"
        df_preview = pd.read_excel(temp_input_path, sheet_name="CONGLOMERADO", engine="openpyxl")
        nombres_profesionales = sorted(df_preview["NOMBRE DEL PROFESIONAL"].dropna().unique())

        # Selección individual
        nombre_seleccionado = st.selectbox("👤 Selecciona el profesional:", nombres_profesionales)

        # Botón para generar archivo de un profesional
        if nombre_seleccionado and st.button("🚀 Generar archivo individual"):
            generador = CuadroFacturacionGenerator()

            with st.spinner("⏳ Generando archivo, por favor espera..."):
                temp_output = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{nombre_seleccionado.replace(' ', '_')}.xlsx")
                temp_output_path = temp_output.name
                temp_output.close()

                generador.generar_filtrado_por_profesional(temp_input_path, temp_output_path, [nombre_seleccionado])

            st.success(f"✅ Archivo generado para {nombre_seleccionado}. Descárgalo a continuación:")

            with open(temp_output_path, "rb") as f:
                st.download_button(
                    label=f"📥 Descargar {nombre_seleccionado}",
                    data=f,
                    file_name=f"CUADRO_{nombre_seleccionado.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{nombre_seleccionado}"
                )

        # Botón para generar archivo con TODOS
        if st.button("🚀 Generar archivo con TODOS"):
            generador = CuadroFacturacionGenerator()

            with st.spinner("⏳ Generando archivo con todos los profesionales..."):
                temp_output_all = tempfile.NamedTemporaryFile(delete=False, suffix="_TODOS.xlsx")
                temp_output_all_path = temp_output_all.name
                temp_output_all.close()

                # Pasamos la lista completa de terapeutas
                generador.generar_filtrado_por_profesional(temp_input_path, temp_output_all_path, nombres_profesionales)

            st.success("✅ Archivo generado con TODOS los profesionales. Descárgalo a continuación:")

            with open(temp_output_all_path, "rb") as f:
                st.download_button(
                    label="📥 Descargar TODOS",
                    data=f,
                    file_name="CUADRO_TODOS.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_todos"
                )

    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {e}")
    finally:
        os.remove(temp_input_path)
