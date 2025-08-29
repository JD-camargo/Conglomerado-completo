import streamlit as st
import pandas as pd
import tempfile
import os
import unicodedata
import re
from CuadroFacturacionGenerator import CuadroFacturacionGenerator

st.set_page_config(page_title="Generador de Cuadro de Facturación", layout="centered")
st.title("🧾 Generador de Cuadro de Facturación")
st.markdown("Sube el Excel del conglomerado, elige un profesional o genera el archivo de **todos**.")

# ----- helpers locales para detectar la columna del profesional -----
def _norm(s):
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^a-z0-9 ]+", "", s)
    return s

def _find_col(candidates, cols):
    norm_map = {_norm(c): c for c in cols}
    for cand in candidates:
        n = _norm(cand)
        if n in norm_map:
            return norm_map[n]
    for cand in candidates:
        n = _norm(cand)
        for k, original in norm_map.items():
            if n in k or k in n:
                return original
    return None

CANDIDATAS_PRO = [
    "NOMBRE DEL PROFESIONAL",
    "Nombre completo de profesional",
    "Nombre del profesional",
    "Profesional",
    "PROFESIONAL",
    "Terapeuta",
    "Nombre completo del profesional",
]

uploaded_file = st.file_uploader("📤 Cargar archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_input:
        temp_input.write(uploaded_file.read())
        temp_input_path = temp_input.name

    try:
        # Intentar leer CONGLOMERADO; si no, primera hoja
        xls = pd.ExcelFile(temp_input_path, engine="openpyxl")
        sheet = "CONGLOMERADO" if "CONGLOMERADO" in xls.sheet_names else xls.sheet_names[0]
        df_preview = pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")

        # Detectar columna del profesional automáticamente
        col_prof = _find_col(CANDIDATAS_PRO, df_preview.columns)
        if not col_prof:
            st.error(
                "No pude detectar la columna del profesional. "
                f"Columnas disponibles: {list(df_preview.columns)}"
            )
            st.stop()

        # Nombres para el selectbox
        nombres_profesionales = sorted(df_preview[col_prof].dropna().astype(str).unique())

        st.caption(f"Columna detectada para profesional: **{col_prof}**")

        # Selección individual
        nombre_seleccionado = st.selectbox("👤 Selecciona el profesional:", nombres_profesionales)

        cols_btn = st.columns(2)
        generador = CuadroFacturacionGenerator()

        # Botón individual
        if nombre_seleccionado and cols_btn[0].button("🚀 Generar archivo individual"):
            with st.spinner("⏳ Generando archivo..."):
                temp_output = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{_norm(nombre_seleccionado)}.xlsx")
                temp_output_path = temp_output.name
                temp_output.close()
                generador.generar_filtrado_por_profesional(
                    temp_input_path, temp_output_path, [nombre_seleccionado]
                )
            st.success(f"✅ Archivo generado para {nombre_seleccionado}.")
            with open(temp_output_path, "rb") as f:
                st.download_button(
                    label=f"📥 Descargar {nombre_seleccionado}",
                    data=f,
                    file_name=f"CUADRO_{nombre_seleccionado.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{_norm(nombre_seleccionado)}",
                )

        # Botón todos
        if cols_btn[1].button("🚀 Generar archivo con TODOS"):
            with st.spinner("⏳ Generando archivo con todos..."):
                temp_output_all = tempfile.NamedTemporaryFile(delete=False, suffix="_TODOS.xlsx")
                temp_output_all_path = temp_output_all.name
                temp_output_all.close()
                generador.generar_filtrado_por_profesional(
                    temp_input_path, temp_output_all_path, nombres_profesionales
                )
            st.success("✅ Archivo generado con TODOS los profesionales.")
            with open(temp_output_all_path, "rb") as f:
                st.download_button(
                    label="📥 Descargar TODOS",
                    data=f,
                    file_name="CUADRO_TODOS.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_todos",
                )

    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {e}")
    finally:
        os.remove(temp_input_path)
