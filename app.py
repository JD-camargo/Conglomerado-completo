# app.py
import streamlit as st
import tempfile, os
from CuadroFacturacionGenerator import CuadroFacturacionGenerator

st.set_page_config(page_title="Generador de Cuadro de Facturación", layout="centered")
st.title("🧾 Generador de Cuadro de Facturación - IPS")
st.markdown("Sube el Excel **CONGLOMERADO**, elige un profesional o descarga el cuadro de todas las profesionales.")

uploaded_file = st.file_uploader("📤 Cargar archivo Excel (.xlsx)", type=["xlsx"])
if not uploaded_file:
    st.info("Sube el archivo del conglomerado para ver opciones.")
    st.stop()

# Guardar temporalmente el archivo subido
with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
    tmp.write(uploaded_file.read())
    temp_input_path = tmp.name

generador = CuadroFacturacionGenerator()

try:
    # Listar profesionales detectados
    nombres, col_prof = generador.listar_profesionales(temp_input_path)
    st.caption(f"Columna detectada para profesional: **{col_prof}**")
    st.dataframe  # placeholder

    # Mostrar vista previa (primeras filas)
    import pandas as pd
    preview_df = pd.read_excel(temp_input_path, sheet_name="CONGLOMERADO", engine="openpyxl") if "CONGLOMERADO" in pd.ExcelFile(temp_input_path, engine="openpyxl").sheet_names else pd.read_excel(temp_input_path, engine="openpyxl")
    st.subheader("Vista previa del archivo (primeras 8 filas)")
    st.dataframe(preview_df.head(8))

    # UI: selección individual y botones
    nombre_seleccionado = st.selectbox("👤 Selecciona el profesional:", nombres)

    c1, c2, c3 = st.columns([1,1,1])

    # Botón: generar/descargar individual
    if c1.button("📥 Descargar seleccionado"):
        with st.spinner("Generando archivo individual..."):
            import tempfile, os
            temp_out = tempfile.NamedTemporaryFile(delete=False, suffix=f"_{nombre_seleccionado.replace(' ', '_')}.xlsx")
            temp_out_path = temp_out.name
            temp_out.close()
            generador.generar_filtrado(temp_input_path, temp_out_path, [nombre_seleccionado])
            # leer bytes y eliminar archivo temporal
            with open(temp_out_path, "rb") as f:
                data = f.read()
            os.remove(temp_out_path)
        st.success(f"Archivo listo: {nombre_seleccionado}")
        st.download_button(
            label=f"📥 Descargar {nombre_seleccionado}",
            data=data,
            file_name=f"CUADRO_{nombre_seleccionado.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_individual_{nombre_seleccionado}"
        )

    # Botón: generar/descargar todo combinado (una hoja)
    if c2.button("📥 Descargar TODOS (combinado)"):
        with st.spinner("Generando archivo combinado con todos los profesionales..."):
            temp_out_all = tempfile.NamedTemporaryFile(delete=False, suffix="_TODOS_COMBINADO.xlsx")
            temp_out_all_path = temp_out_all.name
            temp_out_all.close()
            generador.generar_filtrado(temp_input_path, temp_out_all_path, nombres)
            with open(temp_out_all_path, "rb") as f:
                data_all = f.read()
            os.remove(temp_out_all_path)
        st.success("Archivo combinado listo")
        st.download_button(
            label="📥 Descargar CUADRO_TODOS.xlsx",
            data=data_all,
            file_name="CUADRO_TODOS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_todos_combined"
        )

    # Botón: generar workbook con hoja por profesional
    if c3.button("📥 Descargar TODOS (por hojas)"):
        with st.spinner("Generando workbook por profesional (una hoja por terapeuta)..."):
            temp_out_wb = tempfile.NamedTemporaryFile(delete=False, suffix="_TODOS_POR_HOJA.xlsx")
            temp_out_wb_path = temp_out_wb.name
            temp_out_wb.close()
            generador.generar_workbook_por_profesional(temp_input_path, temp_out_wb_path, nombres)
            with open(temp_out_wb_path, "rb") as f:
                data_wb = f.read()
            os.remove(temp_out_wb_path)
        st.success("Workbook listo (una hoja por terapeuta)")
        st.download_button(
            label="📥 Descargar CUADRO_TODOS_POR_HOJA.xlsx",
            data=data_wb,
            file_name="CUADRO_TODOS_POR_HOJA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_todos_por_hoja"
        )

except Exception as e:
    st.error(f"❌ Error: {e}")

finally:
    # borrar el archivo subido temporal
    try:
        if os.path.exists(temp_input_path):
            os.remove(temp_input_path)
    except:
        pass
