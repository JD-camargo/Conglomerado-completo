# app.py
import streamlit as st
from CuadroFacturacionGenerator import CuadroFacturacionGenerator

st.set_page_config(page_title="Generador de Cuadro de Facturación", layout="centered")

st.title("📊 Generador de Cuadro de Facturación")

# Formulario para ingresar datos
st.subheader("Ingrese los datos de facturación")

datos = []
num_items = st.number_input("Número de ítems", min_value=1, max_value=20, step=1)

for i in range(num_items):
    st.markdown(f"### Ítem {i+1}")
    descripcion = st.text_input(f"Descripción {i+1}", key=f"desc_{i}")
    cantidad = st.number_input(f"Cantidad {i+1}", min_value=1, key=f"cant_{i}")
    valor_unitario = st.number_input(f"Valor unitario {i+1}", min_value=0.0, format="%.2f", key=f"vu_{i}")

    if descripcion:
        datos.append({
            "item": i+1,
            "descripcion": descripcion,
            "cantidad": cantidad,
            "valor_unitario": valor_unitario
        })

# Botón para generar archivo
if st.button("Generar Excel"):
    if datos:
        generator = CuadroFacturacionGenerator()
        archivo = generator.generar(datos)

        with open(archivo, "rb") as f:
            st.download_button(
                label="📥 Descargar Cuadro de Facturación",
                data=f,
                file_name=archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("Debe ingresar al menos un ítem con descripción.")
