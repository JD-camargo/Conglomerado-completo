import pandas as pd

class CuadroFacturacionGenerator:
    def __init__(self):
        pass

    def generar_filtrado_por_profesional(self, input_path, output_path, lista_profesionales):
        """
        Genera un cuadro de facturación filtrado por uno o varios profesionales.

        Parámetros:
        - input_path: ruta al archivo Excel original (conglomerado).
        - output_path: ruta donde se guardará el Excel generado.
        - lista_profesionales: lista con uno o varios nombres de profesionales.
        """

        # Cargar hoja "CONGLOMERADO"
        df = pd.read_excel(input_path, sheet_name="CONGLOMERADO", engine="openpyxl")

        # Filtrar solo los profesionales seleccionados
        df_filtrado = df[df["NOMBRE DEL PROFESIONAL"].isin(lista_profesionales)]

        # (Opcional) Aquí podrías hacer cálculos adicionales, ej:
        # df_filtrado["VALOR_FACTURAR"] = df_filtrado["SESIONES"] * df_filtrado["TARIFA"]

        # Guardar resultado en Excel
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df_filtrado.to_excel(writer, sheet_name="FACTURACION", index=False)
