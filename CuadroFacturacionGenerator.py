import pandas as pd

class CuadroFacturacionGenerator:
    def __init__(self):
        pass

    def generar_filtrado_por_profesional(self, input_path, output_path, lista_profesionales):
        """
        Genera un cuadro de facturación filtrado por uno o varios profesionales,
        manteniendo las mismas columnas del archivo original.
        """

        # Cargar hoja "CONGLOMERADO"
        df = pd.read_excel(input_path, sheet_name="CONGLOMERADO", engine="openpyxl")

        # Filtrar solo los profesionales seleccionados
        df_filtrado = df[df["Nombre completo de profesional"].isin(lista_profesionales)]

        # 🔹 Definir las columnas en el mismo orden que el archivo original
        columnas_salida = [
            "TIPO CONTRATO (OPS O NOMINA)",
            "CC Profesional",
            "Nombre completo de profesional",
            "Area",
            "Nombre completo de Usuario",
            "Doc Usuario",
            "No Autorización",
            "SES AUTOR",
            "Fecha Inicial",
            "Fecha Final",
            "NO de sesiones",
            "AUTOR",
            "GLOSAS",
            "RECONOCE LA EMPRESA",
            "Fechas de atención DIAS Y MESES",
            "Valor"
        ]

        # Asegurarnos de que solo se guarden esas columnas en el mismo orden
        df_filtrado = df_filtrado[columnas_salida]

        # Guardar resultado en Excel
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df_filtrado.to_excel(writer, sheet_name="FACTURACION", index=False)
