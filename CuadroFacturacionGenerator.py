import pandas as pd
import os

class CuadroFacturacionGenerator:
    def __init__(self, file):
        self.file = file
        self.df = None

    def load_data(self):
        """Carga el archivo de Excel respetando los encabezados tal como vienen"""
        try:
            self.df = pd.read_excel(self.file, dtype=str)  # Forzamos a string para evitar los "1"
            self.df = self.df.fillna("")  # Rellenamos vacíos con ""
        except Exception as e:
            raise Exception(f"Error al leer el archivo: {e}")

    def generar_por_profesional(self, output_dir="output"):
        """Genera un archivo por cada profesional"""
        if self.df is None:
            raise Exception("Primero debe cargar los datos con load_data()")

        os.makedirs(output_dir, exist_ok=True)

        # Obtener lista de profesionales únicos
        profesionales = self.df["Nombre completo de profesional"].unique()

        archivos = []
        for profesional in profesionales:
            df_pro = self.df[self.df["Nombre completo de profesional"] == profesional].copy()

            # Mover la columna 'TIPO CONTRATO (OPS O NOMINA)' al final
            if "TIPO CONTRATO (OPS O NOMINA)" in df_pro.columns:
                cols = [c for c in df_pro.columns if c != "TIPO CONTRATO (OPS O NOMINA)"]
                cols.append("TIPO CONTRATO (OPS O NOMINA)")
                df_pro = df_pro[cols]

            filename = os.path.join(output_dir, f"{profesional}.xlsx")
            df_pro.to_excel(filename, index=False)
            archivos.append(filename)

        return archivos

    def generar_consolidado(self, output_dir="output"):
        """Genera un solo archivo con todo"""
        if self.df is None:
            raise Exception("Primero debe cargar los datos con load_data()")

        os.makedirs(output_dir, exist_ok=True)

        # Mover la columna 'TIPO CONTRATO (OPS O NOMINA)' al final
        df = self.df.copy()
        if "TIPO CONTRATO (OPS O NOMINA)" in df.columns:
            cols = [c for c in df.columns if c != "TIPO CONTRATO (OPS O NOMINA)"]
            cols.append("TIPO CONTRATO (OPS O NOMINA)")
            df = df[cols]

        filename = os.path.join(output_dir, "Consolidado_Facturacion.xlsx")
        df.to_excel(filename, index=False)
        return filename
