import pandas as pd
import os

class CuadroFacturacionGenerator:
    def __init__(self, file):
        self.file = file
        self.df = None

    def load_data(self):
        """Carga el archivo respetando encabezados reales"""
        try:
            # 👇 Importante: header=0 para tomar la primera fila como encabezados
            # dtype=str para que no convierta a número y no aparezcan "1"
            self.df = pd.read_excel(self.file, header=0, dtype=str)
            self.df = self.df.fillna("")  # reemplaza NaN por vacío
        except Exception as e:
            raise Exception(f"Error al leer el archivo: {e}")

    def generar_por_profesional(self, output_dir="output"):
        """Genera un Excel por cada profesional"""
        if self.df is None:
            raise Exception("Primero debe cargar los datos con load_data()")

        os.makedirs(output_dir, exist_ok=True)
        archivos = []

        for profesional in self.df["Nombre completo de profesional"].unique():
            df_pro = self.df[self.df["Nombre completo de profesional"] == profesional].copy()

            # mover columna "TIPO CONTRATO (OPS O NOMINA)" al final si existe
            if "TIPO CONTRATO (OPS O NOMINA)" in df_pro.columns:
                cols = [c for c in df_pro.columns if c != "TIPO CONTRATO (OPS O NOMINA)"]
                cols.append("TIPO CONTRATO (OPS O NOMINA)")
                df_pro = df_pro[cols]

            filename = os.path.join(output_dir, f"{profesional}.xlsx")
            df_pro.to_excel(filename, index=False)
            archivos.append(filename)

        return archivos

    def generar_consolidado(self, output_dir="output"):
        """Genera un Excel con todos los registros"""
        if self.df is None:
            raise Exception("Primero debe cargar los datos con load_data()")

        os.makedirs(output_dir, exist_ok=True)
        df = self.df.copy()

        if "TIPO CONTRATO (OPS O NOMINA)" in df.columns:
            cols = [c for c in df.columns if c != "TIPO CONTRATO (OPS O NOMINA)"]
            cols.append("TIPO CONTRATO (OPS O NOMINA)")
            df = df[cols]

        filename = os.path.join(output_dir, "Consolidado_Facturacion.xlsx")
        df.to_excel(filename, index=False)
        return filename
