import pandas as pd
from collections import defaultdict
from datetime import datetime

class CuadroFacturacionGenerator:

    def _formatear_fechas(self, fechas):
        fechas_ordenadas = sorted(fechas, key=lambda x: datetime.strptime(x, "%Y-%m-%d"))
        fechas_dict = defaultdict(list)

        meses_es = {
            "January": "enero", "February": "febrero", "March": "marzo",
            "April": "abril", "May": "mayo", "June": "junio",
            "July": "julio", "August": "agosto", "September": "septiembre",
            "October": "octubre", "November": "noviembre", "December": "diciembre"
        }

        for fecha in fechas_ordenadas:
            dt = datetime.strptime(fecha, "%Y-%m-%d")
            mes = dt.strftime("%B")
            dia = str(dt.day)
            fechas_dict[mes].append(dia)

        fechas_formateadas = []
        for mes, dias in fechas_dict.items():
            mes_es = meses_es.get(mes, mes)
            fechas_formateadas.append(f"{', '.join(dias)} {mes_es}")

        return ", ".join(fechas_formateadas)


    def _procesar_dataframe(self, df):
        df_filtered = df[[
            "DOC PROFESIONAL", "NOMBRE DEL PROFESIONAL", "Tipo de nota",
            "Documento", "NOMBRE USUARIO", "FECHA INI AUT", "FECHA FINAL", "AUT", "FECHA ATENCION"
        ]]

        sesiones_dict = defaultdict(lambda: {"count": 0, "fechas": []})

        for _, row in df_filtered.iterrows():
            clave = (
                row["DOC PROFESIONAL"], row["NOMBRE DEL PROFESIONAL"], row["Tipo de nota"],
                row["Documento"], row["NOMBRE USUARIO"], row["AUT"], row["FECHA INI AUT"],
                row["FECHA FINAL"]
            )
            sesiones_dict[clave]["count"] += 1
            sesiones_dict[clave]["fechas"].append(row["FECHA ATENCION"].date().isoformat())

        datos_expandidos = []
        for clave, valores in sesiones_dict.items():
            doc_profesional, nombre_profesional, tipo_nota, documento, nombre_usuario, autorizacion, fecha_ini_aut, fecha_final = clave
            fechas_atencion = self._formatear_fechas(valores["fechas"])

            fila = [
                doc_profesional, nombre_profesional, tipo_nota,
                nombre_usuario, documento,
                autorizacion, fecha_ini_aut, fecha_final,
                valores["count"], fechas_atencion
            ]
            datos_expandidos.append(fila)

        df_grouped = pd.DataFrame(datos_expandidos, columns=[
            "CC Profesional", "Nombre completo de profesional", "Area",
            "Nombre completo de Usuario", "Doc Usuario", "No Autorización",
            "Fecha Inicial", "Fecha Final", "NO de sesiones",
            "Fechas de atención DIAS Y MESES"
        ])

        # columnas extra
        df_grouped.insert(7, "SES AUTOR", "")
        df_grouped.insert(11, "AUTOR", "")
        df_grouped.insert(12, "GLOSAS", "")
        df_grouped.insert(13, "RECONOCE LA EMPRESA", "")
        df_grouped["Valor"] = df_grouped["NO de sesiones"] * 4500
        df_grouped["TIPO CONTRATO (OPS O NOMINA)"] = "Nómina"

        return df_grouped


    def generar_filtrado_por_profesional(self, conglomerado_path, output_path, nombre_profesional):
        df = pd.read_excel(conglomerado_path, sheet_name="CONGLOMERADO", engine="openpyxl")
        df = df[df["NOMBRE DEL PROFESIONAL"] == nombre_profesional]

        if df.empty:
            raise ValueError(f"No se encontraron registros para el profesional: {nombre_profesional}")

        df_grouped = self._procesar_dataframe(df)
        df_grouped.to_excel(output_path, sheet_name="CUADRO SESIONES REALIZADAS", index=False, engine="openpyxl")


    def generar_todos(self, conglomerado_path, output_path):
        df = pd.read_excel(conglomerado_path, sheet_name="CONGLOMERADO", engine="openpyxl")
        df_grouped = self._procesar_dataframe(df)
        df_grouped.to_excel(output_path, sheet_name="CUADRO SESIONES REALIZADAS", index=False, engine="openpyxl")
