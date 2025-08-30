# CuadroFacturacionGenerator.py
import pandas as pd
import unicodedata
import re
from typing import List
from pathlib import Path

# ---------------- utilidades ----------------
def _norm(texto: str) -> str:
    if texto is None:
        return ""
    s = unicodedata.normalize("NFKD", str(texto))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^a-z0-9 ]+", "", s)
    return s

def _find_col(candidates: List[str], df_columns: List[str]):
    """Busca la primera columna del df_columns que coincida (normalizada) con cualquiera de candidates."""
    norm_map = {_norm(col): col for col in df_columns}
    # Busqueda exacta normalizada
    for cand in candidates:
        n = _norm(cand)
        if n in norm_map:
            return norm_map[n]
    # Búsqueda parcial (por si hay sufijos/prefijos)
    for cand in candidates:
        n = _norm(cand)
        for k, original in norm_map.items():
            if n in k or k in n:
                return original
    return None

def _open_conglomerado(input_path: str) -> pd.DataFrame:
    xls = pd.ExcelFile(input_path, engine="openpyxl")
    sheet = "CONGLOMERADO" if "CONGLOMERADO" in xls.sheet_names else xls.sheet_names[0]
    return pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")

def _safe_sheet_name(name: str) -> str:
    # Quitar caracteres inválidos y limitar a 31 chars
    s = re.sub(r'[:\\/*?\[\]]', '_', str(name))
    s = s[:31]
    if not s:
        s = "Hoja"
    return s

# ---------------- clase principal ----------------
class CuadroFacturacionGenerator:
    def __init__(self):
        # columnas finales en el orden exacto que quieres (TIPO CONTRATO al final)
        self.columnas_finales = [
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
            "Valor",
            "TIPO CONTRATO (OPS O NOMINA)",
        ]

        # sinónimos candidatos para detectar la columna del profesional
        self.candidatas_profesional = [
            "Nombre completo de profesional",
            "NOMBRE DEL PROFESIONAL",
            "Nombre del profesional",
            "Profesional",
            "PROFESIONAL",
            "Terapeuta",
            "Nombre completo del profesional",
        ]

        # opcional: sinónimos para otras columnas si quieres expandir
        self.synonyms = {
            "CC Profesional": ["CC Profesional", "cc profesional", "cc", "cedula profesional", "cedula"],
            "Nombre completo de profesional": ["Nombre completo de profesional", "nombre del profesional", "profesional", "terapeuta"],
            "Area": ["Area", "Área", "especialidad", "departamento"],
            "Nombre completo de Usuario": ["Nombre completo de Usuario", "nombre del usuario", "usuario"],
            "Doc Usuario": ["Doc Usuario", "Documento Usuario", "Documento del Usuario", "doc usuario"],
            "No Autorización": ["No Autorización", "No Autorizacion", "No. Autorización", "no autorizacion"],
            "SES AUTOR": ["SES AUTOR", "SES_AUTOR", "ses autor"],
            "Fecha Inicial": ["Fecha Inicial", "Fecha Inicio", "FechaInicio"],
            "Fecha Final": ["Fecha Final", "Fecha Fin", "FechaFin"],
            "NO de sesiones": ["NO de sesiones", "No de sesiones", "Nº de sesiones", "numero de sesiones", "no sesiones"],
            "AUTOR": ["AUTOR", "Autor"],
            "GLOSAS": ["GLOSAS", "Glosas"],
            "RECONOCE LA EMPRESA": ["RECONOCE LA EMPRESA", "Reconoce la empresa"],
            "Fechas de atención DIAS Y MESES": ["Fechas de atención DIAS Y MESES", "Fechas de atencion", "Fechas de atención"],
            "Valor": ["Valor", "VALOR", "Valor a facturar", "precio"],
            "TIPO CONTRATO (OPS O NOMINA)": ["TIPO CONTRATO (OPS O NOMINA)", "TIPO CONTRATO", "Tipo de contrato", "Tipo contrato", "Contrato"],
        }

    # ---------------- utilitarios públicos ----------------
    def listar_profesionales(self, input_path: str):
        """Devuelve (lista_nombres, nombre_columna_detectada). Lanzará KeyError si no encuentra la columna del profesional."""
        df = _open_conglomerado(input_path)
        col_prof = _find_col(self.candidatas_profesional, df.columns)
        if not col_prof:
            raise KeyError(f"No se encontró la columna del profesional. Columnas disponibles: {list(df.columns)}")
        nombres = sorted(df[col_prof].dropna().astype(str).unique())
        return nombres, col_prof

    def generar_filtrado(self, input_path: str, output_path: str, lista_profesionales: List[str]):
        """
        Genera un Excel (una hoja llamada FACTURACION) con las columnas en el orden deseado.
        lista_profesionales: lista con los nombres (si quieres todos, pasa la lista completa).
        """
        df = _open_conglomerado(input_path)

        # detectar columna profesional
        col_prof = _find_col(self.candidatas_profesional, df.columns)
        if not col_prof:
            raise KeyError("No se encontró la columna del profesional en el archivo.")
        lista_limpia = [str(x).strip() for x in lista_profesionales if str(x).strip()]
        df_filtrado = df[df[col_prof].astype(str).isin(lista_limpia)]

        # Construir DataFrame de salida con columnas fijas (si no existen, crear vacías)
        salida = pd.DataFrame(index=df_filtrado.index)
        for final_col in self.columnas_finales:
            # buscar la columna original correspondiente (usando synonyms si existen)
            candidates = self.synonyms.get(final_col, [final_col])
            encontrada = _find_col(candidates, df_filtrado.columns)
            if encontrada:
                salida[final_col] = df_filtrado[encontrada]
            else:
                salida[final_col] = ""  # columna vacía si no existe

        # Guardar
        salida.to_excel(output_path, sheet_name="FACTURACION", index=False)
        return output_path

    def generar_workbook_por_profesional(self, input_path: str, output_path: str, lista_profesionales: List[str]):
        """
        Genera un workbook con una hoja por profesional (nombre de hoja = terapeuta).
        Cada hoja tiene las mismas columnas en el orden deseado.
        """
        df = _open_conglomerado(input_path)
        col_prof = _find_col(self.candidatas_profesional, df.columns)
        if not col_prof:
            raise KeyError("No se encontró la columna del profesional en el archivo.")
        lista_limpia = [str(x).strip() for x in lista_profesionales if str(x).strip()]

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for nombre in lista_limpia:
                df_f = df[df[col_prof].astype(str) == nombre]
                salida = pd.DataFrame(index=df_f.index)
                for final_col in self.columnas_finales:
                    candidates = self.synonyms.get(final_col, [final_col])
                    encontrada = _find_col(candidates, df_f.columns)
                    if encontrada:
                        salida[final_col] = df_f[encontrada]
                    else:
                        salida[final_col] = ""
                sheet = _safe_sheet_name(nombre)
                # si no hay filas, aún creamos la hoja vacía con headers
                salida.to_excel(writer, sheet_name=sheet, index=False)
        return output_path
