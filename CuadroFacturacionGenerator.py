# CuadroFacturacionGenerator.py
import pandas as pd
import unicodedata
import re
from typing import List

# ---------- utilidades ----------
def _norm(texto: str) -> str:
    """Normaliza texto para comparar columnas (quita tildes, simbolos, minusculas)."""
    if texto is None:
        return ""
    s = unicodedata.normalize("NFKD", str(texto))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^a-z0-9 ]+", "", s)
    return s

def _find_col(candidates: List[str], df_columns: List[str]):
    """
    Busca la primera columna de df_columns que coincida (normalizado) con cualquier candidato.
    Devuelve el nombre original de la columna (sin normalizar) o None.
    """
    norm_map = {_norm(col): col for col in df_columns}
    # búsqueda exacta normalizada
    for cand in candidates:
        n = _norm(cand)
        if n in norm_map:
            return norm_map[n]
    # búsqueda parcial (por si hay prefijos/sufijos)
    for cand in candidates:
        n = _norm(cand)
        for k, original in norm_map.items():
            if n in k or k in n:
                return original
    return None

def _open_conglomerado(input_path: str) -> pd.DataFrame:
    """Abre la hoja 'CONGLOMERADO' si existe, si no abre la primera hoja."""
    xls = pd.ExcelFile(input_path, engine="openpyxl")
    sheet = "CONGLOMERADO" if "CONGLOMERADO" in xls.sheet_names else xls.sheet_names[0]
    return pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")

def _safe_sheet_name(name: str) -> str:
    s = re.sub(r'[:\\/*?\[\]]', '_', str(name))
    s = s[:31]
    if not s:
        s = "Hoja"
    return s

# ---------- clase principal ----------
class CuadroFacturacionGenerator:
    def __init__(self):
        # columnas finales en el orden exacto que pides (TIPO CONTRATO al final)
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

        # candidatas para detectar la columna del profesional
        self.candidatas_profesional = [
            "Nombre completo de profesional",
            "NOMBRE DEL PROFESIONAL",
            "Nombre del profesional",
            "Profesional",
            "PROFESIONAL",
            "Terapeuta",
            "Nombre completo del profesional",
        ]

        # sinónimos útiles para mapear columnas existentes a las columnas finales
        self.synonyms = {
            "CC Profesional": ["CC Profesional", "cc profesional", "cc", "cedula profesional", "cedula"],
            "Nombre completo de profesional": ["Nombre completo de profesional", "nombre del profesional", "profesional", "terapeuta"],
            "Area": ["Area", "Área", "especialidad", "departamento"],
            "Nombre completo de Usuario": ["Nombre completo de Usuario", "nombre del usuario", "usuario"],
            "Doc Usuario": ["Doc Usuario", "Documento Usuario", "Documento del Usuario", "doc usuario", "documento usuario"],
            "No Autorización": ["No Autorización", "No Autorizacion", "No. Autorización", "no autorizacion", "num autorizacion"],
            "SES AUTOR": ["SES AUTOR", "SES_AUTOR", "ses autor", "sesiones autorizadas"],
            "Fecha Inicial": ["Fecha Inicial", "Fecha Inicio", "FechaInicio", "fecha inicio"],
            "Fecha Final": ["Fecha Final", "Fecha Fin", "FechaFin", "fecha fin"],
            "NO de sesiones": ["NO de sesiones", "No de sesiones", "Nº de sesiones", "numero de sesiones", "no sesiones"],
            "AUTOR": ["AUTOR", "Autor"],
            "GLOSAS": ["GLOSAS", "Glosas"],
            "RECONOCE LA EMPRESA": ["RECONOCE LA EMPRESA", "Reconoce la empresa"],
            "Fechas de atención DIAS Y MESES": ["Fechas de atención DIAS Y MESES", "Fechas de atencion", "Fechas de atención", "fechas de atencion"],
            "Valor": ["Valor", "VALOR", "Valor a facturar", "precio"],
            "TIPO CONTRATO (OPS O NOMINA)": ["TIPO CONTRATO (OPS O NOMINA)", "TIPO CONTRATO", "Tipo de contrato", "Tipo contrato", "Contrato"],
        }

    # ---------------- métodos públicos ----------------
    def listar_profesionales(self, input_path: str):
        """
        Retorna (lista_nombres, nombre_columna_detectada).
        Lanza KeyError si no encuentra la columna del profesional.
        """
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
        df_filtrado = df[df[col_prof].astype(str).isin(lista_limpia)].reset_index(drop=True)

        # si no hay filas, crear un DataFrame vacío con los headers deseados
        if df_filtrado.shape[0] == 0:
            salida = pd.DataFrame(columns=self.columnas_finales)
        else:
            salida = pd.DataFrame()
            for final_col in self.columnas_finales:
                candidates = self.synonyms.get(final_col, [final_col])
                encontrada = _find_col(candidates, df_filtrado.columns)
                if encontrada:
                    # usar .reset_index(drop=True) para alinear por posición y evitar issues con índices
                    salida[final_col] = df_filtrado[encontrada].reset_index(drop=True)
                else:
                    salida[final_col] = [""] * len(df_filtrado)

        # Guardar Excel con la hoja FACTURACION
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
                df_f = df[df[col_prof].astype(str) == nombre].reset_index(drop=True)
                if df_f.shape[0] == 0:
                    salida = pd.DataFrame(columns=self.columnas_finales)
                else:
                    salida = pd.DataFrame()
                    for final_col in self.columnas_finales:
                        candidates = self.synonyms.get(final_col, [final_col])
                        encontrada = _find_col(candidates, df_f.columns)
                        if encontrada:
                            salida[final_col] = df_f[encontrada].reset_index(drop=True)
                        else:
                            salida[final_col] = [""] * len(df_f)

                sheet = _safe_sheet_name(nombre)
                salida.to_excel(writer, sheet_name=sheet, index=False)

        return output_path
