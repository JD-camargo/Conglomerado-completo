import pandas as pd
import unicodedata
import re

# ---------- Utilidades ----------
def _norm(texto: str) -> str:
    """Normaliza: sin tildes, minúsculas, sin signos, espacios colapsados."""
    if texto is None:
        return ""
    s = unicodedata.normalize("NFKD", str(texto))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^a-z0-9 ]+", "", s)
    return s

def _find_col(candidates, df_columns):
    """Busca en columnas del DF la primera coincidencia normalizada con cualquiera de 'candidates'."""
    norm_map = {_norm(col): col for col in df_columns}
    # match exact normalizado
    for cand in candidates:
        n = _norm(cand)
        if n in norm_map:
            return norm_map[n]
    # match parcial (por si hay sufijos/prefijos)
    for cand in candidates:
        n = _norm(cand)
        for k, original in norm_map.items():
            if n in k or k in n:
                return original
    return None

def _open_conglomerado(input_path):
    """Abre la hoja CONGLOMERADO; si no existe, abre la primera hoja."""
    xls = pd.ExcelFile(input_path, engine="openpyxl")
    sheet = "CONGLOMERADO" if "CONGLOMERADO" in xls.sheet_names else xls.sheet_names[0]
    return pd.read_excel(xls, sheet_name=sheet, engine="openpyxl")


class CuadroFacturacionGenerator:
    def __init__(self):
        # Columnas deseadas en el Excel de salida (orden fijo)
        self.columnas_salida_deseadas = [
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
            "NO de sesione
