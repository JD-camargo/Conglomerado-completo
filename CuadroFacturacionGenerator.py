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
            "NO de sesiones",
            "AUTOR",
            "GLOSAS",
            "RECONOCE LA EMPRESA",
            "Fechas de atención DIAS Y MESES",
            "Valor",
        ]
        # Sinónimos para detectar la columna del profesional
        self.candidatas_profesional = [
            "NOMBRE DEL PROFESIONAL",
            "Nombre completo de profesional",
            "Nombre del profesional",
            "Profesional",
            "PROFESIONAL",
            "Terapeuta",
            "Nombre completo del profesional",
        ]

    def generar_filtrado_por_profesional(self, input_path, output_path, lista_profesionales):
        """
        Genera un Excel filtrado por uno o varios profesionales.
        - Mantiene las columnas en el mismo orden que self.columnas_salida_deseadas,
          usando las que existan en el archivo origen.
        """
        df = _open_conglomerado(input_path)

        # Detectar columna del profesional de forma robusta
        col_prof = _find_col(self.candidatas_profesional, df.columns)
        if not col_prof:
            raise KeyError(
                "No se encontró la columna del profesional. "
                f"Columnas disponibles: {list(df.columns)}"
            )

        # Filtrado
        lista_limpia = [str(x).strip() for x in lista_profesionales if str(x).strip()]
        df_filtrado = df[df[col_prof].astype(str).isin(lista_limpia)]

        # Mapear columnas de salida a las existentes (en el mismo orden)
        columnas_existentes = []
        for deseada in self.columnas_salida_deseadas:
            encontrada = _find_col([deseada], df_filtrado.columns)
            if encontrada:
                columnas_existentes.append(encontrada)

        # Si no se encontró ninguna de las deseadas, exportar tal cual filtrado
        if not columnas_existentes:
            columnas_existentes = list(df_filtrado.columns)

        # Guardar resultado
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df_filtrado[columnas_existentes].to_excel(
                writer, sheet_name="FACTURACION", index=False
            )
