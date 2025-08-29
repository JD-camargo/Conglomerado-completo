import pandas as pd
import unicodedata
import re

class CuadroFacturacionGenerator:
    def __init__(self, sheet_name="CONGLOMERADO"):
        self.sheet_name = sheet_name

        # Columnas de salida en el orden que mencionaste
        self.columnas_canonicas = [
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

        # Sinónimos mínimos para detectar columnas (puedes ampliar si lo necesitas)
        self.sinonimos = {
            "Nombre completo de profesional": [
                "NOMBRE DEL PROFESIONAL",
                "Nombre del profesional",
                "Profesional",
                "Profesional tratante",
                "Therapist Name",
                "Professional Name",
            ],
        }

    # -------- utilidades internas --------
    @staticmethod
    def _norm(s):
        s = str(s)
        s = unicodedata.normalize("NFD", s)
        s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")  # quita acentos
        s = s.lower()
        s = re.sub(r"[^a-z0-9]+", " ", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s

    def _colmap(self, cols):
        return {self._norm(c): c for c in cols}

    def _match(self, colmap, objetivo):
        candidates = [objetivo] + self.sinonimos.get(objetivo, [])
        for cand in candidates:
            key = self._norm(cand)
            if key in colmap:
                return colmap[key]
        return None

    def detectar_columna_profesional(self, df: pd.DataFrame) -> str:
        colmap = self._colmap(df.columns)
        # objetivo canónico
        found = self._match(colmap, "Nombre completo de profesional")
        if found:
            return found
        # fallback duro por si acaso
        if self._norm("NOMBRE DEL PROFESIONAL") in colmap:
            return colmap[self._norm("NOMBRE DEL PROFESIONAL")]
        raise KeyError(
            "No se encontró la columna del profesional. Asegúrate de que exista "
            "'Nombre completo de profesional' o 'NOMBRE DEL PROFESIONAL'."
        )

    # -------- método principal --------
    def generar_filtrado_por_profesional(self, input_path, output_path, lista_profesionales):
        # Lee la hoja
        df = pd.read_excel(input_path, sheet_name=self.sheet_name, engine="openpyxl")

        # Detecta columna del profesional
        col_prof = self.detectar_columna_profesional(df)

        # Filtro por lista (limpiando espacios)
        nombres = [str(n).strip() for n in lista_profesionales]
        df_filtrado = df[df[col_prof].astype(str).str.strip().isin(nombres)].copy()

        # Construir salida con tus encabezados; si alguno no existe, lo deja vacío
        colmap = self._colmap(df_filtrado.columns)
        salida = pd.DataFrame()
        for canon in self.columnas_canonicas:
            origen = self._match(colmap, canon)
            if origen is None:
                # si no está, crea la columna vacía para mantener la estructura
                salida[canon] = pd.Series([pd.NA] * len(df_filtrado))
            else:
                salida[canon] = df_filtrado[origen]

        # Guardar
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            salida.to_excel(writer, sheet_name="FACTURACION", index=False)
