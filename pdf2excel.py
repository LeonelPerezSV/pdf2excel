# pdf2excel.py
import tabula
import pandas as pd
from pathlib import Path

def _extraer_tablas(pdf: Path, *, encoding: str = "cp1252"):
    """
    Devuelve una lista de DataFrames con las tablas detectadas.
    1º intenta con lattice=True; si no halla nada, prueba con stream=True.
    """
    # --- primer intento: lattice (líneas visibles) --------------------------
    tablas = tabula.read_pdf(
        pdf.as_posix(),
        pages="all",
        multiple_tables=True,
        lattice=True,
        encoding=encoding            # evita UnicodeDecodeError en Windows
    )
    if tablas:
        return tablas

    # --- segundo intento: stream (sin líneas de tabla) ---------------------
    tablas = tabula.read_pdf(
        pdf.as_posix(),
        pages="all",
        multiple_tables=True,
        lattice=False,               # ≈ stream
        encoding=encoding
    )
    return tablas

def pdf_a_excel(pdf_path: str | Path, excel_path: str | Path) -> Path:
    """
    Convierte todas las tablas detectadas en un PDF a un archivo Excel.
    Devuelve la ruta del Excel generado.
    """
    pdf_path = Path(pdf_path).expanduser().resolve()
    excel_path = Path(excel_path).expanduser().resolve()

    if not pdf_path.exists():
        raise FileNotFoundError(f"⛔ El archivo no existe: {pdf_path}")

    tablas = _extraer_tablas(pdf_path)
    if not tablas:
        raise RuntimeError("⛔ No se detectaron tablas. "
                           "Prueba con otro PDF o verifica que no sea escaneado.")

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        for i, df in enumerate(tablas, start=1):
            df.to_excel(writer, sheet_name=f"Tabla_{i}", index=False)

    return excel_path

# --- ejecución directa desde consola --------------------------------------
if __name__ == "__main__":
    SRC = Path(r"C:\Users\Contabilidad_04\PycharmProjects\pdf2excel\mifactura.pdf")
    DST = SRC.with_suffix(".xlsx")
    outfile = pdf_a_excel(SRC, DST)
    print(f"✅ Conversión terminada → {outfile}")
