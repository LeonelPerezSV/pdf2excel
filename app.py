import streamlit as st
from pdf2excel import pdf_a_excel
import tempfile
import os
import zipfile
from pathlib import Path

st.set_page_config(page_title="PDF → Excel", page_icon="📑")
st.title("📑 Conversor Masivo PDF ➜ Excel")

# ► Permite seleccionar múltiples PDFs a la vez
drop_zone = st.file_uploader(
    "Sube uno o más archivos PDF", type="pdf", accept_multiple_files=True
)


if drop_zone and st.button("Convertir todos"):
    with st.spinner("Convirtiendo archivos…"):
        temp_dir = Path(tempfile.mkdtemp())
        excels_generados = []

        # ─── Guarda cada PDF y lo convierte ────────────────────────────────────
        for archivo in drop_zone:
            try:
                # 1) Guardar PDF temporalmente
                pdf_path = temp_dir / archivo.name
                with open(pdf_path, "wb") as f:
                    f.write(archivo.read())

                # 2) Convertir a Excel
                xlsx_path = pdf_path.with_suffix(".xlsx")
                pdf_a_excel(pdf_path, xlsx_path)
                excels_generados.append(xlsx_path)

                st.success(f"✅ {archivo.name} → {xlsx_path.name}")
            except Exception as err:
                st.error(f"❌ Error procesando {archivo.name}: {err}")

        # ─── Ofrecer la descarga ──────────────────────────────────────────────
        if not excels_generados:
            st.warning("No se generaron archivos Excel.")
        elif len(excels_generados) == 1:
            # Caso 1: solo un Excel → descarga directa
            xlsx_path = excels_generados[0]
            with open(xlsx_path, "rb") as f:
                st.download_button(
                    label="⬇️ Descargar Excel", data=f, file_name=xlsx_path.name
                )
        else:
            # Caso 2: varios Excel → empaquetar en ZIP
            zip_path = temp_dir / "excels_convertidos.zip"
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
                for xlsx in excels_generados:
                    zipf.write(xlsx, arcname=xlsx.name)

            with open(zip_path, "rb") as f:
                st.download_button(
                    label="📦 Descargar ZIP con todos los Excel",
                    data=f,
                    file_name="excels_convertidos.zip",
                )

        # Nota: la carpeta Descargas depende del navegador del usuario.
        # El archivo se guardará en la carpeta de descargas predeterminada
        # del navegador, normalmente "Descargas" en Windows.

    # Limpieza opcional (se mantiene si quieres conservar los temporales)
    # import shutil; shutil.rmtree(temp_dir, ignore_errors=True)
