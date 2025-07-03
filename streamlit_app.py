import streamlit as st
import openpyxl
from io import BytesIO
import traceback
from copy import copy
from openpyxl.cell.cell import MergedCell

st.set_page_config(page_title="Consolidador a XLSX", layout="wide")
st.title("Consolidador REM")

st.warning(
    "**Importante:** Este método reconstruye el archivo final desde cero para garantizar que no esté dañado. "
    "Se copiarán valores, formatos de celda, anchos/altos y celdas combinadas. **Gráficos, imágenes y tablas dinámicas de la plantilla no serán transferidos.**"
)

# --- Funciones de Ayuda para Copiar la Plantilla ---
def copy_sheet_properties(source_ws, target_ws):
    """Copia propiedades de la hoja como anchos/altos y celdas combinadas."""
    for col_letter, dim in source_ws.column_dimensions.items():
        target_ws.column_dimensions[col_letter].width = dim.width
    for row_idx, dim in source_ws.row_dimensions.items():
        target_ws.row_dimensions[row_idx].height = dim.height
    for merged_range in source_ws.merged_cells.ranges:
        target_ws.merge_cells(str(merged_range))

def copy_cell(source_cell, target_cell):
    """Copia valor y todos los estilos de una celda a otra."""
    target_cell.value = source_cell.value
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

# --- Interfaz de Streamlit ---
st.header("1. Sube tu Archivo de Plantilla")
template_file = st.file_uploader(
    "Este archivo definirá la estructura y formato del resultado.",
    type=["xlsx", "xlsm"]
)

st.header("2. Sube los Archivos con Datos a Sumar")
data_files = st.file_uploader(
    "De estos archivos solo se tomarán los valores numéricos.",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True
)

if 'processed_file' not in st.session_state:
    st.session_state.processed_file = None
    st.session_state.file_name = None

if st.button("✨ Consolidar Archivos"):
    if not template_file:
        st.error("❌ Por favor, sube un archivo de plantilla.")
    elif not data_files:
        st.error("❌ Por favor, sube al menos un archivo de datos.")
    else:
        with st.spinner("Iniciando proceso..."):
            try:
                # PASO 1: Sumar los datos de los archivos de datos
                sumas_consolidadas = {}
                progress_bar = st.progress(0, text="Paso 1/3: Sumando datos...")
                for i, data_file in enumerate(data_files):
                    data_file.seek(0)
                    wb_data = openpyxl.load_workbook(data_file, data_only=True)
                    for hoja_nombre in wb_data.sheetnames:
                        if hoja_nombre not in sumas_consolidadas:
                            sumas_consolidadas[hoja_nombre] = {}
                        ws_data = wb_data[hoja_nombre]
                        for fila in ws_data.iter_rows():
                            for celda in fila:
                                if isinstance(celda.value, (int, float)):
                                    ref = celda.coordinate
                                    sumas_consolidadas[hoja_nombre][ref] = sumas_consolidadas[hoja_nombre].get(ref, 0) + celda.value
                    progress_bar.progress((i + 1) / len(data_files), text=f"Paso 1/3: Leyendo {data_file.name}")

                # PASO 2: Reconstruir la plantilla en un nuevo libro de trabajo
                progress_bar.progress(0, text="Paso 2/3: Reconstruyendo la plantilla...")
                template_file.seek(0)
                wb_template = openpyxl.load_workbook(template_file)
                wb_final = openpyxl.Workbook()
                wb_final.remove(wb_final.active)

                for i, hoja_nombre in enumerate(wb_template.sheetnames):
                    source_ws = wb_template[hoja_nombre]
                    target_ws = wb_final.create_sheet(title=hoja_nombre)
                    copy_sheet_properties(source_ws, target_ws)
                    for fila in source_ws.iter_rows():
                        for celda in fila:
                            if isinstance(celda, MergedCell):
                                continue
                            new_cell = target_ws.cell(row=celda.row, column=celda.column)
                            copy_cell(celda, new_cell)
                    progress_bar.progress((i + 1) / len(wb_template.sheetnames), text=f"Paso 2/3: Copiando hoja '{hoja_nombre}'...")

                # PASO 3: Escribir las sumas en el nuevo libro reconstruido
                progress_bar.progress(0, text="Paso 3/3: Escribiendo valores consolidados...")
                for hoja_nombre, celdas in sumas_consolidadas.items():
                    if hoja_nombre in wb_final.sheetnames:
                        ws_final = wb_final[hoja_nombre]
                        for celda_ref, valor_sumado in celdas.items():
                            try:
                                ws_final[celda_ref].value = valor_sumado
                            except Exception: pass
                
                progress_bar.progress(1.0, text="Generando archivo final...")
                output = BytesIO()
                wb_final.save(output)
                output.seek(0)
                
                st.session_state.processed_file = output
                st.session_state.file_name = "Rem_consolidados.xlsx"
                progress_bar.empty()

            except Exception as e:
                st.error(f"Ocurrió un error crítico: {e}")
                st.error(traceback.format_exc())
                st.session_state.processed_file = None

# Lógica de descarga
if st.session_state.processed_file is not None:
    st.success("✅ ¡Consolidación completada!")
    st.download_button(
        label="📥 Descargar Rem_consolidados.xlsx",
        data=st.session_state.processed_file,
        file_name=st.session_state.file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    if st.button("Limpiar y empezar de nuevo"):
        st.session_state.processed_file = None
        st.session_state.file_name = None
        st.rerun()
