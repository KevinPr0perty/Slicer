import io
import math
import zipfile
from copy import copy

import streamlit as st
from openpyxl import load_workbook, Workbook


# ---------- Helpers ----------

def copy_cell(src_cell, dst_cell):
    """Copy value + formatting from one cell to another."""
    dst_cell.value = src_cell.value
    if src_cell.has_style:
        dst_cell._style = copy(src_cell._style)
    dst_cell.number_format = src_cell.number_format
    dst_cell.font = copy(src_cell.font)
    dst_cell.fill = copy(src_cell.fill)
    dst_cell.border = copy(src_cell.border)
    dst_cell.alignment = copy(src_cell.alignment)
    dst_cell.protection = copy(src_cell.protection)
    dst_cell.comment = src_cell.comment


def copy_sheet_layout(src_ws, dst_ws):
    """Copy sheet-level layout (best effort)."""
    dst_ws.freeze_panes = src_ws.freeze_panes
    dst_ws.sheet_format = copy(src_ws.sheet_format)
    dst_ws.sheet_properties = copy(src_ws.sheet_properties)
    dst_ws.page_setup = copy(src_ws.page_setup)
    dst_ws.page_margins = copy(src_ws.page_margins)
    dst_ws.print_options = copy(src_ws.print_options)

    for col_letter, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col_letter].width = dim.width
        dst_ws.column_dimensions[col_letter].hidden = dim.hidden
        dst_ws.column_dimensions[col_letter].outlineLevel = dim.outlineLevel
        dst_ws.column_dimensions[col_letter].collapsed = dim.collapsed

    for r, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[r].height = dim.height
        dst_ws.row_dimensions[r].hidden = dim.hidden
        dst_ws.row_dimensions[r].outlineLevel = dim.outlineLevel
        dst_ws.row_dimensions[r].collapsed = dim.collapsed

    for merged in list(src_ws.merged_cells.ranges):
        dst_ws.merge_cells(str(merged))


def split_excel_to_xlsx_bytes(
    file_bytes: bytes,
    sheet_name: str | None,
    chunk_size: int = 999,
    header_rows: int = 2,
    max_col_override: int | None = None,
):
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb[sheet_name] if sheet_name else wb.active

    max_col = max_col_override or ws.max_column
    max_row = ws.max_row

    data_start = header_rows + 1  # row 3 by default
    data_rows = max(0, max_row - header_rows)

    if data_rows <= 0:
        raise ValueError("No data rows found (expected data starting at row 3).")

    parts = math.ceil(data_rows / chunk_size)
    outputs: list[tuple[str, bytes]] = []

    for i in range(parts):
        start_idx = i * chunk_size
        end_idx = min((i + 1) * chunk_size, data_rows)

        out_wb = Workbook()
        out_ws = out_wb.active
        out_ws.title = ws.title

        copy_sheet_layout(ws, out_ws)

        # Copy header rows
        for r in range(1, header_rows + 1):
            for c in range(1, max_col + 1):
                copy_cell(ws.cell(row=r, column=c), out_ws.cell(row=r, column=c))

        # Copy data rows
        out_row = data_start
        for k in range(start_idx, end_idx):
            src_r = data_start + k
            for c in range(1, max_col + 1):
                copy_cell(ws.cell(row=src_r, column=c), out_ws.cell(row=out_row, column=c))
            out_row += 1

        bio = io.BytesIO()
        out_wb.save(bio)
        bio.seek(0)

        outputs.append((f"part_{i+1}_of_{parts}.xlsx", bio.read()))

    return outputs


def make_zip(files: list[tuple[str, bytes]]) -> bytes:
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, fbytes in files:
            zf.writestr(fname, fbytes)
    zip_buffer.seek(0)
    return zip_buffer.read()


# ---------- Streamlit UI ----------

st.set_page_config(page_title="Temu Excel Splitter", layout="centered")

st.title("Temu Excel Inventory Splitter")
st.write(
    "Upload a Temu inventory Excel file. "
    "Each output file keeps the **first two rows** and contains **up to 999 data rows**."
)

uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])

chunk_size = st.number_input(
    "Data rows per file",
    min_value=100,
    max_value=999,
    value=999,
    step=1,
)

header_rows = st.number_input(
    "Header rows to repeat",
    min_value=1,
    max_value=10,
    value=2,
    step=1,
)

sheet_name = st.text_input("Sheet name (optional)", value="")

max_col_input = st.number_input(
    "Max columns to copy (0 = auto)",
    min_value=0,
    max_value=500,
    value=0,
    step=1,
)

max_col_override = None if max_col_input == 0 else int(max_col_input)

if uploaded:
    st.success(f"Loaded: {uploaded.name}")

    if st.button("Split & Generate ZIP"):
        try:
            parts = split_excel_to_xlsx_bytes(
                file_bytes=uploaded.read(),
                sheet_name=sheet_name.strip() or None,
                chunk_size=int(chunk_size),
                header_rows=int(header_rows),
                max_col_override=max_col_override,
            )

            zip_bytes = make_zip(parts)

            st.success(f"Done! Generated {len(parts)} files.")
            st.download_button(
                label="Download ZIP",
                data=zip_bytes,
                file_name="temu_split_files.zip",
                mime="application/zip",
            )

        except Exception as e:
            st.error(f"Error: {e}")
