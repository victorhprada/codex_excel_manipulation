from __future__ import annotations

from io import BytesIO
from copy import copy
import unicodedata

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


CENTER_SHEET_NAME = "Detalhado"
COLUMN_ESTABELECIMENTO = "ESTABELECIMENTO"
CHECKOUT_COLUMN = "CHECKOUT"

COST_SHEET_NAME = "Custo empresa"
DISCOUNT_SHEET_NAME = "Desconto folha"

COST_FILTER_VALUE = "TARIFA RESGATE LIMITE PARA FLEX"
DISCOUNT_FILTER_VALUE = "RESGATE LIMITE PARA FLEX"

OVERVIEW_SHEET_NAME = "Overview"
OVERVIEW_TOTAL_LABEL = "TOTAL DA EMPRESA"
OVERVIEW_CHECKOUT_FOLHA_LABEL = "Checkouts Folha colab."
OVERVIEW_CHECKOUT_EMPRESA_LABEL = "Checkouts a pagar Empresa"
OVERVIEW_CUSTO_EMPRESA_LABEL = "Custo empresa (Taxa tarifas)"
OVERVIEW_A_DEBITAR_LABEL = "A debitar em folha"
OVERVIEW_TOTAL_FUNC_LABEL = "TOTAL DO FUNCIONÁRIO"
OVERVIEW_TOTAL_FECHAMENTO_LABEL = "TOTAL DO FECHAMENTO"

COST_HEADER_ESTABELECIMENTO = "ESTABELECIMENTO"
COST_HEADER_CHECKOUT = "CHECKOUT"
COST_HEADER_DEBITO = "DEBITO EM FOLHA"
COST_HEADER_DEBITO_ACCENT = "DÉBITO EM FOLHA"


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return "".join(
        char for char in unicodedata.normalize("NFD", text)
        if unicodedata.category(char) != "Mn"
    )


def find_label_cell(sheet, label: str):
    target = normalize_text(label)
    for row in sheet.iter_rows():
        for cell in row:
            if normalize_text(cell.value) == target:
                return cell
    return None


def find_value_cell(sheet, label_cell):
    for cell in sheet[label_cell.row]:
        if cell.column > label_cell.column and cell.value not in (None, ""):
            return cell
    return None


def find_header_column(sheet, labels: set[str]) -> str | None:
    normalized = {normalize_text(label) for label in labels}
    for cell in sheet[1]:
        if normalize_text(cell.value) in normalized:
            return cell.column_letter
    return None


def copy_row_style(source_row, target_row) -> None:
    for source_cell, target_cell in zip(source_row, target_row):
        if source_cell.font:
            target_cell.font = copy(source_cell.font)
        if source_cell.fill:
            target_cell.fill = copy(source_cell.fill)
        if source_cell.border:
            target_cell.border = copy(source_cell.border)
        if source_cell.alignment:
            target_cell.alignment = copy(source_cell.alignment)
        target_cell.number_format = source_cell.number_format


def process_excel(uploaded_file: BytesIO) -> BytesIO:
    bytes_data = uploaded_file.getvalue()
    excel_file = pd.ExcelFile(BytesIO(bytes_data))

    detailed_frame = pd.read_excel(excel_file, sheet_name=CENTER_SHEET_NAME)

    checkout_filled = (
        detailed_frame[CHECKOUT_COLUMN].notna()
        & detailed_frame[CHECKOUT_COLUMN].astype(str).str.strip().ne("")
    )

    cost_no_checkout = detailed_frame[
        detailed_frame[COLUMN_ESTABELECIMENTO].isin(
            [COST_FILTER_VALUE, DISCOUNT_FILTER_VALUE]
        )
        & ~checkout_filled
    ]

    cost_checkout_empresa = detailed_frame[
        (detailed_frame[COLUMN_ESTABELECIMENTO] == COST_FILTER_VALUE) & checkout_filled
    ]

    cost_checkout_folha = detailed_frame[
        (detailed_frame[COLUMN_ESTABELECIMENTO] == DISCOUNT_FILTER_VALUE)
        & checkout_filled
    ]

    discount_frame = detailed_frame[
        (detailed_frame[COLUMN_ESTABELECIMENTO] == DISCOUNT_FILTER_VALUE)
        & ~checkout_filled
    ]

    title_empresa = pd.DataFrame([{detailed_frame.columns[0]: "Checkouts Empresa"}])
    title_folha = pd.DataFrame([{detailed_frame.columns[0]: "Checkouts Folha colab"}])

    title_empresa = title_empresa.reindex(columns=detailed_frame.columns, fill_value="")
    title_folha = title_folha.reindex(columns=detailed_frame.columns, fill_value="")

    cost_frame = pd.concat(
        [
            cost_no_checkout,
            title_empresa,
            cost_checkout_empresa,
            title_folha,
            cost_checkout_folha,
        ],
        ignore_index=True,
    )

    workbook = load_workbook(BytesIO(bytes_data))
    overview_sheet = workbook[OVERVIEW_SHEET_NAME]

    checkout_folha_cell = find_label_cell(overview_sheet, OVERVIEW_CHECKOUT_FOLHA_LABEL)
    checkout_empresa_cell = find_label_cell(overview_sheet, OVERVIEW_CHECKOUT_EMPRESA_LABEL)
    custo_empresa_cell = find_label_cell(overview_sheet, OVERVIEW_CUSTO_EMPRESA_LABEL)
    total_empresa_cell = find_label_cell(overview_sheet, OVERVIEW_TOTAL_LABEL)
    a_debitar_cell = find_label_cell(overview_sheet, OVERVIEW_A_DEBITAR_LABEL)
    total_func_cell = find_label_cell(overview_sheet, OVERVIEW_TOTAL_FUNC_LABEL)
    total_fechamento_cell = find_label_cell(overview_sheet, OVERVIEW_TOTAL_FECHAMENTO_LABEL)

    # 👉 REMOÇÃO DE LINHAS TOTALMENTE VAZIAS (acabamento visual)
    rows_to_delete = []
    for row in overview_sheet.iter_rows():
        if all(cell.value in (None, "") for cell in row):
            rows_to_delete.append(row[0].row)

    for row_idx in sorted(rows_to_delete, reverse=True):
        overview_sheet.delete_rows(row_idx)

    for sheet_name in (COST_SHEET_NAME, DISCOUNT_SHEET_NAME):
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]

    cost_sheet = workbook.create_sheet(COST_SHEET_NAME)
    for row in dataframe_to_rows(cost_frame, index=False, header=True):
        cost_sheet.append(row)

    discount_sheet = workbook.create_sheet(DISCOUNT_SHEET_NAME)
    for row in dataframe_to_rows(discount_frame, index=False, header=True):
        discount_sheet.append(row)

    cost_debito_col = find_header_column(
        cost_sheet, {COST_HEADER_DEBITO, COST_HEADER_DEBITO_ACCENT}
    )
    cost_estabelecimento_col = find_header_column(cost_sheet, {COST_HEADER_ESTABELECIMENTO})
    cost_checkout_col = find_header_column(cost_sheet, {COST_HEADER_CHECKOUT})

    if cost_debito_col and cost_estabelecimento_col and cost_checkout_col:
        find_value_cell(overview_sheet, checkout_folha_cell).value = (
            f"=SUMIFS('Custo empresa'!{cost_debito_col}:{cost_debito_col},"
            f"'Custo empresa'!{cost_estabelecimento_col}:{cost_estabelecimento_col},"
            f"\"{DISCOUNT_FILTER_VALUE}\","
            f"'Custo empresa'!{cost_checkout_col}:{cost_checkout_col},\"<>\")"
        )

        find_value_cell(overview_sheet, checkout_empresa_cell).value = (
            f"=SUMIFS('Custo empresa'!{cost_debito_col}:{cost_debito_col},"
            f"'Custo empresa'!{cost_estabelecimento_col}:{cost_estabelecimento_col},"
            f"\"{COST_FILTER_VALUE}\","
            f"'Custo empresa'!{cost_checkout_col}:{cost_checkout_col},\"<>\")"
        )

        find_value_cell(overview_sheet, custo_empresa_cell).value = (
            f"=SUMIFS('Custo empresa'!{cost_debito_col}:{cost_debito_col},"
            f"'Custo empresa'!{cost_checkout_col}:{cost_checkout_col},\"=\")"
        )

    total_empresa_value = find_value_cell(overview_sheet, total_empresa_cell)
    total_empresa_value.value = (
        f"=SUM({find_value_cell(overview_sheet, checkout_folha_cell).coordinate};"
        f"{find_value_cell(overview_sheet, checkout_empresa_cell).coordinate};"
        f"{find_value_cell(overview_sheet, custo_empresa_cell).coordinate})"
    )

    a_debitar_value = find_value_cell(overview_sheet, a_debitar_cell)
    a_debitar_value.value = "=SUM('Desconto folha'!M:M)"

    total_func_value = find_value_cell(overview_sheet, total_func_cell)
    total_func_value.value = f"={a_debitar_value.coordinate}"

    total_fechamento_value = overview_sheet.cell(
        row=total_fechamento_cell.row + 1,
        column=total_fechamento_cell.column,
    )
    total_fechamento_value.value = (
        f"={total_empresa_value.coordinate}+{total_func_value.coordinate}"
    )

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def main() -> None:
    st.title("Gerador de Relatório Excel")

    uploaded_file = st.file_uploader(
        "Selecione o arquivo Excel (.xlsx)", type=["xlsx"]
    )

    if uploaded_file and st.button("Processar arquivo"):
        output = process_excel(uploaded_file)
        st.download_button(
            "Baixar arquivo processado",
            output,
            file_name=f"processado_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
