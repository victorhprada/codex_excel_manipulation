from __future__ import annotations

from io import BytesIO
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

# Labels existentes / alvo
OVERVIEW_SECTION_LABEL = "PARTE DA EMPRESA"
OVERVIEW_VALUE_HEADER_LABEL = "VALOR"
OVERVIEW_TOTAL_LABEL = "TOTAL DA EMPRESA"
OVERVIEW_TOTAL_FECHAMENTO_LABEL = "TOTAL DO FECHAMENTO"

# Linhas que já existem no Overview (serão reutilizadas/renomeadas)
OVERVIEW_CHECKOUT_PAGAR_LABEL = "Checkouts a pagar"
OVERVIEW_TAXA_ADMIN_LABEL = "Taxa administrativa"
OVERVIEW_SUBSIDIOS_LABEL = "Subsídios"
OVERVIEW_CREDITOS_LABEL = "Créditos inseridos"  # vamos “limpar” sem deletar linha

# Novos nomes (regras)
OVERVIEW_CHECKOUT_FOLHA_LABEL = "Checkouts Folha colab."
OVERVIEW_CHECKOUT_EMPRESA_LABEL = "Checkouts a pagar Empresa"
OVERVIEW_CUSTO_EMPRESA_LABEL = "Custo empresa (Taxa tarifas)"

OVERVIEW_A_DEBITAR_LABEL = "A debitar em folha"
OVERVIEW_TOTAL_FUNC_LABEL = "TOTAL DO FUNCIONÁRIO"

# Headers na aba Custo empresa
COST_HEADER_ESTABELECIMENTO = "ESTABELECIMENTO"
COST_HEADER_CHECKOUT = "CHECKOUT"
COST_HEADER_DEBITO = "DEBITO EM FOLHA"
COST_HEADER_DEBITO_ACCENT = "DÉBITO EM FOLHA"


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return "".join(
        char for char in unicodedata.normalize("NFD", text) if unicodedata.category(char) != "Mn"
    )


def find_label_cell(sheet, label: str):
    target = normalize_text(label)
    for row in sheet.iter_rows():
        for cell in row:
            if normalize_text(cell.value) == target:
                return cell
    return None


def find_header_column_letter(sheet, labels: set[str]) -> str | None:
    normalized = {normalize_text(label) for label in labels}
    for cell in sheet[1]:
        if normalize_text(cell.value) in normalized:
            return cell.column_letter
    return None


def copy_row_style(source_row, target_row) -> None:
    for source_cell, target_cell in zip(source_row, target_row):
        target_cell.font = source_cell.font
        target_cell.fill = source_cell.fill
        target_cell.border = source_cell.border
        target_cell.alignment = source_cell.alignment
        target_cell.number_format = source_cell.number_format


def get_overview_value_col(overview_sheet):
    """
    Descobre a coluna da tabela de valores do Overview.
    Estratégia: encontrar "PARTE DA EMPRESA" e na MESMA linha achar "VALOR".
    """
    section_cell = find_label_cell(overview_sheet, OVERVIEW_SECTION_LABEL)
    if not section_cell:
        return None

    row = overview_sheet[section_cell.row]
    for cell in row:
        if normalize_text(cell.value) == normalize_text(OVERVIEW_VALUE_HEADER_LABEL):
            return cell.column  # índice numérico
    return None


def get_overview_value_cell(overview_sheet, label_cell, value_col: int):
    """
    Retorna a célula de valor da mesma linha do label, na coluna `value_col`,
    mesmo que esteja vazia.
    """
    if not label_cell or not value_col:
        return None
    return overview_sheet.cell(row=label_cell.row, column=value_col)


def process_excel(uploaded_file: BytesIO) -> BytesIO:
    bytes_data = uploaded_file.getvalue()
    excel_file = pd.ExcelFile(BytesIO(bytes_data))

    if CENTER_SHEET_NAME not in excel_file.sheet_names:
        available = ", ".join(excel_file.sheet_names)
        raise ValueError(f"A aba '{CENTER_SHEET_NAME}' nao foi encontrada. Disponiveis: {available}")

    detailed_frame = pd.read_excel(excel_file, sheet_name=CENTER_SHEET_NAME)

    if COLUMN_ESTABELECIMENTO not in detailed_frame.columns:
        raise ValueError(f"A coluna '{COLUMN_ESTABELECIMENTO}' nao existe na aba '{CENTER_SHEET_NAME}'.")
    if CHECKOUT_COLUMN not in detailed_frame.columns:
        raise ValueError(f"A coluna '{CHECKOUT_COLUMN}' nao existe na aba '{CENTER_SHEET_NAME}'.")

    checkout_filled = (
        detailed_frame[CHECKOUT_COLUMN].notna()
        & detailed_frame[CHECKOUT_COLUMN].astype(str).str.strip().ne("")
    )

    # Base Custo empresa:
    cost_no_checkout = detailed_frame[
        detailed_frame[COLUMN_ESTABELECIMENTO].isin([COST_FILTER_VALUE, DISCOUNT_FILTER_VALUE])
        & ~checkout_filled
    ]
    cost_checkout_empresa = detailed_frame[
        (detailed_frame[COLUMN_ESTABELECIMENTO] == COST_FILTER_VALUE) & checkout_filled
    ]
    cost_checkout_folha = detailed_frame[
        (detailed_frame[COLUMN_ESTABELECIMENTO] == DISCOUNT_FILTER_VALUE) & checkout_filled
    ]

    # Base Desconto folha:
    discount_frame = detailed_frame[
        (detailed_frame[COLUMN_ESTABELECIMENTO] == DISCOUNT_FILTER_VALUE) & ~checkout_filled
    ]

    # “Títulos” dentro da aba Custo empresa (mantemos como antes)
    title_empresa = pd.DataFrame([{detailed_frame.columns[0]: "Checkouts Empresa"}]).reindex(
        columns=detailed_frame.columns, fill_value=""
    )
    title_folha = pd.DataFrame([{detailed_frame.columns[0]: "Checkouts Folha colab"}]).reindex(
        columns=detailed_frame.columns, fill_value=""
    )

    cost_frame = pd.concat(
        [cost_no_checkout, title_empresa, cost_checkout_empresa, title_folha, cost_checkout_folha],
        ignore_index=True,
    )

    workbook = load_workbook(BytesIO(bytes_data))
    overview_sheet = workbook[OVERVIEW_SHEET_NAME] if OVERVIEW_SHEET_NAME in workbook.sheetnames else None

    # Remove e recria as abas tabula
