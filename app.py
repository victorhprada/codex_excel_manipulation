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
OVERVIEW_FILTER_VALUES = {"Taxa administrativa", "Subsídios", "Créditos inseridos"}
OVERVIEW_SECTION_LABEL = "PARTE DA EMPRESA"
OVERVIEW_TOTAL_LABEL = "TOTAL DA EMPRESA"
OVERVIEW_CHECKOUT_FOLHA_LABEL = "Checkouts Folha colab."
OVERVIEW_CHECKOUT_EMPRESA_LABEL = "Checkouts a pagar Empresa"
OVERVIEW_CUSTO_EMPRESA_LABEL = "Custo empresa (Taxa tarifas)"
OVERVIEW_A_DEBITAR_LABEL = "A debitar em folha"
OVERVIEW_TOTAL_FUNC_LABEL = "TOTAL DO FUNCIONÁRIO"
OVERVIEW_TOTAL_FECHAMENTO_LABEL = "TOTAL DO FECHAMENTO"


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


def find_value_cell(sheet, label_cell):
    for cell in sheet[label_cell.row]:
        if cell.column > label_cell.column and cell.value not in (None, ""):
            return cell
    return None


def process_excel(uploaded_file: BytesIO) -> BytesIO:
    bytes_data = uploaded_file.getvalue()
    excel_file = pd.ExcelFile(BytesIO(bytes_data))

    if CENTER_SHEET_NAME not in excel_file.sheet_names:
        available = ", ".join(excel_file.sheet_names)
        raise ValueError(
            f"A aba '{CENTER_SHEET_NAME}' nao foi encontrada. Disponiveis: {available}"
        )

    # Lemos a aba base para aplicar os filtros das novas abas.
    detailed_frame = pd.read_excel(excel_file, sheet_name=CENTER_SHEET_NAME)
    if COLUMN_ESTABELECIMENTO not in detailed_frame.columns:
        raise ValueError(
            f"A coluna '{COLUMN_ESTABELECIMENTO}' nao existe na aba '{CENTER_SHEET_NAME}'."
        )

    if CHECKOUT_COLUMN not in detailed_frame.columns:
        raise ValueError(
            f"A coluna '{CHECKOUT_COLUMN}' nao existe na aba '{CENTER_SHEET_NAME}'."
        )

    checkout_filled = (
        detailed_frame[CHECKOUT_COLUMN].notna()
        & detailed_frame[CHECKOUT_COLUMN].astype(str).str.strip().ne("")
    )

    # Aplicamos os filtros solicitados, mantendo a mesma estrutura de colunas.
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
    title_folha = pd.DataFrame(
        [{detailed_frame.columns[0]: "Checkouts Folha colab"}]
    )
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
    overview_sheet = (
        workbook[OVERVIEW_SHEET_NAME]
        if OVERVIEW_SHEET_NAME in workbook.sheetnames
        else None
    )
    checkout_folha_cell = None
    checkout_empresa_cell = None
    custo_empresa_cell = None
    total_empresa_cell = None
    a_debitar_cell = None
    total_func_cell = None
    total_fechamento_cell = None
    if overview_sheet:
        normalized_filters = {normalize_text(value) for value in OVERVIEW_FILTER_VALUES}
        rows_to_remove = {
            cell.row
            for row in overview_sheet.iter_rows()
            for cell in row
            if normalize_text(cell.value) in normalized_filters
        }
        for row_idx in sorted(rows_to_remove, reverse=True):
            overview_sheet.delete_rows(row_idx)

        checkout_folha_cell = find_label_cell(
            overview_sheet, OVERVIEW_CHECKOUT_FOLHA_LABEL
        )
        checkout_empresa_cell = find_label_cell(
            overview_sheet, OVERVIEW_CHECKOUT_EMPRESA_LABEL
        )
        custo_empresa_cell = find_label_cell(overview_sheet, OVERVIEW_CUSTO_EMPRESA_LABEL)
        total_empresa_cell = find_label_cell(overview_sheet, OVERVIEW_TOTAL_LABEL)
        a_debitar_cell = find_label_cell(overview_sheet, OVERVIEW_A_DEBITAR_LABEL)
        total_func_cell = find_label_cell(overview_sheet, OVERVIEW_TOTAL_FUNC_LABEL)
        total_fechamento_cell = find_label_cell(
            overview_sheet, OVERVIEW_TOTAL_FECHAMENTO_LABEL
        )

    for sheet_name in (COST_SHEET_NAME, DISCOUNT_SHEET_NAME):
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]

    cost_sheet = workbook.create_sheet(COST_SHEET_NAME)
    for row in dataframe_to_rows(cost_frame, index=False, header=True):
        cost_sheet.append(row)

    discount_sheet = workbook.create_sheet(DISCOUNT_SHEET_NAME)
    for row in dataframe_to_rows(discount_frame, index=False, header=True):
        discount_sheet.append(row)

    if overview_sheet:
        total_empresa_value = None
        if total_empresa_cell:
            total_empresa_value = find_value_cell(overview_sheet, total_empresa_cell)
            if (
                total_empresa_value
                and checkout_folha_cell
                and checkout_empresa_cell
                and custo_empresa_cell
            ):
                checkouts_folha_value = find_value_cell(
                    overview_sheet, checkout_folha_cell
                )
                checkouts_empresa_value = find_value_cell(
                    overview_sheet, checkout_empresa_cell
                )
                custo_empresa_value = find_value_cell(
                    overview_sheet, custo_empresa_cell
                )
                company_cells = [
                    value_cell.coordinate
                    for value_cell in (
                        checkouts_folha_value,
                        checkouts_empresa_value,
                        custo_empresa_value,
                    )
                    if value_cell
                ]
                if company_cells:
                    total_empresa_value.value = f"=SUM({','.join(company_cells)})"
        total_func_value = None
        if a_debitar_cell:
            value_cell = find_value_cell(overview_sheet, a_debitar_cell)
            if value_cell:
                value_cell.value = "=SUM('Desconto folha'!M:M)"
        if total_func_cell:
            total_func_value = find_value_cell(overview_sheet, total_func_cell)
            if total_func_value and a_debitar_cell:
                a_debitar_value = find_value_cell(overview_sheet, a_debitar_cell)
                if a_debitar_value:
                    total_func_value.value = f"={a_debitar_value.coordinate}"
        if total_fechamento_cell and total_empresa_value and total_func_value:
            total_fechamento_value = find_value_cell(
                overview_sheet, total_fechamento_cell
            )
            if total_fechamento_value:
                total_fechamento_value.value = (
                    f"={total_empresa_value.coordinate}+{total_func_value.coordinate}"
                )

    output_buffer = BytesIO()
    workbook.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer


def main() -> None:
    st.title("Gerador de Relatorio Excel")
    st.write(
        "Envie um arquivo Excel (.xlsx) com as abas originais e gere um novo arquivo "
        "com as abas adicionais 'Custo empresa' e 'Desconto folha'."
    )

    uploaded_file = st.file_uploader(
        "Selecione o arquivo Excel (.xlsx)",
        type=["xlsx"],
    )

    if uploaded_file is None:
        st.info("Nenhum arquivo carregado. Envie um arquivo Excel para continuar.")
        return

    if st.button("Processar arquivo"):
        try:
            output_buffer = process_excel(uploaded_file)
        except ValueError as exc:
            st.error(str(exc))
            return
        except Exception as exc:  # noqa: BLE001
            st.error(f"Erro inesperado ao processar o arquivo: {exc}")
            return

        output_filename = f"processado_{uploaded_file.name}"
        st.success("Arquivo processado com sucesso.")
        st.download_button(
            label="Baixar arquivo processado",
            data=output_buffer,
            file_name=output_filename,
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )


if __name__ == "__main__":
    main()
