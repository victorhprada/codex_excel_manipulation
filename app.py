from __future__ import annotations

from io import BytesIO

import pandas as pd
import streamlit as st


SOURCE_SHEET_NAME = "Detalhado"
COMPILED_SHEET_NAME = "Compilado por funcionário"
COLUMN_ESTABELECIMENTO = "ESTABELECIMENTO"
CHECKOUT_COLUMN = "CHECKOUT"
COST_SHEET_NAME = "Custo empresa"
DISCOUNT_SHEET_NAME = "Desconto folha"
COST_FILTER_VALUE = "TARIFA RESGATE LIMITE PARA FLEX"
DISCOUNT_FILTER_VALUE = "RESGATE LIMITE PARA FLEX"
OUTPUT_FILENAME = "relatorio_processado.xlsx"


def process_excel(uploaded_file: BytesIO) -> BytesIO:
    excel_file = pd.ExcelFile(uploaded_file)

    if SOURCE_SHEET_NAME not in excel_file.sheet_names:
        available = ", ".join(excel_file.sheet_names)
        raise ValueError(
            f"A aba '{SOURCE_SHEET_NAME}' nao foi encontrada. Disponiveis: {available}"
        )

    if COMPILED_SHEET_NAME not in excel_file.sheet_names:
        available = ", ".join(excel_file.sheet_names)
        raise ValueError(
            f"A aba '{COMPILED_SHEET_NAME}' nao foi encontrada. Disponiveis: {available}"
        )

    # Lemos a aba base para aplicar os filtros das novas abas.
    detailed_frame = pd.read_excel(excel_file, sheet_name=SOURCE_SHEET_NAME)
    if COLUMN_ESTABELECIMENTO not in detailed_frame.columns:
        raise ValueError(
            f"A coluna '{COLUMN_ESTABELECIMENTO}' nao existe na aba '{SOURCE_SHEET_NAME}'."
        )

    compiled_frame = pd.read_excel(excel_file, sheet_name=COMPILED_SHEET_NAME)
    if CHECKOUT_COLUMN not in compiled_frame.columns:
        raise ValueError(
            f"A coluna '{CHECKOUT_COLUMN}' nao existe na aba '{COMPILED_SHEET_NAME}'."
        )

    key_columns = [
        column
        for column in detailed_frame.columns
        if column in compiled_frame.columns and column != CHECKOUT_COLUMN
    ]
    if not key_columns:
        raise ValueError(
            "Nao foi possivel identificar colunas em comum para cruzar o checkout."
        )

    checkout_filled = (
        compiled_frame[CHECKOUT_COLUMN].notna()
        & compiled_frame[CHECKOUT_COLUMN].astype(str).str.strip().ne("")
    )
    compiled_keys = pd.MultiIndex.from_frame(
        compiled_frame.loc[checkout_filled, key_columns].astype(str)
    )
    detailed_keys = pd.MultiIndex.from_frame(
        detailed_frame.loc[:, key_columns].astype(str)
    )
    checkout_mask = detailed_keys.isin(compiled_keys)

    # Aplicamos os filtros solicitados, mantendo a mesma estrutura de colunas.
    cost_frame = detailed_frame[
        (detailed_frame[COLUMN_ESTABELECIMENTO] == COST_FILTER_VALUE) | checkout_mask
    ]
    discount_frame = detailed_frame[
        (detailed_frame[COLUMN_ESTABELECIMENTO] == DISCOUNT_FILTER_VALUE)
        & ~checkout_mask
    ]

    # Gravamos todas as abas originais e adicionamos as novas abas filtradas.
    output_buffer = BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        for sheet_name in excel_file.sheet_names:
            frame = pd.read_excel(excel_file, sheet_name=sheet_name)
            frame.to_excel(writer, sheet_name=sheet_name, index=False)

        cost_frame.to_excel(writer, sheet_name=COST_SHEET_NAME, index=False)
        discount_frame.to_excel(writer, sheet_name=DISCOUNT_SHEET_NAME, index=False)

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

        st.success("Arquivo processado com sucesso.")
        st.download_button(
            label="Baixar arquivo processado",
            data=output_buffer,
            file_name=OUTPUT_FILENAME,
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )


if __name__ == "__main__":
    main()
