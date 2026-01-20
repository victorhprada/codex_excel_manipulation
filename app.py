from __future__ import annotations

from io import BytesIO

import pandas as pd
import streamlit as st


CENTER_SHEET_NAME = "Detalhado"
COLUMN_ESTABELECIMENTO = "ESTABELECIMENTO"
CHECKOUT_COLUMN = "CHECKOUT"
COST_SHEET_NAME = "Custo empresa"
DISCOUNT_SHEET_NAME = "Desconto folha"
COST_FILTER_VALUE = "TARIFA RESGATE LIMITE PARA FLEX"
DISCOUNT_FILTER_VALUE = "RESGATE LIMITE PARA FLEX"
OUTPUT_FILENAME = "relatorio_processado.xlsx"


def process_excel(uploaded_file: BytesIO) -> BytesIO:
    excel_file = pd.ExcelFile(uploaded_file)

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
        (detailed_frame[COLUMN_ESTABELECIMENTO] == COST_FILTER_VALUE)
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
