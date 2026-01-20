import argparse
from pathlib import Path

import pandas as pd


COST_SHEET_NAME = "Custo empresa"
DISCOUNT_SHEET_NAME = "Desconto folha"
SOURCE_SHEET_NAME = "Detalhado"
COLUMN_ESTABELECIMENTO = "Estabelecimento"
COST_FILTER_VALUE = "TARIFA RESGATE LIMITE PARA FLEX"
DISCOUNT_FILTER_VALUE = "RESGATE LIMITE PARA FLEX"


def build_output_excel(input_path: Path, output_path: Path) -> None:
    excel_file = pd.ExcelFile(input_path)

    if SOURCE_SHEET_NAME not in excel_file.sheet_names:
        available = ", ".join(excel_file.sheet_names)
        raise ValueError(
            f"A aba '{SOURCE_SHEET_NAME}' nao foi encontrada. Disponiveis: {available}"
        )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name in excel_file.sheet_names:
            frame = pd.read_excel(excel_file, sheet_name=sheet_name)
            frame.to_excel(writer, sheet_name=sheet_name, index=False)

        detailed_frame = pd.read_excel(excel_file, sheet_name=SOURCE_SHEET_NAME)
        if COLUMN_ESTABELECIMENTO not in detailed_frame.columns:
            raise ValueError(
                f"A coluna '{COLUMN_ESTABELECIMENTO}' nao existe na aba '{SOURCE_SHEET_NAME}'."
            )

        cost_frame = detailed_frame[
            detailed_frame[COLUMN_ESTABELECIMENTO] == COST_FILTER_VALUE
        ]
        discount_frame = detailed_frame[
            detailed_frame[COLUMN_ESTABELECIMENTO] == DISCOUNT_FILTER_VALUE
        ]

        cost_frame.to_excel(writer, sheet_name=COST_SHEET_NAME, index=False)
        discount_frame.to_excel(writer, sheet_name=DISCOUNT_SHEET_NAME, index=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Gera um novo arquivo Excel com as abas originais e duas abas filtradas."
        )
    )
    parser.add_argument("input", type=Path, help="Caminho do arquivo Excel de entrada")
    parser.add_argument("output", type=Path, help="Caminho do arquivo Excel de saida")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    build_output_excel(args.input, args.output)


if __name__ == "__main__":
    main()
