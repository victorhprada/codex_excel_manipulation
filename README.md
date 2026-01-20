# Excel Filter Helper

Aplicacao simples em Python com Streamlit para receber um arquivo Excel (.xlsx) e gerar outro arquivo com todas as abas originais, alem de duas abas filtradas a partir de **Detalhado**.

## Requisitos

- Python 3.10+
- pandas
- openpyxl
- streamlit

Instale dependencias:

```bash
pip install -r requirements.txt
```

## Uso

```bash
streamlit run app.py
```

## Regras implementadas

- Mantem todas as abas originais exatamente como estao.
- Cria as abas adicionais:
  - **Custo empresa**: linhas em que `ESTABELECIMENTO` == `TARIFA RESGATE LIMITE PARA FLEX`
    e linhas em que `ESTABELECIMENTO` == `RESGATE LIMITE PARA FLEX` com **CHECKOUT** preenchido
  - **Desconto folha**: linhas em que `ESTABELECIMENTO` == `RESGATE LIMITE PARA FLEX`
    e **CHECKOUT** vazio ou nulo
- Ambas mantem a mesma estrutura de colunas da aba **Detalhado**.
- O arquivo final e disponibilizado para download com o nome **relatorio_processado.xlsx**.
- Registros com **CHECKOUT** preenchido na aba **Detalhado** sao sempre
  classificados como **Custo empresa** e removidos de **Desconto folha**.
- A aba **Custo empresa** e organizada com blocos visuais para separar registros sem
  checkout e registros com checkout (empresa e folha) ao final da tabela.

## Observacoes

- O arquivo de entrada deve conter a aba **Detalhado**.
- A coluna **ESTABELECIMENTO** precisa existir na aba **Detalhado**.
- A coluna **CHECKOUT** precisa existir na aba **Detalhado**.
