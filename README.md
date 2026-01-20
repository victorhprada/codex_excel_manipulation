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
  - **Custo empresa**: linhas em que `Estabelecimento` == `TARIFA RESGATE LIMITE PARA FLEX`
  - **Desconto folha**: linhas em que `Estabelecimento` == `RESGATE LIMITE PARA FLEX`
- Ambas mantem a mesma estrutura de colunas da aba **Detalhado**.
- O arquivo final e disponibilizado para download com o nome **relatorio_processado.xlsx**.

## Observacoes

- O arquivo de entrada deve conter a aba **Detalhado**.
- A coluna **Estabelecimento** precisa existir na aba **Detalhado**.
