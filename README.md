# Excel Filter Helper

Aplicacao simples em Python para receber um arquivo Excel (.xlsx) e gerar outro arquivo com todas as abas originais, alem de duas abas filtradas a partir de **Detalhado**.

## Requisitos

- Python 3.10+
- pandas
- openpyxl

Instale dependencias:

```bash
pip install -r requirements.txt
```

## Uso

```bash
python app.py caminho/entrada.xlsx caminho/saida.xlsx
```

## Regras implementadas

- Mantem todas as abas originais exatamente como estao.
- Cria as abas adicionais:
  - **Custo empresa**: linhas em que `Estabelecimento` == `TARIFA RESGATE LIMITE PARA FLEX`
  - **Desconto folha**: linhas em que `Estabelecimento` == `RESGATE LIMITE PARA FLEX`
- Ambas mantem a mesma estrutura de colunas da aba **Detalhado**.

## Observacoes

- O arquivo de entrada deve conter a aba **Detalhado**.
- A coluna **Estabelecimento** precisa existir na aba **Detalhado**.
