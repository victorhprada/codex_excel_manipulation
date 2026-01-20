# Excel Filter Helper

Ferramenta interna para **tratamento e organização de relatórios de consumo**, criada para
eliminar ajustes manuais em Excel, evitar erros de cálculo e facilitar a leitura dos dados
pela área administrativa e financeira.

---

## 🎯 Qual problema esta ferramenta resolve?

Os relatórios de consumo exportados pelo sistema:

- Exigem **edições manuais recorrentes**
- Misturam custos da empresa com valores descontados em folha
- Apresentam **checkouts (demissões)** misturados com funcionários ativos
- Possuem totais **fixos**, que não se ajustam quando linhas são removidas
- Têm layout complexo, o que dificulta ajustes sem quebrar o formato

👉 Isso gera **retrabalho**, **risco de erro** e **perda de tempo** para o time.

---

## ✅ O que a ferramenta faz

A partir de um **único upload de Excel**, a aplicação gera um novo arquivo:

### 📊 Organização dos dados
- Mantém todas as abas originais do relatório
- Cria as abas:
  - **Custo empresa**
  - **Desconto folha**
- Aplica regras claras para separar:
  - Custos assumidos pela empresa
  - Valores descontados em folha
  - Checkouts (demissões)

### 👥 Tratamento de checkouts
- Registros com **CHECKOUT preenchido**:
  - Não aparecem em *Desconto folha*
  - São tratados como **Custo empresa**
- Dentro da aba **Custo empresa**, os checkouts são separados visualmente em:
  - **Checkouts Empresa**
  - **Checkouts Folha colab**

### 🧾 Overview confiável
- Remove automaticamente as linhas:
  - Subsídios
  - Taxa administrativa
- Mantém **100% do layout original**
- Aplica **fórmulas de soma no TOTAL DA EMPRESA**, garantindo:
  - Recalculo automático
  - Zero inconsistência ao remover ou adicionar linhas

### 📁 Arquivo final
- Mantém o nome original do arquivo
- Adiciona o prefixo:
  - `processado_`
- Exemplo:
