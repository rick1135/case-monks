# Case Monks — Limpeza e Análise de Oportunidades

## 1) Visão Geral da Solução

Este case organiza uma base comercial com problemas de qualidade para transformar dados brutos em informação confiável para decisão de negócio.

Em termos práticos, a solução:

- limpa inconsistências estruturais da base de oportunidades;
- elimina duplicidades que inflavam o pipeline;
- padroniza campos críticos (datas, valores e moedas);
- gera entregáveis executivos para tomada de decisão.

Resumo do resultado final:

- Base original: **413 linhas**
- Base limpa: **261 linhas**
- Duplicatas removidas: **152 linhas**

## 2) Fluxo de Trabalho (Ponta a Ponta)

1. **Limpeza de dados (`src/cleaning.py`)**
   Aplica regras de qualidade, deduplicação e padronização.
2. **Relatório de qualidade (`src/error_report.py`)**
   Gera visão executiva dos problemas encontrados e impacto.
3. **Análise comercial (`src/analysis_report.py`)**
   Produz visualizações e insights de pipeline.
4. **Apresentação final (`src/presentation.py`)**
   Consolida narrativa executiva em PDF (5 páginas).

## 3) Estrutura do Projeto

```text
case-monks/
|-- opps_corrupted.xlsx        # Base original (entrada)
|-- opps_corrigido.xlsx        # Base limpa (saída)
|-- relatorio_erros.html       # Relatório de qualidade de dados
|-- analise.html               # Análise comercial com gráficos
|-- apresentacao.pdf           # Apresentação executiva final
|-- README.md
`-- src/
    |-- cleaning.py            # Limpeza e padronização dos dados
    |-- error_report.py        # Relatório de erros e riscos
    |-- analysis_report.py     # Análise comercial
    `-- presentation.py        # Geração da apresentação em PDF
```

## 4) Dependências e Setup

Pré-requisitos:

- Python 3.10+
- `pip`

Instalação das bibliotecas:

```bash
python -m pip install pandas numpy openpyxl plotly reportlab
```

## 5) Como Executar (Comandos Exatos)

No diretório raiz do projeto:

```bash
python src/cleaning.py
python src/error_report.py
python src/analysis_report.py
python src/presentation.py
```

Arquivos gerados:

- `opps_corrigido.xlsx`
- `relatorio_erros.html`
- `analise.html`
- `apresentacao.pdf`

## 6) O que Cada Artefato Responde

- **`opps_corrigido.xlsx`**
  “Qual é a base confiável para gestão e análise comercial?”

- **`relatorio_erros.html`**
  “Quais problemas de qualidade existiam e qual o impacto no negócio?”

- **`analise.html`**
  “Como o pipeline está distribuído e onde estão os principais focos de ação?”

- **`apresentacao.pdf`**
  “Qual a narrativa final para liderança: contexto, achados, recomendações e aprendizados?”

## 7) Premissas de Negócio (Crítico)

### 7.1 Remoção de 152 duplicatas

A remoção de **152 duplicatas** foi necessária para evitar distorções de gestão, principalmente:

- inflação artificial de pipeline;
- previsão de receita superestimada;
- priorização comercial incorreta.

Critério adotado para manter apenas 1 registro por oportunidade (`Opportunity_ID`):

1. **Maior completude de campos** (registro com mais informação útil).
2. Em caso de empate, **`Created_Date` mais antiga** (registro pioneiro, mais estável para histórico).

Esse critério melhora consistência analítica e reduz ruído operacional sem descaracterizar o histórico comercial.

### 7.2 Tratamento de nulos

Regras aplicadas:

- remoção de `Opportunity_ID` nulo/vazio (sem chave mínima de rastreabilidade);
- preenchimento de nulos categóricos com **`Unknown`**;
- preenchimento de nulos numéricos com **`0`**, após parse robusto.

Racional de negócio: preservar continuidade das análises, explicitar ausência de informação e evitar que dashboards/indicadores quebrem por dados faltantes.

### 7.3 Regra cambial sem conversão FX

Regra adotada:

- harmonizar `Amount_Currency` e `Total_Product_Amount_Currency` na mesma linha;
- padronizar código monetário;
- **não converter valores entre moedas**.

Justificativa: sem taxa e data de câmbio confiáveis no dataset, qualquer conversão FX introduziria risco de distorção financeira. A prioridade foi consistência de classificação, não recalcular valores.

## 8) Limitações e Próximos Passos

Limitações atuais:

- sem política de FX histórico para consolidação em moeda única;
- dependência de qualidade de entrada no CRM;
- insights sujeitos à atualização da base e sazonalidade comercial.

Próximos passos recomendados:

1. Criar validações obrigatórias por estágio no CRM (campos críticos).
2. Bloquear duplicidade na origem com chave canônica (`Opportunity_ID` + `Account_ID`).
3. Instituir monitoramento semanal de qualidade (duplicatas, nulos críticos, datas inválidas, outliers).
4. Definir governança formal de dados com responsável de negócio por domínio comercial.
