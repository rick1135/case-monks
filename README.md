# Case Monks — Limpeza e Análise de Oportunidades

## 1) Visão Geral da Solução

Este case foi desenvolvido como um pipeline RevOps **em fases rigorosas (checkpoints)** para garantir qualidade, rastreabilidade e confiabilidade da análise:

1. **Extração e inspeção inicial da base**
2. **Limpeza e padronização dos dados**
3. **Validação da qualidade (relatório de erros)**
4. **Análise de pipeline e geração de insights**
5. **Síntese executiva final**

Essa abordagem faseada permite rastrear claramente **o que entrou, o que foi transformado, por que foi transformado e qual impacto cada ajuste teve no resultado final**.

Resumo do resultado:

- Base original: **413 linhas**
- Base final limpa: **261 linhas**
- Duplicatas removidas: **152 linhas**

## 2) Estrutura do Projeto

```text
case-monks/
|-- opps_corrupted.xlsx        # Base original (entrada)
|-- opps_corrigido.xlsx        # Base limpa (saída)
|-- relatorio_erros.html       # Validação e auditoria de qualidade
|-- analise.html               # Análise comercial com gráficos e insights
|-- apresentacao.pdf           # Apresentação executiva final
|-- README.md
`-- src/
    |-- cleaning.py            # Limpeza e padronização
    |-- error_report.py        # Relatório de erros
    `-- analysis_report.py     # Relatório analítico
```

## 3) Dependências e Setup

Pré-requisitos:

- Python 3.10+
- `pip`

Instalação das bibliotecas:

```bash
python -m pip install pandas numpy openpyxl plotly
```

## 4) Como Executar (Ponta a Ponta)

No diretório raiz do projeto:

```bash
python src/cleaning.py
python src/error_report.py
python src/analysis_report.py
```

Arquivos gerados:

- `opps_corrigido.xlsx`
- `relatorio_erros.html`
- `analise.html`

Arquivo final de apresentação (já pronto para submissão):

- `apresentacao.pdf`

## 5) Premissas de Negócio

### 5.1 Remoção de duplicatas (152 linhas)

A base estava inflada por duplicidades de oportunidade, então removi 152 linhas para evitar:

- inflação artificial do pipeline;
- previsão de receita superestimada;
- priorização comercial incorreta.

Regra aplicada por `Opportunity_ID`:

1. manter o registro com **maior completude** (mais campos úteis preenchidos);
2. em empate, manter o de **`Created_Date` mais antiga**.

Com esse critério, melhoro a consistência analítica e reduzo ruído operacional sem descaracterizar o histórico comercial.

### 5.2 Tratamento de nulos

- `Opportunity_ID` nulo/vazio: remoção do registro (sem chave mínima de rastreabilidade).
- Campos categóricos nulos: preenchimento com `Unknown`.
- Campos numéricos nulos: preenchimento com `0` após parse robusto.

Meu racional de negócio aqui foi preservar continuidade das análises, explicitar ausência de informação e evitar que dashboards/indicadores quebrem por dados faltantes.

### 5.3 Consistência cambial

- Harmonização entre `Amount_Currency` e `Total_Product_Amount_Currency`.
- Padronização de código monetário por linha.
- **Sem conversão FX** (não há taxa/data de câmbio confiáveis no dataset).

## 6) Interação com IA

Durante o desenvolvimento, utilizei LLMs como **assistentes de codificação** em um fluxo de revisão técnica Sênior/Pleno para:

- estruturar e refinar scripts Python;
- revisar regras de limpeza e validação;
- elevar clareza e robustez dos entregáveis técnicos.

Além disso, o **Google NotebookLM** foi utilizado para:

- sintetizar a narrativa de negócios;
- apoiar a interpretação dos insights extraídos;
- gerar o design e a narrativa do arquivo final **`apresentacao.pdf`**.

## 7) Entregáveis Finais do Case

- `opps_corrigido.xlsx`
- `relatorio_erros.html`
- `analise.html`
- `apresentacao.pdf`
- `README.md`

## 8) Limitações e Próximos Passos

Limitações atuais:

- ausência de política de FX histórico para consolidação em moeda única;
- dependência de qualidade da entrada no CRM;
- insights sensíveis à atualização de base e sazonalidade.

Próximos passos recomendados:

1. Validation rules por estágio no CRM.
2. Bloqueio de duplicidade por chave canônica (`Opportunity_ID` + `Account_ID`).
3. Monitoramento recorrente de qualidade de dados (SLA e data owner).
4. Evolução para camada contínua de governança RevOps.
