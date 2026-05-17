# Processo Seletivo Estágio | Case | Dados & Analytics (Mercado de Capitais e M&A)
## Documentação Técnica e Analítica do Projeto

> **Área:** Deal & Capital Markets (DCM) — Análise de Pipeline Comercial  
> **Ferramentas:** Excel · Power Query · Tabelas Dinâmicas · Gráficos Dinâmicos  
> **Objetivo:** Consolidar, tratar e analisar bases operacionais fragmentadas para geração de inteligência comercial e suporte à decisão executiva no contexto de M&A e Mercado de Capitais.

---

## Sumário

1. [Introdução e Contexto do Problema](#1-introdução-e-contexto-do-problema)
2. [Estrutura dos Dados](#2-estrutura-dos-dados)
3. [Principais Inconsistências Identificadas](#3-principais-inconsistências-identificadas)
4. [Processo de Limpeza e Transformação](#4-processo-de-limpeza-e-transformação)
5. [Modelagem Analítica e Regras de Negócio](#5-modelagem-analítica-e-regras-de-negócio)
6. [Decisões Técnicas](#6-decisões-técnicas)
7. [Fluxo ETL Completo](#7-fluxo-etl-completo)
8. [Ferramentas Utilizadas](#8-ferramentas-utilizadas)
9. [KPIs e Métricas Desenvolvidas](#9-kpis-e-métricas-desenvolvidas)
10. [Insights Acionáveis](#10-insights-acionáveis)
11. [Melhorias Futuras](#11-melhorias-futuras)
12. [Conclusão Executiva](#12-conclusão-executiva)

---

## 1. Introdução e Contexto do Problema

Em operações de Mercado de Capitais e M&A, a consolidação e interpretação de dados desempenham papel estratégico na identificação de oportunidades comerciais, priorização de clientes e suporte à tomada de decisão executiva.

Neste case, o desafio consiste em integrar diferentes bases de dados relacionadas ao pipeline de deals, interações comerciais, histórico de operações, informações cambiais e cadastro de clientes, considerando problemas comuns encontrados em ambientes corporativos reais, como inconsistências cadastrais, divergências de nomenclatura, formatos distintos, duplicidades e ausência de informações relevantes.

A proposta busca transformar dados dispersos em uma visão analítica estruturada e orientada a negócio, permitindo acompanhar indicadores de performance comercial, estimar receitas esperadas, avaliar o avanço das oportunidades no funil e apoiar estratégias de priorização comercial.

Além do tratamento e modelagem dos dados, o desenvolvimento do projeto contempla a construção de métricas, análises exploratórias e recomendações executivas acionáveis, com foco em gerar insights relevantes para liderança e apoiar decisões baseadas em dados.


---

## 2. Estrutura dos Dados

O projeto consolidou cinco bases operacionais distintas, cada uma com granularidade e propósito específicos:

```
Case-Itau.xlsx
├── interactions_log          → Log de interações comerciais com clientes
├── deals_pipeline            → Histórico de estágios e operações do pipeline
├── company_master            → Cadastro mestre de clientes (tabela dimensional)
├── fx_rates                  → Taxas de câmbio de referência por data e moeda
├── taxas_diarias             → Tabela auxiliar de conversão e normalização BRL
├── Media_Tipo                → Médias históricas de fee por tipo de negócio
├── Media_Tipo_Cargo          → Médias históricas de fee por tipo + cargo bancário
├── Dashboard                 → Painel executivo consolidado
└── PIVOT                     → Tabelas dinâmicas de apoio analítico
```

### Descrição das tabelas principais

| Tabela | Granularidade | Chave Principal | Papel no Modelo |
|---|---|---|---|
| `interactions_log` | 1 linha = 1 interação | `interaction_id` | Fato — interações comerciais |
| `deals_pipeline` | 1 linha = 1 deal/estágio | `deal_id` | Fato — pipeline de negócios |
| `company_master` | 1 linha = 1 cliente | `client_id` | Dimensão — cadastro de clientes |
| `fx_rates` | 1 linha = 1 taxa por data/moeda | `data + moeda` | Dimensão — referência cambial |
| `Media_Tipo` | 1 linha = 1 tipo de negócio | `tipo_negocio` | Dimensão — imputação de fee |

A separação entre tabelas fato e dimensão foi uma decisão deliberada para garantir que os joins fossem realizados de forma controlada, sem duplicação de registros e com rastreabilidade completa das transformações aplicadas.


---

## 3. Principais Inconsistências Identificadas

Antes de qualquer transformação, foi realizada uma varredura analítica completa das bases para mapear todos os problemas de qualidade. O diagnóstico revelou inconsistências em múltiplas camadas — de formato, tipagem, semântica e integridade referencial.

### 3.1 Datas inconsistentes — `interactions_log`

A coluna de data da tabela de interações apresentava dois formatos distintos coexistindo na mesma série histórica:

- **Formato 1:** Data simples com sufixo `Z` (padrão UTC ISO 8601) — ex: `2024-03-15Z`
- **Formato 2:** Data completa com horário e fuso horário — ex: `2024-03-15T14:32:00+03:00`

O Power Query não conseguia converter esses registros diretamente para o tipo `DATA`, pois interpretava os sufixos como parte do valor textual. Qualquer tentativa de tipagem direta gerava erros em massa ou silenciosamente descartava registros, comprometendo a integridade da série temporal.

### 3.2 Problemas de encoding — `company_master`

A coluna `client_name_canonical` apresentava caracteres especiais corrompidos, resultado de inconsistência de encoding na origem do arquivo. Nomes com acentuação (ã, ç, é, ó) apareciam como sequências ilegíveis, impossibilitando joins confiáveis entre a tabela de clientes e as demais bases.

### 3.3 Tipagem incorreta — `deals_pipeline`

Múltiplos problemas de tipagem foram identificados na tabela principal de operações:

- Colunas de data importadas como texto, impedindo cálculos de ciclo de vida e aging
- Valores monetários ultrapassando o limite do tipo `Currency` do Excel em alguns registros
- Coluna `Probabilidade` com separadores decimais mistos (`.` e `,`), fazendo com que o Power Query interpretasse parte dos valores como texto e parte como número — sem gerar erro visível, apenas silenciosamente distorcendo os cálculos

### 3.4 Valores monetários em múltiplas moedas sem normalização

A coluna de valor financeiro das operações continha registros em `BRL`, `USD` e `EUR` sem qualquer indicação de taxa de conversão aplicada. Somar esses valores diretamente — como seria feito em uma análise ingênua — produziria um número sem significado financeiro real, misturando moedas como se fossem equivalentes.

### 3.5 Classificações ambíguas e campos técnicos sem semântica de negócio

- Campos booleanos representados por `0` e `1` sem legenda
- Status operacionais em inglês (`Won`, `Lost`, `Open`) sem padronização para o contexto brasileiro
- Classificações internas `low`, `med`, `high` sem definição formal documentada
- Coluna `tracking` com valores em inglês sem tradução ou mapeamento para PT-BR

### 3.6 Inconsistências operacionais — anomalias de processo

Durante a análise, foram identificadas duas categorias de anomalias que vão além de problemas técnicos de formato — são inconsistências de processo com impacto direto na confiabilidade dos indicadores:

**Anomalia 1 — NO_SHOW / Sem Resposta com dados operacionais preenchidos:**
Registros classificados como "Não Comparecimento" ou "Sem Resposta" apresentavam simultaneamente anotações internas e tempo de duração diferente de `null`. Isso indica que houve esforço operacional real documentado, mesmo sem retorno formal do cliente.

**Anomalia 2 — Falsos Positivos e Falsos Negativos no pipeline:**

| Cenário | Status Final | Duração/Tentativa | Risco Analítico |
|---|---|---|---|
| Falso Positivo | `Proposta Aceita` | `null` | Aceitação sem evidência de interação real |
| Falso Negativo | `Resultado Imparcial` | `null` | Neutro sem fundamento operacional registrado |

Contabilizar esses registros como resultados consolidados infla artificialmente o win rate e distorce as projeções de receita do pipeline.


---

## 4. Processo de Limpeza e Transformação

Todo o processo de ETL foi executado via Power Query, garantindo que cada transformação fosse documentada como uma etapa auditável na consulta. A ordem das etapas foi planejada para respeitar dependências — por exemplo, a normalização de datas precede qualquer join temporal, e a padronização de moedas precede qualquer cálculo financeiro.

### 4.1 Normalização de datas — `interactions_log`

**Problema:** Dois formatos de data coexistindo na mesma coluna, ambos com sufixos que impediam conversão direta.

**Solução aplicada:**

```excel
=VALOR(ESQUERDA([Data]; 10))
```

A lógica extrai os primeiros 10 caracteres de qualquer string de data — que correspondem invariavelmente à porção `AAAA-MM-DD` — e converte o resultado para um valor numérico de data reconhecido pelo Excel. Em seguida, a formatação foi aplicada para exibição no padrão brasileiro `DD/MM/AAAA`.

**Por que essa abordagem:** É simples, determinística e não depende de parsing complexo. Funciona para ambos os formatos identificados sem necessidade de lógica condicional. 

**Resultado:** 100% dos registros convertidos para tipo `DATA` sem perda de linhas.

---

### 4.2 Correção de encoding — `company_master`

**Problema:** Caracteres especiais corrompidos na coluna `client_name_canonical`, comprometendo a integridade dos joins entre bases.

**Solução aplicada:** Ajuste do encoding de origem diretamente na consulta do Power Query, alterando a codificação do arquivo fonte para `UTF-8`. Os casos residuais que não foram corrigidos automaticamente — geralmente nomes com combinações incomuns de caracteres — foram tratados manualmente via substituição direta na consulta.

**Impacto:** A correção foi crítica para garantir que os joins entre `interactions_log`, `deals_pipeline` e `company_master` fossem realizados corretamente. Um nome de cliente corrompido não encontraria correspondência na tabela dimensional, gerando registros órfãos e distorcendo qualquer análise por cliente.

---

### 4.3 Reclassificação de campos booleanos — `company_master`

**Problema:** Campos indicadores representados por `0` e `1` sem semântica clara para usuários de negócio.

**Solução aplicada:** Substituição via Power Query para `"Privado"` e `"Público"`, tornando os campos diretamente interpretáveis em tabelas dinâmicas e no dashboard sem necessidade de legenda técnica.

**Decisão de design:** Manter os valores originais como referência em coluna auxiliar foi considerado, mas descartado por adicionar complexidade desnecessária. O campo booleano original não seria utilizado em nenhum cálculo — apenas em filtros e segmentações visuais.

---

### 4.4 Padronização completa — `deals_pipeline`

Esta foi a tabela com maior volume de transformações, dado que é a tabela fato central do modelo.

#### Tipagem de datas e valores monetários

Todas as colunas de data foram forçadas para o tipo `DATA` no Power Query. Valores monetários foram padronizados como `MOEDA` em BRL após a normalização cambial (descrita na seção 4.5).

#### Correção da coluna `Probabilidade`

**Problema:** A coluna apresentava separadores decimais mistos. Registros com `.` como separador eram interpretados como número pelo Power Query, enquanto registros com `,` eram interpretados como texto — sem gerar erro, apenas produzindo `null` silencioso nos cálculos.

**Solução aplicada:** Criação da coluna `Probabilidade Correção` via transformação no Power Query:

```powerquery
= Text.Replace([Probabilidade], ".", ",")
```

Após a substituição, a coluna foi convertida para tipo `Número Decimal`, garantindo que todos os valores fossem tratados uniformemente como proporções entre 0 e 1.

**Decisão de auditoria:** O campo original `Probabilidade` foi preservado na consulta como referência de auditoria. A coluna `Probabilidade Correção` é a utilizada em todos os cálculos subsequentes.


#### Padronização de status operacional

Os status originais em inglês foram mapeados para um padrão único em PT-BR:

| Status Original | Padrão Adotado |
|---|---|
| `Won` / `Closed Won` | `Lucro` |
| `Lost` / `Closed Lost` | `Perda` |
| `Open` / `Active` | `Aberto` |

#### Criação da coluna `Rastreamento Final`

A coluna `tracking` original continha valores em inglês (`Timing`, `Strategy`, `Market conditions`, `Unknown`) que representam o motivo associado ao status do deal. Foi criada a coluna `Rastreamento Final` com a tradução para PT-BR:

| Valor Original | Rastreamento Final |
|---|---|
| `Timing` | `Dentro do Prazo` |
| `Strategy` | `Estratégia` |
| `Market conditions` | `Condições Mercado` |
| `Unknown` | `Desconhecido` |


---

### 4.5 Normalização cambial — FX Rates

**Problema:** Valores financeiros das operações registrados em `BRL`, `USD` e `EUR` sem taxa de conversão aplicada. Somar esses valores diretamente produziria um número financeiramente inválido.

**Tentativas iniciais e por que falharam:**

A primeira abordagem foi alterar a localidade/região do Excel para forçar a interpretação dos valores. Isso não resolveu o problema de moedas mistas — apenas alterou a formatação de exibição. A segunda tentativa foi converter diretamente para o tipo `MOEDA` no Power Query, o que gerou erros em registros com valores em USD e EUR, removendo linhas e criando inconsistências.

**Solução definitiva — tabela auxiliar FX Rates:**

Foi criada uma tabela auxiliar com a seguinte estrutura:

| Data | Moeda | Taxa (para BRL) |
|---|---|---|
| 01/01/2024 | USD | 4,92 |
| 01/01/2024 | EUR | 5,34 |
| ... | ... | ... |

O processo de normalização seguiu três etapas no Power Query:

1. **Merge** da tabela `deals_pipeline` com a tabela `fx_rates` usando as chaves `Data` + `Moeda`
2. **Expansão** da coluna de taxa para trazer o valor correto por linha
3. **Criação de coluna personalizada** para o valor convertido:

```powerquery
= if [Moeda] = "BRL" then [Valor Negócio]
  else [Valor Negócio] * [Taxa_BRL]
```

**Resultado:** Nenhuma linha foi perdida. Todas as operações passaram a ter um valor financeiro comparável em BRL, independentemente da moeda original. A coluna `Valor Negócio BRL` tornou-se a base para todos os cálculos financeiros subsequentes.

---

## 5. Modelagem Analítica e Regras de Negócio

Com os dados limpos e padronizados, a etapa seguinte foi construir as métricas derivadas que transformam dados operacionais em inteligência comercial.

### 5.1 Imputação de `fee_bps` ausentes

**Problema:** Parte das operações não possuía o campo `fee_bps` (pontos-base de comissão) preenchido. Utilizar uma média global para imputação seria tecnicamente incorreto — a estrutura de fee de uma operação de M&A difere fundamentalmente da emissão de um título de dívida (BOND) ou de uma debênture.

**Solução — hierarquia de imputação:**

Foi implementada uma lógica hierárquica em três níveis:

1. **Nível 1 (prioridade máxima):** Se o valor original de `fee_bps` existe na operação, ele é mantido sem alteração
2. **Nível 2:** Se o campo está vazio, utiliza-se a média histórica calculada por **Tipo de Negócio + Cargo Bancário** (tabela `Media_Tipo_Cargo`)
3. **Nível 3 (fallback):** Se não há registros suficientes para o agrupamento Tipo + Cargo, utiliza-se a média por **Tipo de Negócio** apenas (tabela `Media_Tipo`)

Essa hierarquia garante que a imputação seja sempre a mais específica possível, reduzindo o viés introduzido pela substituição.

O resultado foi armazenado na coluna `expected_fee_brl`, que é a utilizada em todos os cálculos de receita.

---

### 5.2 Cálculo da receita esperada bruta — `expected_fee_brl`

Com o `fee_bps_final` disponível para todas as linhas, o cálculo da comissão bruta esperada foi implementado como coluna personalizada no Power Query:

```powerquery
([Valor Negócio BRL] / 100000000000) * ([fee_bps_final] / 10000)
```


### 5.3 Métricas de ciclo de vida e aging

**Cycle time (`cycle_time_dias`):** Diferença em dias entre a data de criação do deal e a data de fechamento (para deals concluídos). Permite identificar a duração média do ciclo comercial por tipo de operação e setor.

**Aging (`aging_dias`):** Para deals com status `Aberto`, calcula o tempo em dias desde a criação até a data atual. Permite identificar operações estagnadas no pipeline que podem requerer ação comercial.

Ambas as métricas foram calculadas diretamente no Power Query usando operações de subtração entre colunas do tipo `DATA`.


---

## 6. Decisões Técnicas

Esta seção documenta as principais decisões de design e as alternativas consideradas, com o racional de cada escolha.

### 6.1 Por que Power Query e não fórmulas diretas no Excel?

O Power Query foi escolhido como motor principal de transformação por três razões:

1. **Auditabilidade:** Cada etapa de transformação fica registrada como um passo nomeado na consulta, criando um histórico completo e reversível de todas as operações aplicadas
2. **Reprodutibilidade:** Ao atualizar os dados de origem, todas as transformações são reaplicadas automaticamente sem intervenção manual
3. **Separação de camadas:** Mantém os dados brutos intactos na origem e as transformações isoladas na camada de consulta, seguindo o princípio de não modificar a fonte

### 6.2 Por que preservar colunas originais como auditoria?

Em todas as transformações que alteraram valores existentes (probabilidade, status, rastreamento), a coluna original foi preservada com o nome original e a coluna transformada recebeu um novo nome. Isso garante:

- Rastreabilidade completa de cada transformação
- Capacidade de reverter ou comparar valores antes/depois
- Transparência para auditoria interna ou revisão por terceiros

### 6.3 Por que não usar média global para imputação de `fee_bps`?

Uma média global de fee ignoraria a heterogeneidade estrutural dos produtos de DCM. O fee de uma operação de M&A (tipicamente 1-2% do valor da transação) é fundamentalmente diferente do fee de uma emissão de debêntures ou de um BOND. Usar uma média global introduziria viés sistemático nas projeções de receita, superestimando fees de produtos de baixo custo e subestimando os de alto custo.


### 6.4 Tratamento de falsos positivos e negativos

A decisão foi **não excluir** esses registros do dataset, mas **isolá-los analiticamente**. Excluir registros sem validação operacional seria uma intervenção irreversível nos dados. Isolá-los permite que a liderança tome a decisão de tratamento com base em contexto de negócio — que o analista de dados não possui.


---

## 7. Fluxo ETL Completo

O diagrama abaixo representa o fluxo completo de processamento dos dados, da origem até a camada analítica:

```
FONTES BRUTAS
│
├── interactions_log (CSV/Sheet)
│   ├── [ETL] Normalização de datas → VALOR(ESQUERDA(..., 10))
│   ├── [ETL] Correção de encoding → UTF-8 + ajuste manual residual
│   └── → interactions_log_clean
│
├── deals_pipeline (Sheet)
│   ├── [ETL] Tipagem de datas → tipo DATA
│   ├── [ETL] Correção de probabilidade → coluna Probabilidade Correção
│   ├── [ETL] Reclassificação de status → Lucro / Perda / Aberto
│   ├── [ETL] Tradução de tracking → Rastreamento Final (PT-BR)
│   ├── [ETL] Merge com fx_rates → coluna Valor Negócio BRL
│   ├── [ETL] Imputação fee_bps → fee_bps_final (hierarquia 3 níveis)
│   ├── [CALC] expected_fee_brl = (Valor BRL / 1e11) * (fee_bps_final / 10000)
│   ├── [CALC] expected_revenue = expected_fee_brl * Probabilidade Final
│   ├── [CALC] cycle_time_dias = Data Fechamento - Data Criação
│   └── → deals_pipeline_clean
│
├── company_master (Sheet)
│   ├── [ETL] Reclassificação booleanos → Privado / Público
│   ├── [ETL] Padronização de cabeçalhos
│   
│
└── fx_rates (Sheet)
    ├── [ETL] Validação de cobertura de datas e moedas
    └── → fx_rates_clean (tabela de lookup)

                    ↓

CAMADA ANALÍTICA (Power Query / Excel)
│
├── Joins: deals_pipeline_clean ←→ company_master_clean (via client_id)
├── Joins: deals_pipeline_clean ←→ fx_rates_clean (via data + moeda)
├── Joins: deals_pipeline_clean ←→ Media_Tipo / Media_Tipo_Cargo (via tipo + cargo)
│
└── TABELA ANALÍTICA CONSOLIDADA
    ├── Todas as métricas derivadas
    ├── Todos os campos padronizados
    └── Pronta para Tabelas Dinâmicas e Dashboard

                    ↓

DASHBOARD EXECUTIVO
├── Distribuição de status (Lucro / Perda / Aberto)
├── Volume financeiro por setor e tipo
├── Win rate consolidado
├── Aging médio do pipeline
├── Efetividade comercial por canal
├── Top 10 deals por receita esperada
└── Segmentações dinâmicas por setor e produto
```


---

## 8. Ferramentas Utilizadas

### Microsoft Excel

O Excel foi utilizado como ambiente principal de desenvolvimento, visualização e entrega. Suas funcionalidades de Tabelas Dinâmicas e Gráficos Dinâmicos permitiram construir o dashboard executivo com segmentações interativas sem necessidade de ferramentas externas de BI.

Fórmulas estruturadas foram utilizadas pontualmente para transformações que não eram viáveis diretamente no Power Query — como a extração de datas com `=VALOR(ESQUERDA(...))` — e para validações de consistência entre colunas.

### Power Query

O Power Query foi o motor central de ETL do projeto. Todas as transformações estruturais — tipagem, normalização, joins, imputação, criação de colunas calculadas — foram implementadas como etapas auditáveis dentro das consultas.

A escolha pelo Power Query em detrimento de fórmulas diretas no Excel foi deliberada: garante que o pipeline de transformação seja reproduzível, documentado e atualizável sem retrabalho manual.

As principais funcionalidades utilizadas foram:

| Funcionalidade | Aplicação no Projeto |
|---|---|
| `Table.TransformColumnTypes` | Tipagem de datas e valores monetários |
| `Table.ReplaceValue` | Substituição de separadores decimais |
| `Table.AddColumn` | Criação de colunas calculadas e condicionais |
| `Table.NestedJoin` | Merge com tabelas de lookup (FX, Media_Tipo) |
| `Table.RenameColumns` | Padronização de cabeçalhos |
| `Table.SelectColumns` | Seleção e reordenação de colunas para entrega |


---

## 9. KPIs e Métricas Desenvolvidas

O modelo analítico foi estruturado para responder às principais perguntas do case:

### KPIs de Pipeline

| Métrica | Definição | Relevância |
|---|---|---|
| **Win Rate** | `Deals Lucro / (Deals Lucro + Deals Perda)` | Efetividade comercial geral |
| **Receita Esperada Total** | `Σ expected_revenue` | Projeção de caixa do pipeline |
| **Ticket Médio** | `Σ Valor Negócio BRL / n deals` | Perfil financeiro das operações |
| **Aging Médio** | `Média de aging_dias (deals Abertos)` | Saúde do pipeline em aberto |
| **Cycle Time Médio** | `Média de cycle_time_dias (deals fechados)` | Eficiência do ciclo comercial |

### KPIs de Interações Comerciais

| Métrica | Definição | Relevância |
|---|---|---|
| **Taxa de Resposta** | `Interações com retorno / Total de interações` | Engajamento da carteira |
| **Índice de Retrabalho** | `Interações NO_SHOW ou Sem Resposta com duração > 0` | Custo operacional oculto |
| **Efetividade por Canal** | `Conversões por tipo de comunicação` | Otimização de abordagem |

### KPIs de Risco Analítico

| Métrica | Definição | Relevância |
|---|---|---|
| **Falsos Positivos** | `Deals "Proposta Aceita" com duração null` | Risco de inflação do pipeline |
| **Falsos Negativos** | `Deals "Imparcial" com duração null` | Risco de subnotificação |
| **Cobertura de fee_bps** | `% de deals com fee original vs. imputado` | Qualidade da base de receita |

---

## 10. Insights Acionáveis

### 10.1 A inteligência oculta no "Sem Resposta"

A análise das interações revelou que registros classificados como `NO_SHOW` ou `Sem Resposta` frequentemente continham anotações detalhadas e tempo de duração preenchido. Isso prova que houve esforço operacional real — múltiplas tentativas, tempo alocado, registro de contexto — mesmo sem retorno formal do cliente.

A interpretação estratégica é contraintuitiva: **se a instituição insiste repetidamente em um cliente que não responde, é porque esse cliente possui uma carteira valiosa.** O "Sem Resposta" não é ausência de interesse institucional — é evidência de alto potencial percebido.

Isso abre caminho para um KRI de eficiência comercial: cruzar o índice de retrabalho por cliente com o valor de carteira estimado, identificando os casos onde o custo de servir está desproporcionalmente alto em relação ao retorno esperado.

### 10.2 Concentração de deals nas etapas iniciais do funil

A análise de distribuição por estágio revelou concentração de operações nas fases de `Pitch` e `Origination`, com queda abrupta nas etapas de `Closing`. Isso pode indicar gargalos operacionais na transição entre estágios intermediários — um ponto de atenção para a gestão do pipeline.

### 10.3 Heterogeneidade de fee por produto

A análise das médias por tipo de negócio confirmou a hipótese que motivou a imputação hierárquica: a dispersão de `fee_bps` entre produtos é significativa. Usar uma média global teria introduzido erro sistemático nas projeções de receita de produtos com estrutura de fee muito diferente da média.

### 10.4 Impacto dos falsos sinais no win rate

Os registros identificados como potenciais falsos positivos (Proposta Aceita sem interação registrada) representam um risco direto de inflação do win rate. Em um contexto de fechamento de trimestre, esses registros podem criar uma percepção de performance comercial superior à realidade operacional.


---

## 11. Melhorias Futuras

### 11.1 Entity Resolution mais robusta

A padronização de nomes de clientes entre bases foi realizada via `client_id` e correção de encoding. Uma evolução natural seria implementar uma lógica de similaridade textual para identificar clientes que aparecem com nomes ligeiramente diferentes entre bases — por exemplo, "Itaú BBA" vs. "Itaú BBA S.A." vs. "ITAU BBA". No contexto atual (Excel/Power Query), isso poderia ser feito com uma tabela de-para manual para os casos mais críticos.

### 11.2 Validação operacional dos falsos positivos/negativos

Os registros identificados como potenciais falsos sinais foram isolados analiticamente, mas a validação definitiva requer confirmação operacional da área de negócio. Uma melhoria futura seria criar um fluxo de revisão periódica desses registros com a equipe comercial.

### 11.3 Atualização automática das taxas FX

A tabela de taxas cambiais foi construída manualmente para o período coberto pelos dados. Em um ambiente de produção, essa tabela seria alimentada automaticamente via conexão com uma fonte de dados de mercado. No contexto Excel/Power Query, uma alternativa viável seria uma conexão com uma planilha Google Sheets pública que consolide taxas diárias do Banco Central.

### 11.4 KRI de retrabalho operacional

O insight sobre clientes com alto índice de `NO_SHOW` e `Sem Resposta` abre espaço para um indicador formal de eficiência comercial. A construção desse KRI dependeria de um cruzamento entre o log de interações e o valor de carteira estimado por cliente — informação que pode ser derivada do `deals_pipeline` mas que requereria validação com a área de negócio.

---

## 12. Conclusão Executiva

Este projeto demonstrou que o trabalho de qualidade de dados em um ambiente de Mercado de Capitais vai muito além de limpeza técnica. Cada decisão de transformação tem impacto direto na confiabilidade das projeções financeiras, na precisão dos indicadores comerciais e na qualidade das decisões estratégicas que serão tomadas com base nesses dados.

O processo cobriu cinco dimensões fundamentais:

1. **Integridade técnica** — normalização de formatos, tipagem correta, eliminação de ambiguidades de encoding e separadores
2. **Consistência semântica** — padronização de status, traduções, reclassificações com racional documentado
3. **Confiabilidade financeira** — normalização cambial rigorosa, imputação hierárquica de fee, cálculo de receita ajustada por risco
4. **Pensamento crítico** — identificação de anomalias operacionais (falsos positivos/negativos, retrabalho oculto) que não seriam detectadas por uma análise superficial
5. **Comunicação executiva** — transformação de dados operacionais brutos em indicadores acionáveis e visualizações interpretáveis pela liderança

O resultado é um modelo analítico que não apenas responde às perguntas óbvias do pipeline — quantos deals, qual o volume, qual o win rate — mas que também levanta as perguntas que a liderança ainda não sabia que precisava fazer: onde está o retrabalho oculto? Quais registros de sucesso são, na verdade, ruído de sistema? Onde o esforço comercial está sendo investido de forma desproporcional ao retorno esperado?

---

*Desenvolvido como case técnico para processo seletivo — Área de Deal & Capital Markets (DCM) · Itaú BBA*  
*Ferramentas: Microsoft Excel · Power Query · Tabelas Dinâmicas · Gráficos Dinâmicos*

