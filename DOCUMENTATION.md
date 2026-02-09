# Magic Formula — Documentação Técnica

Documentação do script `magicFormula.py`: estrutura, funções, fórmulas e fluxo de dados.

---

## 1. Visão geral

O script calcula o **Magic Index** (Fórmula Mágica de Joel Greenblatt) para ações brasileiras listadas em `simbols.py`, usando dados do Yahoo Finance via `yahooquery`.  
Gera uma planilha Excel com empresas positivas (ordenadas pelo Magic Index) e, em aba separada, empresas com EBIT negativo (ordenadas por EBIT).

- **Entrada:** lista de tickers em `simbols.py` (ex.: `AALR3`, `ABCB4`).
- **Saída:** arquivo Excel em `output/magicFormula_DDMMAAAA-HHMMSS.xlsx` (e opcionalmente `_partial_apos_erro` em caso de erro).

---

## 2. Dependências

| Pacote       | Uso                          |
|-------------|------------------------------|
| `yahooquery`| Dados financeiros (preço, balanço, DRE, etc.) |
| `pandas`    | DataFrames e export para Excel |
| `openpyxl`  | Motor Excel para `.xlsx`     |
| `curl_cffi` | Sessão HTTP (impersonate Chrome) para yahooquery |
| `python-dateutil` | Parsing de datas (`parse`) |

Instalação: `pip install -r requirements.txt`

---

## 3. Estrutura do módulo

```
magicFormula.py
├── main()                  # Ponto de entrada; mede tempo e chama startProcess()
├── save_spreadsheet()      # Monta DataFrames e grava Excel (final ou parcial)
├── startProcess()          # Loop por tickers; coleta dados e trata erros
├── generateData(simbol)    # Para um ticker: busca dados e retorna dict ou None
├── calculate_ebit()        # EBIT a partir da DRE
├── calculate_ey()          # Earning Yield (EBIT / EV)
├── calculate_price_momentum()   # Momentum de preço (6 meses)
├── calculate_total_debt()      # Dívida total do balanço
└── calculate_total_debt_alt()  # Cálculo alternativo de dívida
```

Locale: `pt_BR` para números.

---

## 4. Fluxo de execução

1. **main()**  
   Chama `startProcess()` e imprime o tempo total.

2. **startProcess()**  
   - Inicializa `all_data` e `negative_ebit_data`.  
   - Para cada `ticker` em `simbolos`:  
     - Chama `generateData(ticker)` (em thread).  
     - Se retornar dict:  
       - EBIT &lt; 0 → `negative_ebit_data`.  
       - Caso contrário → `all_data`.  
     - Em exceção: imprime erro, chama `save_spreadsheet(..., suffix='_partial_apos_erro')` e continua para o próximo ticker.  
   - Ao final, chama `save_spreadsheet(all_data, negative_ebit_data)` para o arquivo final.

3. **save_spreadsheet(all_data, negative_ebit_data, suffix='')**  
   - Constrói DataFrames, ordena (Magic Index ou EBIT), garante pasta `output/`, grava o Excel e retorna o caminho do arquivo.

---

## 5. Funções principais

### 5.1 `save_spreadsheet(all_data, negative_ebit_data, suffix='')`

- **Objetivo:** Gerar o Excel a partir das listas de resultados.
- **Parâmetros:**
  - `all_data`: lista de dicts (empresas com EBIT ≥ 0).
  - `negative_ebit_data`: lista de dicts (EBIT &lt; 0).
  - `suffix`: string opcional no nome do arquivo (ex.: `'_partial_apos_erro'`).
- **Comportamento:**
  - Cria DataFrames, ordena por `MagicIndex` (positivas) e por `Ebit (Lajir)` (negativas).
  - Nome do arquivo: `magicFormula_DDMMAAAA-HHMMSS{suffix}.xlsx`.
- **Retorno:** caminho absoluto do arquivo gerado.

---

### 5.2 `startProcess()`

- Itera sobre `simbolos`.
- Para cada ticker usa `ThreadPoolExecutor` e `generateData(ticker)`.
- Agrupa resultados em `all_data` e `negative_ebit_data`.
- Em caso de exceção: salva planilha parcial com sufixo `_partial_apos_erro` e continua.
- No fim chama `save_spreadsheet()` para o resultado final.

---

### 5.3 `generateData(simbol)`

- **Entrada:** ticker sem sufixo (ex.: `AALR3`). Internamente usa `simbol_.SA` para Yahoo.
- **Saída:**  
  - Um **dict** com todas as colunas descritas na seção 6, ou  
  - **None** (dados insuficientes, ex.: sem preço atual).

Resumo do que a função faz:

1. Cria sessão `curl_cffi` (impersonate Chrome) e `yf.Ticker(simbol_.SA, ...)`.
2. Obtém preço atual; se falhar, retorna `None`.
3. Nome, setor, dividend yield, recomendação, price momentum (6 meses).
4. Define frequência (anual `'a'` ou trimestral `'q'`) com base na data do último balanço.
5. DRE e EBIT via `calculate_ebit()`; se EBIT inválido, retorna `None` (ou dict para EBIT negativo, conforme trecho de código para “Ebit negativo”).
6. Balanço, valuation, market cap, classificação (NANOCAP … LARGECAP).
7. ROIC: EBIT / EV (EV = Patrimônio Líquido + Dívida Líquida; dívida por `calculate_total_debt` / `calculate_total_debt_alt`).
8. Métricas adicionais: CGL por ação, VPA, dívida líquida/EBIT, etc.
9. Earning Yield com `calculate_ey(ebit, balance, CP, valuation)`.
10. Magic Index = EY + ROIC (e opcionalmente Magic Momentum com peso em price momentum).
11. Monta e retorna o dict de saída.

Qualquer exceção não tratada dentro de `generateData` propaga para `startProcess`, que salva a planilha parcial e segue para o próximo ticker.

---

## 6. Fórmulas e helpers

### 6.1 EBIT — `calculate_ebit(income_statement)`

Tenta, em ordem:

1. Campo `EBIT` da DRE.  
2. EBIT = `NetIncome + InterestExpense + TaxProvision`.  
3. EBIT = `OperatingIncome + OtherIncomeExpense`.  
4. EBIT = `EBITDA - ReconciledDepreciation`.

Retorna valor numérico ou `None` se não conseguir.

---

### 6.2 Earning Yield — `calculate_ey(ebit, balance, current_stock_price, valuation)`

- Primeiro tenta usar `EnterpriseValue` de `valuation`.  
- Senão: EV = Market Cap + Total Debt - Cash (com `calculate_total_debt` / `calculate_total_debt_alt` e `OrdinarySharesNumber`).  
- EY = EBIT / EV (retorno em decimal; no restante do código é multiplicado por 100 para percentual).

---

### 6.3 Price Momentum — `calculate_price_momentum(ticker, months=6)`

- Série de preços de fechamento em `ticker.history(period=f'{months}mo')`.  
- Momentum = (preço mais recente - preço mais antigo) / preço mais antigo.  
- Retorno: percentual (ex.: 15.5 para 15,5%).

---

### 6.4 Dívida total — `calculate_total_debt(balance)` e `calculate_total_debt_alt(balance)`

- **calculate_total_debt:** tenta `TotalDebt`; depois combina dívida de curto e longo prazo; fallbacks com totais e provisions.  
- **calculate_total_debt_alt:** fórmula alternativa com TotalAssets, Goodwill, Equity e LongTermProvisions; ou fallback com provisions.

Usados para EV e, indiretamente, para ROIC e EY.

---

## 7. Dicionário de dados (saída)

Campos do dict retornado por `generateData` (e colunas do Excel):

| Coluna                         | Descrição resumida                    |
|--------------------------------|----------------------------------------|
| Ticker                         | Símbolo (ex.: AALR3)                   |
| Empresa                        | Nome longo (Yahoo)                      |
| Setor                          | Setor da empresa                       |
| CapType                        | NANOCAP, MICROCAP, SMALLCAP, MIDCAP, LARGECAP |
| MagicIndex                     | EY + ROIC (empresas com EBIT positivo) |
| MagicMomentumIndex             | Índice com peso em momentum            |
| Price Momentum                 | Variação de preço em 6 meses (%)       |
| EarningYield                   | EBIT/EV (%)                            |
| ROIC                           | EBIT/EV (%) (capital tangível)         |
| DividendosPercentual           | Dividend yield (%)                     |
| PrecoAcao                      | Preço atual                            |
| PrecoAcao6meses                | Preço 6 meses atrás                    |
| DifPrecoAcao                   | Diferença de preço                     |
| Capital de Giro Liquido por Ação | CGL por ação                         |
| CGL/PrecoAcao                  | Razão ou '-'                           |
| Valor Patrimonial por Ação     | VPA                                    |
| PrecoAcao / VPA                | 'subvalorizada' ou '-'                 |
| Dívida Líquida                 | (TotalDebt - Cash) / EBIT              |
| RecomendacaoCompraVenda        | recommendationKey (Yahoo)               |
| Ebit (Lajir)                   | EBIT                                   |
| CapitalTangivelEmpresa        | EV (capital tangível)                  |
| ValorMercadoEmpresa            | Market cap                             |

Para empresas com EBIT negativo, o dict tem estrutura similar, com foco em `Ebit (Lajir)` e sem Magic Index/EY/ROIC completos.

---

## 8. Tratamento de erros e planilha parcial

- **Por ticker:** qualquer exceção em `generateData(ticker)` é capturada em `startProcess()`.  
- **Ação:** imprime mensagem do tipo `Erro ao processar {ticker}: {e}`, chama `save_spreadsheet(all_data, negative_ebit_data, suffix='_partial_apos_erro')` e **continua** para o próximo ticker.  
- **Arquivo parcial:** `output/magicFormula_DDMMAAAA-HHMMSS_partial_apos_erro.xlsx` com todos os tickers já processados com sucesso até o momento do erro.  
- **Arquivo final:** ao terminar o loop, sempre é gerado o Excel final com todos os resultados obtidos na execução.

Assim, um erro em um ticker não interrompe o processo e não se perde o trabalho já feito.

---

## 9. Configuração e uso

- **Tickers:** editar a lista `simbolos` em `simbols.py`.  
- **Execução:**  
  `python3 magicFormula.py`  
- **Saída:** pasta `output/` no diretório atual; arquivos `magicFormula_*.xlsx`.

---

## 10. Abas do Excel

1. **Empresas Positivas** — EBIT ≥ 0; ordenação por `MagicIndex` decrescente.  
2. **Empresas EBIT Negativo** — EBIT &lt; 0; ordenação por `Ebit (Lajir)` crescente (menos negativo primeiro).  
A segunda aba só é criada se houver pelo menos uma empresa com EBIT negativo.

---

Esta documentação descreve o comportamento e a estrutura do `magicFormula.py` para manutenção e extensão do script.
