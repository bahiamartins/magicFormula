import yahooquery as yf
import locale
import pandas as pd
from concurrent.futures.thread import ThreadPoolExecutor as Executor
import time
import datetime
from dateutil.parser import parse
import os
from curl_cffi import requests

from simbols import simbolos

locale.setlocale(locale.LC_ALL, 'pt_BR')


def main():
    t1 = time.perf_counter()
    startProcess()
    t2 = time.perf_counter()
    print(f'Rodou em :{t2 - t1} segundos')


def startProcess():

    all_data = []
    negative_ebit_data = []  # New list for companies with negative EBIT

    print('Processando Stocks')

    for ticker in simbolos:

        with Executor() as executor:
            r = executor.submit(generateData, ticker)
            #print(r.result())
            if r.result():
                if r.result().get('Ebit (Lajir)', 0) < 0:
                    negative_ebit_data.append(r.result())
                else:
                    all_data.append(r.result())

    df = pd.DataFrame(all_data)
    df = df.sort_values(by='MagicIndex', ascending=False, ignore_index=True)

    df_negative_ebit = pd.DataFrame(negative_ebit_data)
    if not df_negative_ebit.empty:
        df_negative_ebit = df_negative_ebit.sort_values(by='Ebit (Lajir)', ascending=True, ignore_index=True)

    if not os.path.exists("output"):
        os.makedirs('output')
        
    output = os.path.join(os.getcwd(), 'output/')
    fileName = f'magicFormula_{datetime.datetime.now().strftime("%d%m%Y-%H%M%S")}.xlsx'
    filePath = os.path.join(output, fileName)
    
    # Create Excel writer object
    with pd.ExcelWriter(filePath, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Empresas Positivas', index=False)
        if not df_negative_ebit.empty:
            df_negative_ebit.to_excel(writer, sheet_name='Empresas EBIT Negativo', index=False)


def generateData(simbol):

    simbol_ = f'{simbol}.SA'
    print('------')
    print('')
    print('')
    print('Stock ', simbol)

    session = requests.Session(impersonate="chrome")
    
    ticker = yf.Ticker(simbol_, 
        asynchronous=True,
        progress=True, 
        session=session
    )

    try:
        CP = float(ticker.financial_data[simbol_]['currentPrice'])
    except:
        return None

    # Get company name and sector early
    try:
        name = ticker.price[simbol_]['longName']
    except:
        name = None
    
    try:
        sector = ticker.asset_profile[simbol_]['sector']
    except:
        sector = '---'

    # Get dividend yield early
    try:
        DY = round(ticker.summary_detail[simbol_]['dividendYield']*100, 2)
    except:
        DY = 0

    # Get recommendation early
    recommendationTrend = ticker.financial_data[simbol_]
    try:
        recommendationKey = recommendationTrend['recommendationKey']
    except:
        recommendationKey = None

    # price momentum

    # prMo = (CP - CPn) / CPn
    # CP = Closing price in the current period
    # CPn = Closing price N periods ago
    # considerar n = 6 meses

    prMo = calculate_price_momentum(ticker)
    
    try:
        CPn = float(ticker.history(period='6mo')['close'][0])

        pr = CP - CPn
        #print('CP - CPn ', pr)

        prMo = pr / CPn
        prMo = round(prMo*100, 2)
        
    except:
        CPn = None
        prMo = None
        pr = None

    #check if companies have quarters data depending on the month
    # 31-Março
    # 30-Junho
    # 30-Setembro
    # 31-Dezembro

    frequency = 'a'

    keyStats = ticker.key_stats

    try:
        recentQuarter = parse(keyStats[simbol_]['mostRecentQuarter']).date()
    except:
        recentQuarter = None

    if recentQuarter and recentQuarter.month != 12:
        frequency = 'q'

    if recentQuarter and recentQuarter.month == 12:
        if datetime.date.today().month == 1:
            frequency = 'q'
    
    print('frequency ', frequency)

    income_statement = ticker.income_statement(frequency=frequency).iloc[['-1']]
    ebit = calculate_ebit(income_statement)

    print('EBIT ', ebit)
    
    balance = ticker.balance_sheet(frequency=frequency).iloc[['-1']]
    valuation = ticker.valuation_measures.iloc[['-1']]

    try:
        marketCap = int(valuation.loc[:, 'MarketCap'].iloc[0])
    except:
        pass
    try:    
        marketCap =  int(ticker.price[simbol_]['marketCap'])
    except:
        marketCap = 0

    print('marketCap ', marketCap)

    capType = 'LARGECAP'
    makCp = marketCap
    if makCp <= 50000000: #50.000.000
        capType = 'NANOCAP'
    if makCp > 50000000 and makCp <= 300000000: #300.000.000
        capType = 'MICROCAP'
    if makCp > 300000000 and makCp <= 2000000000: #2.000.000.000
        capType = 'SMALLCAP'
    if makCp > 2000000000 and makCp <= 10000000000: # 10.000.000.000
        capType = 'MIDCAP'
    if makCp > 10000000000: # 10.000.000.000
        capType = 'LARGECAP'
    
    
    #ROIC
    # Retorno sobre Capital
    # ROIC = EBIT / EV sendo EV = (Patrimônio Líquido + Dívida Líquida)
    # EV = capital de giro líquido + ativos fixos líquidos

    try:
        TotalEquity = balance.loc[:,'TotalEquityGrossMinorityInterest'].iloc[0]  # Patrimônio Líquido
    except:
        TotalEquity = balance.loc[:,'StockholdersEquity'].iloc[0]

    TotalDebt = calculate_total_debt(balance)
    if not TotalDebt:
        TotalDebt = calculate_total_debt_alt(balance)

    try:
        Cash = balance.loc[:,'CashAndCashEquivalents'].iloc[0]
    except:
        Cash = balance.loc[:,'CashCashEquivalentsAndShortTermInvestments'].iloc[0]
    
    EV = TotalEquity + (TotalDebt - Cash)
    EV = int(EV)

    print('EV ', EV)

    # stocks undervalued
    # Capital de Giro Líquido
    # CGL por acao = (CurrentAssets - TotalLiabilitiesNetMinorityInterest) / OrdinarySharesNumber
    # CGL por Ação > Preço da Ação e a empresa tem:
    # Baixa dívida (NetDebt),
    # Fluxo de caixa positivo (CashFlow)

    try:
        cash_flow = ticker.cash_flow(frequency=frequency).iloc[['-1']]
        FreeCashFlow = cash_flow.loc[:,'FreeCashFlow'].iloc[0]
        CurrentAssets = balance.loc[:,'CurrentAssets'].iloc[0]
        ordinarySharesNumber = balance.loc[:,'OrdinarySharesNumber'].iloc[0]
        currentLiabilities = balance.loc[:,'CurrentLiabilities'].iloc[0]

        cglPorAcao = (int(FreeCashFlow) + int(CurrentAssets) - int(currentLiabilities)) / int(ordinarySharesNumber)
        cglPorAcao = round(cglPorAcao, 2)
    except:
        cglPorAcao = None

    print('Capital de Giro Liquido por Ação ', cglPorAcao)

    # Valor Patrimonial por Ação
    # VPA = StockholdersEquity / OrdinarySharesNumber
    # Se o preço da ação está abaixo do VPA, pode indicar subvalorização.
    
    try:
        stockholdersEquity = balance.loc[:,'StockholdersEquity'].iloc[0]
        ordinarySharesNumber = balance.loc[:,'OrdinarySharesNumber'].iloc[0]

        vpa = int(stockholdersEquity) / int(ordinarySharesNumber)
        vpa = round(vpa, 2)
    except:
        vpa = None

    print('Valor Patrimonial por Ação ', vpa)
    
    # Dívida Líquida = TotalDebt - CashAndCashEquivalents  
    # Formula = (TotalDebt - CashAndCashEquivalents) / EBIT

    try:
        totalDebt = balance.loc[:,'TotalDebt'].iloc[0]
        cashAndCashEquivalents = balance.loc[:,'CashAndCashEquivalents'].iloc[0]

        dl = (int(totalDebt) - int(cashAndCashEquivalents)) / ebit
        dl = round(dl, 2)
    except:
        dl = None

    print('Dívida Líquida ', dl)


    if EV is None or ebit < 1:
        print('Ebit negativo')
        # Instead of returning None, we'll return the data for negative EBIT companies
        data = {
            'Ticker': simbol,
            'Empresa': name,
            'Setor': sector,
            'CapType': capType,
            'Price Momentum': prMo,
            'DividendosPercentual': DY,
            'PrecoAcao': CP,
            'PrecoAcao6meses': CPn,
            'DifPrecoAcao': pr,
            'Capital de Giro Liquido por Ação': cglPorAcao,
            'Valor Patrimonial por Ação': vpa,
            'Dívida Líquida': dl,
            'RecomendacaoCompraVenda': recommendationKey,
            'Ebit (Lajir)': ebit,
            'CapitalTangivelEmpresa': EV,
            'ValorMercadoEmpresa': marketCap,
            
        }
        return data

    ROIC = ebit / EV
    ROIC = round(ROIC*100, 2)

    print('ROIC ', ROIC)

    if ROIC <= 0:
        print('Sem retorno de capital')
    
    # Earning Yield
    #Resultado de Rendimento
    # EY = EBIT / Valor de Mercado da Empresa

    if ebit > 1:
        EY = calculate_ey(ebit, balance, CP, valuation)
        EY = round(EY*100, 2)
        print('EY ', EY)
    else:
        print('Earning Yield negativo')
    
    # magic index
    MAGIC_IDX = round(EY + ROIC, 2)

    #index com price momentum
    if all([ROIC, EY, prMo]):
        # Ponderação: 50% fundamentos, 50% momentum
        magic_momentum_idx = (0.4 * ROIC + 0.4 * EY) + (0.2 * prMo)
    else:
        magic_momentum_idx = None
    
    #print('IDX ', magic_momentum_idx)
    
    if CPn:
        CPn = round(CPn, 2)
    
    if pr:
        pr = round(pr, 2)
    
    data = {
        'Ticker': simbol,
        'Empresa': name,
        'Setor': sector,
        'CapType': capType,
        'MagicIndex': MAGIC_IDX,
        'MagicMomentumIndex': magic_momentum_idx,
        'Price Momentum': prMo,
        'EarningYield': EY,
        'ROIC': ROIC,
        'DividendosPercentual': DY,
        'PrecoAcao': CP,
        'PrecoAcao6meses': CPn,
        'DifPrecoAcao': pr,
        'Capital de Giro Liquido por Ação': cglPorAcao,
        'Valor Patrimonial por Ação': vpa,
        'Dívida Líquida': dl,
        'RecomendacaoCompraVenda': recommendationKey,
        'Ebit (Lajir)': ebit,
        'CapitalTangivelEmpresa': EV,
        'ValorMercadoEmpresa': marketCap
    }
    
    return data


def calculate_ebit(income_statement):
    try:
        # Tenta usar o EBIT direto (se disponível)
        ebit = income_statement.loc[:,'EBIT'].iloc[0]
        if not pd.isna(ebit):
            return ebit
    except:
        pass
        
    try:
        # Método 1 (NetIncome + Juros + Impostos)
        net_income = income_statement.loc[:,'NetIncome'].iloc[0]
        interest_expense = income_statement.loc[:,'InterestExpense'].iloc[0]
        tax_provision = income_statement.loc[:,'TaxProvision'].iloc[0]
        ebit = net_income + interest_expense + tax_provision
        if not pd.isna(ebit):
            return ebit
    except:
        pass
        
    try:
        # Método 2 (OperatingIncome + Itens Não Operacionais)
        operating_income = income_statement.loc[:,'OperatingIncome'].iloc[0]
        other_income = income_statement.loc[:,'OtherIncomeExpense'].iloc[0]
        ebit = operating_income + other_income
        if not pd.isna(ebit):
            return ebit
    except:
        pass
        
    try:
        # Método 3 (EBITDA - Depreciação)
        ebitda = income_statement.loc[:,'EBITDA'].iloc[0]
        depreciation = income_statement.loc[:,'ReconciledDepreciation'].iloc[0]
        ebit = ebitda - depreciation
        if not pd.isna(ebit):
            return ebit
    except:
        return None

#calcular earning yield
def calculate_ey(ebit, balance, current_stock_price, valuation):

    #if EnterpriseValue
    try:
        ev = valuation.loc[:, 'EnterpriseValue'].iloc[0]
        if not pd.isna(ev):
            ey = ebit / ev
            ey = round(ey*100, 2)
            return ey
    except:
        pass


    # 1. Calcular Market Cap
    shares_outstanding = balance.loc[:,'OrdinarySharesNumber'].iloc[0]
    market_cap = current_stock_price * shares_outstanding
    
    # 2. Calcular Total Debt
    total_debt = calculate_total_debt(balance)
    if not total_debt:
        total_debt = calculate_total_debt_alt(balance)
    
    # 3. Obter Cash
    cash = balance.loc[:,'CashAndCashEquivalents'].iloc[0]
    
    # 4. Calcular Enterprise Value
    ev = market_cap + total_debt - cash
    
    # 5. Calcular Earning Yield
    ey = ebit / ev if ev != 0 else 0
    
    return ey


def calculate_price_momentum(ticker, months=6):
    # Interpretação do Price Momentum:
    # Momentum	    Interpretação	              Sinal
    # > +15%	    Forte tendência de alta	      📈 Bullish
    # +5% a +15%	Tendência moderada de alta	  ↗️ Positivo
    # -5% a +5%	    Lateralizado/Neutro	          ➡️ Neutro
    # -5% a -15%	Tendência moderada de baixa   ↘️ Negativo
    # < -15%	    Forte tendência de baixa	  📉 Bearish
    
    try:
        # Obter dados históricos ordenados cronologicamente (mais antigo primeiro)
        hist = ticker.history(period=f'{months}mo')
        
        # Verificar se há dados suficientes
        if len(hist) < 2:
            return None
        
        # Preço mais antigo (início do período)
        oldest_price = hist['close'].iloc[0]
        
        # Preço mais recente (final do período)
        latest_price = hist['close'].iloc[-1]
        
        # Cálculo do momentum
        momentum = (latest_price - oldest_price) / oldest_price
        return round(momentum * 100, 2)  # Retorna em percentual
    
    except Exception as e:
        print(f"Erro no cálculo de momentum: {str(e)}")
        return None


def calculate_total_debt(balance):

    try:
        TotalDebt = balance.loc[:,'TotalDebt'].iloc[0]  # Dívida total
        print(f"TotalDebt 1: {str(TotalDebt)}")
        if pd.notna(TotalDebt):
            return TotalDebt
        print('no totaldebt')
    except:
        pass

    try:
        # Se não tiver 'TotalDebt', calcule:
        CurrentDebt = balance.loc[:,'CurrentDebtAndCapitalLeaseObligation'].iloc[0]
        print(f"CurrentDebt: {str(CurrentDebt)}")
        if pd.notna(CurrentDebt):
            return CurrentDebt
        print('no currentdebt')
    except:
        CurrentDebt = None
    
    if CurrentDebt:

        try:
          LongTermDebt = balance.loc[:,'LongTermDebtAndCapitalLeaseObligation'].iloc[0]
          print(f"LongTermDebt: {str(LongTermDebt)}")
          if pd.notna(LongTermDebt):
            return LongTermDebt
          print('no longtermdebt')
        except:
          LongTermDebt = 0

        if CurrentDebt is not None:
            TotalDebt = CurrentDebt + LongTermDebt
            print(f"TotalDebt 2: {str(TotalDebt)}")
            if pd.notna(TotalDebt):
                return TotalDebt
    
    try:
        # Extração dos campos necessários
        total_assets = balance.loc[:,'TotalAssets'].iloc[0]
        if pd.notna(total_assets):
            return total_assets
        print('no totalassets')

    except:
        total_assets = None
    
    if total_assets:

        try:
            net_tangible_assets = balance.loc[:,'NetTangibleAssets'].iloc[0]
            if pd.notna(net_tangible_assets):
                return net_tangible_assets
            print('no nettangibleassets')
        except:
            net_tangible_assets = None
        
        try:
          long_term_provisions = balance.loc[:,'LongTermProvisions'].iloc[0]
          if pd.notna(long_term_provisions):
            return long_term_provisions
          print('no longtermprovisions')
        except:
          long_term_provisions = 0

        # Cálculo da dívida total
        total_debt = (total_assets - net_tangible_assets) + long_term_provisions
        if pd.notna(total_debt):
          print(f"total_debt: {str(total_debt)}")
          return max(total_debt, 0)  # Garante valor não-negativo
    try:
        # 1. Dívida de Curto Prazo (Current Debt)
        current_debt = balance['CurrentCapitalLeaseObligation'].iloc[0]
        
        # 2. Dívida de Longo Prazo (Long-Term Debt)
        # Calculada como o total de arrendamentos menos a parcela corrente
        lease_obligations = balance['CapitalLeaseObligations'].iloc[0]
        current_lease = balance['CurrentCapitalLeaseObligation'].iloc[0]
        long_term_debt = lease_obligations - current_lease
        
        # 3. Total Debt
        if pd.notna(current_debt) and pd.notna(lease_obligations) and pd.notna(current_lease):
            total_debt = current_debt + long_term_debt
            print(f"TotalDebt 3: {str(total_debt)}")
            return total_debt
    
    except KeyError as e:
        print(f"Campo não encontrado: {str(e)}")
    except Exception as e:
        print(f"Erro no cálculo: {str(e)}")
        return None


def calculate_total_debt_alt(balance):
    try:
        total_assets = balance.loc[:,'TotalAssets'].iloc[0]
        goodwill_intangibles = balance.loc[:,'GoodwillAndOtherIntangibleAssets'].iloc[0]
        equity = balance.loc[:,'CommonStockEquity'].iloc[0]
        long_term_provisions = balance.loc[:,'LongTermProvisions'].iloc[0]
        
        # Fórmula alternativa
        total_debt = total_assets - goodwill_intangibles - equity + long_term_provisions
        return max(total_debt, 0)
        
    except:
        pass

    try:
        # Fallback simplificado
        total_debt = balance['LongTermProvisions'].iloc[0] * 2
        if not pd.isna(total_debt):
            return total_debt
    except:
        return None


if __name__ == '__main__':
    main()