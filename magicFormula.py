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

    try:
        #iloc -1 pega ultma linha do pandas
        ebit = ticker.income_statement(frequency=frequency).iloc[['-1']].loc[:,'EBIT']
        ebit = int(ebit.iloc[0])
    except:
        ebit = 0

    print('EBIT ', ebit)
    
    balance = ticker.balance_sheet(frequency=frequency).iloc[['-1']]

    try:
        marketCap = int(balance.loc[:,'TotalAssets'].iloc[0])
    except:
        marketCap = 0

    try:
        marketCap = int(ticker.price[simbol_]['marketCap'])
    except:
        marketCap = 0
        #return None

    print('marketCap ', marketCap)

    capType = 'LARGECAP'
    makCp = int(marketCap)
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
    # ROIC = EBIT / EV
    # EV = capital de giro líquido + ativos fixos líquidos

    try:
        TotalAssets = balance.loc[:,'TotalAssets'].iloc[0]
        TotalAssets = int(TotalAssets)
    except:
        TotalAssets = 0

    try:
        MachineryFurnitureEquipment = balance.loc[:,'MachineryFurnitureEquipment'].iloc[0]
        MachineryFurnitureEquipment = int(MachineryFurnitureEquipment)
    except:
        MachineryFurnitureEquipment = 0
    
    EV = TotalAssets + MachineryFurnitureEquipment
    EV = int(EV)

    print('EV ', EV)
    
    if not ebit > 1 or EV is None or EV == 0:
        print('Ebit negativo')
        # Instead of returning None, we'll return the data for negative EBIT companies
        data = {
            'Ticker': simbol,
            'Empresa': name,
            'Setor': sector,
            'Price Momentum': prMo,
            'DividendosPercentual': DY,
            'PrecoAcao': CP,
            'PrecoAcao6meses': CPn,
            'DifPrecoAcao': pr,
            'RecomendacaoCompraVenda': recommendationKey,
            'Ebit (Lajir)': ebit,
            'CapitalTangivelEmpresa': EV,
            'ValorMercadoEmpresa': marketCap,
            'CapType': capType
        }
        return data

    ROIC = ebit / EV
    ROIC = round(ROIC*100, 2)

    print('ROIC ', ROIC)

    if ROIC <= 0:
        print('Sem retorno de capital')
        return None
    
    # Earning Yield
    #Resultado de Rendimento
    # EY = EBIT / Valor de Mercado da Empresa
    #invCap = Valor de Mercado da Empresa = valor de mercado + débito líquido remunerado a juros

    try:
        currentLiab = balance.loc[:,'CurrentLiabilities']
        currentLiab = int(currentLiab.iloc[0])
    except:
        currentLiab = 0
    
    invCap = marketCap + currentLiab

    print('invCap', invCap)

    if ebit > 1:
        EY = ebit / invCap
        EY = round(EY*100, 2)
    else:
        print('Earning Yield negativo')
        return None
    
    # magic index
    MAGIC_IDX = round(EY + ROIC, 2)

    #index com price momentum
    magic_momentum_idx = None
    if prMo:
        magic_momentum_idx = MAGIC_IDX + prMo
    
    #print('IDX ', magic_momentum_idx)
    
    if CPn:
        CPn = round(CPn, 2)
    
    if pr:
        pr = round(pr, 2)
    
    data = {
        'Ticker': simbol,
        'Empresa': name,
        'Setor': sector,
        'MagicIndex': MAGIC_IDX,
        'MagicMomentumIndex': magic_momentum_idx,
        'Price Momentum': prMo,
        'EarningYield': EY,
        'ROIC': ROIC,
        'DividendosPercentual': DY,
        'PrecoAcao': CP,
        'PrecoAcao6meses': CPn,
        'DifPrecoAcao': pr,
        'RecomendacaoCompraVenda': recommendationKey,
        'Ebit (Lajir)': ebit,
        'CapitalTangivelEmpresa': EV,
        'ValorMercadoEmpresa': marketCap,
        'CapType': capType
    }
    
    return data


if __name__ == '__main__':
    main()