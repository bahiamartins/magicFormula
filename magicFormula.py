import yahooquery as yf
import locale
import pandas as pd
from concurrent.futures.thread import ThreadPoolExecutor as Executor
import time
import datetime
import os

from simbols import simbolos

locale.setlocale(locale.LC_ALL, 'pt_BR')

def main():
    t1 = time.perf_counter()
    startProcess()
    t2 = time.perf_counter()
    print(f'Rodou em :{t2 - t1} segundos')


def startProcess():

    all_data = []

    print('Processando Stocks')

    for ticker in simbolos:

        with Executor() as executor:
            r = executor.submit(generateData, ticker)
            #print(r.result())
            if r.result():
                all_data.append(r.result())

    df = pd.DataFrame(all_data)
    df = df.sort_values(by='MagicIndex', ascending=False, ignore_index=True)

    if not os.path.exists("output"):
        os.makedirs('output')
        
    output = os.path.join(os.getcwd(), 'output/')
    fileName = f'magicFormula_{datetime.datetime.now().strftime("%d%m%Y-%H%M%S")}.xlsx'
    filePath = os.path.join(output, fileName)
    df.to_excel(filePath, index = False)
    

def generateData(simbol):

    simbol_ = f'{simbol}.SA'
    print('------')
    print('')
    print('')
    print('Stock ', simbol)
    
    ticker = yf.Ticker(simbol_, 
        asynchronous=True
    )

    try:
        CP = float(ticker.financial_data[simbol_]['currentPrice'])
    except:
        return None

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
        prMo = None


    try:
        #pegar ebit atualizado
        #iloc -1 pega ultma linha do pandas
        ebit = ticker.income_statement(frequency='q').iloc[['-1']].loc[:,'EBIT']
        balance = ticker.balance_sheet(frequency='q').iloc[['-1']]
    except:
        #pegar ebit anual
        try:
            ebit = ticker.income_statement(frequency='a').iloc[['-1']].loc[:,'EBIT']
            balance = ticker.balance_sheet(frequency='q').iloc[['-1']]
        except:
            return None
    
    ebit = float(ebit.iloc[0])
    
    print('EBIT ', ebit)
    

    try:
        marketCap = int(ticker.price[simbol_]['marketCap'])
    except:
        marketCap = 0
        #return None

    print('marketCap ', marketCap)

    #ROIC
    # Retorno sobre Capital
    # ROIC = EBIT / EV
    # EV = capital de giro líquido + ativos fixos líquidos

    try:
        EV = balance.loc[:,'TotalAssets'] + balance.loc[:,'MachineryFurnitureEquipment']
        EV = int(EV.iloc[0])
    except:
        return None

    print('EV ', EV)
    
    if not ebit > 1 or EV is None:
        print('Ebit negativo')
        return None

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

    try:
        DY = round(ticker.summary_detail[simbol_]['dividendYield']*100, 2)
    except:
        DY = 0
    
    #filtra só empresas que pagam dividendos maior que x%
    #if not DY > 3:
    #    print('Dividendos baixos')
    #    return None
        
    recommendationTrend = ticker.financial_data[simbol_]

    try:
        sector = ticker.asset_profile[simbol_]['sector']
    except:
        sector = '---'
    
    try:
        name = ticker.price[simbol_]['longName']
    except:
        name = None
    
    data = {
        'Ticker': simbol,
        'Empresa': name,
        'Setor': sector,
        'MagicIndex': MAGIC_IDX,
        'MagicMomentumIndex': magic_momentum_idx,
        'Price Momentum': prMo,
        'EarningYield': EY,
        'ROIC': ROIC,
        'DividendosPercentual': DY ,
        'PrecoAcao': CP,
        'PrecoAcao6meses': round(CPn, 2),
        'DifPrecoAcao': round(pr, 2),
        'RecomendacaoCompraVenda': recommendationTrend['recommendationKey'],
        'MarketCap': marketCap,
        'Ebit (Lajir)': ebit,
        'CapitalTangivelEmpresa': EV,
        'ValorMercadoEmpresa': invCap
    }
    
    return data


if __name__ == '__main__':
    main()