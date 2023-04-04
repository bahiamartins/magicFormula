import yahooquery as yf
import locale
import pandas as pd
from concurrent.futures.thread import ThreadPoolExecutor as Executor
import time
import datetime
import os

locale.setlocale(locale.LC_ALL, 'pt_BR')


simbolos = [
    'RRRP3',
    'ALPA4',
    'ABEV3',
    'AMBP3',
    'AMER3',
    'ARZZ3',
    'ASAI3',
    'AURE3',
    'AZUL4',
    'B3SA3',
    'BPAN4',
    'BBSE3',
    'BRML3',
    'BBDC3',
    'BBDC4',
    'BRAP4',
    'BBAS3',
    'BRKM5',
    'BRFS3',
    'BPAC11',
    'CRFB3',
    'CBAV3',
    'CCRO3',
    'CMIG4',
    'CIEL3',
    'COGN3',
    'CPLE6',
    'CSAN3',
    'CPFE3',
    'CMIN3',
    'CVCB3',
    'CYRE3',
    'DXCO3',
    'ECOR3',
    'ELET3',
    'ELET6',
    'EMBR3',
    'ENAT3',
    'ENBR3',
    'ENGI11',
    'ENEV3',
    'EGIE3',
    'EQTL3',
    'EZTC3',
    'FLRY3',
    'GGBR4',
    'GOAU4',
    'GOLL4',
    'NTCO3',
    'SOMA3',
    'HAPV3',
    'HYPE3',
    #'IGTI11',
    'IRBR3',
    'ITSA4',
    'ITUB4',
    'JBSS3',
    'KLBN11',
    'RENT3',
    'LWSA3',
    'LREN3',
    'MDIA3',
    'MGLU3',
    'MRFG3',
    'CASH3',
    'BEEF3',
    'MOVI3',
    'MRVE3',
    'MULT3',
    'PCAR3',
    #'PETR3',
    'PETR4',
    'PRIO3',
    'PETZ3',
    'PSSA3',
    'QUAL3',
    'RADL3',
    'RAIZ4',
    'RDOR3',
    'RAIL3',
    'SBSP3',
    'SANB11',
    'STBP3',
    'SMTO3',
    'CSNA3',
    'SLCE3',
    'SULA11',
    'SUZB3',
    'TAEE11',
    'VIVT3',
    'TIMS3',
    'TOTS3',
    'UGPA3',
    'USIM5',
    'VALE3',
    'VAMO3',
    'VIIA3',
    'VBBR3',
    'WEGE3',
    'YDUQ3',
]

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
    tkr = ticker.all_modules[simbol_]
    
    # price momentum

    # prMo = (CP - CPn) / CPn
    # CP = Closing price in the current period
    # CPn = Closing price N periods ago
    # considerar n = 6 meses
    
    try:
        CPn = ticker.history(period='6mo')['close'][0]
        #print('CPn ', CPn)

        CP = tkr['financialData']['currentPrice']
        #print('CP ', CP)

        pr = CP - CPn
        #print('CP - CPn ', pr)

        prMo = pr / CPn
        prMo = round(prMo*100, 2)
        
    except:
        prMo = None

    #print(tkr['balanceSheetHistoryQuarterly']['balanceSheetStatements'])
    #print('')
    #print(tkr['incomeStatementHistoryQuarterly']['incomeStatementHistory'])

    try:
        #pegar ebit atualizado
        ebit = tkr['incomeStatementHistoryQuarterly']['incomeStatementHistory'][0]['ebit']

        balance = tkr['balanceSheetHistoryQuarterly']['balanceSheetStatements'][0]
    except:
        #pegar ebit anual
        try:
            ebit = tkr['incomeStatementHistory']['incomeStatementHistory'][0]['ebit']
            balance = tkr['balanceSheetHistory']['balanceSheetStatements'][0]
        except:
            return None

    try:
        marketCap = tkr['summaryDetail']['marketCap']
    except:
        marketCap = 0
        return None

    #ROIC
    # Retorno sobre Capital
    # ROIC = EBIT / EV
    # EV = capital de giro líquido + ativos fixos líquidos

    try:
        EV = balance['totalCurrentAssets'] + balance['propertyPlantEquipment']
    except:
        return None
    
    if not ebit > 1 or EV == 0:
        print('Ebit negativo')
        return None

    ROIC = ebit / EV
    ROIC = round(ROIC*100, 2)

    if ROIC <= 0:
        print('Sem retorno de capital')
        return None
    
    #Valor de Mercado da Empresa
    #print(balance)
    #print('')
    #print(tkr['financialData'])

    # Earning Yield
    #Resultado de Rendimento
    # EY = EBIT / Valor de Mercado da Empresa
    #invCap = Valor de Mercado da Empresa = valor de mercado + débito líquido remunerado a juros
    
    invCap = marketCap + balance['totalCurrentLiabilities']

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
        DY = round(tkr['summaryDetail']['dividendYield']*100, 2)
    except:
        DY = 0
    
    #filtra só empresas que pagam dividendos maior que x%
    #if not DY > 3:
    #    print('Dividendos baixos')
    #    return None
        
    recommendationTrend = tkr['recommendationTrend']['trend'][3]
    buy = recommendationTrend['strongBuy'] + recommendationTrend['buy']
    sell = recommendationTrend['strongSell'] + recommendationTrend['sell']

    try:
        sector = tkr['summaryProfile']['sector']
    except:
        sector = '---'
    
    data = {
        'Ticker': simbol,
        'Empresa': tkr['quoteType']['longName'],
        'Setor': sector,
        'MagicIndex': MAGIC_IDX,
        'MagicMomentumIndex': magic_momentum_idx,
        'Price Momentum': prMo,
        'EarningYield': EY,
        'ROIC': ROIC,
        'DividendosPercentual': DY ,
        'PrecoAcao': tkr['financialData']['currentPrice'],
        'PrecoAcao6meses': round(CPn, 2),
        'DifPrecoAcao': round(pr, 2),
        'RecomendacaoCompraYahoo': buy,
        'RecomendacaoVendaYahoo': sell,
        'MarketCap': marketCap,
        'Ebit (Lajir)': ebit,
        'CapitalTangivelEmpresa': EV,
        'ValorMercadoEmpresa': invCap
    }
    
    return data


if __name__ == '__main__':
    main()