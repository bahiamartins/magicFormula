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
    'PETR3',
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

    output = os.path.join(os.getcwd(), 'output/')
    fileName = f'magicFormula_{datetime.datetime.now().strftime("%d%m%Y-%H%M%S")}.xlsx'
    filePath = os.path.join(output, fileName)
    df.to_excel(filePath, index = False)


def generateData(simbol):

    simbol = f'{simbol}.SA'
    print('------')
    print('')
    print('')
    print('Stock ', simbol)
    tkr = yf.Ticker(simbol).all_modules[simbol]

    #print(tkr['balanceSheetHistoryQuarterly']['balanceSheetStatements'])
    #print('')
    #print(tkr['incomeStatementHistoryQuarterly']['incomeStatementHistory'])

    #somar ebit por quarter
    #3º quarter
    balanceDate = tkr['balanceSheetHistoryQuarterly']['balanceSheetStatements'][0]['endDate']
    if '09-30' in balanceDate:
        #somar ebit 3 quarters
        ebit1 = tkr['incomeStatementHistoryQuarterly']['incomeStatementHistory'][0]['ebit']
        ebit2 = tkr['incomeStatementHistoryQuarterly']['incomeStatementHistory'][1]['ebit']
        ebit3 = tkr['incomeStatementHistoryQuarterly']['incomeStatementHistory'][2]['ebit']

        ebit = ebit1 + ebit2 + ebit3

        balance = tkr['balanceSheetHistoryQuarterly']['balanceSheetStatements'][0]

    elif '06-30' in balanceDate:
        #somar ebit 2 quarters
        ebit1 = tkr['incomeStatementHistoryQuarterly']['incomeStatementHistory'][0]['ebit']
        ebit2 = tkr['incomeStatementHistoryQuarterly']['incomeStatementHistory'][1]['ebit']

        ebit = ebit1 + ebit2

        balance = tkr['balanceSheetHistoryQuarterly']['balanceSheetStatements'][0]

    elif '03-31' in balanceDate:
        #pegar ebit atualizado
        ebit = tkr['incomeStatementHistoryQuarterly']['incomeStatementHistory'][0]['ebit']

        balance = tkr['balanceSheetHistoryQuarterly']['balanceSheetStatements'][0]
    else:
        #pegar ebit anual
        ebit = tkr['incomeStatementHistory']['incomeStatementHistory'][0]['ebit']

        balance = tkr['balanceSheetHistory']['balanceSheetStatements'][0]

    try:
        marketCap = tkr['summaryDetail']['marketCap']
    except:
        marketCap = '---'

    #ROIC
    # Retorno sobre Capital
    # ROIC = EBIT / EV
    # EV = capital de giro líquido + ativos fixos líquidos

    try:
        capitalGiro = tkr['financialData']['operatingCashflow']
    except:
        try:
            capitalGiro = tkr['financialData']['freeCashflow']
        except:
            try:
                capitalGiro = balance['totalCurrentAssets']
            except:
                print('sem capital giro')
                return None
    ativoFixoLiquido = balance['netTangibleAssets']
    EV = capitalGiro + ativoFixoLiquido
    
    if not ebit > 1 or EV == 0:
        print('Ebit negativo')
        return None

    ROIC = ebit / EV
    ROIC = round(ROIC, 2)

    if ROIC <= 0:
        print('Sem retorno de capital')
        return None
    
    #Valor de Mercado da Empresa
    #print(balance)
    #print('')
    #print(tkr['financialData'])
    try:
        totalDebt = tkr['financialData']['totalDebt']
    except:
        totalDebt = balance['shortLongTermDebt'] + balance['longTermDebt']
    
    try:
        totalStockholderEquity = balance['totalStockholderEquity']
    except:
        totalStockholderEquity = 0
    try:
        goodWill = balance['goodWill']
    except:
        goodWill = 0
    try:
        cash = balance['cash']
    except:
        cash = 0
    try:
        retainedEarnings = balance['retainedEarnings']
    except:
        retainedEarnings = 0

    # Earning Yield
    #Resultado de Rendimento
    # EY = EBIT / Valor de Mercado da Empresa
    #invCap = Valor de Mercado da Empresa = valor de mercado + débito líquido remunerado a juros
    invCap = totalDebt + totalStockholderEquity + goodWill + cash #+ retainedEarnings
    
    if ebit > 1:
        EY = ebit / invCap
        EY = round(EY, 2)
    else:
        print('Earning Yield negativo')
        return None
    
    # magic index
    MAGIC_IDX = round(EY + ROIC, 2)
    
    try:
        DY = round(tkr['summaryDetail']['dividendYield']*100, 2)
    except:
        DY = 0
    
    #filtra só empresas que pagam dividendos maior que x%
    if not DY > 3:
        print('Dividendos baixos')
        return None
        
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
        'EarningYield': EY,
        'ROIC': ROIC,
        'DividendosPercentual': DY ,
        'PrecoAcao': tkr['financialData']['currentPrice'],
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