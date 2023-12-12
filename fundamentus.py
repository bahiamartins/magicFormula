from bs4 import BeautifulSoup
import requests
from simbols import simbolos
import locale
import pandas as pd
from concurrent.futures.thread import ThreadPoolExecutor as Executor
import time
import datetime
import os


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
    fileName = f'magicFormula_fundamentus_{datetime.datetime.now().strftime("%d%m%Y-%H%M%S")}.xlsx'
    filePath = os.path.join(output, fileName)
    df.to_excel(filePath, index = False)


def generateData(simbol):

    print('Stock ', simbol)
    
    ticker = simbol
    url = f'https://www.fundamentus.com.br/detalhes.php?papel={ticker}'
    agent = {"User-Agent":'Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.115 Safari/537.36'}
    page = requests.get(url, headers=agent)
    soup = BeautifulSoup(page.content, 'html.parser')

    base = soup.find_all('table')

    #current price
    try:
        CP = base[0].select('tr > td')[3].span.text
        CP = float(CP.replace(',', '.'))
    except:
        CP = 0
    
    # ROIC
    try:
        ROIC = base[2].select('tr')[7].select('td')[5].span.text
        ROIC = float(ROIC.replace('%', '').replace(',', '.'))
    except:
        ROIC = 0


    # Earning Yield
    try:
        EBIT = base[4].select('tr')[3].select('td')[1].span.text
        EBIT = int(EBIT.replace('.', ''))
    except:
        EBIT = 0

    try:
        EV = base[1].select('tr')[1].select('td')[1].span.text
        EV = int(EV.replace('.', ''))
        EY = EBIT / EV
        EY = round(EY*100, 2)
    except:
        EY = 0
    

    # magic index
    try:
        MAGIC_IDX = round(EY + ROIC, 2)
    except:
        MAGIC_IDX = 0

    try:
        DY = base[2].select('tr')[8].select('td')[3].span.text
        DY = float(DY.replace('%', '').replace(',', '.'))
    except:
        DY = 0

    name = base[0].select('tr')[2].select('td')[1].span.text
    sector = base[0].select('tr')[3].select('td')[1].span.a.text
    
    data = {
        'Ticker': simbol,
        'Empresa': name,
        'Setor': sector,
        'MagicIndex': MAGIC_IDX,
        'EarningYield': EY,
        'ROIC': ROIC,
        'DividendosPercentual': DY ,
        'PrecoAcao': CP
    }
    
    return data


if __name__ == '__main__':
    main()