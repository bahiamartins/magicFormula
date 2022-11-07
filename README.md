# Programa Fórmula Mágica

Script criado seguindo as regras do Livro

### A Fórmula Mágica de Joel Greenblatt para Bater o Mercado de Ações

```bash
# criando o ambiente virtual
$ python3 -m venv .venv
# Iniciando o ambiente virtual
$ source .venv/bin/activate
# Instalando as dependencias
$ pip install -r requirements.txt
```

O programa pode ser executado usando o seguinte comando
```bash
$ python3 magicFormula.py
```

Aí ele cria uma planilha excel na pasta output.

### Como funciona a fórmula

**Passo 1 - Retorno sobre o Capital - ROIC**

O primeiro passo é calcular o quanto uma empresa devolve a seus acionistas pelo valor investido

LAJIR (Lucro antes de juros e impostos) == EBIT

ROIC = EBIT / Capital Giro Liquido + Ativos Fixos Liquidos)


**Passo 2 - Resultados de Rendimento - Earning Yield (EY)**

O segundo passo é calcular o quanto uma empresa gera de retorno

EY = EBIT / Valor da Empresa

Sendo o Valor da Empresa = o valor de mercado + débito líquido


**Passo 3 - Index da Fórmula Mágica**

Agora é só somar os resultados dos passos 1 e 2 e classificar por maior pontuação.
Pronto!

Agora você já tem uma lista potencial de empresas para comprar ações que estào, no momento de gerar o Magic Index, baratas em relação ao mercado.


**Passo 4 - Bonus**

Eu filtrei no código apenas empresas que pagam mais de 4% de dividendos.

Bons Investimentos!!!


