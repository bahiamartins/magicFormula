# Programa Fórmula Mágica

Script criado seguindo as regras do Livro

### A Fórmula Mágica de Joel Greenblatt para Bater o Mercado de Ações


Link do Livro [https://www.amazon.com.br/F%C3%B3rmula-M%C3%A1gica-Greenblatt-Bater-Mercado/dp/8557173601]


### Rodando o script

**windows**

```bash
# criando o ambiente virtual
$ py -m venv .venv       
# Iniciando o ambiente virtual
$ .venv/Scripts/activate.ps1
# Instalando as dependencias
$ pip install --upgrade pip
$ pip install -r requirements.txt
```

**unix**

```bash
# criando o ambiente virtual
$ python3 -m venv .venv
# Iniciando o ambiente virtual
$ source .venv/bin/activate
# Instalando as dependencias
$ pip install --upgrade pip
$ pip install -r requirements.txt
```

**O programa pode ser executado usando o seguinte comando**

```bash
$ python3 magicFormula.py
```

OU

```bash
$ python3 fundamentus.py
```

A diferença é que o magicFormula usa a API do Yahoo e o outro faz scrap no site fundamentus
A API do Yahoo costuma ficar fora do ar, então o site fundamentus pode ser uma alternativa

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

Agora você já tem uma lista potencial de empresas para comprar ações que estão, no momento de gerar o Magic Index, potencialmente baratas em relação ao mercado e que poderão se valorizar em um período de tempo aceitável.


**Classes de Empresas**

Para facilitar comparar empresas por porte, foi inserida a coluna CapType que permite filtrar empreas pelo valor de mercado de acordo com seus grupos:

- Nanocaps – Até 50 milhões
- Microcaps – Até 300 milhões
- Small Caps – Até 2 bilhões
- Mid Caps – Até 10 bilhões
- Large Caps – Até 200 bilhões
- Mega Caps – Mais de 200 bilhões

Lembrando que no Brasil não há empresas no momento no perfil Mega Caps

Bons Investimentos!!!


