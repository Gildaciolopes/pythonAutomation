import pandas as pd
print(pd.__version__)

tabela_vendas = pd.read_excel('Vendas.xlsx')

pd.set_option('display.max_columns', None)
print(tabela_vendas)


# faturamento por loja
faturamento = tabela_vendas.groupby('ID Loja')['Valor Final'].sum()
print(faturamento)

# quantidade de produtos vendidos por loja
quantidade = tabela_vendas.groupby('ID Loja')['Quantidade'].sum()
print(quantidade)