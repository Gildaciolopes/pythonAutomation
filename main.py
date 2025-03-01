import pandas as pd
print(pd.__version__)

tabela_vendas = pd.read_excel('Vendas.xlsx')

pd.set_option('display.max_columns', None)
print(tabela_vendas)


# faturamento por loja
faturamento = tabela_vendas['ID Loja', 'Valor Final'].grtoupby('ID Loja').sum()
print(faturamento)