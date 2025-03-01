import pandas as pd

tabela_vendas = pd.read_excel('Vendas.xlsx')

pd.set_option('display.max_columns', None)
print(tabela_vendas)


# faturamento por loja
faturamento = tabela_vendas.groupby('ID Loja')['Valor Final'].sum()
faturamento_df = faturamento.to_frame()
print(faturamento_df)

# quantidade de produtos vendidos por loja
quantidade = tabela_vendas.groupby('ID Loja')['Quantidade'].sum()
quantidade_df = quantidade.to_frame()
print(quantidade_df)

print('-' * 50)
# ticket médio por produto em cada loja
ticket_medio = (faturamento / quantidade).to_frame(name="Ticket Médio")
print(ticket_medio)
