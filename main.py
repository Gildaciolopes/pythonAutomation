import pandas as pd
import win32com.client as win32

# Ler a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
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

# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'contato.gildaciolopes@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento_df.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade_df.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Gildácio</p>
'''

mail.Send()

print('Email Enviado')