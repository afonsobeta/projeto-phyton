import pandas as pd
import win32com.client as win32
from pandas.core.common import not_none
# importar a base de dados
tabela_vendas= pd.read_excel("vendas.xlsx")

# visualizar a base de dados
pd.set_option("display.max_columns", None)
print(tabela_vendas)

# faturamento por loja
faturamento = tabela_vendas[["ID Loja", "Valor Final"]].groupby ("ID Loja").sum()
print (faturamento)

quantidade = tabela_vendas[["ID Loja", "Quantidade"]].groupby ("ID Loja").sum()
# quantidade de produtos vendidos por loja
print (quantidade)

print ("-" * 50)
# ticket medio por produto em cada loja
ticket_medio = (faturamento["Valor Final"] / quantidade["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
print (ticket_medio)

# enviar um email com relatorio
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'email@exemplo.com'
mail.Subject = 'Relatorios de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatorio de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Medio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Medio': 'R${:,.2f}'.format})}

<p>Qualquer duvida estou a disposicao.</p>

<p>Att.,</p>
<p>Afonso</p>
'''

mail.Send()

print ('Email Enviado')