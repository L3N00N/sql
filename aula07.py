import pandas as pd
import win32com.client as win32
tabela_vendas = pd.read_excel('Vendas.xlsx')

#mostrar somente tabelas escolhidas
faturamento= tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
#iprint(faturamento)

#quantidade produtos vendidos por loja
quantidade= tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
#print(quantd)

#ticket medio por prduto de cada loja
ticket_medio= (faturamento['Valor Final']/ quantidade['Quantidade']).to_frame()
ticket_medio= ticket_medio.rename(columns={0: 'Ticket Medio'})
print(ticket_medio)

#enviar email
outlook= win32.Dispatch('outlook.application')
mail= outlook.CreateItem(0)
mail.To= 'leonardo123preuss@gmail.com'
mail.Subject= 'Relatorio de vendas por loja'
mail.HTMLBody= f''''
<p>Prezados,
Segue o relatorio de vendas por cada loja. </p>

<p>faturamento</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>quantidade vendida</p>
{quantidade.to_html(formatters={'Quantidade': 'R${:,.2f}'.format})}

<p>ticket medio dos produtos de cada loja</p>
{ticket_medio.to_html(formatters={'Ticket Medio': 'R${:,.2f}'.format})}

<p>Qualquer duvida estou a disposição</p>

<p>Att...
Preuss</p>
'''

mail.Send()