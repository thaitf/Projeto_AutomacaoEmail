import pandas as pd
import win32com.client as win32
import os

# importar tabela
tabela_vendas = pd.read_excel('vendas.xlsx')

# pd.set_option('display.max_columns', None)  
# Exibir todas as colunas para verificar se todas as colunas estao sendo lidas
# print(tabela_vendas.head())  
# Exibir as primeiras linhas da tabela para verificar se os dados foram lidos corretamente   

#faturamento por loja = filtra depois agrupa o que foi filtrado [sum = somando oas outras colunas]
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja', as_index=False).sum()
# as_index=False indica que o indice sera uma coluna e nao o index do pandas
#faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum().reset_index()
#outra forma de fazer o mesmo
faturamento.reset_index(drop=True, inplace=True)
# remove o índice atual e cria um novo de 0 a n-1
faturamento.index += 1
# soma 1 ao índice para começar do 1

#quantidade de produto vendido por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja', as_index=False).sum()
quantidade.reset_index(drop=True, inplace=True)
# remove o índice atual e cria um novo de 0 a n-1
quantidade.index += 1
# soma 1 ao índice para começar do 1

#tiket medio por loja = faturamento dividido pela quantidade de produtos vendidos
ticket = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
#transforma a divisao em uma tabela
ticket = ticket.rename(columns={0: 'Ticket Medio'}) 
#renomeia a coluna 0 para Ticket Medio
ticket['ID Loja'] = quantidade['ID Loja']
ticket = ticket[['ID Loja', 'Ticket Medio']]
#adiciona a coluna 'ID Loja' ao DataFrame ticket
ticket.reset_index(drop=True, inplace=True)
# remove o índice atual e cria um novo de 0 a n-1
ticket.index += 1
# soma 1 ao índice para começar do 1

#junta as tabelas de faturamento, quantidade e ticket médio
final = pd.merge(faturamento, quantidade, on='ID Loja')
final = pd.merge(final, ticket, on='ID Loja')

# salvar Excel com caminho absoluto
caminho = os.path.abspath('relatorio.xlsx')
final.to_excel(caminho, index=False)
#envia excel com os dados finais 

#envia email com o relatório
outlook = win32.Dispatch('Outlook.Application')
mail = outlook.CreateItem(0)
mail.To = 'thaitf@outlook.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados(as),</p>

<p>Segue o relatório de vendas por loja e em anexo o arquivo em excel </p>

<p>Faturamento: </p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}
<!--formatters = formata o valor da coluna 'Valor Final' para o formato de moeda brasileira (R$)
to_html = converte o DataFrame em uma tabela HTML para ser exibida no email -->

<p>Quantidade Vendida: </p>
{quantidade.to_html()}

<p>Ticket Médio:</p> 
{ticket.to_html(formatters={'Ticket Medio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida, estou à disposição.</p>
<p>Agradeço a atenção,</p>
<p>Thaís Tanaka</p>

'''
#''' ''' 3 aspas simples indica que o texto é um bloco de texto,
#  ou seja, pode ter várias linhas
mail.Attachments.Add(Source=caminho)
mail.Send()

print('Email Enviado')
