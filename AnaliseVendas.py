import win32com.client as win32
import pandas as pd

# importar a base de dados
Tabela_Vendas = pd.read_excel('Vendas.xlsx')


# visualizar a base de dados
# impedir que o pyton oculte tabelas
pd.set_option('display.max_columns', None)
# print(Tabela_Vendas)

# Faturamento total por loja
Faturamento_Por_Loja = Tabela_Vendas[[
    'ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(Faturamento_Por_Loja)
print('_' * 50)

# Quantidade de produtos vendidos por loja
Qty_Produtos_Vendidos = Tabela_Vendas[[
    'ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(Qty_Produtos_Vendidos)
print('_' * 50)

# Ticket médio por produto em cada loja
Ticket_Medio_Por_Produto = (
    Faturamento_Por_Loja['Valor Final'] / Qty_Produtos_Vendidos['Quantidade']).to_frame()
print(Ticket_Medio_Por_Produto)
Ticket_Medio_Por_Produto = Ticket_Medio_Por_Produto.rename(
    columns={0: 'Ticket Médio'})  # renomeando o nome da coluna

# enviar um email com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'hiagosa9@gmail.com'
mail.subject = 'Relatório De Vendas Por loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada loja</p>

<p>Faturamento:
</p>
{Faturamento_Por_Loja.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}

<p>Quantidade Vendida
<p>
{Qty_Produtos_Vendidos.to_html()}

<p>Ticket Médio dos porodutos em cada loja:
</p>
{Ticket_Medio_Por_Produto.to_html(formatters={'Ticket Médio':'R${:,.2f}'.format})}
<p>Qualquer dúvida estou à disposição</p>
<p>Att.,</p>
<p>Hiago Sátiro</p>'''
mail.Send()
print('Email Enviado!')
