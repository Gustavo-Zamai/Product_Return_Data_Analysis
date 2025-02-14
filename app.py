import os
import pandas as pd
import plotly.express as px
import win32com.client as win32

data_list = os.listdir("Vendas")
#print(data_list)

joined_table = pd.DataFrame()

for file in data_list:
    if "Devolucoes" in file:
        table = pd.read_csv(f"Vendas/{file}")
        joined_table = joined_table._append(table)
        #print(joined_table)

# Produto mais devolvida
product_table = joined_table.groupby("Produto").sum()
product_table = product_table[["Quantidade Devolvida"]].sort_values(by="Quantidade Devolvida", ascending=False)
print(product_table)

graphic = px.bar(product_table, x=product_table.index, y="Quantidade Devolvida")
graphic.show()

print("-" * 50)
# Loja com mais devoluções
store_table = joined_table.groupby("Loja").sum()
store_table = store_table[["Quantidade Devolvida"]].sort_values(by="Quantidade Devolvida", ascending=False)
print(store_table)

bar_graphic = px.bar(store_table, x=store_table.index, y="Quantidade Devolvida")
bar_graphic.show()

# send email as a report
outlook = win32.Dispatch("outlook.application")
mail = outlook.CreateItem(0)
mail.To = "gustavosimaozamai@gmail.com"
mail.Subject = "Relatório de Devoluções"
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o relatório de vendas por cada loja.</p>

<p>Produtos com maior número de devoluções:</p>
{product_table.to_html()}

<p>Quantidade Vendida:</p>
{store_table.to_html()}

<p>Qualquer dúvida estou à disposição.</p>
<p>Att.,</p>
<p>Gustavo.</p>
'''

mail.Send()