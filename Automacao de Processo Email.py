#!/usr/bin/env python
# coding: utf-8

# ### Passo 1 - Importar Arquivos e Bibliotecas

# In[16]:


import pandas as pd
from pathlib import Path
import shutil
import win32com.client as win32
import os           

vendas_df = pd.read_excel(r"Bases de Dados\Vendas.xlsx")
email_df = pd.read_excel(r"Bases de Dados\Emails.xlsx")
lojas_df = pd.read_csv(r"Bases de Dados\Lojas.csv", encoding="ISO-8859-1",sep=";")

dic_loja = {}
vendas_df = vendas_df.merge(lojas_df, on="ID Loja")
for loja in lojas_df['Loja']:
    dic_loja[loja] = vendas_df.loc[vendas_df["Loja"] == loja, :]
 
dia_indicador = vendas_df['Data'].max()

caminho_backup_lojas = Path(r"Backup Arquivos Lojas")

arquivos_pasta_backup = caminho_backup_lojas.iterdir()
lista_lojas_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dic_loja:
    if loja not in lista_lojas_backup:
        nova_pasta = caminho_backup_lojas / loja
        nova_pasta.mkdir()
        
    nome_arquivo = "{}_{}_{}.xlsx".format(dia_indicador.month, dia_indicador.day, loja)
    local_arquivo = caminho_backup_lojas / loja / nome_arquivo
    dic_loja[loja].to_excel(local_arquivo)
    
meta_dia_fat = 1000
meta_dia_produto = 4
meta_dia_ticket = 600

meta_ano_fat = 1650000
meta_ano_produto = 120
meta_ano_ticket = 500

for loja in dic_loja:
    venda_loja = dic_loja[loja]
    venda_loja_dia = venda_loja.loc[venda_loja["Data"]==dia_indicador, :]

    #fat
    total_dia = venda_loja_dia['Valor Final'].sum()
    total_ano = venda_loja['Valor Final'].sum()
#     print(total_dia, total_ano, sep="\n")


    #diversidade de produtos
    produtos_ano = len(venda_loja['Produto'].unique())
#     print(produtos_ano)
    produtos_dia = len(venda_loja_dia['Produto'].unique())
#     print(produtos_dia)

    #ticket médio
    valor_venda = venda_loja.groupby('Código Venda').sum(numeric_only=True)
    ticket_ano = valor_venda['Valor Final'].mean()
#     print(f'{ticket_ano:.2f}')

    valor_venda_dia = venda_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_dia = valor_venda_dia['Valor Final'].mean()
#     print(f"{ticket_dia:.2f}")
#     print('-----')
    
    outlook = win32.Dispatch("outlook.application")


    email_gerente = email_df.loc[email_df["Loja"]==loja, "E-mail"].values[0]
    nome_gerente = email_df.loc[email_df["Loja"]==loja, "Gerente"].values[0]

    mail = outlook.CreateItem(0)
    mail.To = email_gerente
    mail.Subject = f"OnePage dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}"

    if total_dia >= meta_dia_fat:
        cor_fat_dia = "green"
    else:
        cor_fat_dia = "red"

    if produtos_dia >= meta_dia_produto:
        cor_prod_dia = "green"
    else:
        cor_prod_dia = "red"

    if ticket_dia >= meta_dia_ticket:
        cor_ticket_dia = "green"
    else:
        cor_ticket_dia = "red"
    #
    #
    if total_ano >= meta_ano_fat:
        cor_fat_ano = "green"
    else:
        cor_fat_ano = "red"

    if produtos_ano >= meta_ano_produto:
        cor_prod_ano = "green"
    else:
        cor_prod_ano = "red"

    if ticket_ano >= meta_ano_ticket:
        cor_ticket_ano = "green"
    else:
        cor_ticket_ano = "red"


    mail.HTMLBody = f"""
    <p>Bom dia, {nome_gerente}!</p>
    <p>O resultado de ontem <strong>{dia_indicador.day}/{dia_indicador.month}</strong> da loja <strong>{loja}</strong> foi:</p>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor dia</th>
        <th>Meta dia</th>
        <th>Cenário dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R$ {total_dia:.2f}</td>
        <td style="text-align: center">R$ {meta_dia_fat:.2f}</td>
        <td style="text-align: center"><font color={cor_fat_dia}>◙</td>
      </tr>
      <tr>
        <td>Diversidade Prod</td>
        <td style="text-align: center">{produtos_dia}</td>
        <td style="text-align: center">{meta_dia_produto}</td>
        <td style="text-align: center"><font color={cor_prod_dia}>◙</td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R$ {ticket_dia:.2f}</td>
        <td style="text-align: center">R$ {meta_dia_ticket:.2f}</td>
        <td style="text-align: center"><font color={cor_ticket_dia}>◙</td>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor dia</th>
        <th>Meta dia</th>
        <th>Cenário dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R$ {total_ano:.2f}</td>
        <td style="text-align: center">R$ {meta_ano_fat:.2f}</td>
        <td style="text-align: center"><font color={cor_fat_ano}>◙</td>
      </tr>
      <tr>
        <td>Diversidade Prod</td>
        <td style="text-align: center">{produtos_ano}</td>
        <td style="text-align: center">{meta_ano_produto}</td>
        <td style="text-align: center"><font color={cor_prod_ano}>◙</td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R$ {ticket_ano:.2f}</td>
        <td style="text-align: center">R$ {meta_ano_ticket:.2f}</td>
        <td style="text-align: center"><font color={cor_ticket_ano}>◙</td>
      </tr>
    </table>

    <p>Segue em anexo planilha com todos os dados para mais detalhes.</p>

    <p>Qualquer dúvida estou a disposição.</p>
    <p>Att., Lucas</p>
    """
    caminho = Path.cwd() / caminho_backup_lojas / loja / f"{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx"
    attachment = caminho
    mail.Attachments.Add(str(attachment))
    mail.Send()
    print(f"email da loja {loja} enviado")
    
faturamento_lojas = vendas_df.groupby('Loja')[['Loja','Valor Final']].sum(numeric_only=True)
ranking_fat_lojas = faturamento_lojas.sort_values(by="Valor Final", ascending=False)

nome_arquivo = "{}_{}_Ranking_Anual.xlsx".format(dia_indicador.month, dia_indicador.day)
ranking_fat_lojas.to_excel(r"Backup Arquivos Lojas/{}".format(nome_arquivo))

##

faturamento_lojas_dia = vendas_df.loc[vendas_df['Data']==dia_indicador, :]
faturamento_lojas_dia = faturamento_lojas_dia.groupby('Loja')[['Loja','Valor Final']].sum(numeric_only=True)
ranking_fat_lojas_dia = faturamento_lojas_dia.sort_values(by="Valor Final", ascending=False)

nome_arquivo = "{}_{}_Ranking_Dias.xlsx".format(dia_indicador.month, dia_indicador.day)
ranking_fat_lojas_dia.to_excel(r"Backup Arquivos Lojas/{}".format(nome_arquivo))

outlook = win32.Dispatch("outlook.application")


email_gerente = email_df.loc[email_df["Loja"]=="Diretoria", "E-mail"].values[0]


mail = outlook.CreateItem(0)
mail.To = email_gerente
mail.Subject = f"{dia_indicador.day}/{dia_indicador.month}_Ranking de Lojas"
mail.Body = f"""
Prezados, bom dia

Melhor loja do dia em faturamento {ranking_fat_lojas_dia.index[0]} com faturamento de R${ranking_fat_lojas_dia.iloc[0, 0]:.2f}.
Pior loja do dia em faturamento {ranking_fat_lojas_dia.index[-1]} com faturamento de R${ranking_fat_lojas_dia.iloc[-1, 0]:.2f}.

Melhor loja do Ano em faturamento {ranking_fat_lojas.index[0]} com faturamento de R${ranking_fat_lojas.iloc[0, 0]:.2f}.
Pior loja do Ano em faturamento {ranking_fat_lojas.index[-1]} com faturamento de R${ranking_fat_lojas.iloc[-1, 0]:.2f}.

Segue anexo do ranking das lojas de dia e ano

Qualquer dúvida estou a disposição.

Att., 

Lucas
"""

attachment = Path.cwd() / caminho_backup_lojas / f"{dia_indicador.month}_{dia_indicador.day}_Ranking_Anual.xlsx"
mail.Attachments.Add(str(attachment))
attachment = Path.cwd() / caminho_backup_lojas / f"{dia_indicador.month}_{dia_indicador.day}_Ranking_Dias.xlsx"
mail.Attachments.Add(str(attachment))

mail.Send()

print(f"email da Diretoria enviado")


# In[ ]:




