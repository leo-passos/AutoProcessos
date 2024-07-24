#!/usr/bin/env python
# coding: utf-8

# ### Importar Arquivos e Bibliotecas

# In[1]:


import pandas as pd
import pathlib
import win32com.client as win32


# In[2]:


#importar as bases de dados
#corrigir o probema de Encoding e Separadores do .csv

emails = pd.read_excel(r'Bases de Dados/Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados/Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados/Vendas.xlsx')

display(emails.head(3))
display(lojas.head(3))
display(vendas.head(3))


# ### Definir Criar uma Tabela para cada Loja e Definir o dia do Indicador

# In[3]:


#Incluir o nome da loja em Vendas
vendas = vendas.merge(lojas, on='ID Loja')


# In[4]:


#Criar uma tabela para cada uma das lojas, para análisar as vendas separadamentes
dic_lojas = {}
for loja in lojas['Loja']:
    dic_lojas[loja] = vendas.loc[vendas['Loja']==loja,:]


# In[5]:


#Formatação das Datas e Horas das Tabelas
dia_indicador = vendas['Data'].max()
print('{}/{}'.format(dia_indicador.day, dia_indicador.month))


# ### Salvar a planilha na pasta de backup

# In[6]:


#identificar se a pasta já existe
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

arquivo_pasta_backup = caminho_backup.iterdir()

lista_nomes_backup = []
for arquivo in arquivo_pasta_backup:
    lista_nomes_backup.append(arquivo.name)

for loja in dic_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja #função do Pathlib 
        nova_pasta.mkdir()


# In[7]:


#salvar dentro da pasta
nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month,dia_indicador.day,loja)
local_arquivo = caminho_backup / loja / nome_arquivo
    
dic_lojas[loja].to_excel(local_arquivo)


# ### Calcular o indicador para 1 loja
# Indicadores:
# - Faturamento
# - Diversidade de Produtos
# - Ticket Médio
# 
# ### Enviar por e-mail para o gerente
# Com o relatório de cada loja

# In[8]:


#Definição de Metas 
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtde_produto_dia = 4
meta_qtde_produto_ano = 120
meta_ticket_medio_dia = 500
meta_ticket_medio_ano = 500


# In[9]:


for loja in dic_lojas:
    
    vendas_loja = dic_lojas[loja]
    vendas_loja_dia =  vendas_loja.loc[vendas_loja['Data']==dia_indicador , :]

    #Faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    #Diversidade de Produtos
    qtde_produto_ano = len(vendas_loja['Produto'].unique())
    qtde_produto_dia = len(vendas_loja_dia['Produto'].unique())

    #Ticket Médio
    valor_venda = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_ano = valor_venda['Valor Final'].mean()

    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
    
    outlok = win32.Dispatch('outlook.application')

    nome = emails.loc[emails['Loja']==loja,'Gerente'].values[0]

    mail = outlok.CreateItem(0)
    mail.To = emails.loc[emails['Loja']==loja,'E-mail'].values[0]
    mail.CC = ''
    mail.BCC = ''
    mail.Subject = f'OnePage Dia{dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
    #mail.Body = 'Texto do Email'
    #ou mail.HTMLBody = <p>Corpo do Email</p>

    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produto_dia >= meta_qtde_produto_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produto_ano >= meta_qtde_produto_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticket_medio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticket_medio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    mail.HTMLBody = f'''
        <p>Bom dia, {nome}</p>

    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produto_dia}</td>
        <td style="text-align: center">{meta_qtde_produto_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticket_medio_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produto_ano}</td>
        <td style="text-align: center">{qtde_produto_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticket_medio_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
      </tr>
    </table>

    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

    <p>Qualquer dúvida estou à disposição.</p>
    <p>Att., Lira</p>
    
    '''
    

    #Anexos (pode colocar quantos quiser)
    attachment  = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))

    mail.Send()
    print('E-mail da Loja {} enviado para o(a) gerente {}'.format(loja, nome))


# ### Automatizar todas as lojas

# In[10]:


display(vendas)


# In[11]:


faturamento_lojas = vendas.groupby('Loja')[['Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)
display(faturamento_lojas_ano)

nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

vendas_dia = vendas.loc[vendas['Data']==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)
display(faturamento_lojas_dia)

nome_arquivo = '{}_{}_Ranking Dia.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))


# ### Criar ranking e enviar e-mail para a Diretoria
# - Melhor loja do Dia em Faturamento
# - Pior loja do Dia em Faturamento
# - Melhor loja do Ano em Faturamento
# - Pior loja do Ano em Faturamento

# In[12]:


#enviar o e-mail
outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']=='Diretoria', 'E-mail'].values[0]
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'
mail.Body = f'''

Prezados, bom dia

Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}
Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, 0]:.2f}

Melhor loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}
Pior loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1, 0]:.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição.

Att.,
Lira
'''

# Anexos (pode colocar quantos quiser):
attachment  = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment  = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))


mail.Send()
print('E-mail da Diretoria enviado')

