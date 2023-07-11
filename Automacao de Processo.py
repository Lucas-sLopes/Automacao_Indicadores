#!/usr/bin/env python
# coding: utf-8

# ### Passo 1 - Importar Arquivos e Bibliotecas

# In[1]:


# Bibliotecas
import pandas as pd
import pathlib
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')


# In[2]:


# Arquivos
df_email = pd.read_excel(r'Bases de Dados/Emails.xlsx')
df_lojas = pd.read_csv(r'Bases de Dados/Lojas.csv', encoding='latin1', sep=';')
df_vendas = pd.read_excel(r'Bases de Dados/Vendas.xlsx')


# ### Passo 2 - Definir Criar uma Tabela para cada Loja e Definir o dia do Indicador

# In[3]:


#Mesclar tabelas para conter as lojas na tabela de venda
vendas = df_vendas.merge(df_lojas, on='ID Loja')


# In[4]:


#Cria um dicionario com uma tabela de cada loja
dic_lojas = {}
for loja in df_lojas['Loja']:
    dic_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :]


# In[7]:


#Definir dia do indicador
dia_indicador = vendas['Data'].max()


# ### Passo 3 - Salvar a planilha na pasta de backup

# In[8]:


caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

arquivos_backup = caminho_backup.iterdir()

#Lista para verificar se o arquivo da loja existe dentro da lista
lista_backup = [arquivo.name for arquivo in arquivos_backup]

for loja in dic_lojas:
    if loja not in lista_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()
    nome_arquivo = '{}-{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.year, loja)
    local_arquivo = caminho_backup / loja / nome_arquivo
    
    dic_lojas[loja].to_excel(local_arquivo)


# ### Passo 4 - Calcular o indicador para 1 loja

# In[29]:


loja = 'Norte Shopping'
loja_vendas = dic_lojas[loja]
vendas_loja_dia = loja_vendas.loc[loja_vendas['Data'] == dia_indicador, :]

#Faturamento
faturamento_ano = loja_vendas['Valor Final'].sum()
faturamento_dia = vendas_loja_dia['Valor Final'].sum()
print(f'O faturamento do ano da loja {loja} foi de R$ {faturamento_ano:.2f}')
print(f'O faturamento do dia da loja {loja} foi de R$ {faturamento_dia:.2f}')

#Diversidade de produtos
qtd_produtos_ano = len(loja_vendas['Produto'].unique())
print(f'A loja {qtd_produtos_ano} vendeu produtos diferentes no ano')

qtd_produtos_dia = len(vendas_loja_dia['Produto'].unique())
print(f'Foram vendidos hoje {qtd_produtos_dia} produtos diferentes')

#ticket medio anual
valor_venda = loja_vendas.groupby('Código Venda').sum()
ticket_medio_ano = valor_venda['Valor Final'].mean()
print(f'A média de vendas anual da loja {loja} foi de R${ticket_medio_ano:.2f}')

#ticket medio Dia
valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
ticket_media_dia = valor_venda_dia['Valor Final'].mean()
print(f'A média de vendas da loja {loja} no dia foi de R${ticket_media_dia:.2f}')


# ### Passo 5 - Enviar por e-mail para o gerente

# In[ ]:





# ### Passo 6 - Automatizar todas as lojas

# ### Passo 7 - Criar ranking para diretoria

# In[ ]:





# ### Passo 8 - Enviar e-mail para diretoria

# In[ ]:




