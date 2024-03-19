import sys
sys.path.append('C:/Automacoes/Projetos_Syngenta')
from utilities.service_now_download_report import dowload_report_sn
from utilities.others import get_email_password
import os
import pandas as pd
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.lists.list import List
from office365.sharepoint.listitems.listitem import ListItem
import requests


folder_path = os.path.dirname(__file__)
login_data = get_email_password(0)
email = login_data[0]
password = login_data[1]
link = r'https://syngenta.service-now.com/sys_report_template.do?jvar_report_id=52db39a01bdef05094ddc9506e4bcbc1'

file_path = dowload_report_sn(link, folder_path, email, password)

df = pd.read_excel(file_path)
df = df.sort_values('Opened for')
df['Concatenate'] = '<br><b><a href=https://syngenta.service-now.com/assessment_take2.do?sysparm_assessable_sysid=' + df['Sys ID'] + '&sysparm_assessable_type=d8844bbedb741340fa1570a5ae9619ec>' + df['Number'] + ' - ' + df['Short description'] + '</a></b>'
df = df.reset_index(drop=True)

df1 = pd.DataFrame({'email': [], 'textjoin': [], '1nome': []})

nome = str
textjoin = list
indice_output = -1
for i in range(0, len(df)):

    if i == 0:
        nova_linha = pd.DataFrame({'email': [df['Email'][i]], 'textjoin': [df['Concatenate'][i]], '1nome': [df['Opened for'][i]]})
        df1 = pd.concat([df1, nova_linha], ignore_index=True)
        indice_output = indice_output+1

    elif df['Opened for'][i] == df['Opened for'][i-1]:

        df1['textjoin'][indice_output] = df1['textjoin'][indice_output] + ';' + df['Concatenate'][i]
    
    else:
        nova_linha = pd.DataFrame({'email': [df['Email'][i]], 'textjoin': [df['Concatenate'][i]], '1nome': [df['Opened for'][i]]})
        df1 = pd.concat([df1, nova_linha], ignore_index=True)
        indice_output = indice_output+1

df1['1nome'] = df1['1nome'].str.split().str[0]
df1 = df1.astype(str)


site_url = 'https://syngenta.sharepoint.com/sites/Testes314'
login_data = get_email_password(1)
generic_username = login_data[0]
generic_password = login_data[1]

ctx_auth = AuthenticationContext(site_url)

if ctx_auth.acquire_token_for_user(generic_username, generic_password):
    ctx = ClientContext(site_url, ctx_auth)
    print('Usuário autenticado com sucesso')
else:
    print('Falha na autenticação do usuário')
    exit()

list_name = 'lembrete_pesquisa'
sp_list = ctx.web.lists.get_by_title(list_name)
for i, row in df1.iterrows():
    item_properties = {'Title': row['email'], 'textjoin': row['textjoin'], 'nome': row['1nome']}
    new_item = sp_list.add_item(item_properties)
    ctx.execute_query()

print('Itens adicionados com sucesso!')
