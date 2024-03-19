import sys
sys.path.append('C:/Automacoes/Projetos_Syngenta')
from utilities.service_now_download_report import dowload_report_sn
from utilities.others import get_email_password
import pandas as pd
import datetime
import requests
import os
from performance_RH.atualizacao_csat import *
from performance_RH.atualizacao_issues import *

#Definir a data para entregar o sample check
def define_limite_date_to_analises_samplecheck():
    data_atual = datetime.date.today()
    dia_semana_atual = data_atual.weekday()
    dias_para_quinta = (3 - dia_semana_atual) % 7
    proxima_quinta = data_atual + datetime.timedelta(days=dias_para_quinta)
    print("A data da próxima quinta-feira é:", proxima_quinta)
    proxima_quinta_str = proxima_quinta.strftime("%Y-%m-%d")
    return proxima_quinta_str

def randon_df(df_excel):
    df_excel = df_excel.sample(frac=1).reset_index(drop=True)
    return df_excel

def treatment_df(df_excel):
    df_excel = df_excel.astype(str)
    df_excel['Closed'] = df_excel['Closed'].str[:10]
    return df_excel

def create_new_df_empty(df_excel):
    novo_df = pd.DataFrame(columns=df_excel.columns)
    return novo_df

def add_rows_sorted_on_new_df(df_excel, novo_df):
    for i, row in df_excel.iterrows():
        assigned_to = row['Assigned to']
        novo_df.loc[i] = row
        contagem = (novo_df['Assigned to'] == assigned_to).sum()
        novo_df.loc[i, 'Count'] = contagem
    novo_df = novo_df.sort_values(by='Count', ascending=True)
    return novo_df

def define_owner_list(team, owners_cg_syngenta, owners_benefits):
    if team == 'CG & Syngenta':
        owners = owners_cg_syngenta
    else:
        owners = owners_benefits
    return owners

def fill_df_whith_owners(novo_df, owners):
    count = 0
    countx5 = 0
    for i, row in novo_df.iterrows():
        novo_df.loc[i, 'Owner'] = owners[countx5]
        count += 1
        if count == 5:
            count = 0
            countx5 += 1
        if countx5 == len(owners):
            break
    novo_df = novo_df[novo_df['Owner'] != 'nan']
    return novo_df

def update_smartsheet_samplecheck(novo_df, webhook_url):
    data_to_send_json = novo_df.to_json(orient='records')
    
    response = requests.post(webhook_url, data= data_to_send_json, headers={"Content-Type": "application/json"})

    if response.status_code == 200:
        print(f"\033[32mChamados enviados para a Smartsheet!\033[m")
    elif response.status_code == 202:
        print(f"\033[32mChamados enviados para a Smartsheet!\033[m")
    else:
        print("\033[31mErro ao enviar dados para o Power Automate:", response.status_code)
        print(response.text)


def main():
    owners_cg_syngenta = ["paulo.silva@syngenta.com"]
    
    "paulo.silva@syngenta.com"
    "arthur.monteiro@syngenta.com"
    "sue.santana@syngenta.com"
    "rafael.biral@syngenta.com"

    owners_benefits = ["natacha.rodrigues@syngenta.com",
                       "camila.severiano@syngenta.com",
                       "juliana.pierro@syngenta.com"]


    link = r'https://syngenta.service-now.com/sys_report_template.do?jvar_report_id=b0216a8e93e1f5d043d4fec47aba10eb'
    arquivo = {'dados':[{'time': 'CG & Syngenta',
                        'sheets_id': 3391140826244996,
                        'http': r'https://prod-158.westeurope.logic.azure.com:443/workflows/b05aa97f931d41ad84e2302842aec057/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=qUtfSLW_720hz5DNPEODVQr8B6C6wzwYitlWHWD6tsQ'},
                        {'time': 'Benefits',
                        'sheets_id': 803030158337924,
                        'http': r'https://prod-228.westeurope.logic.azure.com:443/workflows/f6d782fe28814a958b04b67d7b06664a/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=hhkbsxcg7r8KSOvgrADXr-pDo8VEDsk6eoErjldFPgA'}
                        ]}

    folder_path = os.path.dirname(__file__)
    login_data = get_email_password(0)
    email = login_data[0]
    password = login_data[1]

    file_path = dowload_report_sn(link, folder_path, email, password)

    for dados in arquivo['dados']:

        sheet_id = dados['sheets_id']
        team = dados['time']
        webhook_url = dados['http']

        df_excel = convert_xl_to_dataframe(file_path)

        df_excel = filter_assignmentgroup(df_excel, sheet_id, 803030158337924, 'SYN-BRA-BEN-HR')

        df_excel = treatment_df(df_excel)
        
        df_excel = randon_df(df_excel)
        
        novo_df = create_new_df_empty(df_excel)

        novo_df = add_rows_sorted_on_new_df(df_excel, novo_df)

        proxima_quinta_str = define_limite_date_to_analises_samplecheck()

        novo_df = novo_df.assign(Review_data=proxima_quinta_str)

        owners = define_owner_list(team, owners_cg_syngenta, owners_benefits)

        novo_df = fill_df_whith_owners(novo_df, owners)

        update_smartsheet_samplecheck(novo_df, webhook_url)

if __name__ == "__main__":
    main()
