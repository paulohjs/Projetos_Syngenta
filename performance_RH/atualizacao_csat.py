import sys
sys.path.append('C:/Automacoes/Projetos_Syngenta')
from utilities.service_now_download_report import dowload_report_sn
from utilities.others import get_email_password, get_smartsheet_key
import smartsheet
import pandas as pd
import requests
import os


def convert_xl_to_dataframe(file_path):
    xl_sheet = pd.read_excel(file_path)
    df_excel = pd.DataFrame(xl_sheet)
    return df_excel

def __convert_type_of_data__(df_excel, sheet_id):
    df_excel = df_excel.astype(str)
    if sheet_id == 7343913095718788:
        df_excel['Value'] = df_excel['Value'].astype(int)
    return df_excel

def __convert_datetime_to_dateonly__(df_excel, sheet_id):
    df_excel['Opened'] = df_excel['Opened'].astype(str)
    df_excel['Opened'] = df_excel['Opened'].str[:10]
    df_excel['Closed'] = df_excel['Closed'].astype(str)
    df_excel['Closed'] = df_excel['Closed'].str[:10]
    if sheet_id == 7343913095718788:
        df_excel['Updated'] = df_excel['Updated'].astype(str)
        df_excel['Updated'] = df_excel['Updated'].str[:10]
    return df_excel

def query_smartsheet_data(smartsheet_token, sheet_id):
    smartsheet_client = smartsheet.Smartsheet(smartsheet_token)
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)
    return sheet

def convert_smartsheetdata_to_dataframe(sheet):
    data = []
    for row in sheet.rows:
        row_data = [cell.value for cell in row.cells]
        data.append(row_data)
    column_titles = [column.title for column in sheet.columns]
    df_smartsheet = pd.DataFrame(data, columns=column_titles)
    return df_smartsheet

def verify_if_tickets_not_contais_on_smartsheet(df_excel, df_smartsheet):
    is_contained = df_excel['Number'].isin(df_smartsheet['Ticket_Case'])
    not_contained = df_excel[~is_contained]
    return not_contained

def update_smartsheet(not_contained, webhook_url, sheet_name):
    count_not_countained = not_contained['Number'].count()

    if count_not_countained > 0:
        print(f'\nForam encontrados {count_not_countained} chamados novos, dentre eles:')
        print(not_contained['Number'].values, f'\n')

        data_to_send_json = not_contained.to_json(orient='records', indent=4)

        response = requests.post(webhook_url, data= data_to_send_json, headers={"Content-Type": "application/json"})

        if response.status_code == 200:
            print(f"\033[32mChamados enviados para a Smartsheet {sheet_name}!\033[m")
        elif response.status_code == 202:
            print(f"\033[32mChamados enviados para a Smartsheet {sheet_name}!\033[m")
        else:
            print("\033[31mErro ao enviar dados para o Power Automate:", response.status_code)
            print(response.text)
    else:
        print(f'\033[31mNão foram encontrados chamados novos para a Smartsheet {sheet_name}!!\033[m')

def cleanup_folder(file_path):
    if os.path.exists(file_path):
        os.remove(file_path)

def main():
    smartsheet_token = get_smartsheet_key()

    reports = {'arquivos': [
        {'link': r'https://syngenta.service-now.com/sys_report_template.do?jvar_report_id=fb474a3f4708219051cf6d72e36d4353',
        'http': r'https://prod-124.westeurope.logic.azure.com:443/workflows/fd630c10f08d47c480109b38f98b6c4b/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=monYox6dkYikBM1_xbvV4niMIeHEntwK3XCZFMFYVGg',
        'sheet_id': 7343913095718788,
        'sheet_name': 'Analise Cases CSAT'}]}

    folder_path = os.path.dirname(__file__)
    
    login_data = get_email_password(0)
    email = login_data[0]
    password = login_data[1]

    for arquivo in reports['arquivos']:

        link = arquivo['link']
        webhook_url = arquivo['http']
        sheet_id = arquivo['sheet_id']
        sheet_name = arquivo['sheet_name']

        file_path = dowload_report_sn(link, folder_path, email, password)

        df_excel = convert_xl_to_dataframe(file_path)

        df_excel = __convert_type_of_data__(df_excel, sheet_id)
        
        df_excel = __convert_datetime_to_dateonly__(df_excel, sheet_id)

        sheet = query_smartsheet_data(smartsheet_token, sheet_id)

        df_smartsheet = convert_smartsheetdata_to_dataframe(sheet)

        not_contained = verify_if_tickets_not_contais_on_smartsheet(df_excel, df_smartsheet)

        update_smartsheet(not_contained, webhook_url, sheet_name)

        cleanup_folder(file_path)

if __name__ == "__main__":
    main()
