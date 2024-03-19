import sys
sys.path.append('C:/Automacoes/Projetos_Syngenta')
from utilities.service_now_download_report import dowload_report_sn
from utilities.others import get_email_password, get_smartsheet_key
import pandas as pd
import os
from performance_RH.atualizacao_csat import *


def __convert_type_of_data__(df_excel):
    df_excel['Employee number'].fillna(0, inplace=True)
    df_excel = df_excel.astype(str)
    df_excel['Number of impacted/ potentially impacted employees'] = df_excel['Number of impacted/ potentially impacted employees'].astype(int)
    df_excel['Employee number'] = df_excel['Employee number'].astype(float).apply(round)
    df_excel['Employee number'] = df_excel['Employee number'].astype(int)
    
    return df_excel

def __convert_datetime_to_dateonly__(df_excel):
    df_excel['Opened'] = df_excel['Opened'].astype(str)
    df_excel['Opened'] = df_excel['Opened'].str[:10]
    df_excel['Closed'] = df_excel['Closed'].astype(str)
    df_excel['Closed'] = df_excel['Closed'].str[:10]

    return df_excel

def filter_assignmentgroup(df_excel, sheet_id_actual, sheet_id_filter_if_true, assignment_group):
    if sheet_id_actual == sheet_id_filter_if_true:
        df_excel = df_excel.loc[df_excel['Assignment group'] == assignment_group]
    else:
        df_excel = df_excel.loc[df_excel['Assignment group'] != assignment_group]
    return df_excel

def main():
    smartsheet_token = get_smartsheet_key()
    link = r'https://syngenta.service-now.com/sys_report_template.do?jvar_report_id=21c7b3fb47cc6950cbd1efb2e36d43ab'
    reports = {'arquivos': [
        {'http': r'https://prod-145.westeurope.logic.azure.com:443/workflows/07eb6b784ef747eeabc434deec07e1aa/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=XaugGq7jA6RMmoss9S1bx8k_dJ3_3a_XZYRXoG_bbhg',
        'sheet_id': 6399567966037892,
        'sheet_name': 'Analise de Issues (Beneficios)',
        'column_id': 4597583127046020},
        {'http': r'https://prod-98.westeurope.logic.azure.com:443/workflows/e33a1d84ad8e4effb7fa8effcf276e66/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=9NNPRBrSf51BWFgIqXM6NPJ1sDYh86fbiOPYQ22zBYU',
        'sheet_id': 7475485111281540,
        'sheet_name': 'Analise de Issues (CG + Syngenta)',
        'column_id': 4902497693263748}]}

    folder_path = os.path.dirname(__file__)
    login_data = get_email_password(0)
    email = login_data[0]
    password = login_data[1]

    file_path = dowload_report_sn(link, folder_path, email, password)

    for arquivo in reports['arquivos']:

        webhook_url = arquivo['http']
        sheet_id = arquivo['sheet_id']
        sheet_name = arquivo['sheet_name']

        df_excel = convert_xl_to_dataframe(file_path)

        df_excel = __convert_type_of_data__(df_excel)

        df_excel = __convert_datetime_to_dateonly__(df_excel)

        df_excel = filter_assignmentgroup(df_excel, sheet_id, 6399567966037892, 'SYN-BRA-BEN-HR')

        sheet = query_smartsheet_data(smartsheet_token, sheet_id)

        df_smartsheet = convert_smartsheetdata_to_dataframe(sheet)

        not_contained = verify_if_tickets_not_contais_on_smartsheet(df_excel, df_smartsheet)

        update_smartsheet(not_contained, webhook_url, sheet_name)

    cleanup_folder(file_path)

if __name__ == "__main__":
    main()
