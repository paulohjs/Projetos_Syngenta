import smartsheet
import pandas as pd
import requests
from performance_RH.atualizacao_csat import *
from performance_RH.atualizacao_issues import *

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


def update_pdf_title(pdf_path= str, new_title= str):
    from PyPDF2 import PdfReader, PdfWriter

    # Abrir o PDF existente
    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)
        writer = PdfWriter()

        # Copiar todas as páginas do PDF original para o novo PDF
        for page in reader.pages:
            writer.add_page(page)

        # Modificar os metadados
        metadata = reader.metadata
        new_metadata = {**metadata, **{'/Title': f'{new_title}]'}}

        writer.add_metadata(new_metadata)

        # Salvar o novo PDF
        with open(pdf_path, 'wb') as new_file:
            writer.write(new_file)