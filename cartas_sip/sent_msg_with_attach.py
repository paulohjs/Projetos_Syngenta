import os
import pandas as pd
import win32com.client as win32

sheet = pd.read_excel(r'C:\Automacoes\Projetos_Syngenta\testes\sip.xlsm',sheet_name='Sheet2')

dataframe = pd.DataFrame(sheet)

for i, row in dataframe.iterrows():
    manager_name = row['manager_name']
    manager_email = row['manager_email']
    print(manager_name)

    # Configuração do caminho local e do arquivo
    local_folder_path = r'C:\Users\s1002322\Syngenta\Brazil HROps - Comp Cycle - Zip_SIP'
    file_to_attach = f"{row['manager_name']}.zip"

    # Iniciar o cliente do Outlook
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    # Configurar o e-mail
    mail.SentOnBehalfOfName = 'syngenta_brasil.operacoes_de_rh@syngenta.com'
    mail.To = manager_email
    mail.Subject = 'Cartas de SIP (Sales Incentive Plan) 2023-2024'
    mail.HTMLBody = f"""<html>
<head></head>
<body>
    <p>Prezado(a) {manager_name}, espero que esteja tudo bem!</p>
     
    <p>Anexo a carta referente ao valor bruto do pagamento do <strong>SIP 2023-2024</strong> que será creditado em <strong>28/03/2024</strong>.</p>
     
    <p>Como os documentos possuem informações confidenciais de remuneração, para acessar as cartas digite o número de sua <strong>matrícula Workday (employee ID)</strong> como senha, ao abrir o arquivo.</p>

    <p><span style="color: red;">Atenção, esta é uma caixa de e-mail não monitorada, por favor não responder este email.</span></p>
    <p>Se você tiver alguma dúvida ou problema na abertura do anexo, entre em contato com a central de RH pelo telefone <strong>(11) 5643 2979</strong> ou abra um chamado pelo seguinte link: 
    <a href="https://syngenta.service-now.com/syn_esc?id=sc_cat_item&sys_id=433cc105dbbd0b005f2dffc41d9619b2" target="_blank">
    Service Desk</a>.</p>

    <p>Atenciosamente,</p>
    <p>Syngenta Recursos Humanos</p>
</body>
</html>
"""

    # Anexar arquivo
    attachment_path = os.path.join(local_folder_path, file_to_attach)
    mail.Attachments.Add(attachment_path)

    # Enviar o e-mail
    mail.Send()