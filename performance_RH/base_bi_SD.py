import sys
sys.path.append('C:/Automacoes/Projetos_Syngenta')
from utilities.service_now_download_report import dowload_report_sn
from utilities.others import get_email_password
import requests
import openpyxl
from time import sleep

#Define the URLs and folder_paths to use on the function
reports = {'files': [
    {'link': 'https://syngenta.service-now.com/sys_report_template.do?jvar_report_id=f745ace79756815051d0359fe153af12', 
     'file_path': 'C:\Power BI\Service Now - Customer Satisfaction'},
    {'link': 'https://syngenta.service-now.com/sys_report_template.do?jvar_report_id=798e78d397528d1051d0359fe153afd5',
     'file_path': 'C:\Power BI\Service Now - Cases Opened'},
    {'link': 'https://syngenta.service-now.com/sys_report_template.do?jvar_report_id=2ed315f4c34629d0883212a4e401315b',
     'file_path': 'C:\Power BI\Service Now - Cases Closed'}
]}

#Get email and password from JSON
login_data = get_email_password(0)
email = login_data[0]
password = login_data[1]

#Dowload the reports with the function sndowload
for file in reports['files']:
    file_path = dowload_report_sn(file['link'], file['file_path'], email, password)
    workbook = openpyxl.load_workbook(file_path)
    workbook.save(file_path)
    sleep(3)
    workbook.close()

#request to update the dataset with power automate
url = "https://prod-63.westeurope.logic.azure.com:443/workflows/9132d3907c714553ad0a36d34c7204f1/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=tA13l83-grPO22OWbRKiLmTIK-hrz93Zyeieww-f4cg"
data = {"function": "run_flow"}
response = requests.post(url, json=data)
if response.status_code == 200:
    print("\033[32mThe flow in power automate requested to update dataset\033[m")
elif response.status_code == 202:
    print("\033[32mThe flow in power automate requested to update dataset\033[m")
else:
    print("\033[31mAn error occurred while updating the dataset.")
    print("Error code:", response.status_code, "\033[m")
