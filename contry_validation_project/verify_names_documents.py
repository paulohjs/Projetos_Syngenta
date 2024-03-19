import sys
sys.path.append('C:/Automacoes/Projetos_Syngenta')
import utilities.browser_functions as browser
from utilities.others import get_email_password
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from time import sleep
import sqlite3

def document_informations(driver, check_exist_column,i ,index, dn, dtw):

    index = 1
    
    while True:
        if 'Processo de negócios' in check_exist_column:
            try:
                processo = driver.find_element(By.XPATH, f'//tbody/tr[{index}]/td[3]').text
            except NoSuchElementException:
                print("Elemento não encontrado, encerrando o loop.")
                break

            if processo == "":
                try:
                    document_name = driver.find_element(By.XPATH, f'//tbody/tr[{index}]/td[{dn}]').text
                    document_type_workday = driver.find_element(By.XPATH, f'//tbody/tr[{index}]/td[{dtw}]').text
                    document_folder_workday = driver.find_element(By.XPATH, f'//tbody/tr[{index}]/td[2]').text
                    document_id = document_name[28:36]
                    document_id = int(document_id)
                    document_date = document_name[17:27]

                    print(f'''employee_id: {i}
                    Nome do documento: {document_name}
                    ID do documento: {document_id}
                    Categoria do documento: {document_type_workday}
                    Pasta do documento: {document_folder_workday}
                    Data do documento: {document_date}
                    ''')
                    
                    query = '''
                    INSERT INTO documents_workday (document_id_wd, employee_id_wd, document_name, document_folder, document_type, document_date)
                    VALUES (?, ?, ?, ?, ?, ?)
                    '''
                    connector = sqlite3.connect('bd_nomes.sdb')
                    cursor = connector.cursor()
                    cursor.execute(query, (document_id, i, document_name,  document_folder_workday, document_type_workday, document_date))
                    connector.commit()
                    connector.close()

                except sqlite3.Error as e:
                    print(f"Ocorreu um erro ao inserir os dados: {e}")
                    connector.close()
                except NoSuchElementException:
                    print("Elemento não encontrado, encerrando o loop.")
                    break
                except Exception as e:
                    print(f"Ocorreu um erro: {e}")
                finally:
                    index += 1
            else:
                index += 1
       
        else:
            try:
                document_name = driver.find_element(By.XPATH, f'//tbody/tr[{index}]/td[{dn}]').text
                document_type_workday = driver.find_element(By.XPATH, f'//tbody/tr[{index}]/td[{dtw}]').text
                document_folder_workday = driver.find_element(By.XPATH, f'//tbody/tr[{index}]/td[2]').text
                document_id = document_name[28:36]
                document_id = int(document_id)
                document_date = document_name[17:27]

                print(f'''employee_id: {i}
                Nome do documento: {document_name}
                ID do documento: {document_id}
                Categoria do documento: {document_type_workday}
                Pasta do documento: {document_folder_workday}
                Data do documento: {document_date}
                ''')
                
                query = '''
                INSERT INTO documents_workday (document_id_wd, employee_id_wd, document_name, document_folder, document_type, document_date)
                VALUES (?, ?, ?, ?, ?, ?)
                '''
                connector = sqlite3.connect('bd_nomes.sdb')
                cursor = connector.cursor()
                cursor.execute(query, (document_id, i, document_name,  document_folder_workday, document_type_workday, document_date))
                connector.commit()
                connector.close()

            except sqlite3.Error as e:
                print(f"Ocorreu um erro ao inserir os dados: {e}")
                connector.close()
            except NoSuchElementException:
                print("Elemento não encontrado, encerrando o loop.")
                break
            except Exception as e:
                print(f"Ocorreu um erro: {e}")
            finally:
                index += 1

def workday():
    email, password = get_email_password(0)
    options = browser.set_browser_options('Paulo')
    service = browser.set_browser_service_update_chromedriver()
    driver = browser.open_browser(service, options, 'https://wd3.myworkday.com/syngenta/d/home.htmld', minimize_window = False, set_standard_window_size = False)
    wait = WebDriverWait(driver, 30)
    browser.microsoft_login(driver, email, password, 'https://wd3.myworkday.com')
    sleep(3)

    employee_id = []
    if driver.current_url == 'https://wd3.myworkday.com/wday/authgwy/syngenta/login.htmld':
        skip_button = driver.find_element(By.XPATH, '//*[@data-automation-id="linkButton"]')
        skip_button.click()
    sleep(10)
    for i in employee_id:
        find = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="wd-searchInput"]/input')))
        sleep(3)
        find.clear()
        sleep(3)
        find.send_keys(i)
        sleep(1)
        button_find = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="wd-searchInput"]/div')))
        sleep(0.5)
        button_find.click()
        sleep(1)
        employee_name = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@class="pexsearch-5lrqs3"]')))
        sleep(0.5)
        employee_name.click()
        sleep(1)
        try:
            more_button = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@data-automation-id="workerProfileMoreLinkMenuItemWrapper"]')))
            sleep(0.5)
            more_button.click()
        except:
            None
        sleep(1)
        personal_button = wait.until(EC.visibility_of_element_located((By.XPATH, '//li[@class="WPOR"][7]')))
        sleep(0.5)
        personal_button.click()
        sleep(1)
        document_button = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@class="WJE2 WLC2"][2]')))
        sleep(0.5)
        document_button.click()
        sleep(1)

        a = wait.until(EC.visibility_of_element_located((By.XPATH, '//thead/tr[1]/th[3]'))) #esse tem que esperar aparecer
        check_exist_column = driver.find_element(By.XPATH, f'//thead/tr[1]/th[3]').text
        print(check_exist_column)
        if 'Processo de negócios' in check_exist_column:
            dn = 6
            dtw = 4
        else:
            dn = 5
            dtw = 3
        
        index = 1
        document_informations(driver, check_exist_column,i ,index, dn, dtw)
        clear = driver.find_element(By.XPATH, '//*[@data-automation-id="searchInputClearTextIcon"]')
        clear.click()
    
    driver.quit()

def aconso():
    employee_id = []
    email, password = get_email_password(0)
    options = browser.set_browser_options('Paulo')
    service = browser.set_browser_service_update_chromedriver()
    driver = browser.open_browser(service, options, 'https://aconso.cloud/aconso/noFramesMain.jsp?_override_url=login/login_gate_desktop.jsp&mandant=d068', minimize_window = False, set_standard_window_size = False)
    wait = WebDriverWait(driver, 30)
    browser.microsoft_login(driver, email, password, 'https://aconso.cloud/aconso')
    sleep(3)

    try:
        dropdown = driver.find_element(By.XPATH, '//*[@id="loginComponent_locale"]')
        select = Select(dropdown)
        select.select_by_visible_text('English')
        sleep(3)
        login_button = driver.find_element(By.XPATH, '//*[@id="Ajax_loginComponent"]/input')
        login_button.click()
        sleep(10)
    except:
        None

    for i in employee_id:

        wait = WebDriverWait(driver, 30)
        driver.get('https://aconso.cloud/frontend/users')
        sleep(2)
        search_input = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@class="form-control mt-3 form-search"]')))
        sleep(0.5)
        search_input.send_keys(i)
        sleep(2)
        search_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@class="btn btn-primary btn-customSearch"]')))
        sleep(0.5)
        search_button.click()
        sleep(2)
        employee_card = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@class="col-sm-6 col-md-4 col-lg-3"][1]')))
        sleep(0.5)
        employee_card.click()
        sleep(2)
        documents_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@class="btn btn-toDocuments hidden-print"]')))
        sleep(0.5)
        documents_button.click()
        sleep(5)
        
        index = 1
        while True:
            try:
                document = driver.find_element(By.XPATH, f'//div[1]/div[2]/main/div[3]/div[2]/div[2]/div/div[2]/div/div/ul/li[{index}]')
                driver.execute_script("arguments[0].scrollIntoView(true);", document)
                document_name = document.text
                document_id = document.get_attribute('data-document-id')
                index_date = document_name.find("-") + 2
                index_type = document_name.find("-") - 1
                document_category = document_name[:index_type]
                document_date = document_name[index_date:]

                print(f'''employee_id: {i}')
                Nome do documento: {document_name}
                ID do documento: {document_id}
                Categoria do documento: {document_category}
                Data do documento: {document_date}
                ''')

                query = '''
                INSERT INTO documents_aconso (document_id_ac, employee_id_ac, document_name, document_type, document_date)
                VALUES (?, ?, ?, ?, ?)
                '''
                connector = sqlite3.connect('bd_nomes.sdb')
                cursor = connector.cursor()

                # Inserir os dados no banco de dados
                try:
                    cursor.execute(query, (document_id, i, document_name,  document_category, document_date))
                    connector.commit()
                except sqlite3.Error as e:
                    None
                    #print(f"Ocorreu um erro ao inserir os dados: {e}")

                # Fechar a conexão com o banco de dados
                connector.close()

                index += 1
            except:
                break

    driver.quit()


aconso()
