import sys
sys.path.append('C:/Automacoes/Projetos_Syngenta')
import utilities.browser_functions as browser
from utilities.others import get_email_password
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import sqlite3


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
    sleep(3)

    for i in employee_id:
        find = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="wd-searchInput"]/input')))
        sleep(1)
        find.send_keys(i)
        sleep(1)
        button_find = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="wd-searchInput"]/div')))
        button_find.click()
        sleep(1)
        a = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@class="pexsearch-5lrqs3"]')))
        employee_name = driver.find_element(By.XPATH, '//*[@class="pexsearch-5lrqs3"]').text
        
        conn = sqlite3.connect('bd_nomes.sdb')
        cursor = conn.cursor()
        cursor.execute('''
        UPDATE employee_names
        SET name_workday = ?
        WHERE employee_id = ?
        ''', (employee_name, i))
        conn.commit()
        conn.close()

        clear = driver.find_element(By.XPATH, '//*[@data-automation-id="searchInputClearTextIcon"]')
        clear.click()
        sleep(3)
    
    driver.quit()

def aconso():
    employee_id = []
    email, password = get_email_password(0)
    options = browser.set_browser_options('Paulo')
    service = browser.set_browser_service_update_chromedriver()
    driver = browser.open_browser(service, options, 'https://wd3.myworkday.com/syngenta/d/home.htmld', minimize_window = False, set_standard_window_size = False)
    browser.microsoft_login(driver, email, password, 'https://wd3.myworkday.com')
    sleep(3)

    dropdown = driver.find_element(By.XPATH, '//*[@id="loginComponent_locale"]')
    select = Select(dropdown)
    select.select_by_visible_text('English')
    sleep(3)
    login_button = driver.find_element(By.XPATH, '//*[@id="Ajax_loginComponent"]/input')
    login_button.click()
    sleep(10)

    for i in employee_id:
        driver.get('https://aconso.cloud/frontend/users')
        sleep(5)
        search_input = driver.find_element(By.XPATH, '//*[@class="form-control mt-3 form-search"]')
        search_input.send_keys(i)
        sleep(2)
        search_button = driver.find_element(By.XPATH, '//*[@class="btn btn-primary btn-customSearch"]')
        search_button.click()
        sleep(5)
        employee_card = driver.find_element(By.XPATH, '//*[@class="col-sm-6 col-md-4 col-lg-3"][1]')
        employee_card.click()
        sleep(5)
        first_name = driver.find_element(By.XPATH, '//*[@class="text text--md"][5]').text
        last_name = driver.find_element(By.XPATH, '//*[@class="text text--md"][7]').text
        sleep(5)

        start_index = first_name.find(":") + 1
        first_name = first_name[start_index:].strip()
        start_index = last_name.find(":") + 1
        last_name = last_name[start_index:].strip()
        employee_name = f'{first_name} {last_name}'

        conn = sqlite3.connect('bd_nomes.sdb')
        cursor = conn.cursor()
        cursor.execute('''
        UPDATE employee_names
        SET name_aconso = ?
        WHERE employee_id = ?
        ''', (employee_name, i))
        conn.commit()
        conn.close()

    driver.quit()
    

workday()
