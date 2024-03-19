#Function to display a warning window
import os
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from .browser_functions import *
from .others import *


def __switch_to_iframe__(driver, link):
    """
    Switches the context of the WebDriver to an iframe found in the current page.

    Args:
        driver (WebDriver): The WebDriver instance controlling the browser.
        link (str): The URL used to locate the iframe by its 'src' attribute.
    """

    wait = WebDriverWait(driver,180)
    src_value = link[32:]
    iframe_xpath = f"//iframe[@src='{src_value}']"
    iframe_element = driver.find_elements(By.XPATH, iframe_xpath)
    if iframe_element:
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, iframe_xpath)))
        print('\033[32mSwitched to Iframe!\033[m')
    else:
        print('\033[32mThe page not have the Iframe, stayed on the main page!\033[m')

def __capture_informations_of_report__(driver):
    """
    Captures the report name and constructs the file name from the report's first column.

    Args:
        driver (WebDriver): The WebDriver instance controlling the browser.

    Returns:
        str: The constructed file name for the report.
    """

    wait = WebDriverWait(driver,180)
    report_name = wait.until(EC.visibility_of_element_located((By.XPATH, '//div[@class="list-title"]')))
    first_collumn = wait.until(EC.element_to_be_clickable((By.XPATH, '(//th)[3]')))
    file_name_without_extension = first_collumn.get_attribute("glide_field")
    index = file_name_without_extension.index('.')
    file_name = f'{file_name_without_extension[:index]}.xlsx'
    print(f'Report name:\033[34m {report_name.text} \033[m\nFile name:\033[34m {file_name} \033[m')
    return file_name

def __transform_paths__(destination_path, file_name):
    """
    Transforms the paths for the origin and destination of the report file.

    Args:
        destination_path (str): The destination path where the report file will be moved.
        file_name (str): The name of the report file.

    Returns:
        tuple: A tuple containing the destination file path and the original file path.
    """

    destination_path_file = f'{destination_path}\\{file_name}'
    pc_user = os.getlogin()
    origin_path_file = f'C:\\Users\\{pc_user}\\Downloads\\{file_name}'
    return destination_path_file, origin_path_file

def __query_excel_file__(driver):
    """
    Initiates the query for an Excel report file within the web page.

    Args:
        driver (WebDriver): The WebDriver instance controlling the browser.
    """

    wait = WebDriverWait(driver, 15)
    actions = ActionChains(driver)
    js_code = """
    var xpathExpression = "(//th)[3]";
    var resultadoXPath = document.evaluate(xpathExpression, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
    var elementoDesejado = resultadoXPath.singleNodeValue;

    // Criar um novo evento de clique com o botão direito
    var eventoCliqueDireito = new MouseEvent("contextmenu", {
        bubbles: true,
        cancelable: true,
        view: window,
        button: 2 // 2 representa o botão direito do mouse
        });

    // Disparar o evento de clique com o botão direito no elemento desejado
    elementoDesejado.dispatchEvent(eventoCliqueDireito);
    """
    driver.execute_script(js_code)
    export_options_button = wait.until(EC.visibility_of_element_located((By.XPATH, '//div[@data-context-menu-label="Export"]')))
    actions.move_to_element(export_options_button).perform()
    sleep(1)
    excel_button = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[6]/div[2]')))
    sleep(1)
    excel_button.click()
    print('\033[32mExcel report query requested!\033[m')

def __wait_heavy_query__(driver):
    """
    Waits for a heavy query to complete by interacting with the 'Wait For It' option if available.

    Args:
        driver (WebDriver): The WebDriver instance controlling the browser.
    """

    actions = ActionChains(driver)
    wait = WebDriverWait(driver, 5)
    try:
        wait_for_it_buttton = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="export_wait"]')))
        sleep(4)
        actions.move_to_element(wait_for_it_buttton).perform()
        wait_for_it_buttton.click()
        print('\033[33mWarning: Heavy file, Wait For It option active.\033[m')
    except:
        None

def __verify_query_run_normaly__(driver):
    """
    Verifies that the query is running normally and waits for the download button to become enabled.

    Args:
        driver (WebDriver): The WebDriver instance controlling the browser.

    Raises:
        Exception: If the query appears to be broken.
    """

    wait = WebDriverWait(driver, 15)
    previous_text = None
    while True:
        actual_text = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="poll_text"]'))).text
        print(actual_text)
        if previous_text is not None and actual_text == previous_text:
            raise Exception('The query is broken')
        else:
            try:
                download_button = driver.find_element(By.XPATH, '//*[@id="download_button"]')
                if download_button.is_enabled():
                    print(f'\033[32mExported all rows of report!\033[m')
                    break
            except:
                None
        previous_text = actual_text
        sleep(5)

def __click_download_button__(driver):
    """
    Clicks the download button to start downloading the report file.

    Args:
        driver (WebDriver): The WebDriver instance controlling the browser.
    """

    download_button = driver.find_element(By.XPATH, '//*[@id="download_button"]')
    download_button.click()
    print(f'\033[33mDownloading the file...\033[m')

def __close_browser__(driver):
    """
    Closes the browser controlled by the WebDriver.

    Args:
        driver (WebDriver): The WebDriver instance controlling the browser.
    """

    driver.quit()
    print('\033[32mBrowser closed!\033[m')
    sleep(2)


def dowload_report_sn(link= str, destination_path= str, email=str, password= str):
    """
    Automates the process of downloading a report from a ServiceNow instance.

    This function handles the entire process from opening the browser, logging in to Microsoft,
    navigating to the report, and downloading it to a specified destination path.

    Args:
        link (str): The URL of the ServiceNow report to be downloaded.
        destination_path (str): The destination path where the report file will be saved.
        email (str): The email address for the Microsoft account.
        password (str): The password for the Microsoft account.

    Returns:
        str: The destination path where the report file was saved.
    """

    options = set_browser_options('Paulo')

    service = set_browser_service_update_chromedriver()

    driver = open_browser(service, options, link)

    microsoft_login(driver, email, password, 'https://syngenta.service-now.com/')

    #Switch to iframe if have it
    __switch_to_iframe__(driver, link)

    file_name = __capture_informations_of_report__(driver)

    paths = __transform_paths__(destination_path, file_name)
    destination_path_file = paths[0]
    origin_path_file = paths[1]

    cleanup_folders(destination_path_file, origin_path_file)

    __query_excel_file__(driver)

    #If the file is large the 'Wait for it' function is actived
    __wait_heavy_query__(driver)

    __verify_query_run_normaly__(driver)

    __click_download_button__(driver)

    move_file_dowloaded(origin_path_file, destination_path_file)

    __close_browser__(driver)

    return destination_path_file