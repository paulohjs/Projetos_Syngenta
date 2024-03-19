#Function to display a warning window
import os
from time import sleep
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def set_browser_options(user):
    """
    Sets the options for the Chrome browser.

    Returns:
        options: An instance of ChromeOptions with the configured settings.
    """

    dir_path = os.path.dirname(__file__)
    profile = os.path.join(dir_path, "profile", f"{user}")
    options = webdriver.ChromeOptions()
    options.add_argument(r"user-data-dir={}".format(profile))
    options.add_argument("--disable-cache")
    print('\033[32mConfigured the options of browser!\033[m')
    return options

def set_browser_service_update_chromedriver():
    """
    Configures the update service for the Chrome browser.

    Returns:
        Service: An instance of Service configured for the Chrome browser.
    """

    service = Service(ChromeDriverManager().install())
    print('\033[32mConfigured the update service of browser!\033[m')
    return service

def open_browser(service, options, link, minimize_window = True, set_standard_window_size = True):
    """
    Opens the Chrome browser with the specified service and options, and navigates to the given link.

    Args:
        service (Service): The service configuration for the browser.
        options (ChromeOptions): The options for the browser.
        link (str): The URL to navigate to after opening the browser.

    Returns:
        WebDriver: An instance of Chrome's WebDriver with the browser opened.
    """

    driver = webdriver.Chrome(options=options, service= service)

    if set_standard_window_size == True:
        width = 1024  # Largura em pixels
        height = 768  # Altura em pixels
        driver.set_window_size(width, height)

    if minimize_window == True:
        driver.minimize_window()

    driver.get(link)
    print('\033[32mOpened of browser on the page!\033[m')
    return driver

def microsoft_login(driver, email, password, main_https):
    """
    Attempts to log in to a Microsoft account within a web page using the provided email and password.
    The function continues to attempt login until the current URL of the driver matches the main_https parameter,
    indicating that the login has been completed and the main application page has been reached.

    Args:
        driver (WebDriver): The WebDriver instance controlling the browser.
        email (str): The email address for the Microsoft account.
        password (str): The password for the Microsoft account.
        main_https (str): The HTTPS URL of the main application page to be reached after a successful login.

    """

    wait = WebDriverWait(driver, 15)
    target_url = True
    count = 0
    while target_url:
        if main_https in driver.current_url:
            target_url = False
            print('\033[32mLogin completed!\033[m')
        elif 'login.microsoftonline.com' in driver.current_url:
            sleep(5)

            #Try to complete each login step (it is not always necessary to complete all steps)
            try:
                validation_code = driver.find_element(By.XPATH, '//*[@id="idRichContext_DisplaySign"]')
                if validation_code:
                    validation_code = driver.find_element(By.XPATH, '//*[@id="idRichContext_DisplaySign"]')
                    validation_code_text = validation_code.text
                    count +=1
                    if count == 1:
                        print(f'Validation code: {validation_code_text}')
            except:
                None

            try:
                email_field = driver.find_element(By.XPATH, '//*[@id="i0116"]')
                if email_field:
                    email_field.send_keys(email)
                    sleep(1)
                    next_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]')))
                    next_button.click()
            except:
                None

            try:
                password_field = driver.find_element(By.XPATH, '//*[@id="i0118"]')
                if password_field:
                    password_field.send_keys(password)
                    sleep(1)
                    signin_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]')))
                    signin_button.click()
            except:
                None
        sleep(2)

def close_browser(driver):
    """
    Closes the browser controlled by the WebDriver.

    Args:
        driver (WebDriver): The WebDriver instance controlling the browser.
    """

    driver.quit()
    print('\033[32mBrowser closed!\033[m')
    sleep(2)