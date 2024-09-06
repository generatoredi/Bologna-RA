from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from time import sleep
from tqdm import tqdm
import math
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
#########################################



driver = webdriver.Chrome()
driver.get("https://sba.unibo.it/it/almare/collezioni/banche-dati")


def press_import_button(driver):
    # Press the button
    driver.find_element(By.XPATH, "//*[contains(text(), 'Actions')]").click()
    sleep(10)
    driver.find_element(By.XPATH, "//*[contains(text(), 'Export companies')]").click()

def import_company_info_aux(driver, ii, nbr_comp, chunk_size):
    # Open the export company stuff, put it in a while loop so it does not bug
    ready = 0
    while ready == 0:
        try:
            press_import_button(driver)
            ready = 1
        except NoSuchElementException:
            ready = 0

    # Switch to txt format
    element_present = EC.presence_of_element_located((By.XPATH, "//select[@id='component_FormatTypeSelectedId']/option[@value='ExcelDisplay2007']"))
    WebDriverWait(driver, 30).until(element_present)
    driver.find_element(By.XPATH, "//select[@id='component_FormatTypeSelectedId']/option[@value='ExcelDisplay2007']").click()

    # Put the range of companies we want
    driver.find_element(By.XPATH, "//select[@id='component_RangeOptionSelectedId']/option[contains(text(),'A range')]").click()

    # Sending the information about which range of companies we want
    driver.find_element(By.NAME, "component.From").send_keys(str(ii * chunk_size + 1))
    driver.find_element(By.NAME, "component.To").send_keys(str(min((ii + 1) * chunk_size, nbr_comp)))

    # Check the option to send the export via e-mail
    sleep(3)
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.NAME, "component.Send")))
    send_element = driver.find_element(By.NAME, "component.Send")
    send_element.click()

    # Ensure the "Export" button is clickable
    export_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[@class='button submit ok']")))

    # Scroll to the element to bring it into view (optional)
    actions = ActionChains(driver)
    actions.move_to_element(export_button).perform()

    # Click the "Export" button using JavaScript
    driver.execute_script("arguments[0].click();", export_button)

    # Wait for a few seconds (you may adjust the sleep duration as needed)
    sleep(3)

    # Close the window
    element_present_close = EC.presence_of_element_located((By.XPATH, '//*[@data-popup-close="cancel"]'))
    WebDriverWait(driver, 1200).until(element_present_close)
    sleep(1)
    driver.find_element(By.XPATH, '//*[@data-popup-close="cancel"]').click()

    # Sleep on it before continuing
    sleep(10)

def import_company_info(driver, upper_bound=10000000000, chunk_size=2_500, start_point=0):
    nbr_comp = driver.find_element(By.CSS_SELECTOR, "h3[data-total-records-count]").get_attribute("data-total-records-count")
    nbr_comp = min(int(nbr_comp), upper_bound)
    nbr_run = math.ceil(int(nbr_comp) / chunk_size)
    start_iteration = max(start_point // chunk_size, 0)

    for ii in tqdm(range(start_iteration, nbr_run)):
        try:
            import_company_info_aux(driver, ii, nbr_comp, chunk_size)
        except (NoSuchElementException, TimeoutException, ElementClickInterceptedException):
            import_company_info_aux(driver, ii, nbr_comp, chunk_size)

# Replace 5000 with your desired starting point
import_company_info(driver, upper_bound=10000000000, chunk_size=2_500, start_point=0)
