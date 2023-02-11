from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import ElementNotVisibleException
import re
from time import sleep
import json
import argparse


def download_report(report, start_date, end_date):
    # load configuration file
    try:
        config = open('config.json', encoding='utf8', errors='ignore')
        config_dict = json.load(config)
    except FileNotFoundError:
        print('The configuration file was not found.')
    except:
        print('The configuration file could not be opened.')

    edupage_url = config_dict['url_links']['edupage']
    chromedriver_path = config_dict['paths']['chromedriver_path']
    username = config_dict['credentials']['login']
    password = config_dict['credentials']['password']

    # initiate webdriver
    options = webdriver.ChromeOptions()
    #options.add_argument("--start-maximized")
    service = Service(executable_path=chromedriver_path)
    driver = webdriver.Chrome(service=service, options=options)

    # load edupage
    try:
        driver.get(edupage_url)
    except Exception as e:
        print("Page was not loaded:", e.message)
        driver.quit()

    # remove cookies pop up
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "a.flat-button.flat-button-greend.eu-cookie-closeBtn"))).click()
    except ElementNotVisibleException:
        print("Cookies window didn't load")

    # login
    try:
        login_input = driver.find_element(By.ID, 'home_Login_1e1')
        login_input.clear()
        login_input.send_keys(username)

        password_input = driver.find_element(By.ID, 'home_Login_1e2')
        password_input.clear()
        password_input.send_keys(password)

        login = driver.find_element(By.XPATH,
                                    "//span[@class='skgdFormValue skgdSubmitRow']//input[@class='skgdFormSubmit']")
        login.click()
        print('Login successful')
    except Exception as e:
        print('Login unsuccessful due to error:', e.message)
        driver.quit()

    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "div.ebicon.ebicon-module-vyucovanie"))).click()
    except ElementNotVisibleException:
        print("Edupage was not loaded.")
        driver.quit()

    # open Reports card
    try:
        reports_card = driver.find_element(By.CSS_SELECTOR, "div.ebicon.ebicon-module-reports")
        reports_card.click()
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'asc-combo')]"))).click()
    except Exception as e:
        print('Reports window was not loaded:', e.message)
        driver.quit()

    # set report type
    menu = driver.find_element(By.XPATH, '/html/body/div[2]/div[12]/div[3]')
    menu_str = menu.text
    menu_list = menu_str.split("\n")

    input_rex = re.compile(".*" + report + ".*")
    input_list = list(filter(input_rex.match, menu_list))
    report = driver.find_element(By.LINK_TEXT, input_list[0])
    report.click()

    # set date range
    try:
        from_date = driver.find_element(By.XPATH,
                                        "/html/body/div[2]/div[2]/div/div/div[2]/div/div/div[1]/div[3]/div/div[3]/input[2]")
        from_date.clear()
        from_date.send_keys(start_date)

        to_date = driver.find_element(By.XPATH,
                                      "/html/body/div[2]/div[2]/div/div/div[2]/div/div/div[1]/div[3]/div/div[3]/span/input[2]")
        to_date.clear()
        to_date.send_keys(end_date)
    except Exception as e:
        print('Date buttons not found:', e.message)

    # set page orientation
    try:
        page_setup_dropdown = driver.find_element(By.XPATH, "//span[contains(@title, 'A4 - Na')]")
        page_setup_dropdown.click()
        page_setup = driver.find_element(By.PARTIAL_LINK_TEXT, "Na výšku")
        page_setup.click()
    except Exception as e:
        print('Page orientation button not found:', e.message)

    # refresh and save
    try:
        refresh_button = driver.find_element(By.XPATH,
                                             "//div[contains(@class, 'right')]//div[contains(@class, 'asc-ribbon-button')]//div[contains(@class, 'middle')]//img")
        refresh_button.click()
    except Exception as e:
        print('Refresh buttons not found:', e.message)

    sleep(3)

    try:
        save_button = driver.find_element(By.XPATH, "//button[contains(@class, 'asc-button button-gray menu')]")
        save_button.click()
        save_setup = driver.find_element(By.PARTIAL_LINK_TEXT, "PDF")
        save_setup.click()
    except Exception as e:
        print('Save button not found', e.message)

    sleep(20)

    driver.quit()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-r", "--report", type=str)
    parser.add_argument("-s", "--startdate", type=str, help='Date in DD.MM.YYYY format')
    parser.add_argument("-e", "--enddate", type=str, help='Date in DD.MM.YYYY format')
    args = vars(parser.parse_args())

    report = args["report"]
    start_date = args["startdate"]
    end_date = args["enddate"]

    download_report(report, start_date, end_date)
