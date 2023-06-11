from selenium import webdriver
import time
import json
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import chromedriver_autoinstaller
from selenium.webdriver import DesiredCapabilities
import re
import csv
import requests
import pandas as pd

def visibil_element(driver, by, selector, wait=10): ### web element find and search
    element = False
    if by == 'xpath':
        byselector = By.XPATH
    try:
        element = WebDriverWait(driver, wait).until(
            EC.visibility_of_element_located((byselector, selector)))
    except Exception :
        element = False
    return element

def login(driver):
    driver.get('https://orappl.candy.it/giastech/index.jsf')
    visibil_element(driver, 'xpath', '//input[@type="text"]', wait=30)
    driver.find_element(By.XPATH, '//input[@type="text"]').send_keys('78001048')
    driver.find_element(By.XPATH, '//input[@type="password"]').send_keys('Wm1Mark')
    driver.find_element(By.XPATH, '//button[@type="submit"]').click()
    driver.get('https://orappl.candy.it/documentation/model_find.jsf')
    visibil_element(driver, 'xpath', '//input[@type="text"]', wait=30)

def return_rows(driver,code):
    if code[0].isdigit():
        search = driver.find_element(By.XPATH, '(//input[@type="text"])[1]')
        search.clear()
        search.send_keys(code)
        search.send_keys(Keys.RETURN)
        driver.find_element(By.XPATH, '(//input[@type="text"])[1]').clear()
        time.sleep(1)
        data = driver.find_elements(By.XPATH, '//td[@class="content"]//tbody//a')
        if len(data) != 0:
            return len(data)

    search = driver.find_element(By.XPATH, '(//input[@type="text"])[2]')
    search.clear()
    search.send_keys(code)
    search.send_keys(Keys.RETURN)
    driver.find_element(By.XPATH, '(//input[@type="text"])[2]').clear()
    time.sleep(1)
    data = driver.find_elements(By.XPATH, '//td[@class="content"]//tbody//a')
    return len(data)

def download_pdf(url, save_path):
    response = requests.get(url)
    with open(save_path, 'wb') as file:
        file.write(response.content)

def pdf_link(driver):
    while True:
        logs_raw = driver.get_log("performance")
        logs = [json.loads(lr["message"])["message"] for lr in logs_raw]
        xhr_logs = [log for log in logs if "Network.responseReceived" in log["method"]]
        if xhr_logs:
            break
    network_response_data = None
    for log in xhr_logs:
        try:
            request_id = log["params"]["requestId"]
            resp_url = log["params"]["response"]["url"]
            if ("model_content.jsf" in resp_url):
                network_response_data = log['params']['response']['headers']['Content-Disposition']
                pdf_url = re.findall(r'"(.*?)"', network_response_data)[0]
                # print(pdf_url)
                return pdf_url
        except:
            pass

def details_collect(driver,data_id):
    pin_sin = visibil_element(driver, 'xpath', '//span[@id="body:form_content_model:mod_id"]', wait=30).text
    pin = pin_sin.split('-')[0].strip()
    sin = pin_sin.split('-')[1].strip()
    try:
        machine_type = visibil_element(driver, 'xpath', '//span[@id="body:form_content_model:mod_productLine"]', wait=30).text
    except Exception:
        machine_type = " "
    try:
        brand = visibil_element(driver, 'xpath', '//span[@id="body:form_content_model:mod_brand"]', wait=30).text.replace("Brand", '').strip()
    except Exception:
        brand = " "
        
    all_row = driver.find_elements(By.XPATH, '//table[@rules="all"]//tr')
    all_data = []
    save_location = f'PDF/HVR_ED_{pin}.pdf'
    if data_id[0] in ['0', 0]:
        data_id = f"##{str(data_id)}"
    if pin[0] in ['0', 0]:
        pin = f"##{str(pin)}"
    for x in range (len(all_row)):
        try:
            note = driver.find_element(By.XPATH, f'(//table[@rules="all"]//tr)[{x+1}]//td[2]').text
        except Exception:
            data = [data_id,pin,sin,machine_type,brand," "," "," "," "," "," "," "," "]
            all_data.append(data)
            break
        ref = driver.find_element(By.XPATH, f'(//table[@rules="all"]//tr)[{x+1}]//td[3]').text
        Code = driver.find_element(By.XPATH, f'(//table[@rules="all"]//tr)[{x+1}]//td[4]').text
        Haier_product_code = driver.find_element(By.XPATH, f'(//table[@rules="all"]//tr)[{x+1}]//td[5]').text
        Description = driver.find_element(By.XPATH, f'(//table[@rules="all"]//tr)[{x+1}]//td[6]').text
        Begin = driver.find_element(By.XPATH, f'(//table[@rules="all"]//tr)[{x+1}]//td[7]').text
        End = driver.find_element(By.XPATH, f'(//table[@rules="all"]//tr)[{x+1}]//td[8]').text
        try:
            if ref[0] in ['0', 0]:
                ref = f"##{str(ref)}"
        except Exception:pass
        try:
            if Code[0] in ['0', 0]:
                Code = f"##{str(Code)}"
        except Exception:pass
        try:
            if Begin[0] in ['0', 0]:
                Begin = f"##{str(Begin)}"
        except Exception:pass
        try:
            if End[0] in ['0', 0]:
                End = f"##{str(End)}"
        except Exception:pass
        try:
            if Haier_product_code[0] in ['0', 0]:
                Haier_product_code = f"##{str(Haier_product_code)}"
        except Exception:pass
        save_location_csv = f'HVR_ED_{pin}.pdf'
        data = [data_id,pin,sin,machine_type,brand,note,ref,Code,Haier_product_code,Description,Begin,End,save_location_csv]
        all_data.append(data)
    try:
        link_pdf = pdf_link(driver)
        link_pdf = link_pdf.split("\n")[0]
        download_pdf(link_pdf, save_location)
    except Exception:
        return all_data
    return all_data

def save_data(data,csv_filename):
    with open(csv_filename, mode='a', newline='',encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(["ID","Pin","SIN","Machine Type","Brand","Note","Ref","Code","Haier Product Code","Description","Begin","End","PDF File",])
        writer.writerow(data)

def code_check(driver,code,data_id,csv_filename):
    data_len = return_rows(driver,code)
    if data_len != 0:
        for pg in range(50): 
            data_le_1 = driver.find_elements(By.XPATH, '//tr[@class="normalRow"]')
            data_le_2 = driver.find_elements(By.XPATH, '//tr[@class="alternateRow"]')
            data_le = (len(data_le_1)) + (len(data_le_2))
            for i in range(data_le):
                if i != 0:
                    driver.get('https://orappl.candy.it/documentation/model_find.jsf')
                row = visibil_element(driver, 'xpath', f'(//td[@class="content"]//tbody//a)[{i+1}]', wait=30)
                row.click()
                collected_data = details_collect(driver,data_id)
                for data in collected_data:
                    save_data(data,csv_filename)
            try:
                driver.get('https://orappl.candy.it/documentation/model_find.jsf')
                driver.find_element(By.XPATH , '//img[@alt="Next"]')
                driver.find_element(By.XPATH, f"(//table)[9]//tr//child::td//child::a//child::span[text()='{pg+2}']").click()
                print(f"Page > {pg+2}")
                time.sleep(3)
            except Exception:
                break

# driver = webdriver.Chrome(ChromeDriverManager().install())

driver_path = chromedriver_autoinstaller.install()
print(f"=> Driver Path is: {driver_path}")
print('=> Browser Opening....')
capabilities = DesiredCapabilities.CHROME
capabilities["goog:loggingPrefs"] = {"performance": "ALL"}
s = Service(driver_path)
driver = webdriver.Chrome(service=s, desired_capabilities=capabilities)
driver.maximize_window()

login(driver)


df = pd.read_excel('Hoover Scraping Codes for Kawsar_Log 30.5.23.xlsx', header=None)

csv_filename = "data.csv"

minimum = 1
maximum = 1000
count = maximum-minimum
for x in range(count):
    print(f">> {x+1} >> {count+1}")
    data_id = str(df.iloc[minimum + x, 0])
    code = str(df.iloc[minimum + x, 1])
    code_check(driver,code,data_id,csv_filename)

time.sleep(5)
driver.close()