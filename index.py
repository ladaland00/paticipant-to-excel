import json
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException
import urllib

import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

options = Options()

# Initialize an instance of the chrome driver (browser)
options.page_load_strategy = 'eager'
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 10)

# Load JSON data into a DataFrame
json_data = pd.read_json('data.json')

flat_data = []
error_data = []
for index, item in json_data.iterrows():
    # Check if "coveredEntities" is not empty
    print(index)
    try:
        driver.get("https://www.dataprivacyframework.gov/s/participant-search/participant-detail?id=" +
                   item["id"]+"&status=Inactive")
        wait.until(EC.visibility_of_element_located(
            (By.XPATH, "//*[@class='slds-grid slds-wrap slds-p-around_small']")))
        parent = driver.find_element(
            By.XPATH, "//*[@class='slds-grid slds-wrap slds-p-around_small']")
        print("parent", parent.text)
        name = parent.find_element(
            By.XPATH, "./div[1]/div[1]")
        print("Name", name.text)
        title = parent.find_element(
            By.XPATH, "./div[1]/div[2]")
        email = parent.find_element(
            By.XPATH, "./div[2]/div[1]/a")
        phone = parent.find_element(
            By.XPATH, "./div[2]/div[2]/a")
        address1 = parent.find_element(
            By.XPATH, "./div[1]/div[3]")
        address2 = parent.find_element(
            By.XPATH, "./div[1]/div[4]")
        address3 = parent.find_element(
            By.XPATH, "./div[1]/div[5]")
        flat_data.append({
            "Company": item["name"],
            "Name":  name.text,
            "Title": title.text,  # You can fill this in as needed
            "Email": email.text,  # You can fill this in as needed
            "Address 1": address1.text,
            "Address 2":  address2.text,
            "Address 3": address3.text,
            "Phone": phone.text  # You can fill this in as needed
        })
    except NoSuchElementException:
        error_data.append({
            "id": item["id"],
            "reason":"No element"
            })
        print("No data")
    except TimeoutException:
        time.sleep(5)
        error_data.append({
            "id": item["id"],
            "reason":"No element"
            })
        print("No data")
df = pd.DataFrame(flat_data)
de = pd.DataFrame(error_data)

# Create an Excel writer
excel_writer = pd.ExcelWriter('output.xlsx', engine='xlsxwriter')

# Write the DataFrame to the Excel file
df.to_excel(excel_writer, sheet_name='Sheet1', index=False)
de.to_excel(excel_writer, sheet_name='Sheet2', index=False)
# Save the Excel file

excel_writer.close()
