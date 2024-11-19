# Import the required libraries
import traceback
from openpyxl import load_workbook
from sheetfu import SpreadsheetApp, Table
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import EdgeOptions as Options
from selenium.webdriver import EdgeService as Service
from subprocess import CREATE_NO_WINDOW
import datetime
import queue
import time


def indent_exists_on_webpage(text_box):
    service = Service()
    options = Options()
    # options.add_argument("--headless")
    # options.add_argument("--no-sandbox")
    # options.add_argument("--window-size=1280,720")
    # options.add_argument("--disable-gpu")
    options.add_experimental_option("detach", True)
    service.creation_flags = CREATE_NO_WINDOW
    driver = webdriver.Edge(service=service, options=options)
    driver.get("http://172.17.3.57:8080/IntraNet/Imports/IndentAprovalQry.jsp?myusrid=901519")
    table = driver.find_element(By.ID, "tablea").find_element(By.TAG_NAME, "tbody")
    rows = table.find_elements(By.TAG_NAME, "tr")
    for row in rows:
        tds = row.find_elements(By.TAG_NAME, "td")
        data = tds[0].get_attribute("innerText")
        if indent_number in data:
            return True
    return False


def run_check():
    try:
        sa = SpreadsheetApp('encoded-net-397411-ac78804e5820.json')
        spreadsheet = sa.open_by_id('1dIC0YJrQn2YdiJawl-etvttB83PK-ZQfWcM0zHkNKe8')
        sheet = spreadsheet.get_sheet_by_name('Sheet1')
        data_range = sheet.get_data_range()
        table = Table(data_range)

        for item in table:
            indent_number = item.get_field_value('Indent')
            # status_queue.put(f"Checking Approval for Indent: {indent_number}")
            if "Approve" in item.get_field_value("Approval_Status"):
                continue
            if not indent_exists_on_webpage(indent_number):
                item.set_field_value("Approval_Status", "Approved")
                item.set_field_value("Bot_Updated_Status", "Updated")
                item.set_field_value("Bot_Updated_Dt", datetime.datetime.now().strftime("%d-%m-%Y %H:%M"))
                item.set_field_value("UpdatedBy_Sys", "Bot skipped")
                # status_queue.put(f"Bot skipped approval for Indent: {indent_number}")
                table.commit()
    except Exception as e:
        print(f"Some Issue Approving data {repr(e)}")
        print(traceback.format_exc())
        # status_queue.put(f"Some Issue Approving data {repr(e)} rerunning after 30min")
    finally:
        time.sleep(900)
        run_check()


run_check()
