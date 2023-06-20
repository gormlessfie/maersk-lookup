from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from openpyxl import Workbook
from datetime import datetime

def clears_cookies_page(driver):
    # Wait until the cookies accept element is visible
    cookies_accept_element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.XPATH, "//button[@aria-label='Allow all']"))
    )

    # Click on the cookies accept button
    cookies_accept_element.click()

def search(driver, tracker):
    # Fills out input box with tracking #
    input_box = driver.find_element(By.XPATH, "//input[@placeholder='Enter a tracking ID']")
    input_box.send_keys(tracker)
    input_box.send_keys(Keys.ENTER)

def retrieve_date_info(driver):
    try:
        # Wait until the date element is visible
        eta_date_element = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//dd[@class='container-info__text container-info__text--date']"))
        )
        
        #Collects date element
        return eta_date_element.text
    
    except NoSuchElementException:
        # Handle the case where the search element or any other expected element is not found
        return "Date not found"

    except Exception as e:
        # Handle any other unexpected exception
        print("An unexpected error occurred:", str(e))
        return "Date not found"

def go_home(driver):
    # Returns to main page
    home_button = driver.find_element(By.XPATH, "//a[@class='ign-header__logo ign-track']//*[name()='svg']")
    home_button.click()
    
def export_dates(list_searched, worksheet):
    for entry in list_searched:
        row = entry
        worksheet.append(row)
        
def format_date(date):
    # Parse the input string into a datetime object
    date_object = datetime.strptime(date, "%d %b %Y %H:%M")

    # Format the date as "month/day"
    formatted_date = date_object.strftime("%m/%d")
    return formatted_date

# Setup Selenium
driver = webdriver.Firefox()
driver.get("https://www.maersk.com/")
list_tracking_numbers = open('list-trackers.txt', 'r').readlines()

# Setup excel workbook
workbook = Workbook()
worksheet = workbook.active
worksheet.title = "Shipping Date Changes"
worksheet.column_dimensions['A'].width = 25

clears_cookies_page(driver)

list_searched = []
for tracker in list_tracking_numbers:
    search(driver, tracker)
    
    try:
        date = format_date(retrieve_date_info(driver))
    except ValueError:
        date = "No ETA Found"
    
    tracker = tracker.strip()
    row = [tracker, date]
    list_searched.append(row)
    go_home(driver)

export_dates(list_searched, worksheet)
workbook.save("output/maersk_shipping_dates_changes.xlsx")

driver.close()