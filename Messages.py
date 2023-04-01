import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import random
import string
import time
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# Open the Excel file
wb = openpyxl.load_workbook('Book1.xlsx')
sheet = wb.active
BASE_URL = "https://web.whatsapp.com/"
CHAT_URL = "https://web.whatsapp.com/send?phone={phone}&text&type=phone_number&app_absent=1"
# Get the phone numbers and messages from the Excel sheet
phone_numbers = [cell.value for cell in sheet['A']]
for i in range(len(phone_numbers)):
    print(phone_numbers[i])
messages = [cell.value for cell in sheet['B']]

# Open WhatsApp Web
'''driver = webdriver.Chrome()
driver.get("https://web.whatsapp.com/")'''
chrome_options = Options()
chrome_options.add_argument("start-maximized")
user_data_dir = ''.join(random.choices(string.ascii_letters, k=8))
chrome_options.add_argument("--E:/4thYearProject/" + user_data_dir)
chrome_options.add_argument("--incognito")

browser = webdriver.Chrome(
    ChromeDriverManager().install(),
    options=chrome_options,
)

browser.get(BASE_URL)
browser.maximize_window()

for i in range(len(phone_numbers)):
    # Find the search box and search for the phone number
    print("hello")
    browser.get(CHAT_URL.format(phone=phone_numbers[i]))
    time.sleep(3)

    # Find the message box and send the message
    inp_xpath = (
    '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]')
    input_box = WebDriverWait(browser, 60).until(
    expected_conditions.presence_of_element_located((By.XPATH, inp_xpath)))
    input_box.send_keys(messages[i])
    input_box.send_keys(Keys.ENTER)

    time.sleep(8)  # Wait for 8 seconds before sending the next message

# Close the browser

