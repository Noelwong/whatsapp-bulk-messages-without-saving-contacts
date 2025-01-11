# Program to send bulk messages through WhatsApp web from an excel sheet without saving contact numbers
import pathlib

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import pandas
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# Get excel path from user and validate
while True:
    excel_path = input("Please enter the path to your Excel file (e.g., '/Users/Desktop/WhatsApp_Bulk_Message/Recipients_data.xlsx'): ")
    if pathlib.Path(excel_path).is_file():
        try:
            excel_data = pandas.read_excel(excel_path, sheet_name='Recipients')
            break
        except Exception as e:
            print(f"Error reading the Excel file: {e}")
            print("Please try again with a valid Excel file.")
    else:
        print("File not found. Please check the path and try again.")

count = 0

chrome_options = Options()
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.get('https://web.whatsapp.com')


input("Press ENTER after login into Whatsapp Web and your chats are visiable.")
for column in excel_data['Contact'].tolist():
    try:
        phoneNumber = '852'+str(excel_data['Contact'][count])
        message = excel_data['Message'][0].replace(';', '%0A')
        url = 'https://web.whatsapp.com/send?phone={}&text={}'.format(phoneNumber, message)
        sent = False
        # It tries 3 times to send a message in case if there any error occurred
        driver.get(url)
        try:
            # Try to find either English or Chinese send button
            send_button = WebDriverWait(driver, 35).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[aria-label="Send"], button[aria-label="傳送"]'))
            )
            sleep(2)
            send_button.click()
            sent = True
            sleep(5)
            print('Message sent to: ' + str(excel_data['Contact'][count]))
        except Exception as e:
            print("Sorry message could not sent to " + str(excel_data['Contact'][count]))
        count = count + 1
    except Exception as e:
        print('Failed to send message to ' + str(excel_data['Contact'][count]) + str(e))
driver.quit()
print("The script executed successfully.")
