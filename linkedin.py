import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import random

from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# Use the provided credentials to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('C:/Users/user/Documents/sturdy-airport-386809-76d2c0ad810c.json', scope)
gc = gspread.authorize(creds)

# Open the sheet and get all values in the column containing the LinkedIn URLs
doc = gc.open('Saswat Mishra')
worksheet = doc.worksheet('Linkedin URL')
data = worksheet.get_all_values()
urls = [row[1] for row in data[100:300]]

# Set the path to the ChromeDriver executable
webdriver_path = 'C:/Users/user/Downloads/chromedriver.exe'

# Specify the path to the user data directory and profile directory
user_data_dir = 'C:/Users/user/AppData/Local/Google/Chrome/User Data'
profile_directory = 'Profile 15'

# Create a Service object and specify the path to the ChromeDriver executable
service = Service(webdriver_path)

# Create ChromeOptions and set the user data directory and profile directory
chrome_options = Options()
chrome_options.add_argument(f'--user-data-dir={user_data_dir}')
chrome_options.add_argument(f'--profile-directory={profile_directory}')

# Pass the Service object and ChromeOptions to the ChromeDriver
driver = webdriver.Chrome(service=service, options=chrome_options)

# Iterate over the URLs and update the corresponding rows with name and company information
for i, url in enumerate(urls, start=101):  # Start from row 2 since row 1 is header
    try:
        # Open a new tab and navigate to a blank page
        driver.execute_script("window.open();")  # Open a new tab
        driver.switch_to.window(driver.window_handles[-1])  # Switch to the new tab
        driver.get('about:blank')  # Load a blank page in the new tab

        # Open the specified URL
        driver.get(url)

        # Get the name and company information
        element = driver.find_element(By.CLASS_NAME, 'text-heading-xlarge')
        name = element.text
        element1 = driver.find_element(By.CLASS_NAME, 'text-body-medium')
        company = element1.text
        element2 = driver.find_element(By.CLASS_NAME, 'inline-show-more-text')
        company2 = element2.text
        



        # Update the corresponding rows in the worksheet with name and company information
        worksheet.update(f'C{i}', [[name]])  # Update column 3 (C) with the name
        worksheet.update(f'D{i}', [[company]])  # Update column 4 (D) with the company
        worksheet.update(f'E{i}', [[company2]])
        # Add a random delay between 5 and 10 seconds before moving on to the next URL
        time.sleep(random.uniform(5, 10))

    except Exception as e:
        print(f"Error processing URL: {url}")
        print(str(e))

        # Write "NA" in the corresponding cells for name and company information
        worksheet.update(f'C{i}', [['NA']])
        worksheet.update(f'D{i}', [['NA']])
