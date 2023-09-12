# Import necessary libraries
import openpyxl
import time
import os
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By

# Define the main function
def runner(default_directory, rpa_webpage):
    # Initialize the WebDriver
    driver = web_driver(default_directory)

    # Navigate to the RPA Challenge website
    driver.get(rpa_webpage)

    # Download a file from the website
    file_downloader(driver)

    # Process data from an Excel workbook to list with dictionaries
    data = workbook_processor(default_directory)

    # Start the challenge on the website
    start_challenge(driver)

    # Submit data to the website
    submit_data(data, driver)

    # Close the WebDriver
    driver.quit()

# Function to initialize the WebDriver with default download directory
def web_driver(default_directory):
    # Configure Chrome options for file downloads and browser window detachment
    options = webdriver.ChromeOptions()
    pref = {"download.default_directory": default_directory}
    options.add_experimental_option("prefs", pref)
    options.add_experimental_option("detach", True)

    # Initialize the WebDriver with the configured options
    driver = webdriver.Chrome(options=options)
    return driver

# Function to download a file from the website
def file_downloader(driver):
    # Find and click the download link on the website
    download_file = driver.find_element(By.PARTIAL_LINK_TEXT, 'download')
    download_file.click()

    # Wait for a few seconds for file to download (adjust as needed)
    time.sleep(2)



# Function to process data from an Excel workbook downloaded from challenge webpage
def workbook_processor(default_directory):
    # Load from default download directory last downloaded file
    files = os.listdir(default_directory)
    files_sorted = sorted(files, key=lambda x: os.path.getmtime(os.path.join(default_directory, x))) #sort files by modified time
    last_downloaded_file = files_sorted[-1] # Chose last element of the list

    # Load the Excel workbook from a specified file path
    workbook = openpyxl.load_workbook(f"{default_directory}\{last_downloaded_file}")
    worksheet = workbook.active

    # Initialize data and header lists
    data = []
    header = []

    # Iterate through rows in the worksheet
    for row in worksheet.iter_rows(values_only=True):
        person_data = {}

        # If header is empty, populate it with the first row
        if not header:
            for cell in row:
                if cell is None:
                    break
                header.append(cell)
            continue # to avoid adding first row from sheet to dictionary values

        # Populate person_data dictionary with data from the row
        for header_item, cell in zip(header, row):
            if cell is not None:
                person_data[header_item] = cell

        # Append person_data to the data list
        data.append(person_data)

    return data

# Function to start the challenge on the website
def start_challenge(driver):
    # Find and click the start button on the website
    start_button = driver.find_element(By.XPATH, "/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button")
    start_button.click()

# Function to submit data to the website
def submit_data(data, driver):
    for item in data:
        # Fill out web forms with data from the Excel sheet
        driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelFirstName"]').send_keys(item["First Name"])
        driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelLastName"]').send_keys(item["Last Name "])
        driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelCompanyName"]').send_keys(item["Company Name"])
        driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelRole"]').send_keys(item["Role in Company"])
        driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelAddress"]').send_keys(item["Address"])
        driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelEmail"]').send_keys(item["Email"])
        driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelPhone"]').send_keys(item["Phone Number"])

        # Click the submit button on the website
        submit_button = driver.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input')
        submit_button.click()

# Call the main runner function parsing default download directory for web driver and file name
runner(r"C:\RPA Challenge", "https://rpachallenge.com/")
