import time
import workbook_processor
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

class RPAExcecutor:
    
    processor = None
    
    
    def __init__(self):
        self.processor = workbook_processor.WorkbookProcessor()
    
    
    def init_driver(self):
        self.options = webdriver.ChromeOptions()
        self.driver = webdriver.Chrome(options=self.options)
        return self.driver
        
        
    def download_excel_file(self, url):
        self.driver.get(url)
        download_button = self.driver.find_element(By.PARTIAL_LINK_TEXT, 'download')
        download_button.click()
        time.sleep(5)
                    
    def start_challenge(self):
        start_button = self.driver.find_element(By.ID, 'start')
        start_button.click()
        
    def submit_form(self, filepath):
        data = self.processor.process(filepath)
        
        for item in data:
            for key, value in item.items():
                xpath_string = f"label{key}".replace(" ", "")
                
                try:
                    field = self.driver.find_element(By.XPATH, f'//input[@ng-reflect-name="{xpath_string}"]')
                    field.clear()
                    print (f"sfilled box with value {value}")
                    field.send_keys(value)
                except NoSuchElementException:
                    print(f"element {xpath_string} not forund" )
                    continue
        
        submit_button = self.driver.find_element(By.NAME, 'submit')
        submit_button.click()
        
        #https://github.com/Lucas-Meira22/Challenge-RPA/blob/main/Challenge-RPA/main.py
        
    
    def close_browser(self):
        self.driver.quith()
        
    def execute_rpa_script(self, url, filepath):
        self.init_driver()
        #self.download_excel_file(url)
        #self.start_challenge()
        self.submit_form(filepath)
        self.close_browser()
        
        
        
        
        

        
        
    