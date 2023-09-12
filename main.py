import rpa_executor

driver_path = r"C:\Users\nivankiv\Documents\chrome-win64\chrome.exe"
file_path = r"C:\Users\nivankiv\Downloads\challenge.xlsx"
challenge_url = "https://rpachallenge.com/"


excecutor = rpa_executor.RPAExcecutor()

excecutor.execute_rpa_script(challenge_url, file_path)

#processor = workbook_processor.WorkbookProcessor(r"C:\Users\nivankiv\Downloads\challenge.xlsx")


