import openpyxl

class WorkbookProcessor:
    
    data = None
        
        
    def __init__(self):
        self.data = []
        
    def process(self, workbook_url):
        workbook = openpyxl.load_workbook(workbook_url)
        worksheet = workbook.active
        
        header = []
        
        for row in worksheet.iter_rows(values_only = True):
            person_data = {}
            
            if header.__len__() == 0:
                for cell in row:
                    if cell == None:
                        break
                    header.append(cell)
                continue
            
            for header_item, cell in zip(header, row):
                if cell == None:
                    continue
                
                person_data[header_item] = cell
                
            self.data.append(person_data)
                
        return self.data

    