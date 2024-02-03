
import openpyxl
from openpyxl import load_workbook

wb = openpyxl.load_workbook("/Users/danielfonseca/repos/ShipData/data/ship_maintence.xlsx")

def deleteDuplicateRows(workbook):
    ws=workbook['Ship_Maintenance_Cleaned']
    rows = ws.iter_rows(min_row=1)
    seenIDS = set()
    resultSheet = workbook.create_sheet('No Duplicates')
    
    for row in rows:
        if row[2].value not in seenIDS:
            seenIDS.add(row[2].value)
            row_with_values = [cell.value for cell in row]
            resultSheet.append(row_with_values)
    
    return workbook.save("/Users/danielfonseca/repos/ShipData/data/resultData.xlsx")

def countUniqueRows(worksheet):
    rows = worksheet.iter_rows(min_row=2)
    seenIDS = set()

    for row in rows:
        if row[2].value not in seenIDS:
            seenIDS.add(row[2].value)

    return len(seenIDS)
            


    


            
            
    





    
    
    
    



