import openpyxl
wb = openpyxl.load_workbook("data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2_PROCESADO.xlsx", read_only=True)
sheet = wb["Sheet1"]
print("Sheet1 max_row in processed file:", sheet.max_row)
