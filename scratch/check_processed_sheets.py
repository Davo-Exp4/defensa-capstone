import openpyxl
wb = openpyxl.load_workbook("data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2_PROCESADO.xlsx", read_only=True)
print("Cohorte 2 Processed Sheets:", wb.sheetnames)
