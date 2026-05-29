import openpyxl
wb = openpyxl.load_workbook("data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2_PROCESADO.xlsx", read_only=True)
sheet = wb["Sheet1"]
headers = [cell.value for cell in sheet[1]]
print("Sheet1 Headers:")
for idx, h in enumerate(headers):
    print(f"  {idx}: {h}")
