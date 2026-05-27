import openpyxl

wb = openpyxl.load_workbook("data/cohorte_pasada_procesado.xlsx", data_only=True)
s_docente = wb["CALIFICACION-DOCENTE"]

print("CALIFICACION-DOCENTE columns:", [cell.value for cell in s_docente[1]])
# Print all rows to see if we can find a group number or other things
for r in range(2, s_docente.max_row + 1):
    vals = [cell.value for cell in s_docente[r]]
    if any(vals):
        print(vals)
