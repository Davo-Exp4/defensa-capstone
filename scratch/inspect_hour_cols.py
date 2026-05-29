import openpyxl
import pandas as pd

wb = openpyxl.load_workbook("data/presentaciones_crcronograma.xlsx", data_only=True)
sheet = wb["Hoja1"]
headers = [c.value for c in sheet[1]]
print("Headers:", headers)

rows = []
for r in range(2, sheet.max_row+1):
    row_vals = [cell.value for cell in sheet[r]]
    if any(row_vals):
        rows.append(row_vals)

df = pd.DataFrame(rows, columns=headers)

target_students = [
    "CARRION PARDO STALIN SAMUEL",
    "AMORES FALCONI RODRIGO SEBASTIAN",
    "PONCE TAMAYO JUAN CARLOS",
    "RON LARCO MAURO DANILO",
    "DAVILA VALLEJO NICOLAS SALOMON"
]

print("\n--- Targeted Student Records ---")
for s in target_students:
    # Use simple upper / substring to be robust
    matches = df[df["NOMBRE"].apply(lambda x: str(x).strip().upper() if x else "").str.contains(s.split()[0])]
    if not matches.empty:
        for idx, row in matches.iterrows():
            print(f"\nStudent: {row['NOMBRE']}")
            print(f"  HORARIO REGULAR: {row.get('HORARIO REGULAR')}")
            print(f"  DÍA DEFENSA: {row.get('DÍA DEFENSA')}")
            print(f"  HORA INICIO: {row.get('HORA INICIO')}")
            print(f"  HORA FIN: {row.get('HORA FIN')}")
            print(f"  HORA: {row.get('HORA')}")
            print(f"  SALA: {row.get('SALA')}")
