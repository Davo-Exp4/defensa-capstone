import pandas as pd
df = pd.read_excel("data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2_PROCESADO.xlsx", sheet_name="Calculo")
print("Total rows:", len(df))
print(df[["Seleccione el nombre del Estudiante", "Cuenta de Seleccione su nombre (Evaluador)", "Nota ponderada"]].to_string())
