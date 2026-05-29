import pandas as pd
df_calc_hist = pd.read_excel("data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2_PROCESADO.xlsx", sheet_name="Calculo")
print("First row of sheet (which might be headers if we don't skip):")
print(df_calc_hist.head(2).to_string())
print("\nHeaders:")
print(list(df_calc_hist.columns))
