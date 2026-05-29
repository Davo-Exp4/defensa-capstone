import pandas as pd
df = pd.read_excel("data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2(1-360).xlsx")
print("Columns:")
for idx, col in enumerate(df.columns):
    print(f"  {idx}: {col}")
