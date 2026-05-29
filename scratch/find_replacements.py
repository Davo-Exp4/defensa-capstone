import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import pandas as pd
from src.cleaner import normalize_name
from src.engine import check_replacement

df_ind = pd.read_excel("data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2(1-360).xlsx")

# Find student and evaluator column names
student_col = [c for c in df_ind.columns if "Seleccione el nombre del Estudiante" in c][0]
evaluator_col = [c for c in df_ind.columns if "Seleccione su nombre (Evaluador)" in c][0]

replacements = []
for idx, row in df_ind.iterrows():
    eval_raw = row[evaluator_col]
    sub_name = row["Name"]
    sub_email = row["Email"]
    stud_name = row[student_col]
    
    if check_replacement(eval_raw, sub_name):
        replacements.append({
            "ID": row["ID"],
            "Selected (Dropdown)": eval_raw,
            "Logged In (M365)": sub_name,
            "Email": sub_email,
            "Student": stud_name
        })

df_rep = pd.DataFrame(replacements)
print(f"Total replacements found: {len(df_rep)}")
if not df_rep.empty:
    print("\nReplacements list:")
    print(df_rep.to_string())
