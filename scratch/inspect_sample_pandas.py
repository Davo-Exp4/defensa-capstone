import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import pandas as pd
from src.cleaner import normalize_name

df_ind, df_calc, df_comp, _ = pd.read_excel("data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2(1-360).xlsx"), None, None, None

# Find student and evaluator column names
student_col = [c for c in df_ind.columns if "Seleccione el nombre del Estudiante" in c][0]
evaluator_col = [c for c in df_ind.columns if "Seleccione su nombre (Evaluador)" in c][0]

df_ind["Student_Norm"] = df_ind[student_col].apply(normalize_name)
df_ind["Evaluator_Norm"] = df_ind[evaluator_col].apply(normalize_name)

# Scheduled
df_sched = pd.read_excel("data/presentaciones_crcronograma.xlsx", sheet_name="Hoja1")
df_sched["Student_Norm"] = df_sched["NOMBRE"].apply(normalize_name)

sample_students = [
    "RAMOS VASCONEZ XAVIER ALEJANDRO",
    "AGUILAR RODRIGUEZ CAMILA ESTEFANIA",
    "CHIRIBOGA TERAN JUAN MARTIN"
]

for s in sample_students:
    print(f"\n==================== {s} ====================")
    s_norm = normalize_name(s)
    
    # Schedule info
    row = df_sched[df_sched["Student_Norm"] == s_norm]
    if not row.empty:
        r = row.iloc[0]
        print(f"Scheduled jurors: Titular={r['DOCENTE TITULACIÓN']}, Tutor={r['TUTOR']}, Tercer={r['TERCER DOCENTE']}, Adic={r['DOCENTE ADICIONAL']}")
    else:
        print("Not found in schedule!")
        
    # Submissions info
    sub = df_ind[df_ind["Student_Norm"] == s_norm]
    print(f"Submissions ({len(sub)}):")
    for _, item in sub.iterrows():
        print(f"  ID: {item['ID']} | Selected Evaluator: {item[evaluator_col]} | M365 Name: {item['Name']} | Email: {item['Email']}")
