import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import pandas as pd
from src.cleaner import normalize_name

df_ind = pd.read_excel("data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2(1-360).xlsx")
student_col = [c for c in df_ind.columns if "Seleccione el nombre del Estudiante" in c][0]
evaluator_col = [c for c in df_ind.columns if "Seleccione su nombre (Evaluador)" in c][0]

df_ind["Student_Norm"] = df_ind[student_col].apply(normalize_name)

counts = df_ind["Student_Norm"].value_counts()
students_4 = counts[counts == 4].index.tolist()

df_sched = pd.read_excel("data/presentaciones_crcronograma.xlsx", sheet_name="Hoja1")
df_sched["Student_Norm"] = df_sched["NOMBRE"].apply(normalize_name)

print(f"Students with 4 reviews: {len(students_4)}")
for s_norm in students_4:
    # Get raw name
    s_raw = df_ind[df_ind["Student_Norm"] == s_norm][student_col].iloc[0]
    print(f"\n==================== Student: {s_raw} ====================")
    
    # Schedule
    row = df_sched[df_sched["Student_Norm"] == s_norm]
    if not row.empty:
        r = row.iloc[0]
        print(f"Scheduled: Titular={r['DOCENTE TITULACIÓN']}, Tutor={r['TUTOR']}, Tercer={r['TERCER DOCENTE']}, Adic={r['DOCENTE ADICIONAL']}")
    else:
        print("Not scheduled!")
        
    # Reviews
    sub = df_ind[df_ind["Student_Norm"] == s_norm]
    for _, item in sub.iterrows():
        print(f"  ID: {item['ID']} | Selected: {item[evaluator_col]} | M365: {item['Name']} | Email: {item['Email']}")
