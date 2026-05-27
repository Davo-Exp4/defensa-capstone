import pandas as pd
import openpyxl
from src.engine import process_oral_defense

raw_excel = "data/cohorte_pasada_crudo.xlsx"
processed_excel = "data/cohorte_pasada_procesado.xlsx"

df_ind_ours, df_calc_ours, df_comp_ours, _ = process_oral_defense(raw_excel, processed_excel)
df_calc_hist = pd.read_excel(processed_excel, sheet_name="Calculo", header=2)

df_calc_hist.columns = [str(c).strip() for c in df_calc_hist.columns]
df_calc_ours.columns = [str(c).strip() for c in df_calc_ours.columns]

df_calc_hist = df_calc_hist[df_calc_hist["Seleccione el nombre del Estudiante"].notna()]
df_calc_hist = df_calc_hist[df_calc_hist["Seleccione el nombre del Estudiante"].str.upper() != "TOTAL GENERAL"]

df_calc_ours["student_key"] = df_calc_ours["Seleccione el nombre del Estudiante"].str.strip().str.upper()
df_calc_hist["student_key"] = df_calc_hist["Seleccione el nombre del Estudiante"].str.strip().str.upper()

ours_indexed = df_calc_ours.set_index("student_key")
hist_indexed = df_calc_hist.set_index("student_key")

mismatched_students_count = 0
for s_key in ours_indexed.index:
    if s_key in hist_indexed.index:
        count_ours = ours_indexed.loc[s_key, "Cuenta de Seleccione su nombre (Evaluador)"]
        count_hist = hist_indexed.loc[s_key, "Cuenta de Seleccione su nombre (Evaluador)"]
        
        note_ours = ours_indexed.loc[s_key, "Nota ponderada"]
        note_hist = hist_indexed.loc[s_key, "Nota ponderada"]
        
        if abs(note_ours - note_hist) > 1e-5:
            mismatched_students_count += 1
            print(f"Student: '{ours_indexed.loc[s_key, 'Seleccione el nombre del Estudiante']}'")
            print(f"  Count: Ours={count_ours}, Hist={count_hist}")
            print(f"  Note:  Ours={note_ours:.4f}, Hist={note_hist:.4f}")

print(f"\nTotal mismatched students: {mismatched_students_count}")
