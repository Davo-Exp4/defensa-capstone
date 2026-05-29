import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.engine import process_oral_defense, process_capstone_written

print("=== TESTING ORAL DEFENSE COMPLIANCE ===")
df_ind, df_calc, df_comp, df_sched = process_oral_defense(
    "data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2(1-360).xlsx",
    "data/presentaciones_crcronograma.xlsx",
    exclude_duplicates=True
)

print("Compliance shape:", df_comp.shape)
print("Compliance columns:", list(df_comp.columns))
print("Any replacements in compliance?", df_comp["Is_Replacement"].any())
print("Replacements count in compliance:", df_comp["Is_Replacement"].sum())
if df_comp["Is_Replacement"].any():
    print(df_comp[df_comp["Is_Replacement"] == True][["Docente", "Docente_Real", "Estudiante", "Estado"]].head(5))

print("\n=== TESTING WRITTEN CAPSTONE COMPLIANCE ===")
df_ind_w, df_calc_w, df_comp_w, df_sched_w = process_capstone_written(
    "data/EVALUACIÓN PROYECTO CAPSTONE - COHORTE 2(1-39).xlsx",
    "data/presentaciones_crcronograma.xlsx",
    exclude_duplicates=True
)

print("Compliance shape:", df_comp_w.shape)
print("Compliance columns:", list(df_comp_w.columns))
print("Any replacements in compliance?", df_comp_w["Is_Replacement"].any())
print("Replacements count in compliance:", df_comp_w["Is_Replacement"].sum())

print("\nSUCCESS!")
