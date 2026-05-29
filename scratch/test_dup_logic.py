import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.engine import process_oral_defense, process_capstone_written

print("=== TESTING ORAL DEFENSE DUPLICATES ===")
# Case 1: No exclusion
df_ind_1, df_calc_1, df_comp_1, _ = process_oral_defense(
    "data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2(1-360).xlsx",
    "data/presentaciones_crcronograma.xlsx",
    exclude_duplicates=False
)
print("Without exclusion:")
print("  Total individual reviews:", len(df_ind_1))
print("  Total students in Calc:", len(df_calc_1))

# Case 2: With exclusion
df_ind_2, df_calc_2, df_comp_2, _ = process_oral_defense(
    "data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2(1-360).xlsx",
    "data/presentaciones_crcronograma.xlsx",
    exclude_duplicates=True
)
print("With exclusion:")
print("  Total individual reviews:", len(df_ind_2))
print("  Total students in Calc:", len(df_calc_2))
print("  Difference in reviews:", len(df_ind_1) - len(df_ind_2))

# Let's inspect a specific duplicate student to verify score change
camila_1 = df_calc_1[df_calc_1["Seleccione el nombre del Estudiante"].str.contains("CAMILA ESTEFANIA")]
camila_2 = df_calc_2[df_calc_2["Seleccione el nombre del Estudiante"].str.contains("CAMILA ESTEFANIA")]

print("\nCamila Estefania (Without Exclusion):")
if not camila_1.empty:
    print(f"  Reviews count: {camila_1.iloc[0]['Cuenta de Seleccione su nombre (Evaluador)']}")
    print(f"  Nota ponderada: {camila_1.iloc[0]['Nota ponderada']:.4f}")

print("Camila Estefania (With Exclusion):")
if not camila_2.empty:
    print(f"  Reviews count: {camila_2.iloc[0]['Cuenta de Seleccione su nombre (Evaluador)']}")
    print(f"  Nota ponderada: {camila_2.iloc[0]['Nota ponderada']:.4f}")

print("\n=== TESTING WRITTEN CAPSTONE ===")
df_ind_w, df_calc_w, _, _ = process_capstone_written(
    "data/EVALUACIÓN PROYECTO CAPSTONE - COHORTE 2(1-39).xlsx",
    "data/presentaciones_crcronograma.xlsx",
    exclude_duplicates=True
)
print("With exclusion:")
print("  Total individual reviews:", len(df_ind_w))
print("  Total students in Calc:", len(df_calc_w))
print("SUCCESS!")
