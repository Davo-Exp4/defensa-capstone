import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import pandas as pd
import openpyxl
from src.engine import process_oral_defense, CRITERIA_MAP

def run_blackbox_test():
    print("====================================================")
    print("EJECUTANDO PRUEBA DE CAJA NEGRA (GOLD STANDARD)")
    print("====================================================")
    
    raw_excel = "data/cohorte_pasada_crudo.xlsx"
    processed_excel = "data/cohorte_pasada_procesado.xlsx"
    
    # Test A: Processing the completed Sheet1 inside processed_excel (which has all 556 evaluations)
    print(f"\n---> TEST A: Procesando la hoja Sheet1 de '{processed_excel}' (Base de datos completa)")
    df_ind_a, df_calc_a, df_comp_a, _ = process_oral_defense(processed_excel, processed_excel)
    df_calc_a.columns = [str(c).strip() for c in df_calc_a.columns]
    
    # Load historical processed data (Calculo sheet, header is on row 3 -> header=2)
    df_calc_hist = pd.read_excel(processed_excel, sheet_name="Calculo", header=2)
    df_calc_hist.columns = [str(c).strip() for c in df_calc_hist.columns]
    df_calc_hist = df_calc_hist[df_calc_hist["Seleccione el nombre del Estudiante"].notna()]
    df_calc_hist = df_calc_hist[df_calc_hist["Seleccione el nombre del Estudiante"].str.upper() != "TOTAL GENERAL"]
    
    df_calc_a["student_key"] = df_calc_a["Seleccione el nombre del Estudiante"].str.strip().str.upper()
    df_calc_hist["student_key"] = df_calc_hist["Seleccione el nombre del Estudiante"].str.strip().str.upper()
    
    ours_indexed_a = df_calc_a.set_index("student_key")
    hist_indexed = df_calc_hist.set_index("student_key")
    
    mismatches_a = 0
    checked_students_a = 0
    
    columns_to_compare = [
        "Cuenta de Seleccione su nombre (Evaluador)",
        "Nota ponderada",
    ]
    for key, c_info in CRITERIA_MAP.items():
        columns_to_compare.append(c_info["clean_name"].strip())
        columns_to_compare.append(c_info["qual_name"].strip())
        
    for s_key in ours_indexed_a.index:
        if s_key in hist_indexed.index:
            checked_students_a += 1
            row_ours = ours_indexed_a.loc[s_key]
            row_hist = hist_indexed.loc[s_key]
            student_name = row_ours["Seleccione el nombre del Estudiante"]
            
            for col in columns_to_compare:
                val_ours = row_ours[col]
                val_hist = row_hist[col]
                
                if isinstance(val_ours, (int, float)) and isinstance(val_hist, (int, float)):
                    if abs(val_ours - val_hist) > 1e-5:
                        print(f"   [ERROR] TEST A: Mismatch para {student_name} en '{col}': Nuestro={val_ours:.4f}, Histórico={val_hist:.4f}")
                        mismatches_a += 1
                else:
                    if str(val_ours).strip().lower() != str(val_hist).strip().lower():
                        print(f"   [ERROR] TEST A: Mismatch para {student_name} en '{col}': Nuestro='{val_ours}', Histórico='{val_hist}'")
                        mismatches_a += 1
                        
    # Compare note final label
    ours_rubrica_col = "Nota Rubrica Promedio"
    hist_rubrica_col = None
    for col in hist_indexed.columns:
        if "rubrica" in col.lower() and "promedio" in col.lower():
            hist_rubrica_col = col
            break
    if not hist_rubrica_col:
        for col in hist_indexed.columns:
            if "rubrica" in col.lower() or "nota rubrica" in col.lower():
                hist_rubrica_col = col
                break
                
    if hist_rubrica_col:
        for s_key in ours_indexed_a.index:
            if s_key in hist_indexed.index:
                val_ours = ours_indexed_a.loc[s_key, ours_rubrica_col]
                val_hist = hist_indexed.loc[s_key, hist_rubrica_col]
                if str(val_ours).strip().lower() != str(val_hist).strip().lower():
                    print(f"   [ERROR] TEST A: Mismatch de nota final para {ours_indexed_a.loc[s_key, 'Seleccione el nombre del Estudiante']}: Nuestro='{val_ours}', Histórico='{val_hist}'")
                    mismatches_a += 1

    print(f"   TEST A COMPLETADO. Alumnos validados: {checked_students_a}, Diferencias: {mismatches_a}")

    # Test B: Processing raw_excel (subset, filtering comparisons for matching evaluation count only)
    print(f"\n---> TEST B: Procesando '{raw_excel}' (Base de datos cruda parcial)")
    df_ind_b, df_calc_b, df_comp_b, _ = process_oral_defense(raw_excel, processed_excel)
    df_calc_b.columns = [str(c).strip() for c in df_calc_b.columns]
    df_calc_b["student_key"] = df_calc_b["Seleccione el nombre del Estudiante"].str.strip().str.upper()
    ours_indexed_b = df_calc_b.set_index("student_key")
    
    mismatches_b = 0
    checked_students_b = 0
    
    for s_key in ours_indexed_b.index:
        if s_key in hist_indexed.index:
            count_ours = ours_indexed_b.loc[s_key, "Cuenta de Seleccione su nombre (Evaluador)"]
            count_hist = hist_indexed.loc[s_key, "Cuenta de Seleccione su nombre (Evaluador)"]
            
            if count_ours == count_hist:
                checked_students_b += 1
                row_ours = ours_indexed_b.loc[s_key]
                row_hist = hist_indexed.loc[s_key]
                student_name = row_ours["Seleccione el nombre del Estudiante"]
                
                for col in columns_to_compare:
                    val_ours = row_ours[col]
                    val_hist = row_hist[col]
                    
                    if isinstance(val_ours, (int, float)) and isinstance(val_hist, (int, float)):
                        if abs(val_ours - val_hist) > 1e-5:
                            print(f"   [ERROR] TEST B: Mismatch para {student_name} en '{col}': Nuestro={val_ours:.4f}, Histórico={val_hist:.4f}")
                            mismatches_b += 1
                    else:
                        if str(val_ours).strip().lower() != str(val_hist).strip().lower():
                            print(f"   [ERROR] TEST B: Mismatch para {student_name} en '{col}': Nuestro='{val_ours}', Histórico='{val_hist}'")
                            mismatches_b += 1
                            
                if hist_rubrica_col:
                    val_ours = ours_indexed_b.loc[s_key, ours_rubrica_col]
                    val_hist = hist_indexed.loc[s_key, hist_rubrica_col]
                    if str(val_ours).strip().lower() != str(val_hist).strip().lower():
                        print(f"   [ERROR] TEST B: Mismatch de nota final para {student_name}: Nuestro='{val_ours}', Histórico='{val_hist}'")
                        mismatches_b += 1

    print(f"   TEST B COMPLETADO. Alumnos validados (con igual jurado): {checked_students_b}, Diferencias: {mismatches_b}")

    print("\n====================================================")
    print("RESUMEN DE RESULTADOS DE CAJA NEGRA")
    print("====================================================")
    print(f"TEST A: {mismatches_a} discrepancias (Base de datos completa).")
    print(f"TEST B: {mismatches_b} discrepancias (Mismos evaluadores).")
    
    if mismatches_a == 0 and mismatches_b == 0:
        print("\n🎉 ¡EXITO! El 100% de los cálculos coincide matemáticamente con los datos históricos de validación en ambos escenarios.")
        return True
    else:
        print("\n❌ FALLIDO: Se encontraron discrepancias. Revisar lógica en src/engine.py.")
        return False

if __name__ == "__main__":
    run_blackbox_test()
