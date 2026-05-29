import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import pandas as pd
import openpyxl
from src.engine import process_oral_defense, CRITERIA_MAP

def run_blackbox_test():
    print("====================================================")
    print("EJECUTANDO PRUEBA DE CAJA NEGRA (COHORTE 2)")
    print("====================================================")
    
    raw_excel = "data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2_PROCESADO.xlsx"
    processed_excel = "data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2_PROCESADO.xlsx"
    schedule_excel = "data/presentaciones_crcronograma.xlsx"
    
    print(f"\n---> Procesando '{raw_excel}' con cronograma...")
    # By default, to match the original cohort 2 processed sheet, we should probably set exclude_duplicates=False or True depending on what that sheet did.
    # Let's run both and see which one matches the processed file!
    for excl in [False, True]:
        print(f"\n--- Probando con exclude_duplicates={excl} ---")
        df_ind, df_calc, df_comp, _ = process_oral_defense(raw_excel, schedule_excel, exclude_duplicates=excl)
        df_calc.columns = [str(c).strip() for c in df_calc.columns]
        
        # Load historical processed data
        df_calc_hist = pd.read_excel(processed_excel, sheet_name="Calculo")
        df_calc_hist.columns = [str(c).strip() for c in df_calc_hist.columns]
        df_calc_hist = df_calc_hist[df_calc_hist["Seleccione el nombre del Estudiante"].notna()]
        df_calc_hist = df_calc_hist[df_calc_hist["Seleccione el nombre del Estudiante"].str.upper() != "TOTAL GENERAL"]
        
        df_calc["student_key"] = df_calc["Seleccione el nombre del Estudiante"].str.strip().str.upper()
        df_calc_hist["student_key"] = df_calc_hist["Seleccione el nombre del Estudiante"].str.strip().str.upper()
        
        ours_indexed = df_calc.set_index("student_key")
        hist_indexed = df_calc_hist.set_index("student_key")
        
        mismatches = 0
        checked_students = 0
        
        columns_to_compare = [
            "Cuenta de Seleccione su nombre (Evaluador)",
            "Nota ponderada",
        ]
        for key, c_info in CRITERIA_MAP.items():
            columns_to_compare.append(c_info["clean_name"].strip())
            columns_to_compare.append(c_info["qual_name"].strip())
            
        for s_key in ours_indexed.index:
            if s_key in hist_indexed.index:
                checked_students += 1
                row_ours = ours_indexed.loc[s_key]
                row_hist = hist_indexed.loc[s_key]
                student_name = row_ours["Seleccione el nombre del Estudiante"]
                
                for col in columns_to_compare:
                    if col not in row_ours or col not in row_hist:
                        continue
                    val_ours = row_ours[col]
                    val_hist = row_hist[col]
                    
                    if isinstance(val_ours, (int, float)) and isinstance(val_hist, (int, float)):
                        if abs(val_ours - val_hist) > 1e-5:
                            # Print only top 10 mismatches to not flood
                            if mismatches < 10:
                                print(f"   [ERROR] Mismatch para {student_name} en '{col}': Nuestro={val_ours:.4f}, Histórico={val_hist:.4f}")
                            mismatches += 1
                    else:
                        if str(val_ours).strip().lower() != str(val_hist).strip().lower():
                            if mismatches < 10:
                                print(f"   [ERROR] Mismatch para {student_name} en '{col}': Nuestro='{val_ours}', Histórico='{val_hist}'")
                            mismatches += 1
                            
        print(f"   Alumnos validados: {checked_students}, Diferencias: {mismatches}")
        if mismatches == 0:
            print(f"🎉 ¡COINCIDENCIA PERFECTA CON EXCLUDE_DUPLICATES={excl}!")

if __name__ == "__main__":
    run_blackbox_test()
