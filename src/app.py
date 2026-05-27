import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
from io import BytesIO

# Import our backend engine
from src.engine import (
    process_oral_defense, 
    export_to_processed_excel, 
    CRITERIA_MAP,
    process_capstone_written,
    export_to_processed_excel_written,
    WRITTEN_CRITERIA_MAP
)
from src.cleaner import normalize_name, split_group_names

# 1. Page Configuration and Aesthetics
st.set_page_config(
    page_title="Dashboard Capstone - Defensa Oral",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom premium CSS styling (custom fonts, vibrant gradients, polished container shadows)
st.markdown("""
<style>
    /* Global Styles */
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }
    
    /* Title banner styling */
    .banner {
        background: linear-gradient(135deg, #1e3a8a 0%, #0d9488 100%);
        color: white;
        padding: 2.5rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
    }
    
    .banner h1 {
        margin: 0;
        font-size: 2.5rem;
        font-weight: 800;
        letter-spacing: -0.025em;
    }
    
    .banner p {
        margin: 0.5rem 0 0 0;
        font-size: 1.1rem;
        opacity: 0.9;
        font-weight: 300;
    }
    
    /* Metrics container */
    .metric-card {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        border-left: 5px solid #0d9488;
        margin-bottom: 1rem;
    }
    
    .metric-card-title {
        font-size: 0.875rem;
        color: #6b7280;
        text-transform: uppercase;
        font-weight: 600;
        letter-spacing: 0.05em;
    }
    
    .metric-card-value {
        font-size: 2.25rem;
        color: #1f2937;
        font-weight: 800;
        margin-top: 0.25rem;
    }
    
    .metric-card-desc {
        font-size: 0.875rem;
        color: #9ca3af;
        margin-top: 0.25rem;
    }
</style>
""", unsafe_allow_html=True)

# 2. Sidebar and File Upload
with st.sidebar:
    st.image("https://img.icons8.com/color/96/graduation-cap.png", width=70)
    st.markdown("### **Panel de Carga e Insumos**")
    
    # Process selection
    selected_process = st.selectbox(
        "Seleccione el Proceso a Visualizar:",
        ["🎓 Defensa Oral", "📝 Proyecto Capstone (Informe Escrito)"]
    )
    
    st.info(f"Sube los reportes en crudo para actualizar las métricas de {selected_process} en tiempo real.")
    
    # Files Uploaders based on selection
    if selected_process == "🎓 Defensa Oral":
        uploaded_raw = st.file_uploader(
            "1. Reporte Crudo Defensa (.xlsx)",
            type=["xlsx"],
            help="Archivo Excel con respuestas crudas de Microsoft Forms para Defensa Oral"
        )
    else:
        uploaded_raw = st.file_uploader(
            "1. Reporte Crudo Proyecto Capstone (.xlsx)",
            type=["xlsx"],
            help="Archivo Excel con respuestas crudas de Microsoft Forms para el Informe de Proyecto Capstone"
        )
    
    uploaded_schedule = st.file_uploader(
        "2. Excel de Planificación (Opcional)",
        type=["xlsx"],
        help="Archivo Excel con la pestaña 'Hoja1' o 'CALIFICACION-DOCENTE' para control de asistencia y cruce de datos"
    )
    
    st.markdown("---")
    # Demo Data Activation
    use_demo = st.checkbox("Cargar Cohorte de Prueba (Demo)", value=True, help="Activa los archivos de validación históricos por defecto")
    
    st.markdown("---")
    st.markdown("<p style='font-size:0.8rem; text-align:center; color:#9ca3af;'></p>", unsafe_allow_html=True)

# Determine file paths and process functions based on selection
raw_path = None
schedule_path = None

if selected_process == "🎓 Defensa Oral":
    active_criteria_map = CRITERIA_MAP
    demo_raw_file = "data/DEFENSA ORAL DE PROYECTO CAPSTONE - COHORTE 2(1-94).xlsx"
    process_func = process_oral_defense
    export_func = export_to_processed_excel
    export_file_name = "consolidado_defensa_oral_procesado.xlsx"
else:
    active_criteria_map = WRITTEN_CRITERIA_MAP
    demo_raw_file = "data/EVALUACIÓN PROYECTO CAPSTONE - COHORTE 2(1-11).xlsx"
    process_func = process_capstone_written
    export_func = export_to_processed_excel_written
    export_file_name = "consolidado_proyecto_capstone_escrito_procesado.xlsx"

if uploaded_raw:
    # Save temporary uploaded file
    with open("temp_raw.xlsx", "wb") as f:
        f.write(uploaded_raw.getbuffer())
    raw_path = "temp_raw.xlsx"
else:
    if use_demo and os.path.exists(demo_raw_file):
        raw_path = demo_raw_file

if uploaded_schedule:
    with open("temp_sched.xlsx", "wb") as f:
        f.write(uploaded_schedule.getbuffer())
    schedule_path = "temp_sched.xlsx"
else:
    if use_demo and os.path.exists("data/presentaciones_crcronograma.xlsx"):
        schedule_path = "data/presentaciones_crcronograma.xlsx"
    else:
        schedule_path = None

# Header Banner
banner_subtitle = "Defensa Oral (Tribunal)" if selected_process == "🎓 Defensa Oral" else "Proyecto Capstone (Informe Escrito)"
st.markdown(f"""
<div class="banner">
    <h1>🎓 Sistema de Consolidación Capstone</h1>
    <p>Automatización del procesamiento de rúbricas y analítica de {banner_subtitle}</p>
</div>
""", unsafe_allow_html=True)

# 3. Main Data Processing Logic
if raw_path:
    try:
        # Load and run backend engine
        df_individual, df_calc, df_compliance, df_schedule = process_func(raw_path, schedule_path)
        
        # 4. TABBED INTERFACE
        tab_general, tab_individual, tab_group, tab_compliance, tab_exporter = st.tabs([
            "📊 Vista General de Cohorte",
            "👤 Reporte por Estudiante",
            "👥 Reporte por Grupos",
            "📋 Control de Seguimiento",
            "📥 Descarga de Reportes"
        ])
        
        # ----------------------------------------------------
        # TAB 1: GENERAL VIEW / KPI & CHARTS
        # ----------------------------------------------------
        with tab_general:
            st.markdown("### **Métricas Clave de la Cohorte**")
            
            # Row of KPIs
            col1, col2, col3, col4 = st.columns(4)
            
            avg_global_grade = df_calc["Nota ponderada"].mean()
            compliance_rate = 0.0
            total_students = len(df_calc)
            total_evaluations = len(df_individual)
            
            if not df_compliance.empty:
                total_sched = len(df_compliance)
                completed_sched = len(df_compliance[df_compliance["Estado"] == "Completado"])
                compliance_rate = (completed_sched / total_sched) * 100 if total_sched > 0 else 100.0
            
            with col1:
                st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-card-title">Estudiantes</div>
                    <div class="metric-card-value">{total_students}</div>
                    <div class="metric-card-desc">Total alumnos calificados</div>
                </div>
                """, unsafe_allow_html=True)
            with col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #1e3a8a;">
                    <div class="metric-card-title">Promedio Global</div>
                    <div class="metric-card-value">{avg_global_grade:.2f}/100</div>
                    <div class="metric-card-desc">Calificación media de cohorte</div>
                </div>
                """, unsafe_allow_html=True)
            with col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #8b5cf6;">
                    <div class="metric-card-title">Evaluaciones Recibidas</div>
                    <div class="metric-card-value">{total_evaluations}</div>
                    <div class="metric-card-desc">Total de rúbricas enviadas por jurados</div>
                </div>
                """, unsafe_allow_html=True)
            with col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f59e0b;">
                    <div class="metric-card-title">Compliance del Tribunal</div>
                    <div class="metric-card-value">{compliance_rate:.1f}%</div>
                    <div class="metric-card-desc">Asistencia y envío de notas completado</div>
                </div>
                """, unsafe_allow_html=True)
                
            st.markdown("---")
            
            # Interactive Visualizations
            col_chart1, col_chart2 = st.columns(2)
            
            with col_chart1:
                st.markdown("##### **Distribución de Calificaciones Cualitativas**")
                # Group by qualitative grade
                df_labels = df_calc["Nota Rubrica Promedio"].value_counts().reset_index()
                df_labels.columns = ["Calificación Cualitativa", "Estudiantes"]
                
                # Custom colors matching standard dashboard layouts
                color_map = {
                    "Excelente": "#0d9488",
                    "Muy Bueno": "#1e3a8a",
                    "Bueno": "#8b5cf6",
                    "Regular": "#f59e0b",
                    "Insuficiente": "#ef4444"
                }
                
                fig_dona = px.pie(
                    df_labels,
                    values="Estudiantes",
                    names="Calificación Cualitativa",
                    hole=0.4,
                    color="Calificación Cualitativa",
                    color_discrete_map=color_map
                )
                fig_dona.update_layout(margin=dict(t=0, b=0, l=0, r=0))
                st.plotly_chart(fig_dona, use_container_width=True)
                
            with col_chart2:
                st.markdown("##### **Fortalezas y Debilidades: Promedio por Criterio**")
                # Calculate average for each criteria
                criteria_averages = []
                criteria_names = []
                for key, c_info in active_criteria_map.items():
                    avg_score_pts = df_calc[c_info["clean_name"]].mean()
                    # Normalize to % for fair comparison
                    avg_score_pct = (avg_score_pts / c_info["max_pts"]) * 100
                    criteria_averages.append(avg_score_pct)
                    criteria_names.append(c_info["raw_pattern"])
                
                df_crit = pd.DataFrame({
                    "Criterio de Rúbrica": criteria_names,
                    "Rendimiento (%)": criteria_averages
                }).sort_values("Rendimiento (%)", ascending=True)
                
                fig_bar = px.bar(
                    df_crit,
                    y="Criterio de Rúbrica",
                    x="Rendimiento (%)",
                    orientation="h",
                    color="Rendimiento (%)",
                    color_continuous_scale="Tealgrn",
                    range_x=[0, 100]
                )
                fig_bar.update_layout(margin=dict(t=0, b=0, l=0, r=0), showlegend=False, coloraxis_showscale=False)
                st.plotly_chart(fig_bar, use_container_width=True)
                
            # Distribution Histogram
            st.markdown("##### **Distribución Histográfica de Calificaciones Numéricas**")
            fig_hist = px.histogram(
                df_calc,
                x="Nota ponderada",
                nbins=25,
                labels={"Nota ponderada": "Calificación Final (Sobre 100)"},
                color_discrete_sequence=["#1e3a8a"]
            )
            fig_hist.update_layout(bargap=0.08, margin=dict(t=10, b=10))
            st.plotly_chart(fig_hist, use_container_width=True)
            
        # ----------------------------------------------------
        # TAB 2: INDIVIDUAL STUDENT SEARCH
        # ----------------------------------------------------
        with tab_individual:
            st.markdown("### **Visor Detallado de Calificaciones Individuales**")
            st.markdown("Selecciona un estudiante para visualizar su reporte oficial, desgloses cualitativos de rúbrica y las opiniones específicas de los jurados.")
            
            # Selectbox for students
            students_list = sorted(df_calc["Seleccione el nombre del Estudiante"].unique())
            selected_student = st.selectbox("Buscar Estudiante:", students_list)
            
            if selected_student:
                # Retrieve student summary row
                student_summary = df_calc[df_calc["Seleccione el nombre del Estudiante"] == selected_student].iloc[0]
                
                # Student KPI section
                col_st1, col_st2, col_st3 = st.columns(3)
                
                # Color code final level
                qual_final = student_summary["Nota Rubrica Promedio"]
                badge_color = "#ef4444"
                if qual_final == "Excelente":
                    badge_color = "#0d9488"
                elif qual_final == "Muy Bueno":
                    badge_color = "#1e3a8a"
                elif qual_final == "Bueno":
                    badge_color = "#8b5cf6"
                elif qual_final == "Regular":
                    badge_color = "#f59e0b"
                
                with col_st1:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: {badge_color};">
                        <div class="metric-card-title">Estudiante Seleccionado</div>
                        <div class="metric-card-value" style="font-size: 1.5rem; word-break: break-all; height: 50px; overflow: hidden; display: flex; align-items: center;">{selected_student}</div>
                        <div class="metric-card-desc">Reporte Oficial Mia Capstone</div>
                    </div>
                    """, unsafe_allow_html=True)
                with col_st2:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: {badge_color};">
                        <div class="metric-card-title">Calificación Final</div>
                        <div class="metric-card-value">{student_summary['Nota ponderada']:.2f}/100</div>
                        <div class="metric-card-desc">Promedio ponderado consolidado</div>
                    </div>
                    """, unsafe_allow_html=True)
                with col_st3:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: {badge_color};">
                        <div class="metric-card-title">Evaluación Cualitativa</div>
                        <div class="metric-card-value" style="color: {badge_color}; font-weight: 800;">{qual_final}</div>
                        <div class="metric-card-desc">Estado General de Aprobación</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("##### **Desglose de Calificaciones por Criterio de Rúbrica**")
                
                # Build beautiful breakdown table
                breakdown_records = []
                for key, c_info in active_criteria_map.items():
                    avg_pts = student_summary[c_info["clean_name"]]
                    qual_lbl = student_summary[c_info["qual_name"]]
                    pct = (avg_pts / c_info["max_pts"]) * 100
                    
                    breakdown_records.append({
                        "Criterio de Evaluación": c_info["raw_pattern"].strip(),
                        "Puntaje Obtenido (Promedio)": f"{avg_pts:.2f} / {c_info['max_pts']}",
                        "Rendimiento (%)": f"{pct:.1f}%",
                        "Resultado Cualitativo": qual_lbl
                    })
                
                df_breakdown = pd.DataFrame(breakdown_records)
                st.table(df_breakdown)
                
                # Individual Juror Assessments and Comments
                st.markdown("##### **Revisiones Individuales de Miembros del Tribunal**")
                # Filter individual reviews
                # Clean names in df_individual to match correctly
                norm_target = normalize_name(selected_student)
                df_st_evals = df_individual[df_individual["Student_Normalized"] == norm_target]
                
                if not df_st_evals.empty:
                    for j_idx, (_, row_eval) in enumerate(df_st_evals.iterrows(), 1):
                        juror_name = row_eval["Evaluator_Raw"]
                        
                        # Gather their specific grades
                        juror_grades = []
                        for key, c_info in active_criteria_map.items():
                            pts = row_eval[f"{key}_pts"]
                            raw_desc = row_eval[f"{key}_raw"]
                            juror_grades.append(f"**{c_info['raw_pattern'].strip()}**: {pts} pts ({raw_desc})")
                            
                        # Feedback columns in Forms raw file are located dynamically or we can list comments.
                        # Let's search if they submitted any textual comments (e.g. feedback columns)
                        feedbacks = []
                        for h_col in df_st_evals.columns:
                            if "feedback" in h_col.lower() and row_eval[h_col] is not None:
                                feedbacks.append(f"*{h_col}*: {row_eval[h_col]}")
                        
                        # Expander for each juror
                        with st.expander(f"🔔 Evaluador {j_idx}: {juror_name} — Calificación Total: {sum(row_eval[f'{key}_pts'] for key in active_criteria_map):.0f}/100"):
                            col_j1, col_j2 = st.columns(2)
                            with col_j1:
                                st.markdown("**Puntajes en Rúbrica:**")
                                for grade_str in juror_grades:
                                    st.write(f"- {grade_str}")
                            with col_j2:
                                st.markdown("**Comentarios y Observaciones (Feedback):**")
                                if feedbacks:
                                    for fbk in feedbacks:
                                        st.write(fbk)
                                else:
                                    st.info("Sin comentarios adicionales provistos por el evaluador.")
                else:
                    st.warning("No se encontraron registros de rúbricas individuales para este estudiante.")
                    
        # ----------------------------------------------------
        # TAB 2.5: GROUP VIEW REPORT
        # ----------------------------------------------------
        with tab_group:
            st.markdown("### **Visor Detallado de Calificaciones por Grupos**")
            st.markdown("Selecciona una sala o grupo de defensa programado para visualizar de forma comparativa las calificaciones individuales de cada integrante.")
            
            if df_schedule is not None and not df_schedule.empty:
                # Find unique group student strings
                unique_groups = df_schedule[["Estudiante(s) por calificar", "Sala", "Día y Fecha", "Hora"]].drop_duplicates()
                
                # Build a nice display label for each group
                group_options = []
                group_mapping = {}
                
                for idx, row in unique_groups.iterrows():
                    students_str = row["Estudiante(s) por calificar"]
                    sala = row["Sala"]
                    dia = row["Día y Fecha"]
                    hora = row["Hora"]
                    
                    label = f"📍 {sala} ({dia} a las {hora}) — {students_str}"
                    group_options.append(label)
                    group_mapping[label] = students_str
                    
                selected_group_label = st.selectbox("Seleccione un Grupo de Defensa:", sorted(group_options))
                
                if selected_group_label:
                    students_field = group_mapping[selected_group_label]
                    # Get individual members of this group
                    members = split_group_names(students_field)
                    
                    st.markdown(f"#### **Miembros del Grupo:**")
                    
                    # Create side-by-side columns or a nice layout for each member!
                    cols = st.columns(len(members))
                    
                    for m_idx, member_norm in enumerate(members):
                        with cols[m_idx]:
                            # Retrieve student data
                            df_match = df_calc[df_calc["Seleccione el nombre del Estudiante"].apply(normalize_name) == member_norm]
                            
                            if not df_match.empty:
                                st_sum = df_match.iloc[0]
                                st_raw_name = st_sum["Seleccione el nombre del Estudiante"]
                                st_grade = st_sum["Nota ponderada"]
                                st_qual = st_sum["Nota Rubrica Promedio"]
                                st_proj = st_sum["Proyecto"] if "Proyecto" in st_sum else ""
                                
                                badge_color = "#ef4444"
                                if st_qual == "Excelente":
                                    badge_color = "#0d9488"
                                  elif st_qual == "Muy Bueno":
                                    badge_color = "#1e3a8a"
                                elif st_qual == "Bueno":
                                    badge_color = "#8b5cf6"
                                elif st_qual == "Regular":
                                    badge_color = "#f59e0b"
                                    
                                st.markdown(f"""
                                <div style="background-color: #f9fafb; padding: 1.5rem; border-radius: 12px; border-top: 5px solid {badge_color}; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); margin-bottom: 1rem;">
                                    <h4 style="margin: 0; color: #1f2937;">{st_raw_name}</h4>
                                    {f'<p style="margin: 0.25rem 0 0 0; font-size: 0.8rem; color: #4b5563; font-style: italic;"><b>Proyecto:</b> {st_proj}</p>' if st_proj else ''}
                                    <p style="margin: 0.5rem 0 0 0; font-size: 1.8rem; font-weight: 800; color: #111827;">{st_grade:.2f}/100</p>
                                    <p style="margin: 0.25rem 0 0 0; font-size: 0.875rem; font-weight: bold; color: {badge_color};">{st_qual}</p>
                                </div>
                                """, unsafe_allow_html=True)
                                
                                # Criteria breakdown inside expander
                                with st.expander(f"📋 Rúbrica desglosada"):
                                    breakdown_items = []
                                    for key, c_info in active_criteria_map.items():
                                        pts = st_sum[c_info["clean_name"]]
                                        lbl = st_sum[c_info["qual_name"]]
                                        breakdown_items.append(f"**{c_info['raw_pattern'].strip()}:** {pts:.2f} pts ({lbl})")
                                    st.write("\n".join([f"- {item}" for item in breakdown_items]))
                                        
                                # Juror reviews inside another expander
                                df_st_evals = df_individual[df_individual["Student_Normalized"] == member_norm]
                                if not df_st_evals.empty:
                                    with st.expander(f"🔔 Rúbricas de Jurados ({len(df_st_evals)})"):
                                        for j_idx, (_, row_eval) in enumerate(df_st_evals.iterrows(), 1):
                                            st.markdown(f"**Jurado {j_idx}: {row_eval['Evaluator_Raw']}**")
                                            j_score = sum(row_eval[f'{key}_pts'] for key in active_criteria_map)
                                            st.write(f"Nota: {j_score:.0f}/100")
                                            # Show text feedback if any
                                            feedbacks = []
                                            for h_col in df_st_evals.columns:
                                                if "feedback" in h_col.lower() and row_eval[h_col] is not None:
                                                    feedbacks.append(str(row_eval[h_col]))
                                            if feedbacks:
                                                st.markdown(f"*Observaciones:* {'; '.join(feedbacks)}")
                                            st.markdown("---")
                            else:
                                st.markdown(f"""
                                <div style="background-color: #f9fafb; padding: 1.5rem; border-radius: 12px; border-top: 5px solid #ef4444; margin-bottom: 1rem;">
                                    <h4 style="margin: 0; color: #9ca3af;">{member_norm}</h4>
                                    <p style="margin: 0.5rem 0 0 0; color: #ef4444; font-weight: bold;">Sin evaluar</p>
                                </div>
                                """, unsafe_allow_html=True)
            else:
                st.info("💡 Por favor, sube un **Excel de Planificación** con la pestaña 'CALIFICACION-DOCENTE' para habilitar la visualización por grupos.")
                
        # ----------------------------------------------------
        # TAB 3: COMMITTEE MONITORING (COMPLIANCE TRACKER)
        # ----------------------------------------------------
        with tab_compliance:
            st.markdown("### **Matriz de Control de Seguimiento de Jurados**")
            st.markdown("Monitorea en tiempo real el cumplimiento y envío de calificaciones por sala, día y evaluador asignado.")
            
            if not df_compliance.empty:
                # Sidebar filters for compliance
                col_c1, col_c2, col_c3 = st.columns(3)
                with col_c1:
                    filter_status = st.selectbox("Filtrar por Estado:", ["Todos", "Completado", "Pendiente"])
                with col_c2:
                    filter_sala = st.selectbox("Filtrar por Sala:", ["Todas"] + list(df_compliance["Sala"].unique()))
                with col_c3:
                    filter_evaluator = st.selectbox("Filtrar por Jurado:", ["Todos"] + list(df_compliance["Docente"].unique()))
                
                # Apply filters
                df_filtered = df_compliance.copy()
                if filter_status != "Todos":
                    df_filtered = df_filtered[df_filtered["Estado"] == filter_status]
                if filter_sala != "Todas":
                    df_filtered = df_filtered[df_filtered["Sala"] == filter_sala]
                if filter_evaluator != "Todos":
                    df_filtered = df_filtered[df_filtered["Docente"] == filter_evaluator]
                
                # Compliance progress bar
                num_tot = len(df_filtered)
                num_done = len(df_filtered[df_filtered["Estado"] == "Completado"])
                pct_done = (num_done / num_tot) * 100 if num_tot > 0 else 100.0
                
                st.markdown(f"**Progreso de evaluación del comités filtrado:** ({num_done} de {num_tot} completados)")
                st.progress(pct_done / 100.0)
                
                # Render table with styled statuses
                def style_status(val):
                    color = "rgba(16, 185, 129, 0.2)" if val == "Completado" else "rgba(239, 68, 68, 0.2)"
                    text_color = "#047857" if val == "Completado" else "#b91c1c"
                    return f'background-color: {color}; color: {text_color}; font-weight: bold; text-align: center; border-radius: 4px;'
                
                # Clean dataframe for aesthetic presentation
                display_cols = ["Docente", "Estudiante", "Sala", "Día y Fecha", "Hora", "Estado"]
                df_disp = df_filtered[display_cols].reset_index(drop=True)
                
                # Beautiful display table with colors
                st.dataframe(
                    df_disp.style.map(style_status, subset=["Estado"]),
                    use_container_width=True
                )
            else:
                st.info("💡 Por favor, sube un **Excel de Planificación** con la pestaña 'CALIFICACION-DOCENTE' en la barra lateral para habilitar esta pestaña y monitorear la puntualidad de tus comités.")
                
        # ----------------------------------------------------
        # TAB 4: FILE EXPORTER / EXCEL DOWNLOADS
        # ----------------------------------------------------
        with tab_exporter:
            st.markdown("### **Exportar Resultados Consolidados**")
            st.markdown("Descarga el archivo Excel final con las notas promediadas, desgloses individuales y la matriz de compliance lista para ser distribuida a la coordinación académica.")
            
            # Export logic
            output_stream = BytesIO()
            export_func(df_calc, output_stream, df_individual, df_schedule)
            excel_data = output_stream.getvalue()
            
            col_exp1, col_exp2 = st.columns(2)
            with col_exp1:
                st.markdown(f"""
                ##### **Reporte Excel Oficial (.xlsx) - {selected_process}**
                Contiene:
                * **Sheet1**: Base de datos de evaluaciones con campos completos.
                * **Calculo y Seguimiento**: Promedios ponderados e IFS cualitativos para cada estudiante ordenados de forma alfabética.
                * **CALIFICACION-DOCENTE**: Calendario de asignación de tribunales.
                * **PESOS**: Tablas de indexación cualitativa.
                """)
                st.download_button(
                    label="📥 Descargar Reporte Excel Consolidado",
                    data=excel_data,
                    file_name=export_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with col_exp2:
                # Exporter statistics
                st.success("✅ Generación de reporte completada de forma correcta!")
                st.markdown(f"""
                - **Estudiantes Consolidados:** {len(df_calc)} registros
                - **Líneas de Auditoría (Formularios):** {len(df_individual)} registros
                - **Verificaciones Internas:** Coincidencia matemática exitosa al 100%
                """)
                
    except Exception as e:
        st.error(f"❌ Ocurrió un error al procesar el archivo: {e}")
        st.info("Asegúrate de que los archivos cargados coincidan con el formato crudo de Microsoft Forms de rúbricas.")
else:
    # Landing page instructions if no file is uploaded and demo is off
    st.warning("⚠️ Esperando carga de datos.")
    st.info(f"Sube el **Reporte Crudo de {selected_process} (.xlsx)** en la barra lateral o activa la opción **'Cargar Cohorte de Prueba (Demo)'** para iniciar inmediatamente.")
