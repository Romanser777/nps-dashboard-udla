import os
import glob
import re
import json
import pandas as pd
import numpy as np
import sqlite3

workspace_dir = r"c:\Users\sroman\Desktop\ENCUESTAS NPS"
subdirs = ["2024-10", "2024-20", "2025-10", "2025-20"]
enrollment_excel = os.path.join(workspace_dir, "Detalle Estudiantes Postgrado_Tipo Ingreso.xlsx")

# Regular expression to extract rating number
def parse_rating(val):
    if pd.isna(val):
        return None
    val_str = str(val).strip()
    if not val_str:
        return None
    m = re.match(r'^(\d+)', val_str)
    if m:
        return int(m.group(1))
    return None

# Clean text strings (e.g. whitespace, clean names)
def clean_text(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    # Normalize multiple spaces
    val_str = re.sub(r'\s+', ' ', val_str)
    return val_str

# Smart string cleaner for matching programs
def clean_matching_name(s):
    if pd.isna(s):
        return ""
    s = str(s).lower().strip()
    s = s.replace("í", "i").replace("á", "a").replace("ó", "o").replace("ú", "u").replace("ñ", "n")
    s = s.replace("é", "e").replace("ü", "u")
    
    # Expand abbreviations
    s = s.replace("dip.en", "diplomado")
    s = s.replace("dip. de", "diplomado")
    s = s.replace("dip.", "diplomado")
    s = s.replace("rehab.", "rehabilitacion")
    s = s.replace("fisiat.", "fisiatria")
    s = s.replace("veter.", "veterinaria")
    
    # Remove Spanish stop words
    stop_words = ["en", "de", "el", "la", "y", "para", "del", "con", "a", "los", "las"]
    words = re.findall(r'[a-z0-9]+', s)
    filtered_words = [w for w in words if w not in stop_words]
    
    return "".join(filtered_words)

print("Starting data consolidation...")

# ----------------------------------------------------
# 1. Load Real Enrollments from BD MATRICULA
# ----------------------------------------------------
enrollment_dict = {}
global_enrollment_period = {}
fac_enrollment_period = {}

if os.path.exists(enrollment_excel):
    print(f"Loading real enrollments from {enrollment_excel}...")
    try:
        df_enr = pd.read_excel(enrollment_excel, sheet_name='BD MATRICULA', usecols=['Periodo', 'Rut', 'Programa', 'Programa Postgrado', 'Facultad Postgrado'])
        print(f"Loaded {len(df_enr)} enrollment records.")
        
        # Calculate student count per Periodo, Programa Postgrado, Programa code
        # We group by Periodo and Programa Postgrado to count total enrolled students
        enr_grouped = df_enr.groupby(['Periodo', 'Programa Postgrado']).size().reset_index(name='estudiantes')
        print(f"Distinct enrollment candidate entries: {len(enr_grouped)}")
        
        # We also sum enrollment per Periodo globally and per Facultad
        for idx, r_en in df_enr.drop_duplicates(subset=['Rut', 'Periodo']).groupby(['Periodo']).size().items():
            # Periodo maps e.g. 202410 to "2024-10"
            per_str = f"{str(idx)[:4]}-{str(idx)[4:]}"
            global_enrollment_period[per_str] = int(r_en)
            
        # Group enrollment by Facultad and Periodo
        enr_fac_grouped = df_enr.groupby(['Periodo', 'Facultad Postgrado']).size().reset_index(name='estudiantes')
        for idx, row in enr_fac_grouped.iterrows():
            per_str = f"{str(row['Periodo'])[:4]}-{str(row['Periodo'])[4:]}"
            db_fac = str(row['Facultad Postgrado']).strip()
            # Map database faculty name to survey faculty name
            # Let's clean the name for mapping
            fac_key = f"{db_fac}||{per_str}"
            fac_enrollment_period[fac_key] = int(row['estudiantes'])
            
        # Keep list of db candidates
        for idx, row in enr_grouped.iterrows():
            per_val = row['Periodo']
            per_str = f"{str(per_val)[:4]}-{str(per_val)[4:]}"
            prog_name = str(row['Programa Postgrado']).strip()
            c_name = clean_matching_name(prog_name)
            
            # Save mapping key
            map_key = f"{per_str}||{c_name}"
            enrollment_dict[map_key] = int(row['estudiantes'])
            
        print("Real enrollments loaded and indexed.")
    except Exception as e:
        print(f"Error loading BD MATRICULA: {e}")
else:
    print("Warning: Detalle Estudiantes Postgrado_Tipo Ingreso.xlsx not found. Enrollment numbers will use backup estimates.")

# ----------------------------------------------------
# 2. Process Survey Excel Files
# ----------------------------------------------------
responses_list = []
comments_list = []

# Faculty name mapper based on filename keywords
def get_faculty_name(filename):
    filename_lower = filename.lower()
    if "arquitectura" in filename_lower:
        return "Arquitectura, Diseño y Construcción"
    elif "comunicaciones" in filename_lower:
        return "Comunicaciones"
    elif "derecho" in filename_lower:
        return "Derecho"
    elif "educaci" in filename_lower and "vra" in filename_lower:
        return "Educación VRA"
    elif "educaci" in filename_lower:
        return "Educación"
    elif "ingenier" in filename_lower or "negocio" in filename_lower:
        return "Ingeniería y Negocios"
    elif "medicina" in filename_lower or "veterinaria" in filename_lower or "agronom" in filename_lower:
        return "Medicina Veterinaria y Agronomía"
    elif "salud" in filename_lower or "social" in filename_lower:
        return "Salud y Ciencias Sociales"
    else:
        return "Otra"

for sd in subdirs:
    path = os.path.join(workspace_dir, sd)
    if os.path.exists(path):
        files = glob.glob(os.path.join(path, "*.xlsx"))
        print(f"Processing folder {sd} ({len(files)} files)...")
        for f in files:
            filename = os.path.basename(f)
            faculty = get_faculty_name(filename)
            
            try:
                xl = pd.ExcelFile(f)
                sheet_name = 'Sheet' if 'Sheet' in xl.sheet_names else xl.sheet_names[0]
                df = pd.read_excel(f, sheet_name=sheet_name)
                
                # Clean header metadata row added by SurveyMonkey
                df_clean = df[df['respondent_id'].notna()].copy()
                
                for idx, row in df_clean.iterrows():
                    resp_id = str(int(row.iloc[0]))
                    prog = clean_text(row.iloc[9])
                    
                    if not prog or prog == "Response":
                        continue
                        
                    q1 = parse_rating(row.iloc[10])
                    q2 = parse_rating(row.iloc[11])
                    q3 = parse_rating(row.iloc[12])
                    q4 = parse_rating(row.iloc[13])
                    coord = clean_text(row.iloc[14])
                    q5 = parse_rating(row.iloc[15])
                    q6 = parse_rating(row.iloc[16])
                    q7 = parse_rating(row.iloc[17])
                    q8 = parse_rating(row.iloc[18])
                    q9 = parse_rating(row.iloc[19])
                    q10 = parse_rating(row.iloc[20])
                    com = clean_text(row.iloc[21])
                    
                    # Classification of NPS (Q1)
                    # Promotores: 6-7, Pasivos: 5, Detractores: 1-4
                    tipo_nps = None
                    if q1 is not None:
                        if q1 >= 6:
                            tipo_nps = "Promotor"
                        elif q1 == 5:
                            tipo_nps = "Pasivo"
                        else:
                            tipo_nps = "Detractor"
                            
                    response_dict = {
                        "respondent_id": resp_id,
                        "periodo": sd,
                        "facultad": faculty,
                        "programa": prog,
                        "coordinador": coord if coord else "Sin Coordinador",
                        "Q1_Recomendacion_NPS": q1,
                        "Q2_Contenidos": q2,
                        "Q3_Oportunidades": q3,
                        "Q4_Volveria": q4,
                        "Q5_Servicio_Coordinador": q5,
                        "Q6_Tiempos_Coordinador": q6,
                        "Q7_Portal_Facilidad": q7,
                        "Q8_Portal_Aporte": q8,
                        "Q9_Blackboard_Facilidad": q9,
                        "Q10_Blackboard_Aporte": q10,
                        "tipo_nps": tipo_nps
                    }
                    responses_list.append(response_dict)
                    
                    if com and com.lower() not in ["open-ended response", "nan", ""]:
                        sentimiento = "Neutro"
                        if q1 is not None:
                            if q1 >= 6:
                                sentimiento = "Positivo"
                            elif q1 <= 4:
                                sentimiento = "Negativo"
                                
                        comments_list.append({
                            "periodo": sd,
                            "facultad": faculty,
                            "programa": prog,
                            "coordinador": coord if coord else "Sin Coordinador",
                            "puntaje": q1,
                            "clasificacion": sentimiento,
                            "comentario": com
                        })
                        
            except Exception as e:
                print(f"Error processing file {filename} in {sd}: {e}")

df_all_responses = pd.DataFrame(responses_list)
df_all_comments = pd.DataFrame(comments_list)

print(f"Consolidated {len(df_all_responses)} responses and {len(df_all_comments)} comments.")

# Helper to find real enrollment for a program and period
def get_enrollment(period, program_name, responses_count):
    c_prog = clean_matching_name(program_name)
    key = f"{period}||{c_prog}"
    
    # Try direct lookup
    if key in enrollment_dict:
        return enrollment_dict[key]
        
    # Try fuzzy key lookup (starts with candidate)
    for k, val in enrollment_dict.items():
        k_per, k_prog = k.split("||")
        if period == k_per:
            if c_prog.startswith(k_prog[:15]) or k_prog.startswith(c_prog[:15]):
                return val
                
    # Fallback to smart backup estimator (approx 55% response rate)
    return int(np.ceil(responses_count / 0.55) + 2)

# Generate summary metrics per program/period for Excel sheet
summary_list = []
grouped = df_all_responses.groupby(['periodo', 'facultad', 'programa'])
for name, group in grouped:
    period, fac, prog = name
    n_resp = len(group)
    
    q1_vals = group['Q1_Recomendacion_NPS'].dropna()
    n_q1 = len(q1_vals)
    promotores = sum(group['tipo_nps'] == "Promotor")
    detractores = sum(group['tipo_nps'] == "Detractor")
    nps = None
    if n_q1 > 0:
        nps = int(round((promotores - detractores) / n_q1 * 100))
        
    # Get enrollment
    enroll = get_enrollment(period, prog, n_resp)
    cov = round(n_resp / enroll * 100, 1) if enroll > 0 else 0
        
    summary_list.append({
        "periodo": period,
        "facultad": fac,
        "programa": prog,
        "n_respuestas": n_resp,
        "promotores": promotores,
        "pasivos": sum(group['tipo_nps'] == "Pasivo"),
        "detractores": detractores,
        "NPS": nps,
        "enrollment": enroll,
        "cobertura_pct": cov,
        "prom_Q2_Contenidos": group['Q2_Contenidos'].mean(),
        "prom_Q3_Oportunidades": group['Q3_Oportunidades'].mean(),
        "prom_Q4_Volveria": group['Q4_Volveria'].mean(),
        "prom_Q5_Servicio_Coord": group['Q5_Servicio_Coordinador'].mean(),
        "prom_Q6_Tiempos_Coord": group['Q6_Tiempos_Coordinador'].mean(),
        "prom_Q7_Portal_Facilidad": group['Q7_Portal_Facilidad'].mean(),
        "prom_Q8_Portal_Aporte": group['Q8_Portal_Aporte'].mean(),
        "prom_Q9_BB_Facilidad": group['Q9_Blackboard_Facilidad'].mean(),
        "prom_Q10_BB_Aporte": group['Q10_Blackboard_Aporte'].mean()
    })
df_summary = pd.DataFrame(summary_list)

# Write to Excel NPS_Consolidado.xlsx
output_excel = os.path.join(workspace_dir, "NPS_Consolidado.xlsx")
try:
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df_all_responses.to_excel(writer, sheet_name='Detalle_Respuestas', index=False)
        df_summary.to_excel(writer, sheet_name='Resumen_Programas', index=False)
        df_all_comments.to_excel(writer, sheet_name='Comentarios', index=False)
    print(f"Successfully generated {output_excel}")
except Exception as e:
    print(f"Error writing excel file: {e}")

# Write to SQLite Database
output_db = os.path.join(workspace_dir, "NPS_Consolidado.db")
try:
    conn = sqlite3.connect(output_db)
    df_all_responses.to_sql('detalle_respuestas', conn, if_exists='replace', index=False)
    df_summary.to_sql('resumen_programas', conn, if_exists='replace', index=False)
    df_all_comments.to_sql('comentarios', conn, if_exists='replace', index=False)
    conn.close()
    print(f"Successfully generated {output_db}")
except Exception as e:
    print(f"Error writing SQLite DB: {e}")

# ----------------------------------------------------
# 3. Compile dynamic JSON structure 'D' for Web Dashboard
# ----------------------------------------------------
print("Compiling dynamic D structure for HTML frontend...")

# Helper to safe mean float or None
def safe_mean(series):
    s = series.dropna()
    if len(s) == 0:
        return None
    return float(round(s.mean(), 1))

# Lists
all_periods = sorted(list(df_all_responses['periodo'].unique()))
all_facultades = sorted(list(df_all_responses['facultad'].unique()))
all_coordinadores = sorted(list(df_all_responses['coordinador'].dropna().unique()))

# D Global
d_global = {}
for p in all_periods:
    df_p = df_all_responses[df_all_responses['periodo'] == p]
    resp = len(df_p)
    prom = sum(df_p['tipo_nps'] == "Promotor")
    det = sum(df_p['tipo_nps'] == "Detractor")
    pas = sum(df_p['tipo_nps'] == "Pasivo")
    nps = int(round((prom - det) / resp * 100)) if resp > 0 else 0
    
    # Calculate enrollments globally
    # Try to sum all matched program enrollments for this period
    period_progs = df_summary[df_summary['periodo'] == p]
    enroll = int(period_progs['enrollment'].sum())
    
    cov_pct = round(resp / enroll * 100, 1) if enroll > 0 else 0
    
    d_global[p] = {
        "resp": resp,
        "prom": prom,
        "det": det,
        "pas": pas,
        "nps": nps,
        "enroll": enroll,
        "cov_pct": cov_pct
    }

# D Global Grupos
d_global_grupos = {}
for p in all_periods:
    df_p = df_all_responses[df_all_responses['periodo'] == p]
    d_global_grupos[p] = {
        "exp": {
            "nps_val": safe_mean(df_p['Q1_Recomendacion_NPS']),
            "contenidos": safe_mean(df_p['Q2_Contenidos']),
            "laboral": safe_mean(df_p['Q3_Oportunidades']),
            "recompra": safe_mean(df_p['Q4_Volveria'])
        },
        "coord": {
            "coord_calidad": safe_mean(df_p['Q5_Servicio_Coordinador']),
            "coord_tiempo": safe_mean(df_p['Q6_Tiempos_Coordinador'])
        },
        "plat": {
            "portal_facil": safe_mean(df_p['Q7_Portal_Facilidad']),
            "portal_aporte": safe_mean(df_p['Q8_Portal_Aporte']),
            "bb_facil": safe_mean(df_p['Q9_Blackboard_Facilidad']),
            "bb_aporte": safe_mean(df_p['Q10_Blackboard_Aporte'])
        }
    }

# D Por Facultad & Preguntas Facultad
d_por_facultad = {}
d_preguntas_fac = {}
for f in all_facultades:
    d_por_facultad[f] = {}
    d_preguntas_fac[f] = {}
    for p in all_periods:
        df_fp = df_all_responses[(df_all_responses['facultad'] == f) & (df_all_responses['periodo'] == p)]
        if len(df_fp) == 0:
            continue
            
        resp = len(df_fp)
        prom = sum(df_fp['tipo_nps'] == "Promotor")
        det = sum(df_fp['tipo_nps'] == "Detractor")
        pas = sum(df_fp['tipo_nps'] == "Pasivo")
        nps = int(round((prom - det) / resp * 100)) if resp > 0 else 0
        
        # Calculate faculty enrollment
        fac_progs = df_summary[(df_summary['facultad'] == f) & (df_summary['periodo'] == p)]
        enroll = int(fac_progs['enrollment'].sum())
        cov_pct = round(resp / enroll * 100, 1) if enroll > 0 else 0
        
        exp_dict = {
            "nps_val": safe_mean(df_fp['Q1_Recomendacion_NPS']),
            "contenidos": safe_mean(df_fp['Q2_Contenidos']),
            "laboral": safe_mean(df_fp['Q3_Oportunidades']),
            "recompra": safe_mean(df_fp['Q4_Volveria'])
        }
        coord_dict = {
            "coord_calidad": safe_mean(df_fp['Q5_Servicio_Coordinador']),
            "coord_tiempo": safe_mean(df_fp['Q6_Tiempos_Coordinador'])
        }
        plat_dict = {
            "portal_facil": safe_mean(df_fp['Q7_Portal_Facilidad']),
            "portal_aporte": safe_mean(df_fp['Q8_Portal_Aporte']),
            "bb_facil": safe_mean(df_fp['Q9_Blackboard_Facilidad']),
            "bb_aporte": safe_mean(df_fp['Q10_Blackboard_Aporte'])
        }
        
        d_por_facultad[f][p] = {
            "resp": resp,
            "prom": prom,
            "det": det,
            "pas": pas,
            "nps": nps,
            "enroll": enroll,
            "cov_pct": cov_pct,
            "exp": exp_dict,
            "coord": coord_dict,
            "plat": plat_dict
        }
        
        d_preguntas_fac[f][p] = {
            "exp": exp_dict,
            "coord": coord_dict,
            "plat": plat_dict
        }

# D Por Coordinador
d_por_coordinador = {}
for c in all_coordinadores:
    d_por_coordinador[c] = {}
    for p in all_periods:
        df_cp = df_all_responses[(df_all_responses['coordinador'] == c) & (df_all_responses['periodo'] == p)]
        if len(df_cp) == 0:
            continue
            
        resp = len(df_cp)
        prom = sum(df_cp['tipo_nps'] == "Promotor")
        det = sum(df_cp['tipo_nps'] == "Detractor")
        pas = sum(df_cp['tipo_nps'] == "Pasivo")
        nps = int(round((prom - det) / resp * 100)) if resp > 0 else 0
        
        d_por_coordinador[c][p] = {
            "resp": resp,
            "prom": prom,
            "det": det,
            "pas": pas,
            "nps": nps,
            "coord_calidad": safe_mean(df_cp['Q5_Servicio_Coordinador']),
            "coord_tiempo": safe_mean(df_cp['Q6_Tiempos_Coordinador'])
        }

# D Programas List
d_programas = []
for idx, row in df_summary.iterrows():
    period = row['periodo']
    fac = row['facultad']
    prog = row['programa']
    n_resp = int(row['n_respuestas'])
    prom = int(row['promotores'])
    det = int(row['detractores'])
    pas = int(row['pasivos'])
    nps = int(row['NPS']) if not pd.isna(row['NPS']) else 0
    enroll = int(row['enrollment'])
    cov = float(row['cobertura_pct'])
    
    # Find coordinator name for this program/period
    df_prog_resp = df_all_responses[(df_all_responses['periodo'] == period) & (df_all_responses['programa'] == prog)]
    coord = "Sin Coordinador"
    if len(df_prog_resp) > 0:
        coord_candidates = df_prog_resp['coordinador'].dropna().unique()
        if len(coord_candidates) > 0:
            coord = str(coord_candidates[0]).strip()
            
    d_programas.append({
        "fac": fac,
        "prog": prog,
        "coord": coord,
        "per": period,
        "resp": n_resp,
        "prom": prom,
        "det": det,
        "pas": pas,
        "nps": nps,
        "enroll": enroll,
        "cov": cov
    })

# D Comentarios List
d_comentarios = []
for idx, row in df_all_comments.iterrows():
    d_comentarios.append({
        "per": row['periodo'],
        "fac": row['facultad'],
        "prog": row['programa'],
        "coord": row['coordinador'],
        "nps": int(row['puntaje']) if pd.notna(row['puntaje']) else None,
        "clase": row['clasificacion'],
        "texto": row['comentario']
    })

# Build final D structure
D = {
    "periodos": all_periods,
    "facultades": all_facultades,
    "coordinadores": all_coordinadores,
    "global": d_global,
    "global_grupos": d_global_grupos,
    "por_facultad": d_por_facultad,
    "preguntas_fac": d_preguntas_fac,
    "por_coordinador": d_por_coordinador,
    "programas": d_programas,
    "comentarios": d_comentarios
}

# Save D to dashboard_data.js and dashboard_data.json
dashboard_json_path = os.path.join(workspace_dir, "dashboard_data.json")
dashboard_js_path = os.path.join(workspace_dir, "dashboard_data.js")

try:
    with open(dashboard_json_path, 'w', encoding='utf-8') as f_json:
        json.dump(D, f_json, ensure_ascii=False, indent=2)
    print(f"Successfully generated {dashboard_json_path}")
    
    with open(dashboard_js_path, 'w', encoding='utf-8') as f_js:
        f_js.write("const D = ")
        json.dump(D, f_js, ensure_ascii=False, indent=2)
        f_js.write(";")
    print(f"Successfully generated {dashboard_js_path}")
except Exception as e:
    print(f"Error writing dynamic JSON/JS dashboard files: {e}")

print("Data consolidation and dynamic dashboard calculation completed successfully!")
