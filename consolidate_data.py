import os
import glob
import re
import json
import pandas as pd
import numpy as np
import sqlite3
from datetime import datetime

# ─────────────────────────────────────────────────────────────────
#  CONFIGURACIÓN — cambia solo esta ruta si mueves la carpeta
# ─────────────────────────────────────────────────────────────────
workspace_dir = os.path.dirname(os.path.abspath(__file__))
subdirs = ["2024-10", "2024-20", "2025-10", "2025-20"]
enrollment_excel = os.path.join(workspace_dir, "Detalle Estudiantes Postgrado_Tipo Ingreso.xlsx")
coordinadores_excel = os.path.join(workspace_dir, "Coordinadores por periodos.xlsx")
html_template   = os.path.join(workspace_dir, "index.html")
output_html     = os.path.join(workspace_dir, "Dashboard_NPS_UDLA.html")

# ─────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────
def parse_rating(val):
    if pd.isna(val):
        return None
    m = re.match(r'^(\d+)', str(val).strip())
    return int(m.group(1)) if m else None

def clean_text(val):
    if pd.isna(val):
        return ""
    return re.sub(r'\s+', ' ', str(val).strip())

def clean_matching_name(s):
    if pd.isna(s):
        return ""
    s = str(s).lower().strip()
    for src, dst in [("í","i"),("á","a"),("ó","o"),("ú","u"),("ñ","n"),("é","e"),("ü","u")]:
        s = s.replace(src, dst)
    for src, dst in [("dip.en","diplomado"),("dip. de","diplomado"),("dip.","diplomado"),
                     ("rehab.","rehabilitacion"),("fisiat.","fisiatria"),("veter.","veterinaria")]:
        s = s.replace(src, dst)
    stop = {"en","de","el","la","y","para","del","con","a","los","las"}
    return "".join(w for w in re.findall(r'[a-z0-9]+', s) if w not in stop)

def safe_mean(series):
    s = series.dropna()
    return float(round(s.mean(), 1)) if len(s) else None

def get_faculty_name(filename):
    fn = filename.lower()
    if "arquitectura" in fn:                              return "Arquitectura, Diseño y Construcción"
    if "comunicaciones" in fn:                            return "Comunicaciones"
    if "derecho" in fn:                                   return "Derecho"
    if "educaci" in fn and "vra" in fn:                   return "Educación VRA"
    if "educaci" in fn:                                   return "Educación"
    if "ingenier" in fn or "negocio" in fn:               return "Ingeniería y Negocios"
    if "medicina" in fn or "veterinaria" in fn or "agronom" in fn: return "Medicina Veterinaria y Agronomía"
    if "salud" in fn or "social" in fn:                   return "Salud y Ciencias Sociales"
    return "Otra"

print("=" * 60)
print("  ACTUALIZADOR DE DASHBOARD NPS — UDLA")
print(f"  {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
print("=" * 60)

# ─────────────────────────────────────────────────────────────────
#  1. MATRÍCULAS REALES
# ─────────────────────────────────────────────────────────────────
enrollment_dict = {}
global_enrollment_period = {}
fac_enrollment_period = {}

if os.path.exists(enrollment_excel):
    print("\n[1/4] Cargando matrículas reales...")
    try:
        df_enr = pd.read_excel(
            enrollment_excel, sheet_name='BD MATRICULA',
            usecols=['Periodo','Rut','Programa','Programa Postgrado','Facultad Postgrado']
        )
        print(f"      {len(df_enr)} registros de matrícula cargados.")

        enr_grouped = df_enr.groupby(['Periodo','Programa Postgrado']).size().reset_index(name='estudiantes')

        for idx, r_en in df_enr.drop_duplicates(subset=['Rut','Periodo']).groupby(['Periodo']).size().items():
            per_str = f"{str(idx)[:4]}-{str(idx)[4:]}"
            global_enrollment_period[per_str] = int(r_en)

        enr_fac = df_enr.groupby(['Periodo','Facultad Postgrado']).size().reset_index(name='estudiantes')
        for _, row in enr_fac.iterrows():
            per_str = f"{str(row['Periodo'])[:4]}-{str(row['Periodo'])[4:]}"
            fac_enrollment_period[f"{str(row['Facultad Postgrado']).strip()}||{per_str}"] = int(row['estudiantes'])

        for _, row in enr_grouped.iterrows():
            per_str = f"{str(row['Periodo'])[:4]}-{str(row['Periodo'])[4:]}"
            prog_name = str(row['Programa Postgrado']).strip()
            enrollment_dict[f"{per_str}||{clean_matching_name(prog_name)}"] = int(row['estudiantes'])

        print("      Matrículas indexadas correctamente.")
    except Exception as e:
        print(f"      ERROR al cargar matrículas: {e}")
else:
    print("\n[1/4] AVISO: Archivo de matrículas no encontrado. Se usarán estimaciones.")

# ─────────────────────────────────────────────────────────────────
#  2. PROCESAR ENCUESTAS EXCEL
# ─────────────────────────────────────────────────────────────────
print("\n[2/4] Procesando encuestas por período...")
responses_list = []
comments_list  = []

# Cargar coordinadores por programa y período
coord_map = {}
if os.path.exists(coordinadores_excel):
    xl_coord = pd.read_excel(coordinadores_excel, sheet_name=None)
    for sheet_name, df_c in xl_coord.items():
        per = sheet_name.strip()
        df_c.columns = [str(c).strip() for c in df_c.columns]
        coord_col = [c for c in df_c.columns if 'oordinador' in c][0]
        prog_col  = [c for c in df_c.columns if 'rograma' in c or 'ombre' in c][0]
        for _, row in df_c.iterrows():
            c_val = str(row[coord_col]).strip() if pd.notna(row[coord_col]) else ''
            p_val = str(row[prog_col]).strip()  if pd.notna(row[prog_col])  else ''
            if c_val and p_val and c_val != 'nan' and p_val != 'nan':
                c_val = c_val.replace('Vanessa Alvarez', 'Vanessa Álvarez')
                c_val = c_val.replace('Paula Díaz Espinoza', 'Paula Díaz')
                c_val = c_val.strip()
                coord_map[(per, p_val.upper().strip())] = c_val
    print(f"      {len(coord_map)} asignaciones de coordinador cargadas.")

# Mapeo manual para programas cuyos nombres difieren entre encuesta y archivo de coordinadores
coord_manual = {
    ('2024-20', 'DIPLOMADO EN BIOESTADÍSTICA'): 'Sergio Román',
    ('2025-10', 'DIPLOMADO EN AGROECOLOGÍA'): 'Valentina Uribarri',
    ('2025-10', 'DIPLOMADO EN INTELIGENCIA ARTIFICIAL EN LOS NEGOCIOS'): 'Sergio Román',
    ('2025-10', 'DIPLOMADO EN PRODUCCIÓN DE EVENTOS EMPRESARIALES, TECNOLOGÍAS E INTELIGENCIA ARTIFICIAL'): 'Carolina Castro',
    ('2025-20', 'DIPLOMADO EN PERITAJE PSICOSOCIAL EN MATERIA PENAL Y FAMILIA'): 'Valentina Uribarri',
    ('2025-20', 'DIPLOMADO EN INVESTIGACIÓN – ACCIÓN APLICADA A LA DOCENCIA UNIVERSITARIA'): 'Vanessa Álvarez',
    ('2025-20', 'DIPLOMADO EN BIOESTADÍSTICA'): 'Sergio Román',
}

for sd in subdirs:
    path  = os.path.join(workspace_dir, sd)
    files = glob.glob(os.path.join(path, "*.xlsx")) if os.path.exists(path) else []
    print(f"      {sd}: {len(files)} archivos")

    for f in files:
        filename = os.path.basename(f)
        faculty  = get_faculty_name(filename)
        try:
            xl = pd.ExcelFile(f)
            sheet = 'Sheet' if 'Sheet' in xl.sheet_names else xl.sheet_names[0]
            df = pd.read_excel(f, sheet_name=sheet)
            df_clean = df[df['respondent_id'].notna()].copy()

            for _, row in df_clean.iterrows():
                prog = clean_text(row.iloc[9])
                if not prog or prog == "Response":
                    continue

                q1  = parse_rating(row.iloc[10])
                q2  = parse_rating(row.iloc[11])
                q3  = parse_rating(row.iloc[12])
                q4  = parse_rating(row.iloc[13])
                coord = clean_text(row.iloc[14])
                coord = coord.replace("Vanessa Alvarez", "Vanessa Álvarez")
                coord = coord.replace("Paula Díaz Espinoza", "Paula Díaz")

                prog_upper = prog.upper().strip()

                # 1. Mapa manual (máxima prioridad)
                coord_from_map = coord_manual.get((sd, prog_upper))

                # 2. Coincidencia exacta en archivo
                if not coord_from_map:
                    coord_from_map = coord_map.get((sd, prog_upper))

                # 3. Prefijo
                if not coord_from_map:
                    for (map_per, map_prog), map_coord in coord_map.items():
                        if map_per == sd and (map_prog.startswith(prog_upper[:45]) or prog_upper[:45] in map_prog):
                            coord_from_map = map_coord
                            break

                if coord_from_map:
                    coord = coord_from_map
                elif not coord:
                    coord = "Sin Coordinador"
                q5  = parse_rating(row.iloc[15])
                q6  = parse_rating(row.iloc[16])
                q7  = parse_rating(row.iloc[17])
                q8  = parse_rating(row.iloc[18])
                q9  = parse_rating(row.iloc[19])
                q10 = parse_rating(row.iloc[20])
                com = clean_text(row.iloc[21])

                tipo_nps = None
                if q1 is not None:
                    if q1 >= 6:
                        tipo_nps = "Promotor"
                    elif q1 >= 4:
                        tipo_nps = "Pasivo"
                    else:
                        tipo_nps = "Detractor"

                responses_list.append({
                    "respondent_id": str(int(row.iloc[0])),
                    "periodo": sd, "facultad": faculty,
                    "programa": prog,
                    "coordinador": coord if coord else "Sin Coordinador",
                    "Q1_Recomendacion_NPS": q1, "Q2_Contenidos": q2,
                    "Q3_Oportunidades": q3, "Q4_Volveria": q4,
                    "Q5_Servicio_Coordinador": q5, "Q6_Tiempos_Coordinador": q6,
                    "Q7_Portal_Facilidad": q7, "Q8_Portal_Aporte": q8,
                    "Q9_Blackboard_Facilidad": q9, "Q10_Blackboard_Aporte": q10,
                    "tipo_nps": tipo_nps
                })

                if com and com.lower() not in ["open-ended response","nan",""]:
                    sentimiento = "Neutro"
                    if q1 is not None:
                        sentimiento = "Positivo" if q1 >= 6 else ("Negativo" if q1 <= 3 else "Neutro")
                    comments_list.append({
                        "periodo": sd, "facultad": faculty,
                        "programa": prog,
                        "coordinador": coord if coord else "Sin Coordinador",
                        "puntaje": q1, "clasificacion": sentimiento, "comentario": com
                    })
        except Exception as e:
            print(f"        ERROR en {filename}: {e}")

df_all  = pd.DataFrame(responses_list)
df_coms = pd.DataFrame(comments_list)
print(f"      Total: {len(df_all)} respuestas, {len(df_coms)} comentarios.")

# ─────────────────────────────────────────────────────────────────
#  3. CALCULAR RESUMEN Y ESTRUCTURA D
# ─────────────────────────────────────────────────────────────────
print("\n[3/4] Calculando métricas y construyendo estructura de datos...")

def get_enrollment(period, program_name, n_resp):
    key = f"{period}||{clean_matching_name(program_name)}"
    if key in enrollment_dict:
        return enrollment_dict[key]
    for k, val in enrollment_dict.items():
        k_per, k_prog = k.split("||")
        if period == k_per:
            c = clean_matching_name(program_name)
            if c.startswith(k_prog[:15]) or k_prog.startswith(c[:15]):
                return val
    return int(np.ceil(n_resp / 0.55) + 2)

# Resumen por programa
summary_list = []
for (period, fac, prog), grp in df_all.groupby(['periodo','facultad','programa']):
    n   = len(grp)
    q1  = grp['Q1_Recomendacion_NPS'].dropna()
    pr  = int((grp['tipo_nps'] == "Promotor").sum())
    de  = int((grp['tipo_nps'] == "Detractor").sum())
    pa  = int((grp['tipo_nps'] == "Pasivo").sum())
    nps = int(round((pr - de) / len(q1) * 100)) if len(q1) else 0
    enr = get_enrollment(period, prog, n)
    summary_list.append({
        "periodo": period, "facultad": fac, "programa": prog,
        "n_respuestas": n, "promotores": pr, "pasivos": pa, "detractores": de,
        "NPS": nps, "enrollment": enr, "cobertura_pct": round(n/enr*100,1) if enr else 0,
        "prom_Q2": safe_mean(grp['Q2_Contenidos']),
        "prom_Q3": safe_mean(grp['Q3_Oportunidades']),
        "prom_Q4": safe_mean(grp['Q4_Volveria']),
        "prom_Q5": safe_mean(grp['Q5_Servicio_Coordinador']),
        "prom_Q6": safe_mean(grp['Q6_Tiempos_Coordinador']),
        "prom_Q7": safe_mean(grp['Q7_Portal_Facilidad']),
        "prom_Q8": safe_mean(grp['Q8_Portal_Aporte']),
        "prom_Q9": safe_mean(grp['Q9_Blackboard_Facilidad']),
        "prom_Q10": safe_mean(grp['Q10_Blackboard_Aporte']),
    })
df_sum = pd.DataFrame(summary_list)

all_periods = sorted(df_all['periodo'].unique().tolist())
all_facs    = sorted(df_all['facultad'].unique().tolist())
all_coords  = sorted(df_all['coordinador'].dropna().unique().tolist())

# Global
d_global = {}
for p in all_periods:
    g = df_all[df_all['periodo'] == p]
    pr = int((g['tipo_nps']=="Promotor").sum())
    de = int((g['tipo_nps']=="Detractor").sum())
    pa = int((g['tipo_nps']=="Pasivo").sum())
    n = len(g)
    nps = int(round((pr - de) / n * 100)) if n else 0
    enr = int(df_sum[df_sum['periodo']==p]['enrollment'].sum())
    d_global[p] = {"resp":n,"prom":int(pr),"det":int(de),"pas":int(pa),
                   "nps":nps,"enroll":enr,"cov_pct":round(n/enr*100,1) if enr else 0}

# Global grupos
d_global_grupos = {}
for p in all_periods:
    g = df_all[df_all['periodo']==p]
    d_global_grupos[p] = {
        "exp":  {"nps_val":safe_mean(g['Q1_Recomendacion_NPS']),"contenidos":safe_mean(g['Q2_Contenidos']),
                 "laboral":safe_mean(g['Q3_Oportunidades']),"recompra":safe_mean(g['Q4_Volveria'])},
        "coord":{"coord_calidad":safe_mean(g['Q5_Servicio_Coordinador']),"coord_tiempo":safe_mean(g['Q6_Tiempos_Coordinador'])},
        "plat": {"portal_facil":safe_mean(g['Q7_Portal_Facilidad']),"portal_aporte":safe_mean(g['Q8_Portal_Aporte']),
                 "bb_facil":safe_mean(g['Q9_Blackboard_Facilidad']),"bb_aporte":safe_mean(g['Q10_Blackboard_Aporte'])}
    }

# Por facultad
d_por_fac = {}
d_preg_fac = {}
for f in all_facs:
    d_por_fac[f] = {}
    d_preg_fac[f] = {}
    for p in all_periods:
        g = df_all[(df_all['facultad']==f) & (df_all['periodo']==p)]
        if len(g) == 0:
            continue
        n  = len(g)
        pr = int((g['tipo_nps']=="Promotor").sum())
        de = int((g['tipo_nps']=="Detractor").sum())
        pa = int((g['tipo_nps']=="Pasivo").sum())
        nps = int(round((pr-de)/n*100)) if n else 0
        enr = int(df_sum[(df_sum['facultad']==f)&(df_sum['periodo']==p)]['enrollment'].sum())
        exp   = {"nps_val":safe_mean(g['Q1_Recomendacion_NPS']),"contenidos":safe_mean(g['Q2_Contenidos']),
                 "laboral":safe_mean(g['Q3_Oportunidades']),"recompra":safe_mean(g['Q4_Volveria'])}
        coord = {"coord_calidad":safe_mean(g['Q5_Servicio_Coordinador']),"coord_tiempo":safe_mean(g['Q6_Tiempos_Coordinador'])}
        plat  = {"portal_facil":safe_mean(g['Q7_Portal_Facilidad']),"portal_aporte":safe_mean(g['Q8_Portal_Aporte']),
                 "bb_facil":safe_mean(g['Q9_Blackboard_Facilidad']),"bb_aporte":safe_mean(g['Q10_Blackboard_Aporte'])}
        d_por_fac[f][p]  = {"resp":n,"prom":pr,"det":de,"pas":pa,"nps":nps,"enroll":enr,
                             "cov_pct":round(n/enr*100,1) if enr else 0,"exp":exp,"coord":coord,"plat":plat}
        d_preg_fac[f][p] = {"exp":exp,"coord":coord,"plat":plat}

# Por coordinador
d_por_coord = {}
for c in all_coords:
    d_por_coord[c] = {}
    for p in all_periods:
        g = df_all[(df_all['coordinador']==c)&(df_all['periodo']==p)]
        if len(g) == 0:
            d_por_coord[c][p] = None
            continue
        n  = len(g)
        pr = int((g['tipo_nps']=="Promotor").sum())
        de = int((g['tipo_nps']=="Detractor").sum())
        pa = int((g['tipo_nps']=="Pasivo").sum())
        nps = int(round((pr-de)/n*100)) if n else 0
        d_por_coord[c][p] = {"resp":n,"prom":pr,"det":de,"pas":pa,"nps":nps,
                              "coord_calidad":safe_mean(g['Q5_Servicio_Coordinador']),
                              "coord_tiempo":safe_mean(g['Q6_Tiempos_Coordinador'])}

# Programas
d_programas = []
for _, row in df_sum.iterrows():
    period, fac, prog = row['periodo'], row['facultad'], row['programa']
    df_pr = df_all[(df_all['periodo']==period)&(df_all['programa']==prog)]
    coords = df_pr['coordinador'].dropna().unique()
    coord  = str(coords[0]).strip() if len(coords) else "Sin Coordinador"
    d_programas.append({
        "fac":fac,"prog":prog,"coord":coord,"per":period,
        "resp":int(row['n_respuestas']),"prom":int(row['promotores']),
        "det":int(row['detractores']),"pas":int(row['pasivos']),
        "nps":int(row['NPS']) if pd.notna(row['NPS']) else 0,
        "enroll":int(row['enrollment']),"cov":float(row['cobertura_pct'])
    })

# Comentarios
d_comentarios = []
for _, row in df_coms.iterrows():
    d_comentarios.append({
        "per":row['periodo'],"fac":row['facultad'],"prog":row['programa'],
        "coord":row['coordinador'],
        "nps":int(row['puntaje']) if pd.notna(row['puntaje']) else None,
        "clase":row['clasificacion'],"texto":row['comentario']
    })

# Estructura D final
D = {
    "periodos": all_periods, "facultades": all_facs, "coordinadores": all_coords,
    "global": d_global, "global_grupos": d_global_grupos,
    "por_facultad": d_por_fac, "preguntas_fac": d_preg_fac,
    "por_coordinador": d_por_coord,
    "programas": d_programas, "comentarios": d_comentarios
}

# ─────────────────────────────────────────────────────────────────
#  4. GUARDAR ARCHIVOS DE SALIDA
# ─────────────────────────────────────────────────────────────────
print("\n[4/4] Generando archivos de salida...")

# 4a. Excel consolidado
output_excel = os.path.join(workspace_dir, "NPS_Consolidado.xlsx")
try:
    df_resp_export = df_all.copy()
    df_com_export  = df_coms.copy()
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df_resp_export.to_excel(writer, sheet_name='Detalle_Respuestas', index=False)
        df_sum.to_excel(writer, sheet_name='Resumen_Programas', index=False)
        df_com_export.to_excel(writer, sheet_name='Comentarios', index=False)
    print(f"      ✓ {os.path.basename(output_excel)}")
except Exception as e:
    print(f"      ✗ Excel: {e}")

# 4b. SQLite
output_db = os.path.join(workspace_dir, "NPS_Consolidado.db")
try:
    conn = sqlite3.connect(output_db)
    df_all.to_sql('detalle_respuestas', conn, if_exists='replace', index=False)
    df_sum.to_sql('resumen_programas',  conn, if_exists='replace', index=False)
    df_coms.to_sql('comentarios',       conn, if_exists='replace', index=False)
    conn.close()
    print(f"      ✓ {os.path.basename(output_db)}")
except Exception as e:
    print(f"      ✗ SQLite: {e}")

# 4c. dashboard_data.js y .json (para GitHub Pages)
json_str = json.dumps(D, ensure_ascii=False, indent=2)
try:
    with open(os.path.join(workspace_dir, "dashboard_data.json"), 'w', encoding='utf-8') as f:
        f.write(json_str)
    with open(os.path.join(workspace_dir, "dashboard_data.js"), 'w', encoding='utf-8') as f:
        f.write("const D = ")
        f.write(json_str)
        f.write(";")
    print(f"      ✓ dashboard_data.js / dashboard_data.json")
except Exception as e:
    print(f"      ✗ dashboard_data: {e}")

# 4d. HTML autocontenido — el objetivo principal
# Lee index.html y reemplaza <script src="dashboard_data.js"></script>
# por el JSON embebido directamente en una etiqueta <script>
if os.path.exists(html_template):
    try:
        with open(html_template, 'r', encoding='utf-8') as f:
            html = f.read()

        # Calcular totales para el subtítulo dinámico
        total_resp  = len(df_all)
        total_facs  = len(all_facs)
        total_cords = len(all_coords)
        periodo_label = f"{all_periods[0]} a {all_periods[-1]}" if all_periods else ""
        fecha_gen = datetime.now().strftime('%d/%m/%Y %H:%M')

        # Reemplazar la carga externa por datos embebidos
        embedded_script = (
            f'<script>\n'
            f'/* Datos generados automáticamente el {fecha_gen} */\n'
            f'const D = {json_str};\n'
            f'</script>'
        )
        html_out = re.sub(
            r'<script\s+src=["\']dashboard_data\.js["\']></script>',
            embedded_script,
            html,
            flags=re.IGNORECASE
        )

        # Actualizar el subtítulo con cifras reales
        html_out = re.sub(
            r'(<p[^>]*>)([^<]*períodos[^<]*)</p>',
            f'<p>Comparativo {len(all_periods)} períodos · {total_resp:,} respuestas · {total_facs} facultades · {total_cords} coordinadores</p>',
            html_out
        )

        with open(output_html, 'w', encoding='utf-8') as f:
            f.write(html_out)

        size_kb = os.path.getsize(output_html) // 1024
        print(f"      ✓ {os.path.basename(output_html)}  ({size_kb} KB — autocontenido)")
    except Exception as e:
        print(f"      ✗ HTML autocontenido: {e}")
        import traceback; traceback.print_exc()
else:
    print(f"      ✗ No se encontró {html_template} — omitiendo generación de HTML.")

print("\n" + "=" * 60)
print("  ¡Consolidación completada!")
print(f"  Períodos: {', '.join(all_periods)}")
print(f"  Respuestas: {len(df_all):,}   Comentarios: {len(df_coms):,}")
print(f"  HTML listo: {os.path.basename(output_html)}")
print("=" * 60)
