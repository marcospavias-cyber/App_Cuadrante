import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime
from datetime import timedelta
import io
import random
import calendar
from itertools import groupby
from operator import itemgetter

# ==============================================================================
# 1. CONFIGURACIÓN Y CONSTANTES
# ==============================================================================

st.set_page_config(layout="wide", page_title="Gestor V52.0 - Final Precision", page_icon="🚒")

# --- ESTILOS VISUALES ---
st.markdown("""
<style>
    .stApp { background-color: #f8f9fa; }
    h1 { color: #d32f2f; }
    div[data-testid="stMetricValue"] { font-size: 1.4rem; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# --- CRÉDITOS ---
st.markdown("<h5 style='text-align: center; color: #888;'>Arquitectura de Precisión - V52.0</h5>", unsafe_allow_html=True)
st.title("🚒 Gestor de Cuadrantes Inteligente")

TEAMS = ['A', 'B', 'C']
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# --- ESTRATEGIAS ---
STRATEGIES = {
    "standard": {
        "name": "🛡️ Estándar Elástica (Recomendada)",
        "blocks": [
            {"dur": 10, "cred": 4, "label": "📦 Bloque 10d (4 Cr)"},
            {"dur": 10, "cred": 3, "label": "📦 Bloque 10d (3 Cr)"},
            {"dur": 9,  "cred": 3, "label": "📦 Bloque 9d (3 Cr)"},
            {"dur": 4,  "cred": 1, "label": "🧩 Relleno 4d (1 Cr)"},
            {"dur": 1,  "cred": 1, "label": "🧩 Día Suelto (1 Cr)"}
        ],
        "auto_recipe": [
            {"dur": 10, "target": 4}, {"dur": 10, "target": 3}, 
            {"dur": 10, "target": 3}, {"dur": 9,  "target": 3}
        ]
    },
    "safe": {
        "name": "🔢 Matemática Pura",
        "blocks": [
            {"dur": 12, "cred": 4, "label": "📦 Largo 12d (4 Cr)"},
            {"dur": 9,  "cred": 3, "label": "📦 Medio 9d (3 Cr)"},
            {"dur": 6,  "cred": 2, "label": "📦 Corto 6d (2 Cr)"},
            {"dur": 4,  "cred": 1, "label": "🧩 Relleno 4d (1 Cr)"},
            {"dur": 1,  "cred": 1, "label": "🧩 Día Suelto (1 Cr)"}
        ],
        "auto_recipe": [
            {"dur": 12, "target": 4}, {"dur": 12, "target": 4}, 
            {"dur": 9, "target": 3}, {"dur": 6, "target": 2}
        ]
    }
}

DEFAULT_ROSTER = [
    {"ID_Puesto": "Jefe A",       "Nombre": "Jefe A",       "Turno": "A", "Rol": "Jefe",       "SV": False},
    {"ID_Puesto": "Subjefe A",    "Nombre": "Subjefe A",    "Turno": "A", "Rol": "Subjefe",    "SV": False},
    {"ID_Puesto": "Cond A",       "Nombre": "Cond A",       "Turno": "A", "Rol": "Conductor",  "SV": False},
    {"ID_Puesto": "Bombero A1",   "Nombre": "Bombero A1",   "Turno": "A", "Rol": "Bombero",    "SV": False},
    {"ID_Puesto": "Bombero A2",   "Nombre": "Bombero A2",   "Turno": "A", "Rol": "Bombero",    "SV": False},
    {"ID_Puesto": "Bombero A3",   "Nombre": "Bombero A3",   "Turno": "A", "Rol": "Bombero",    "SV": False},
    {"ID_Puesto": "Jefe B",       "Nombre": "Jefe B",       "Turno": "B", "Rol": "Jefe",       "SV": False},
    {"ID_Puesto": "Subjefe B",    "Nombre": "Subjefe B",    "Turno": "B", "Rol": "Subjefe",    "SV": False},
    {"ID_Puesto": "Cond B",       "Nombre": "Cond B",       "Turno": "B", "Rol": "Conductor",  "SV": False},
    {"ID_Puesto": "Bombero B1",   "Nombre": "Bombero B1",   "Turno": "B", "Rol": "Bombero",    "SV": False},
    {"ID_Puesto": "Bombero B2",   "Nombre": "Bombero B2",   "Turno": "B", "Rol": "Bombero",    "SV": False},
    {"ID_Puesto": "Bombero B3",   "Nombre": "Bombero B3",   "Turno": "B", "Rol": "Bombero",    "SV": False},
    {"ID_Puesto": "Jefe C",       "Nombre": "Jefe C",       "Turno": "C", "Rol": "Jefe",       "SV": False},
    {"ID_Puesto": "Subjefe C",    "Nombre": "Subjefe C",    "Turno": "C", "Rol": "Subjefe",    "SV": False},
    {"ID_Puesto": "Cond C",       "Nombre": "Cond C",       "Turno": "C", "Rol": "Conductor",  "SV": False},
    {"ID_Puesto": "Bombero C1",   "Nombre": "Bombero C1",   "Turno": "C", "Rol": "Bombero",    "SV": False},
    {"ID_Puesto": "Bombero C2",   "Nombre": "Bombero C2",   "Turno": "C", "Rol": "Bombero",    "SV": False},
    {"ID_Puesto": "Bombero C3",   "Nombre": "Bombero C3",   "Turno": "C", "Rol": "Bombero",    "SV": False},
]

# ==============================================================================
# 2. LÓGICA DE NEGOCIO (CORE)
# ==============================================================================

def get_short_id(name, role, turn):
    """Genera un ID corto para visualizar en celdas pequeñas (ej: JB, C1, B2)."""
    if role == "Jefe": return f"J{turn}"
    if role == "Subjefe": return f"S{turn}"
    if role == "Conductor": return f"C{turn}"
    if "Bombero" in name:
        parts = name.split()
        if len(parts) > 1:
            suffix = parts[-1]
            if len(suffix) >= 2:
                return f"B{suffix[-1]}{turn}"
    return f"{name[:3]}{turn}"

@st.cache_data
def generate_base_schedule(year):
    """Genera el patrón A-B-C para todo el año."""
    is_leap = calendar.isleap(year)
    total_days = 366 if is_leap else 365
    status = {'A': 0, 'B': 2, 'C': 1} 
    schedule = {team: [] for team in TEAMS}
    for _ in range(total_days):
        for t in TEAMS:
            schedule[t].append('T' if status[t] == 0 else 'L')
            status[t] = (status[t] + 1) % 3
    return schedule, total_days

def is_in_night_period(day_idx, year, night_periods):
    """Verifica si un día cae dentro de un periodo de nocturnidad."""
    current_date = datetime.date(year, 1, 1) + datetime.timedelta(days=day_idx)
    for start, end in night_periods:
        if start <= current_date <= end: return True
    return False

@st.cache_data
def get_night_transition_dates(night_periods):
    """Obtiene fechas críticas de cambio de nocturnidad."""
    dates = set()
    for start, end in night_periods:
        dates.add(end) 
    return dates

def calculate_stats(roster_df, requests, year):
    """Calcula estadísticas en tiempo real (Créditos y Días Naturales)."""
    base_sch, _ = generate_base_schedule(year)
    stats = {}
    for _, p in roster_df.iterrows():
        stats[p['Nombre']] = {'credits': 0, 'natural': 0}
    for req in requests:
        name = req['Nombre']
        if name not in stats: continue
        s_idx = req['Inicio'].timetuple().tm_yday - 1
        e_idx = req['Fin'].timetuple().tm_yday - 1
        row = roster_df[roster_df['Nombre'] == name].iloc[0]
        nat = (e_idx - s_idx) + 1
        cred = 0
        for d in range(s_idx, e_idx + 1):
            if base_sch[row['Turno']][d] == 'T': cred += 1
        stats[name]['credits'] += cred
        stats[name]['natural'] += nat
    return stats

def book_slot_gen(start_idx, duration, person, occupation_map):
    """Marca un slot como ocupado en el mapa temporal."""
    for i in range(start_idx, start_idx + duration):
        if i not in occupation_map: occupation_map[i] = []
        occupation_map[i].append(person)

# --- DETECTOR DE CONFLICTOS CENTRALIZADO ---
def analyze_slot(start_idx, duration, person, occupation_map_T, base_sch, year, transition_dates, daily_absent, daily_roles):
    """
    Analiza si un bombero puede irse de vacaciones en un rango dado.
    Reglas: Cupo global (2), Coincidencia Categoría, Coincidencia Turno, Nocturnas.
    """
    total_days = len(base_sch['A'])
    if start_idx + duration > total_days: return False, "Fuera de año", None
    
    my_start_natural = start_idx
    my_end_natural = start_idx + duration - 1

    # 1. Reglas Diarias
    for i in range(start_idx, start_idx + duration):
        # REGLA SAGRADA: Máximo 2 personas ausentes en toda la plantilla
        if len(daily_absent[i]) >= 2:
            return False, f"Cupo lleno (2 pers) el día {i+1}", "Sistema"

        # Regla Nocturna (Transición)
        d_obj = datetime.date(year, 1, 1) + timedelta(days=i)
        if d_obj in transition_dates:
            if base_sch[person['Turno']][i] == 'T': 
                return False, "Conflicto Nocturna (Día Transición)", "Nocturna"
        
        # Regla: No coincidir con alguien de tu mismo turno
        occupants_T = occupation_map_T.get(i, [])
        for occ in occupants_T:
            if occ['Turno'] == person['Turno']: 
                return False, f"Coincide turno T con {occ['Nombre']}", occ['Nombre']
    
    # 2. Regla Categoría (excepto Bomberos)
    if person['Rol'] != 'Bombero':
        for d_check in range(my_start_natural, my_end_natural + 1):
            people_today = daily_absent[d_check]
            roles_today = daily_roles[d_check]
            for p_name, p_role in zip(people_today, roles_today):
                if p_name != person['Nombre'] and p_role == person['Rol']:
                    return False, f"Coincidencia Categoría con {p_name}", p_name
                        
    return True, "OK", None


# ==============================================================================
# 3. ALGORITMO GENERADOR (VERSION INTELIGENTE - 3 FASES)
# ==============================================================================

def auto_generate_schedule(roster_df, year, night_periods, strategy_key, current_reqs):
    base_sch, total_days = generate_base_schedule(year)
    transition_dates = get_night_transition_dates(night_periods)
    
    # --- Construir Mapas de Ocupación Iniciales ---
    occupation_map_T = {i:[] for i in range(total_days)}
    daily_absent = {i: [] for i in range(total_days)}
    daily_roles = {i: [] for i in range(total_days)}
    
    for req in current_reqs:
        p_row = roster_df[roster_df['Nombre'] == req['Nombre']]
        if p_row.empty: continue
        p_row = p_row.iloc[0]
        s = req['Inicio'].timetuple().tm_yday - 1
        e = req['Fin'].timetuple().tm_yday - 1
        s = max(0, s); e = min(total_days - 1, e)
        for d in range(s, e + 1):
            daily_absent[d].append(req['Nombre'])
            daily_roles[d].append(p_row['Rol'])
            if base_sch[p_row['Turno']][d] == 'T':
                occupation_map_T[d].append(p_row)

    generated_requests = []
    people = roster_df.to_dict('records')
    priority_order = ["Jefe", "Subjefe", "Conductor", "Bombero"]
    
    # --- PASO CRÍTICO: Barajar aleatoriamente para justicia social ---
    random.shuffle(people)
    people.sort(key=lambda x: priority_order.index(x['Rol'])) # Respetar jerarquía pero con orden interno aleatorio
    
    RECIPE = STRATEGIES[strategy_key]['auto_recipe']
    
    for person in people:
        my_slots = []
        credits_got = 0
        natural_days_got = 0
        
        # Calcular estado actual del bombero
        my_existing = [r for r in current_reqs if r['Nombre'] == person['Nombre']]
        for r in my_existing:
            s = r['Inicio'].timetuple().tm_yday - 1
            e = r['Fin'].timetuple().tm_yday - 1
            dur = (e - s) + 1
            my_slots.append((s, dur))
            natural_days_got += dur
            for k in range(s, e+1):
                if 0 <= k < total_days and base_sch[person['Turno']][k] == 'T':
                    credits_got += 1
        
        # ======================================================
        # FASE 1: BLOQUES PRINCIPALES (Ideal)
        # ======================================================
        current_recipe = RECIPE.copy()
        current_recipe.sort(key=lambda x: x['dur'], reverse=True)
        
        for block in current_recipe:
            if credits_got >= 13: break
            duration = block['dur']
            target = block['target']
            
            valid_starts = []
            for d in range(0, total_days - duration):
                # REGLA SAGRADA: Cupo >= 2 bloquea
                if len(daily_absent[d]) >= 2: continue
                
                # Verificar que todo el bloque cabe
                block_broken = False
                for k in range(d, d+duration):
                    if len(daily_absent[k]) >= 2: block_broken = True; break
                if block_broken: continue

                c = sum([1 for k in range(d, d+duration) if base_sch[person['Turno']][k] == 'T'])
                if c == target:
                     is_valid, _, _ = analyze_slot(d, duration, person, occupation_map_T, base_sch, year, transition_dates, daily_absent, daily_roles)
                     if is_valid: valid_starts.append(d)
            
            if valid_starts:
                start = random.choice(valid_starts)
                overlap = False
                for ms in my_slots:
                    if not (start + duration - 1 < ms[0] or start > ms[0] + ms[1] - 1): 
                        overlap = True; break
                
                if not overlap:
                    book_slot_gen(start, duration, person, occupation_map_T)
                    for k in range(start, start + duration):
                        daily_absent[k].append(person['Nombre'])
                        daily_roles[k].append(person['Rol'])
                    my_slots.append((start, duration))
                    credits_got += target
                    natural_days_got += duration
                    generated_requests.append({
                        "Nombre": person['Nombre'],
                        "Inicio": datetime.date(year, 1, 1) + timedelta(days=start),
                        "Fin": datetime.date(year, 1, 1) + timedelta(days=start+duration-1)
                    })

        # ======================================================
        # FASE 2: RESCATE (Asegurar los 13 Créditos)
        # ======================================================
        # Si falló la fase 1 por cupo lleno, buscamos huecos pequeños
        rescue_blocks = [{"dur": 4, "target": 1}, {"dur": 3, "target": 1}, {"dur": 1, "target": 1}]
        
        for r_block in rescue_blocks:
            # Intentos repetidos para llenar los créditos que falten
            for _ in range(15): 
                if credits_got >= 13: break
                duration = r_block['dur']
                target = r_block['target']
                
                valid_starts = []
                for d in range(0, total_days - duration):
                    if len(daily_absent[d]) >= 2: continue
                    block_broken = False
                    for k in range(d, d+duration):
                         if len(daily_absent[k]) >= 2: block_broken = True; break
                    if block_broken: continue
                    
                    c = sum([1 for k in range(d, d+duration) if base_sch[person['Turno']][k] == 'T'])
                    if c >= target:
                         is_valid, _, _ = analyze_slot(d, duration, person, occupation_map_T, base_sch, year, transition_dates, daily_absent, daily_roles)
                         if is_valid: valid_starts.append(d)
                
                if valid_starts:
                    start = random.choice(valid_starts)
                    overlap = False
                    for ms in my_slots:
                        if not (start + duration - 1 < ms[0] or start > ms[0] + ms[1] - 1): 
                            overlap = True; break
                    
                    if not overlap:
                        book_slot_gen(start, duration, person, occupation_map_T)
                        for k in range(start, start + duration):
                            daily_absent[k].append(person['Nombre'])
                            daily_roles[k].append(person['Rol'])
                        my_slots.append((start, duration))
                        credits_got += target
                        natural_days_got += duration
                        generated_requests.append({
                            "Nombre": person['Nombre'],
                            "Inicio": datetime.date(year, 1, 1) + timedelta(days=start),
                            "Fin": datetime.date(year, 1, 1) + timedelta(days=start+duration-1)
                        })

        # ======================================================
        # FASE 3: RELLENO DE NATURALES (Hasta 39 días)
        # ======================================================
        # Si tiene los créditos pero le faltan días para disfrutar
        needed_natural = 39 - natural_days_got
        if needed_natural > 0:
            potential_days = list(range(total_days))
            random.shuffle(potential_days)
            
            for d in potential_days:
                if needed_natural <= 0: break
                
                # 1. Regla Sagrada: Cupo
                if len(daily_absent[d]) >= 2: continue
                
                # 2. Evitar solapamiento propio
                overlap = False
                for ms in my_slots:
                    if not (d < ms[0] or d > ms[0] + ms[1] - 1): overlap = True; break
                if overlap: continue

                # 3. Solo días libres si ya cumplió créditos, o días T si aun le faltan (caso raro)
                is_working_day = (base_sch[person['Turno']][d] == 'T')
                if credits_got >= 13 and is_working_day: continue
                
                is_valid, _, _ = analyze_slot(d, 1, person, occupation_map_T, base_sch, year, transition_dates, daily_absent, daily_roles)
                
                if is_valid:
                    if is_working_day:
                         book_slot_gen(d, 1, person, occupation_map_T)
                         credits_got += 1
                    
                    daily_absent[d].append(person['Nombre'])
                    daily_roles[d].append(person['Rol'])
                    my_slots.append((d, 1))
                    natural_days_got += 1
                    needed_natural -= 1
                    
                    generated_requests.append({
                        "Nombre": person['Nombre'],
                        "Inicio": datetime.date(year, 1, 1) + timedelta(days=d),
                        "Fin": datetime.date(year, 1, 1) + timedelta(days=d)
                    })

    return generated_requests

# ==============================================================================
# 4. RENDERIZADO VISUAL Y EXCEL
# ==============================================================================

@st.cache_data
def render_global_occupation_calendar(year, roster_df, requests, night_periods):
    """Renderiza el mapa de calor HTML de ocupación."""
    base_sch, total_days = generate_base_schedule(year)
    transition_dates = get_night_transition_dates(night_periods)
    occ_map = {d: [] for d in range(total_days)}
    
    for req in requests:
        name = req['Nombre']
        if name not in roster_df['Nombre'].values: continue
        person_row = roster_df[roster_df['Nombre'] == name].iloc[0]
        turn = person_row['Turno']
        s = req['Inicio'].timetuple().tm_yday - 1
        e = req['Fin'].timetuple().tm_yday - 1
        for d in range(s, e+1):
            occ_map[d].append(get_short_id(name, person_row['Rol'], turn))

    html = "<div style='font-family:monospace; font-size:9px; overflow-x: auto;'>"
    html += """
    <div style='display:flex; gap:10px; margin-bottom:10px; font-size:11px; font-weight:bold;'>
        <span style='background:#d4edda; color:#155724; padding:2px 6px; border:1px solid #c3e6cb;'>🟩 DISPONIBLE</span>
        <span style='background:#FFF3CD; color:#856404; padding:2px 6px; border:1px solid #FFEEBA;'>🟧 ÚLTIMA PLAZA</span>
        <span style='background:#F8D7DA; color:#721c24; padding:2px 6px; border:1px solid #F5C6CB;'>🟥 COMPLETO (2/2)</span>
    </div>
    """
    html += "<div style='display:flex; margin-bottom:2px;'><div style='width:35px;'></div>"
    for d in range(1, 32):
        html += f"<div style='width:32px; text-align:center; color:#888;'>{d}</div>"
    html += "</div>"

    for m_idx, mes in enumerate(MESES):
        m_num = m_idx + 1
        days_in_month = calendar.monthrange(year, m_num)[1]
        html += f"<div style='display:flex; margin-bottom:2px;'><div style='width:35px; font-weight:bold; padding-top:8px;'>{mes}</div>"
        for d in range(1, 32):
            if d <= days_in_month:
                dt = datetime.date(year, m_num, d)
                d_idx = dt.timetuple().tm_yday - 1
                occupants = occ_map[d_idx]
                count = len(occupants)
                
                if count == 0: bg = "#d4edda"; txt_col = "#155724"
                elif count == 1: bg = "#FFF3CD"; txt_col = "#856404"
                else: bg = "#F8D7DA"; txt_col = "#721c24" # Rojo si >= 2

                border = "1px solid #fff"
                if dt in transition_dates: border = "2px solid red"
                label = "<br>".join(occupants)
                html += f"<div style='width:32px; height:30px; background-color:{bg}; color:{txt_col}; text-align:center; border:{border}; border-radius:2px; font-size:8px; line-height:9px; display:flex; align-items:center; justify-content:center;'>{label}</div>"
            else:
                html += "<div style='width:32px;'></div>"
        html += "</div>"
    html += "</div>"
    return html

def render_annual_calendar(year, team, base_sch, night_periods, custom_schedule=None):
    """Renderiza el calendario individual de un bombero."""
    html = f"<div style='font-family:monospace; font-size:10px;'>"
    html += """
    <div style='display:flex; gap:10px; margin-bottom:5px; font-size:11px; font-weight:bold;'>
        <span style='background:#d4edda; color:#155724; padding:2px 5px; border:1px solid #c3e6cb;'>T (Guardia)</span>
        <span style='background:#FFC000; color:#000; padding:2px 5px; border:1px solid #DAA520;'>V (Vacaciones)</span>
        <span style='background:#1E7E34; color:white; padding:2px 5px;'>T (Noche)</span>
    </div>
    """
    html += "<div style='display:flex; margin-bottom:2px;'><div style='width:30px;'></div>"
    for d in range(1, 32):
        html += f"<div style='width:20px; text-align:center; color:#888;'>{d}</div>"
    html += "</div>"
    for m_idx, mes in enumerate(MESES):
        m_num = m_idx + 1
        days_in_month = calendar.monthrange(year, m_num)[1]
        html += f"<div style='display:flex; margin-bottom:2px;'><div style='width:30px; font-weight:bold;'>{mes}</div>"
        for d in range(1, 32):
            if d <= days_in_month:
                dt = datetime.date(year, m_num, d)
                d_idx = dt.timetuple().tm_yday - 1
                state = base_sch[team][d_idx]
                final_val = state
                if custom_schedule: final_val = custom_schedule[d_idx]
                
                bg_color = "#eee"; text_color = "#ccc"; border = "1px solid #fff"
                
                if str(final_val).startswith('T'): 
                    bg_color = "#d4edda"; text_color = "#155724"
                    if is_in_night_period(d_idx, year, night_periods):
                        bg_color = "#1E7E34"; text_color = "white"
                elif final_val == 'V':
                    bg_color = "#FFC000"; text_color = "#000"
                
                if dt in get_night_transition_dates(night_periods): border = "2px solid red"
                
                html += f"<div style='width:20px; background-color:{bg_color}; color:{text_color}; text-align:center; border:{border}; border-radius:2px;'>{state[0]}</div>"
            else:
                html += "<div style='width:20px;'></div>"
        html += "</div>"
    html += "</div>"
    return html

def validate_and_generate_final(roster_df, requests, year, night_periods, strategy_key="standard"):
    """Calcula el cuadrante final y asigna coberturas automáticas."""
    base_schedule_turn, total_days = generate_base_schedule(year)
    final_schedule = {} 
    
    for _, row in roster_df.iterrows():
        final_schedule[row['Nombre']] = base_schedule_turn[row['Turno']].copy()

    day_vacations = {i: [] for i in range(total_days)}
    unavailable_map = {i: set() for i in range(total_days)}

    # Marcar Vacaciones
    for req in requests:
        name = req['Nombre']
        s_idx = req['Inicio'].timetuple().tm_yday - 1
        e_idx = req['Fin'].timetuple().tm_yday - 1
        
        for d in range(s_idx, e_idx + 1):
            unavailable_map[d].add(name)
            if final_schedule[name][d] == 'T':
                day_vacations[d].append(name)
                final_schedule[name][d] = 'V'
            else:
                final_schedule[name][d] = 'V(L)'

    adjustments_log = []
    coverage_counts = {name: 0 for name in roster_df['Nombre']}
    
    # Asignar Coberturas
    for d in range(total_days):
        absent_people = day_vacations[d]
        if not absent_people: continue
        
        potential_coverers = []
        for _, candidate in roster_df.iterrows():
            c_name = candidate['Nombre']
            # Debe estar libre y no ocupado
            if final_schedule[c_name][d] == 'L' and c_name not in unavailable_map[d]:
                # Verificar descanso (no doblar)
                prev_day = final_schedule[c_name][d-1] if d > 0 else 'L'
                next_day = final_schedule[c_name][d+1] if d < total_days-1 else 'L'
                if str(prev_day).startswith('T') or str(next_day).startswith('T'): continue
                
                potential_coverers.append(c_name)
        
        # Aleatoriedad y Justicia en la cobertura
        random.shuffle(potential_coverers)
        potential_coverers.sort(key=lambda x: coverage_counts[x])
        
        for missing in absent_people:
            # Filtro básico de rol
            role_missing = roster_df[roster_df['Nombre']==missing].iloc[0]['Rol']
            found = None
            for cand in potential_coverers:
                role_cand = roster_df[roster_df['Nombre']==cand].iloc[0]['Rol']
                compatible = False
                if role_missing in ["Jefe", "Subjefe"] and role_cand in ["Jefe", "Subjefe"]: compatible = True
                elif role_missing == "Conductor" and role_cand == "Conductor": compatible = True
                elif role_missing == "Bombero": compatible = True 
                
                if compatible:
                    found = cand
                    break
            
            if found:
                final_schedule[found][d] = f"T*({missing})"
                coverage_counts[found] += 1
                adjustments_log.append((d, found, missing))
                potential_coverers.remove(found)
                unavailable_map[d].add(found)

    return final_schedule, adjustments_log, coverage_counts

def create_final_excel(schedule, roster_df, year, requests, adjustments_log):
    """Genera el archivo Excel descargable."""
    wb = Workbook()
    
    # Estilos Excel
    s_T = PatternFill("solid", fgColor="C6EFCE")
    s_V = PatternFill("solid", fgColor="FFC000")
    s_Cov = PatternFill("solid", fgColor="FFC7CE")
    
    font_bold = Font(bold=True)
    align_c = Alignment(horizontal="center", vertical="center")
    border_all = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    # HOJA 1
    ws1 = wb.active; ws1.title = "Cuadrante"
    curr_row = 1
    for t in TEAMS:
        ws1.cell(curr_row, 1, f"TURNO {t}").font = Font(bold=True, color="FFFFFF"); 
        ws1.cell(curr_row, 1).fill = PatternFill("solid", fgColor="000080")
        curr_row += 1
        
        members = roster_df[roster_df['Turno'] == t]
        for _, p in members.iterrows():
            nm = p['Nombre']
            ws1.cell(curr_row, 1, f"{nm} ({p['Rol']})").font = font_bold
            
            for m_idx, mes in enumerate(MESES):
                days_in_month = calendar.monthrange(year, m_idx+1)[1]
                for d in range(1, days_in_month + 1):
                    dt = datetime.date(year, m_idx+1, d)
                    d_idx = dt.timetuple().tm_yday - 1
                    col_idx = 2 + d_idx 
                    val = schedule[nm][d_idx]
                    cell = ws1.cell(curr_row, col_idx)
                    cell.border = border_all
                    
                    if val == 'T': cell.fill = s_T; cell.value = "T"
                    elif val == 'V': cell.fill = s_V; cell.value = "V"
                    elif str(val).startswith('T*'): cell.fill = s_Cov; cell.value = "T*"
            curr_row += 1
        curr_row += 1

    # HOJA 2 ESTADISTICAS
    ws2 = wb.create_sheet("Estadísticas")
    ws2.append(["Nombre", "Rol", "Total Guardias Trabajadas", "Total Vacaciones (Naturales)"])
    for _, p in roster_df.iterrows():
        nm = p['Nombre']
        worked = sum([1 for d in schedule[nm] if str(d).startswith('T')])
        nat_days = 0
        for r in requests:
            if r['Nombre'] == nm:
                nat_days += (r['Fin'] - r['Inicio']).days + 1
        ws2.append([nm, p['Rol'], worked, nat_days])

    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out

# ==============================================================================
# 5. UI STREAMLIT
# ==============================================================================

if 'raw_requests_df' not in st.session_state:
    st.session_state.raw_requests_df = pd.DataFrame(columns=["Nombre", "Inicio", "Fin"])
if 'nights' not in st.session_state: st.session_state.nights = []
if 'roster_data' not in st.session_state: st.session_state.roster_data = pd.DataFrame(DEFAULT_ROSTER)

current_requests = st.session_state.raw_requests_df.to_dict('records')
year_val = 2026 
stats = calculate_stats(st.session_state.roster_data, current_requests, year_val)

with st.sidebar:
    st.header("Panel de Control")
    strategy_key = st.selectbox("🎯 Estrategia", list(STRATEGIES.keys()), format_func=lambda x: STRATEGIES[x]['name'])
    
    st.markdown("---")
    if st.button("🗑️ Borrar Todo"):
        st.session_state.raw_requests_df = pd.DataFrame(columns=["Nombre", "Inicio", "Fin"])
        st.rerun()

    st.markdown("---")
    if st.button("🎲 Rellenar Automático"):
        with st.spinner("Optimizando cuadrante (Fase 1, 2 y 3)..."):
            new_reqs = auto_generate_schedule(st.session_state.roster_data, year_val, st.session_state.nights, strategy_key, current_reqs)
            if new_reqs:
                df_new = pd.DataFrame(new_reqs)
                st.session_state.raw_requests_df = pd.concat([st.session_state.raw_requests_df, df_new], ignore_index=True)
                st.success(f"Generados {len(new_reqs)} periodos nuevos.")
                st.rerun()
            else:
                st.warning("No se encontraron más huecos disponibles.")

# MAIN LAYOUT
st.subheader("🌍 Mapa de Calor (Ocupación Global)")
st.caption("Verde: Libre | Rojo: Cupo lleno (2 personas)")
st.markdown(render_global_occupation_calendar(year_val, st.session_state.roster_data, current_requests, st.session_state.nights), unsafe_allow_html=True)

col_1, col_2 = st.columns([1, 2])

with col_1:
    st.subheader("Selección Manual")
    sel_person = st.selectbox("Bombero:", st.session_state.roster_data['Nombre'])
    if sel_person:
        curr_s = stats[sel_person]
        st.metric("Créditos (Guardias)", f"{curr_s['credits']} / 13", delta=13-curr_s['credits'])
        st.metric("Naturales", f"{curr_s['natural']} / 39", delta=39-curr_s['natural'])
        
        d_range = st.date_input("Elegir fechas", [])
        if len(d_range) == 2:
            if st.button("Añadir Vacaciones"):
                new_row = {"Nombre": sel_person, "Inicio": d_range[0], "Fin": d_range[1]}
                st.session_state.raw_requests_df = pd.concat([st.session_state.raw_requests_df, pd.DataFrame([new_row])], ignore_index=True)
                st.rerun()
        
        my_reqs = st.session_state.raw_requests_df[st.session_state.raw_requests_df['Nombre'] == sel_person]
        if not my_reqs.empty:
            st.write("Periodos asignados:")
            st.dataframe(my_reqs)

with col_2:
    if sel_person:
        st.subheader(f"Calendario de {sel_person}")
        p_data = st.session_state.roster_data[st.session_state.roster_data['Nombre']==sel_person].iloc[0]
        base_sch, _ = generate_base_schedule(year_val)
        preview_sch = base_sch[p_data['Turno']].copy()
        
        # Pinta vacaciones actuales
        for r in current_requests:
            if r['Nombre'] == sel_person:
                s = r['Inicio'].timetuple().tm_yday - 1
                e = r['Fin'].timetuple().tm_yday - 1
                for k in range(s, e+1):
                    preview_sch[k] = 'V'
        
        st.markdown(render_annual_calendar(year_val, p_data['Turno'], base_sch, st.session_state.nights, preview_sch), unsafe_allow_html=True)

st.divider()
if st.button("📥 Generar Excel Final", type="primary"):
    sch, adj, counts = validate_and_generate_final(st.session_state.roster_data, current_requests, year_val, st.session_state.nights)
    excel_file = create_final_excel(sch, st.session_state.roster_data, year_val, current_requests, adj)
    st.download_button("Descargar Cuadrante .xlsx", excel_file, f"Cuadrante_{year_val}.xlsx")
