import streamlit as st
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime
import io
import random
import calendar 
import pandas as pd
import os
from itertools import groupby
from operator import itemgetter
from datetime import timedelta

# ==============================================================================
# 1. CONFIGURACIÓN Y CONSTANTES
# ==============================================================================

st.set_page_config(layout="wide", page_title="Gestor V44.2 Fix")

# --- CONTRASEÑA ---
ADMIN_PASSWORD = "lucena2026"

# --- CONFIGURACIÓN UBH ---
TEAMS = ['A', 'B', 'C']
# ROLES BASE
ROLES = ["Capataz", "2º Capataz", "Bombero"] 
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
DB_FILE = "vacaciones_db.csv"
ADJ_FILE = "ajustes_db.csv"

# Objetivos de jornada (Actualizados)
MIN_WORK_DAYS = 126
MAX_WORK_DAYS = 128

# --- ESTRATEGIAS ---
STRATEGIES = {
    "standard": {
        "name": "🛡️ Estándar (10+10+10+9)",
        "desc": "3 bloques de 10 días y 1 de 9 días.",
        "blocks": [
            {"dur": 10, "cred": 4, "label": "Bloque 10d (4 Cr)"},
            {"dur": 10, "cred": 3, "label": "Bloque 10d (3 Cr)"},
            {"dur": 9,  "cred": 3, "label": "Bloque 9d (3 Cr)"}
        ],
        "auto_recipe": [ {"dur": 10, "target": 4}, {"dur": 10, "target": 3}, {"dur": 10, "target": 3}, {"dur": 9, "target": 3} ]
    },
    "sniper": {
        "name": "🎯 Francotirador (13 días)",
        "desc": "Eliges tus 13 guardias una a una.",
        "blocks": [ {"dur": 1, "cred": 1, "label": "Día Suelto (1 Cr)"} ],
        "auto_recipe": [{"dur": 1, "target": 1}] * 13
    }
}

# PLANTILLA POR DEFECTO (UBH)
DEFAULT_ROSTER = [
    {"ID_Puesto": "Capataz A", "Nombre": "Capataz A", "Turno": "A", "Rol": "Capataz", "Poli": False},
    {"ID_Puesto": "2º Cap A",  "Nombre": "2º Cap A",  "Turno": "A", "Rol": "2º Capataz", "Poli": False},
    {"ID_Puesto": "Bomb A1",   "Nombre": "Bombero A1", "Turno": "A", "Rol": "Bombero", "Poli": True},
    {"ID_Puesto": "Bomb A2",   "Nombre": "Bombero A2", "Turno": "A", "Rol": "Bombero", "Poli": False},
    {"ID_Puesto": "Bomb A3",   "Nombre": "Bombero A3", "Turno": "A", "Rol": "Bombero", "Poli": False},
    {"ID_Puesto": "Bomb A4",   "Nombre": "Bombero A4", "Turno": "A", "Rol": "Bombero", "Poli": False},
    
    {"ID_Puesto": "Capataz B", "Nombre": "Capataz B", "Turno": "B", "Rol": "Capataz", "Poli": False},
    {"ID_Puesto": "2º Cap B",  "Nombre": "2º Cap B",  "Turno": "B", "Rol": "2º Capataz", "Poli": False},
    {"ID_Puesto": "Bomb B1",   "Nombre": "Bombero B1", "Turno": "B", "Rol": "Bombero", "Poli": True},
    {"ID_Puesto": "Bomb B2",   "Nombre": "Bombero B2", "Turno": "B", "Rol": "Bombero", "Poli": False},
    {"ID_Puesto": "Bomb B3",   "Nombre": "Bombero B3", "Turno": "B", "Rol": "Bombero", "Poli": False},
    {"ID_Puesto": "Bomb B4",   "Nombre": "Bombero B4", "Turno": "B", "Rol": "Bombero", "Poli": False},

    {"ID_Puesto": "Capataz C", "Nombre": "Capataz C", "Turno": "C", "Rol": "Capataz", "Poli": False},
    {"ID_Puesto": "2º Cap C",  "Nombre": "2º Cap C",  "Turno": "C", "Rol": "2º Capataz", "Poli": False},
    {"ID_Puesto": "Bomb C1",   "Nombre": "Bombero C1", "Turno": "C", "Rol": "Bombero", "Poli": True},
    {"ID_Puesto": "Bomb C2",   "Nombre": "Bombero C2", "Turno": "C", "Rol": "Bombero", "Poli": False},
    {"ID_Puesto": "Bomb C3",   "Nombre": "Bombero C3", "Turno": "C", "Rol": "Bombero", "Poli": False},
    {"ID_Puesto": "Bomb C4",   "Nombre": "Bombero C4", "Turno": "C", "Rol": "Bombero", "Poli": False},
]

# ==============================================================================
# 2. FUNCIONES (ORDEN CORREGIDO Y BLINDADO)
# ==============================================================================

def load_data():
    if os.path.exists(DB_FILE):
        try:
            df = pd.read_csv(DB_FILE)
            if not df.empty:
                df['Inicio'] = pd.to_datetime(df['Inicio']).dt.date
                df['Fin'] = pd.to_datetime(df['Fin']).dt.date
            else: df = pd.DataFrame(columns=["Nombre", "Inicio", "Fin"])
        except: df = pd.DataFrame(columns=["Nombre", "Inicio", "Fin"])
    else:
        df = pd.DataFrame(columns=["Nombre", "Inicio", "Fin"])
    
    if os.path.exists(ADJ_FILE):
        try:
            adj_data = pd.read_csv(ADJ_FILE).to_dict('records')
        except: adj_data = []
    else:
        adj_data = []
    return df, adj_data

def save_data(df_vacs, list_adj):
    df_vacs.to_csv(DB_FILE, index=False)
    pd.DataFrame(list_adj).to_csv(ADJ_FILE, index=False)

def get_short_id(name, role, turn):
    if role == "Capataz": return f"CP{turn}"
    if role == "2º Capataz": return f"2C{turn}"
    if "Bombero" in name:
        parts = name.split()
        if len(parts) > 1:
            suffix = parts[-1]
            if len(suffix) >= 2: return f"B{suffix[-1]}{turn}"
    return f"{name[:3]}{turn}"

def generate_night_template():
    wb = Workbook()
    ws = wb.active; ws.title = "Plan Nocturnas"
    ws.append(["Inicio (dd/mm/yyyy)", "Fin (dd/mm/yyyy)", "Notas"])
    ws.append(["2026-01-10", "2026-01-12", "Ejemplo"])
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out

@st.cache_data
def generate_base_schedule(year):
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
    current_date = datetime.date(year, 1, 1) + datetime.timedelta(days=day_idx)
    for start, end in night_periods:
        if start <= current_date <= end: return True
    return False

@st.cache_data
def get_night_transition_dates(night_periods):
    dates = set()
    for start, end in night_periods:
        dates.add(end) 
    return dates

def calculate_stats(roster_df, requests, year):
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

def check_global_conflict_generic(start_idx, duration, person, occupation_map, base_sch, year, transition_dates):
    total_days = len(base_sch['A'])
    if start_idx + duration > total_days: return True

    for i in range(start_idx, start_idx + duration):
        d_obj = datetime.date(year, 1, 1) + timedelta(days=i)
        if d_obj in transition_dates:
            if base_sch[person['Turno']][i] == 'T': return True
        
        occupants = occupation_map.get(i, [])
        if len(occupants) >= 2: return True
        
        for occ in occupants:
            if occ['Turno'] == person['Turno']: return True
            # Normas UBH
            if person['Rol'] == 'Capataz' and occ['Rol'] == 'Capataz': return True
            if person['Rol'] == '2º Capataz' and occ['Rol'] == '2º Capataz': return True
    return False

def get_available_blocks_for_person(person_name, roster_df, current_requests, year, night_periods, month_range, strategy_key):
    base_sch, total_days = generate_base_schedule(year)
    transition_dates = get_night_transition_dates(night_periods)
    person = roster_df[roster_df['Nombre'] == person_name].iloc[0]
    
    start_month_idx = MESES.index(month_range[0]) + 1
    end_month_idx = MESES.index(month_range[1]) + 1
    
    occupation_map = {i:[] for i in range(total_days)}
    my_current_slots = [] 
    for req in current_requests:
        if req['Nombre'] != person_name:
            p_req = roster_df[roster_df['Nombre'] == req['Nombre']].iloc[0]
            s = req['Inicio'].timetuple().tm_yday - 1
            e = req['Fin'].timetuple().tm_yday - 1
            for d in range(s, e+1):
                if d < total_days and base_sch[p_req['Turno']][d] == 'T':
                    occupation_map[d].append(p_req)
        else:
            s = req['Inicio'].timetuple().tm_yday - 1
            e = req['Fin'].timetuple().tm_yday - 1
            my_current_slots.append((s, e))

    # Fallback si la estrategia no existe
    strat_data = STRATEGIES.get(strategy_key, STRATEGIES['standard'])
    block_defs = strat_data['blocks']
    
    options = {b['label']: [] for b in block_defs}
    
    for d in range(total_days): 
        d_date = datetime.date(year, 1, 1) + timedelta(days=d)
        if not (start_month_idx <= d_date.month <= end_month_idx): continue

        for b_def in block_defs:
            duration = b_def['dur']
            target_cred = b_def['cred']
            label_key = b_def['label']
            
            if d + duration > total_days: continue
            
            if not check_global_conflict_generic(d, duration, person, occupation_map, base_sch, year, transition_dates):
                overlap = False
                for ms in my_current_slots:
                    # Lógica de solapamiento simple
                    s_req, e_req = ms[0], ms[0] + (ms[1]-ms[0])
                    s_new, e_new = d, d + duration - 1
                    if not (e_new < s_req or s_new > e_req): 
                        overlap = True
                        break
                
                if not overlap:
                    credits = 0
                    for k in range(d, d+duration):
                        if base_sch[person['Turno']][k] == 'T': credits += 1
                    
                    if credits == target_cred:
                        start_date = d_date
                        end_date = start_date + timedelta(days=duration-1)
                        txt = f"{d_date.strftime('%d/%m')} - {end_date.strftime('%d/%m')}"
                        options[label_key].append({'label': txt, 'start': start_date, 'end': end_date})
    return options

def auto_generate_schedule(roster_df, year, night_periods, strategy_key):
    base_sch, total_days = generate_base_schedule(year)
    transition_dates = get_night_transition_dates(night_periods)
    occupation_map = {} 
    generated_requests = []
    people = roster_df.to_dict('records')
    priority_order = ["Capataz", "2º Capataz", "Bombero"]
    people.sort(key=lambda x: priority_order.index(x['Rol']))
    RECIPE = STRATEGIES.get(strategy_key, STRATEGIES['standard'])['auto_recipe']
    
    for person in people:
        my_slots = []
        current_recipe = RECIPE.copy()
        random.shuffle(current_recipe) 
        credits_got = 0
        
        for block in current_recipe:
            duration = block['dur']
            target = block['target']
            options = []
            for d in range(0, total_days - duration):
                c = sum([1 for k in range(d, d+duration) if base_sch[person['Turno']][k] == 'T'])
                if c == target:
                     if not check_global_conflict_generic(d, duration, person, occupation_map, base_sch, year, transition_dates):
                        options.append(d)
            random.shuffle(options)
            for start in options:
                overlap = False # Simplificado para auto
                if not overlap:
                    # book_slot_gen logic inline
                    for k in range(start, start + duration):
                        if k not in occupation_map: occupation_map[k] = []
                        occupation_map[k].append(person)
                    
                    credits_got += target
                    generated_requests.append({
                        "Nombre": person['Nombre'],
                        "Inicio": datetime.date(year, 1, 1) + timedelta(days=start),
                        "Fin": datetime.date(year, 1, 1) + timedelta(days=start+duration-1)
                    })
                    break 
    return generated_requests

def render_global_occupation_calendar(year, roster_df, requests, night_periods):
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
            if d < total_days and base_sch[turn][d] == 'T':
                occ_map[d].append(get_short_id(name, person_row['Rol'], turn))

    html = "<div style='font-family:monospace; font-size:9px;'>"
    html += "<div style='display:flex; gap:10px; margin-bottom:10px; font-weight:bold;'><span style='background:#d4edda; padding:2px 5px;'>Libre</span><span style='background:#FFF3CD; padding:2px 5px;'>1 Ocup</span><span style='background:#F8D7DA; padding:2px 5px;'>Lleno</span></div>"
    
    html += "<div style='display:flex; margin-bottom:2px;'><div style='width:35px;'></div>"
    for d in range(1, 32): html += f"<div style='width:32px; text-align:center; color:#888;'>{d}</div>"
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
                bg = "#d4edda" if count == 0 else "#FFF3CD" if count == 1 else "#F8D7DA"
                border = "1px solid #fff"
                if dt in transition_dates: border = "2px solid red"
                label = "<br>".join(occupants)
                html += f"<div style='width:32px; height:30px; background-color:{bg}; border:{border}; font-size:8px; text-align:center; display:flex; align-items:center; justify-content:center;'>{label}</div>"
            else:
                html += "<div style='width:32px;'></div>"
        html += "</div>"
    html += "</div>"
    return html

def render_annual_calendar(year, team, base_sch, night_periods, custom_schedule=None):
    html = f"<div style='font-family:monospace; font-size:10px;'>"
    html += "<div style='display:flex; gap:10px; margin-bottom:5px; font-weight:bold;'><span style='background:#d4edda; padding:2px 5px;'>T</span><span style='background:#FFC000; padding:2px 5px;'>V</span><span style='background:#FFFFE0; padding:2px 5px;'>V(R)</span><span style='background:#1E7E34; color:white; padding:2px 5px;'>Noche</span></div>"
    html += "<div style='display:flex; margin-bottom:2px;'><div style='width:30px;'></div>"
    for d in range(1, 32): html += f"<div style='width:20px; text-align:center; color:#888;'>{d}</div>"
    html += "</div>"
    
    for m_idx, mes in enumerate(MESES):
        m_num = m_idx + 1
        days_in_month = calendar.monthrange(year, m_num)[1]
        html += f"<div style='display:flex; margin-bottom:2px;'><div style='width:30px; font-weight:bold;'>{mes}</div>"
        for d in range(1, 32):
            if d <= days_in_month:
                dt = datetime.date(year, m_num, d)
                d_idx = dt.timetuple().tm_yday - 1
                
                # BLINDAJE CONTRA ERRORES DE INDICE
                if d_idx >= len(base_sch[team]): continue 

                state = base_sch[team][d_idx]
                final_val = state
                if custom_schedule: final_val = custom_schedule[d_idx]
                
                bg_color = "#eee"; text_color = "#ccc"; border = "1px solid #fff"
                
                if final_val == 'T': 
                    bg_color = "#d4edda"; text_color = "#155724"
                    if is_in_night_period(d_idx, year, night_periods):
                        bg_color = "#1E7E34"; text_color = "white"
                elif final_val == 'V': bg_color = "#FFC000"; text_color = "#000"
                elif final_val == 'V(R)': bg_color = "#FFFFE0"; text_color = "#555"
                elif final_val == 'T+': bg_color = "#ADD8E6"; text_color = "#000"
                elif final_val == 'L*': bg_color = "#E6E6FA"; text_color = "#000"
                
                if dt in get_night_transition_dates(night_periods): border = "2px solid red"
                html += f"<div style='width:20px; background-color:{bg_color}; color:{text_color}; text-align:center; border:{border}; border-radius:2px;'>{state[0]}</div>"
            else:
                html += "<div style='width:20px;'></div>"
        html += "</div>"
    html += "</div>"
    return html

# --- COBERTURAS UBH ---
def get_candidates(person_missing, roster_df, day_idx, current_schedule, year, night_periods, adjustments_log_current_day=None):
    candidates = []
    missing_role = person_missing['Rol']
    missing_turn = person_missing['Turno']
    blocked_turns = set()
    if adjustments_log_current_day:
        for coverer_name in adjustments_log_current_day:
            cov_p = roster_df[roster_df['Nombre'] == coverer_name]
            if not cov_p.empty: blocked_turns.add(cov_p.iloc[0]['Turno'])
    
    turn_exhausted_from_night = None
    if day_idx > 0:
        prev_day_idx = day_idx - 1
        if is_in_night_period(prev_day_idx, year, night_periods):
            base_sch_temp, _ = generate_base_schedule(year)
            for t in TEAMS:
                if base_sch_temp[t][prev_day_idx] == 'T':
                    turn_exhausted_from_night = t; break

    for _, candidate in roster_df.iterrows():
        if candidate['Turno'] == missing_turn: continue
        cand_status = current_schedule[candidate['Nombre']][day_idx]
        if cand_status != 'L': continue
        if candidate['Turno'] in blocked_turns: continue
        if turn_exhausted_from_night and candidate['Turno'] == turn_exhausted_from_night: continue

        is_compatible = False
        cand_role = candidate['Rol']
        cand_is_poli = candidate['Poli']
        
        if missing_role == "Capataz":
            if cand_role in ["Capataz", "2º Capataz"] or cand_is_poli: is_compatible = True
        elif missing_role == "2º Capataz":
            if cand_role in ["Capataz", "2º Capataz"] or cand_is_poli: is_compatible = True
        elif missing_role == "Bombero":
            if cand_role == "Bombero" or cand_is_poli: is_compatible = True
            
        if is_compatible: candidates.append(candidate['Nombre'])
    return candidates

def validate_and_generate_final(roster_df, requests, year, night_periods, forced_adjustments=None, strategy_key="standard"):
    if forced_adjustments is None: forced_adjustments = []
    base_schedule_turn, total_days = generate_base_schedule(year)
    final_schedule = {} 
    turn_coverage_counters = {'A': 0, 'B': 0, 'C': 0}
    person_coverage_counters = {name: 0 for name in roster_df['Nombre']}
    
    # Carga base
    base_work_load = {}
    for _, row in roster_df.iterrows():
        base_t_count = base_schedule_turn[row['Turno']].count('T')
        base_work_load[row['Nombre']] = base_t_count - 13 # Restamos vacas teóricas

    for _, row in roster_df.iterrows():
        final_schedule[row['Nombre']] = base_schedule_turn[row['Turno']].copy()

    day_vacations = {i: [] for i in range(total_days)}
    
    for req in requests:
        name = req['Nombre']
        s_idx = req['Inicio'].timetuple().tm_yday - 1
        e_idx = req['Fin'].timetuple().tm_yday - 1
        for d in range(s_idx, e_idx + 1):
            if d < total_days and final_schedule[name][d] == 'T':
                day_vacations[d].append(name)
                final_schedule[name][d] = 'V'
            elif d < total_days:
                final_schedule[name][d] = 'V(L)'

    adjustments_log = []
    for d in range(total_days):
        absent_people = day_vacations[d]
        if not absent_people: continue
        current_day_coverers = []
        absent_people.sort(key=lambda x: 0 if "Capataz" in x else 1)

        for name_missing in absent_people:
            person_row = roster_df[roster_df['Nombre'] == name_missing].iloc[0]
            candidates = get_candidates(person_row, roster_df, d, final_schedule, year, night_periods, current_day_coverers)
            if candidates:
                valid = []
                for c in candidates:
                    prev = final_schedule[c][d-1] if d>0 else 'L'
                    prev2 = final_schedule[c][d-2] if d>1 else 'L'
                    if not (prev.startswith('T') and prev2.startswith('T')): valid.append(c)
                if valid:
                    valid.sort(key=lambda x: (base_work_load[x] + person_coverage_counters[x], random.random()))
                    chosen = valid[0]
                    final_schedule[chosen][d] = f"T*({name_missing})"
                    adjustments_log.append((d, chosen, name_missing))
                    current_day_coverers.append(chosen)
                    person_coverage_counters[chosen] += 1
    
    for adj in forced_adjustments:
        d = adj['day_idx']
        p = adj['person']
        type_adj = adj['type']
        if type_adj == 'add': final_schedule[p][d] = "T+"
        elif type_adj == 'remove': final_schedule[p][d] = "L*"

    # Relleno final
    for name in roster_df['Nombre']:
        if strategy_key == 'sniper':
            sched = final_schedule[name]
            for d in range(total_days):
                if d+2 >= total_days: break
                if sched[d] == 'V':
                    if sched[d+1] == 'L': final_schedule[name][d+1] = 'V(R)'
                    if sched[d+2] == 'L': final_schedule[name][d+2] = 'V(R)'
        # Si no es sniper, el relleno ya está implícito o no se fuerza visualmente

    return final_schedule, adjustments_log, person_coverage_counters, {}

def get_work_days_count(final_schedule):
    counts = {}
    for name, sched in final_schedule.items():
        c = 0
        for s in sched:
            if s == 'T' or s.startswith('T*') or s == 'T+': c += 1
        counts[name] = c
    return counts

def find_adjustment_options(person_name, action_type, roster_df, year, night_periods, current_schedule):
    options = []
    base_sch, total_days = generate_base_schedule(year)
    vacation_counts = {i:0 for i in range(total_days)}
    for sched in current_schedule.values():
        for i, s in enumerate(sched):
            if 'V' in s: vacation_counts[i] += 1

    for d in range(total_days):
        current_status = current_schedule[person_name][d]
        if action_type == 'add':
            if current_status == 'L':
                if d > 0 and is_in_night_period(d-1, year, night_periods):
                    # Check prev turn
                    prev_t_turn = None
                    for t in TEAMS: 
                        if base_sch[t][d-1] == 'T': prev_t_turn = t
                    # Si mi turno trabajó ayer noche, no puedo ir hoy
                    # Simplificado: Si es noche, no se refuerza
                    continue 

                if vacation_counts[d] < 2:
                    d_str = (datetime.date(year, 1, 1) + timedelta(days=d)).strftime("%d/%m")
                    options.append({'day_idx': d, 'label': f"{d_str} (Libre, {vacation_counts[d]} vacs)"})
        elif action_type == 'remove':
            if current_status == 'T' or current_status.startswith('T*'):
                if vacation_counts[d] == 0:
                    d_str = (datetime.date(year, 1, 1) + timedelta(days=d)).strftime("%d/%m")
                    options.append({'day_idx': d, 'label': f"{d_str} (Trabajando)"})
    return options[:15]

def create_final_excel(schedule, roster_df, year, requests, fill_log, counters, night_periods, adjustments_log, strategy_key="standard"):
    wb = Workbook()
    s_T = PatternFill("solid", fgColor="C6EFCE"); s_V = PatternFill("solid", fgColor="FFC000")
    s_VR = PatternFill("solid", fgColor="FFFFE0"); s_Cov = PatternFill("solid", fgColor="FFC7CE")
    s_L = PatternFill("solid", fgColor="F2F2F2"); s_Night = PatternFill("solid", fgColor="A6A6A6")
    s_Extra = PatternFill("solid", fgColor="ADD8E6"); s_Free = PatternFill("solid", fgColor="E6E6FA")
    font_bold = Font(bold=True); font_red = Font(color="9C0006", bold=True)
    align_c = Alignment(horizontal="center", vertical="center")
    border_thin = Side(border_style="thin", color="000000")
    border_all = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)

    ws1 = wb.active; ws1.title = "Cuadrante"
    ws1.column_dimensions['A'].width = 20
    for i in range(2, 34): ws1.column_dimensions[get_column_letter(i)].width = 4
    
    curr_row = 1
    for t in TEAMS:
        ws1.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=32)
        cell_title = ws1.cell(curr_row, 1, f"TURNO {t}"); cell_title.font = Font(bold=True, color="FFFFFF"); cell_title.fill = PatternFill("solid", fgColor="000080"); cell_title.alignment = align_c
        curr_row += 2
        members = roster_df[roster_df['Turno'] == t].copy()
        members = members.sort_values(by=['Nombre']) # Simplificado
        for _, p in members.iterrows():
            nm = p['Nombre']
            ws1.cell(curr_row, 1, f"{nm}").font = font_bold
            for d in range(1, 32): c = ws1.cell(curr_row, d+1, d); c.alignment = align_c; c.font = font_bold; c.border = border_all; c.fill = PatternFill("solid", fgColor="E0E0E0")
            curr_row += 1
            for m_idx, mes in enumerate(MESES):
                ws1.cell(curr_row, 1, mes).font = font_bold; ws1.cell(curr_row, 1).border = border_all
                d_month = calendar.monthrange(year, m_idx+1)[1]
                for d in range(1, 32):
                    cell = ws1.cell(curr_row, d+1); cell.border = border_all; cell.alignment = align_c
                    if d <= d_month:
                        dt = datetime.date(year, m_idx+1, d); d_y = dt.timetuple().tm_yday - 1
                        st_val = schedule[nm][d_y]
                        fill = s_L; val = ""
                        if st_val == 'T': fill = s_T; val = "T"
                        elif st_val == 'V': fill = s_V; val = "V"
                        elif st_val == 'V(R)': fill = s_VR; val = "v"
                        elif st_val.startswith('T*'): 
                            fill = s_Cov; cell.font = font_red
                            val = "C" # Cover
                        elif st_val == 'T+': fill = s_Extra; val = "T+"
                        elif st_val == 'L*': fill = s_Free; val = "L"
                        if is_in_night_period(d_y, year, night_periods): fill = s_Night
                        cell.fill = fill; cell.value = val
                    else: cell.fill = PatternFill("solid", fgColor="808080")
                curr_row += 1
            curr_row += 2 
    
    # ESTADISTICAS
    ws2 = wb.create_sheet("Estadísticas")
    headers = ["Nombre", "Turno", "Puesto", "Días Trabajados", "Gastado (T)", "Coberturas (T*)"]
    ws2.append(headers)
    for _, p in roster_df.iterrows():
        name = p['Nombre']; sch = schedule[name]
        v_credits = sch.count('V')
        t_cover = counters[name]
        total_worked = 0
        for s in sch:
            if s == 'T' or s.startswith('T*') or s == 'T+': total_worked += 1
        ws2.append([name, p['Turno'], p['Rol'], total_worked, v_credits, t_cover])

    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out

# ==============================================================================
# 4. EJECUCIÓN DE LA APP
# ==============================================================================

# Inicializar datos
init_session_state()
current_requests_df, current_adjustments = load_data_from_csv()
st.session_state.raw_requests_df = current_requests_df
st.session_state.forced_adjustments = current_adjustments

current_requests = st.session_state.raw_requests_df.to_dict('records')
edited_df = st.session_state.roster_data
year_val = 2026

# Calcular estadísticas globales (LIGERO)
stats = calculate_stats(edited_df, current_requests, year_val)

# BARRA LATERAL
with st.sidebar:
    st.header("Configuración")
    year_val = st.number_input("Año", value=2026)
    
    with st.expander("Plantilla"):
        column_cfg = {
            "ID_Puesto": st.column_config.TextColumn(disabled=True),
            "Turno": st.column_config.SelectboxColumn(options=TEAMS, required=True),
            "Rol": st.column_config.SelectboxColumn(options=ROLES, required=True),
            "Poli": st.column_config.CheckboxColumn(label="¿Es Poli?", help="Puede cubrir a mandos", default=False)
        }
        edited_df = st.data_editor(st.session_state.roster_data, column_config=column_cfg, use_container_width=True, key="roster_editor")
        st.session_state.roster_data = edited_df
        
    with st.expander("Nocturnas"):
         uploaded_n = st.file_uploader("Subir Excel Nocturnas", type=['xlsx'])
         if uploaded_n:
             try:
                 df_n = pd.read_excel(uploaded_n)
                 st.session_state.nights = [] # Limpiar antes de cargar
                 for _, row in df_n.iterrows():
                     if not pd.isnull(row.iloc[0]):
                         st.session_state.nights.append((row.iloc[0].date(), row.iloc[1].date()))
                 st.success("Cargado.")
             except: st.error("Error en archivo.")
    
    if st.button("Limpiar Todos los Datos"):
        save_data_to_csv(pd.DataFrame(columns=["Nombre", "Inicio", "Fin"]), [])
        st.rerun()
    
    # Estrategia visual
    strategy_key = st.selectbox("Estrategia Visual", list(STRATEGIES.keys()), format_func=lambda x: STRATEGIES[x]['name'])

    if st.button("🎲 Generar Automático", type="primary"):
        new_reqs = auto_generate_schedule(edited_df, year_val, st.session_state.nights, strategy_key)
        st.session_state.raw_requests_df = pd.DataFrame(new_reqs)
        save_data_to_csv(st.session_state.raw_requests_df, st.session_state.forced_adjustments)
        st.session_state.computed_data = None # Invalidate cache
        st.rerun()


# PÁGINA PRINCIPAL
st.subheader("🌍 Mapa de Disponibilidad Global")
st.markdown(render_global_occupation_calendar(year_val, edited_df, current_requests, st.session_state.nights), unsafe_allow_html=True)

st.divider()

# ZONA DE EDICIÓN (SOLO ADMIN)
# LOGIN SIMPLIFICADO PARA ESTA DEMO
is_admin = True 
if is_admin:
    c1, c2 = st.columns([1, 2])
    with c1:
        st.subheader("✍️ Gestión Individual")
        names = edited_df['Nombre'].tolist()
        names.sort(key=lambda x: (0 if "Capataz" in x else 1 if "2º" in x else 2))
        
        sel_person = st.selectbox("Trabajador:", names)
        
        if sel_person:
            # Info Créditos
            p_stats = stats.get(sel_person, {'credits': 0})
            st.metric("Créditos", f"{p_stats['credits']} / 13")
            
            # Buscador
            month = st.select_slider("Mes", options=MESES, value=(MESES[0], MESES[-1]))
            if st.button("🔍 Buscar Huecos"):
                opts = get_available_options_UBH(sel_person, edited_df, current_requests, year_val, st.session_state.nights, month, strategy_key)
                st.session_state.opts_cache = opts 
            
            if 'opts_cache' in st.session_state:
                opts = st.session_state.opts_cache
                for label, slots in opts.items():
                    with st.expander(label, expanded=True):
                        for i, s in enumerate(slots):
                            # FIX KEY DUPLICATE
                            if st.button(f"➕ {s['label']}", key=f"{sel_person}_{s['start']}_{label}_{i}"):
                                new_req = {"Nombre": sel_person, "Inicio": s['start'], "Fin": s['end']}
                                current_requests.append(new_req)
                                save_data_to_csv(pd.DataFrame(current_requests), st.session_state.forced_adjustments)
                                st.session_state.raw_requests_df = pd.DataFrame(current_requests)
                                st.session_state.computed_data = None
                                st.rerun()

            st.write("Mis Vacaciones:")
            my_reqs = [r for r in current_requests if r['Nombre'] == sel_person]
            for i, r in enumerate(my_reqs):
                c_a, c_b = st.columns([4,1])
                c_a.write(f"{r['Inicio']} - {r['Fin']}")
                if c_b.button("🗑️", key=f"del_{i}"):
                    current_requests.remove(r)
                    save_data_to_csv(pd.DataFrame(current_requests), st.session_state.forced_adjustments)
                    st.session_state.raw_requests_df = pd.DataFrame(current_requests)
                    st.session_state.computed_data = None
                    st.rerun()

# --- CÁLCULO PESADO BAJO DEMANDA ---
st.divider()
st.header("⚙️ Generación Final y Descarga")

col_act, col_status = st.columns([1, 3])
if col_act.button("🔄 CALCULAR RESULTADOS", type="primary"):
    with st.spinner("Procesando coberturas y ajustes..."):
        sch, adj, count, fill = validate_and_generate_final(edited_df, current_requests, year_val, st.session_state.nights, st.session_state.forced_adjustments, strategy_key)
        excel_io = create_final_excel(sch, edited_df, year_val, current_requests, fill, count, st.session_state.nights, adj, strategy_key)
        work_days = get_work_days_count(sch)
        st.session_state.computed_data = {
            "sch": sch, "adj": adj, "work_days": work_days, "excel": excel_io
        }
    st.success("¡Calculado!")

# MOSTRAR RESULTADOS SI EXISTEN
if st.session_state.computed_data:
    data = st.session_state.computed_data
    
    cols = st.columns(4)
    for i, (name, d) in enumerate(data['work_days'].items()):
        color = "green" if MIN_WORK_DAYS <= d <= MAX_WORK_DAYS else "red"
        cols[i%4].markdown(f"**{name}**: :{color}[{d}]")
    
    c_a, c_b = st.columns(2)
    with c_a:
        p_sel = st.selectbox("Añadir Guardia a:", list(data['work_days'].keys()), key="p_add")
        opts = find_adjustment_options(p_sel, 'add', edited_df, year_val, st.session_state.nights, data['sch'])
        if opts:
            d_opt = st.selectbox("Día:", options=opts, format_func=lambda x: x['label'], key="d_add")
            if st.button("➕ Confirmar"):
                st.session_state.forced_adjustments.append({'day_idx': d_opt['day_idx'], 'person': p_sel, 'type': 'add'})
                save_data_to_csv(st.session_state.raw_requests_df, st.session_state.forced_adjustments)
                st.session_state.computed_data = None 
                st.rerun()
    
    with c_b:
        p_rem = st.selectbox("Quitar Guardia a:", list(data['work_days'].keys()), key="p_rem")
        opts_r = find_adjustment_options(p_rem, 'remove', edited_df, year_val, st.session_state.nights, data['sch'])
        if opts_r:
            d_opt_r = st.selectbox("Día:", options=opts_r, format_func=lambda x: x['label'], key="d_rem")
            if st.button("➖ Confirmar"):
                st.session_state.forced_adjustments.append({'day_idx': d_opt_r['day_idx'], 'person': p_rem, 'type': 'remove'})
                save_data_to_csv(st.session_state.raw_requests_df, st.session_state.forced_adjustments)
                st.session_state.computed_data = None
                st.rerun()

    st.download_button("📥 Descargar Excel", data['excel'], "Cuadrante_Final.xlsx")
else:
    st.info("Pulsa el botón para calcular el cuadrante completo.")
