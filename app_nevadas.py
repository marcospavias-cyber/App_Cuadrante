import streamlit as st
import pandas as pd
import io
from datetime import datetime, time

# Configuración de la página
st.set_page_config(page_title="Gestor de Rutas Pro", layout="wide")

st.title("📝 Gestor de Partes de Ruta (Automático)")

# --- 1. DATOS INTEGRADOS ---
DATOS_RUTAS = [
    {'layer': 'R1-1', 'ESTADO': 'CERRADA', 'HIELO': None, 'DISTANCIA': 3.21, 'PASADAS': 2.0, 'RECURSO': None, 'ACTUACION': 'SALERO', 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '1º', 'TOTAL_KM': 0.0, 'PERSONAL': 'vol morella', 'RECORRIDO': 'CV108 Ballestar a CV107 La Pobla de Benifassar', 'H.INICIO': 1000.0, 'H.FIN': 1300.0, 'TIEMPO': 3.0},
    {'layer': 'R1-2', 'ESTADO': 'CADENAS', 'HIELO': 'HIELO', 'DISTANCIA': 15.73, 'PASADAS': 3.0, 'RECURSO': 'BRP531', 'ACTUACION': 'CUCHILLA', 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '2º', 'TOTAL_KM': 0.0, 'PERSONAL': 'ubf morella', 'RECORRIDO': 'CV105 cruce CV107 a La Senia', 'H.INICIO': 1130.0, 'H.FIN': 1300.0, 'TIEMPO': 130.0},
    {'layer': 'R1-3', 'ESTADO': 'ABIERTA', 'HIELO': None, 'DISTANCIA': 11.08, 'PASADAS': 4.0, 'RECURSO': None, 'ACTUACION': 'CUÑA', 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '3º', 'TOTAL_KM': 0.0, 'PERSONAL': None, 'RECORRIDO': 'CV106 CV105 a Fredes', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R1-4', 'ESTADO': 'CADENAS', 'HIELO': None, 'DISTANCIA': 5.08, 'PASADAS': 5.0, 'RECURSO': None, 'ACTUACION': 'CUÑA SALERO', 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '4º', 'TOTAL_KM': 0.0, 'PERSONAL': None, 'RECORRIDO': 'Camino rural Fredes a Boixar', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    # ... (Puedes añadir el resto de tu lista aquí)
]

@st.cache_data
def obtener_dataframes():
    return pd.DataFrame(DATOS_RUTAS)

df = obtener_dataframes()

# Lista de columnas finales
COLUMNAS_FINALES = [
    'layer', 'RUTA', 'TRAMO', 'RECORRIDO', 
    'ESTADO', 'HIELO', 'ACTUACION',
    'DISTANCIA', 'PASADAS', 'TOTAL_KM',
    'RECURSO', 'PERSONAL', 'FUNDENTE',
    'FECHA', 'H.INICIO', 'H.FIN', 'TIEMPO'
]

# Inicializar sesión
if 'seleccionados' not in st.session_state:
    st.session_state['seleccionados'] = pd.DataFrame(columns=COLUMNAS_FINALES)

# --- FUNCIONES DE AYUDA PARA FORMATOS ---

def parsear_fecha(valor):
    """Convierte el string 'YYYY-MM-DD' a objeto date. Si falla, devuelve hoy."""
    if pd.isna(valor) or valor == "None":
        return datetime.now().date()
    try:
        return datetime.strptime(str(valor), '%Y-%m-%d').date()
    except:
        return datetime.now().date()

def parsear_hora(valor):
    """Convierte float tipo 1300.0 a objeto time (13:00). Si falla, devuelve 00:00."""
    if pd.isna(valor) or valor == "None":
        return time(9, 0) # Hora por defecto: 09:00
    try:
        # Asumimos formato HHMM (ej. 1300.0 -> 13:00)
        val_int = int(float(valor))
        hora = val_int // 100
        minuto = val_int % 100
        # Validar rangos
        if hora > 23: hora = 0
        if minuto > 59: minuto = 0
        return time(hora, minuto)
    except:
        return time(9, 0)

def safe_float(val):
    """Intenta convertir a float para los campos numéricos"""
    try:
        return float(val)
    except:
        return 0.0

def safe_str(val):
    if pd.isna(val) or val == "None": return ""
    return str(val)


# --- ÁREA DE EDICIÓN ---

opciones_estado = [x for x in df['ESTADO'].unique() if pd.notna(x)] + ["SIN INCIDENCIAS"]
opciones_hielo = [x for x in df['HIELO'].unique() if pd.notna(x)] + ["NO"]
opciones_actuacion = [x for x in df['ACTUACION'].unique() if pd.notna(x)] + ["NINGUNA"]

with st.container(border=True):
    st.subheader("1. Seleccionar Tramo")
    
    # Label Combo
    df['label_combo'] = df['layer'] + " | " + df['RECORRIDO']
    seleccion_usuario = st.selectbox("Elige la Ruta/Capa:", df['label_combo'].tolist())
    
    # Datos Fila
    fila_datos = df[df['label_combo'] == seleccion_usuario].iloc[0]
    
    st.divider()
    
    with st.form("form_ruta_completo"):
        st.subheader("2. Editar Datos")
        
        # INFO FIJA
        c1, c2, c3 = st.columns([1, 1, 3])
        c1.info(f"**Ruta:** {fila_datos['RUTA']}")
        c2.info(f"**Tramo:** {fila_datos['TRAMO']}")
        c3.info(f"**Recorrido:** {fila_datos['RECORRIDO']}")
        
        st.markdown("---")
        
        # DESPLEGABLES
        col_e1, col_e2, col_e3 = st.columns(3)
        
        with col_e1:
            idx_est = 0
            if pd.notna(fila_datos['ESTADO']) and fila_datos['ESTADO'] in opciones_estado:
                idx_est = opciones_estado.index(fila_datos['ESTADO'])
            val_estado = st.selectbox("Estado", opciones_estado, index=idx_est)
            
        with col_e2:
            idx_hielo = 0
            if pd.notna(fila_datos['HIELO']) and fila_datos['HIELO'] in opciones_hielo:
                idx_hielo = opciones_hielo.index(fila_datos['HIELO'])
            else:
                if "NO" in opciones_hielo: idx_hielo = opciones_hielo.index("NO")
            val_hielo = st.selectbox("Hielo", opciones_hielo, index=idx_hielo)
            
        with col_e3:
            idx_act = 0
            if pd.notna(fila_datos['ACTUACION']) and fila_datos['ACTUACION'] in opciones_actuacion:
                idx_act = opciones_actuacion.index(fila_datos['ACTUACION'])
            else:
                if "NINGUNA" in opciones_actuacion: idx_act = opciones_actuacion.index("NINGUNA")
            val_actuacion = st.selectbox("Actuación", opciones_actuacion, index=idx_act)

        st.markdown("---")
        
        # NUMÉRICOS Y CÁLCULO
        st.markdown("##### Datos Kilométricos")
        ca1, ca2, ca3, ca4 = st.columns(4)
        
        # Distancia (Number Input para poder calcular)
        val_distancia = ca1.number_input("Distancia (km)", value=safe_float(fila_datos.get('DISTANCIA')), step=0.1, format="%.2f")
        
        # Pasadas (Number Input)
        val_pasadas = ca2.number_input("Pasadas", value=safe_float(fila_datos.get('PASADAS')), step=1.0, format="%.0f")
        
        # TOTAL KM (Deshabilitado - Se calcula al enviar)
        # Mostramos un placeholder visual, el valor real se calcula abajo
        ca3.text_input("Total Km (Auto)", value="Se calculará al guardar", disabled=True)
        
        val_recurso = ca4.text_input("Recurso", value=safe_str(fila_datos.get('RECURSO')))
        
        st.markdown("##### Fechas y Tiempos")
        cb1, cb2, cb3, cb4 = st.columns(4)
        
        # FECHA (Date Input)
        fecha_default = parsear_fecha(fila_datos.get('FECHA'))
        val_fecha = cb1.date_input("Fecha", value=fecha_default)
        
        # HORAS (Time Input)
        hora_ini_default = parsear_hora(fila_datos.get('H.INICIO'))
        val_hinicio = cb2.time_input("Hora Inicio", value=hora_ini_default)
        
        hora_fin_default = parsear_hora(fila_datos.get('H.FIN'))
        val_hfin = cb3.time_input("Hora Fin", value=hora_fin_default)
        
        # TIEMPO TOTAL (Calculable o manual)
        val_tiempo = cb4.text_input("Tiempo Total (min)", value=safe_str(fila_datos.get('TIEMPO')))
        
        # OTROS
        st.markdown("##### Otros")
        cc1, cc2 = st.columns(2)
        val_personal = cc1.text_input("Personal", value=safe_str(fila_datos.get('PERSONAL')))
        val_fundente = cc2.text_input("Fundente", value=safe_str(fila_datos.get('FUNDENTE')))

        st.markdown("<br>", unsafe_allow_html=True)
        submit_btn = st.form_submit_button("➕ Añadir y Calcular", type="primary")
        
        if submit_btn:
            # --- CÁLCULO AUTOMÁTICO ---
            calculo_total_km = val_distancia * val_pasadas
            
            # --- CONVERTIR FECHAS A TEXTO PARA CSV ---
            # Streamlit devuelve objetos date/time, los pasamos a string para el CSV
            str_fecha = val_fecha.strftime('%Y-%m-%d')
            str_hinicio = val_hinicio.strftime('%H:%M')
            str_hfin = val_hfin.strftime('%H:%M')
            
            nueva_fila = {
                'layer': fila_datos['layer'],
                'RUTA': fila_datos['RUTA'],
                'TRAMO': fila_datos['TRAMO'],
                'RECORRIDO': fila_datos['RECORRIDO'],
                'ESTADO': val_estado,
                'HIELO': val_hielo,
                'ACTUACION': val_actuacion,
                'DISTANCIA': val_distancia,
                'PASADAS': val_pasadas,
                'TOTAL_KM': calculo_total_km,  # <--- AQUÍ SE GUARDA EL CÁLCULO
                'RECURSO': val_recurso,
                'PERSONAL': val_personal,
                'FUNDENTE': val_fundente,
                'FECHA': str_fecha,
                'H.INICIO': str_hinicio,
                'H.FIN': str_hfin,
                'TIEMPO': val_tiempo
            }
            
            st.session_state['seleccionados'] = pd.concat([st.session_state['seleccionados'], pd.DataFrame([nueva_fila])], ignore_index=True)
            st.success(f"✅ Añadido: {fila_datos['layer']} | Total Km calculados: {calculo_total_km:.2f}")

# --- TABLA Y DESCARGA ---
st.subheader("📋 Tabla Formada")

df_resultados = st.session_state['seleccionados']

if not df_resultados.empty:
    st.dataframe(df_resultados, use_container_width=True, hide_index=True)
    
    csv_buffer = io.BytesIO()
    df_resultados.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
    csv_bytes = csv_buffer.getvalue()
    
    col_d1, col_d2 = st.columns([1, 5])
    with col_d1:
        st.download_button(
            label="💾 Descargar CSV",
            data=csv_bytes,
            file_name="parte_rutas.csv",
            mime="text/csv",
        )
    with col_d2:
        if st.button("🗑️ Limpiar Tabla"):
            st.session_state['seleccionados'] = pd.DataFrame(columns=COLUMNAS_FINALES)
            st.rerun()
else:
    st.info("Utiliza el formulario para añadir tramos.")
