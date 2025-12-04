import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Gestor de Rutas Completo", layout="wide")

st.title("📝 Gestor de Partes de Ruta (Completo)")

# --- 1. DATOS INTEGRADOS ---
# Estructura completa basada en tu archivo original
DATOS_RUTAS = [
    {'layer': 'R1-1', 'ESTADO': 'CERRADA', 'HIELO': None, 'DISTANCIA': 3.21, 'PASADAS': 2.0, 'RECURSO': None, 'ACTUACION': 'SALERO', 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '1º', 'TOTAL_KM': 0.0, 'PERSONAL': 'vol morella', 'RECORRIDO': 'CV108 Ballestar a CV107 La Pobla de Benifassar', 'H.INICIO': 1000.0, 'H.FIN': 1300.0, 'TIEMPO': 3.0},
    {'layer': 'R1-2', 'ESTADO': 'CADENAS', 'HIELO': 'HIELO', 'DISTANCIA': 15.73, 'PASADAS': 3.0, 'RECURSO': 'BRP531', 'ACTUACION': 'CUCHILLA', 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '2º', 'TOTAL_KM': 0.0, 'PERSONAL': 'ubf morella', 'RECORRIDO': 'CV105 cruce CV107 a La Senia', 'H.INICIO': 1130.0, 'H.FIN': 1300.0, 'TIEMPO': 130.0},
    {'layer': 'R1-3', 'ESTADO': 'ABIERTA', 'HIELO': None, 'DISTANCIA': 11.08, 'PASADAS': 4.0, 'RECURSO': None, 'ACTUACION': 'CUÑA', 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '3º', 'TOTAL_KM': 0.0, 'PERSONAL': None, 'RECORRIDO': 'CV106 CV105 a Fredes', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R1-4', 'ESTADO': 'CADENAS', 'HIELO': None, 'DISTANCIA': 5.08, 'PASADAS': 5.0, 'RECURSO': None, 'ACTUACION': 'CUÑA SALERO', 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '4º', 'TOTAL_KM': 0.0, 'PERSONAL': None, 'RECORRIDO': 'Camino rural Fredes a Boixar', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R1-5', 'ESTADO': 'CADENAS', 'HIELO': 'PLACAS', 'DISTANCIA': 19.27, 'PASADAS': 3.0, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '5º', 'TOTAL_KM': 0.0, 'PERSONAL': None, 'RECORRIDO': 'CV109 cruce CV105 Boixar a CV105 cruce CV110 Herbes.', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R1-6', 'ESTADO': 'CERRADA', 'HIELO': None, 'DISTANCIA': 7.57, 'PASADAS': 2.0, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '6º', 'TOTAL_KM': 0.0, 'PERSONAL': None, 'RECORRIDO': 'CV109 cruce CV105 Boixar a Coratxar', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R1-7', 'ESTADO': 'HIELO', 'HIELO': None, 'DISTANCIA': 8.43, 'PASADAS': 4.0, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '7º', 'TOTAL_KM': 0.0, 'PERSONAL': None, 'RECORRIDO': 'CV105 cruce CV109 Boixar a CV107', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R1-8', 'ESTADO': 'PLACAS', 'HIELO': None, 'DISTANCIA': 7.3, 'PASADAS': 50.0, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '8º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'La Senia a Cases del Riu a Rosell por CV100', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    # Se han simplificado el resto de filas vacías para no hacer el código infinito,
    # pero el sistema funciona igual para nuevas entradas o si pegas el resto de tu lista aquí.
    {'layer': 'R11-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 2.14, 'PASADAS': 3.0, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 11, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Accesos Castellfort', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R3-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 12.46, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 3, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV14 Morella a Forcall', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
]
# NOTA: Puedes añadir el resto de tu lista DATOS_RUTAS anterior aquí dentro si lo necesitas.

@st.cache_data
def obtener_dataframes():
    return pd.DataFrame(DATOS_RUTAS)

df = obtener_dataframes()

# Lista de todas las columnas que queremos en el resultado final
COLUMNAS_FINALES = [
    'layer', 'RUTA', 'TRAMO', 'RECORRIDO', # Fijas
    'ESTADO', 'HIELO', 'ACTUACION',        # Desplegables
    'DISTANCIA', 'PASADAS', 'TOTAL_KM',    # Numéricas / Texto
    'RECURSO', 'PERSONAL', 'FUNDENTE',     # Recursos
    'FECHA', 'H.INICIO', 'H.FIN', 'TIEMPO' # Tiempos
]

# Inicializar sesión para guardar datos añadidos
if 'seleccionados' not in st.session_state:
    st.session_state['seleccionados'] = pd.DataFrame(columns=COLUMNAS_FINALES)

# --- FUNCIONES AUXILIARES ---
def safe_value(val):
    """Devuelve el valor como string si existe, o vacío si es nan/None"""
    if pd.isna(val) or val == "None":
        return ""
    return str(val)

# --- 2. ÁREA DE SELECCIÓN Y EDICIÓN ---

# Opciones para los desplegables principales
opciones_estado = [x for x in df['ESTADO'].unique() if pd.notna(x)] + ["SIN INCIDENCIAS"]
opciones_hielo = [x for x in df['HIELO'].unique() if pd.notna(x)] + ["NO"]
opciones_actuacion = [x for x in df['ACTUACION'].unique() if pd.notna(x)] + ["NINGUNA"]

with st.container(border=True):
    st.subheader("1. Seleccionar Tramo")
    
    # Creamos label combo para buscar mejor
    df['label_combo'] = df['layer'] + " | " + df['RECORRIDO']
    seleccion_usuario = st.selectbox("Elige la Ruta/Capa:", df['label_combo'].tolist())
    
    # Recuperamos fila original
    fila_datos = df[df['label_combo'] == seleccion_usuario].iloc[0]
    
    st.divider()
    
    with st.form("form_ruta_completo"):
        st.subheader("2. Editar Datos")
        
        # --- BLOQUE 1: INFORMACIÓN FIJA ---
        c1, c2, c3 = st.columns([1, 1, 3])
        c1.info(f"**Ruta:** {fila_datos['RUTA']}")
        c2.info(f"**Tramo:** {fila_datos['TRAMO']}")
        c3.info(f"**Recorrido:** {fila_datos['RECORRIDO']}")
        
        st.markdown("---")
        
        # --- BLOQUE 2: ESTADOS (Desplegables) ---
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
        
        # --- BLOQUE 3: DATOS NUMÉRICOS Y RECURSOS (Rellenables manualmente) ---
        # Fila A
        ca1, ca2, ca3, ca4 = st.columns(4)
        val_distancia = ca1.text_input("Distancia", value=safe_value(fila_datos.get('DISTANCIA')))
        val_pasadas = ca2.text_input("Pasadas", value=safe_value(fila_datos.get('PASADAS')))
        val_totalkm = ca3.text_input("Total Km", value=safe_value(fila_datos.get('TOTAL_KM')))
        val_recurso = ca4.text_input("Recurso", value=safe_value(fila_datos.get('RECURSO')))
        
        # Fila B
        cb1, cb2, cb3, cb4 = st.columns(4)
        val_personal = cb1.text_input("Personal", value=safe_value(fila_datos.get('PERSONAL')))
        val_fundente = cb2.text_input("Fundente", value=safe_value(fila_datos.get('FUNDENTE')))
        val_fecha = cb3.text_input("Fecha (AAAA-MM-DD)", value=safe_value(fila_datos.get('FECHA')))
        val_tiempo = cb4.text_input("Tiempo Total", value=safe_value(fila_datos.get('TIEMPO')))
        
        # Fila C (Horas)
        cc1, cc2 = st.columns(2)
        val_hinicio = cc1.text_input("Hora Inicio", value=safe_value(fila_datos.get('H.INICIO')))
        val_hfin = cc2.text_input("Hora Fin", value=safe_value(fila_datos.get('H.FIN')))

        st.markdown("<br>", unsafe_allow_html=True)
        submit_btn = st.form_submit_button("➕ Añadir Registro", type="primary")
        
        if submit_btn:
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
                'TOTAL_KM': val_totalkm,
                'RECURSO': val_recurso,
                'PERSONAL': val_personal,
                'FUNDENTE': val_fundente,
                'FECHA': val_fecha,
                'TIEMPO': val_tiempo,
                'H.INICIO': val_hinicio,
                'H.FIN': val_hfin
            }
            # Añadir a session_state
            st.session_state['seleccionados'] = pd.concat([st.session_state['seleccionados'], pd.DataFrame([nueva_fila])], ignore_index=True)
            st.success(f"Añadido: {fila_datos['layer']}")

# --- 3. TABLA Y DESCARGA ---
st.subheader("📋 Tabla Formada")

df_resultados = st.session_state['seleccionados']

if not df_resultados.empty:
    st.dataframe(df_resultados, use_container_width=True, hide_index=True)
    
    # Preparar CSV para descarga
    csv_buffer = io.BytesIO()
    df_resultados.to_csv(csv_buffer, index=False, encoding='utf-8-sig') # utf-8-sig para que Excel lea bien las tildes
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
    st.info("Utiliza el formulario de arriba para añadir tramos a la tabla.")
