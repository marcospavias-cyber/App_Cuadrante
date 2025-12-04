import streamlit as st
import pandas as pd

# Configuración de la página
st.set_page_config(page_title="Gestor de Rutas", layout="wide")

st.title("📝 Gestor de Partes de Ruta")

# --- 1. DATOS INTEGRADOS ---
DATOS_RUTAS = [
    {'layer': 'R1-1', 'ESTADO': 'CERRADA', 'HIELO': None, 'ACTUACION': 'SALERO', 'RUTA': 1, 'TRAMO': '1º', 'RECORRIDO': 'CV108 Ballestar a CV107 La Pobla de Benifassar'},
    {'layer': 'R1-2', 'ESTADO': 'CADENAS', 'HIELO': 'HIELO', 'ACTUACION': 'CUCHILLA', 'RUTA': 1, 'TRAMO': '2º', 'RECORRIDO': 'CV105 cruce CV107 a La Senia'},
    {'layer': 'R1-3', 'ESTADO': 'ABIERTA', 'HIELO': None, 'ACTUACION': 'CUÑA', 'RUTA': 1, 'TRAMO': '3º', 'RECORRIDO': 'CV106 CV105 a Fredes'},
    {'layer': 'R1-4', 'ESTADO': 'CADENAS', 'HIELO': None, 'ACTUACION': 'CUÑA SALERO', 'RUTA': 1, 'TRAMO': '4º', 'RECORRIDO': 'Camino rural Fredes a Boixar'},
    {'layer': 'R1-5', 'ESTADO': 'CADENAS', 'HIELO': 'PLACAS', 'ACTUACION': None, 'RUTA': 1, 'TRAMO': '5º', 'RECORRIDO': 'CV109 cruce CV105 Boixar a CV105 cruce CV110 Herbes.'},
    {'layer': 'R1-6', 'ESTADO': 'CERRADA', 'HIELO': None, 'ACTUACION': None, 'RUTA': 1, 'TRAMO': '6º', 'RECORRIDO': 'CV109 cruce CV105 Boixar a Coratxar'},
    {'layer': 'R1-7', 'ESTADO': 'HIELO', 'HIELO': None, 'ACTUACION': None, 'RUTA': 1, 'TRAMO': '7º', 'RECORRIDO': 'CV105 cruce CV109 Boixar a CV107'},
    {'layer': 'R1-8', 'ESTADO': 'PLACAS', 'HIELO': None, 'ACTUACION': None, 'RUTA': 1, 'TRAMO': '8º', 'RECORRIDO': 'La Senia a Cases del Riu a Rosell por CV100'},
    {'layer': 'R11-1', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 11, 'TRAMO': '1º', 'RECORRIDO': 'Accesos Castellfort'},
    {'layer': 'R11-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 11, 'TRAMO': '2º', 'RECORRIDO': 'CV124 de CV126 a CV12'},
    {'layer': 'R11-3', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 11, 'TRAMO': '3º', 'RECORRIDO': 'CV126 Castellfort a Vilafranca'},
    {'layer': 'R11-4', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 11, 'TRAMO': '4º', 'RECORRIDO': 'CV15 Vilafranca a Llosar a Pista rural Vilafranca a Portell de Morella'},
    {'layer': 'R11-5', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 11, 'TRAMO': '5º', 'RECORRIDO': 'Pista rual Portell de Morella a la Cuba cruce CV120'},
    {'layer': 'R11-6', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 11, 'TRAMO': '6º', 'RECORRIDO': 'CV120 desde La Cuba a cruce con Todolella'},
    {'layer': 'R11-7', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 11, 'TRAMO': '7º', 'RECORRIDO': 'Pista rural La Mata Todolella desde CV120 a Cinctorres'},
    {'layer': 'R11-8', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 11, 'TRAMO': '8º', 'RECORRIDO': 'CV124 de Cinctorres a Castellfort'},
    {'layer': 'R12-1', 'ESTADO': 'CADENAS', 'HIELO': None, 'ACTUACION': None, 'RUTA': 12, 'TRAMO': '1º', 'RECORRIDO': 'CV175 Villahermosa a CV190'},
    {'layer': 'R12-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 12, 'TRAMO': '2º', 'RECORRIDO': 'CV176 de CV175 a CV190'},
    {'layer': 'R12-3', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 12, 'TRAMO': '3º', 'RECORRIDO': 'CV175 de Villahermosa a Puertomingalvo'},
    {'layer': 'R15-1', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 15, 'TRAMO': '1º', 'RECORRIDO': 'CV207 Ballacas a CV209 Pina de Montalgrao'},
    {'layer': 'R15-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 15, 'TRAMO': '2º', 'RECORRIDO': 'CV207 de CV209 Pina de Montalgrao a Villanueva de Viver'},
    {'layer': 'R15-3', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 15, 'TRAMO': '3º', 'RECORRIDO': 'CV207 Villanueva de Vives a CV20 por Fuente la Reina'},
    {'layer': 'R15-4', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 15, 'TRAMO': '4º', 'RECORRIDO': 'CV208 de CV207 a CV20 por Los Pastores'},
    {'layer': 'R15-5', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 15, 'TRAMO': '5º', 'RECORRIDO': 'CV2093 de Pina de Montalgrao a CV2092'},
    {'layer': 'R16-1', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 16, 'TRAMO': '1º', 'RECORRIDO': 'CV200 de Segorbe a Aín'},
    {'layer': 'R16-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 16, 'TRAMO': '2º', 'RECORRIDO': 'CV223 Ain a Alcudia a CV215 a Algimia a CV213 Matet a P.Villamalur'},
    {'layer': 'R17-1', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 17, 'TRAMO': '1º', 'RECORRIDO': 'CV203 de Caudiel a CV205'},
    {'layer': 'R17-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 17, 'TRAMO': '2º', 'RECORRIDO': 'CV 205 cruce CV203 a cruce CV202'},
    {'layer': 'R17-3', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 17, 'TRAMO': '3º', 'RECORRIDO': 'CV202 de cruce CV205 a Villamalur'},
    {'layer': 'R17-4', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 17, 'TRAMO': '4º', 'RECORRIDO': 'CV205 cruce CV202 a Cruce CV201 y CV201 cruce CV202 a Artesa'},
    {'layer': 'R2-1', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 2, 'TRAMO': '1º', 'RECORRIDO': 'CV111 de N232 a Vallibona'},
    {'layer': 'R2-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 2, 'TRAMO': '2º', 'RECORRIDO': 'Pista rural dels Llivis entre 232 y CV12'},
    {'layer': 'R2-3', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 2, 'TRAMO': '3º', 'RECORRIDO': 'Pista rural La Llacua acceso desde CV12'},
    {'layer': 'R2-4', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 2, 'TRAMO': '4º', 'RECORRIDO': 'Pista rural de la Cana desde CV12 a CV124'},
    {'layer': 'R2-5', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 2, 'TRAMO': '5º', 'RECORRIDO': 'Pista rural La Vega del Moll de CV125 a CV12'},
    {'layer': 'R23-1', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 23, 'TRAMO': '1º', 'RECORRIDO': 'CV117 de Morella a Xiva de Morella'},
    {'layer': 'R23-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 23, 'TRAMO': '2º', 'RECORRIDO': 'CV105 CV110 de N232 a lim provincia Teruel Herbes'},
    {'layer': 'R23-3', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 23, 'TRAMO': '3º', 'RECORRIDO': 'CV105 de cruce CV110 a Castell de Cabres'},
    {'layer': 'R23-4', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 23, 'TRAMO': '4º', 'RECORRIDO': 'CV1050 Acceso a Herbeset'},
    {'layer': 'R26-1', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 26, 'TRAMO': '1º', 'RECORRIDO': 'CV193 de Lucena del Cid a Argelita'},
    {'layer': 'R26-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 26, 'TRAMO': '2º', 'RECORRIDO': 'CV194 de Argelita a CV1970 Giraba de Abajo'},
    {'layer': 'R26-3', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 26, 'TRAMO': '3º', 'RECORRIDO': 'CV1970 de CV194 a Giraba'},
    {'layer': 'R26-4', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 26, 'TRAMO': '4º', 'RECORRIDO': 'CV194 de Giraba a lim provincia en Cortes de Arenoso'},
    {'layer': 'R26-5', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 26, 'TRAMO': '5º', 'RECORRIDO': 'Camino Mas de la Llosa de CV190 a limite provincia'},
    {'layer': 'R27-1', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 27, 'TRAMO': '1º', 'RECORRIDO': 'CV171 de Adzeneta a Xodos'},
    {'layer': 'R27-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 27, 'TRAMO': '2º', 'RECORRIDO': 'CV170 y CV169 de Adzeneta a Collado CV170 por Benafigos'},
    {'layer': 'R27-3', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 27, 'TRAMO': '3º', 'RECORRIDO': 'CV172 Acceso Benafigos'},
    {'layer': 'R3-1', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 3, 'TRAMO': '1º', 'RECORRIDO': 'Accesos a Morella son CV1160 y CV1170'},
    {'layer': 'R3-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 3, 'TRAMO': '2º', 'RECORRIDO': 'CV14 Morella a Forcall'},
    {'layer': 'R3-3', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 3, 'TRAMO': '3º', 'RECORRIDO': 'Forcall a Todolella por CV120 y CV122'},
    {'layer': 'R3-4', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 3, 'TRAMO': '4º', 'RECORRIDO': 'Todolella a Olocau del Rey por CV120 y CV121'},
    {'layer': 'R3-5', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 3, 'TRAMO': '5º', 'RECORRIDO': 'Olocau del Rey a Tronchon por CV123 y A226'},
    {'layer': 'R3-6', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 3, 'TRAMO': '6º', 'RECORRIDO': 'Olocau del Rey a Bordon por CV121 y TE'},
    {'layer': 'R3-7', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 3, 'TRAMO': '7º', 'RECORRIDO': 'Olocau del Rey a Todolella por CV 122'},
    {'layer': 'R4-1', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 4, 'TRAMO': '1º', 'RECORRIDO': 'San Mateo a Xert por CV132'},
    {'layer': 'R4-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 4, 'TRAMO': '2º', 'RECORRIDO': 'Rosell a Bel por CV104'},
    {'layer': 'R5-1', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 5, 'TRAMO': '1º', 'RECORRIDO': 'Accesos a Morella CV1160 y CV1170'},
    {'layer': 'R5-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 5, 'TRAMO': '2º', 'RECORRIDO': 'CV125 de CV14 a Cinctorres'},
    {'layer': 'R5-3', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 5, 'TRAMO': '3º', 'RECORRIDO': 'Morella a Forcal de CV14 con CV125 a CV124 Forcall'},
    {'layer': 'R5-4', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 5, 'TRAMO': '4º', 'RECORRIDO': 'CV14 desde CV124 por Villores CV119 Ortells CV1171 y Palanques CV118'},
    {'layer': 'R5-5', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 5, 'TRAMO': '5º', 'RECORRIDO': 'Camino Rural Vega del Moll La Mina de CV125 a CV12'},
    {'layer': 'R5-6', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 5, 'TRAMO': '6º', 'RECORRIDO': 'Pista Rural Sierra de Palos'},
    {'layer': 'R5-7', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 5, 'TRAMO': '7º', 'RECORRIDO': 'Pista de la Carcellera'},
    {'layer': 'R7-1', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 7, 'TRAMO': '1º', 'RECORRIDO': 'Accesos Vilafranca'},
    {'layer': 'R7-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 7, 'TRAMO': '2º', 'RECORRIDO': 'CV15 de Vilafranca a Ares por acceso CV1260'},
    {'layer': 'R7-3', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 7, 'TRAMO': '3º', 'RECORRIDO': 'CV124 de CV12 a Castellfort'},
    {'layer': 'R7-4', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 7, 'TRAMO': '4º', 'RECORRIDO': 'CV126 Castellfort a Villafranca'},
    {'layer': 'R7-5', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 7, 'TRAMO': '5º', 'RECORRIDO': 'Pista rural Vilafranca a portell de Morella'},
    {'layer': 'R8-1', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 8, 'TRAMO': '1º', 'RECORRIDO': 'CV166 de Benasal a CV15'},
    {'layer': 'R8-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 8, 'TRAMO': '2º', 'RECORRIDO': 'CV166 de Benassal a Sant Pau'},
    {'layer': 'R8-3', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 8, 'TRAMO': '3º', 'RECORRIDO': 'CV166 de CV15 a Vilar de Canes y CV168 Vilardecanes a CV15'},
    {'layer': 'R8-4', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 8, 'TRAMO': '4º', 'RECORRIDO': 'CV15 de CV165 a CV166'},
    {'layer': 'R8-5', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 8, 'TRAMO': '5º', 'RECORRIDO': 'CV128 CV15 a Cati a L\' Avella CV1270 y N232'},
    {'layer': 'R8-6', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 8, 'TRAMO': '6º', 'RECORRIDO': 'CV167 Benasal hasta final cruce CV15'},
    {'layer': 'R9-1', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 9, 'TRAMO': '1º', 'RECORRIDO': 'CV170 desde CV1720 a Vistabella poblacion'},
    {'layer': 'R9-2', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 9, 'TRAMO': '2º', 'RECORRIDO': 'Pista rural Vistabella a Sant Joan de Penyagolosa'},
    {'layer': 'R9-3', 'ESTADO': None, 'HIELO': None, 'ACTUACION': None, 'RUTA': 9, 'TRAMO': '3º', 'RECORRIDO': 'CV 1720 de cruce CV170 a Sierra del Boi'}
]

# Inicializamos DataFrames
@st.cache_data
def obtener_dataframes():
    df_base = pd.DataFrame(DATOS_RUTAS)
    return df_base

df = obtener_dataframes()

# Inicializar sesión para guardar datos añadidos
if 'seleccionados' not in st.session_state:
    st.session_state['seleccionados'] = pd.DataFrame(columns=['layer', 'RUTA', 'TRAMO', 'ESTADO', 'HIELO', 'ACTUACION', 'RECORRIDO'])

# --- 2. ÁREA DE SELECCIÓN Y EDICIÓN ---

# Opciones para los desplegables (sacadas de todos los valores posibles del CSV)
opciones_estado = [x for x in df['ESTADO'].unique() if pd.notna(x)] + ["SIN INCIDENCIAS"]
opciones_hielo = [x for x in df['HIELO'].unique() if pd.notna(x)] + ["NO"]
opciones_actuacion = [x for x in df['ACTUACION'].unique() if pd.notna(x)] + ["NINGUNA"]

with st.container(border=True):
    st.subheader("1. Seleccionar Tramo")
    
    # Selector principal (Columna A - layer)
    # Creamos una etiqueta amigable que combine Layer + Recorrido
    df['label_combo'] = df['layer'] + " | " + df['RECORRIDO']
    
    seleccion_usuario = st.selectbox("Elige la Ruta/Capa:", df['label_combo'].tolist())
    
    # Recuperamos los datos de la fila seleccionada
    fila_datos = df[df['label_combo'] == seleccion_usuario].iloc[0]
    
    st.divider()
    
    # Formulario para rellenar datos
    with st.form("form_ruta"):
        st.subheader("2. Rellenar Estado")
        
        # Columnas fijas (Información no editable)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.info(f"**Ruta:** {fila_datos['RUTA']}")
        with col2:
            st.info(f"**Tramo:** {fila_datos['TRAMO']}")
        with col3:
            st.info(f"**Recorrido:** {fila_datos['RECORRIDO']}")
            
        st.markdown("---")
        
        # Columnas editables (Desplegables)
        c_estado, c_hielo, c_actuacion = st.columns(3)
        
        with c_estado:
            # Pre-seleccionar valor si existe, sino el primero
            idx_est = 0
            if pd.notna(fila_datos['ESTADO']) and fila_datos['ESTADO'] in opciones_estado:
                idx_est = opciones_estado.index(fila_datos['ESTADO'])
            val_estado = st.selectbox("Estado (Col B)", opciones_estado, index=idx_est)
            
        with c_hielo:
            idx_hielo = 0
            if pd.notna(fila_datos['HIELO']) and fila_datos['HIELO'] in opciones_hielo:
                idx_hielo = opciones_hielo.index(fila_datos['HIELO'])
            else:
                 # Por defecto ponemos 'NO' si está vacío en el original
                if "NO" in opciones_hielo: idx_hielo = opciones_hielo.index("NO")
            val_hielo = st.selectbox("Hielo (Col C)", opciones_hielo, index=idx_hielo)
            
        with c_actuacion:
            idx_act = 0
            if pd.notna(fila_datos['ACTUACION']) and fila_datos['ACTUACION'] in opciones_actuacion:
                idx_act = opciones_actuacion.index(fila_datos['ACTUACION'])
            else:
                if "NINGUNA" in opciones_actuacion: idx_act = opciones_actuacion.index("NINGUNA")
            val_actuacion = st.selectbox("Actuación (Col G)", opciones_actuacion, index=idx_act)

        # Botón de añadir
        submit_btn = st.form_submit_button("➕ Añadir a la lista", type="primary")
        
        if submit_btn:
            # Crear nueva fila
            nueva_fila = {
                'layer': fila_datos['layer'],
                'RUTA': fila_datos['RUTA'],
                'TRAMO': fila_datos['TRAMO'],
                'ESTADO': val_estado,
                'HIELO': val_hielo,
                'ACTUACION': val_actuacion,
                'RECORRIDO': fila_datos['RECORRIDO']
            }
            # Añadir a session_state usando pd.concat
            st.session_state['seleccionados'] = pd.concat([st.session_state['seleccionados'], pd.DataFrame([nueva_fila])], ignore_index=True)
            st.success("¡Tramo añadido correctamente!")

# --- 3. TABLA DE RESULTADOS ACUMULADA ---
st.subheader("📋 Rutas Añadidas")

if not st.session_state['seleccionados'].empty:
    st.dataframe(
        st.session_state['seleccionados'], 
        use_container_width=True,
        hide_index=True
    )
    
    # Botón para borrar la lista si te equivocas
    if st.button("🗑️ Borrar todo"):
        st.session_state['seleccionados'] = pd.DataFrame(columns=['layer', 'RUTA', 'TRAMO', 'ESTADO', 'HIELO', 'ACTUACION', 'RECORRIDO'])
        st.rerun()
else:
    st.info("Aún no has añadido ninguna ruta.")
