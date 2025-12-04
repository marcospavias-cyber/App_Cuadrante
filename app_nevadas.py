import streamlit as st
import pandas as pd
import numpy as np

# Configuración de la página
st.set_page_config(page_title="Buscador de Rutas", layout="wide")

st.title("📊 Visor de Rutas y Tarifas")

# --- DATOS INTEGRADOS (No se requiere subir archivo) ---
# Estos datos provienen de tu archivo 'hoja calculo rutas.xlsx - rutas.csv'
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

@st.cache_data
def cargar_datos():
    # Creamos el DataFrame directamente desde la variable
    df = pd.DataFrame(DATOS_RUTAS)
    return df

df = cargar_datos()

if df is not None:
    # --- DEFINICIÓN DE COLUMNAS ---
    # Usamos los nombres reales del CSV en lugar de índices (A, B, C...)
    col_A = 'layer'       # Input búsqueda
    col_B = 'ESTADO'      # Filtro 1
    col_C = 'HIELO'       # Filtro 2
    col_G = 'ACTUACION'   # Filtro 3
    
    col_J = 'RUTA'        # Info Relacionada
    col_K = 'TRAMO'       # Info Relacionada
    col_N = 'RECORRIDO'   # Resultado (Output)
    
    # Columnas fijas a visualizar
    cols_visualizar = [col_A, col_J, col_K, col_N]

    # --- BARRA LATERAL (FILTROS) ---
    st.sidebar.header("🔍 Filtros")

    def obtener_opciones(columna):
        # Filtramos valores nulos (None/NaN) y convertimos a string
        opciones = df[columna].dropna().astype(str).unique().tolist()
        opciones.sort()
        return ["Todos"] + opciones

    filtro_B = st.sidebar.selectbox(f"Filtrar por {col_B}", obtener_opciones(col_B))
    filtro_C = st.sidebar.selectbox(f"Filtrar por {col_C}", obtener_opciones(col_C))
    filtro_G = st.sidebar.selectbox(f"Filtrar por {col_G}", obtener_opciones(col_G))

    # --- ÁREA PRINCIPAL ---

    st.markdown(f"### Buscar por {col_A}")
    busqueda_A = st.text_input(f"Escribe el valor de {col_A} (ej. R1-2)", "")

    # --- LÓGICA DE FILTRADO ---
    df_filtrado = df.copy()

    # Aplicar filtros de desplegables
    if filtro_B != "Todos":
        df_filtrado = df_filtrado[df_filtrado[col_B].astype(str) == filtro_B]
    
    if filtro_C != "Todos":
        df_filtrado = df_filtrado[df_filtrado[col_C].astype(str) == filtro_C]

    if filtro_G != "Todos":
        df_filtrado = df_filtrado[df_filtrado[col_G].astype(str) == filtro_G]

    # Aplicar filtro de texto (Columna A - layer)
    if busqueda_A:
        df_filtrado = df_filtrado[df_filtrado[col_A].astype(str).str.contains(busqueda_A, case=False, na=False)]

    # --- MOSTRAR RESULTADOS ---
    st.divider()
    st.subheader("Resultados")
    
    if not df_filtrado.empty:
        st.dataframe(df_filtrado[cols_visualizar], use_container_width=True, hide_index=True)
    else:
        st.info("No hay datos para mostrar con los filtros actuales.")

    # Destacar resultado único (Input A -> Output N)
    if len(df_filtrado) == 1:
        valor_N = df_filtrado.iloc[0][col_N]
        st.success(f"📍 Resultado exacto en {col_N}: **{valor_N}**")
    elif len(df_filtrado) == 0:
        st.warning("No se encontraron resultados.")
