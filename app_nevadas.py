import streamlit as st
import pandas as pd

# Configuración de la página
st.set_page_config(page_title="Buscador de Rutas", layout="wide")

st.title("📊 Visor de Rutas y Tarifas")

# Función para cargar los datos
@st.cache_data
def cargar_datos():
    try:
        # Leemos el archivo ODS
        # engine='odf' es necesario para archivos .ods
        df = pd.read_excel("hoja calculo rutas.ods", engine="odf")
        return df
    except FileNotFoundError:
        st.error("No se encuentra el archivo 'hoja calculo rutas.ods'. Asegúrate de que está en la misma carpeta.")
        return None
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        return None

df = cargar_datos()

if df is not None:
    # --- MAPEO DE COLUMNAS ---
    # Python empieza a contar en 0. 
    # A=0, B=1, C=2, G=6, J=9, K=10, N=13
    try:
        col_A = df.columns[0]  # Input principal
        col_B = df.columns[1]  # Desplegable 1
        col_C = df.columns[2]  # Desplegable 2
        col_G = df.columns[6]  # Desplegable 3
        
        col_J = df.columns[9]  # Info relacionada
        col_K = df.columns[10] # Info relacionada
        col_N = df.columns[13] # Info resultado (Output)
        
        cols_visualizar = [col_A, col_J, col_K, col_N]
        
    except IndexError:
        st.error("El archivo no tiene suficientes columnas. Verifica el formato.")
        st.stop()

    # --- BARRA LATERAL (FILTROS) ---
    st.sidebar.header("🔍 Filtros")

    # Función auxiliar para crear listas únicas sin repetir y ordenadas
    def obtener_opciones(columna):
        opciones = df[columna].dropna().astype(str).unique().tolist()
        opciones.sort()
        return ["Todos"] + opciones

    # Desplegables para B, C y G (sin repetir valores)
    filtro_B = st.sidebar.selectbox(f"Filtrar por {col_B} (Col B)", obtener_opciones(col_B))
    filtro_C = st.sidebar.selectbox(f"Filtrar por {col_C} (Col C)", obtener_opciones(col_C))
    filtro_G = st.sidebar.selectbox(f"Filtrar por {col_G} (Col G)", obtener_opciones(col_G))

    # --- ÁREA PRINCIPAL ---

    # Input para la Columna A
    st.markdown(f"### Buscar por {col_A}")
    busqueda_A = st.text_input(f"Escribe el valor de {col_A} (ej. contenido de celda A2)", "")

    # --- LÓGICA DE FILTRADO ---
    df_filtrado = df.copy()

    # Aplicar filtros de desplegables
    if filtro_B != "Todos":
        df_filtrado = df_filtrado[df_filtrado[col_B].astype(str) == filtro_B]
    
    if filtro_C != "Todos":
        df_filtrado = df_filtrado[df_filtrado[col_C].astype(str) == filtro_C]

    if filtro_G != "Todos":
        df_filtrado = df_filtrado[df_filtrado[col_G].astype(str) == filtro_G]

    # Aplicar filtro de texto (Columna A)
    if busqueda_A:
        # Filtramos si contiene el texto (insensible a mayúsculas/minúsculas)
        df_filtrado = df_filtrado[df_filtrado[col_A].astype(str).str.contains(busqueda_A, case=False, na=False)]

    # --- MOSTRAR RESULTADOS ---
    st.divider()
    
    # 1. Mostrar la tabla con las columnas fijas solicitadas (A, J, K, N)
    st.subheader("Resultados")
    st.dataframe(df_filtrado[cols_visualizar], use_container_width=True, hide_index=True)

    # 2. Destacar el resultado de la columna N si hay una coincidencia única
    # Esto cumple con "si pongo la celda A2 aparezca la info correspondiente en N2"
    if len(df_filtrado) == 1:
        valor_N = df_filtrado.iloc[0][col_N]
        st.success(f"📍 Resultado exacto en {col_N}: **{valor_N}**")
    elif len(df_filtrado) == 0:
        st.warning("No se encontraron resultados con esos criterios.")
