import streamlit as st
import pandas as pd
import os

# Configuración de la página
st.set_page_config(page_title="Gestión de Rutas", layout="wide")

def cargar_datos():
    nombre_archivo = "prueba.xlsx - Hoja1.csv"
    if os.path.exists(nombre_archivo):
        df = pd.read_csv(nombre_archivo)
        # Limpiamos los espacios en blanco de los nombres de las columnas
        df.columns = df.columns.str.strip()
        # Aseguramos que las columnas existan si el archivo estaba vacío
        columnas_requeridas = ['Ruta', 'Cadenas', 'Hielo', 'Tramos', 'Pasadas', 'Distancias']
        for col in columnas_requeridas:
            if col not in df.columns:
                df[col] = None
        return df, nombre_archivo
    else:
        # Si no existe, creamos la estructura vacía
        columnas = ['Ruta', 'Cadenas', 'Hielo', 'Tramos', 'Pasadas', 'Distancias']
        return pd.DataFrame(columns=columnas), nombre_archivo

def ordenar_por_ruta(df):
    """
    Ordena el DataFrame basándose en el número contenido en la columna 'Ruta'.
    Esto evita que 'Ruta 10' aparezca antes que 'Ruta 2'.
    """
    if df.empty or 'Ruta' not in df.columns:
        return df
    
    # Creamos una columna temporal extrayendo solo el número de la Ruta
    # r'(\d+)' busca cualquier secuencia de dígitos
    df['_temp_sort'] = df['Ruta'].astype(str).str.extract(r'(\d+)').astype(float)
    
    # Ordenamos por ese número (y ponemos los que no tengan número al final)
    df = df.sort_values(by='_temp_sort', na_position='last')
    
    # Eliminamos la columna temporal para que no se vea
    df = df.drop(columns=['_temp_sort'])
    return df

def main():
    st.title("Aplicación de Control de Rutas")
    
    # 1. Cargar la base de la tabla
    df, nombre_archivo = cargar_datos()

    # Sección de entrada de datos
    st.subheader("Añadir / Editar Información")
    
    with st.form("formulario_rutas"):
        col1, col2, col3 = st.columns(3)
        col4, col5, col6 = st.columns(3)
        
        nuevo_registro = {}
        
        # Columna 1: Ruta (AHORA ES UN DESPLEGABLE DE 1 a 10)
        with col1:
            opciones_ruta = [f"Ruta {i}" for i in range(1, 11)]
            nuevo_registro["Ruta"] = st.selectbox("Ruta", opciones_ruta)
            
        # Columna 2: Cadenas
        with col2:
            nuevo_registro["Cadenas"] = st.selectbox("Cadenas", ["Si", "No"])
            
        # Columna 3: Hielo
        with col3:
            nuevo_registro["Hielo"] = st.selectbox("Hielo", ["Si", "No"])
            
        # Columna 4: Tramos
        with col4:
            nuevo_registro["Tramos"] = st.selectbox("Tramos", ["Si", "No"])
            
        # Columna 5: Pasadas (Número Natural)
        with col5:
            nuevo_registro["Pasadas"] = st.number_input("Pasadas", min_value=0, step=1, format="%d")
            
        # Columna 6: Distancias
        with col6:
            nuevo_registro["Distancias"] = st.selectbox("Distancias", ["Si", "No"])
            
        # Botón para enviar
        enviado = st.form_submit_button("Añadir a la Tabla")
        
        if enviado:
            # Crear nueva fila
            nueva_fila = pd.DataFrame([nuevo_registro])
            
            # Verificar si esa Ruta ya existe para actualizarla o añadirla
            # (Opcional: si quieres que se dupliquen quita este if/else y deja solo el concat)
            if not df.empty and nuevo_registro["Ruta"] in df["Ruta"].values:
                st.info(f"Actualizando información existente para {nuevo_registro['Ruta']}")
                # Borramos la antigua y ponemos la nueva (o podrías solo actualizar)
                df = df[df["Ruta"] != nuevo_registro["Ruta"]]
                df = pd.concat([df, nueva_fila], ignore_index=True)
            else:
                df = pd.concat([df, nueva_fila], ignore_index=True)
            
            # Ordenamos la tabla antes de guardar
            df = ordenar_por_ruta(df)
            
            # Guardamos en CSV
            df.to_csv(nombre_archivo, index=False)
            st.success("¡Registro guardado y tabla ordenada!")
            st.rerun()

    # Mostrar la tabla actualizada y ordenada
    st.subheader("Listado de Rutas (Ordenado)")
    
    # Aseguramos el orden visualmente también por si acaso se carga sin guardar cambios
    df_visual = ordenar_por_ruta(df)
    st.dataframe(df_visual, use_container_width=True, hide_index=True)

if __name__ == "__main__":
    main()
