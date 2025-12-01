import streamlit as st
import pandas as pd
import os

# Configuración de la página
st.set_page_config(page_title="Gestión de Rutas", layout="wide")

def cargar_datos():
    nombre_archivo = "prueba.xlsx - Hoja1.csv"
    if os.path.exists(nombre_archivo):
        df = pd.read_csv(nombre_archivo)
        # Limpiamos los espacios en blanco de los nombres de las columnas (ej: "Ruta " -> "Ruta")
        df.columns = df.columns.str.strip()
        return df, nombre_archivo
    else:
        st.error(f"No se encontró el archivo: {nombre_archivo}")
        return pd.DataFrame(), nombre_archivo

def main():
    st.title("Aplicación de Control de Rutas")
    
    # 1. Cargar la base de la tabla
    df, nombre_archivo = cargar_datos()
    
    if df.empty and 'Ruta' not in df.columns:
        st.warning("El archivo CSV parece estar vacío o no tiene el formato correcto.")
        # Creamos la estructura si el archivo está vacío basándonos en tu petición
        columnas = ['Ruta', 'Cadenas', 'Hielo', 'Tramos', 'Pasadas', 'Distancias']
        df = pd.DataFrame(columns=columnas)

    # Sección de entrada de datos (Los desplegables)
    st.subheader("Añadir / Editar Información")
    
    with st.form("formulario_rutas"):
        col1, col2, col3 = st.columns(3)
        col4, col5, col6 = st.columns(3)
        
        # Diccionario para guardar los valores del formulario
        nuevo_registro = {}
        
        # Columna 1: Ruta (Texto libre, ya que suele ser el nombre)
        with col1:
            nuevo_registro["Ruta"] = st.text_input("Ruta (Nombre)")
            
        # Columna 2: Cadenas (Si/No)
        with col2:
            nuevo_registro["Cadenas"] = st.selectbox("Cadenas", ["Si", "No"])
            
        # Columna 3: Hielo (Si/No)
        with col3:
            nuevo_registro["Hielo"] = st.selectbox("Hielo", ["Si", "No"])
            
        # Columna 4: Tramos (Si/No)
        with col4:
            nuevo_registro["Tramos"] = st.selectbox("Tramos", ["Si", "No"])
            
        # Columna 5: Pasadas (Número Natural)
        with col5:
            # step=1 asegura números enteros, min_value=0 asegura naturales
            nuevo_registro["Pasadas"] = st.number_input("Pasadas", min_value=0, step=1, format="%d")
            
        # Columna 6: Distancias (Si/No - Según tu instrucción)
        with col6:
            nuevo_registro["Distancias"] = st.selectbox("Distancias", ["Si", "No"])
            
        # Botón para enviar
        enviado = st.form_submit_button("Añadir a la Tabla")
        
        if enviado:
            if nuevo_registro["Ruta"]: # Verificamos que se haya puesto un nombre
                # Convertimos el diccionario a un DataFrame de una fila
                nueva_fila = pd.DataFrame([nuevo_registro])
                
                # Concatenamos con el DataFrame existente
                df = pd.concat([df, nueva_fila], ignore_index=True)
                
                # Guardamos en el CSV
                df.to_csv(nombre_archivo, index=False)
                st.success("¡Registro añadido correctamente!")
                st.rerun() # Recargar la página para ver los cambios
            else:
                st.error("Por favor, escribe un nombre para la Ruta.")

    # Mostrar la tabla actualizada
    st.subheader("Base de Datos Actual")
    st.dataframe(df, use_container_width=True)

if __name__ == "__main__":
    main()