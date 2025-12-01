import streamlit as st
import pandas as pd
import os

# Configuración de la página
st.set_page_config(page_title="Gestión de Rutas", layout="wide")

def cargar_datos():
    nombre_archivo = "prueba.xlsx - Hoja1.csv"
    # Definimos las columnas incluyendo la nueva "Zona"
    columnas_ordenadas = ['Zona', 'Ruta', 'Cadenas', 'Hielo', 'Tramos', 'Pasadas', 'Distancias']
    
    if os.path.exists(nombre_archivo):
        df = pd.read_csv(nombre_archivo)
        df.columns = df.columns.str.strip()
        
        # Añadimos columnas que falten (por si el archivo es antiguo)
        for col in columnas_ordenadas:
            if col not in df.columns:
                df[col] = None
                
        # Reordenamos las columnas para que se vean bien
        df = df[columnas_ordenadas]
        return df, nombre_archivo
    else:
        return pd.DataFrame(columns=columnas_ordenadas), nombre_archivo

def ordenar_por_ruta(df):
    """Ordena el DataFrame por el número de la Ruta."""
    if df.empty or 'Ruta' not in df.columns:
        return df
    
    # Extraemos el número para ordenar correctamente (1, 2, ... 10)
    df['_temp_sort'] = df['Ruta'].astype(str).str.extract(r'(\d+)').astype(float)
    df = df.sort_values(by=['_temp_sort'], na_position='last')
    df = df.drop(columns=['_temp_sort'])
    return df

def main():
    st.title("Aplicación de Control de Rutas")
    
    # Cargar datos
    df, nombre_archivo = cargar_datos()

    # --- FORMULARIO DE ENTRADA ---
    st.subheader("Añadir / Editar Información")
    
    with st.form("formulario_rutas"):
        # Organizamos en columnas (Ahora 7 campos, usamos 4 arriba y 3 abajo o similar)
        # Fila 1
        c1, c2, c3, c4 = st.columns(4)
        # Fila 2
        c5, c6, c7 = st.columns(3)
        
        nuevo_registro = {}
        
        # 1. ZONA (Controla las rutas)
        with c1:
            zona_seleccionada = st.selectbox("Zona", ["Lucena", "Villahermosa"])
            nuevo_registro["Zona"] = zona_seleccionada
        
        # 2. RUTA (Filtrada según la Zona)
        with c2:
            if zona_seleccionada == "Lucena":
                # De la 1 a la 5
                rutas_posibles = [f"Ruta {i}" for i in range(1, 6)]
            else: # Villahermosa
                # De la 5 a la 10
                rutas_posibles = [f"Ruta {i}" for i in range(5, 11)]
            
            nuevo_registro["Ruta"] = st.selectbox("Ruta", rutas_posibles)

        # 3. CADENAS
        with c3:
            nuevo_registro["Cadenas"] = st.selectbox("Cadenas", ["Si", "No"])
            
        # 4. HIELO
        with c4:
            nuevo_registro["Hielo"] = st.selectbox("Hielo", ["Si", "No"])
            
        # 5. TRAMOS
        with c5:
            nuevo_registro["Tramos"] = st.selectbox("Tramos", ["Si", "No"])
            
        # 6. PASADAS
        with c6:
            nuevo_registro["Pasadas"] = st.number_input("Pasadas", min_value=0, step=1, format="%d")
            
        # 7. DISTANCIAS
        with c7:
            nuevo_registro["Distancias"] = st.selectbox("Distancias", ["Si", "No"])
            
        # Botón de enviar
        enviado = st.form_submit_button("Guardar en Tabla")
        
        if enviado:
            nueva_fila = pd.DataFrame([nuevo_registro])
            
            # Lógica de actualización: Si la ruta ya existe, la reemplazamos
            if not df.empty and nuevo_registro["Ruta"] in df["Ruta"].values:
                st.info(f"Actualizando datos de la {nuevo_registro['Ruta']}...")
                df = df[df["Ruta"] != nuevo_registro["Ruta"]]
                df = pd.concat([df, nueva_fila], ignore_index=True)
            else:
                df = pd.concat([df, nueva_fila], ignore_index=True)
            
            # Ordenar y Guardar
            df = ordenar_por_ruta(df)
            df.to_csv(nombre_archivo, index=False)
            st.success("¡Registro guardado correctamente!")
            st.rerun()

    # --- VISUALIZACIÓN DE LA TABLA ---
    st.subheader("Estado Actual de las Rutas")
    
    # Mostramos la tabla (asegurando el orden visual)
    df_visual = ordenar_por_ruta(df)
    st.dataframe(df_visual, use_container_width=True, hide_index=True)

if __name__ == "__main__":
    main()
