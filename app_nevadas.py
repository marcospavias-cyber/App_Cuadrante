import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime, time, date

# Configuración de la página
st.set_page_config(page_title="Gestor de Rutas", layout="wide")

st.title("📝 Gestor de Partes de Ruta")

# --- 1. BASE DE DATOS COMPLETA (77 RUTAS) ---
# Ordenada naturalmente: R1, R2, R3... R9, R11...
DATOS_RUTAS = [
    {'layer': 'R1-1', 'ESTADO': 'CERRADA', 'HIELO': None, 'DISTANCIA': 3.21, 'PASADAS': 2.0, 'RECURSO': None, 'ACTUACION': 'SALERO', 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '1º', 'TOTAL_KM': 0.0, 'PERSONAL': 'vol morella', 'RECORRIDO': 'CV108 Ballestar a CV107 La Pobla de Benifassar', 'H.INICIO': 1000.0, 'H.FIN': 1300.0, 'TIEMPO': 3.0},
    {'layer': 'R1-2', 'ESTADO': 'CADENAS', 'HIELO': 'HIELO', 'DISTANCIA': 15.73, 'PASADAS': 3.0, 'RECURSO': 'BRP531', 'ACTUACION': 'CUCHILLA', 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '2º', 'TOTAL_KM': 0.0, 'PERSONAL': 'ubf morella', 'RECORRIDO': 'CV105 cruce CV107 a La Senia', 'H.INICIO': 1130.0, 'H.FIN': 1300.0, 'TIEMPO': 130.0},
    {'layer': 'R1-3', 'ESTADO': 'ABIERTA', 'HIELO': None, 'DISTANCIA': 11.08, 'PASADAS': 4.0, 'RECURSO': None, 'ACTUACION': 'CUÑA', 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '3º', 'TOTAL_KM': 0.0, 'PERSONAL': None, 'RECORRIDO': 'CV106 CV105 a Fredes', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R1-4', 'ESTADO': 'CADENAS', 'HIELO': None, 'DISTANCIA': 5.08, 'PASADAS': 5.0, 'RECURSO': None, 'ACTUACION': 'CUÑA SALERO', 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '4º', 'TOTAL_KM': 0.0, 'PERSONAL': None, 'RECORRIDO': 'Camino rural Fredes a Boixar', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R1-5', 'ESTADO': 'CADENAS', 'HIELO': 'PLACAS', 'DISTANCIA': 19.27, 'PASADAS': 3.0, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '5º', 'TOTAL_KM': 0.0, 'PERSONAL': None, 'RECORRIDO': 'CV109 cruce CV105 Boixar a CV105 cruce CV110 Herbes.', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R1-6', 'ESTADO': 'CERRADA', 'HIELO': None, 'DISTANCIA': 7.57, 'PASADAS': 2.0, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '6º', 'TOTAL_KM': 0.0, 'PERSONAL': None, 'RECORRIDO': 'CV109 cruce CV105 Boixar a Coratxar', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R1-7', 'ESTADO': 'HIELO', 'HIELO': None, 'DISTANCIA': 8.43, 'PASADAS': 4.0, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '7º', 'TOTAL_KM': 0.0, 'PERSONAL': None, 'RECORRIDO': 'CV105 cruce CV109 Boixar a CV107', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R1-8', 'ESTADO': 'PLACAS', 'HIELO': None, 'DISTANCIA': 7.3, 'PASADAS': 50.0, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': '2025-09-26', 'RUTA': 1, 'TRAMO': '8º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'La Senia a Cases del Riu a Rosell por CV100', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R2-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 16.09, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 2, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV111 de N232 a Vallibona', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R2-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 6.13, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 2, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Pista rural dels Llivis entre 232 y CV12', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R2-3', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 11.11, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 2, 'TRAMO': '3º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Pista rural La Llacua acceso desde CV12', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R2-4', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 11.4, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 2, 'TRAMO': '4º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Pista rural de la Cana desde CV12 a CV124', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R2-5', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 4.48, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 2, 'TRAMO': '5º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Pista rural La Vega del Moll de CV125 a CV12', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R3-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 2.97, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 3, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Accesos a Morella son CV1160 y CV1170', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R3-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 12.46, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 3, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV14 Morella a Forcall', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R3-3', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 4.53, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 3, 'TRAMO': '3º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Forcall a Todolella por CV120 y CV122', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R3-4', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 14.25, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 3, 'TRAMO': '4º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Todolella a Olocau del Rey por CV120 y CV121', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R3-5', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 21.72, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 3, 'TRAMO': '5º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Olocau del Rey a Tronchon por CV123 y A226', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R3-6', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 10.28, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 3, 'TRAMO': '6º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Olocau del Rey a Bordon por CV121 y TE', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R3-7', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 11.18, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 3, 'TRAMO': '7º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Olocau del Rey a Todolella por CV 122', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R4-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 4.83, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 4, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'San Mateo a Xert por CV132', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R4-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 8.97, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 4, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Rosell a Bel por CV104', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R5-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 7.43, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 5, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Accesos a Morella CV1160 y CV1170', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R5-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 11.67, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 5, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV125 de CV14 a Cinctorres', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R5-3', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 15.21, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 5, 'TRAMO': '3º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Morella a Forcal de CV14 con CV125 a CV124 Forcall', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R5-4', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 15.46, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 5, 'TRAMO': '4º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV14 desde CV124 por Villores CV119 Ortells CV1171 y Palanques CV118', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R5-5', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 4.42, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 5, 'TRAMO': '5º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Camino Rural Vega del Moll La Mina de CV125 a CV12', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R5-6', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 9.02, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 5, 'TRAMO': '6º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Pista Rural Sierra de Palos', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R5-7', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 2.81, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 5, 'TRAMO': '7º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Pista de la Carcellera', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R7-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 4.54, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 7, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Accesos Vilafranca', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R7-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 11.62, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 7, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV15 de Vilafranca a Ares por acceso CV1260', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R7-3', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 7.39, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 7, 'TRAMO': '3º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV124 de CV12 a Castellfort', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R7-4', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 9.7, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 7, 'TRAMO': '4º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV126 Castellfort a Villafranca', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R7-5', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 14.44, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 7, 'TRAMO': '5º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Pista rural Vilafranca a portell de Morella', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R8-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 3.84, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 8, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV166 de Benasal a CV15', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R8-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 28.06, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 8, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV166 de Benassal a Sant Pau', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R8-3', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 10.61, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 8, 'TRAMO': '3º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV166 de CV15 a Vilar de Canes y CV168 Vilardecanes a CV15', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R8-4', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 6.46, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 8, 'TRAMO': '4º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV15 de CV165 a CV166', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R8-5', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 38.07, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 8, 'TRAMO': '5º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': "CV128 CV15 a Cati a L' Avella CV1270 y N232", 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R8-6', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 17.31, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 8, 'TRAMO': '6º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV167 Benasal hasta final cruce CV15', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R9-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 4.25, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 9, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV170 desde CV1720 a Vistabella poblacion', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R9-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 7.61, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 9, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Pista rural Vistabella a Sant Joan de Penyagolosa', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R9-3', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 5.88, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 9, 'TRAMO': '3º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV 1720 de cruce CV170 a Sierra del Boi', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R11-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 2.14, 'PASADAS': 3.0, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 11, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Accesos Castellfort', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R11-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 6.5, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 11, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV124 de CV126 a CV12', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R11-3', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 9.61, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 11, 'TRAMO': '3º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV126 Castellfort a Vilafranca', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R11-4', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 16.98, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 11, 'TRAMO': '4º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV15 Vilafranca a Llosar a Pista rural Vilafranca a Portell de Morella', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R11-5', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 10.91, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 11, 'TRAMO': '5º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Pista rual Portell de Morella a la Cuba cruce CV120', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R11-6', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 5.66, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 11, 'TRAMO': '6º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV120 desde La Cuba a cruce con Todolella', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R11-7', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 20.53, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 11, 'TRAMO': '7º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Pista rural La Mata Todolella desde CV120 a Cinctorres', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R11-8', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 14.78, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 11, 'TRAMO': '8º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV124 de Cinctorres a Castellfort', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R12-1', 'ESTADO': 'CADENAS', 'HIELO': None, 'DISTANCIA': 10.45, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 12, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV175 Villahermosa a CV190', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R12-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 5.28, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 12, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV176 de CV175 a CV190', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R12-3', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 12.2, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 12, 'TRAMO': '3º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV175 de Villahermosa a Puertomingalvo', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R15-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 4.84, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 15, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV207 Ballacas a CV209 Pina de Montalgrao', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R15-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 5.84, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 15, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV207 de CV209 Pina de Montalgrao a Villanueva de Viver', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R15-3', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 15.91, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 15, 'TRAMO': '3º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV207 Villanueva de Vives a CV20 por Fuente la Reina', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R15-4', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 13.46, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 15, 'TRAMO': '4º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV208 de CV207 a CV20 por Los Pastores', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R15-5', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 6.09, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 15, 'TRAMO': '5º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV2093 de Pina de Montalgrao a CV2092', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R16-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 19.28, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 16, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV200 de Segorbe a Aín', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R16-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 32.93, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 16, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV223 Ain a Alcudia a CV215 a Algimia a CV213 Matet a P.Villamalur', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R17-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 29.15, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 17, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV203 de Caudiel a CV205', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R17-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 5.7, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 17, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV 205 cruce CV203 a cruce CV202', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R17-3', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 8.44, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 17, 'TRAMO': '3º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV202 de cruce CV205 a Villamalur', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R17-4', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 7.46, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 17, 'TRAMO': '4º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV205 cruce CV202 a Cruce CV201 y CV201 cruce CV202 a Artesa', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R23-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 7.84, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 23, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV117 de Morella a Xiva de Morella', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R23-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 18.6, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 23, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV105 CV110 de N232 a lim provincia Teruel Herbes', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R23-3', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 10.52, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 23, 'TRAMO': '3º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV105 de cruce CV110 a Castell de Cabres', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R23-4', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 0.6, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 23, 'TRAMO': '4º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV1050 Acceso a Herbeset', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R26-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 22.45, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 26, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV193 de Lucena del Cid a Argelita', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R26-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 8.66, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 26, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV194 de Argelita a CV1970 Giraba de Abajo', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R26-3', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 1.09, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 26, 'TRAMO': '3º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV1970 de CV194 a Giraba', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R26-4', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 41.92, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 26, 'TRAMO': '4º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV194 de Giraba a lim provincia en Cortes de Arenoso', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R26-5', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 3.94, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 26, 'TRAMO': '5º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'Camino Mas de la Llosa de CV190 a limite provincia', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R27-1', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 15.09, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 27, 'TRAMO': '1º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV171 de Adzeneta a Xodos', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R27-2', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 16.17, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 27, 'TRAMO': '2º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV170 y CV169 de Adzeneta a Collado CV170 por Benafigos', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None},
    {'layer': 'R27-3', 'ESTADO': None, 'HIELO': None, 'DISTANCIA': 0.37, 'PASADAS': None, 'RECURSO': None, 'ACTUACION': None, 'FUNDENTE': None, 'FECHA': None, 'RUTA': 27, 'TRAMO': '3º', 'TOTAL_KM': None, 'PERSONAL': None, 'RECORRIDO': 'CV172 Acceso Benafigos', 'H.INICIO': None, 'H.FIN': None, 'TIEMPO': None}
]

@st.cache_data
def obtener_dataframes():
    return pd.DataFrame(DATOS_RUTAS)

df = obtener_dataframes()

COLUMNAS_FINALES = [
    'layer', 'RUTA', 'TRAMO', 'RECORRIDO', 
    'ESTADO', 'HIELO', 'ACTUACION',
    'DISTANCIA', 'PASADAS', 'TOTAL_KM',
    'RECURSO', 'PERSONAL', 'FUNDENTE',
    'FECHA', 'H.INICIO', 'H.FIN', 'TIEMPO'
]

if 'seleccionados' not in st.session_state:
    st.session_state['seleccionados'] = pd.DataFrame(columns=COLUMNAS_FINALES)

# --- FUNCIONES AUXILIARES ---

def parsear_fecha(valor):
    if pd.isna(valor) or valor == "None":
        return datetime.now().date()
    try:
        return datetime.strptime(str(valor), '%Y-%m-%d').date()
    except:
        return datetime.now().date()

def parsear_hora(valor):
    if pd.isna(valor) or valor == "None":
        return time(9, 0)
    try:
        val_int = int(float(valor))
        hora = val_int // 100
        minuto = val_int % 100
        if hora > 23: hora = 0
        if minuto > 59: minuto = 0
        return time(hora, minuto)
    except:
        return time(9, 0)

def safe_float(val):
    try: return float(val)
    except: return 0.0

def safe_str(val):
    if pd.isna(val) or val == "None": return ""
    return str(val)

# --- ÁREA DE EDICIÓN ---

opciones_estado = [x for x in df['ESTADO'].unique() if pd.notna(x)] + ["SIN INCIDENCIAS"]
opciones_hielo = [x for x in df['HIELO'].unique() if pd.notna(x)] + ["NO"]
opciones_actuacion = [x for x in df['ACTUACION'].unique() if pd.notna(x)] + ["NINGUNA"]

with st.container(border=True):
    st.subheader("1. Seleccionar Tramo")
    
    # Creamos label combo
    # IMPORTANTE: Convertimos a string por si hay algún dato vacío que rompa el concat
    df['label_combo'] = df['layer'].astype(str) + " | " + df['RECORRIDO'].astype(str)
    
    seleccion_usuario = st.selectbox("Elige la Ruta/Capa:", df['label_combo'].tolist())
    
    # Recuperamos fila original
    fila_datos = df[df['label_combo'] == seleccion_usuario].iloc[0]
    
    st.divider()
    
    with st.form("form_ruta_pro"):
        st.subheader("2. Editar Datos del Parte")
        
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
        st.markdown("##### 📏 Kilómetros (Cálculo Automático)")
        ca1, ca2, ca3, ca4 = st.columns(4)
        
        val_distancia = ca1.number_input("Distancia (km)", value=safe_float(fila_datos.get('DISTANCIA')), step=0.1, format="%.2f")
        val_pasadas = ca2.number_input("Pasadas", value=safe_float(fila_datos.get('PASADAS')), step=1.0, format="%.0f")
        
        # Campo bloqueado, solo visual
        ca3.text_input("Total Km", value="Auto-calc al guardar", disabled=True)
        
        val_recurso = ca4.text_input("Recurso", value=safe_str(fila_datos.get('RECURSO')))
        
        st.markdown("##### 🕒 Tiempo (Cálculo Automático)")
        cb1, cb2, cb3, cb4 = st.columns(4)
        
        val_fecha = cb1.date_input("Fecha", value=parsear_fecha(fila_datos.get('FECHA')))
        val_hinicio = cb2.time_input("Hora Inicio", value=parsear_hora(fila_datos.get('H.INICIO')))
        val_hfin = cb3.time_input("Hora Fin", value=parsear_hora(fila_datos.get('H.FIN')))
        
        # Campo bloqueado, solo visual
        cb4.text_input("Tiempo (min)", value="Auto-calc al guardar", disabled=True)
        
        st.markdown("##### 👷 Personal y Material")
        cc1, cc2 = st.columns(2)
        val_personal = cc1.text_input("Personal", value=safe_str(fila_datos.get('PERSONAL')))
        val_fundente = cc2.text_input("Fundente", value=safe_str(fila_datos.get('FUNDENTE')))

        st.markdown("<br>", unsafe_allow_html=True)
        submit_btn = st.form_submit_button("➕ Calcular y Añadir", type="primary")
        
        if submit_btn:
            # 1. CÁLCULO KM
            calculo_total_km = val_distancia * val_pasadas
            
            # 2. CÁLCULO TIEMPO
            dt_inicio = datetime.combine(datetime.today(), val_hinicio)
            dt_fin = datetime.combine(datetime.today(), val_hfin)
            
            diferencia = dt_fin - dt_inicio
            minutos_totales = diferencia.total_seconds() / 60
            
            # Corrección si pasa de medianoche (ej. 23:00 a 01:00)
            if minutos_totales < 0:
                minutos_totales += 1440
            
            # Formatear strings para CSV
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
                'TOTAL_KM': round(calculo_total_km, 2),
                'RECURSO': val_recurso,
                'PERSONAL': val_personal,
                'FUNDENTE': val_fundente,
                'FECHA': str_fecha,
                'H.INICIO': str_hinicio,
                'H.FIN': str_hfin,
                'TIEMPO': round(minutos_totales)
            }
            
            st.session_state['seleccionados'] = pd.concat([st.session_state['seleccionados'], pd.DataFrame([nueva_fila])], ignore_index=True)
            st.success(f"✅ Guardado: {fila_datos['layer']} | {minutos_totales:.0f} min | {calculo_total_km:.2f} km")

# --- TABLA Y DESCARGA ---
st.subheader("📋 Tabla de Partes")

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
    st.info("Utiliza el formulario de arriba para añadir tramos.")
