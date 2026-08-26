# app_conciliacion_v40_hibrida_corregida_final.py
#
# FIX 1: Bloqueo de 4 dígitos eliminado (Permite cruzar Refs cortas).
# FIX 2: Textos gerenciales ultra-cortos. 
# FIX 3: Nueva Agrupación IP (Mezcla y suma) por Fecha y Sector como Plan B.
# FIX 4: Diferencia de Fecha de 1 día pasa a Morado. > 1 día pasa a Salmón.

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from difflib import SequenceMatcher
from datetime import datetime

st.set_page_config(page_title="Conciliación Integral CLM", layout="wide")
st.markdown('''
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stApp {max-width: 100%;}
    </style>
''', unsafe_allow_html=True)

st.title("🏦 Conciliación Automatizada — Motor Integral v40 (Híbrida) 🤖")
st.write("Sube tu archivo consolidado.")
st.caption(
    "Selecciona opcion 'Tarde' solo despues de depurar los pendientes del primer cruce."
)

with st.expander("⚙️ Parámetros de tolerancia"):
    TOPE_DIAS_ALERTA = st.slider(
        "Días máximos permitidos para cruzar con desfase de fecha (1 día = Morado, 2+ días = Salmón).",
        1, 10, 4
    )
    tol_valor_purpura = st.number_input(
        "Diferencia máxima de valor ($) para alerta MORADA",
        min_value=1, value=500, step=50, max_value=500
    )
    tol_valor_abs_general = st.number_input(
        "Diferencia absoluta general de valor para alertar (más allá del morado) ($)",
        min_value=1, value=5000, step=100
    )
    tol_valor_pct_general = st.number_input(
        "Diferencia relativa general de valor para alertar (%)",
        min_value=0.00, value=0.00, step=0.05
    ) / 100
    multiplo_redondo = st.selectbox("Múltiplo para valor 'redondo' (alta ambigüedad)", [50000, 100000], index=1)
    
    umbral_fuzzy_nequi = st.slider(
        "Sensibilidad de coincidencia difusa para 'NEQUI' en legalizaciones DZ (0.60 = más flexible, 0.95 = más estricto)",
        0.60, 0.95, 0.72, step=0.01
    )

    st.divider()
    modo_tarde = st.checkbox(
        "🌅 Activar Casilla 'Tarde' (Segunda pasada en Pendientes)",
        value=False,
        help="Ejecuta una segunda pasada profunda (T1 a T6) y genera una pestaña ordenada de menor a mayor."
    )

COLOR_AZUL = "#C5D9F1"      # Conciliado perfecto
COLOR_VERDE = "#A9D18E"     # DZ multiposición sin conciliar
COLOR_SALMON = "#F5B7A1"    # Diferencia de fecha > 1 día
COLOR_MORADO = "#C39BD3"    # Diferencia de valor O fecha de exactamente 1 día
COLOR_DURAZNO = "#FAD7A0"   # Reclasificación de banco
COLOR_BLANCO = "#FFFFFF"    # Pendiente / Sugerencias estándar
COLOR_GRIS = "#D0CECE"      # Cruces IP/CB (Regla 3 estricta o mezcla por fecha/sector)
COLOR_AMARILLO = "#FFF2CC"  # Sugerencias del Modo Tarde

archivo_subido = st.file_uploader("Selecciona el archivo de Excel o CSV", type=['xlsx', 'csv'])

if archivo_subido is not None:
    if st.button("🚀 Ejecutar Conciliación", use_container_width=True):
        try:
            with st.spinner("Ejecutando motor híbrido v40... Esto puede tomar unos segundos."):

                # =====================================================
                # 1. LECTURA
                # =====================================================
                if archivo_subido.name.lower().endswith('.csv'):
                    df = pd.read_csv(archivo_subido)
                else:
                    hojas = pd.read_excel(archivo_subido, sheet_name=None)
                    hojas_validas = [h for h in hojas.values() if not h.dropna(how='all').empty]
                    if not hojas_validas:
                        st.error("El archivo no contiene hojas con datos.")
                        st.stop()
                    df = pd.concat(hojas_validas, ignore_index=True)

                df.columns = df.columns.str.strip()

                col_A = 'Asignación' if 'Asignación' in df.columns else ('Asignacion' if 'Asignacion' in df.columns else None)
                col_B = 'Nº documento' if 'Nº documento' in df.columns else 'Nº doc.'
                col_C = 'Clase de documento' if 'Clase de documento' in df.columns else ('Clase doc.' if 'Clase doc.' in df.columns else None)
                col_D = 'Fe.contabilización' if 'Fe.contabilización' in df.columns else ('Fecha de documento' if 'Fecha de documento' in df.columns else None)
                col_F = 'Fecha valor' if 'Fecha valor' in df.columns else 'Fe-valor'
                col_G = 'Clave contabiliz.' if 'Clave contabiliz.' in df.columns else 'CT'
                col_H = 'Referencia'
                col_I = 'Importe en moneda local' if 'Importe en moneda local' in df.columns else 'Importe en ML'
                col_K = 'Texto' if 'Texto' in df.columns else None
                col_novedad = 'Doc.compensación' if 'Doc.compensación' in df.columns else ('Novedad' if 'Novedad' in df.columns else None)
                col_banco = 'Clave referencia 3'

                if not col_A:
                    st.error("No se encontró la columna Asignación (A).")
                    st.stop()

                requeridas = [col_H, col_G, col_F, col_I, col_banco, col_B]
                faltantes = [c for c in requeridas if c not in df.columns]
                if faltantes:
                    st.error(f"No se encontraron estas columnas obligatorias: {faltantes}")
                    st.stop()

                usar_ipcb = col_C is not None
                columnas_originales = list(df.columns)

                # =====================================================
                # 2. AUTOCOMPLETADO DE BANCO
                # =====================================================
                mapeo_cuentas_banco = {
                    "1110056001": "CUENTA 1110056001", "1110056101": "BANCO DE BOGOTA",
                    "1110056201": "BANCO DAVIBANK S.A.", "1110056301": "BANCOLOMBIA S.A.",
                    "1110056401": "BANCO CAJA SOCIAL S.", "1110056501": "BANCO DAVIVIENDA S.A",
                    "1110056601": "BANCO BILBAO VIZCAYA", "1110056701": "BANCO AGRARIO DE COL",
                    "1120055001": "BANCO COMERCIAL AV V", "1120055101": "BANCO DE OCCIDENTE",
                    "1120055301": "BANCO GNB SUDAMERIS",
                }
                bancos_completados = []
                current_bank = None
                for _, row in df.iterrows():
                    asig_val = str(row.get(col_A, ""))
                    banco_val = row.get(col_banco, None)
                    if "cuenta de mayor" in asig_val.lower():
                        m = re.search(r'(\d{6,})', asig_val)
                        if m:
                            current_bank = mapeo_cuentas_banco.get(m.group(1), f"CUENTA {m.group(1)} (sin mapear)")
                    if pd.notnull(banco_val) and str(banco_val).strip().lower() not in ("", "nan"):
                        current_bank = str(banco_val).strip()
                    bancos_completados.append(current_bank)
                df[col_banco] = bancos_completados
                df = df[~df[col_A].astype(str).str.contains("cuenta de mayor", case=False, na=False)].copy()

                # =====================================================
                # 3. LIMPIEZA BASE
                # =====================================================
                df[col_B] = pd.to_numeric(df[col_B], errors='coerce')
                filas_antes = len(df)
                filas_descartadas = df[df[col_B].isna() | df[col_G].isna()].copy()
                df = df.dropna(subset=[col_B, col_G]).reset_index(drop=True)
                filas_excluidas = filas_antes - len(df)

                df = df.sort_values(by=[col_B], ascending=True).reset_index(drop=True)
                df['ID_Linea'] = df.index

                df[col_G] = df[col_G].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                df[col_banco] = df[col_banco].astype(str).str.strip()
                df[col_I] = pd.to_numeric(df[col_I], errors='coerce').fillna(0)
                df['Abs_I'] = df[col_I].abs()

                df['Fecha_F'] = pd.to_datetime(df[col_F], errors='coerce', dayfirst=True)
                df[col_F] = df['Fecha_F'].dt.date

                if col_D:
                    df['Fecha_D'] = pd.to_datetime(df[col_D], errors='coerce', dayfirst=True)
                    df['Periodo_D'] = df['Fecha_D'].dt.to_period('M').astype(str)
                    df.loc[df['Fecha_D'].isna(), 'Periodo_D'] = 'SIN_FECHA_D'
                else:
                    df['Periodo_D'] = 'SIN_FECHA_D'

                df['Estado_Conciliacion'] = 'Pendiente'
                df['Comentario'] = ''
                df['Candidatos_Conciliacion'] = ''

                df['B_Repite'] = df.groupby(col_B)[col_B].transform('count') > 1

                # =====================================================
                # 4. SECTORIZACIÓN
                # =====================================================
                mapeo_referencias_dist = {
                    "11760923": "Dist Acopi", "11761277": "Dist Acopi", "11761293": "Dist Acopi",
                    "11761327": "Dist Acopi", "11761301": "Dist Acopi", "12273934": "Dist Acopi",
                    "11761319": "Dist Acopi", "12273900": "Dist Acopi", "12273926": "Dist Acopi",
                    "14632012": "Dist Acopi", "15186547": "Dist Acopi", "13048756": "Dist Acopi",
                    "15186539": "Dist Acopi", "16219602": "Dist Acopi", "16591240": "Dist Acopi",
                    "16634586": "Dist Acopi", "14885164": "Dist Acopi", "19827765": "Dist Acopi",
                    "11761350": "Dist Buga", "12161154": "Dist Buga", "14294946": "Dist Buga",
                    "15926645": "Dist Buga", "17608589": "Dist Buga",
                    "11831583": "Dist Dosquebradas", "12161162": "Dist Dosquebradas",
                    "12161121": "Dist Dosquebradas", "12161139": "Dist Dosquebradas",
                    "12874475": "Dist Dosquebradas", "15190309": "Dist Dosquebradas",
                    "14468144": "Dist Dosquebradas", "12500773": "Dist Dosquebradas",
                    "14468151": "Dist Dosquebradas", "14651459": "Dist Dosquebradas",
                    "15444946": "Dist Dosquebradas", "16062176": "Dist Dosquebradas",
                    "20836698": "Dist Dosquebradas", "72806854": "Dist Dosquebradas",
                    "20719829": "Dist Dosquebradas",
                    "15536188": "Dist Pasto", "12637294": "Dist Pasto", "11844685": "Dist Pasto",
                    "20235651": "Dist Pasto", "15536170": "Dist Pasto", "17549197": "Dist Pasto",
                    "17608605": "Dist Pasto", "17968405": "VENTA EN LINEA"
                }
                
                pares_ip = [
                    ("3001", "11760923"), ("3002", "11761277"), ("3003", "11761293"),
                    ("3004", "11761327"), ("3005", "11761301"), ("3006", "12273934"),
                    ("3007", "11761319"), ("3008", "12273900"), ("3009", "12273926"),
                    ("3010", "14632012"), ("3011", "15186547"), ("3012", "13048756"),
                    ("3013", "15186539"), ("3200", "16219602"), ("3201", "16591240"),
                    ("3202", "16634586"), ("2005", "14885164"), ("3203", "19827765"),
                    ("2001", "11761350"), ("2002", "12161154"), ("2003", "14294946"),
                    ("2210", "15926645"), ("4002", "11831583"), ("4001", "12161162"),
                    ("4003", "12161121"), ("4004", "12161139"), ("4005", "12874475"),
                    ("4006", "15190309"), ("4008", "12500773"), ("4009", "14468151"),
                    ("4010", "14651459"), ("4200", "15444946"), ("4253", "16062176"),
                    ("4007", "20836698"), ("4202", "20836698"), ("4203", "72806854"),
                    ("4201", "20719829"), ("6101", "15536188"), ("6102", "12637294"),
                    ("6103", "11844685"), ("6106", "15536170"), ("6108", "17549197")
                ]
                dict_8_to_list4 = {}
                for r4, r8 in pares_ip:
                    if r8 not in dict_8_to_list4:
                        dict_8_to_list4[r8] = []
                    dict_8_to_list4[r8].append(r4)

                def clasificar_sector(row):
                    texto_k = str(row.get(col_K, "")) if col_K else ""
                    texto_nov = str(row.get(col_novedad, "")) if col_novedad else ""
                    h_val = str(row.get(col_H, "")).strip()
                    h_val = re.sub(r'\.0$', '', h_val)

                    t = f"{texto_k} {texto_nov}".upper()
                    if 'D502' in t: return 'Dist Buga'
                    if 'D503' in t: return 'Dist Acopi'
                    if 'D504' in t: return 'Dist Dosquebradas'
                    if 'D505' in t: return 'Dist Pasto'

                    if h_val.isdigit():
                        num = int(h_val)
                        if 2000 <= num <= 2999: return 'Dist Buga'
                        if 3000 <= num <= 3999: return 'Dist Acopi'
                        if 4000 <= num <= 4999: return 'Dist Dosquebradas'
                        if 6000 <= num <= 6999: return 'Dist Pasto'

                    if h_val in mapeo_referencias_dist: return mapeo_referencias_dist[h_val]

                    return 'Sin clasificar'

                df['Sector'] = df.apply(clasificar_sector, axis=1)

                def obtener_ref_homologada(row):
                    texto = f" {row.get(col_H,'')} {row.get(col_A,'')} {row.get(col_K,'') if col_K else ''} {row.get(col_novedad,'') if col_novedad else ''} ".upper()
                    n8 = re.findall(r' \d{8} ', texto)
                    for n in n8:
                        if n in dict_8_to_list4:
                            return n
                    n4 = re.findall(r' \d{4} ', texto)
                    for n in n4:
                        for k8, list_4 in dict_8_to_list4.items():
                            if n in list_4:
                                return k8
                    return None

                # =====================================================
                # DETECCIÓN NEQUI
                # =====================================================
                def _similitud(palabra, objetivo='NEQUI'):
                    return SequenceMatcher(None, palabra, objetivo).ratio()

                def contiene_nequi_fuzzy(texto, umbral):
                    if not texto: return False
                    texto_up = str(texto).upper()
                    if 'NEQUI' in texto_up: return True
                    palabras = re.findall(r'[A-ZÑ]{3,8}', texto_up)
                    for p in palabras:
                        if _similitud(p, 'NEQUI') >= umbral: return True
                    return False

                def es_nequi_flexible(row, umbral):
                    g_val = str(row.get(col_G, '')).strip()
                    c_val = str(row.get(col_C, '')).strip().upper() if usar_ipcb else ''

                    if g_val == '50':
                        h_raw = str(row.get(col_H, '')).strip()
                        h_clean = re.sub(r'\.0$', '', h_raw)
                        if h_clean.isdigit():
                            h_num = int(h_clean)
                            if 100_000 <= h_num <= 999_999_999: return True
                            if 1_000_000_000 <= h_num <= 1_399_999_999: return True
                        return False

                    if g_val == '40':
                        val_a = str(row.get(col_A, '')).strip().upper()
                        if val_a == 'T' or val_a.startswith('T-') or val_a.startswith('T/') or val_a == '/':
                            return True
                            
                        if c_val == 'DZ':
                            texto_completo = f"{row.get(col_K, '') if col_K else ''} {row.get(col_A, '')} {row.get(col_H, '')}"
                            return contiene_nequi_fuzzy(texto_completo, umbral)

                    return False

                def limpiar_numero(v):
                    if pd.isna(v): return ''
                    t = re.sub(r'\.0$', '', str(v).strip())
                    m = re.findall(r'\d+', t)
                    return m[0] if m else ''

                def solo_digitos(valor):
                    return re.sub(r'\D', '', str(valor or ''))

                def referencias_se_contienen(valor_a, valor_h):
                    a = solo_digitos(valor_a)
                    h = solo_digitos(valor_h)
                    return bool(a and h and (a in h or h in a))

                df['Es_Nequi'] = df.apply(lambda r: es_nequi_flexible(r, umbral_fuzzy_nequi), axis=1)
                df['H_Limpia'] = df[col_H].apply(limpiar_numero)
                df['A_Limpia'] = df[col_A].apply(limpiar_numero)

                if usar_ipcb:
                    df['Es_IP_G40'] = (df[col_C].astype(str).str.upper() == 'IP') & (df[col_G] == '40')
                    df['Es_CB_G50'] = False
                else:
                    df['Es_IP_G40'] = False
                    df['Es_CB_G50'] = False

                # =====================================================
                # FUNCIONES AUXILIARES
                # =====================================================
                usados = set()
                parejas_registradas = []

                def formato_linea(idl):
                    r = df.loc[df['ID_Linea'] == idl].iloc[0]
                    d = str(int(r[col_B])) if pd.notna(r[col_B]) else ""
                    c = str(r[col_C]) if usar_ipcb and pd.notna(r.get(col_C)) else ""
                    g = str(r[col_G])
                    i = f"${r['Abs_I']:,.0f}"
                    f = str(r[col_F])
                    etiqueta_clase = f"{c}=" if c and c.lower() != 'nan' else ""
                    return f"{d} ({etiqueta_clase}G{g}, {i}, F{f})"

                def resumen_docs(sub_df):
                    return ", ".join(str(int(d)) for d in sub_df[col_B].tolist())

                def escribir_estado(indices, estado, forzar=False):
                    if not indices: return
                    if forzar:
                        df.loc[df['ID_Linea'].isin(indices), 'Estado_Conciliacion'] = estado
                    else:
                        mask = df['ID_Linea'].isin(indices) & (df['Estado_Conciliacion'] == 'Pendiente')
                        df.loc[mask, 'Estado_Conciliacion'] = estado

                def escribir_comentario(idl, texto, append=True):
                    mask = df['ID_Linea'] == idl
                    if append:
                        anterior = str(df.loc[mask, 'Comentario'].iloc[0])
                        df.loc[mask, 'Comentario'] = texto if not anterior else anterior + " | " + texto
                    else:
                        df.loc[mask, 'Comentario'] = texto

                def escribir_candidatos(idl, texto):
                    df.loc[df['ID_Linea'] == idl, 'Candidatos_Conciliacion'] = texto

                def diferencia_dias_fila(id40, id50):
                    f40 = df.loc[df['ID_Linea'] == id40, 'Fecha_F'].iloc[0]
                    f50 = df.loc[df['ID_Linea'] == id50, 'Fecha_F'].iloc[0]
                    if pd.isna(f40) or pd.isna(f50): return None
                    return abs((f40 - f50).days)

                def fecha_dentro_de_4_dias(id40, id50):
                    dias = diferencia_dias_fila(id40, id50)
                    return dias is not None and dias <= TOPE_DIAS_ALERTA

                def registrar_pareja_por_fecha(id40, id50, ignorar=None):
                    dias = diferencia_dias_fila(id40, id50)
                    if dias is None or dias > TOPE_DIAS_ALERTA:
                        comentario = f'Excede límite de días permitidos ({dias} días).'
                        for idx in [id40, id50]: escribir_comentario(idx, comentario, append=False)
                        return False

                    texto = f'{formato_linea(id40)} | {formato_linea(id50)}'
                    
                    if dias == 0:
                        estado = 'Conciliado'
                        msg = 'Cruce exacto.'
                    elif dias <= 1:
                        estado = 'Diferencia de fecha'
                        msg = f'Diferencia de fecha: {dias} día(s).'
                    else:
                        estado = 'Diferencia de fecha extensa'
                        msg = f'Diferencia de fecha extensa: {dias} día(s).'

                    for idx in [id40, id50]:
                        df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = estado
                        df.loc[df['ID_Linea'] == idx, 'Comentario'] = msg
                        df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto
                    usados.update([id40, id50])
                    parejas_registradas.append((id40, id50))
                    return True

                def gate_seguridad(id40, id50, exigir_importe_exacto=True, tolerancia_valor=None, ignorar_sector=False):
                    res = df.loc[df['ID_Linea'].isin([id40, id50])]
                    ra = res[res['ID_Linea'] == id40].iloc[0]
                    rb = res[res['ID_Linea'] == id50].iloc[0]

                    resultado = {
                        'ok': False, 'motivo': '', 'dif_dias': None,
                        'mismo_banco': True, 'banco_a': '', 'banco_b': '',
                        'dif_valor': None, 'pct_valor': None,
                        'mismo_sector': True, 'es_nequi': False, 'es_ip': False
                    }

                    clase_a = str(ra.get(col_C, '')).strip().upper() if usar_ipcb else ''
                    clase_b = str(rb.get(col_C, '')).strip().upper() if usar_ipcb else ''
                    es_ip = (clase_a == 'IP') or (clase_b == 'IP')
                    resultado['es_ip'] = es_ip

                    fa, fb = ra['Fecha_F'], rb['Fecha_F']
                    if pd.isna(fa) or pd.isna(fb):
                        resultado['motivo'] = "Fecha F inválida"
                        return resultado

                    dif_dias = abs((fa - fb).days)
                    resultado['dif_dias'] = dif_dias

                    if dif_dias > TOPE_DIAS_ALERTA:
                        resultado['motivo'] = "Excede límite de días permitidos"
                        return resultado

                    def limpiar_nombre_banco(nombre_banco):
                        nombre = str(nombre_banco).upper()
                        if 'BANCOLOMBIA' in nombre: return 'Bancolombia'
                        if 'DAVIVIENDA' in nombre: return 'Davivienda'
                        if 'DAVIBANK' in nombre: return 'Davivienda'
                        if 'BOGOTA' in nombre: return 'Banco de Bogotá'
                        if 'CAJA SOCIAL' in nombre: return 'Banco Caja Social'
                        if 'BILBAO' in nombre or 'BBVA' in nombre: return 'BBVA'
                        if 'AGRARIO' in nombre: return 'Banco Agrario'
                        if 'AV V' in nombre or 'VILLAS' in nombre: return 'Banco AV Villas'
                        if 'OCCIDENTE' in nombre: return 'Banco de Occidente'
                        if 'SUDAMERIS' in nombre: return 'Banco GNB Sudameris'
                        return nombre.title()

                    banco_a_crudo = str(ra[col_banco]).strip()
                    banco_b_crudo = str(rb[col_banco]).strip()
                    
                    resultado['banco_a'] = limpiar_nombre_banco(banco_a_crudo)
                    resultado['banco_b'] = limpiar_nombre_banco(banco_b_crudo)

                    if es_ip:
                        resultado['mismo_banco'] = True
                    else:
                        resultado['mismo_banco'] = (banco_a_crudo == banco_b_crudo)

                    sector_a = str(ra.get('Sector', '')).strip()
                    sector_b = str(rb.get('Sector', '')).strip()
                    
                    if not ignorar_sector:
                        if sector_a not in ('', 'Sin clasificar') and sector_b not in ('', 'Sin clasificar'):
                            if sector_a != sector_b:
                                resultado['motivo'] = "Sector distinto"
                                resultado['mismo_sector'] = False
                                return resultado

                    imp_a = abs(ra[col_I])
                    imp_b = abs(rb[col_I])
                    dif_valor = round(abs(imp_a - imp_b), 2)
                    max_imp = max(imp_a, imp_b, 1)
                    pct_valor = dif_valor / max_imp
                    resultado['dif_valor'] = dif_valor
                    resultado['pct_valor'] = pct_valor

                    if exigir_importe_exacto and dif_valor > 0:
                        tope = tolerancia_valor if tolerancia_valor is not None else 0
                        if dif_valor > tope:
                            resultado['motivo'] = f"Importe distinto"
                            return resultado

                    resultado['es_nequi'] = bool(ra.get('Es_Nequi', False)) or bool(rb.get('Es_Nequi', False))
                    resultado['ok'] = True
                    resultado['motivo'] = "OK"
                    return resultado

                def clasificar_y_registrar(id40, id50, ignorar_sector=False):
                    res = gate_seguridad(id40, id50, exigir_importe_exacto=True, tolerancia_valor=tol_valor_purpura, ignorar_sector=ignorar_sector)
                    if not res['ok']:
                        return False, res['motivo']

                    texto_candidatos = f"{formato_linea(id40)} | {formato_linea(id50)}"
                    comentario_40 = []
                    comentario_50 = []
                    
                    dias = res['dif_dias']
                    if res['es_ip']:
                        if dias > 1:
                            estado_final = 'Diferencia de fecha extensa'
                            msg = f"Diferencia de fecha extensa: {dias} día(s)."
                        elif dias == 1:
                            estado_final = 'Diferencia de fecha'
                            msg = f"Diferencia de fecha: {dias} día(s)."
                        elif res['dif_valor'] and res['dif_valor'] > 0:
                            estado_final = 'Diferencia de valor'
                            msg = f"Diferencia de valor: ${res['dif_valor']:,.0f}."
                        else:
                            estado_final = 'Cruce Múltiple IP/CB'
                            msg = "Cruce exacto (POS)."
                            
                        if msg:
                            comentario_40.append(msg)
                            comentario_50.append(msg)

                    elif not res['mismo_banco']:
                        estado_final = 'Reclasificacion de Banco'
                        comentario_40.append(f"Registrado en '{res['banco_a']}'; debe ser '{res['banco_b']}'.")
                        comentario_50.append(f"Registrado en '{res['banco_b']}'; debe ser '{res['banco_a']}'.")
                    elif dias > 1:
                        estado_final = 'Diferencia de fecha extensa'
                        msg = f"Diferencia de fecha extensa: {dias} día(s)."
                        comentario_40.append(msg)
                        comentario_50.append(msg)
                    elif dias == 1:
                        estado_final = 'Diferencia de fecha'
                        msg = f"Diferencia de fecha: {dias} día(s)."
                        comentario_40.append(msg)
                        comentario_50.append(msg)
                    elif res['dif_valor'] and res['dif_valor'] > 0:
                        estado_final = 'Diferencia de valor'
                        msg = f"Diferencia de valor: ${res['dif_valor']:,.0f}."
                        comentario_40.append(msg)
                        comentario_50.append(msg)
                    else:
                        estado_final = 'Conciliado'
                        comentario_40.append("Cruce exacto.")
                        comentario_50.append("Cruce exacto.")

                    if res['es_nequi']:
                        comentario_40.append("[Nequi]")
                        comentario_50.append("[Nequi]")

                    c40_str = " ".join(comentario_40)
                    c50_str = " ".join(comentario_50)

                    escribir_estado([id40, id50], estado_final, forzar=True)
                    escribir_candidatos(id40, texto_candidatos)
                    escribir_candidatos(id50, texto_candidatos)
                    escribir_comentario(id40, c40_str, append=False)
                    escribir_comentario(id50, c50_str, append=False)

                    parejas_registradas.append((id40, id50))
                    return True, estado_final

                # =====================================================
                # Regla 3a: IP Homologados Agrupado ESTRICTO BD
                # =====================================================
                ind_ip_exacto = set()
                ind_ip_tolerancia = set()
                if usar_ipcb:
                    df['Ref_H_Homologada'] = df.apply(obtener_ref_homologada, axis=1)

                    df_ip = df[(df[col_C].astype(str).str.upper() == 'IP') & (df[col_G] == '40') & df['Ref_H_Homologada'].notna()]
                    df_cb = df[(df[col_C].astype(str).str.upper() == 'CB') & (df[col_G] == '50') & df['Ref_H_Homologada'].notna()]

                    if not df_ip.empty and not df_cb.empty:
                        grp_ip = df_ip.groupby([col_banco, 'Sector', 'Ref_H_Homologada'])['Abs_I'].sum().reset_index(name='S_IP')
                        grp_cb = df_cb.groupby([col_banco, 'Sector', 'Ref_H_Homologada'])['Abs_I'].sum().reset_index(name='S_CB')
                        m = pd.merge(grp_cb, grp_ip, on=[col_banco, 'Sector', 'Ref_H_Homologada'])
                        m['DifV'] = (m['S_CB'] - m['S_IP']).abs()
                        max_s = m[['S_CB', 'S_IP']].max(axis=1).clip(lower=1)
                        m['Pct'] = m['DifV'] / max_s

                        exactos = m[m['DifV'].round(2) == 0]
                        con_tol = m[(m['DifV'] > 0) & ((m['DifV'] <= tol_valor_abs_general) | (m['Pct'] <= tol_valor_pct_general))]

                        def procesar_grupo_ip(fila, es_exacto):
                            b, s, rh = fila[col_banco], fila['Sector'], fila['Ref_H_Homologada']
                            sub_ip = df_ip[(df_ip[col_banco] == b) & (df_ip['Sector'] == s) & (df_ip['Ref_H_Homologada'] == rh)]
                            sub_cb = df_cb[(df_cb[col_banco] == b) & (df_cb['Sector'] == s) & (df_cb['Ref_H_Homologada'] == rh)]
                            ip_ids = [i for i in sub_ip['ID_Linea'].tolist() if i not in usados]
                            cb_ids = [i for i in sub_cb['ID_Linea'].tolist() if i not in usados]
                            if not ip_ids or not cb_ids: return
                            usados.update(ip_ids + cb_ids)
                            texto_cand = " | ".join(formato_linea(i) for i in ip_ids + cb_ids)
                            if es_exacto:
                                estado = 'Cruce Múltiple IP/CB'
                                ind_ip_exacto.update(ip_ids + cb_ids)
                                txt = "Cruce exacto homologado (POS)."
                            else:
                                estado = 'Diferencia de valor'
                                ind_ip_tolerancia.update(ip_ids + cb_ids)
                                txt = f"Diferencia de valor (${fila['DifV']:,.0f}) (POS)."
                            for idx in ip_ids + cb_ids:
                                df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = estado
                                df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                                df.loc[df['ID_Linea'] == idx, 'Comentario'] = txt

                        for _, fila in exactos.iterrows(): procesar_grupo_ip(fila, es_exacto=True)
                        for _, fila in con_tol.iterrows(): procesar_grupo_ip(fila, es_exacto=False)

                # ================================================================
                # Regla 3b: MEZCLA Y SUMA IP (Por Fecha y Sector) Plan B
                # ================================================================
                if usar_ipcb:
                    df_ip_pend = df[(df[col_C].astype(str).str.upper() == 'IP') & (df[col_G] == '40') & (~df['ID_Linea'].isin(usados))]
                    df_cb_pend = df[(df[col_C].astype(str).str.upper() == 'CB') & (df[col_G] == '50') & (~df['ID_Linea'].isin(usados))]

                    if not df_ip_pend.empty and not df_cb_pend.empty:
                        grp_ip2 = df_ip_pend.groupby([col_banco, 'Sector', 'Fecha_F'])['Abs_I'].sum().reset_index(name='S_IP')
                        grp_cb2 = df_cb_pend.groupby([col_banco, 'Sector', 'Fecha_F'])['Abs_I'].sum().reset_index(name='S_CB')
                        m2 = pd.merge(grp_cb2, grp_ip2, on=[col_banco, 'Sector', 'Fecha_F'])
                        m2['DifV'] = (m2['S_CB'] - m2['S_IP']).abs()
                        max_s2 = m2[['S_CB', 'S_IP']].max(axis=1).clip(lower=1)
                        m2['Pct'] = m2['DifV'] / max_s2

                        exactos2 = m2[m2['DifV'].round(2) == 0]
                        con_tol2 = m2[(m2['DifV'] > 0) & ((m2['DifV'] <= tol_valor_abs_general) | (m2['Pct'] <= tol_valor_pct_general))]

                        def procesar_grupo_ip_fallback(fila, es_exacto):
                            b, s, f = fila[col_banco], fila['Sector'], fila['Fecha_F']
                            sub_ip = df_ip_pend[(df_ip_pend[col_banco] == b) & (df_ip_pend['Sector'] == s) & (df_ip_pend['Fecha_F'] == f)]
                            sub_cb = df_cb_pend[(df_cb_pend[col_banco] == b) & (df_cb_pend['Sector'] == s) & (df_cb_pend['Fecha_F'] == f)]
                            
                            ip_ids = [i for i in sub_ip['ID_Linea'].tolist() if i not in usados]
                            cb_ids = [i for i in sub_cb['ID_Linea'].tolist() if i not in usados]
                            if not ip_ids or not cb_ids: return
                            
                            usados.update(ip_ids + cb_ids)
                            texto_cand = " | ".join(formato_linea(i) for i in ip_ids + cb_ids)
                            
                            if es_exacto:
                                estado = 'Sugerencia - Cruce IP por fecha y sector'
                                txt = "Sugerencia POS: Cuadre perfecto por misma Fecha y Sector."
                            else:
                                estado = 'Diferencia de valor'
                                txt = f"Sugerencia POS: Diferencia de valor (${fila['DifV']:,.0f}) por Fecha y Sector."
                                
                            for idx in ip_ids + cb_ids:
                                df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = estado
                                df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                                df.loc[df['ID_Linea'] == idx, 'Comentario'] = txt

                        for _, fila in exactos2.iterrows(): procesar_grupo_ip_fallback(fila, es_exacto=True)
                        for _, fila in con_tol2.iterrows(): procesar_grupo_ip_fallback(fila, es_exacto=False)

                    # Los IP que sobren aquí quedarán pendientes
                    ip_sin_resolver = df[(df[col_C].astype(str).str.upper() == 'IP') & (df[col_G] == '40') & (~df['ID_Linea'].isin(usados))]
                    for idl in ip_sin_resolver['ID_Linea']:
                        escribir_comentario(idl, "Falta Referencia POS.", append=False)

                # =====================================================
                # Regla 8: NEQUI POR TOTALES Y FIFO
                # =====================================================
                df_nequi_dz = df[(df[col_G] == '40') & (df['Es_Nequi'] == True) & (~df['ID_Linea'].isin(usados))]
                df_cb_disponible = df[(df[col_G] == '50') & (df['Es_Nequi'] == True) & (~df['ID_Linea'].isin(usados))]
                
                if usar_ipcb:
                    df_cb_disponible = df_cb_disponible[df_cb_disponible[col_C].astype(str).str.upper() != 'IP']

                if not df_nequi_dz.empty and not df_cb_disponible.empty:
                    for (banco_g, sector_g, fecha_g), grupo_dz in df_nequi_dz.groupby([col_banco, 'Sector', 'Fecha_F']):
                        grupo_dz = grupo_dz[~grupo_dz['ID_Linea'].isin(usados)]
                        if grupo_dz.empty: continue
                        
                        if sector_g in ('', 'Sin clasificar'):
                            grupo_cb = df_cb_disponible[(df_cb_disponible[col_banco] == banco_g) & (df_cb_disponible['Fecha_F'] == fecha_g) & (~df_cb_disponible['ID_Linea'].isin(usados))]
                        else:
                            grupo_cb = df_cb_disponible[(df_cb_disponible[col_banco] == banco_g) & (df_cb_disponible['Sector'].isin([sector_g, 'Sin clasificar'])) & (df_cb_disponible['Fecha_F'] == fecha_g) & (~df_cb_disponible['ID_Linea'].isin(usados))]
                        
                        if grupo_cb.empty: continue

                        n_dz, n_cb = len(grupo_dz), len(grupo_cb)
                        total_dz, total_cb = round(grupo_dz['Abs_I'].sum(), 2), round(grupo_cb['Abs_I'].sum(), 2)

                        dz_ord = grupo_dz.sort_values(col_B).reset_index(drop=True)
                        cb_ord = grupo_cb.sort_values(col_B).reset_index(drop=True)
                        n_parejas = min(n_dz, n_cb)

                        if n_dz == n_cb and total_dz == total_cb:
                            for i in range(n_parejas):
                                id40, id50 = dz_ord.iloc[i]['ID_Linea'], cb_ord.iloc[i]['ID_Linea']
                                texto_cand = f"{formato_linea(id40)} | {formato_linea(id50)}"
                                for idx in (id40, id50):
                                    df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Conciliado'
                                    df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                                    df.loc[df['ID_Linea'] == idx, 'Comentario'] = "Cruce Nequi por totales."
                                usados.update([id40, id50])
                                parejas_registradas.append((id40, id50))
                            usados.update(dz_ord['ID_Linea'].tolist() + cb_ord['ID_Linea'].tolist())
                        else:
                            docs_dz, docs_cb = resumen_docs(dz_ord), resumen_docs(cb_ord)
                            texto_cand = f"Grupo NEQUI: DZ: {docs_dz} || CB: {docs_cb}"
                            for _, fila in pd.concat([dz_ord, cb_ord]).iterrows():
                                idl = fila['ID_Linea']
                                escribir_estado([idl], 'Pendiente o solicitar soporte', forzar=False)
                                if df.loc[df['ID_Linea'] == idl, 'Candidatos_Conciliacion'].iloc[0] == '':
                                    escribir_candidatos(idl, texto_cand)
                                escribir_comentario(idl, "Totales Nequi no cuadran.", append=False)

                # =====================================================
                # Regla 1 — A debe coincidir con H (exacto)
                # =====================================================
                df_40 = df[(df[col_G] == '40') & (~df['Es_IP_G40'])].copy()
                df_50 = df[(df[col_G] == '50') & (~df['Es_CB_G50'])].copy()

                def emparejar_1a1_por_llave(sub40, sub50, llave40, llave50):
                    s40 = sub40[~sub40['ID_Linea'].isin(usados)].copy()
                    s50 = sub50[~sub50['ID_Linea'].isin(usados)].copy()
                    if s40.empty or s50.empty: return
                    s40['_pos'] = s40.groupby(llave40).cumcount()
                    s50['_pos'] = s50.groupby(llave50).cumcount()
                    merged = pd.merge(s40, s50, left_on=llave40 + ['_pos'], right_on=llave50 + ['_pos'], suffixes=('_40', '_50'))
                    for _, r in merged.iterrows():
                        id40, id50 = r['ID_Linea_40'], r['ID_Linea_50']
                        if id40 in usados or id50 in usados: continue
                        ok, _ = clasificar_y_registrar(id40, id50)
                        if ok: usados.update([id40, id50])

                emparejar_1a1_por_llave(df_40, df_50, [col_banco, 'Abs_I', col_A], [col_banco, 'Abs_I', col_H])

                df_40_limpia = df_40[(~df_40['ID_Linea'].isin(usados)) & (df_40['A_Limpia'] != '')].copy()
                df_50_limpia = df_50[(~df_50['ID_Linea'].isin(usados)) & (df_50['H_Limpia'] != '')].copy()
                emparejar_1a1_por_llave(df_40_limpia, df_50_limpia, [col_banco, 'Abs_I', 'A_Limpia'], [col_banco, 'Abs_I', 'H_Limpia'])

                # ================================================================
                # FLEX POR REFERENCIA PARCIAL
                # ================================================================
                pend40_flex = df[(df[col_G] == '40') & (~df['Es_IP_G40']) & (~df['ID_Linea'].isin(usados))]
                pend50_flex = df[(df[col_G] == '50') & (~df['Es_CB_G50']) & (~df['ID_Linea'].isin(usados))]

                for id40, fila40 in pend40_flex.iterrows():
                    candidatos = pend50_flex[
                        (pend50_flex[col_banco] == fila40[col_banco]) &
                        (pend50_flex['Abs_I'] == fila40['Abs_I']) &
                        (~pend50_flex['ID_Linea'].isin(usados))
                    ].copy()
                    candidatos = candidatos[candidatos[col_H].apply(lambda h: referencias_se_contienen(fila40[col_A], h))]

                    if len(candidatos) != 1: continue

                    id50 = candidatos.iloc[0]['ID_Linea']
                    texto = f'{formato_linea(id40)} | {formato_linea(id50)}'
                    for idx in [id40, id50]:
                        df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Conciliado'
                        df.loc[df['ID_Linea'] == idx, 'Comentario'] = "Cruce por referencia parcial."
                        df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto
                    usados.update([id40, id50])
                    parejas_registradas.append((id40, id50))

                # =====================================================
                # Regla 6 EXPLÍCITA — RECLASIFICACIÓN DE BANCO
                # =====================================================
                df_40 = df[(df[col_G] == '40') & (~df['Es_IP_G40']) & (~df['ID_Linea'].isin(usados))].copy()
                df_50 = df[(df[col_G] == '50') & (~df['Es_CB_G50']) & (~df['ID_Linea'].isin(usados))].copy()

                def emparejar_reclasificacion(sub40, sub50, llave40, llave50):
                    s40 = sub40[~sub40['ID_Linea'].isin(usados)].copy()
                    s50 = sub50[~sub50['ID_Linea'].isin(usados)].copy()
                    if s40.empty or s50.empty: return
                    s40['_pos'] = s40.groupby(llave40).cumcount()
                    s50['_pos'] = s50.groupby(llave50).cumcount()
                    merged = pd.merge(s40, s50, left_on=llave40 + ['_pos'], right_on=llave50 + ['_pos'], suffixes=('_40', '_50'))
                    for _, r in merged.iterrows():
                        id40, id50 = r['ID_Linea_40'], r['ID_Linea_50']
                        if id40 in usados or id50 in usados: continue
                        ra = df.loc[df['ID_Linea'] == id40].iloc[0]
                        rb = df.loc[df['ID_Linea'] == id50].iloc[0]
                        if str(ra[col_banco]).strip() == str(rb[col_banco]).strip(): continue
                        ok, _ = clasificar_y_registrar(id40, id50)
                        if ok: usados.update([id40, id50])

                emparejar_reclasificacion(df_40, df_50, ['Abs_I', col_A], ['Abs_I', col_H])
                
                df_40_r6_limpia = df_40[(~df_40['ID_Linea'].isin(usados)) & (df_40['A_Limpia'] != '')].copy()
                df_50_r6_limpia = df_50[(~df_50['ID_Linea'].isin(usados)) & (df_50['H_Limpia'] != '')].copy()
                emparejar_reclasificacion(df_40_r6_limpia, df_50_r6_limpia, ['Abs_I', 'A_Limpia'], ['Abs_I', 'H_Limpia'])

                # ================================================================
                # SECTORIZACION MULTIPLE FIFO
                # ================================================================
                for (banco_g, sector_g, importe_g), lado40 in df[
                    (df[col_G] == '40') & (~df['Es_IP_G40']) & (~df['ID_Linea'].isin(usados)) & (df['Sector'] != 'Sin clasificar')
                ].groupby([col_banco, 'Sector', 'Abs_I']):
                    lado50 = df[
                        (df[col_G] == '50') & (~df['Es_CB_G50']) & (df[col_banco] == banco_g) & (df['Sector'] == sector_g) & (df['Abs_I'] == importe_g) & (~df['ID_Linea'].isin(usados))
                    ].copy()
                    lado40 = lado40.sort_values(col_B)
                    lado50 = lado50.sort_values(col_B)
                    if lado50.empty: continue

                    for id40, id50 in zip(lado40['ID_Linea'], lado50['ID_Linea']):
                        if id40 in usados or id50 in usados: continue
                        registrar_pareja_por_fecha(id40, id50)

                # Sobrantes de sectorización desbalanceada
                df_40 = df[(df[col_G] == '40') & (~df['Es_IP_G40']) & (~df['ID_Linea'].isin(usados))]
                df_50 = df[(df[col_G] == '50') & (~df['Es_CB_G50']) & (~df['ID_Linea'].isin(usados))]
                d40_sect = df_40[df_40['Sector'] != 'Sin clasificar']
                d50_sect = df_50[df_50['Sector'] != 'Sin clasificar']
                if not d40_sect.empty and not d50_sect.empty:
                    for grp, sub40 in d40_sect.groupby([col_banco, 'Abs_I', 'Sector']):
                        b, imp, sector = grp
                        sub50 = d50_sect[(d50_sect[col_banco] == b) & (d50_sect['Abs_I'] == imp) & (d50_sect['Sector'] == sector)]
                        if sub50.empty: continue
                        docs50_txt = resumen_docs(sub50)
                        docs40_txt = resumen_docs(sub40)
                        for _, r in sub40.iterrows():
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Estado_Conciliacion'] = 'Pendiente o solicitar soporte'
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Candidatos_Conciliacion'] = f"{formato_linea(r['ID_Linea'])} | Candidatos posibles: {docs50_txt}"
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Comentario'] = "Descuadre por sector."
                        for _, r in sub50.iterrows():
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Estado_Conciliacion'] = 'Pendiente o solicitar soporte'
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Candidatos_Conciliacion'] = f"{formato_linea(r['ID_Linea'])} | Candidatos posibles: {docs40_txt}"
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Comentario'] = "Descuadre por sector."

                # ================================================================
                # REGLA 7B - DIFERENCIA DE VALOR (Alertas Sugeridas)
                # ================================================================
                pend40_7b = df[(df[col_G] == '40') & (~df['Es_IP_G40']) & (~df['ID_Linea'].isin(usados)) & (df['Estado_Conciliacion'] == 'Pendiente')].copy()
                pend50_7b = df[(df[col_G] == '50') & (~df['Es_CB_G50']) & (~df['ID_Linea'].isin(usados)) & (df['Estado_Conciliacion'] == 'Pendiente')].copy()

                for id40, fila40 in pend40_7b.iterrows():
                    posibles = pend50_7b[
                        (pend50_7b[col_banco] == fila40[col_banco]) &
                        (~pend50_7b['ID_Linea'].isin(usados))
                    ].copy()

                    if posibles.empty: continue
                    
                    if fila40['Sector'] not in ('', 'Sin clasificar'):
                        posibles = posibles[(posibles['Sector'] == fila40['Sector']) | (posibles['Sector'] == 'Sin clasificar')]

                    posibles['_dif_dias'] = (posibles['Fecha_F'] - fila40['Fecha_F']).dt.days.abs().fillna(999)
                    posibles = posibles[posibles['_dif_dias'] <= TOPE_DIAS_ALERTA]
                    if posibles.empty: continue
                    posibles['_dif_valor'] = (posibles['Abs_I'] - fila40['Abs_I']).abs()
                    posibles = posibles[posibles['_dif_valor'] > tol_valor_purpura]
                    if posibles.empty: continue

                    id50 = posibles.sort_values(['_dif_valor', '_dif_dias']).iloc[0]['ID_Linea']
                    diferencia = round(abs(fila40['Abs_I'] - df.loc[id50, 'Abs_I']), 2)
                    texto = f'{formato_linea(id40)} | {formato_linea(id50)}'
                    
                    for idx in [id40, id50]:
                        df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Diferencia de valor'
                        df.loc[df['ID_Linea'] == idx, 'Comentario'] = f'Diferencia de valor: ${diferencia:,.0f}.'
                        df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto
                    usados.update([id40, id50])

                # =====================================================
                # EXCEPCIÓN NEQUI - DOBLE PASADA
                # =====================================================
                df_40 = df[(df[col_G] == '40') & (~df['Es_IP_G40']) & (~df['ID_Linea'].isin(usados))].copy()
                df_50 = df[(df[col_G] == '50') & (~df['Es_CB_G50']) & (~df['ID_Linea'].isin(usados))].copy()

                df_nequi_40 = df_40[df_40['Es_Nequi'] == True]

                def procesar_candidato_nequi(id40, candidatos_50):
                    exactos = candidatos_50[candidatos_50['Abs_I'] == df.loc[df['ID_Linea'] == id40, 'Abs_I'].iloc[0]]
                    if len(exactos) == 1:
                        id50 = exactos.iloc[0]['ID_Linea']
                        ok, _ = clasificar_y_registrar(id40, id50, ignorar_sector=False)
                        if ok: usados.update([id40, id50])
                        return True
                    
                    if len(exactos) > 1:
                        # FIFO Desambiguado Nequi
                        r40 = df.loc[df['ID_Linea'] == id40].iloc[0]
                        importe_r = r40['Abs_I']
                        grupo_cb = candidatos_50[(candidatos_50['Abs_I'] == importe_r) & (~candidatos_50['ID_Linea'].isin(usados))]
                        
                        if not grupo_cb.empty:
                            banco_r = r40[col_banco]
                            fecha_r = r40['Fecha_F']
                            sector_r = r40['Sector']
                            if sector_r in ('', 'Sin clasificar'):
                                grupo_dz = df[(df[col_G] == '40') & (df['Es_Nequi'] == True) & (~df['ID_Linea'].isin(usados)) & (df[col_banco] == banco_r) & (df['Fecha_F'] == fecha_r) & (df['Abs_I'] == importe_r)]
                            else:
                                grupo_dz = df[(df[col_G] == '40') & (df['Es_Nequi'] == True) & (~df['ID_Linea'].isin(usados)) & (df[col_banco] == banco_r) & (df['Sector'].isin([sector_r, 'Sin clasificar'])) & (df['Fecha_F'] == fecha_r) & (df['Abs_I'] == importe_r)]

                            n_dz, n_cb = len(grupo_dz), len(grupo_cb)
                            if n_dz > 0 and n_cb > 0:
                                dz_ord = grupo_dz.sort_values(col_B).reset_index(drop=True)
                                cb_ord = grupo_cb.sort_values(col_B).reset_index(drop=True)
                                n_pares = min(n_dz, n_cb)
                                for i in range(n_pares):
                                    id_dz = dz_ord.iloc[i]['ID_Linea']
                                    id_cb = cb_ord.iloc[i]['ID_Linea']
                                    if id_dz in usados or id_cb in usados: continue
                                    ok, _ = clasificar_y_registrar(id_dz, id_cb, ignorar_sector=False)
                                    if ok: usados.update([id_dz, id_cb])
                                if id40 in usados: return True
                        
                        if id40 not in usados:
                            docs_txt = resumen_docs(exactos)
                            df.loc[df['ID_Linea'] == id40, 'Estado_Conciliacion'] = 'Pendiente o solicitar soporte'
                            df.loc[df['ID_Linea'] == id40, 'Candidatos_Conciliacion'] = f"{formato_linea(id40)} | Candidatos posibles: {docs_txt}"
                            df.loc[df['ID_Linea'] == id40, 'Comentario'] = "Ambigüedad Nequi."
                            return True

                    r40 = df.loc[df['ID_Linea'] == id40].iloc[0]
                    candidatos_50 = candidatos_50.copy()
                    candidatos_50['_dif_val'] = (candidatos_50['Abs_I'] - r40['Abs_I']).abs()
                    con_tol = candidatos_50[candidatos_50['_dif_val'] <= tol_valor_purpura].sort_values('_dif_val')
                    
                    if len(con_tol) == 1:
                        id50 = con_tol.iloc[0]['ID_Linea']
                        ok, _ = clasificar_y_registrar(id40, id50, ignorar_sector=False)
                        if ok: usados.update([id40, id50])
                        return True
                    if len(con_tol) > 1:
                        docs_txt = resumen_docs(con_tol)
                        df.loc[df['ID_Linea'] == id40, 'Estado_Conciliacion'] = 'Pendiente o solicitar soporte'
                        df.loc[df['ID_Linea'] == id40, 'Candidatos_Conciliacion'] = f"{formato_linea(id40)} | Candidatos posibles: {docs_txt}"
                        df.loc[df['ID_Linea'] == id40, 'Comentario'] = "Ambigüedad Nequi (tolerancia)."
                        return True
                        
                    return False

                for _, r40 in df_nequi_40.iterrows():
                    id40 = r40['ID_Linea']
                    if id40 in usados: continue
                    
                    sec = r40['Sector']
                    if sec in ('', 'Sin clasificar'):
                        candidatos_50 = df_50[
                            (df_50[col_banco] == r40[col_banco]) &
                            (df_50['Es_Nequi'] == True) &
                            (~df_50['ID_Linea'].isin(usados))
                        ].copy()
                    else:
                        candidatos_50 = df_50[
                            (df_50[col_banco] == r40[col_banco]) &
                            (df_50['Es_Nequi'] == True) &
                            (df_50['Sector'].isin([sec, 'Sin clasificar'])) &
                            (~df_50['ID_Linea'].isin(usados))
                        ].copy()
                        
                    if candidatos_50.empty: continue

                    candidatos_50['_dif_dias'] = (candidatos_50['Fecha_F'] - r40['Fecha_F']).dt.days.abs().fillna(999)
                    candidatos_50 = candidatos_50[candidatos_50['_dif_dias'] == 0]
                    if candidatos_50.empty: continue

                    procesar_candidato_nequi(id40, candidatos_50)

                for _, r40 in df_nequi_40.iterrows():
                    id40 = r40['ID_Linea']
                    if id40 in usados: continue
                    
                    sec = r40['Sector']
                    if sec in ('', 'Sin clasificar'):
                        candidatos_50 = df_50[
                            (df_50[col_banco] == r40[col_banco]) &
                            (df_50['Es_Nequi'] == True) &
                            (~df_50['ID_Linea'].isin(usados))
                        ].copy()
                    else:
                        candidatos_50 = df_50[
                            (df_50[col_banco] == r40[col_banco]) &
                            (df_50['Es_Nequi'] == True) &
                            (df_50['Sector'].isin([sec, 'Sin clasificar'])) &
                            (~df_50['ID_Linea'].isin(usados))
                        ].copy()
                        
                    if candidatos_50.empty: continue

                    candidatos_50['_dif_dias'] = (candidatos_50['Fecha_F'] - r40['Fecha_F']).dt.days.abs().fillna(999)
                    candidatos_50 = candidatos_50.sort_values('_dif_dias')
                    if candidatos_50.empty: continue

                    procesar_candidato_nequi(id40, candidatos_50)

                # =====================================================
                # REGLA 4 — DOCUMENTOS DZ CON POSICIONES MÚLTIPLES
                # =====================================================
                df_40 = df[(df[col_G] == '40') & (~df['Es_IP_G40']) & (~df['ID_Linea'].isin(usados))].copy()
                df_50 = df[(df[col_G] == '50') & (~df['Es_CB_G50']) & (~df['ID_Linea'].isin(usados))].copy()

                dz_repetidos = df_40[(df_40['B_Repite'] == True) & (df_40['Candidatos_Conciliacion'] == '')]
                for b_doc, grupo in dz_repetidos.groupby(col_B):
                    for _, linea in grupo.iterrows():
                        idl = linea['ID_Linea']
                        if idl in usados: continue
                        candidatos = df_50[(df_50[col_banco] == linea[col_banco]) & (~df_50['ID_Linea'].isin(usados))].copy()
                        if candidatos.empty: continue

                        candidatos['_dif_dias'] = (candidatos['Fecha_F'] - linea['Fecha_F']).dt.days.abs().fillna(999)
                        candidatos = candidatos[candidatos['_dif_dias'] <= TOPE_DIAS_ALERTA]
                        if candidatos.empty: continue

                        candidatos['_dif_valor'] = (candidatos['Abs_I'] - linea['Abs_I']).abs()
                        candidatos['_mismo_sector'] = (candidatos['Sector'] == linea['Sector']) & (linea['Sector'] != 'Sin clasificar')
                        candidatos['_mismo_A_H'] = candidatos[col_H].astype(str).str.strip() == str(linea[col_A]).strip()
                        candidatos_ordenados = candidatos.sort_values(by=['_mismo_A_H', '_mismo_sector', '_dif_valor', '_dif_dias'], ascending=[False, False, True, True])

                        lineas_candidatos = [f"{rank}. Doc={int(c[col_B])}, H={c[col_H]}, I=${c['Abs_I']:,.0f}, F={c[col_F]}, Sector={c['Sector']}, dif_valor=${c['_dif_valor']:,.0f}, dif_dias={int(c['_dif_dias'])}" for rank, (_, c) in enumerate(candidatos_ordenados.head(5).iterrows(), start=1)]
                        texto_cand_final = f"Línea 40: Doc={int(linea[col_B])}, I=${linea['Abs_I']:,.0f}, F={linea[col_F]} || Candidatos: " + " ; ".join(lineas_candidatos)
                        df.loc[df['ID_Linea'] == idl, 'Estado_Conciliacion'] = 'Pendiente o solicitar soporte'
                        df.loc[df['ID_Linea'] == idl, 'Candidatos_Conciliacion'] = texto_cand_final
                        df.loc[df['ID_Linea'] == idl, 'Comentario'] = "Múltiples posiciones sin cruzar. Ver candidatos."

                # =====================================================
                # ÚLTIMO RECURSO: FIFO CONTROLADO
                # =====================================================
                df_40 = df[(df[col_G] == '40') & (~df['Es_IP_G40']) & (~df['ID_Linea'].isin(usados))].copy()
                df_50 = df[(df[col_G] == '50') & (~df['Es_CB_G50']) & (~df['ID_Linea'].isin(usados))].copy()

                pendientes_40 = df[(df['ID_Linea'].isin(df_40['ID_Linea'])) & (df['Estado_Conciliacion'] == 'Pendiente') & (df['Comentario'] == '') & (df['Candidatos_Conciliacion'] == '')]
                pendientes_50 = df[(df['ID_Linea'].isin(df_50['ID_Linea'])) & (df['Estado_Conciliacion'] == 'Pendiente') & (df['Comentario'] == '') & (df['Candidatos_Conciliacion'] == '')]

                ind_fifo_ok = set()
                ind_fifo_verde_dz = set()

                for grp, sub40 in pendientes_40.groupby([col_banco, 'Abs_I', col_F, 'Sector']):
                    b, imp, f, sector = grp
                    
                    if sector in ('', 'Sin clasificar'):
                        sub50 = pendientes_50[(pendientes_50[col_banco] == b) & (pendientes_50['Abs_I'] == imp) & (pendientes_50[col_F] == f)]
                    else:
                        sub50 = pendientes_50[(pendientes_50[col_banco] == b) & (pendientes_50['Abs_I'] == imp) & (pendientes_50[col_F] == f) & (pendientes_50['Sector'].isin([sector, 'Sin clasificar']))]
                    
                    if sub50.empty: continue
                    s40_ord = sub40[~sub40['ID_Linea'].isin(usados)].sort_values('ID_Linea')
                    s50_ord = sub50[~sub50['ID_Linea'].isin(usados)].sort_values('ID_Linea')
                    n_pares = min(len(s40_ord), len(s50_ord))
                    for i in range(n_pares):
                        id40, id50 = s40_ord.iloc[i]['ID_Linea'], s50_ord.iloc[i]['ID_Linea']
                        texto_cand = f"{formato_linea(id40)} | {formato_linea(id50)}"
                        for idx in (id40, id50):
                            df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Conciliado'
                            df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                            df.loc[df['ID_Linea'] == idx, 'Comentario'] = 'Cruce por orden FIFO.'
                        usados.update([id40, id50])
                        ind_fifo_ok.update([id40, id50])

                    sobrantes40 = s40_ord[~s40_ord['ID_Linea'].isin(usados)]
                    for _, fila in sobrantes40.iterrows():
                        idl = fila['ID_Linea']
                        if bool(fila['B_Repite']):
                            df.loc[df['ID_Linea'] == idl, 'Estado_Conciliacion'] = 'Pendiente o solicitar soporte'
                            df.loc[df['ID_Linea'] == idl, 'Comentario'] = "Múltiples posiciones sin cruzar."
                            ind_fifo_verde_dz.add(idl)

                # ================================================================
                # REGLA 6B — RECLASIFICACIÓN SIN REFERENCIA (último recurso)
                # =====================================================
                df_40 = df[(df[col_G] == '40') & (~df['Es_IP_G40']) & (~df['ID_Linea'].isin(usados))].copy()
                df_50 = df[(df[col_G] == '50') & (~df['Es_CB_G50']) & (~df['ID_Linea'].isin(usados))].copy()
                
                pendientes_40b = df[(df['ID_Linea'].isin(df_40['ID_Linea'])) & (~df['ID_Linea'].isin(usados))]
                pendientes_50b = df[(df['ID_Linea'].isin(df_50['ID_Linea'])) & (~df['ID_Linea'].isin(usados))]

                for (fecha_z, importe_z), grupo40 in pendientes_40b.groupby([col_F, 'Abs_I']):
                    if importe_z > 0 and importe_z % multiplo_redondo == 0:
                        continue
                    grupo40 = grupo40[~grupo40['ID_Linea'].isin(usados)]
                    if grupo40.empty: continue
                    grupo50 = pendientes_50b[(pendientes_50b[col_F] == fecha_z) & (pendientes_50b['Abs_I'] == importe_z) & (~pendientes_50b['ID_Linea'].isin(usados))]
                    if grupo50.empty: continue

                    def limpiar_nombre_banco(nombre_banco):
                        nombre = str(nombre_banco).upper()
                        if 'BANCOLOMBIA' in nombre: return 'Bancolombia'
                        if 'DAVIVIENDA' in nombre: return 'Davivienda'
                        if 'DAVIBANK' in nombre: return 'Davivienda'
                        if 'BOGOTA' in nombre: return 'Banco de Bogotá'
                        if 'CAJA SOCIAL' in nombre: return 'Banco Caja Social'
                        if 'BILBAO' in nombre or 'BBVA' in nombre: return 'BBVA'
                        if 'AGRARIO' in nombre: return 'Banco Agrario'
                        if 'AV V' in nombre or 'VILLAS' in nombre: return 'Banco AV Villas'
                        if 'OCCIDENTE' in nombre: return 'Banco de Occidente'
                        if 'SUDAMERIS' in nombre: return 'Banco GNB Sudameris'
                        return nombre.title()

                    if len(grupo40) == 1 and len(grupo50) == 1:
                        id40 = grupo40.iloc[0]['ID_Linea']
                        id50 = grupo50.iloc[0]['ID_Linea']
                        ra = df.loc[df['ID_Linea'] == id40].iloc[0]
                        rb = df.loc[df['ID_Linea'] == id50].iloc[0]
                        texto_cand = f"{formato_linea(id40)} | {formato_linea(id50)}"
                        if str(ra[col_banco]).strip() == str(rb[col_banco]).strip():
                            estado = 'Pendiente o solicitar soporte'
                            comentario = "Cruce único pero con sector distinto."
                        else:
                            estado = 'Reclasificacion de Banco'
                            banco_a_limpio = limpiar_nombre_banco(ra[col_banco])
                            banco_b_limpio = limpiar_nombre_banco(rb[col_banco])
                            comentario = f"Registrado en '{banco_a_limpio}'; debe ser '{banco_b_limpio}'."
                        for idx in (id40, id50):
                            df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = estado
                            df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                            df.loc[df['ID_Linea'] == idx, 'Comentario'] = comentario
                        usados.update([id40, id50])
                    else:
                        docs40_txt = resumen_docs(grupo40)
                        docs50_txt = resumen_docs(grupo50)
                        for _, r in grupo40.iterrows():
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Estado_Conciliacion'] = 'Pendiente o solicitar soporte'
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Candidatos_Conciliacion'] = f"{formato_linea(r['ID_Linea'])} | Candidatos posibles: {docs50_txt}"
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Comentario'] = "Múltiples candidatos para reclasificación."
                        for _, r in grupo50.iterrows():
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Estado_Conciliacion'] = 'Pendiente o solicitar soporte'
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Candidatos_Conciliacion'] = f"{formato_linea(r['ID_Linea'])} | Candidatos posibles: {docs40_txt}"
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Comentario'] = "Múltiples candidatos para reclasificación."

                # ================================================================
                # VALIDACION FINAL DE FECHA (Sin bloqueo de blancos)
                # ================================================================
                for id40, id50 in parejas_registradas:
                    estado40 = str(df.loc[df['ID_Linea'] == id40, 'Estado_Conciliacion'].iloc[0])
                    estado50 = str(df.loc[df['ID_Linea'] == id50, 'Estado_Conciliacion'].iloc[0])
                    es_flex = 'referencia parcial' in estado40 or 'referencia parcial' in estado50
                    if es_flex: continue

                    dias = diferencia_dias_fila(id40, id50)
                    if dias is not None and dias > TOPE_DIAS_ALERTA:
                        for idx in [id40, id50]:
                            df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Diferencia de fecha > 4 días'
                            df.loc[df['ID_Linea'] == idx, 'Comentario'] = f'Diferencia de fecha extensa ({dias} días).'

                # =====================================================
                # MODO TARDE (SEGUNDA PASADA PROFUNDA T1 - T6)
                # =====================================================
                if modo_tarde:
                    p40 = df[(df[col_G] == '40') & (~df['Es_IP_G40']) & (~df['ID_Linea'].isin(usados))]
                    p50 = df[(df[col_G] == '50') & (~df['Es_CB_G50']) & (~df['ID_Linea'].isin(usados))]
                    for id40, f40 in p40.iterrows():
                        if id40 in usados: continue
                        cand = p50[(p50[col_banco] == f40[col_banco]) & (p50['Abs_I'] == f40['Abs_I']) & (~p50['ID_Linea'].isin(usados))].copy()
                        if cand.empty: continue
                        cand = cand[cand[col_H].apply(lambda h: referencias_se_contienen(f40[col_A], h))]
                        if not cand.empty:
                            id50 = cand.iloc[0]['ID_Linea']
                            for idx in [id40, id50]:
                                df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Conciliado'
                                df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = f"{formato_linea(id40)} | {formato_linea(id50)}"
                                df.loc[df['ID_Linea'] == idx, 'Comentario'] = "Cruce por referencia parcial (forzado)."
                            usados.update([id40, id50])
                            parejas_registradas.append((id40, id50))

                    p40 = df[(df[col_G] == '40') & (~df['Es_IP_G40']) & (~df['ID_Linea'].isin(usados)) & (df['Sector'] != 'Sin clasificar')]
                    for (banco, sector, importe), g40 in p40.groupby([col_banco, 'Sector', 'Abs_I']):
                        g50 = df[(df[col_G] == '50') & (~df['Es_CB_G50']) & (df[col_banco] == banco) & (df['Sector'] == sector) & (df['Abs_I'] == importe) & (~df['ID_Linea'].isin(usados))].copy()
                        if g50.empty: continue
                        g40 = g40.sort_values(col_B)
                        g50 = g50.sort_values(col_B)
                        for id40, id50 in zip(g40['ID_Linea'], g50['ID_Linea']):
                            if id40 in usados or id50 in usados: continue
                            dias = diferencia_dias_fila(id40, id50)
                            if dias is not None:
                                if dias == 0:
                                    est = 'Conciliado'
                                elif dias <= 1:
                                    est = 'Diferencia de fecha'
                                else:
                                    est = 'Diferencia de fecha > 4 días'
                                    
                                for idx in [id40, id50]:
                                    df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = est
                                    df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = f"{formato_linea(id40)} | {formato_linea(id50)}"
                                    df.loc[df['ID_Linea'] == idx, 'Comentario'] = f"Cruce por orden FIFO (dif {dias} días)."
                                usados.update([id40, id50])
                                parejas_registradas.append((id40, id50))

                    p40 = df[(df[col_G] == '40') & (~df['Es_IP_G40']) & (~df['ID_Linea'].isin(usados)) & (df['Sector'] != 'Sin clasificar')]
                    p50 = df[(df[col_G] == '50') & (~df['Es_CB_G50']) & (~df['ID_Linea'].isin(usados))]
                    for id40, f40 in p40.iterrows():
                        cand = p50[(p50[col_banco] == f40[col_banco]) & (p50['Sector'] == f40['Sector']) & (~p50['ID_Linea'].isin(usados))].copy()
                        if cand.empty: continue
                        cand['_dif_dias'] = (cand['Fecha_F'] - f40['Fecha_F']).dt.days.abs().fillna(999)
                        cand = cand[cand['_dif_dias'] <= TOPE_DIAS_ALERTA]
                        if cand.empty: continue
                        cand['_dif_valor'] = (cand['Abs_I'] - f40['Abs_I']).abs()
                        cand = cand[cand['_dif_valor'] > tol_valor_purpura]
                        if not cand.empty:
                            cand = cand.sort_values(['_dif_valor', '_dif_dias'])
                            id50 = cand.iloc[0]['ID_Linea']
                            dif_val = cand.iloc[0]['_dif_valor']
                            dif_d = int(cand.iloc[0]['_dif_dias'])
                            for idx in [id40, id50]:
                                df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Diferencia de valor'
                                df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = f"{formato_linea(id40)} | {formato_linea(id50)}"
                                df.loc[df['ID_Linea'] == idx, 'Comentario'] = f"Diferencia de valor: ${dif_val:,.0f} (dif {dif_d} días)."
                            usados.update([id40, id50])

                    p40_h = df[(df[col_G] == '40') & (~df['Es_IP_G40']) & (~df['ID_Linea'].isin(usados))]
                    p50_h = df[(df[col_G] == '50') & (~df['Es_CB_G50']) & (~df['ID_Linea'].isin(usados))]
                    for imp, g40 in p40_h.groupby('Abs_I'):
                        if imp > 0 and imp % multiplo_redondo == 0: continue
                        g50 = p50_h[p50_h['Abs_I'] == imp]
                        if len(g40) == 1 and len(g50) == 1:
                            id40, id50 = g40.iloc[0]['ID_Linea'], g50.iloc[0]['ID_Linea']
                            dias_lejos = abs((g40.iloc[0]['Fecha_F'] - g50.iloc[0]['Fecha_F']).days)
                            for idx in [id40, id50]:
                                df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Pendiente o solicitar soporte'
                                df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = f"{formato_linea(id40)} | {formato_linea(id50)}"
                                df.loc[df['ID_Linea'] == idx, 'Comentario'] = f"Importe exacto (dif {dias_lejos} días)."
                            usados.update([id40, id50])

                    micro = df[(df[col_G] == '40') & (~df['Es_IP_G40']) & (~df['ID_Linea'].isin(usados)) & (df['Abs_I'] <= 10000)]
                    palabras = ['GMF', 'COMISION', 'IVA', 'RETENCION', '4X1000', 'GRAVAMEN', 'INTERESES', 'RETEICA', 'RETEFUENTE']
                    for id_m, fila_m in micro.iterrows():
                        txt = f"{fila_m.get(col_K, '')} {fila_m.get(col_novedad, '')}".upper()
                        if any(p in txt for p in palabras):
                            df.loc[df['ID_Linea'] == id_m, 'Estado_Conciliacion'] = 'Pendiente o solicitar soporte'
                            df.loc[df['ID_Linea'] == id_m, 'Comentario'] = f"Posible gasto bancario."
                            usados.add(id_m)

                # =====================================================
                # CIERRE
                # =====================================================
                sin_p = df['Estado_Conciliacion'] == 'Pendiente'
                if usar_ipcb:
                    df.loc[sin_p & df['Es_IP_G40'] & (df['Comentario'] == ''), 'Comentario'] = 'Falta Referencia POS.'
                    df.loc[sin_p & ~df['Es_IP_G40'] & (df['Comentario'] == ''), 'Comentario'] = 'Revisión manual requerida.'
                else:
                    df.loc[sin_p & (df['Comentario'] == ''), 'Comentario'] = 'Revisión manual requerida.'

                df_final = df.drop(columns=['ID_Linea', 'Abs_I', 'Fecha_F', 'Fecha_D', 'Es_IP_G40', 'Es_CB_G50'], errors='ignore')
                df_final['Estado_Tecnico'] = df_final['Estado_Conciliacion']
                df_final['Comentario_Tecnico'] = df_final['Comentario']

                # Asignamos el color con la nueva lógica de fechas
                def color_fila(row):
                    est = str(row['Estado_Conciliacion']).strip()
                    com = str(row['Comentario']).strip()
                    
                    if 'Múltiples posiciones sin cruzar' in com: return [f'background-color: {COLOR_VERDE}; color: black'] * len(row)
                    if 'Sugerencia POS' in com or 'Cruce exacto homologado (POS)' in com: 
                        return [f'background-color: {COLOR_GRIS}; color: black'] * len(row)
                        
                    if est == 'Diferencia de fecha > 4 días' or 'extensa' in com: return [f'background-color: {COLOR_SALMON}; color: black'] * len(row)
                    if est == 'Conciliado': return [f'background-color: {COLOR_AZUL}; color: black'] * len(row)
                    if est == 'Reclasificacion de Banco': return [f'background-color: {COLOR_DURAZNO}; color: black'] * len(row)
                    
                    if est == 'Diferencia de fecha' or est == 'Diferencia de valor': return [f'background-color: {COLOR_MORADO}; color: black'] * len(row)
                    if 'forzado' in com or 'Tarde' in com: return [f'background-color: {COLOR_AMARILLO}; color: black'] * len(row)
                    
                    return [f'background-color: {COLOR_BLANCO}; color: black'] * len(row)

                columnas_visibles = columnas_originales + ['Estado_Conciliacion', 'Comentario', 'Candidatos_Conciliacion', 'Sector']

                def vista(df_cualquiera):
                    cols = [c for c in columnas_visibles if c in df_cualquiera.columns]
                    return df_cualquiera[cols].copy()

                cuadre_ok = filas_antes == (len(df) + len(filas_descartadas))
                for c in [c for c in df_final.columns if 'fe.' in c.lower() or 'fecha' in c.lower() or 'fe-' in c.lower()]:
                    df_final[c] = pd.to_datetime(df_final[c], errors='coerce').dt.strftime('%d/%m/%Y')

                output = io.BytesIO()
                b_unicos = [b for b in df_final[col_banco].unique() if str(b).strip().lower() not in ('', 'nan')]
                orden_cuentas = [
                    "1110056001", "1110056101", "1110056201", "1110056301", "1110056401",
                    "1110056501", "1110056601", "1110056701", "1120055001", "1120055101", "1120055301"
                ]
                nombres_ordenados = [mapeo_cuentas_banco.get(c, f"CUENTA {c} (sin mapear)") for c in orden_cuentas]

                def orden_banco(bs):
                    bs = str(bs).strip()
                    if bs in nombres_ordenados: return nombres_ordenados.index(bs)
                    for i, acc in enumerate(orden_cuentas):
                        if acc in bs: return i
                    return 999
                b_unicos = sorted(b_unicos, key=orden_banco)

                pestanas_usadas = set()
                def nombre_pestana(base):
                    nombre = re.sub(r'[\/*?:\[\]]', '-', str(base)[:31])
                    if not nombre.strip() or nombre.lower() == 'nan': nombre = "Sin_Banco"
                    original, cont = nombre, 1
                    while nombre in pestanas_usadas:
                        suf = f"_{cont}"
                        nombre = original[:31-len(suf)] + suf
                        cont += 1
                    pestanas_usadas.add(nombre)
                    return nombre

                advertencias = []
                def hoja_segura(writer, df_hoja, nombre, estilo=True):
                    nf = nombre_pestana(nombre)
                    try:
                        if estilo and not df_hoja.empty:
                            df_hoja.style.apply(lambda row: color_fila(row), axis=1).to_excel(writer, index=False, sheet_name=nf)
                        else:
                            df_hoja.to_excel(writer, index=False, sheet_name=nf)
                    except Exception as e1:
                        try:
                            df_hoja.to_excel(writer, index=False, sheet_name=nf)
                            advertencias.append(f"Hoja '{nf}': sin colores por error de formato ({e1}).")
                        except Exception as e2:
                            advertencias.append(f"Hoja '{nf}': no se pudo escribir ({e2}).")

                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    total_filas = len(df_final)
                    
                    total_azul = int((df_final['Estado_Conciliacion'] == 'Conciliado').sum())
                    total_durazno = int((df_final['Estado_Conciliacion'] == 'Reclasificacion de Banco').sum())
                    
                    mask_salmon_color = df_final['Estado_Conciliacion'].str.contains('> 4 días', na=False)
                    total_salmon = int(mask_salmon_color.sum())
                    
                    mask_morado_color = df_final['Estado_Conciliacion'].isin(['Diferencia de fecha', 'Diferencia de valor'])
                    total_morado = int(mask_morado_color.sum())
                    
                    mask_verde = (df_final[col_G] == '40') & df_final['Comentario'].str.contains('Múltiples posiciones sin cruzar', na=False)
                    total_verde = int(mask_verde.sum())
                    
                    mask_gris = df_final['Comentario'].str.contains('POS', na=False)
                    total_gris = int(mask_gris.sum())
                    
                    total_pendiente = int(df_final['Estado_Conciliacion'].str.contains('Pendiente', na=False).sum())
                    
                    resumen = pd.DataFrame({
                        "Métrica": [
                            "Fecha de procesamiento", "Total filas procesadas",
                            "Azul - Conciliados Exactos",
                            "Gris - Cruces IP/CB (POS o Suma sugerida)",
                            "Verde - Documentos DZ multiposición sin conciliar",
                            "Morado - Diferencia de valor O fecha de 1 día",
                            "Salmón - Diferencia de fecha (> 1 día)",
                            "Durazno - Reclasificación de banco",
                            "Blanco - Pendientes / Otras Sugerencias",
                            "Filas excluidas (sin doc/clave)"
                        ],
                        "Valor": [
                            datetime.now().strftime('%d/%m/%Y %H:%M'), total_filas,
                            total_azul, total_gris, total_verde, total_morado, total_salmon, total_durazno,
                            total_pendiente, filas_excluidas
                        ]
                    })
                    resumen.to_excel(writer, index=False, sheet_name='RESUMEN')
                    pestanas_usadas.add('RESUMEN')

                    df_nov = df_final[df_final[col_G] == '40'].copy()
                    mask_alerta = df_nov['Estado_Conciliacion'].isin(['Diferencia de fecha', 'Diferencia de fecha > 4 días', 'Diferencia de valor', 'Reclasificacion de Banco', 'Pendiente o solicitar soporte'])
                    
                    df_nov = df_nov[mask_alerta]
                    if not df_nov.empty:
                        df_nov = df_nov.sort_values(by=['Estado_Conciliacion', col_I])
                        hoja_segura(writer, vista(df_nov), 'NOVEDADES_Y_PENDIENTES_40', estilo=True)
                    else:
                        hoja_segura(writer, pd.DataFrame(columns=columnas_visibles), 'NOVEDADES_Y_PENDIENTES_40', estilo=False)

                    if modo_tarde:
                        df_tarde = df_final[df_final['Comentario'].str.contains('forzado|gasto', case=False, na=False)].copy()
                        if not df_tarde.empty:
                            df_tarde['Abs_I'] = pd.to_numeric(df_tarde[col_I], errors='coerce').fillna(0).abs()
                            df_tarde = df_tarde.sort_values(by=['Abs_I', col_F])
                            hoja_segura(writer, vista(df_tarde), 'REVISION_TARDE', estilo=True)

                    df_multi = df_final[df_final['B_Repite'] == True].copy() if 'B_Repite' in df_final.columns else pd.DataFrame()
                    if not df_multi.empty:
                        df_multi = df_multi.sort_values(by=[col_banco, col_I, col_F, col_H, col_B])
                        hoja_segura(writer, vista(df_multi), 'REVISAR_POSICIONES_MULTIPLES', estilo=True)

                    for banco in b_unicos:
                        df_b = df_final[df_final[col_banco] == banco].copy().sort_values(by=col_I, ascending=True)
                        if df_b.empty: continue
                        hoja_segura(writer, vista(df_b), str(banco), estilo=True)

                    if not filas_descartadas.empty:
                        hoja_segura(writer, vista(filas_descartadas), 'DESCARTADAS_SIN_DOC_O_CT', estilo=False)

                st.success("¡Conciliación completada con el motor de reglas CLM v40 (Híbrida Corregida Final)!")
                if not cuadre_ok:
                    st.warning("⚠️ Revisa la pestaña DESCARTADAS, el total de filas no coincide.")
                for adv in advertencias:
                    st.warning(f"⚠️ {adv}")

                c1, c2, c3, c4, c5, c6 = st.columns(6)
                c1.metric("Azul", total_azul)
                c2.metric("Gris", total_gris)
                c3.metric("Morado", total_morado)
                c4.metric("Salmón", total_salmon)
                c5.metric("Durazno", total_durazno)
                c6.metric("Pendiente", total_pendiente)

                if filas_excluidas > 0:
                    st.warning(f"⚠️ Se excluyeron {filas_excluidas} filas vacías/totales.")

                st.download_button(
                    label="📥 Descargar Excel con Resultados",
                    data=output.getvalue(),
                    file_name="Conciliacion_CLM_v40_Hibrida_Corregida_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Error técnico detectado: {e}")
            st.exception(e)
