# app_conciliacion_v36_integral.py
#
# VERSION INTEGRAL v36
# Incluye todas las correcciones solicitadas:
# - Comentarios simplificados y jerarquía visual de colores.
# - Rango H Nequi: 100000 - 9999999 y 1000000000 - 1399999999.
# - Flexibilidad para Nequi: Asignación 'T', 'T-', 'T/', '/' y errores de escritura (NEQUI, NEQI, etc).
# - Legalizaciones (G=40, C=DZ): Cruza si tiene palabra/prefijo (no exige número H).
# - Créditos (G=50): Cruza si el rango H se cumple.
# - Desambiguación FIFO estricta para el caso donde Nequi tiene 4 valores idénticos el mismo día.
# - Reparación del bug de `.str.strip()` en Regla 4 (línea 885).
# - Motor a prueba de fallos sin frenos por tipos de datos.

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime

st.set_page_config(page_title="Conciliación Integral CLM", layout="wide")
st.markdown('''
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
''', unsafe_allow_html=True)

st.title("🏦 Conciliación Automatizada — Motor Integral v36 🤖")
st.write("Sube tu archivo consolidado.")
st.caption(
    "Selecciona opcion tarde despues de depurar."
)

with st.expander("⚙️ Parámetros de tolerancia"):
    TOPE_DIAS_ALERTA = st.slider(
        "Días máximos de diferencia de fecha F para alerta (Regla 7, tope fijo 4)",
        1, 4, 4
    )
    tol_valor_purpura = st.number_input(
        "Diferencia máxima de valor ($) para alerta MORADA (Regla morado, tope 500)",
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

    st.divider()
    modo_tarde = st.checkbox(
        "🌅 Activar Casilla 'Tarde' (Segunda pasada en Pendientes)",
        value=False,
        help="Ejecuta una segunda pasada profunda (T1 a T6) y genera una pestaña ordenada de menor a mayor."
    )

COLOR_AZUL = "#C5D9F1"      # Conciliado (todas las reglas cumplen)
COLOR_VERDE = "#A9D18E"     # EXCLUSIVO: DZ multiposición sin conciliar
COLOR_SALMON = "#F5B7A1"    # Diferencia de fecha (hasta 4 días) - Regla 7
COLOR_MORADO = "#C39BD3"    # Diferencia de valor máx $500 - Regla morado
COLOR_DURAZNO = "#FAD7A0"   # Reclasificación de banco - Regla 6
COLOR_BLANCO = "#FFFFFF"    # Pendiente / Sugerencias estándar
COLOR_GRIS = "#D0CECE"      # Cruces múltiples de documento IP (Regla 3)
COLOR_AMARILLO = "#FFF2CC"  # Sugerencias del Modo Tarde

archivo_subido = st.file_uploader("Selecciona el archivo de Excel o CSV", type=['xlsx', 'csv'])

if archivo_subido is not None:
    if st.button("🚀 Ejecutar Conciliación", use_container_width=True):
        try:
            with st.spinner("Ejecutando motor integral de reglas CLM v36... Esto puede tomar unos segundos."):

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

                # ---- Mapeo CLM -> nombres reales de columna ----
                col_A = 'Asignación' if 'Asignación' in df.columns else ('Asignacion' if 'Asignacion' in df.columns else None)
                col_B = 'Nº documento' if 'Nº documento' in df.columns else 'Nº doc.'
                col_C = 'Clase de documento' if 'Clase de documento' in df.columns else ('Clase doc.' if 'Clase doc.' in df.columns else None)
                col_D = 'Fe.contabilización' if 'Fe.contabilización' in df.columns else ('Fecha de documento' if 'Fecha de documento' in df.columns else None)
                col_F = 'Fecha valor' if 'Fecha valor' in df.columns else 'Fe-valor'
                col_G = 'Clave contabiliz.' if 'Clave contabiliz.' in df.columns else 'CT'
                col_H = 'Referencia'
                col_I = 'Importe en moneda local' if 'Importe en moneda local' in df.columns else 'Importe en ML'
                col_K = 'Texto' if 'Texto' in df.columns else None
                col_novedad = 'novedad' if 'novedad' in df.columns else ('Novedad' if 'Novedad' in df.columns else None)
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
                # 2. AUTOCOMPLETADO DE BANCO (Cuentas de Mayor)
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
                # 4. SECTORIZACIÓN (Regla 2)
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
                # NUEVO MAPEO ESTRICTO SEGUN BD
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
                    texto_a = str(row.get(col_A, ""))
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

                    t_full = f"{texto_k} {texto_nov} {texto_a} {h_val}".upper()
                    nums = re.findall(r' \d{4} ', t_full)
                    for n in nums:
                        num = int(n)
                        if 2000 <= num <= 2999: return 'Dist Buga'
                        if 3000 <= num <= 3999: return 'Dist Acopi'
                        if 4000 <= num <= 4999: return 'Dist Dosquebradas'
                        if 6000 <= num <= 6999: return 'Dist Pasto'

                    return 'Sin clasificar'

                df['Sector'] = df.apply(clasificar_sector, axis=1)

                def obtener_ref_homologada(row):
                    texto = f" {row.get(col_H,'')} {row.get(col_A,'')} {row.get(col_K,'') if col_K else ''} {row.get(col_novedad,'') if col_novedad else ''} ".upper()
                    # 1. Check for valid 8-digit codes
                    n8 = re.findall(r'\b\d{8}\b', texto)
                    for n in n8:
                        if n in dict_8_to_list4:
                            return n # Return the 8-digit reference as group ID
                    # 2. Check for valid 4-digit codes mapping to 8-digit code
                    n4 = re.findall(r'\b\d{4}\b', texto)
                    for n in n4:
                        for k8, list_4 in dict_8_to_list4.items():
                            if n in list_4:
                                return k8
                    return None

                def es_nequi(row):
                    val_g = str(row.get(col_G, '')).strip()
                    val_c = str(row.get(col_C, '')).strip().upper() if col_C is not None else ''
                    val_h = str(row.get(col_H, '')).strip()
                    val_h_num_str = re.sub(r'\D', '', val_h)
                    
                    is_h_nequi_range = False
                    if val_h_num_str.isdigit():
                        val_h_num = int(val_h_num_str)
                        if (100000 <= val_h_num <= 9999999) or (1000000000 <= val_h_num <= 1399999999):
                            is_h_nequi_range = True

                    # Recolectamos texto de posibles columnas donde se escriba Nequi
                    texto = f"{row.get(col_K,'') if col_K else ''} {row.get(col_A,'')} {val_h}".upper()
                    if col_novedad:
                        texto += f" {str(row.get(col_novedad, '')).upper()}"

                    # Regex con variaciones ortográficas comunes de Nequi
                    tiene_palabra_nequi = bool(re.search(r'NEQUI|NEQI|NQUI|NEQUY|MEQUI|NEKUI|NEQ', texto))
                    
                    # Verificamos si la Asignación (A) tiene los prefijos/códigos exactos de Nequi
                    val_a = str(row.get(col_A, '')).strip().upper()
                    tiene_prefijo_a = val_a == 'T' or val_a.startswith('T-') or val_a.startswith('T/') or val_a == '/'
                    
                    # Cualquiera de los dos (palabra o prefijo) cuenta como un indicador válido de Nequi
                    es_indicador_texto = tiene_palabra_nequi or tiene_prefijo_a

                    # 1. Validación para Créditos (G = 50)
                    if val_g == '50':
                        if is_h_nequi_range:
                            return True

                    # 2. Validación para Legalizaciones (C = DZ y G = 40)
                    if val_g == '40' and val_c == 'DZ':
                        # SOLO requiere el indicador de texto (no rango numérico en H) para evitar bloqueos
                        if es_indicador_texto:
                            return True
                        return False

                    # 3. Validaciones heredadas para otros documentos (No DZ)
                    if es_indicador_texto: 
                        return True
                        
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

                df['Es_Nequi'] = df.apply(es_nequi, axis=1)
                df['H_Limpia'] = df[col_H].apply(limpiar_numero)
                df['A_Limpia'] = df[col_A].apply(limpiar_numero)

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

                def registrar_pareja_por_fecha(id40, id50, comentario_base):
                    dias = diferencia_dias_fila(id40, id50)
                    if dias is None or dias > TOPE_DIAS_ALERTA:
                        comentario = f'Cruce bloqueado: diferencia F={dias} dias; el maximo permitido es {TOPE_DIAS_ALERTA} dias.'
                        for idx in [id40, id50]: escribir_comentario(idx, comentario, append=False)
                        return False

                    texto = f'{formato_linea(id40)} | {formato_linea(id50)}'
                    if dias == 0:
                        estado = 'Conciliado - grupo azul'
                    else:
                        estado = 'Diferencia de fecha - grupo salmon'

                    for idx in [id40, id50]:
                        df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = estado
                        df.loc[df['ID_Linea'] == idx, 'Comentario'] = f'{comentario_base} Diferencia F={dias} dia(s), maximo permitido={TOPE_DIAS_ALERTA}.'
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
                        resultado['motivo'] = f"Diferencia de fecha F fuera de rango: {dif_dias} dias > {TOPE_DIAS_ALERTA}"
                        return resultado

                    banco_a = str(ra[col_banco]).strip()
                    banco_b = str(rb[col_banco]).strip()
                    resultado['banco_a'] = banco_a
                    resultado['banco_b'] = banco_b

                    if es_ip:
                        resultado['mismo_banco'] = True
                    else:
                        resultado['mismo_banco'] = (banco_a == banco_b)

                    sector_a = str(ra.get('Sector', '')).strip()
                    sector_b = str(rb.get('Sector', '')).strip()
                    
                    if not ignorar_sector:
                        if sector_a not in ('', 'Sin clasificar') and sector_b not in ('', 'Sin clasificar'):
                            if sector_a != sector_b:
                                resultado['motivo'] = f"Sector distinto ({sector_a} vs {sector_b})"
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
                            resultado['motivo'] = f"Importe distinto (dif ${dif_valor:,.0f})"
                            return resultado

                    resultado['es_nequi'] = bool(ra.get('Es_Nequi', False)) or bool(rb.get('Es_Nequi', False))
                    resultado['ok'] = True
                    resultado['motivo'] = "Cumple reglas transversales"
                    return resultado

                def clasificar_y_registrar(id40, id50, base_txt, ignorar_sector=False):
                    res = gate_seguridad(id40, id50, exigir_importe_exacto=True, tolerancia_valor=tol_valor_purpura, ignorar_sector=ignorar_sector)
                    if not res['ok']:
                        return False, res['motivo']

                    texto_candidatos = f"{formato_linea(id40)} | {formato_linea(id50)}"
                    partes_comentario = [base_txt]
                    estado_final = 'Conciliado - Cumple todas las reglas'

                    if res['es_ip']:
                        if res['dif_dias'] and res['dif_dias'] > 0:
                            estado_final = 'Diferencia de fecha'
                            partes_comentario.append(f"Diferencia de fecha: F40 vs F50 difieren {res['dif_dias']} día(s) (tope {TOPE_DIAS_ALERTA}).")
                        elif res['dif_valor'] and res['dif_valor'] > 0:
                            estado_final = 'Diferencia de valor'
                            partes_comentario.append(f"Diferencia de valor: dif=${res['dif_valor']:,.0f} ({res['pct_valor']*100:.2f}%).")
                        else:
                            estado_final = 'Conciliado - IP (banco no evaluado, fecha e importe exactos)'
                    elif not res['mismo_banco']:
                        estado_final = 'Reclasificación de banco'
                        partes_comentario.append(f"Reclasificación de banco: registrado en '{res['banco_a']}'; banco esperado '{res['banco_b']}'.")
                    elif res['dif_dias'] and res['dif_dias'] > 0:
                        estado_final = 'Diferencia de fecha'
                        partes_comentario.append(f"Diferencia de fecha: F40 vs F50 difieren {res['dif_dias']} día(s) (tope {TOPE_DIAS_ALERTA}).")
                    elif res['dif_valor'] and res['dif_valor'] > 0:
                        estado_final = 'Diferencia de valor'
                        partes_comentario.append(f"Diferencia de valor: dif=${res['dif_valor']:,.0f} ({res['pct_valor']*100:.2f}%).")
                    else:
                        estado_final = 'Conciliado - Cumple todas las reglas'

                    if res['es_nequi']:
                        partes_comentario.append("[NEQUI: verificar manual]")

                    comentario_final = " ".join(partes_comentario)

                    for idx in (id40, id50):
                        escribir_estado([idx], estado_final, forzar=True)
                        escribir_candidatos(idx, texto_candidatos)
                        escribir_comentario(idx, comentario_final, append=False)

                    parejas_registradas.append((id40, id50))
                    return True, estado_final

                # =====================================================
                # Regla 3: IP Homologados Agrupado
                # =====================================================
                ind_ip_exacto = set()
                ind_ip_tolerancia = set()
                if usar_ipcb:
                    df['Ref_H_Homologada'] = df.apply(obtener_ref_homologada, axis=1)

                    df_ip = df[(df[col_C].astype(str).str.upper() == 'IP') & (df[col_G] == '40') & df['Ref_H_Homologada'].notna()]
                    df_cb = df[(df[col_C].astype(str).str.upper() == 'CB') & (df[col_G] == '50') & df['Ref_H_Homologada'].notna()]

                    if not df_ip.empty and not df_cb.empty:
                        grp_ip = df_ip.groupby([col_banco, 'Ref_H_Homologada'])['Abs_I'].sum().reset_index(name='S_IP')
                        grp_cb = df_cb.groupby([col_banco, 'Ref_H_Homologada'])['Abs_I'].sum().reset_index(name='S_CB')
                        m = pd.merge(grp_cb, grp_ip, on=[col_banco, 'Ref_H_Homologada'])
                        m['DifV'] = (m['S_CB'] - m['S_IP']).abs()
                        max_s = m[['S_CB', 'S_IP']].max(axis=1).clip(lower=1)
                        m['Pct'] = m['DifV'] / max_s

                        exactos = m[m['DifV'].round(2) == 0]
                        con_tol = m[(m['DifV'] > 0) & ((m['DifV'] <= tol_valor_abs_general) | (m['Pct'] <= tol_valor_pct_general))]

                        def procesar_grupo_ip(fila, es_exacto):
                            b, rh = fila[col_banco], fila['Ref_H_Homologada']
                            sub_ip = df_ip[(df_ip[col_banco] == b) & (df_ip['Ref_H_Homologada'] == rh)]
                            sub_cb = df_cb[(df_cb[col_banco] == b) & (df_cb['Ref_H_Homologada'] == rh)]
                            ip_ids = [i for i in sub_ip['ID_Linea'].tolist() if i not in usados]
                            cb_ids = [i for i in sub_cb['ID_Linea'].tolist() if i not in usados]
                            if not ip_ids or not cb_ids: return
                            usados.update(ip_ids + cb_ids)
                            texto_cand = " | ".join(formato_linea(i) for i in ip_ids + cb_ids)
                            if es_exacto:
                                estado = 'Conciliado - Cruce múltiple IP/CB (Regla 3)'
                                ind_ip_exacto.update(ip_ids + cb_ids)
                                txt = f"Cruce múltiple homologado ({len(ip_ids)} IP = {len(cb_ids)} CB). Ref. homologada: {rh}."
                            else:
                                estado = 'Sugerencia - Cruce múltiple IP/CB con diferencia de valor'
                                ind_ip_tolerancia.update(ip_ids + cb_ids)
                                txt = f"Sugerencia IP/CB con diferencia de valor (${fila['DifV']:,.0f} / {fila['Pct']*100:.2f}%). Suma {len(ip_ids)} IP vs {len(cb_ids)} CB. Ref. homologada: {rh}."
                            for idx in ip_ids + cb_ids:
                                df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = estado
                                df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                                otros = resumen_docs(sub_cb) if idx in ip_ids else resumen_docs(sub_ip)
                                df.loc[df['ID_Linea'] == idx, 'Comentario'] = f"{txt} Docs relacionados: {otros}"

                        for _, fila in exactos.iterrows(): procesar_grupo_ip(fila, es_exacto=True)
                        for _, fila in con_tol.iterrows(): procesar_grupo_ip(fila, es_exacto=False)

                # ================================================================
                # RESTRICCIÓN ESTRICTA IP: Si no homologó arriba, queda bloqueado
                # ================================================================
                if usar_ipcb:
                    ip_sin_resolver = df[(df[col_C].astype(str).str.upper() == 'IP') & (df[col_G] == '40') & (~df['ID_Linea'].isin(usados))]
                    for idl in ip_sin_resolver['ID_Linea']:
                        escribir_comentario(idl, "PDV (IP): Sin coincidencia. Requiere referencia homologada estricta según base de datos.", append=False)

                # =====================================================
                # Regla 8: NEQUI POR TOTALES Y FIFO
                # =====================================================
                ind_nequi8_azul = set()
                ind_nequi8_sugerencia = set()

                df_nequi_dz = df[(df[col_G] == '40') & (df['Es_Nequi'] == True) & (~df['ID_Linea'].isin(usados))]
                df_cb_disponible = df[(df[col_G] == '50') & (~df['ID_Linea'].isin(usados))]
                if usar_ipcb:
                    df_cb_disponible = df_cb_disponible[df_cb_disponible[col_C].astype(str).str.upper() != 'IP']

                if not df_nequi_dz.empty and not df_cb_disponible.empty:
                    for (banco_g, fecha_g), grupo_dz in df_nequi_dz.groupby([col_banco, 'Fecha_F']):
                        grupo_dz = grupo_dz[~grupo_dz['ID_Linea'].isin(usados)]
                        if grupo_dz.empty: continue
                        grupo_cb = df_cb_disponible[(df_cb_disponible[col_banco] == banco_g) & (df_cb_disponible['Fecha_F'] == fecha_g) & (~df_cb_disponible['ID_Linea'].isin(usados))]
                        if grupo_cb.empty: continue

                        n_dz, n_cb = len(grupo_dz), len(grupo_cb)
                        total_dz, total_cb = round(grupo_dz['Abs_I'].sum(), 2), round(grupo_cb['Abs_I'].sum(), 2)

                        dz_ord = grupo_dz.sort_values(col_B).reset_index(drop=True)
                        cb_ord = grupo_cb.sort_values(col_B).reset_index(drop=True)
                        n_parejas = min(n_dz, n_cb)

                        if n_dz == n_cb and total_dz == total_cb:
                            texto_grupo = f"total DZ=${total_dz:,.0f}; total CB=${total_cb:,.0f}; cruce FIFO por B ({n_dz} lineas)."
                            for i in range(n_parejas):
                                id40, id50 = dz_ord.iloc[i]['ID_Linea'], cb_ord.iloc[i]['ID_Linea']
                                texto_cand = f"{formato_linea(id40)} | {formato_linea(id50)}"
                                for idx in (id40, id50):
                                    df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Conciliado - Regla 8 Nequi (total y FIFO)'
                                    df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                                    df.loc[df['ID_Linea'] == idx, 'Comentario'] = f"Regla 8 Nequi: {texto_grupo}"
                                ind_nequi8_azul.update([id40, id50])
                                parejas_registradas.append((id40, id50))
                            usados.update(dz_ord['ID_Linea'].tolist() + cb_ord['ID_Linea'].tolist())
                        else:
                            diff_total = abs(total_dz - total_cb)
                            motivo = f"cantidad DZ={n_dz} vs CB={n_cb}" if n_dz != n_cb else f"totales distintos (dif=${diff_total:,.0f})"
                            docs_dz, docs_cb = resumen_docs(dz_ord), resumen_docs(cb_ord)
                            texto_cand = f"Grupo NEQUI {banco_g} {fecha_g}: DZ candidatos: {docs_dz} || CB candidatos: {docs_cb}"
                            comentario = f"Regla 8 Nequi: grupo NO cuadra exacto ({motivo}). Requiere revisión manual."
                            for _, fila in pd.concat([dz_ord, cb_ord]).iterrows():
                                idl = fila['ID_Linea']
                                escribir_estado([idl], 'Sugerencia - Regla 8 Nequi (revisar total de grupo)', forzar=False)
                                if df.loc[df['ID_Linea'] == idl, 'Candidatos_Conciliacion'].iloc[0] == '':
                                    escribir_candidatos(idl, texto_cand)
                                escribir_comentario(idl, comentario, append=False)
                                ind_nequi8_sugerencia.add(idl)

                # =====================================================
                # Regla 1 — A debe coincidir con H (exacto)
                # AISLAMIENTO ESTRICTO DE IP: Los docs IP (40) ya NO participarán en las reglas genéricas
                # =====================================================
                if usar_ipcb:
                    df_40 = df[(df[col_G] == '40') & (df[col_C].astype(str).str.upper() != 'IP')].copy()
                else:
                    df_40 = df[(df[col_G] == '40')].copy()
                df_50 = df[(df[col_G] == '50')].copy()

                def emparejar_1a1_por_llave(sub40, sub50, llave40, llave50, base_txt):
                    s40 = sub40[~sub40['ID_Linea'].isin(usados)].copy()
                    s50 = sub50[~sub50['ID_Linea'].isin(usados)].copy()
                    if s40.empty or s50.empty: return
                    s40['_pos'] = s40.groupby(llave40).cumcount()
                    s50['_pos'] = s50.groupby(llave50).cumcount()
                    merged = pd.merge(s40, s50, left_on=llave40 + ['_pos'], right_on=llave50 + ['_pos'], suffixes=('_40', '_50'))
                    for _, r in merged.iterrows():
                        id40, id50 = r['ID_Linea_40'], r['ID_Linea_50']
                        if id40 in usados or id50 in usados: continue
                        ok, _ = clasificar_y_registrar(id40, id50, base_txt)
                        if ok: usados.update([id40, id50])

                emparejar_1a1_por_llave(df_40, df_50, [col_banco, 'Abs_I', col_A], [col_banco, 'Abs_I', col_H], "Regla 1: Asignación (A) coincide exacta con Referencia (H).")

                df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
                df_50 = df_50[~df_50['ID_Linea'].isin(usados)]
                emparejar_1a1_por_llave(df_40[df_40['A_Limpia'] != ''], df_50[df_50['H_Limpia'] != ''], [col_banco, 'Abs_I', 'A_Limpia'], [col_banco, 'Abs_I', 'H_Limpia'], "Regla 1 (limpia): Asignación limpia coincide con Referencia limpia.")

                # ================================================================
                # PARCHE v33: FLEX POR REFERENCIA PARCIAL (Unica excepcion a 4 dias)
                # ================================================================
                pend40_flex = df[(df[col_G] == '40') & (~df['ID_Linea'].isin(usados))]
                pend50_flex = df[(df[col_G] == '50') & (~df['ID_Linea'].isin(usados))]

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
                    comentario = (
                        'Flex referencia parcial: A y H se contienen mutuamente; '
                        'banco e importe exactos. Esta excepcion permite ignorar '
                        'el limite normal de 4 dias y queda azul.'
                    )
                    for idx in [id40, id50]:
                        df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Conciliado - Flex referencia parcial'
                        df.loc[df['ID_Linea'] == idx, 'Comentario'] = comentario
                        df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto
                    usados.update([id40, id50])
                    parejas_registradas.append((id40, id50))

                # =====================================================
                # Regla 6 EXPLÍCITA — RECLASIFICACIÓN DE BANCO
                # =====================================================
                df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
                df_50 = df_50[~df_50['ID_Linea'].isin(usados)]

                def emparejar_reclasificacion(sub40, sub50, llave40, llave50, base_txt):
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
                        ok, _ = clasificar_y_registrar(id40, id50, base_txt)
                        if ok: usados.update([id40, id50])

                emparejar_reclasificacion(df_40, df_50, ['Abs_I', col_A], ['Abs_I', col_H], "Regla 6: Asignación (A) coincide con Referencia (H), pero el banco registrado difiere.")
                emparejar_reclasificacion(df_40[df_40['A_Limpia'] != ''], df_50[df_50['H_Limpia'] != ''], ['Abs_I', 'A_Limpia'], ['Abs_I', 'H_Limpia'], "Regla 6 (limpia): Asignación limpia coincide con Referencia limpia, pero el banco registrado difiere.")

                # ================================================================
                # PARCHE v33: SECTORIZACION MULTIPLE FIFO
                # ================================================================
                for (banco_g, sector_g, importe_g), lado40 in df[
                    (df[col_G] == '40') & (~df['ID_Linea'].isin(usados)) & (df['Sector'] != 'Sin clasificar')
                ].groupby([col_banco, 'Sector', 'Abs_I']):
                    lado50 = df[
                        (df[col_G] == '50') & (df[col_banco] == banco_g) & (df['Sector'] == sector_g) & (df['Abs_I'] == importe_g) & (~df['ID_Linea'].isin(usados))
                    ].copy()
                    lado40 = lado40.sort_values(col_B)
                    lado50 = lado50.sort_values(col_B)
                    if lado50.empty: continue

                    for id40, id50 in zip(lado40['ID_Linea'], lado50['ID_Linea']):
                        if id40 in usados or id50 in usados: continue
                        registrar_pareja_por_fecha(id40, id50, f'Sectorizacion FIFO en {sector_g}; importe exacto.')

                # Sobrantes de sectorización desbalanceada
                df_40 = df[(df[col_G] == '40') & (~df['ID_Linea'].isin(usados))]
                df_50 = df[(df[col_G] == '50') & (~df['ID_Linea'].isin(usados))]
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
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Estado_Conciliacion'] = f"Sugerencia - Sectorización desbalanceada ({sector})"
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Candidatos_Conciliacion'] = f"{formato_linea(r['ID_Linea'])} | Candidatos posibles: {docs50_txt}"
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Comentario'] = f"Sector '{sector}' desbalanceado. Créditos candidatos: {docs50_txt}"
                        for _, r in sub50.iterrows():
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Estado_Conciliacion'] = f"Sugerencia - Sectorización desbalanceada ({sector})"
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Candidatos_Conciliacion'] = f"{formato_linea(r['ID_Linea'])} | Candidatos posibles: {docs40_txt}"
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Comentario'] = f"Sector '{sector}' desbalanceado. Débitos candidatos: {docs40_txt}"

                # ================================================================
                # PARCHE v33: REGLA 7B - DIFERENCIA DE VALOR (Alertas Sugeridas)
                # ================================================================
                pend40_7b = df[(df[col_G] == '40') & (~df['ID_Linea'].isin(usados)) & (df['Estado_Conciliacion'] == 'Pendiente')].copy()
                pend50_7b = df[(df[col_G] == '50') & (~df['ID_Linea'].isin(usados)) & (df['Estado_Conciliacion'] == 'Pendiente')].copy()

                for id40, fila40 in pend40_7b.iterrows():
                    if fila40['Sector'] == 'Sin clasificar': continue
                    posibles = pend50_7b[
                        (pend50_7b[col_banco] == fila40[col_banco]) &
                        (pend50_7b['Sector'] == fila40['Sector']) &
                        (~pend50_7b['ID_Linea'].isin(usados))
                    ].copy()

                    if posibles.empty: continue

                    posibles['_dif_dias'] = (posibles['Fecha_F'] - fila40['Fecha_F']).dt.days.abs().fillna(999)
                    posibles = posibles[posibles['_dif_dias'] <= TOPE_DIAS_ALERTA]
                    if posibles.empty: continue
                    posibles['_dif_valor'] = (posibles['Abs_I'] - fila40['Abs_I']).abs()
                    posibles = posibles[posibles['_dif_valor'] > tol_valor_purpura]
                    if posibles.empty: continue

                    id50 = posibles.sort_values(['_dif_valor', '_dif_dias']).iloc[0]['ID_Linea']
                    diferencia = round(abs(fila40['Abs_I'] - df.loc[id50, 'Abs_I']), 2)
                    texto = f'{formato_linea(id40)} | {formato_linea(id50)}'
                    comentario = (
                        f'Regla 7B: diferencia de valor ${diferencia:,.2f}; '
                        f'mismo banco/zona y fecha dentro del limite de {TOPE_DIAS_ALERTA} dias. '
                        f'Supera el tope morado de ${tol_valor_purpura:,.0f}, no se concilia automaticamente.'
                    )
                    for idx in [id40, id50]:
                        df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Sugerencia - Regla 7B diferencia de valor'
                        df.loc[df['ID_Linea'] == idx, 'Comentario'] = comentario
                        df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto
                    usados.update([id40, id50])

                # =====================================================
                # EXCEPCIÓN NEQUI (A = "Nequi", C = DZ) - DOBLE PASADA
                # =====================================================
                df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
                df_50 = df_50[~df_50['ID_Linea'].isin(usados)]

                df_nequi_40 = df_40[df_40['Es_Nequi'] == True]

                def intentar_fifo_nequi(id40, candidatos_50, comentario_base):
                    r40 = df.loc[df['ID_Linea'] == id40].iloc[0]
                    banco_r = r40[col_banco]
                    fecha_r = r40['Fecha_F']
                    importe_r = r40['Abs_I']

                    grupo_dz = df[
                        (df[col_G] == '40') & (df['Es_Nequi'] == True) &
                        (~df['ID_Linea'].isin(usados)) &
                        (df[col_banco] == banco_r) &
                        (df['Fecha_F'] == fecha_r) &
                        (df['Abs_I'] == importe_r)
                    ]
                    grupo_cb = candidatos_50[
                        (candidatos_50['Abs_I'] == importe_r) &
                        (~candidatos_50['ID_Linea'].isin(usados))
                    ]

                    n_dz, n_cb = len(grupo_dz), len(grupo_cb)
                    if n_dz == 0 or n_cb == 0:
                        return False

                    dz_ord = grupo_dz.sort_values(col_B).reset_index(drop=True)
                    cb_ord = grupo_cb.sort_values(col_B).reset_index(drop=True)
                    n_pares = min(n_dz, n_cb)

                    for i in range(n_pares):
                        id_dz = dz_ord.iloc[i]['ID_Linea']
                        id_cb = cb_ord.iloc[i]['ID_Linea']
                        if id_dz in usados or id_cb in usados:
                            continue
                        ok, _ = clasificar_y_registrar(
                            id_dz, id_cb,
                            f"{comentario_base} (FIFO desambiguado: emparejado {n_pares} de {max(n_dz, n_cb)} candidatos disponibles, mismo importe ${importe_r:,.0f})",
                            ignorar_sector=True
                        )
                        if ok:
                            usados.update([id_dz, id_cb])
                    
                    return id40 in usados

                def procesar_candidato_nequi(id40, candidatos_50):
                    exactos = candidatos_50[candidatos_50['Abs_I'] == df.loc[df['ID_Linea'] == id40, 'Abs_I'].iloc[0]]
                    if len(exactos) == 1:
                        id50 = exactos.iloc[0]['ID_Linea']
                        ok, _ = clasificar_y_registrar(id40, id50, "Excepción Nequi (cruce importe exacto)", ignorar_sector=True)
                        if ok:
                            usados.update([id40, id50])
                        return True
                    if len(exactos) > 1:
                        if id40 not in usados:
                            resuelto = intentar_fifo_nequi(id40, exactos, "Excepción Nequi (cruce importe exacto)")
                            if resuelto and id40 in usados:
                                return True
                        if id40 in usados:
                            return True
                        
                        docs_txt = resumen_docs(exactos)
                        df.loc[df['ID_Linea'] == id40, 'Estado_Conciliacion'] = 'Sugerencia - Excepción Nequi ambigua'
                        df.loc[df['ID_Linea'] == id40, 'Candidatos_Conciliacion'] = f"{formato_linea(id40)} | Candidatos posibles: {docs_txt}"
                        df.loc[df['ID_Linea'] == id40, 'Comentario'] = (
                            f"Excepción Nequi: {len(exactos)} candidatos con importe exacto, "
                            "requiere selección manual (no se concilia automático por ambigüedad)."
                        )
                        return True

                    r40 = df.loc[df['ID_Linea'] == id40].iloc[0]
                    candidatos_50 = candidatos_50.copy()
                    candidatos_50['_dif_val'] = (candidatos_50['Abs_I'] - r40['Abs_I']).abs()
                    con_tol = candidatos_50[candidatos_50['_dif_val'] <= tol_valor_purpura].sort_values('_dif_val')
                    
                    if len(con_tol) == 1:
                        id50 = con_tol.iloc[0]['ID_Linea']
                        ok, _ = clasificar_y_registrar(id40, id50, "Excepción Nequi (con diferencia de valor)", ignorar_sector=True)
                        if ok:
                            usados.update([id40, id50])
                        return True
                    if len(con_tol) > 1:
                        docs_txt = resumen_docs(con_tol)
                        df.loc[df['ID_Linea'] == id40, 'Estado_Conciliacion'] = 'Sugerencia - Excepción Nequi ambigua'
                        df.loc[df['ID_Linea'] == id40, 'Candidatos_Conciliacion'] = f"{formato_linea(id40)} | Candidatos posibles: {docs_txt}"
                        df.loc[df['ID_Linea'] == id40, 'Comentario'] = (
                            f"Excepción Nequi: {len(con_tol)} candidatos con diferencia de valor dentro del tope morado, "
                            "requiere selección manual (no se concilia automático por ambigüedad)."
                        )
                        return True
                        
                    return False

                # PRIMERA PASADA: BÚSQUEDA MISMO DÍA
                for _, r40 in df_nequi_40.iterrows():
                    id40 = r40['ID_Linea']
                    if id40 in usados: continue
                    candidatos_50 = df_50[
                        (df_50[col_banco] == r40[col_banco]) &
                        (~df_50['ID_Linea'].isin(usados))
                    ].copy()
                    if candidatos_50.empty: continue

                    # Filtro de sector eliminado para Nequi
                    candidatos_50['_dif_dias'] = (candidatos_50['Fecha_F'] - r40['Fecha_F']).dt.days.abs().fillna(999)
                    candidatos_50 = candidatos_50[candidatos_50['_dif_dias'] == 0]
                    if candidatos_50.empty: continue

                    procesar_candidato_nequi(id40, candidatos_50)

                # SEGUNDA PASADA: BÚSQUEDA CON TOLERANCIA DE DÍAS
                for _, r40 in df_nequi_40.iterrows():
                    id40 = r40['ID_Linea']
                    if id40 in usados: continue
                    candidatos_50 = df_50[
                        (df_50[col_banco] == r40[col_banco]) &
                        (~df_50['ID_Linea'].isin(usados))
                    ].copy()
                    if candidatos_50.empty: continue

                    # Filtro de sector eliminado para Nequi
                    candidatos_50['_dif_dias'] = (candidatos_50['Fecha_F'] - r40['Fecha_F']).dt.days.abs().fillna(999)
                    candidatos_50 = candidatos_50[candidatos_50['_dif_dias'] <= TOPE_DIAS_ALERTA]
                    candidatos_50 = candidatos_50.sort_values('_dif_dias')
                    if candidatos_50.empty: continue

                    procesar_candidato_nequi(id40, candidatos_50)

                # =====================================================
                # REGLA 4 — DOCUMENTOS DZ CON POSICIONES MÚLTIPLES
                # =====================================================
                df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
                df_50 = df_50[~df_50['ID_Linea'].isin(usados)]

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
                        df.loc[df['ID_Linea'] == idl, 'Estado_Conciliacion'] = 'Sugerencia - DZ posiciones múltiples (verificar)'
                        df.loc[df['ID_Linea'] == idl, 'Candidatos_Conciliacion'] = texto_cand_final
                        df.loc[df['ID_Linea'] == idl, 'Comentario'] = f"Documento {int(b_doc)} tiene múltiples posiciones/importes distintos. Se listan candidatos en la columna O."

                # =====================================================
                # ÚLTIMO RECURSO: FIFO CONTROLADO
                # =====================================================
                df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
                df_50 = df_50[~df_50['ID_Linea'].isin(usados)]

                pendientes_40 = df[(df['ID_Linea'].isin(df_40['ID_Linea'])) & (df['Estado_Conciliacion'] == 'Pendiente') & (df['Comentario'] == '') & (df['Candidatos_Conciliacion'] == '')]
                pendientes_50 = df[(df['ID_Linea'].isin(df_50['ID_Linea'])) & (df['Estado_Conciliacion'] == 'Pendiente') & (df['Comentario'] == '') & (df['Candidatos_Conciliacion'] == '')]

                ind_fifo_ok = set()
                ind_fifo_verde_dz = set()

                for grp, sub40 in pendientes_40.groupby([col_banco, 'Abs_I', col_F, 'Sector']):
                    b, imp, f, sector = grp
                    sub50 = pendientes_50[(pendientes_50[col_banco] == b) & (pendientes_50['Abs_I'] == imp) & (pendientes_50[col_F] == f) & (pendientes_50['Sector'] == sector)]
                    if sub50.empty: continue
                    s40_ord = sub40[~sub40['ID_Linea'].isin(usados)].sort_values('ID_Linea')
                    s50_ord = sub50[~sub50['ID_Linea'].isin(usados)].sort_values('ID_Linea')
                    n_pares = min(len(s40_ord), len(s50_ord))
                    for i in range(n_pares):
                        id40, id50 = s40_ord.iloc[i]['ID_Linea'], s50_ord.iloc[i]['ID_Linea']
                        texto_cand = f"{formato_linea(id40)} | {formato_linea(id50)}"
                        for idx in (id40, id50):
                            df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Conciliado - FIFO controlado (última instancia)'
                            df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                            df.loc[df['ID_Linea'] == idx, 'Comentario'] = 'Emparejado por FIFO controlado: misma fecha F, mismo sector, banco e importe.'
                        usados.update([id40, id50])
                        ind_fifo_ok.update([id40, id50])

                    sobrantes40 = s40_ord[~s40_ord['ID_Linea'].isin(usados)]
                    for _, fila in sobrantes40.iterrows():
                        idl = fila['ID_Linea']
                        if bool(fila['B_Repite']):
                            df.loc[df['ID_Linea'] == idl, 'Estado_Conciliacion'] = 'Sugerencia - DZ multiposición sin cruce (verificar)'
                            df.loc[df['ID_Linea'] == idl, 'Comentario'] = f"Documento {int(fila[col_B])} tiene varias posiciones y ésta no encontró pareja exacta."
                            ind_fifo_verde_dz.add(idl)

                # ================================================================
                # REGLA 6B — RECLASIFICACIÓN SIN REFERENCIA (último recurso)
                # =====================================================
                df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
                df_50 = df_50[~df_50['ID_Linea'].isin(usados)]
                pendientes_40b = df[(df['ID_Linea'].isin(df_40['ID_Linea'])) & (~df['ID_Linea'].isin(usados))]
                pendientes_50b = df[(df['ID_Linea'].isin(df_50['ID_Linea'])) & (~df['ID_Linea'].isin(usados))]

                for (fecha_z, importe_z), grupo40 in pendientes_40b.groupby([col_F, 'Abs_I']):
                    if importe_z > 0 and importe_z % multiplo_redondo == 0:
                        continue
                    grupo40 = grupo40[~grupo40['ID_Linea'].isin(usados)]
                    if grupo40.empty:
                        continue
                    grupo50 = pendientes_50b[
                        (pendientes_50b[col_F] == fecha_z) & (pendientes_50b['Abs_I'] == importe_z) &
                        (~pendientes_50b['ID_Linea'].isin(usados))
                    ]
                    if grupo50.empty:
                        continue

                    if len(grupo40) == 1 and len(grupo50) == 1:
                        id40 = grupo40.iloc[0]['ID_Linea']
                        id50 = grupo50.iloc[0]['ID_Linea']
                        ra = df.loc[df['ID_Linea'] == id40].iloc[0]
                        rb = df.loc[df['ID_Linea'] == id50].iloc[0]
                        texto_cand = f"{formato_linea(id40)} | {formato_linea(id50)}"
                        if str(ra[col_banco]).strip() == str(rb[col_banco]).strip():
                            estado = 'Sugerencia - Cruce único sin referencia (sector no coincide)'
                            comentario = ("Único candidato en Fecha valor + importe exactos, mismo "
                                          "banco, pero el Sector no coincide o no está clasificado.")
                        else:
                            estado = 'Reclasificación de banco'
                            comentario = (
                                f"Regla 6B: sin coincidencia de Asignación/Referencia ni de Sector, "
                                f"pero es el ÚNICO candidato con la misma Fecha valor e importe exacto "
                                f"(${importe_z:,.0f}). Registrado en '{ra[col_banco]}'; banco esperado "
                                f"'{rb[col_banco]}'."
                            )
                        for idx in (id40, id50):
                            df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = estado
                            df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                            df.loc[df['ID_Linea'] == idx, 'Comentario'] = comentario
                        usados.update([id40, id50])
                    else:
                        docs40_txt = resumen_docs(grupo40)
                        docs50_txt = resumen_docs(grupo50)
                        for _, r in grupo40.iterrows():
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Estado_Conciliacion'] = 'Sugerencia - Reclasificación con múltiples candidatos'
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Candidatos_Conciliacion'] = f"{formato_linea(r['ID_Linea'])} | Candidatos posibles: {docs50_txt}"
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Comentario'] = f"Fecha+importe exactos (${importe_z:,.0f}) con varios candidatos en distintos bancos: {docs50_txt}"
                        for _, r in grupo50.iterrows():
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Estado_Conciliacion'] = 'Sugerencia - Reclasificación con múltiples candidatos'
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Candidatos_Conciliacion'] = f"{formato_linea(r['ID_Linea'])} | Candidatos posibles: {docs40_txt}"
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Comentario'] = f"Fecha+importe exactos (${importe_z:,.0f}) con varios candidatos en distintos bancos: {docs40_txt}"

                # ================================================================
                # VALIDACION FINAL DE FECHA (Muro de Contencion)
                # ================================================================
                for id40, id50 in parejas_registradas:
                    estado40 = str(df.loc[df['ID_Linea'] == id40, 'Estado_Conciliacion'].iloc[0])
                    estado50 = str(df.loc[df['ID_Linea'] == id50, 'Estado_Conciliacion'].iloc[0])
                    es_flex = 'Flex referencia parcial' in estado40 or 'Flex referencia parcial' in estado50
                    if es_flex: continue

                    dias = diferencia_dias_fila(id40, id50)
                    if dias is None or dias > TOPE_DIAS_ALERTA:
                        for idx in [id40, id50]:
                            df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Pendiente - fecha fuera de rango'
                            df.loc[df['ID_Linea'] == idx, 'Comentario'] = f'Bloqueado por fecha: diferencia F={dias} dias; maximo permitido={TOPE_DIAS_ALERTA}.'
                            df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = ''

                # =====================================================
                # MODO TARDE (SEGUNDA PASADA PROFUNDA T1 - T6)
                # =====================================================
                if modo_tarde:


                    p40 = df[(df[col_G] == '40') & (~df['ID_Linea'].isin(usados))]
                    p50 = df[(df[col_G] == '50') & (~df['ID_Linea'].isin(usados))]
                    for id40, f40 in p40.iterrows():
                        if id40 in usados: continue
                        cand = p50[(p50[col_banco] == f40[col_banco]) & (p50['Abs_I'] == f40['Abs_I']) & (~p50['ID_Linea'].isin(usados))].copy()
                        if cand.empty: continue
                        cand = cand[cand[col_H].apply(lambda h: referencias_se_contienen(f40[col_A], h))]
                        if not cand.empty:
                            id50 = cand.iloc[0]['ID_Linea']
                            for idx in [id40, id50]:
                                df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Conciliado (Tarde) - Flex referencia parcial'
                                df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = f"{formato_linea(id40)} | {formato_linea(id50)}"
                                df.loc[df['ID_Linea'] == idx, 'Comentario'] = "Tarde T2: A y H se contienen mutuamente (ignora límite de días)."
                            usados.update([id40, id50])
                            parejas_registradas.append((id40, id50))

                    p40 = df[(df[col_G] == '40') & (~df['ID_Linea'].isin(usados)) & (df['Sector'] != 'Sin clasificar')]
                    for (banco, sector, importe), g40 in p40.groupby([col_banco, 'Sector', 'Abs_I']):
                        g50 = df[(df[col_G] == '50') & (df[col_banco] == banco) & (df['Sector'] == sector) & (df['Abs_I'] == importe) & (~df['ID_Linea'].isin(usados))].copy()
                        if g50.empty: continue
                        g40 = g40.sort_values(col_B)
                        g50 = g50.sort_values(col_B)
                        for id40, id50 in zip(g40['ID_Linea'], g50['ID_Linea']):
                            if id40 in usados or id50 in usados: continue
                            dias = diferencia_dias_fila(id40, id50)
                            if dias is not None and dias <= TOPE_DIAS_ALERTA:
                                est = 'Conciliado (Tarde) - Sector FIFO' if dias == 0 else 'Diferencia de fecha (Tarde) - Sector FIFO'
                                for idx in [id40, id50]:
                                    df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = est
                                    df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = f"{formato_linea(id40)} | {formato_linea(id50)}"
                                    df.loc[df['ID_Linea'] == idx, 'Comentario'] = f"Tarde T3: Sectorización FIFO, dif {dias} días."
                                usados.update([id40, id50])
                                parejas_registradas.append((id40, id50))

                    p40 = df[(df[col_G] == '40') & (~df['ID_Linea'].isin(usados)) & (df['Sector'] != 'Sin clasificar')]
                    p50 = df[(df[col_G] == '50') & (~df['ID_Linea'].isin(usados))]
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
                                df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Sugerencia (Tarde) - T4 Regla 7B'
                                df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = f"{formato_linea(id40)} | {formato_linea(id50)}"
                                df.loc[df['ID_Linea'] == idx, 'Comentario'] = f"Tarde T4: Regla 7B, diferencia valor ${dif_val:,.2f}, dif {dif_d} días."
                            usados.update([id40, id50])

                    p40_h = df[(df[col_G] == '40') & (~df['ID_Linea'].isin(usados))]
                    p50_h = df[(df[col_G] == '50') & (~df['ID_Linea'].isin(usados))]
                    for imp, g40 in p40_h.groupby('Abs_I'):
                        if imp > 0 and imp % multiplo_redondo == 0: continue
                        g50 = p50_h[p50_h['Abs_I'] == imp]
                        if len(g40) == 1 and len(g50) == 1:
                            id40, id50 = g40.iloc[0]['ID_Linea'], g50.iloc[0]['ID_Linea']
                            dias_lejos = abs((g40.iloc[0]['Fecha_F'] - g50.iloc[0]['Fecha_F']).days)
                            for idx in [id40, id50]:
                                df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Sugerencia (Tarde) - T5 Valor exacto huérfano'
                                df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = f"{formato_linea(id40)} | {formato_linea(id50)}"
                                df.loc[df['ID_Linea'] == idx, 'Comentario'] = f"Tarde T5: Importe exacto (${imp:,.0f}), separados por {dias_lejos} días. Ignora banco y fecha."
                            usados.update([id40, id50])

                    micro = df[(df[col_G] == '40') & (~df['ID_Linea'].isin(usados)) & (df['Abs_I'] <= 10000)]
                    palabras = ['GMF', 'COMISION', 'IVA', 'RETENCION', '4X1000', 'GRAVAMEN', 'INTERESES', 'RETEICA', 'RETEFUENTE']
                    for id_m, fila_m in micro.iterrows():
                        txt = f"{fila_m.get(col_K, '')} {fila_m.get(col_novedad, '')}".upper()
                        if any(p in txt for p in palabras):
                            df.loc[df['ID_Linea'] == id_m, 'Estado_Conciliacion'] = 'Sugerencia (Tarde) - T6 Posible Gasto Bancario'
                            df.loc[df['ID_Linea'] == id_m, 'Comentario'] = f"Tarde T6: Micro-saldo (${fila_m['Abs_I']:,.0f}) con texto de gasto/comisión."
                            usados.add(id_m)

                # =====================================================
                # CIERRE
                # =====================================================
                sin_p = df['Estado_Conciliacion'] == 'Pendiente'
                if usar_ipcb:
                    es_ip = df[col_C].astype(str).str.upper() == 'IP'
                    df.loc[sin_p & es_ip & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia - PDV (requiere referencia homologada o cruce exacto)'
                    df.loc[sin_p & ~es_ip & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia ni sugerencia que cumpla reglas de seguridad.'
                else:
                    df.loc[sin_p & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia ni sugerencia que cumpla reglas de seguridad.'

                # =====================================================
                # SIMPLIFICACIÓN DE ESTADOS, COMENTARIOS Y EXPORTACIÓN
                # =====================================================
                df_final = df.drop(columns=['ID_Linea', 'Abs_I', 'Fecha_F', 'Fecha_D'], errors='ignore')

                # Preservamos los valores enriquecidos/técnicos en columnas ocultas
                df_final['Estado_Tecnico'] = df_final['Estado_Conciliacion']
                df_final['Comentario_Tecnico'] = df_final['Comentario']

                def simplificar_estado(est):
                    est_lower = str(est).lower()
                    if 'conciliado' in est_lower or 'grupo azul' in est_lower or 'flex' in est_lower:
                        return 'Conciliado'
                    elif 'reclasificación' in est_lower and 'múltiples' not in est_lower:
                        return 'Reclasificacion de Banco'
                    elif 'diferencia de fecha' in est_lower or 'grupo salmon' in est_lower:
                        return 'Diferencia de fecha'
                    elif 'diferencia de valor' in est_lower or 'regla 7b' in est_lower:
                        return 'Diferencia de valor'
                    else:
                        return 'Pendiente o solicitar soporte'

                def simplificar_comentario(txt):
                    txt_lower = str(txt).lower()
                    if not txt_lower or txt_lower == 'nan': return ""
                    
                    if 'sin coincidencia' in txt_lower:
                        if 'pdv' in txt_lower or 'ip' in txt_lower:
                            return "Sin coincidencia (Falta Referencia POS)"
                        return "Sin coincidencia"
                    if 'bloquead' in txt_lower or 'fuera de rango' in txt_lower:
                        return "Excede límite de días permitidos"
                    if 'nequi' in txt_lower:
                        if 'fifo desambiguado' in txt_lower: return "Cruce Nequi válido (FIFO)"
                        if 'no cuadra' in txt_lower or 'ambigua' in txt_lower or 'múltiples' in txt_lower or 'varios' in txt_lower or 'candidatos' in txt_lower:
                            return "Revisar Nequi (Ambigüedad o Totales)"
                        return "Cruce Nequi válido"
                    if 'reclasificación' in txt_lower or 'regla 6' in txt_lower:
                        return "Registrado en otro banco"
                    if 'diferencia de fecha' in txt_lower or 'diferencia f=' in txt_lower or 't3' in txt_lower:
                        return "Diferencia de fecha"
                    if 'diferencia de valor' in txt_lower or 'regla 7b' in txt_lower or 'dif=$' in txt_lower or 't4' in txt_lower:
                        return "Diferencia de valor"
                    if 'gasto' in txt_lower or 'comisión' in txt_lower or 't6' in txt_lower:
                        return "Posible gasto bancario"
                    if 'desbalanceado' in txt_lower:
                        return "Descuadre por sector"
                    if 'múltiples posiciones' in txt_lower or 'varias posiciones' in txt_lower:
                        return "Varias posiciones sin cruzar"
                    if 'ip/cb' in txt_lower or 'homologad' in txt_lower or 'pdv' in txt_lower or 't1' in txt_lower:
                        return "Cruce Punto de Venta (POS)"
                    if 'flex' in txt_lower or 'parcial' in txt_lower or 't2' in txt_lower:
                        return "Cruce por referencia parcial"
                    if 'fifo' in txt_lower:
                        return "Cruce por orden FIFO"
                    if 'tarde' in txt_lower or 't5' in txt_lower:
                        return "Cruce forzado (Modo Rescate)"
                    if 'cumple todas' in txt_lower or 'exacto' in txt_lower:
                        return "Cruce exacto"
                    
                    return "Revisión manual requerida"

                # Aplicamos la traducción a las columnas visibles
                df_final['Estado_Conciliacion'] = df_final['Estado_Tecnico'].apply(simplificar_estado)
                df_final['Comentario'] = df_final['Comentario_Tecnico'].apply(simplificar_comentario)

                columnas_visibles = columnas_originales + [
                    'Estado_Conciliacion', 'Comentario', 'Candidatos_Conciliacion', 'Sector'
                ]

                def vista(df_cualquiera):
                    cols = [c for c in columnas_visibles if c in df_cualquiera.columns]
                    return df_cualquiera[cols].copy()

                cuadre_ok = filas_antes == (len(df) + len(filas_descartadas))
                for c in [c for c in df_final.columns if 'fe.' in c.lower() or 'fecha' in c.lower() or 'fe-' in c.lower()]:
                    df_final[c] = pd.to_datetime(df_final[c], errors='coerce').dt.strftime('%d/%m/%Y')

                def color_fila(row):
                    idx = row.name
                    est = str(df_final.loc[idx, 'Estado_Tecnico']).strip().lower()

                    # Prioridad 1: Errores / Bloqueos (Blanco con texto rojo)
                    if 'fecha fuera de rango' in est: 
                        return [f'background-color: {COLOR_BLANCO}; color: red'] * len(row)
                        
                    # Prioridad 2: Multi-posiciones DZ sin cruce (Verde)
                    if 'dz posiciones múltiples' in est or 'dz multiposición sin cruce' in est:
                        return [f'background-color: {COLOR_VERDE}; color: black'] * len(row)
                        
                    # Prioridad 3: Conciliado perfecto o parcial flex (Azul)
                    if 'conciliado' in est or 'grupo azul' in est or 'flex' in est: 
                        return [f'background-color: {COLOR_AZUL}; color: black'] * len(row)
                        
                    # Prioridad 4: Cruce múltiple IP/CB (Gris)
                    if 'cruce múltiple ip/cb' in est: 
                        return [f'background-color: {COLOR_GRIS}; color: black'] * len(row)
                        
                    # Prioridad 5: Reclasificación (Durazno)
                    if 'reclasificación' in est: 
                        return [f'background-color: {COLOR_DURAZNO}; color: black'] * len(row)
                        
                    # Prioridad 6: Diferencia de Fecha (Salmón)
                    if 'diferencia de fecha' in est or 'grupo salmon' in est: 
                        return [f'background-color: {COLOR_SALMON}; color: black'] * len(row)
                        
                    # Prioridad 7: Diferencia de Valor (Morado / Blanco)
                    if 'diferencia de valor' in est:
                        if 'regla 7b' in est: 
                            return [f'background-color: {COLOR_BLANCO}; color: black'] * len(row)
                        return [f'background-color: {COLOR_MORADO}; color: black'] * len(row)
                        
                    # Prioridad 8: Sugerencias Tarde (Amarillo)
                    if '(tarde)' in est or 'tarde t' in est: 
                        return [f'background-color: {COLOR_AMARILLO}; color: black'] * len(row)

                    # Por defecto Blanco
                    return [f'background-color: {COLOR_BLANCO}; color: black'] * len(row)

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
                    
                    total_azul = int(df_final['Estado_Tecnico'].str.contains('Conciliado|grupo azul|Flex', na=False, regex=True).sum())
                    mask_gris = df_final['Estado_Tecnico'].str.contains('cruce múltiple ip/cb', case=False, na=False)
                    mask_durazno = df_final['Estado_Tecnico'].str.contains('reclasificación', case=False, na=False)
                    mask_salmon = df_final['Estado_Tecnico'].str.contains('grupo salmon|diferencia de fecha', case=False, na=False)
                    mask_morado = df_final['Estado_Tecnico'].str.contains('diferencia de valor', case=False, na=False) & ~df_final['Estado_Tecnico'].str.contains('regla 7b', case=False, na=False)
                    mask_azul = df_final['Estado_Tecnico'].str.contains('conciliado|grupo azul|flex', case=False, na=False)
                    mask_fuera_rango = df_final['Estado_Tecnico'].str.contains('fecha fuera de rango', case=False, na=False)
                    mask_amarillo = df_final['Estado_Tecnico'].str.contains('tarde', case=False, na=False)

                    mask_verde = (df_final[col_G] == '40') & df_final['Estado_Tecnico'].str.contains(
                        'dz posiciones múltiples|dz multiposición sin cruce', case=False, na=False, regex=True
                    )
                    total_verde = int(mask_verde.sum())

                    total_salmon = int(mask_salmon.sum())
                    total_morado = int(mask_morado.sum())
                    total_durazno = int(mask_durazno.sum())
                    total_pendiente = int(df_final['Estado_Tecnico'].str.contains('Pendiente', na=False).sum())
                    total_nequi_fifo = int(df_final['Comentario_Tecnico'].str.contains('FIFO desambiguado', case=False, na=False).sum())

                    resumen = pd.DataFrame({
                        "Métrica": [
                            "Fecha de procesamiento", "Total filas procesadas",
                            "Azul - Conciliados Exactos y Flex",
                            "Verde - Documentos DZ multiposición sin conciliar",
                            "Salmón - Diferencia de fecha (Regla 7)",
                            "Morado - Diferencia de valor máx $500",
                            "Durazno - Reclasificación de banco (Regla 6)",
                            "Amarillo - Sugerencias Modo Tarde",
                            "Blanco - Pendientes / Otras Sugerencias",
                            "IP conciliado exacto (Regla 3)", "IP con % de diferencia",
                            "Regla 8 Nequi - Azul (total y FIFO exacto)",
                            "Regla 8 Nequi - Sugerencia (grupo no cuadra exacto)",
                            "Nequi conciliado por FIFO desambiguado (v36)",
                            "FIFO controlado (última instancia)", "DZ verde sin cruce",
                            "Filas excluidas (sin doc/clave)", "Filas con Nº doc. repetido",
                            "Líneas marcadas Nequi (total)",
                        ],
                        "Valor": [
                            datetime.now().strftime('%d/%m/%Y %H:%M'), total_filas,
                            total_azul, total_verde, total_salmon, total_morado, total_durazno,
                            int(mask_amarillo.sum()), total_pendiente,
                            len(ind_ip_exacto), len(ind_ip_tolerancia),
                            len(ind_nequi8_azul), len(ind_nequi8_sugerencia),
                            total_nequi_fifo,
                            len(ind_fifo_ok), len(ind_fifo_verde_dz),
                            filas_excluidas, int(df_final['B_Repite'].sum()) if 'B_Repite' in df_final.columns else 0,
                            int(df['Es_Nequi'].sum()) if 'Es_Nequi' in df.columns else 0,
                        ]
                    })
                    resumen.to_excel(writer, index=False, sheet_name='RESUMEN')
                    pestanas_usadas.add('RESUMEN')

                    df_nov = df_final[df_final[col_G] == '40'].copy()
                    patron_alerta = 'Diferencia de fecha|Diferencia de valor|Reclasificación|grupo salmon|Sugerencia'
                    
                    mask_alerta = df_nov['Estado_Tecnico'].str.contains(patron_alerta, na=False, regex=True)
                    mask_sin_candidato = df_nov['Candidatos_Conciliacion'].astype(str).str.strip().isin(['', 'nan', 'None'])

                    df_nov = df_nov[mask_alerta | mask_sin_candidato]
                    if not df_nov.empty:
                        df_nov = df_nov.sort_values(by=['Estado_Tecnico', col_I])
                        hoja_segura(writer, vista(df_nov), 'NOVEDADES_Y_PENDIENTES_40', estilo=True)
                    else:
                        hoja_segura(writer, pd.DataFrame(columns=columnas_visibles), 'NOVEDADES_Y_PENDIENTES_40', estilo=False)

                    if modo_tarde:
                        df_tarde = df_final[df_final['Estado_Tecnico'].str.contains('tarde|Pendiente', case=False, na=False)].copy()
                        if not df_tarde.empty:
                            df_tarde['Abs_I'] = pd.to_numeric(df_tarde[col_I], errors='coerce').fillna(0).abs()
                            df_tarde = df_tarde.sort_values(by=['Abs_I', col_F])
                            hoja_segura(writer, vista(df_tarde), 'REVISION_TARDE_MENOR_A_MAYOR', estilo=True)

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

                # =====================================================
                # INTERFAZ
                # =====================================================
                st.success("¡Conciliación completada con el motor de reglas CLM v38 (Integral + Nequi FIFO + IP Estricto)!")
                if not cuadre_ok:
                    st.warning("⚠️ Revisa la pestaña DESCARTADAS, el total de filas no coincide.")
                for adv in advertencias:
                    st.warning(f"⚠️ {adv}")

                st.markdown(f'''
**Leyenda de colores:**
- <span style="background-color:{COLOR_AZUL}; padding:2px 8px;">Azul: Conciliado — cumple todas las reglas (A=H, misma F, mismo banco, importe exacto, sector coherente)</span>
- <span style="background-color:{COLOR_VERDE}; padding:2px 8px;">Verde: Documentos DZ (clv=40) con múltiples posiciones que no lograron conciliar.</span>
- <span style="background-color:{COLOR_SALMON}; padding:2px 8px;">Salmón: Diferencia de fecha F (hasta {TOPE_DIAS_ALERTA} días)</span>
- <span style="background-color:{COLOR_MORADO}; padding:2px 8px;">Morado: Diferencia de valor (máx ${tol_valor_purpura:.0f})</span>
- <span style="background-color:{COLOR_DURAZNO}; padding:2px 8px;">Durazno: Reclasificación de banco</span>
- <span style="background-color:{COLOR_AMARILLO}; padding:2px 8px;">Amarillo: Sugerencias del Modo Tarde (Depuración Profunda)</span>
- <span style="background-color:{COLOR_GRIS}; padding:2px 8px;">Gris: Cruces múltiples IP/CB (Regla 3)</span>
- <span style="background-color:{COLOR_BLANCO}; padding:2px 8px; border:1px solid #ccc;">Blanco: Pendientes / Otras Sugerencias / Bloqueos por cruces fuera del límite de días</span>

**Novedades v38 (Integral + Nequi FIFO + IP Estricto):**
- **Nequi Dinámico:** Reconoce Nequi por rango numérico (Créditos) o por palabras clave / códigos "T", "/", "T-" (Legalizaciones), tolerando errores de escritura.
- **Desambiguación FIFO en Nequi:** Resuelve automáticamente cruces cuando hay múltiples documentos Nequi del mismo importe y fecha.
- **Estados y Comentarios simplificados:** La exportación usa frases cortas ("Registrado en otro banco", "Diferencia de fecha") para revisión rápida.
- Bugfixes de limpieza de texto integrados.
''', unsafe_allow_html=True)

                c1, c2, c3, c4, c5, c6 = st.columns(6)
                c1.metric("Azul", total_azul)
                c2.metric("Verde", total_verde)
                c3.metric("Salmón", total_salmon)
                c4.metric("Morado", total_morado)
                c5.metric("Durazno", total_durazno)
                c6.metric("Pendiente", total_pendiente)

                if filas_excluidas > 0:
                    st.warning(f"⚠️ Se excluyeron {filas_excluidas} filas vacías/totales.")

                st.download_button(
                    label="📥 Descargar Excel con Resultados",
                    data=output.getvalue(),
                    file_name="Conciliacion_CLM_v38_Integral_IP_Estricto.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Error técnico detectado: {e}")
            st.exception(e)
