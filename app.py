# app_conciliacion_v26_mejoras.py
#
# REESCRITURA COMPLETA DEL MOTOR segun especificacion CLM entregada por
# el usuario. Incluye optimizaciones de UI (Botón de ejecución) 
# y resolución de bugs (Filtro Novedades, dayfirst en fechas, tolerancia IP/CB).

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

st.title("🏦 Conciliación Automatizada — Motor CLM v26 🤖")
st.write("Sube tu archivo consolidado.")
st.caption(
    "Motor reescrito según especificación CLM (Incluye Parche Estructural v33 con límite de 4 días y Fix Novedades). "
    "A=Asignación, B=Nº doc, C=Clase doc, D=Fecha periodo, F=Fecha valor (PRINCIPAL), "
    "G=Clave, H=Referencia, I=Importe, K=Texto."
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
        min_value=0.00, value=0.00, step=0.05 # FIX: Por defecto 0.0% para exigir importes exactos
    ) / 100
    multiplo_redondo = st.selectbox("Múltiplo para valor 'redondo' (alta ambigüedad)", [50000, 100000], index=1)

COLOR_AZUL = "#C5D9F1"      # Conciliado (todas las reglas cumplen)
COLOR_VERDE = "#A9D18E"     # Sugerencia / DZ multiposición / sectorización / IP%
COLOR_SALMON = "#F5B7A1"    # Diferencia de fecha (hasta 4 días) - Regla 7
COLOR_MORADO = "#C39BD3"    # Diferencia de valor máx $500 - Regla morado
COLOR_DURAZNO = "#FAD7A0"   # Reclasificación de banco - Regla 6
COLOR_BLANCO = "#FFFFFF"    # Pendiente
COLOR_GRIS = "#D0CECE"      # Cruces múltiples de documento IP (Regla 3)

archivo_subido = st.file_uploader("Selecciona el archivo de Excel o CSV", type=['xlsx', 'csv'])

if archivo_subido is not None:
    # NUEVO: Botón para evitar la recarga automática al tocar un parámetro
    if st.button("🚀 Ejecutar Conciliación", use_container_width=True):
        try:
            with st.spinner("Ejecutando motor de reglas CLM... Esto puede tomar unos segundos."):

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
                # FIX: Fallback tipográfico a 'Asignacion' (sin tilde) para lanzar el error adecuado si no existe
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

                # FIX: Añadido dayfirst=True para evitar cruces dd/mm vs mm/dd
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

                # Posiciones repetidas del mismo B (Regla 4)
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
                mapeo_datafono_ref = {
                    "11760923": "3001", "11761277": "3002", "11761293": "3003", "11761327": "3004",
                    "11761301": "3005", "12273934": "3006", "11761319": "3007", "12273900": "3008",
                    "12273926": "3009", "14632012": "3010", "15186547": "3011", "13048756": "3012",
                    "15186539": "3013", "16219602": "3200", "16591240": "3201", "16634586": "3202",
                    "14885164": "2005", "19827765": "3203", "11761350": "2001", "12161154": "2002",
                    "14294946": "2003", "15926645": "2210", "11831583": "4002", "12161162": "4001",
                    "12161121": "4003", "12161139": "4004", "12874475": "4005", "15190309": "4006",
                    "14468144": "4006", "12500773": "4008", "14468151": "4009", "14651459": "4010",
                    "15444946": "4200", "16062176": "4253", "20836698": "4007", "72806854": "4203",
                    "20719829": "4201", "15536188": "6101", "12637294": "6102", "11844685": "6103",
                    "15536170": "6106", "17549197": "6108"
                }

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
                    nums = re.findall(r'\b\d{4}\b', t_full)
                    for n in nums:
                        num = int(n)
                        if 2000 <= num <= 2999: return 'Dist Buga'
                        if 3000 <= num <= 3999: return 'Dist Acopi'
                        if 4000 <= num <= 4999: return 'Dist Dosquebradas'
                        if 6000 <= num <= 6999: return 'Dist Pasto'

                    return 'Sin clasificar'

                df['Sector'] = df.apply(clasificar_sector, axis=1)

                def obtener_ref_homologada(row):
                    texto = f"{row.get(col_H,'')} {row.get(col_A,'')} {row.get(col_K,'') if col_K else ''} {row.get(col_novedad,'') if col_novedad else ''}".upper()
                    n8 = re.findall(r'\b\d{8}\b', texto)
                    for n in n8:
                        if n in mapeo_datafono_ref: return mapeo_datafono_ref[n]
                    n4 = re.findall(r'\b\d{4}\b', texto)
                    for n in n4:
                        if n in mapeo_datafono_ref.values(): return n
                    return None

                def es_nequi(row):
                    texto = f"{row.get(col_K,'') if col_K else ''} {row.get(col_A,'')} {row.get(col_H,'')}".upper()
                    if 'NEQUI' in texto: return True
                    val_a = str(row.get(col_A, '')).strip().upper()
                    if val_a == 'T' or val_a.startswith('T-') or val_a.startswith('T/') or val_a == '/': return True
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
                # FUNCIONES AUXILIARES & PARCHE V33
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

                def gate_seguridad(id40, id50, exigir_importe_exacto=True, tolerancia_valor=None):
                    res = df.loc[df['ID_Linea'].isin([id40, id50])]
                    ra = res[res['ID_Linea'] == id40].iloc[0]
                    rb = res[res['ID_Linea'] == id50].iloc[0]

                    resultado = {
                        'ok': False, 'motivo': '', 'dif_dias': None,
                        'mismo_banco': False, 'banco_a': '', 'banco_b': '',
                        'dif_valor': None, 'pct_valor': None,
                        'mismo_sector': True, 'es_nequi': False
                    }

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
                    resultado['mismo_banco'] = (banco_a == banco_b)

                    sector_a = str(ra.get('Sector', '')).strip()
                    sector_b = str(rb.get('Sector', '')).strip()
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

                def clasificar_y_registrar(id40, id50, base_txt):
                    res = gate_seguridad(id40, id50, exigir_importe_exacto=True, tolerancia_valor=tol_valor_purpura)
                    if not res['ok']:
                        return False, res['motivo']

                    texto_candidatos = f"{formato_linea(id40)} | {formato_linea(id50)}"
                    partes_comentario = [base_txt]
                    estado_final = 'Conciliado - Cumple todas las reglas'

                    if not res['mismo_banco']:
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
                # Regla 3: IP Homologados Grouped
                # =====================================================
                ind_ip_exacto = set()
                ind_ip_tolerancia = set()
                if usar_ipcb:
                    df['Ref_H_Homologada'] = df.apply(obtener_ref_homologada, axis=1)

                    df_ip = df[(df[col_C].astype(str).str.upper() == 'IP') & (df[col_G] == '40') & df['Ref_H_Homologada'].notna()]
                    df_cb = df[(df[col_C].astype(str).str.upper() == 'CB') & (df[col_G] == '50') & df['Ref_H_Homologada'].notna()]

                    if not df_ip.empty and not df_cb.empty:
                        grp_ip = df_ip.groupby([col_banco, col_F, 'Ref_H_Homologada'])['Abs_I'].sum().reset_index(name='S_IP')
                        grp_cb = df_cb.groupby([col_banco, col_F, 'Ref_H_Homologada'])['Abs_I'].sum().reset_index(name='S_CB')
                        m = pd.merge(grp_cb, grp_ip, on=[col_banco, col_F, 'Ref_H_Homologada'])
                        m['DifV'] = (m['S_CB'] - m['S_IP']).abs()
                        max_s = m[['S_CB', 'S_IP']].max(axis=1).clip(lower=1)
                        m['Pct'] = m['DifV'] / max_s

                        exactos = m[m['DifV'].round(2) == 0]
                        con_tol = m[(m['DifV'] > 0) & ((m['DifV'] <= tol_valor_abs_general) | (m['Pct'] <= tol_valor_pct_general))]

                        def procesar_grupo_ip(fila, es_exacto):
                            b, f_val, rh = fila[col_banco], fila[col_F], fila['Ref_H_Homologada']
                            sub_ip = df_ip[(df_ip[col_banco] == b) & (df_ip[col_F] == f_val) & (df_ip['Ref_H_Homologada'] == rh)]
                            sub_cb = df_cb[(df_cb[col_banco] == b) & (df_cb[col_F] == f_val) & (df_cb['Ref_H_Homologada'] == rh)]
                            ip_ids = [i for i in sub_ip['ID_Linea'].tolist() if i not in usados]
                            cb_ids = [i for i in sub_cb['ID_Linea'].tolist() if i not in usados]
                            if not ip_ids or not cb_ids: return
                            usados.update(ip_ids + cb_ids)
                            texto_cand = " | ".join(formato_linea(i) for i in ip_ids + cb_ids)
                            if es_exacto:
                                estado = 'Conciliado - Cruce múltiple IP/CB (Regla 3)'
                                ind_ip_exacto.update(ip_ids + cb_ids)
                                txt = f"Cruce múltiple homologado ({len(ip_ids)} IP = {len(cb_ids)} CB), misma Fecha valor ({f_val}). Ref. homologada: {rh}."
                            else:
                                estado = 'Sugerencia - Cruce múltiple IP/CB con diferencia de valor'
                                ind_ip_tolerancia.update(ip_ids + cb_ids)
                                txt = f"Sugerencia IP/CB con diferencia de valor (${fila['DifV']:,.0f} / {fila['Pct']*100:.2f}%), misma Fecha valor ({f_val}). Suma {len(ip_ids)} IP vs {len(cb_ids)} CB. Ref. homologada: {rh}."
                            for idx in ip_ids + cb_ids:
                                df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = estado
                                df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                                otros = resumen_docs(sub_cb) if idx in ip_ids else resumen_docs(sub_ip)
                                df.loc[df['ID_Linea'] == idx, 'Comentario'] = f"{txt} Docs relacionados: {otros}"

                        for _, fila in exactos.iterrows(): procesar_grupo_ip(fila, es_exacto=True)
                        for _, fila in con_tol.iterrows(): procesar_grupo_ip(fila, es_exacto=False)

                # ================================================================
                # 3. PARCHE v33: IP 1 A 1 POR ZONA, VALOR Y FECHA 
                # ================================================================
                if usar_ipcb:
                    ip40_pend = df[(df[col_C].astype(str).str.upper() == 'IP') & (df[col_G] == '40') & (~df['ID_Linea'].isin(usados))].copy()
                    cb50_pend = df[(df[col_C].astype(str).str.upper() == 'CB') & (df[col_G] == '50') & (~df['ID_Linea'].isin(usados))].copy()

                    for id_ip, fila_ip in ip40_pend.iterrows():
                        if id_ip in usados: continue
                        candidatos = cb50_pend[
                            (cb50_pend[col_banco] == fila_ip[col_banco]) &
                            (cb50_pend['Sector'] == fila_ip['Sector']) &
                            (cb50_pend['Abs_I'] == fila_ip['Abs_I']) &
                            (~cb50_pend['ID_Linea'].isin(usados))
                        ].copy()

                        if candidatos.empty: continue

                        candidatos['_dif_dias'] = (candidatos['Fecha_F'] - fila_ip['Fecha_F']).dt.days.abs().fillna(999)
                        candidatos = candidatos[candidatos['_dif_dias'] <= TOPE_DIAS_ALERTA]

                        if len(candidatos) != 1: continue

                        id_cb = candidatos.iloc[0]['ID_Linea']
                        registrar_pareja_por_fecha(id_ip, id_cb, f'IP/CB 1 a 1 por misma zona {fila_ip["Sector"]}, banco e importe exacto.')

                    ip_sin_resolver = df[(df[col_C].astype(str).str.upper() == 'IP') & (~df['ID_Linea'].isin(usados))]
                    for idl in ip_sin_resolver['ID_Linea']:
                        escribir_comentario(idl, "PDV (IP): requiere referencia homologada de base de datos o coincidencia por Zona.", append=False)

                # ================================================================
                # ALERTA DE RECLASIFICACIÓN DE BANCO EXCLUSIVA IP (Hasta 4 días)
                # ================================================================
                if usar_ipcb:
                    ip40_recl = df[(df[col_C].astype(str).str.upper() == 'IP') & (df[col_G] == '40') & (~df['ID_Linea'].isin(usados))].copy()
                    cb50_recl = df[(df[col_C].astype(str).str.upper() == 'CB') & (df[col_G] == '50') & (~df['ID_Linea'].isin(usados))].copy()

                    for id_ip, fila_ip in ip40_recl.iterrows():
                        if id_ip in usados: continue

                        # Buscar CBs en DISTINTO banco, mismo importe
                        candidatos = cb50_recl[
                            (cb50_recl[col_banco] != fila_ip[col_banco]) &
                            (cb50_recl['Abs_I'] == fila_ip['Abs_I']) &
                            (~cb50_recl['ID_Linea'].isin(usados))
                        ].copy()

                        if candidatos.empty: continue

                        # Aplicar tolerancia de hasta 4 días
                        candidatos['_dif_dias'] = (candidatos['Fecha_F'] - fila_ip['Fecha_F']).dt.days.abs().fillna(999)
                        candidatos = candidatos[candidatos['_dif_dias'] <= TOPE_DIAS_ALERTA]

                        if candidatos.empty: continue

                        # Preferir candidatos de la misma zona si la zona no es 'Sin clasificar'
                        if fila_ip['Sector'] != 'Sin clasificar':
                            cand_zona = candidatos[candidatos['Sector'] == fila_ip['Sector']]
                            if not cand_zona.empty:
                                candidatos = cand_zona
                        
                        if len(candidatos) == 1:
                            id_cb = candidatos.iloc[0]['ID_Linea']
                            banco_ip = str(fila_ip[col_banco]).strip()
                            banco_cb = str(candidatos.iloc[0][col_banco]).strip()
                            dif_d = int(candidatos.iloc[0]['_dif_dias'])
                            
                            texto_cand = f"{formato_linea(id_ip)} | {formato_linea(id_cb)}"
                            comentario = f"Reclasificación de banco (Exclusiva IP): Registrado en '{banco_ip}', esperado en '{banco_cb}'. Diferencia F={dif_d} día(s). Mismo importe."
                            
                            for idx in [id_ip, id_cb]:
                                df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Reclasificación de banco'
                                df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                                df.loc[df['ID_Linea'] == idx, 'Comentario'] = comentario
                            usados.update([id_ip, id_cb])
                            parejas_registradas.append((id_ip, id_cb))
                        elif len(candidatos) > 1:
                            docs_cand = resumen_docs(candidatos)
                            df.loc[df['ID_Linea'] == id_ip, 'Estado_Conciliacion'] = 'Sugerencia - Reclasificación con múltiples candidatos'
                            df.loc[df['ID_Linea'] == id_ip, 'Candidatos_Conciliacion'] = f"{formato_linea(id_ip)} | Candidatos posibles: {docs_cand}"
                            df.loc[df['ID_Linea'] == id_ip, 'Comentario'] = f"Alerta de reclasificación IP: varios candidatos en otros bancos dentro de {TOPE_DIAS_ALERTA} días."


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
                # =====================================================
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
                # 4. PARCHE v33: SECTORIZACION MULTIPLE FIFO
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
                # 6. PARCHE v33: REGLA 7B - DIFERENCIA DE VALOR (Alertas Sugeridas)
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
                    posibles = posibles[posibles['_dif_valor'] > 500]
                    if posibles.empty: continue

                    id50 = posibles.sort_values(['_dif_valor', '_dif_dias']).iloc[0]['ID_Linea']
                    diferencia = round(abs(fila40['Abs_I'] - df.loc[id50, 'Abs_I']), 2)
                    texto = f'{formato_linea(id40)} | {formato_linea(id50)}'
                    comentario = (
                        f'Regla 7B: diferencia de valor ${diferencia:,.2f}; '
                        f'mismo banco/zona y fecha dentro del limite de {TOPE_DIAS_ALERTA} dias. '
                        'No es morado y no se concilia automaticamente.'
                    )
                    for idx in [id40, id50]:
                        df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Sugerencia - Regla 7B diferencia de valor'
                        df.loc[df['ID_Linea'] == idx, 'Comentario'] = comentario
                        df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto
                    usados.update([id40, id50])

                # =====================================================
                # 8. EXCEPCIÓN NEQUI (A = "Nequi", C = DZ)
                # =====================================================
                df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
                df_50 = df_50[~df_50['ID_Linea'].isin(usados)]

                df_nequi_40 = df_40[df_40['Es_Nequi'] == True]

                for _, r40 in df_nequi_40.iterrows():
                    id40 = r40['ID_Linea']
                    if id40 in usados: continue
                    candidatos_50 = df_50[
                        (df_50[col_banco] == r40[col_banco]) &
                        (~df_50['ID_Linea'].isin(usados))
                    ].copy()
                    if candidatos_50.empty: continue

                    if r40['Sector'] != 'Sin clasificar':
                        candidatos_filtrados = candidatos_50[(candidatos_50['Sector'] == r40['Sector']) | (candidatos_50['Sector'] == 'Sin clasificar')]
                        if not candidatos_filtrados.empty: candidatos_50 = candidatos_filtrados
                    
                    candidatos_50['_dif_dias'] = (candidatos_50['Fecha_F'] - r40['Fecha_F']).dt.days.abs().fillna(999)
                    
                    # PARCHE: REGLA DE 4 DIAS ESTRICTA PARA TODOS LOS NEQUI
                    candidatos_50 = candidatos_50[candidatos_50['_dif_dias'] <= TOPE_DIAS_ALERTA]
                    if candidatos_50.empty: continue

                    exactos = candidatos_50[candidatos_50['Abs_I'] == r40['Abs_I']]
                    if not exactos.empty:
                        id50 = exactos.iloc[0]['ID_Linea']
                        ok, _ = clasificar_y_registrar(id40, id50, "Excepción Nequi (cruce importe exacto)")
                        if ok: usados.update([id40, id50])
                        continue

                    candidatos_50['_dif_val'] = (candidatos_50['Abs_I'] - r40['Abs_I']).abs()
                    con_tol = candidatos_50[candidatos_50['_dif_val'] <= tol_valor_purpura].sort_values('_dif_val')
                    if not con_tol.empty:
                        id50 = con_tol.iloc[0]['ID_Linea']
                        ok, _ = clasificar_y_registrar(id40, id50, "Excepción Nequi (con diferencia de valor)")
                        if ok: usados.update([id40, id50])
                        continue

                # =====================================================
                # 9. REGLA 4 — DOCUMENTOS DZ CON POSICIONES MÚLTIPLES
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
                # 10. ÚLTIMO RECURSO: FIFO CONTROLADO
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
                # 10B. REGLA 6B — RECLASIFICACIÓN SIN REFERENCIA (último recurso)
                # ================================================================
                df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
                df_50 = df_50[~df_50['ID_Linea'].isin(usados)]
                pendientes_40b = df[(df['ID_Linea'].isin(df_40['ID_Linea'])) & (~df['ID_Linea'].isin(usados))]
                pendientes_50b = df[(df['ID_Linea'].isin(df_50['ID_Linea'])) & (~df['ID_Linea'].isin(usados))]

                for (fecha_z, importe_z), grupo40 in pendientes_40b.groupby([col_F, 'Abs_I']):
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
                # 7. PARCHE v33: VALIDACION FINAL DE FECHA (Muro de Contencion)
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
                # 11. CIERRE
                # =====================================================
                sin_p = df['Estado_Conciliacion'] == 'Pendiente'
                if usar_ipcb:
                    es_ip = df[col_C].astype(str).str.upper() == 'IP'
                    df.loc[sin_p & es_ip & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia - PDV (requiere referencia homologada o cruce exacto)'
                    df.loc[sin_p & ~es_ip & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia ni sugerencia que cumpla reglas de seguridad.'
                else:
                    df.loc[sin_p & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia ni sugerencia que cumpla reglas de seguridad.'

                # =====================================================
                # 12. EXPORTACIÓN
                # =====================================================
                columnas_visibles = columnas_originales + [
                    'Estado_Conciliacion', 'Comentario', 'Candidatos_Conciliacion', 'Sector'
                ]

                def vista(df_cualquiera):
                    cols = [c for c in columnas_visibles if c in df_cualquiera.columns]
                    return df_cualquiera[cols].copy()

                cuadre_ok = filas_antes == (len(df) + len(filas_descartadas))
                df_final = df.drop(columns=['ID_Linea', 'Abs_I', 'Fecha_F', 'Fecha_D'], errors='ignore')
                for c in [c for c in df_final.columns if 'fe.' in c.lower() or 'fecha' in c.lower() or 'fe-' in c.lower()]:
                    df_final[c] = pd.to_datetime(df_final[c], errors='coerce').dt.strftime('%d/%m/%Y')

                def color_fila(row):
                    est = str(row['Estado_Conciliacion']).strip().lower()
                    if est in ('pendiente', '', 'nan'): return [f'background-color: {COLOR_BLANCO}; color: black'] * len(row)
                    
                    if 'cruce múltiple ip/cb' in est: return [f'background-color: {COLOR_GRIS}; color: black'] * len(row)
                    
                    if 'reclasificación' in est: return [f'background-color: {COLOR_DURAZNO}; color: black'] * len(row)
                    
                    if 'grupo salmon' in est or 'diferencia de fecha' in est: return [f'background-color: {COLOR_SALMON}; color: black'] * len(row)
                    if 'diferencia de valor' in est:
                        if 'regla 7b' in est: return [f'background-color: {COLOR_VERDE}; color: black'] * len(row)
                        return [f'background-color: {COLOR_MORADO}; color: black'] * len(row)
                    
                    if 'dz multiposición' in est or 'dz posiciones múltiples' in est or 'sectorización desbalanceada' in est or 'sugerencia' in est: return [f'background-color: {COLOR_VERDE}; color: black'] * len(row)
                    if 'conciliado' in est or 'grupo azul' in est or 'flex' in est: return [f'background-color: {COLOR_AZUL}; color: black'] * len(row)
                    if 'fecha fuera de rango' in est: return [f'background-color: {COLOR_BLANCO}; color: red'] * len(row)
                    
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
                    nombre = re.sub(r'[\\/*?:\[\]]', '-', str(base)[:31])
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
                            df_hoja.style.apply(lambda row: color_fila(df_hoja.loc[row.name]), axis=1).to_excel(writer, index=False, sheet_name=nf)
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
                    total_azul = int(df_final['Estado_Conciliacion'].str.contains('Conciliado|grupo azul|Flex', na=False, regex=True).sum())
                    total_verde = int(
                        df_final['Estado_Conciliacion'].str.contains('Sugerencia|DZ multiposición|DZ posiciones múltiples|Sectorización desbalanceada|Regla 7B', na=False, regex=True).sum()
                    )
                    total_salmon = int(df_final['Estado_Conciliacion'].str.contains('Diferencia de fecha|grupo salmon', na=False, regex=True).sum())
                    total_morado = int(df_final['Estado_Conciliacion'].str.contains('Diferencia de valor', na=False).sum()) - int(df_final['Estado_Conciliacion'].str.contains('Regla 7B', na=False).sum())
                    total_durazno = int(df_final['Estado_Conciliacion'].str.contains('Reclasificación', na=False).sum())
                    total_pendiente = int(df_final['Estado_Conciliacion'].str.contains('Pendiente', na=False).sum())

                    resumen = pd.DataFrame({
                        "Métrica": [
                            "Fecha de procesamiento", "Total filas procesadas",
                            "Azul - Conciliados Exactos y Flex",
                            "Verde - Sugerencias de Revisión Manual",
                            "Salmón - Diferencia de fecha (Regla 7)",
                            "Morado - Diferencia de valor máx $500",
                            "Durazno - Reclasificación de banco (Regla 6)",
                            "Blanco - Pendientes / Bloqueados",
                            "IP conciliado exacto (Regla 3)", "IP con % de diferencia",
                            "Regla 8 Nequi - Azul (total y FIFO exacto)",
                            "Regla 8 Nequi - Sugerencia (grupo no cuadra exacto)",
                            "FIFO controlado (última instancia)", "DZ verde sin cruce",
                            "Filas excluidas (sin doc/clave)", "Filas con Nº doc. repetido",
                        ],
                        "Valor": [
                            datetime.now().strftime('%d/%m/%Y %H:%M'), total_filas,
                            total_azul, total_verde, total_salmon, total_morado, total_durazno, total_pendiente,
                            len(ind_ip_exacto), len(ind_ip_tolerancia),
                            len(ind_nequi8_azul), len(ind_nequi8_sugerencia),
                            len(ind_fifo_ok), len(ind_fifo_verde_dz),
                            filas_excluidas, int(df_final['B_Repite'].sum()) if 'B_Repite' in df_final.columns else 0,
                        ]
                    })
                    resumen.to_excel(writer, index=False, sheet_name='RESUMEN')
                    pestanas_usadas.add('RESUMEN')

                    # FIX BUG PRINCIPAL: Filtro de novedades con Regex para atrapar todas las sugerencias compuestas
                    df_nov = df_final[df_final[col_G] == '40'].copy()
                    patron_alerta = 'Diferencia de fecha|Diferencia de valor|Reclasificación|grupo salmon|Sugerencia'
                    mask_alerta = df_nov['Estado_Conciliacion'].str.contains(patron_alerta, na=False, regex=True)
                    mask_sin_candidato = df_nov['Candidatos_Conciliacion'].astype(str).str.strip().isin(['', 'nan', 'None'])
                    
                    df_nov = df_nov[mask_alerta | mask_sin_candidato]
                    if not df_nov.empty:
                        df_nov = df_nov.sort_values(by=['Estado_Conciliacion', col_I])
                        hoja_segura(writer, vista(df_nov), 'NOVEDADES_Y_PENDIENTES_40', estilo=True)
                    else:
                        hoja_segura(writer, pd.DataFrame(columns=columnas_visibles), 'NOVEDADES_Y_PENDIENTES_40', estilo=False)

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
                # 13. INTERFAZ
                # =====================================================
                st.success("¡Conciliación completada con el motor de reglas CLM v26!")
                if not cuadre_ok:
                    st.warning("⚠️ Revisa la pestaña DESCARTADAS, el total de filas no coincide.")
                for adv in advertencias:
                    st.warning(f"⚠️ {adv}")

                st.markdown(f'''
**Leyenda de colores:**
- <span style="background-color:{COLOR_AZUL}; padding:2px 8px;">Azul: Conciliado — cumple todas las reglas (A=H, misma F, mismo banco, importe exacto, sector coherente)</span>
- <span style="background-color:{COLOR_VERDE}; padding:2px 8px;">Verde: Sugerencias — revisión de sectorización, Regla 7B (diferencias > 500) y DZ multiposición</span>
- <span style="background-color:{COLOR_SALMON}; padding:2px 8px;">Salmón: Diferencia de fecha F (hasta {TOPE_DIAS_ALERTA} días)</span>
- <span style="background-color:{COLOR_MORADO}; padding:2px 8px;">Morado: Diferencia de valor (máx ${tol_valor_purpura:.0f})</span>
- <span style="background-color:{COLOR_DURAZNO}; padding:2px 8px;">Durazno: Reclasificación de banco</span>
- <span style="background-color:{COLOR_GRIS}; padding:2px 8px;">Gris: Cruces múltiples IP/CB (Regla 3)</span>
- <span style="background-color:{COLOR_BLANCO}; padding:2px 8px; border:1px solid #ccc;">Blanco: Pendientes / Bloqueos por cruces fuera del límite de días</span>

**Reglas clave aplicadas:**
- La fecha D (periodo) es SOLO informativa. La fecha F es la ÚNICA que valida conciliación.
- Los IP cruzan por referencia homologada agrupada, o 1 a 1 por Sector y Valor exactos (Parche v33).
- **Muro de Contención Contable:** Salvo la "Flex Referencia Parcial", NINGÚN cruce superará el tope estricto de {TOPE_DIAS_ALERTA} días de diferencia.
- Los Nequi exigen candidato en mismo banco/sector y aplican la tolerancia morada o estricta de días.
- Regla 7B: Diferencias altas de valor (> $500) en el mismo sector/fecha quedarán sugeridas en VERDE.
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
                    file_name="Conciliacion_CLM_v26_Mejoras_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Error técnico detectado: {e}")
            st.exception(e)
