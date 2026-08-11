# app_conciliacion_v24_color_verde_exclusivo.py
#
# Basado en la v24 original (Correccion de Periodo y Bug de Sobrescritura),
# aplicando una REGLA VISUAL ESTRICTA: El color Verde (#A9D18E) solo pinta
# filas que sean Clave 40 + Clase DZ + No conciliadas. El resto de las
# sugerencias asumen color Amarillo (#FFF2CC).

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime

# ============================================================
# CONFIGURACIÓN DE LA PÁGINA
# ============================================================
st.set_page_config(page_title="Conciliación Integral", layout="wide")

hide_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
"""
st.markdown(hide_style, unsafe_allow_html=True)

st.title("🏦 Conciliación Automatizada 🤖")
st.write("Sube tu archivo consolidado.")
st.caption(
    "v24 (Color Actualizado): Regla visual donde el color Verde aplica ÚNICAMENTE "
    "a los documentos Clave 40, Clase DZ que no lograron conciliar."
)

with st.expander("⚙️ Parámetros de tolerancia para sugerencias (alertas)"):
    tol_dias = st.slider(
        "Días máximos de diferencia para alertar 'error de fecha' (mismo periodo, tope 9)",
        1, 9, 3
    )
    tol_valor_abs = st.number_input("Diferencia absoluta máxima de valor para alertar ($)", min_value=1, value=5000, step=100)
    tol_valor_pct = st.number_input("Diferencia relativa máxima de valor para alertar (%)", min_value=0.01, value=0.5, step=0.01) / 100
    multiplo_redondo = st.selectbox("Múltiplo para considerar un valor 'redondo' (alta ambigüedad)", [50000, 100000], index=1)
    activar_sector_redondos = st.checkbox(
        "Activar 'Sugerencia por Sectorización en Valores Redondos' (banco+importe+periodo+Distribuidora; "
        "NUNCA cruza sectores distintos, requiere que ambos lados tengan sectorización clasificada)",
        value=True
    )

TOPE_DIAS_ALERTA = 9
tol_dias = min(tol_dias, TOPE_DIAS_ALERTA)

COLOR_CONCILIADO = "#C5D9F1"
COLOR_VERIFICAR = "#A9D18E"  # VERDE EXCLUSIVO 40 DZ
COLOR_AMARILLO = "#FFF2CC"   # AMARILLO RESTO DE SUGERENCIAS
COLOR_ALERTA = "#FDEBD0"
COLOR_RECLASIFICAR = "#D7BDE2"
COLOR_PENDIENTE = "#FFFFFF"

archivo_subido = st.file_uploader("Selecciona el archivo de Excel o CSV", type=['xlsx', 'csv'])

if archivo_subido is not None:
    try:
        with st.spinner("Ejecutando motor de reglas, M:N, clasificación multibanco y validación de periodo..."):

            # =========================================================
            # 1. LECTURA Y MAPEO DE COLUMNAS
            # =========================================================
            if archivo_subido.name.lower().endswith('.csv'):
                df = pd.read_csv(archivo_subido)
            else:
                diccionario_hojas = pd.read_excel(archivo_subido, sheet_name=None)
                hojas_validas = [h for h in diccionario_hojas.values() if not h.dropna(how='all').empty]
                if not hojas_validas:
                    st.error("El archivo no contiene hojas con datos.")
                    st.stop()
                df = pd.concat(hojas_validas, ignore_index=True)

            df.columns = df.columns.str.strip()

            col_asignacion = 'Asignación' if 'Asignación' in df.columns else 'Asignaión'
            col_referencia = 'Referencia'
            col_clave = 'Clave contabiliz.' if 'Clave contabiliz.' in df.columns else 'CT'
            col_fecha = 'Fecha valor' if 'Fecha valor' in df.columns else 'Fe-valor'
            col_fecha_contable = 'Fe.contabilización' if 'Fe.contabilización' in df.columns else (
                'Fecha de documento' if 'Fecha de documento' in df.columns else col_fecha
            )
            col_importe = 'Importe en moneda local' if 'Importe en moneda local' in df.columns else 'Importe en ML'
            col_banco = 'Clave referencia 3'
            col_doc = 'Nº documento' if 'Nº documento' in df.columns else 'Nº doc.'
            col_texto = 'Texto' if 'Texto' in df.columns else None
            col_novedad = 'novedad' if 'novedad' in df.columns else ('Novedad' if 'Novedad' in df.columns else None)
            col_clase_doc = 'Clase de documento' if 'Clase de documento' in df.columns else ('Clase doc.' if 'Clase doc.' in df.columns else None)

            columnas_requeridas = [col_referencia, col_clave, col_fecha, col_importe, col_banco, col_doc]
            faltantes = [c for c in columnas_requeridas if c not in df.columns]
            if faltantes:
                st.error(f"No se encontraron estas columnas obligatorias: {faltantes}")
                st.stop()
            if col_asignacion not in df.columns:
                st.error("No se encontró la columna de Asignación en el archivo.")
                st.stop()

            usar_ipcb = col_clase_doc is not None
            columnas_originales = list(df.columns)

            # =========================================================
            # 2. AUTOCOMPLETADO DE BANCOS (Cuentas de Mayor)
            # =========================================================
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
                asig_val = str(row.get(col_asignacion, ""))
                banco_val = row.get(col_banco, None)

                if "cuenta de mayor" in asig_val.lower():
                    match_cuenta = re.search(r'(\d{6,})', asig_val)
                    if match_cuenta:
                        cuenta_num = match_cuenta.group(1)
                        current_bank = mapeo_cuentas_banco.get(cuenta_num, f"CUENTA {cuenta_num} (sin mapear)")

                if pd.notnull(banco_val) and str(banco_val).strip().lower() not in ("", "nan"):
                    current_bank = str(banco_val).strip()
                bancos_completados.append(current_bank)

            df[col_banco] = bancos_completados
            df = df[~df[col_asignacion].astype(str).str.contains("cuenta de mayor", case=False, na=False)].copy()

            # =========================================================
            # 3. LIMPIEZA BASE Y ORDENAMIENTO
            # =========================================================
            df[col_doc] = pd.to_numeric(df[col_doc], errors='coerce')
            filas_antes = len(df)
            filas_descartadas = df[df[col_doc].isna() | df[col_clave].isna()].copy()
            df = df.dropna(subset=[col_doc, col_clave]).reset_index(drop=True)
            filas_excluidas = filas_antes - len(df)

            df = df.sort_values(by=[col_doc], ascending=True).reset_index(drop=True)
            df['ID_Temp'] = df.index

            df[col_clave] = df[col_clave].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df[col_banco] = df[col_banco].astype(str).str.strip()
            df[col_importe] = pd.to_numeric(df[col_importe], errors='coerce').fillna(0)
            df['Abs_Importe'] = df[col_importe].abs()

            df['Fecha_Calc'] = pd.to_datetime(df[col_fecha], errors='coerce')
            df[col_fecha] = df['Fecha_Calc'].dt.date

            df['Estado_Conciliacion'] = 'Pendiente'
            df['Comentario'] = ''
            df['Candidatos_Conciliacion'] = ''

            fecha_contable_calc = pd.to_datetime(df[col_fecha_contable], errors='coerce')
            df['Periodo_Contable'] = fecha_contable_calc.dt.to_period('M').astype(str)
            df.loc[fecha_contable_calc.isna(), 'Periodo_Contable'] = 'SIN_FECHA_CONTABLE'

            grp_multi = [col_banco, 'Abs_Importe', col_fecha, col_referencia, 'Periodo_Contable']
            df['Posiciones_Mismo_Doc'] = df.groupby([col_doc] + grp_multi)[col_doc].transform('count')
            df['Total_Posiciones_Grupo'] = df.groupby(grp_multi)[col_doc].transform('count')
            df['Docs_Unicos_Grupo'] = df.groupby(grp_multi)[col_doc].transform('nunique')
            df['Tiene_Posiciones_Repetidas'] = (df['Posiciones_Mismo_Doc'] > 1)

            df['Doc_ColB_Repite'] = df.groupby(col_doc)[col_doc].transform('count') > 1

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

            def clasificar_distribuidora(row):
                ref_val = str(row.get(col_referencia, "")).strip()
                ref_val = re.sub(r'\.0$', '', ref_val)
                if ref_val in mapeo_referencias_dist: return mapeo_referencias_dist[ref_val]

                texto_val = row.get(col_texto, "") if col_texto else ""
                novedad_val = row.get(col_novedad, "") if col_novedad else ""
                asig_val = str(row.get(col_asignacion, "")) if col_asignacion else ""
                t = f"{texto_val} {novedad_val} {asig_val} {ref_val}".upper()

                if 'DOSQ' in t or 'D504' in t: return 'Dist Dosquebradas'
                if 'ACOPI' in t or 'D503' in t: return 'Dist Acopi'
                if 'PASTO' in t or 'D505' in t: return 'Dist Pasto'
                if 'BUGA' in t or 'D502' in t: return 'Dist Buga'

                numeros = re.findall(r'\b\d{4}\b', t)
                for n in numeros:
                    num = int(n)
                    if 2000 <= num <= 2999: return 'Dist Buga'
                    if 3000 <= num <= 3999: return 'Dist Acopi'
                    if 4000 <= num <= 4999: return 'Dist Dosquebradas'
                    if 6000 <= num <= 6999: return 'Dist Pasto'
                return 'Sin clasificar'

            def obtener_ref_homologada(row):
                texto = f"{row.get(col_referencia,'')} {row.get(col_asignacion,'')} {row.get(col_texto,'') if col_texto else ''} {row.get(col_novedad,'') if col_novedad else ''}".upper()
                numeros_8 = re.findall(r'\b\d{8}\b', texto)
                for num in numeros_8:
                    if num in mapeo_datafono_ref: return mapeo_datafono_ref[num]
                numeros_4 = re.findall(r'\b\d{4}\b', texto)
                for num in numeros_4:
                    if num in mapeo_datafono_ref.values(): return num
                return None

            df['Distribuidora'] = df.apply(clasificar_distribuidora, axis=1)

            def es_nequi(row):
                texto = f"{row.get(col_texto,'') if col_texto else ''} {row.get(col_asignacion,'')} {row.get(col_referencia,'')}".upper()
                return 'NEQUI' in texto

            def limpiar_numero(valor):
                if pd.isna(valor): return ''
                t = re.sub(r'\.0$', '', str(valor).strip())
                m = re.findall(r'\d+', t)
                return m[0] if m else ''

            df['Es_Nequi'] = df.apply(es_nequi, axis=1)
            df['Referencia_Limpia'] = df[col_referencia].apply(limpiar_numero)
            df['Asignacion_Limpia'] = df[col_asignacion].apply(limpiar_numero)

            def set_estado(indices, estado):
                if indices: df.loc[df['ID_Temp'].isin(indices), 'Estado_Conciliacion'] = estado

            def set_comentarios(dic_comentarios):
                for id_temp, texto in dic_comentarios.items(): df.loc[df['ID_Temp'] == id_temp, 'Comentario'] = texto

            def resumen_docs(sub_df): return ", ".join(str(int(d)) for d in sub_df[col_doc].tolist())

            def es_valor_redondo(v): return (v % multiplo_redondo == 0) and v > 0

            # =========================================================
            # GATE DE SEGURIDAD v24
            # =========================================================
            def candidato_seguro(id_a, id_b, exigir_mismo_importe=True, exigir_mismo_periodo=True):
                ra = df.loc[df['ID_Temp'] == id_a].iloc[0]
                rb = df.loc[df['ID_Temp'] == id_b].iloc[0]

                pa, pb = ra['Periodo_Contable'], rb['Periodo_Contable']
                if exigir_mismo_periodo:
                    if pa == 'SIN_FECHA_CONTABLE' or pb == 'SIN_FECHA_CONTABLE': return False, "Sin periodo contable valido", False
                    if pa != pb: return False, f"Periodo distinto ({pa} vs {pb})", False

                fa, fb = ra['Fecha_Calc'], rb['Fecha_Calc']
                if pd.isna(fa) or pd.isna(fb): return False, "Fecha valor invalida", False

                dif_dias = abs((fa - fb).days)
                es_alerta = dif_dias > 0
                if dif_dias > TOPE_DIAS_ALERTA: return False, f"Diferencia de fecha fuera de rango ({dif_dias} dias)", False

                banco_a, banco_b = str(ra[col_banco]).strip(), str(rb[col_banco]).strip()
                if banco_a != banco_b: return False, f"Banco distinto ({banco_a} vs {banco_b})", False

                if exigir_mismo_importe:
                    imp_a, imp_b = abs(ra[col_importe]), abs(rb[col_importe])
                    if round(imp_a, 2) != round(imp_b, 2): return False, "Importe distinto", False

                dist_a, dist_b = str(ra.get('Distribuidora', '')).strip(), str(rb.get('Distribuidora', '')).strip()
                if dist_a not in ('', 'Sin clasificar') and dist_b not in ('', 'Sin clasificar'):
                    if dist_a != dist_b: return False, f"Distribuidora distinta ({dist_a} vs {dist_b})", False

                periodo_distinto_txt = ""
                if not exigir_mismo_periodo and pa != pb and pa != 'SIN_FECHA_CONTABLE' and pb != 'SIN_FECHA_CONTABLE':
                    periodo_distinto_txt = f" [Periodo contable distinto: {pa} vs {pb}, permitido por Referencia exacta]"

                if bool(ra.get('Es_Nequi', False)) or bool(rb.get('Es_Nequi', False)):
                    return True, f"Candidato NEQUI: requiere verificacion manual{periodo_distinto_txt}", es_alerta

                return True, f"Candidato valido (reglas transversales cumplidas){periodo_distinto_txt}", es_alerta

            def formato_doc(id_temp):
                r = df.loc[df['ID_Temp'] == id_temp].iloc[0]
                d = str(int(r[col_doc])) if pd.notna(r[col_doc]) else ""
                c = str(r[col_clase_doc]) if usar_ipcb and pd.notna(r.get(col_clase_doc)) else ""
                k = str(r[col_clave])
                return f"{d} ({c}={k})" if c and c.lower() != 'nan' else f"{d} (Clv {k})"

            def escribir_candidato(id_a, id_b, estado_si_ok, comentario_base, exigir_mismo_importe=True, exigir_mismo_periodo=True):
                ok, motivo, es_alerta = candidato_seguro(id_a, id_b, exigir_mismo_importe, exigir_mismo_periodo)

                if not ok:
                    for idx in (id_a, id_b):
                        mask = df['ID_Temp'] == idx
                        anterior = str(df.loc[mask, 'Comentario'].iloc[0])
                        df.loc[mask, 'Comentario'] = f"Relación bloqueada: {motivo}." if not anterior else anterior + f" | Relación bloqueada: {motivo}."
                    return False

                texto_candidatos = f"{formato_doc(id_a)} | {formato_doc(id_b)}"
                estado_final = "Alerta de texto - Diferencia de fecha (mismo periodo)" if es_alerta else estado_si_ok
                comentario_final = f"{comentario_base} ({motivo})."

                for idx in (id_a, id_b):
                    mask = df['ID_Temp'] == idx
                    if df.loc[mask, 'Estado_Conciliacion'].iloc[0] == 'Pendiente':
                        df.loc[mask, 'Estado_Conciliacion'] = estado_final
                    df.loc[mask, 'Candidatos_Conciliacion'] = texto_candidatos
                    anterior = str(df.loc[mask, 'Comentario'].iloc[0])
                    df.loc[mask, 'Comentario'] = comentario_final if not anterior else anterior + " | " + comentario_final

                return True

            usados_global = set()

            # =========================================================
            # NIVEL 1: CRUCES EXACTOS Y MÚLTIPLES
            # =========================================================
            df_40, df_50 = df[df[col_clave] == '40'].copy(), df[df[col_clave] == '50'].copy()

            llave_1 = [col_banco, 'Abs_Importe', col_fecha, col_referencia]
            df_40['T'], df_50['T'] = df_40.groupby(llave_1).cumcount(), df_50.groupby(llave_1).cumcount()
            c1 = pd.merge(df_40, df_50, on=llave_1 + ['T'], suffixes=('_40', '_50'))

            llave_1_asig = [col_banco, 'Abs_Importe', col_fecha, col_asignacion]
            llave_1_ref = [col_banco, 'Abs_Importe', col_fecha, col_referencia]
            df_40['T2'], df_50['T2'] = df_40.groupby(llave_1_asig).cumcount(), df_50.groupby(llave_1_ref).cumcount()
            c2 = pd.merge(df_40, df_50, left_on=llave_1_asig + ['T2'], right_on=llave_1_ref + ['T2'], suffixes=('_40', '_50'))

            df_40['T3'], df_50['T3'] = df_40.groupby(llave_1_ref).cumcount(), df_50.groupby(llave_1_asig).cumcount()
            c3 = pd.merge(df_40, df_50, left_on=llave_1_ref + ['T3'], right_on=llave_1_asig + ['T3'], suffixes=('_40', '_50'))

            ind_r1 = set(c1['ID_Temp_40']) | set(c1['ID_Temp_50']) | set(c2['ID_Temp_40']) | set(c2['ID_Temp_50']) | set(c3['ID_Temp_40']) | set(c3['ID_Temp_50'])
            ind_r1_limpio = {i for i in ind_r1 if not df.loc[df['ID_Temp'] == i, 'Tiene_Posiciones_Repetidas'].iloc[0]}
            ind_r1_multi = ind_r1 - ind_r1_limpio

            for c in (c1, c2, c3):
                for _, r in c.iterrows():
                    ida, idb = r['ID_Temp_40'], r['ID_Temp_50']
                    if ida in usados_global or idb in usados_global: continue
                    if ida in ind_r1_limpio and idb in ind_r1_limpio:
                        if escribir_candidato(ida, idb, 'Conciliado - Cruce exacto', "Cruce exacto", exigir_mismo_periodo=False):
                            usados_global.update([ida, idb])

            for tid in (ind_r1_multi - usados_global):
                fila_tid = df.loc[df['ID_Temp'] == tid].iloc[0]
                mask = df['ID_Temp'] == tid
                anterior = str(df.loc[mask, 'Comentario'].iloc[0])
                texto = f"Mismo Doc. tiene {int(fila_tid['Posiciones_Mismo_Doc'])} posiciones idénticas."
                df.loc[mask, 'Comentario'] = texto if not anterior else anterior + " | " + texto

            df_p0 = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_p0['A_L'], df_p0['R_L'] = df_p0['Asignacion_Limpia'], df_p0['Referencia_Limpia']
            d40b, d50b = df_p0[df_p0[col_clave] == '40'].drop(columns=['R_L']).copy(), df_p0[df_p0[col_clave] == '50'].drop(columns=['A_L']).copy()
            d40b['Tb'] = d40b.groupby([col_banco, 'Abs_Importe', col_fecha, 'Periodo_Contable', 'A_L']).cumcount()
            d50b['Tb'] = d50b.groupby([col_banco, 'Abs_Importe', col_fecha, 'Periodo_Contable', 'R_L']).cumcount()

            c1b = pd.merge(d40b, d50b, left_on=[col_banco, 'Abs_Importe', col_fecha, 'Periodo_Contable', 'A_L', 'Tb'],
                           right_on=[col_banco, 'Abs_Importe', col_fecha, 'Periodo_Contable', 'R_L', 'Tb'], suffixes=('_40', '_50'))
            c1b = c1b[c1b['A_L'].notna() & (c1b['A_L'] != '')]

            for _, r in c1b.iterrows():
                ida, idb = r['ID_Temp_40'], r['ID_Temp_50']
                if ida in usados_global or idb in usados_global: continue
                if df.loc[df['ID_Temp'].isin([ida, idb]), 'Tiene_Posiciones_Repetidas'].any():
                    df.loc[df['ID_Temp'].isin([ida, idb]), 'Estado_Conciliacion'] = 'Sugerencia fuerte - Doc. con posiciones múltiples (VERIFICAR)'
                    continue
                if escribir_candidato(ida, idb, 'Conciliado - Cruce exacto (Ref limpia)', "Cruce ref limpiada"):
                    usados_global.update([ida, idb])

            ind_1c_ipcb = set()
            if usar_ipcb:
                df_p_ipcb = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
                df_p_ipcb['Ref_H'] = df_p_ipcb.apply(obtener_ref_homologada, axis=1)
                df_vp = df_p_ipcb[df_p_ipcb['Ref_H'].notna()]
                df_ip = df_vp[(df_vp[col_clase_doc].astype(str).str.upper() == 'IP') & (df_vp[col_clave] == '40')]
                df_cb = df_vp[(df_vp[col_clase_doc].astype(str).str.upper() == 'CB') & (df_vp[col_clave] == '50')]

                if not df_ip.empty and not df_cb.empty:
                    grp_ip = df_ip.groupby([col_banco, 'Periodo_Contable', 'Ref_H'])['Abs_Importe'].sum().reset_index(name='S_IP')
                    grp_cb = df_cb.groupby([col_banco, 'Periodo_Contable', 'Ref_H'])['Abs_Importe'].sum().reset_index(name='S_CB')
                    m_ipcb = pd.merge(grp_cb, grp_ip, on=[col_banco, 'Periodo_Contable', 'Ref_H'])
                    matches_ipcb = m_ipcb[m_ipcb['S_CB'].round(2) == m_ipcb['S_IP'].round(2)]

                    for _, m in matches_ipcb.iterrows():
                        b, periodo, rh = m[col_banco], m['Periodo_Contable'], m['Ref_H']
                        sub_ip = df_ip[(df_ip[col_banco] == b) & (df_ip['Periodo_Contable'] == periodo) & (df_ip['Ref_H'] == rh)]
                        sub_cb = df_cb[(df_cb[col_banco] == b) & (df_cb['Periodo_Contable'] == periodo) & (df_cb['Ref_H'] == rh)]

                        ip_ids = [i for i in sub_ip['ID_Temp'].tolist() if i not in usados_global]
                        cb_ids = [i for i in sub_cb['ID_Temp'].tolist() if i not in usados_global]
                        if not ip_ids or not cb_ids: continue

                        if not any(candidato_seguro(a, c, exigir_mismo_importe=False)[0] for a in ip_ids for c in cb_ids): continue
                        
                        ind_1c_ipcb.update(ip_ids + cb_ids); usados_global.update(ip_ids + cb_ids)
                        txt = f"Cruce múltiple IP/CB. Ref: {rh}."
                        for idx in ip_ids + cb_ids:
                            df.loc[df['ID_Temp'] == idx, 'Estado_Conciliacion'] = 'Conciliado - Cruce múltiple'
                            df.loc[df['ID_Temp'] == idx, 'Candidatos_Conciliacion'] = " | ".join(formato_doc(i) for i in ip_ids + cb_ids)
                            otros = resumen_docs(sub_cb) if idx in ip_ids else resumen_docs(sub_ip)
                            df.loc[df['ID_Temp'] == idx, 'Comentario'] = f"{txt} Docs relacionados: {otros}"

            ind_1c_bis_ipcb = set()
            if usar_ipcb:
                df_p_ipcb_bis = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
                df_p_ipcb_bis['Ref_H'] = df_p_ipcb_bis.apply(obtener_ref_homologada, axis=1)
                df_vp_bis = df_p_ipcb_bis[df_p_ipcb_bis['Ref_H'].notna()]
                df_ip_bis = df_vp_bis[(df_vp_bis[col_clase_doc].astype(str).str.upper() == 'IP') & (df_vp_bis[col_clave] == '40')]
                df_cb_bis = df_vp_bis[(df_vp_bis[col_clase_doc].astype(str).str.upper() == 'CB') & (df_vp_bis[col_clave] == '50')]

                if not df_ip_bis.empty and not df_cb_bis.empty:
                    grp_ip_bis = df_ip_bis.groupby([col_banco, 'Periodo_Contable', 'Ref_H'])['Abs_Importe'].sum().reset_index(name='S_IP')
                    grp_cb_bis = df_cb_bis.groupby([col_banco, 'Periodo_Contable', 'Ref_H'])['Abs_Importe'].sum().reset_index(name='S_CB')
                    m_ipcb_bis = pd.merge(grp_cb_bis, grp_ip_bis, on=[col_banco, 'Periodo_Contable', 'Ref_H'])
                    m_ipcb_bis['DifV'] = (m_ipcb_bis['S_CB'] - m_ipcb_bis['S_IP']).abs()
                    max_suma = m_ipcb_bis[['S_CB', 'S_IP']].max(axis=1).clip(lower=1)
                    m_ipcb_bis['Pct'] = m_ipcb_bis['DifV'] / max_suma
                    matches_tolerancia = m_ipcb_bis[(m_ipcb_bis['DifV'] > 0) & ((m_ipcb_bis['DifV'] <= tol_valor_abs) | (m_ipcb_bis['Pct'] <= tol_valor_pct))]

                    for _, m in matches_tolerancia.iterrows():
                        b, periodo, rh = m[col_banco], m['Periodo_Contable'], m['Ref_H']
                        sub_ip = df_ip_bis[(df_ip_bis[col_banco] == b) & (df_ip_bis['Periodo_Contable'] == periodo) & (df_ip_bis['Ref_H'] == rh)]
                        sub_cb = df_cb_bis[(df_cb_bis[col_banco] == b) & (df_cb_bis['Periodo_Contable'] == periodo) & (df_cb_bis['Ref_H'] == rh)]
                        ip_ids = [i for i in sub_ip['ID_Temp'].tolist() if i not in usados_global]
                        cb_ids = [i for i in sub_cb['ID_Temp'].tolist() if i not in usados_global]
                        if not ip_ids or not cb_ids: continue
                        if not any(candidato_seguro(a, c, exigir_mismo_importe=False)[0] for a in ip_ids for c in cb_ids): continue
                        
                        ind_1c_bis_ipcb.update(ip_ids + cb_ids); usados_global.update(ip_ids + cb_ids)
                        txt = f"Sugerencia: cruce múltiple IP/CB con diferencia (${m['DifV']:,.0f}). Ref: {rh}."
                        for idx in ip_ids + cb_ids:
                            df.loc[df['ID_Temp'] == idx, 'Estado_Conciliacion'] = 'Sugerencia - Cruce múltiple IP/CB con diferencia de valor'
                            df.loc[df['ID_Temp'] == idx, 'Candidatos_Conciliacion'] = " | ".join(formato_doc(i) for i in ip_ids + cb_ids)
                            df.loc[df['ID_Temp'] == idx, 'Comentario'] = f"{txt} Docs relacionados: {resumen_docs(sub_cb) if idx in ip_ids else resumen_docs(sub_ip)}"

            # =========================================================
            # NIVEL 2: SECTORIZACION Y VALOR REDONDO
            # =========================================================
            df_p1d = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            if usar_ipcb: df_p1d = df_p1d[df_p1d[col_clase_doc].astype(str).str.upper() != 'IP']
            d40c, d50c = df_p1d[df_p1d['Distribuidora'] != 'Sin clasificar'][df_p1d[col_clave] == '40'], df_p1d[df_p1d['Distribuidora'] != 'Sin clasificar'][df_p1d[col_clave] == '50']

            ind_r1d_f, ind_r1d_a, com_r1d_a, cand_r1d_a = set(), set(), {}, {}

            if not d40c.empty and not d50c.empty:
                for grp, sub40 in d40c.groupby([col_banco, 'Abs_Importe', col_fecha, 'Periodo_Contable', 'Distribuidora']):
                    b, imp, f, periodo, dist = grp
                    sub50 = d50c[(d50c[col_banco] == b) & (d50c['Abs_Importe'] == imp) & (d50c[col_fecha] == f) & (d50c['Periodo_Contable'] == periodo) & (d50c['Distribuidora'] == dist)]
                    if sub50.empty: continue

                    s40_ord = sub40[~sub40['ID_Temp'].isin(usados_global)].sort_values(col_doc)
                    s50_ord = sub50[~sub50['ID_Temp'].isin(usados_global)].sort_values(col_doc)
                    if s40_ord.empty or s50_ord.empty: continue

                    if len(s40_ord) == 1 and len(s50_ord) == 1:
                        ida, idb = s40_ord.iloc[0]['ID_Temp'], s50_ord.iloc[0]['ID_Temp']
                        if escribir_candidato(ida, idb, 'Conciliado - Cruce Distribuidora', f"Candidato único sede ({dist})"):
                            usados_global.update([ida, idb])
                    elif len(s40_ord) == len(s50_ord):
                        for (_, r40), (_, r50) in zip(s40_ord.iterrows(), s50_ord.iterrows()):
                            ida, idb = r40['ID_Temp'], r50['ID_Temp']
                            if ida in usados_global or idb in usados_global: continue
                            if escribir_candidato(ida, idb, "Sugerencia fuerte: Sectorización (FIFO)", f"Sede '{dist}' emparejado FIFO"):
                                ind_r1d_f.update([ida, idb]); usados_global.update([ida, idb])
                    else:
                        for _, r in s40_ord.iterrows():
                            ind_r1d_a.add(r['ID_Temp']); com_r1d_a[r['ID_Temp']] = f"Sede '{dist}' desbalance. Créditos: {resumen_docs(s50_ord)}"
                            cand_r1d_a[r['ID_Temp']] = f"{formato_doc(r['ID_Temp'])} | Posibles (créditos): {resumen_docs(s50_ord)}"
                        for _, r in s50_ord.iterrows():
                            ind_r1d_a.add(r['ID_Temp']); com_r1d_a[r['ID_Temp']] = f"Sede '{dist}' desbalance. Débitos: {resumen_docs(s40_ord)}"
                            cand_r1d_a[r['ID_Temp']] = f"{formato_doc(r['ID_Temp'])} | Posibles (débitos): {resumen_docs(s40_ord)}"

            set_estado(ind_r1d_a, 'Sugerencia: Sugerencia por Distribuidora Multiples')
            set_comentarios(com_r1d_a)
            for idt, txt_cand in cand_r1d_a.items():
                if idt not in usados_global: df.loc[df['ID_Temp'] == idt, 'Candidatos_Conciliacion'] = txt_cand

            df_p = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            if usar_ipcb: df_p = df_p[df_p[col_clase_doc].astype(str).str.upper() != 'IP']
            grp_c = [col_banco, 'Abs_Importe', col_fecha, 'Periodo_Contable']

            df_p40, df_p50 = df_p[df_p[col_clave] == '40'].copy(), df_p[df_p[col_clave] == '50'].copy()
            df_p40['n40'], df_p50['n50'] = df_p40.groupby(grp_c)['ID_Temp'].transform('count'), df_p50.groupby(grp_c)['ID_Temp'].transform('count')

            c_un = pd.merge(df_p40[df_p40['n40'] == 1], df_p50[df_p50['n50'] == 1], on=grp_c, suffixes=('_40', '_50'))
            for _, r in c_un.iterrows():
                ida, idb = r['ID_Temp_40'], r['ID_Temp_50']
                if ida in usados_global or idb in usados_global: continue
                if escribir_candidato(ida, idb, 'Conciliado - Cruce unico', "Único candidato en banco/importe/fecha/periodo"):
                    usados_global.update([ida, idb])

            rem40, rem50 = df_p40[~df_p40['ID_Temp'].isin(usados_global)], df_p50[~df_p50['ID_Temp'].isin(usados_global)]
            ind_amb, com_amb, cand_amb = set(), {}, {}

            for grp, sub40 in rem40.groupby(grp_c):
                b, imp, f, periodo = grp
                sub50 = rem50[(rem50[col_banco] == b) & (rem50['Abs_Importe'] == imp) & (rem50[col_fecha] == f) & (rem50['Periodo_Contable'] == periodo)]
                if sub50.empty: continue

                if es_valor_redondo(imp):
                    s40_ord = sub40[~sub40['ID_Temp'].isin(usados_global)].sort_values(col_doc)
                    s50_ord = sub50[~sub50['ID_Temp'].isin(usados_global)].sort_values(col_doc)
                    if s40_ord.empty or s50_ord.empty: continue
                    if len(s40_ord) == len(s50_ord):
                        for (_, r40), (_, r50) in zip(s40_ord.iterrows(), s50_ord.iterrows()):
                            ida, idb = r40['ID_Temp'], r50['ID_Temp']
                            if ida in usados_global or idb in usados_global: continue
                            if escribir_candidato(ida, idb, "Sugerencia fuerte: Valor redondo (FIFO)", f"Valor redondo (${imp:,.0f}) FIFO per {periodo}"):
                                usados_global.update([ida, idb])
                    else:
                        for _, r in s40_ord.iterrows():
                            ind_amb.add(r['ID_Temp']); com_amb[r['ID_Temp']] = f"Confuso ({len(s40_ord)} vs {len(s50_ord)}). Créditos: {resumen_docs(s50_ord)}"
                            cand_amb[r['ID_Temp']] = f"{formato_doc(r['ID_Temp'])} | Posibles (créditos): {resumen_docs(s50_ord)}"
                        for _, r in s50_ord.iterrows():
                            ind_amb.add(r['ID_Temp']); com_amb[r['ID_Temp']] = f"Confuso ({len(s50_ord)} vs {len(s40_ord)}). Débitos: {resumen_docs(s40_ord)}"
                            cand_amb[r['ID_Temp']] = f"{formato_doc(r['ID_Temp'])} | Posibles (débitos): {resumen_docs(s40_ord)}"
                else:
                    for _, r in sub40.iterrows():
                        ind_amb.add(r['ID_Temp']); com_amb[r['ID_Temp']] = f"{len(sub50)} posibles cruces (mismo periodo). Docs: {resumen_docs(sub50)}"
                        cand_amb[r['ID_Temp']] = f"{formato_doc(r['ID_Temp'])} | Candidatos posibles: {resumen_docs(sub50)}"
                    for _, r in sub50.iterrows():
                        ind_amb.add(r['ID_Temp']); com_amb[r['ID_Temp']] = f"{len(sub40)} posibles cruces (mismo periodo). Docs: {resumen_docs(sub40)}"
                        cand_amb[r['ID_Temp']] = f"{formato_doc(r['ID_Temp'])} | Candidatos posibles: {resumen_docs(sub40)}"

            set_estado(ind_amb - usados_global, 'Sugerencia: Solicitar soporte')
            set_comentarios(com_amb)
            for idt, txt_cand in cand_amb.items():
                if idt not in usados_global: df.loc[df['ID_Temp'] == idt, 'Candidatos_Conciliacion'] = txt_cand

            # =========================================================
            # NIVEL 3: ALERTAS DE FECHA, RECLASIFICACION, DIF VALOR
            # =========================================================
            df_pend = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_pend = df_pend[~df_pend['ID_Temp'].isin(usados_global)]
            df_v = df_pend[df_pend['Referencia_Limpia'] != ''].copy()

            df_4n, df_5n = df_v[df_v[col_clave] == '40'].copy(), df_v[df_v[col_clave] == '50'].copy()

            sA = pd.merge(df_4n, df_5n, on=[col_banco, 'Abs_Importe', 'Referencia_Limpia', 'Periodo_Contable'], suffixes=('_40', '_50'))
            sA['Dif'] = (sA['Fecha_Calc_40'] - sA['Fecha_Calc_50']).dt.days.abs()
            sA = sA[(sA['Dif'] > 0) & (sA['Dif'] <= TOPE_DIAS_ALERTA)].sort_values('Dif').drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            for _, r in sA.iterrows():
                ida, idb = r['ID_Temp_40'], r['ID_Temp_50']
                if ida in usados_global or idb in usados_global: continue
                if escribir_candidato(ida, idb, 'Alerta de texto - Diferencia de fecha (mismo periodo)', f"Referencia {r['Referencia_Limpia']}, difiere {int(r['Dif'])} día(s)"):
                    usados_global.update([ida, idb])

            df_4n, df_5n = df_4n[~df_4n['ID_Temp'].isin(usados_global)], df_5n[~df_5n['ID_Temp'].isin(usados_global)]

            d40_ab, d50_ab = df_4n[df_4n['Asignacion_Limpia'] != ''].copy(), df_5n[df_5n['Referencia_Limpia'] != ''].copy()
            sAB = pd.merge(d40_ab, d50_ab, left_on=[col_banco, 'Abs_Importe', 'Asignacion_Limpia', 'Periodo_Contable'], right_on=[col_banco, 'Abs_Importe', 'Referencia_Limpia', 'Periodo_Contable'], suffixes=('_40', '_50'))
            sAB['Dif_ab'] = (sAB['Fecha_Calc_40'] - sAB['Fecha_Calc_50']).dt.days.abs()
            sAB = sAB[sAB['Dif_ab'] <= TOPE_DIAS_ALERTA].sort_values('Dif_ab').drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            for _, r in sAB.iterrows():
                ida, idb = r['ID_Temp_40'], r['ID_Temp_50']
                if ida in usados_global or idb in usados_global: continue
                estado_ok = 'Conciliado - Cruce exacto (Ref limpia Asig/Ref)' if int(r['Dif_ab']) == 0 else 'Alerta de texto - Diferencia de fecha Ref limpia'
                if escribir_candidato(ida, idb, estado_ok, "Ref limpia Asignación/Referencia"):
                    usados_global.update([ida, idb])

            df_4n, df_5n = df_4n[~df_4n['ID_Temp'].isin(usados_global)], df_5n[~df_5n['ID_Temp'].isin(usados_global)]

            sB = pd.merge(df_4n, df_5n, on=['Abs_Importe', col_fecha, 'Referencia_Limpia'], suffixes=('_40', '_50'))
            sB = sB[sB[f'{col_banco}_40'] != sB[f'{col_banco}_50']].drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            com_B = {}
            for _, r in sB.iterrows():
                ida, idb = r['ID_Temp_40'], r['ID_Temp_50']
                if ida in usados_global or idb in usados_global: continue
                if r['Periodo_Contable_40'] != r['Periodo_Contable_50']: continue
                usados_global.update([ida, idb])
                com_B[ida] = f"Registrado en banco '{r[col_banco+'_50']}'. Doc: {int(r[col_doc+'_50'])}"
                com_B[idb] = f"Registrado en banco '{r[col_banco+'_40']}'. Doc: {int(r[col_doc+'_40'])}"
                df.loc[df['ID_Temp'].isin([ida, idb]), 'Estado_Conciliacion'] = 'Reclasificación de banco'
                df.loc[df['ID_Temp'].isin([ida, idb]), 'Candidatos_Conciliacion'] = f"{formato_doc(ida)} | {formato_doc(idb)}"

            set_comentarios(com_B)

            df_4n, df_5n = df_4n[~df_4n['ID_Temp'].isin(usados_global)], df_5n[~df_5n['ID_Temp'].isin(usados_global)]

            sC = pd.merge(df_4n, df_5n, on=[col_banco, col_fecha, 'Referencia_Limpia'], suffixes=('_40', '_50'))
            sC = sC[sC['Periodo_Contable_40'] == sC['Periodo_Contable_50']]
            sC['DifV'] = (sC['Abs_Importe_40'] - sC['Abs_Importe_50']).abs()
            max_imp = sC[['Abs_Importe_40', 'Abs_Importe_50']].max(axis=1)
            sC['Pct'] = np.where(max_imp == 0, 0, sC['DifV'] / max_imp)
            sC = sC[(sC['DifV'] > 0) & ((sC['DifV'] <= tol_valor_abs) | (sC['Pct'] <= tol_valor_pct))].sort_values('DifV').drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            com_C = {}
            for _, r in sC.iterrows():
                ida, idb = r['ID_Temp_40'], r['ID_Temp_50']
                if ida in usados_global or idb in usados_global: continue
                usados_global.update([ida, idb])
                com_C[ida] = f"Diferencia de ${r['DifV']:,.0f} ({r['Pct']*100:.2f}%). Doc: {int(r[col_doc+'_50'])}"
                com_C[idb] = f"Diferencia de ${r['DifV']:,.0f} ({r['Pct']*100:.2f}%). Doc: {int(r[col_doc+'_40'])}"
                df.loc[df['ID_Temp'].isin([ida, idb]), 'Estado_Conciliacion'] = 'Diferencia en valor'
                df.loc[df['ID_Temp'].isin([ida, idb]), 'Candidatos_Conciliacion'] = f"{formato_doc(ida)} | {formato_doc(idb)}"

            set_comentarios(com_C)

            df_p11d = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_p11d = df_p11d[~df_p11d['ID_Temp'].isin(usados_global)]
            if usar_ipcb: df_p11d = df_p11d[df_p11d[col_clase_doc].astype(str).str.upper() != 'IP']

            d4d, d5d = df_p11d[df_p11d[col_clave] == '40'], df_p11d[df_p11d[col_clave] == '50']
            if not d4d.empty and not d5d.empty:
                sD = pd.merge(d4d, d5d, on=[col_banco, col_fecha], suffixes=('_40', '_50'))
                sD = sD[sD['Periodo_Contable_40'] == sD['Periodo_Contable_50']]
                sD['DifV'] = (sD['Abs_Importe_40'] - sD['Abs_Importe_50']).abs()
                max_impD = sD[['Abs_Importe_40', 'Abs_Importe_50']].max(axis=1)
                sD['Pct'] = np.where(max_impD == 0, 0, sD['DifV'] / max_impD)
                sDt = sD[(sD['DifV'] > 0) & ((sD['DifV'] <= tol_valor_abs) | (sD['Pct'] <= tol_valor_pct))].copy()
                if not sDt.empty:
                    sDt['n4'], sDt['n5'] = sDt.groupby('ID_Temp_40')['ID_Temp_50'].transform('count'), sDt.groupby('ID_Temp_50')['ID_Temp_40'].transform('count')
                    sDu = sDt[(sDt['n4'] == 1) & (sDt['n5'] == 1)]
                    for _, r in sDu.iterrows():
                        ida, idb = r['ID_Temp_40'], r['ID_Temp_50']
                        if ida in usados_global or idb in usados_global: continue
                        usados_global.update([ida, idb])
                        etiq = " [NEQUI: verificar]" if bool(df.loc[df['ID_Temp']==ida, 'Es_Nequi'].iloc[0]) or bool(df.loc[df['ID_Temp']==idb, 'Es_Nequi'].iloc[0]) else ""
                        df.loc[df['ID_Temp']==ida, 'Comentario'] = f"Candidato único con dif. de ${r['DifV']:,.0f}.{etiq} Doc: {int(r[col_doc+'_50'])}"
                        df.loc[df['ID_Temp']==idb, 'Comentario'] = f"Candidato único con dif. de ${r['DifV']:,.0f}.{etiq} Doc: {int(r[col_doc+'_40'])}"
                        df.loc[df['ID_Temp'].isin([ida, idb]), 'Estado_Conciliacion'] = 'Diferencia en valor (NEQUI)'
                        df.loc[df['ID_Temp'].isin([ida, idb]), 'Candidatos_Conciliacion'] = f"{formato_doc(ida)} | {formato_doc(idb)}"

            ind_sector_redondos = set()
            if activar_sector_redondos:
                pend_sector = df[(df['Estado_Conciliacion'] == 'Pendiente') & (df['Candidatos_Conciliacion'] == '') & (df['Abs_Importe'] > 0) & (df['Abs_Importe'] % multiplo_redondo == 0) & (df['Periodo_Contable'] != 'SIN_FECHA_CONTABLE') & (df['Distribuidora'] != 'Sin clasificar')].copy()
                if usar_ipcb: pend_sector = pend_sector[pend_sector[col_clase_doc].astype(str).str.upper() != 'IP']
                for (banco_g, importe_g, periodo_g, dist_g), grupo_sector in pend_sector.groupby([col_banco, 'Abs_Importe', 'Periodo_Contable', 'Distribuidora']):
                    lado_40, lado_50 = grupo_sector[grupo_sector[col_clave] == '40'].sort_values(['Fecha_Calc', 'ID_Temp']), grupo_sector[grupo_sector[col_clave] == '50'].sort_values(['Fecha_Calc', 'ID_Temp'])
                    n_pares_sector = min(len(lado_40), len(lado_50))
                    if n_pares_sector == 0: continue
                    for i in range(n_pares_sector):
                        id_40, id_50 = lado_40.iloc[i]['ID_Temp'], lado_50.iloc[i]['ID_Temp']
                        if id_40 in usados_global or id_50 in usados_global: continue
                        ok, motivo, es_alerta = candidato_seguro(id_40, id_50, exigir_mismo_importe=True)
                        if not ok: continue
                        dif_dias_sector = abs((lado_40.iloc[i]['Fecha_Calc'] - lado_50.iloc[i]['Fecha_Calc']).days)
                        estado_sector = f"Sugerencia fuerte - Sectorización valor redondo ({dist_g}, mismo día)" if dif_dias_sector == 0 else f"Sugerencia fuerte - Sectorización valor redondo ({dist_g}, fechas distintas)"
                        df.loc[df['ID_Temp'].isin([id_40, id_50]), 'Estado_Conciliacion'] = estado_sector
                        df.loc[df['ID_Temp'].isin([id_40, id_50]), 'Candidatos_Conciliacion'] = f"{formato_doc(id_40)} | {formato_doc(id_50)}"
                        df.loc[df['ID_Temp'].isin([id_40, id_50]), 'Comentario'] = f"Sectorización '{dist_g}': valor redondo ${importe_g:,.0f}. Diferencia {dif_dias_sector} día(s)."
                        ind_sector_redondos.update([id_40, id_50]); usados_global.update([id_40, id_50])

            # =========================================================
            # EMPAREJAMIENTO FIFO POSICIONAL FINAL 
            # =========================================================
            def valor_clave_lado(row): return row['Asignacion_Limpia'] if str(row[col_clave]) == '40' else row['Referencia_Limpia']
            df['Valor_Clave_Lado'] = df.apply(valor_clave_lado, axis=1)

            def clave_grupo_ampliado(row):
                if str(row['Valor_Clave_Lado']).strip(): return f"{str(row[col_banco]).strip()}|{str(row['Periodo_Contable']).strip()}|{str(row[col_fecha])}|REF-{str(row['Valor_Clave_Lado']).strip()}"
                return None
            df['Grupo_Ampliado_Key'] = df.apply(clave_grupo_ampliado, axis=1)

            pendientes_sin_candidato = df[(df['Candidatos_Conciliacion'] == '') & (df['Estado_Conciliacion'] == 'Pendiente') & df['Grupo_Ampliado_Key'].notna()].copy()
            ind_fifo_ok, ind_fifo_verde_dz_repetido, ind_fifo_sin_pareja_neutro = set(), set(), set()

            for grupo_key, grupo_df in pendientes_sin_candidato.groupby('Grupo_Ampliado_Key'):
                lado_40, lado_50 = grupo_df[grupo_df[col_clave] == '40'].sort_values('ID_Temp'), grupo_df[grupo_df[col_clave] == '50'].sort_values('ID_Temp')
                n_pares = min(len(lado_40), len(lado_50))

                for i in range(n_pares):
                    id_40, id_50 = lado_40.iloc[i]['ID_Temp'], lado_50.iloc[i]['ID_Temp']
                    etiq = " [NEQUI: verificar manual]" if bool(lado_40.iloc[i].get('Es_Nequi', False)) or bool(lado_50.iloc[i].get('Es_Nequi', False)) else ""
                    df.loc[df['ID_Temp'].isin([id_40, id_50]), 'Candidatos_Conciliacion'] = f"{formato_doc(id_40)} | {formato_doc(id_50)}"
                    df.loc[df['ID_Temp'].isin([id_40, id_50]), 'Estado_Conciliacion'] = 'Conciliado - Candidato principal (posición FIFO)'
                    df.loc[df['ID_Temp'].isin([id_40, id_50]), 'Comentario'] = f"Candidato principal por FIFO dentro del grupo.{etiq}"
                    ind_fifo_ok.update([id_40, id_50])

                for _, fila in lado_40.iloc[n_pares:].iterrows():
                    idx = fila['ID_Temp']
                    if bool(fila['Doc_ColB_Repite']):
                        df.loc[df['ID_Temp']==idx, 'Estado_Conciliacion'] = 'Sugerencia - Clave 40 sin cruce (Doc. repetido en columna B)'
                        df.loc[df['ID_Temp']==idx, 'Comentario'] = f"Clave 40 sin cruce disponible. Doc {int(fila[col_doc])} repite."
                        ind_fifo_verde_dz_repetido.add(idx)
                    else:
                        df.loc[df['ID_Temp']==idx, 'Comentario'] = "Clave 40 sin cruce disponible; doc no repite en columna B."
                        ind_fifo_sin_pareja_neutro.add(idx)

                for _, fila in lado_50.iloc[n_pares:].iterrows():
                    idx = fila['ID_Temp']
                    df.loc[df['ID_Temp']==idx, 'Comentario'] = "Clave 50 excedente sin cruce disponible en su grupo."
                    ind_fifo_sin_pareja_neutro.add(idx)

            sin_p = df['Estado_Conciliacion'] == 'Pendiente'
            if usar_ipcb:
                es_ip = df[col_clase_doc].astype(str).str.upper() == 'IP'
                df.loc[sin_p & es_ip & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia - PDV'
                df.loc[sin_p & ~es_ip & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia segura - revisión manual'
            else:
                df.loc[sin_p & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia segura - revisión manual'

            # =========================================================
            # SELECCIÓN DE COLUMNAS VISIBLES Y FORMATO
            # =========================================================
            columnas_visibles = columnas_originales + ['Estado_Conciliacion', 'Comentario', 'Candidatos_Conciliacion', 'Distribuidora']
            def vista_simplificada(df_cualquiera): return df_cualquiera[[c for c in columnas_visibles if c in df_cualquiera.columns]].copy()

            cuadre_ok = filas_antes == (len(df) + len(filas_descartadas))
            df_final = df.drop(columns=['ID_Temp', 'Abs_Importe', 'Fecha_Calc'], errors='ignore')
            for col_f in [c for c in df_final.columns if 'fe.' in c.lower() or 'fecha' in c.lower() or 'fe-' in c.lower()]:
                df_final[col_f] = pd.to_datetime(df_final[col_f], errors='coerce').dt.strftime('%d/%m/%Y')

            # =========================================================
            # *** REGLA VISUAL ESTRICTA DEL COLOR VERDE ***
            # =========================================================
            def resaltar_conciliados(row):
                est = str(row['Estado_Conciliacion']).strip().lower()
                clave = str(row.get(col_clave, '')).strip()
                clase = str(row.get(col_clase_doc, '')).strip().upper() if col_clase_doc else ''

                es_conciliado = 'conciliado' in est
                es_40_dz = (clave == '40' and clase == 'DZ')

                # 1. REGLA ESTRICTA VERDE: Clave 40, Clase DZ, NO CONCILIADO
                if not es_conciliado and es_40_dz:
                    return [f'background-color: {COLOR_VERIFICAR}; color: black'] * len(row)

                # 2. Pendientes neutros (Blancos)
                if est in ('pendiente', '', 'nan'):
                    return [f'background-color: {COLOR_PENDIENTE}; color: black'] * len(row)

                # 3. Reclasificación (Lila)
                if 'reclasificación' in est or 'otro banco' in est:
                    return [f'background-color: {COLOR_RECLASIFICAR}; color: black'] * len(row)

                # 4. Alertas (Naranja)
                if 'alerta' in est or 'diferencia' in est or 'fecha' in est or 'periodo' in est:
                    return [f'background-color: {COLOR_ALERTA}; color: black'] * len(row)

                # 5. Resto de Sugerencias y FIFO (Amarillo - #FFF2CC)
                if ('sugerencia' in est or ('fifo' in est and 'candidato principal' not in est)
                        or 'multiples' in est or 'múltiples' in est or 'verificar' in est
                        or 'soporte' in est or ('fuerte' in est and 'candidato principal' not in est)
                        or 'sectorización' in est):
                    return [f'background-color: {COLOR_AMARILLO}; color: black'] * len(row)

                # 6. Conciliados (Azul)
                if es_conciliado:
                    return [f'background-color: {COLOR_CONCILIADO}; color: black'] * len(row)

                return [f'background-color: {COLOR_PENDIENTE}; color: black'] * len(row)

            # =========================================================
            # EXPORTACIÓN
            # =========================================================
            output = io.BytesIO()
            b_unicos = [b for b in df_final[col_banco].unique() if str(b).strip().lower() not in ('', 'nan')]
            orden_cuentas = [
                "1110056001", "1110056101", "1110056201", "1110056301",
                "1110056401", "1110056501", "1110056601", "1110056701",
                "1120055001", "1120055101", "1120055301"
            ]
            nombres_ordenados = [mapeo_cuentas_banco.get(c, f"CUENTA {c} (sin mapear)") for c in orden_cuentas]
            def get_bank_order(banco_str):
                if banco_str.strip() in nombres_ordenados: return nombres_ordenados.index(banco_str.strip())
                for i, acc in enumerate(orden_cuentas):
                    if acc in banco_str: return i
                return 999
            b_unicos = sorted(b_unicos, key=get_bank_order)
            nombres_pestanas_usados = set()
            def nombre_pestana_unico(nombre_base):
                nombre = re.sub(r'[\\/*?:\[\]]', '-', str(nombre_base)[:31])
                if not nombre.strip() or nombre.lower() == 'nan': nombre = "Sin_Banco"
                original, contador = nombre, 1
                while nombre in nombres_pestanas_usados:
                    sufijo = f"_{contador}"
                    nombre = original[: 31 - len(sufijo)] + sufijo
                    contador += 1
                nombres_pestanas_usados.add(nombre)
                return nombre

            advertencias_export = []
            def escribir_hoja_segura(writer, df_hoja, nombre_hoja, con_estilo=True):
                nombre_final = nombre_pestana_unico(nombre_hoja)
                try:
                    if con_estilo and not df_hoja.empty:
                        df_hoja.style.apply(lambda row: resaltar_conciliados(df_hoja.loc[row.name]), axis=1).to_excel(writer, index=False, sheet_name=nombre_final)
                    else: df_hoja.to_excel(writer, index=False, sheet_name=nombre_final)
                except Exception as e_estilo:
                    try:
                        df_hoja.to_excel(writer, index=False, sheet_name=nombre_final)
                        advertencias_export.append(f"Hoja '{nombre_final}': formato omitido ({e_estilo}).")
                    except Exception as e_fatal: advertencias_export.append(f"Hoja '{nombre_final}' falló ({e_fatal}).")

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Contadores Generales
                total_filas = len(df_final)
                total_conciliadas = int(df_final['Estado_Conciliacion'].str.contains('Conciliado', na=False).sum())
                total_pendientes = int((df_final['Estado_Conciliacion'] == 'Pendiente').sum())
                total_alertas = int(df_final['Estado_Conciliacion'].str.contains('Alerta|Diferencia', case=False, na=False).sum())
                total_reclasificar = int(df_final['Estado_Conciliacion'].str.contains('Reclasificación', case=False, na=False).sum())
                total_con_candidatos = int((df_final['Candidatos_Conciliacion'] != '').sum())

                # Contador Verde Estricto
                mask_verde = (~df_final['Estado_Conciliacion'].str.contains('Conciliado', case=False, na=False)) & (df_final[col_clave].astype(str).str.strip() == '40')
                if col_clase_doc:
                    mask_verde = mask_verde & (df_final[col_clase_doc].astype(str).str.strip().str.upper() == 'DZ')
                total_verde = int(mask_verde.sum())

                total_sugerencias = int(df_final['Estado_Conciliacion'].str.contains('Sugerencia|FIFO|Verificar|Multiples|Múltiples|Soporte', case=False, na=False).sum())
                
                resumen_df = pd.DataFrame({
                    "Métrica": [
                        "Fecha de procesamiento", "Total filas procesadas", "Conciliadas / Candidato principal (azul)",
                        "Cruce múltiple IP/CB EXACTO", "Sugerencia por Sectorización",
                        "Clave 40 Clase DZ sin conciliar (Verde)", "Otras Sugerencias y Verificaciones (Amarillo)",
                        "Alertas / Diferencias (naranja)", "Reclasificar banco (lila)", "Pendientes (blanco)",
                        "Filas con Candidatos_Conciliacion", "Filas excluidas (sin doc/clave)"
                    ],
                    "Valor": [
                        datetime.now().strftime('%d/%m/%Y %H:%M'), total_filas, total_conciliadas,
                        len(ind_1c_ipcb), len(ind_sector_redondos), total_verde, 
                        total_sugerencias, total_alertas, total_reclasificar, total_pendientes,
                        total_con_candidatos, filas_excluidas
                    ]
                })
                resumen_df.to_excel(writer, index=False, sheet_name='RESUMEN')
                nombres_pestanas_usados.add('RESUMEN')

                df_nov = df_final[~df_final['Estado_Conciliacion'].str.startswith('Conciliado', na=False)].copy()
                df_nov = df_nov[df_nov[col_clave] == '40']
                if not df_nov.empty:
                    df_nov = df_nov.sort_values(by=['Estado_Conciliacion', col_importe])
                    escribir_hoja_segura(writer, vista_simplificada(df_nov), 'NOVEDADES_Y_PENDIENTES_40', con_estilo=True)
                else: escribir_hoja_segura(writer, pd.DataFrame(columns=columnas_visibles), 'NOVEDADES_Y_PENDIENTES_40', con_estilo=False)

                df_multi = df_final[df_final['Tiene_Posiciones_Repetidas'] == True].copy()
                if not df_multi.empty:
                    df_multi = df_multi.sort_values(by=[col_banco, col_importe, col_fecha, col_referencia, col_doc])
                    escribir_hoja_segura(writer, vista_simplificada(df_multi), 'REVISAR_POSICIONES_MULTIPLES', con_estilo=True)

                for banco in b_unicos:
                    df_b = df_final[df_final[col_banco] == banco].copy().sort_values(by=col_importe, ascending=True)
                    if not df_b.empty: escribir_hoja_segura(writer, vista_simplificada(df_b), str(banco), con_estilo=True)

                if not filas_descartadas.empty: escribir_hoja_segura(writer, vista_simplificada(filas_descartadas), 'DESCARTADAS_SIN_DOC_O_CT', con_estilo=False)

            # =========================================================
            # INTERFAZ
            # =========================================================
            st.success("¡Conciliación Integral terminada! Regla visual aplicada para el color Verde exclusivo.")
            if not cuadre_ok: st.warning("⚠️ Revisa la pestaña DESCARTADAS, el total de filas no coincide.")
            if advertencias_export:
                for adv in advertencias_export: st.warning(f"⚠️ {adv}")

            st.markdown(f"""
**Leyenda de colores (6 categorías):**
- <span style="background-color:{COLOR_CONCILIADO}; padding:2px 8px;">Azul: Conciliado / Candidato principal</span>
- <span style="background-color:{COLOR_VERIFICAR}; padding:2px 8px;">Verde: Clave 40 Clase DZ sin conciliar</span>
- <span style="background-color:{COLOR_AMARILLO}; padding:2px 8px;">Amarillo: Resto de Sugerencias y alertas a verificar</span>
- <span style="background-color:{COLOR_ALERTA}; padding:2px 8px;">Naranja: Alerta de fecha o diferencia de valor</span>
- <span style="background-color:{COLOR_RECLASIFICAR}; padding:2px 8px;">Lila: Reclasificación de banco</span>
- <span style="background-color:{COLOR_PENDIENTE}; padding:2px 8px; border:1px solid #ccc;">Blanco: Pendiente general</span>
""", unsafe_allow_html=True)

            c1_, c2_, c3_, c4_, c5_ = st.columns(5)
            c1_.metric("Conciliadas (Azul)", total_conciliadas)
            c2_.metric("Verde (DZ 40 No Conciliado)", total_verde)
            c3_.metric("Alertas/Reclasificar", total_alertas + total_reclasificar)
            c4_.metric("Sugerencias (Amarillo)", total_sugerencias)
            c5_.metric("Pendientes (Blanco)", total_pendientes)

            if filas_excluidas > 0: st.warning(f"⚠️ Se excluyeron {filas_excluidas} filas vacías/totales.")

            st.download_button(label="📥 Descargar Excel con Resultados", data=output.getvalue(), file_name="Conciliacion completa V24 Color Exclusivo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error técnico detectado: {e}")
