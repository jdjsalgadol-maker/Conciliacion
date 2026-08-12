# app_conciliacion_v25_reglas_clm.py
#
# REESCRITURA COMPLETA DEL MOTOR segun especificacion CLM entregada por
# el usuario. Sustituye a v20-v24. Cada bloque de codigo esta comentado
# con el numero de regla que implementa.
#
# =========================================================================
# MAPEO FIJO DE COLUMNAS (CLM = nombre de columna en el prompt)
# =========================================================================
#   A -> AsignaciÃ³n
#   B -> NÂº documento
#   C -> Clase de documento (DZ / CB / IP)
#   D -> Fecha del periodo / Fe.contabilizaciÃ³n (SOLO PERIODO, NO concilia)
#   F -> Fecha valor (Fecha valor) -> ES LA FECHA PRINCIPAL DE CONCILIACION
#   G -> Clave contabiliz. (40 = dÃ©bito/legalizaciÃ³n, 50 = ingreso/cargue)
#   H -> Referencia
#   I -> Importe en moneda local
#   K -> Texto
#   Banco -> Clave referencia 3
#   M -> Estado_Conciliacion (estado + alertas, se escribe en este campo)
#   O -> Candidatos_Conciliacion (candidatos que el usuario puede filtrar)
#
# =========================================================================

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime

st.set_page_config(page_title="ConciliaciÃ³n Integral CLM", layout="wide")
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

st.title("ðŸ¦ ConciliaciÃ³n Automatizada â€” Motor CLM v25 ðŸ¤–")
st.write("Sube tu archivo consolidado.")
st.caption(
    "Motor reescrito segÃºn especificaciÃ³n CLM: A=AsignaciÃ³n, B=NÂº doc, C=Clase doc, "
    "D=Fecha periodo (informativa), F=Fecha valor (PRINCIPAL, debe coincidir), "
    "G=Clave (40 dÃ©bito/legalizaciÃ³n, 50 ingreso/cargue), H=Referencia, I=Importe, "
    "K=Texto, M=Estado, O=Candidatos."
)

with st.expander("âš™ï¸ ParÃ¡metros de tolerancia"):
    TOPE_DIAS_ALERTA = st.slider(
        "DÃ­as mÃ¡ximos de diferencia de fecha F para alerta (Regla 7, tope fijo 4)",
        1, 4, 4
    )
    tol_valor_purpura = st.number_input(
        "Diferencia mÃ¡xima de valor ($) para alerta MORADA (Regla morado, tope 500)",
        min_value=1, value=500, step=50, max_value=500
    )
    tol_valor_abs_general = st.number_input(
        "Diferencia absoluta general de valor para alertar (mÃ¡s allÃ¡ del morado) ($)",
        min_value=1, value=5000, step=100
    )
    tol_valor_pct_general = st.number_input(
        "Diferencia relativa general de valor para alertar (%)",
        min_value=0.01, value=0.5, step=0.01
    ) / 100
    multiplo_redondo = st.selectbox("MÃºltiplo para valor 'redondo' (alta ambigÃ¼edad)", [50000, 100000], index=1)

COLOR_AZUL = "#C5D9F1"      # Conciliado (todas las reglas cumplen)
COLOR_VERDE = "#A9D18E"     # Sugerencia / DZ multiposiciÃ³n / sectorizaciÃ³n / IP%
COLOR_SALMON = "#F5B7A1"    # Diferencia de fecha (hasta 4 dÃ­as) - Regla 7
COLOR_MORADO = "#C39BD3"    # Diferencia de valor mÃ¡x $500 - Regla morado
COLOR_DURAZNO = "#FAD7A0"   # ReclasificaciÃ³n de banco - Regla 6
COLOR_BLANCO = "#FFFFFF"    # Pendiente

archivo_subido = st.file_uploader("Selecciona el archivo de Excel o CSV", type=['xlsx', 'csv'])

if archivo_subido is not None:
    try:
        with st.spinner("Ejecutando motor de reglas CLM..."):

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
            col_A = 'AsignaciÃ³n' if 'AsignaciÃ³n' in df.columns else 'AsignaiÃ³n'
            col_B = 'NÂº documento' if 'NÂº documento' in df.columns else 'NÂº doc.'
            col_C = 'Clase de documento' if 'Clase de documento' in df.columns else ('Clase doc.' if 'Clase doc.' in df.columns else None)
            col_D = 'Fe.contabilizaciÃ³n' if 'Fe.contabilizaciÃ³n' in df.columns else ('Fecha de documento' if 'Fecha de documento' in df.columns else None)
            col_F = 'Fecha valor' if 'Fecha valor' in df.columns else 'Fe-valor'
            col_G = 'Clave contabiliz.' if 'Clave contabiliz.' in df.columns else 'CT'
            col_H = 'Referencia'
            col_I = 'Importe en moneda local' if 'Importe en moneda local' in df.columns else 'Importe en ML'
            col_K = 'Texto' if 'Texto' in df.columns else None
            col_novedad = 'novedad' if 'novedad' in df.columns else ('Novedad' if 'Novedad' in df.columns else None)
            col_banco = 'Clave referencia 3'

            requeridas = [col_H, col_G, col_F, col_I, col_banco, col_B]
            faltantes = [c for c in requeridas if c not in df.columns]
            if faltantes:
                st.error(f"No se encontraron estas columnas obligatorias: {faltantes}")
                st.stop()
            if col_A not in df.columns:
                st.error("No se encontrÃ³ la columna AsignaciÃ³n (A).")
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
            df['ID_Linea'] = df.index  # identificador Ãºnico por LINEA (no por documento)

            df[col_G] = df[col_G].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df[col_banco] = df[col_banco].astype(str).str.strip()
            df[col_I] = pd.to_numeric(df[col_I], errors='coerce').fillna(0)
            df['Abs_I'] = df[col_I].abs()

            df['Fecha_F'] = pd.to_datetime(df[col_F], errors='coerce')
            df[col_F] = df['Fecha_F'].dt.date

            if col_D:
                df['Fecha_D'] = pd.to_datetime(df[col_D], errors='coerce')
                df['Periodo_D'] = df['Fecha_D'].dt.to_period('M').astype(str)
                df.loc[df['Fecha_D'].isna(), 'Periodo_D'] = 'SIN_FECHA_D'
            else:
                df['Periodo_D'] = 'SIN_FECHA_D'

            df['Estado_Conciliacion'] = 'Pendiente'   # columna M
            df['Comentario'] = ''
            df['Candidatos_Conciliacion'] = ''        # columna O

            # Posiciones repetidas del mismo B (Regla 4)
            df['B_Repite'] = df.groupby(col_B)[col_B].transform('count') > 1

            # =====================================================
            # 4. SECTORIZACIÃ“N (Regla 2)
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

                if h_val in mapeo_referencias_dist:
                    return mapeo_referencias_dist[h_val]

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
                if 'NEQUI' in texto:
                    return True
                val_a = str(row.get(col_A, '')).strip().upper()
                if val_a == 'T' or val_a.startswith('T-') or val_a.startswith('T/') or val_a == '/':
                    return True
                return False

            def limpiar_numero(v):
                if pd.isna(v): return ''
                t = re.sub(r'\.0$', '', str(v).strip())
                m = re.findall(r'\d+', t)
                return m[0] if m else ''

            df['Es_Nequi'] = df.apply(es_nequi, axis=1)
            df['H_Limpia'] = df[col_H].apply(limpiar_numero)
            df['A_Limpia'] = df[col_A].apply(limpiar_numero)

            # =====================================================
            # FUNCIONES AUXILIARES
            # =====================================================
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
                if not indices:
                    return
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

            usados = set()

            def gate_seguridad(id40, id50, exigir_importe_exacto=True, tolerancia_valor=None):
                """
                Verifica reglas transversales: Banco, Sector, Importe y Fecha.
                Se aplica estrictamente el TOPE_DIAS_ALERTA (mÃ¡ximo 4 dÃ­as por defecto)
                para prevenir errores contables.
                """
                ra = df.loc[df['ID_Linea'] == id40].iloc[0]
                rb = df.loc[df['ID_Linea'] == id50].iloc[0]

                resultado = {
                    'ok': False, 'motivo': '', 'dif_dias': None,
                    'mismo_banco': False, 'banco_a': '', 'banco_b': '',
                    'dif_valor': None, 'pct_valor': None,
                    'mismo_sector': True, 'es_nequi': False
                }

                fa, fb = ra['Fecha_F'], rb['Fecha_F']
                if pd.isna(fa) or pd.isna(fb):
                    resultado['motivo'] = "Fecha F invÃ¡lida"
                    return resultado
                dif_dias = abs((fa - fb).days)
                resultado['dif_dias'] = dif_dias
                
                # BARRERA DE SEGURIDAD ESTRICTA: Ningun cruce puede exceder el TOPE
                if dif_dias > TOPE_DIAS_ALERTA:
                    resultado['motivo'] = f"Diferencia de fecha F fuera de rango ({dif_dias} dÃ­as > {TOPE_DIAS_ALERTA})"
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
                """
                Aplica el resultado del gate_seguridad y asigna el ESTADO + COLOR correctos.
                """
                res = gate_seguridad(id40, id50, exigir_importe_exacto=True, tolerancia_valor=tol_valor_purpura)
                if not res['ok']:
                    return False, res['motivo']

                texto_candidatos = f"{formato_linea(id40)} | {formato_linea(id50)}"
                partes_comentario = [base_txt]
                estado_final = 'Conciliado - Cumple todas las reglas'

                if not res['mismo_banco']:
                    estado_final = 'ReclasificaciÃ³n de banco'
                    partes_comentario.append(
                        f"ReclasificaciÃ³n de banco: registrado en '{res['banco_a']}'; banco esperado '{res['banco_b']}'."
                    )
                elif res['dif_dias'] and res['dif_dias'] > 0:
                    estado_final = 'Diferencia de fecha'
                    partes_comentario.append(f"Diferencia de fecha: F40 vs F50 difieren {res['dif_dias']} dÃ­a(s) (tope {TOPE_DIAS_ALERTA}).")
                elif res['dif_valor'] and res['dif_valor'] > 0:
                    estado_final = 'Diferencia de valor'
                    partes_comentario.append(
                        f"Diferencia de valor: dif=${res['dif_valor']:,.0f} ({res['pct_valor']*100:.2f}%). "
                        f"{'Dentro del tope de $' + str(int(tol_valor_purpura)) if res['dif_valor'] <= tol_valor_purpura else 'EXCEDE el tope sugerido de $' + str(int(tol_valor_purpura)) + ', revisar con prioridad'}."
                    )
                else:
                    estado_final = 'Conciliado - Cumple todas las reglas'

                if res['es_nequi']:
                    partes_comentario.append("[NEQUI: verificar manual]")

                comentario_final = " ".join(partes_comentario)

                for idx in (id40, id50):
                    # FIX: forzar=True permite sobrescribir los estados de "Sugerencia - Regla 8 Nequi" 
                    # si una regla fuerte posterior confirma el emparejamiento.
                    escribir_estado([idx], estado_final, forzar=True)
                    escribir_candidatos(idx, texto_candidatos)
                    # Usar append=False limpia los comentarios residuales de sugerencias pasadas
                    escribir_comentario(idx, comentario_final, append=False)

                return True, estado_final

            # =====================================================
            # 5. REGLA 3 â€” DOCUMENTOS IP (puntos de venta)
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
                        if not ip_ids or not cb_ids:
                            return
                        usados.update(ip_ids + cb_ids)
                        texto_cand = " | ".join(formato_linea(i) for i in ip_ids + cb_ids)
                        if es_exacto:
                            estado = 'Conciliado - Cruce mÃºltiple IP/CB (Regla 3)'
                            ind_ip_exacto.update(ip_ids + cb_ids)
                            txt = f"Cruce mÃºltiple homologado ({len(ip_ids)} IP = {len(cb_ids)} CB). Ref. homologada: {rh}."
                        else:
                            estado = 'Sugerencia - Cruce mÃºltiple IP/CB con diferencia de valor'
                            ind_ip_tolerancia.update(ip_ids + cb_ids)
                            txt = (
                                f"Sugerencia IP/CB con diferencia de valor "
                                f"(${fila['DifV']:,.0f} / {fila['Pct']*100:.2f}%). "
                                f"Suma {len(ip_ids)} IP vs {len(cb_ids)} CB. Ref. homologada: {rh}."
                            )
                        for idx in ip_ids + cb_ids:
                            df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = estado
                            df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                            otros = resumen_docs(sub_cb) if idx in ip_ids else resumen_docs(sub_ip)
                            df.loc[df['ID_Linea'] == idx, 'Comentario'] = f"{txt} Docs relacionados: {otros}"

                    for _, fila in exactos.iterrows():
                        procesar_grupo_ip(fila, es_exacto=True)
                    for _, fila in con_tol.iterrows():
                        procesar_grupo_ip(fila, es_exacto=False)

                ip_pendiente_zona = df[
                    (df[col_C].astype(str).str.upper() == 'IP') &
                    (~df['ID_Linea'].isin(usados)) &
                    (df['Sector'] != 'Sin clasificar')
                ]
                if not ip_pendiente_zona.empty:
                    cb_disponible_zona = df[
                        (df[col_G] == '50') &
                        (df[col_C].astype(str).str.upper() != 'IP') &
                        (~df['ID_Linea'].isin(usados))
                    ]
                    for (banco_z, fecha_z, sector_z), grupo_ip in ip_pendiente_zona.groupby([col_banco, 'Fecha_F', 'Sector']):
                        grupo_ip = grupo_ip[~grupo_ip['ID_Linea'].isin(usados)]
                        if grupo_ip.empty:
                            continue
                        candidatos_cb = cb_disponible_zona[
                            (cb_disponible_zona[col_banco] == banco_z) &
                            (cb_disponible_zona['Fecha_F'] == fecha_z) &
                            (cb_disponible_zona['Sector'] == sector_z) &
                            (~cb_disponible_zona['ID_Linea'].isin(usados))
                        ]
                        if candidatos_cb.empty:
                            continue
                        for _, fila_ip in grupo_ip.iterrows():
                            if fila_ip['ID_Linea'] in usados:
                                continue
                            match_importe = candidatos_cb[
                                (~candidatos_cb['ID_Linea'].isin(usados)) &
                                (candidatos_cb['Abs_I'] == fila_ip['Abs_I'])
                            ]
                            docs_candidatos = resumen_docs(match_importe) if len(match_importe) else ''
                            if len(match_importe) == 1:
                                id_cb = match_importe.iloc[0]['ID_Linea']
                                texto_cand = f"{formato_linea(fila_ip['ID_Linea'])} | {formato_linea(id_cb)}"
                                for idx in (fila_ip['ID_Linea'], id_cb):
                                    df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Conciliado - IP por Zona (Regla 9)'
                                    df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                                    df.loc[df['ID_Linea'] == idx, 'Comentario'] = f"IP conciliado por misma Zona ({sector_z}) e importe exacto."
                                usados.update([fila_ip['ID_Linea'], id_cb])
                            elif len(match_importe) > 1:
                                escribir_estado([fila_ip['ID_Linea']], 'Sugerencia - IP por Zona (varios candidatos)', forzar=False)
                                escribir_candidatos(fila_ip['ID_Linea'], f"Zona {sector_z}: candidatos CB con mismo importe: {docs_candidatos}")
                                escribir_comentario(fila_ip['ID_Linea'], f"IP con misma Zona ({sector_z}) pero varios CB candidatos con el mismo importe exacto.", append=False)

                ip_sin_resolver = df[(df[col_C].astype(str).str.upper() == 'IP') & (~df['ID_Linea'].isin(usados))]
                for idl in ip_sin_resolver['ID_Linea']:
                    escribir_comentario(idl, "PDV (IP): requiere referencia homologada de base de datos o coincidencia por Zona (Regla 9).", append=False)

            # =====================================================
            # 5B. REGLA 8 â€” NEQUI POR TOTALES Y FIFO (secciÃ³n 11 de la guÃ­a)
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
                    if grupo_dz.empty:
                        continue
                    grupo_cb = df_cb_disponible[
                        (df_cb_disponible[col_banco] == banco_g) &
                        (df_cb_disponible['Fecha_F'] == fecha_g) &
                        (~df_cb_disponible['ID_Linea'].isin(usados))
                    ]
                    if grupo_cb.empty:
                        continue

                    n_dz, n_cb = len(grupo_dz), len(grupo_cb)
                    total_dz = round(grupo_dz['Abs_I'].sum(), 2)
                    total_cb = round(grupo_cb['Abs_I'].sum(), 2)

                    dz_ord = grupo_dz.sort_values(col_B).reset_index(drop=True)
                    cb_ord = grupo_cb.sort_values(col_B).reset_index(drop=True)
                    n_parejas = min(n_dz, n_cb)

                    if n_dz == n_cb and total_dz == total_cb:
                        texto_grupo = f"total DZ=${total_dz:,.0f}; total CB=${total_cb:,.0f}; cruce FIFO por B ({n_dz} lineas)."
                        for i in range(n_parejas):
                            id40 = dz_ord.iloc[i]['ID_Linea']
                            id50 = cb_ord.iloc[i]['ID_Linea']
                            texto_cand = f"{formato_linea(id40)} | {formato_linea(id50)}"
                            for idx in (id40, id50):
                                df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Conciliado - Regla 8 Nequi (total y FIFO)'
                                df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                                df.loc[df['ID_Linea'] == idx, 'Comentario'] = f"Regla 8 Nequi: {texto_grupo}"
                            ind_nequi8_azul.update([id40, id50])
                        usados.update(dz_ord['ID_Linea'].tolist() + cb_ord['ID_Linea'].tolist())
                    else:
                        diff_total = abs(total_dz - total_cb)
                        motivo = (
                            f"cantidad DZ={n_dz} vs CB={n_cb}" if n_dz != n_cb
                            else f"totales distintos (DZ=${total_dz:,.0f} vs CB=${total_cb:,.0f}, dif=${diff_total:,.0f})"
                        )
                        docs_dz = resumen_docs(dz_ord)
                        docs_cb = resumen_docs(cb_ord)
                        texto_cand = f"Grupo NEQUI {banco_g} {fecha_g}: DZ candidatos: {docs_dz} || CB candidatos: {docs_cb}"
                        comentario = (
                            f"Regla 8 Nequi: grupo NO cuadra exacto ({motivo}). "
                            f"No se concilia automÃ¡tico; requiere revisiÃ³n manual del grupo completo."
                        )
                        for _, fila in pd.concat([dz_ord, cb_ord]).iterrows():
                            idl = fila['ID_Linea']
                            escribir_estado([idl], 'Sugerencia - Regla 8 Nequi (revisar total de grupo)', forzar=False)
                            if df.loc[df['ID_Linea'] == idl, 'Candidatos_Conciliacion'].iloc[0] == '':
                                escribir_candidatos(idl, texto_cand)
                            escribir_comentario(idl, comentario, append=False)
                            ind_nequi8_sugerencia.add(idl)

            # =====================================================
            # 6. REGLA 1 â€” A debe coincidir con H (exacto)
            # FIX: Se remueve filtro que bloqueaba a IP, permitiendo cruces 1a1
            # =====================================================
            df_40 = df[(df[col_G] == '40')].copy()
            df_50 = df[(df[col_G] == '50')].copy()

            def emparejar_1a1_por_llave(sub40, sub50, llave40, llave50, base_txt):
                s40 = sub40[~sub40['ID_Linea'].isin(usados)].copy()
                s50 = sub50[~sub50['ID_Linea'].isin(usados)].copy()
                if s40.empty or s50.empty:
                    return
                s40['_pos'] = s40.groupby(llave40).cumcount()
                s50['_pos'] = s50.groupby(llave50).cumcount()
                merged = pd.merge(
                    s40, s50, left_on=llave40 + ['_pos'], right_on=llave50 + ['_pos'],
                    suffixes=('_40', '_50')
                )
                for _, r in merged.iterrows():
                    id40, id50 = r['ID_Linea_40'], r['ID_Linea_50']
                    if id40 in usados or id50 in usados:
                        continue
                    ok, estado_final = clasificar_y_registrar(id40, id50, base_txt)
                    if ok:
                        usados.update([id40, id50])

            emparejar_1a1_por_llave(
                df_40, df_50,
                [col_banco, 'Abs_I', col_A], [col_banco, 'Abs_I', col_H],
                "Regla 1: AsignaciÃ³n (A) coincide exacta con Referencia (H)."
            )

            df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
            df_50 = df_50[~df_50['ID_Linea'].isin(usados)]
            emparejar_1a1_por_llave(
                df_40[df_40['A_Limpia'] != ''], df_50[df_50['H_Limpia'] != ''],
                [col_banco, 'Abs_I', 'A_Limpia'], [col_banco, 'Abs_I', 'H_Limpia'],
                "Regla 1 (limpia): AsignaciÃ³n limpia coincide con Referencia limpia (ej. E3110â†’3110)."
            )

            # =====================================================
            # 6.2 REGLA 1 FLEX: RELACION PARCIAL LIMPIA
            # Resuelve Ej: A="T-9006235207" contiene a H="900623520"
            # Ahora respeta estrictamente el lÃ­mite de <= 4 dÃ­as.
            # =====================================================
            def emparejar_parcial_limpia(sub40, sub50, base_txt):
                s40 = sub40[~sub40['ID_Linea'].isin(usados)].copy()
                s50 = sub50[~sub50['ID_Linea'].isin(usados)].copy()
                if s40.empty or s50.empty: return
                for _, r40 in s40.iterrows():
                    id40 = r40['ID_Linea']
                    if id40 in usados: continue
                    cands = s50[(~s50['ID_Linea'].isin(usados)) & (s50[col_banco] == r40[col_banco]) & (s50['Abs_I'] == r40['Abs_I'])]
                    if cands.empty: continue
                    a_str = str(r40['A_Limpia']).strip()
                    if len(a_str) < 6: continue
                    
                    for _, r50 in cands.iterrows():
                        id50 = r50['ID_Linea']
                        if id50 in usados: continue
                        h_str = str(r50['H_Limpia']).strip()
                        if len(h_str) < 6: continue
                        
                        if (a_str in h_str) or (h_str in a_str):
                            # Ya no se fuerza el max_dias=999. Respeta el gate_seguridad de 4 dÃ­as.
                            ok, _ = clasificar_y_registrar(id40, id50, base_txt)
                            if ok:
                                usados.update([id40, id50])
                                break
            
            df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
            df_50 = df_50[~df_50['ID_Linea'].isin(usados)]
            emparejar_parcial_limpia(df_40, df_50, "Regla 1 (Flex): La asignaciÃ³n se relaciona estrechamente con la Referencia.")

            # =====================================================
            # 6.5 REGLA 6 EXPLÃCITA â€” RECLASIFICACIÃ“N DE BANCO
            # =====================================================
            df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
            df_50 = df_50[~df_50['ID_Linea'].isin(usados)]

            def emparejar_reclasificacion(sub40, sub50, llave40, llave50, base_txt):
                s40 = sub40[~sub40['ID_Linea'].isin(usados)].copy()
                s50 = sub50[~sub50['ID_Linea'].isin(usados)].copy()
                if s40.empty or s50.empty:
                    return
                s40['_pos'] = s40.groupby(llave40).cumcount()
                s50['_pos'] = s50.groupby(llave50).cumcount()
                merged = pd.merge(
                    s40, s50, left_on=llave40 + ['_pos'], right_on=llave50 + ['_pos'],
                    suffixes=('_40', '_50')
                )
                for _, r in merged.iterrows():
                    id40, id50 = r['ID_Linea_40'], r['ID_Linea_50']
                    if id40 in usados or id50 in usados:
                        continue
                    ra = df.loc[df['ID_Linea'] == id40].iloc[0]
                    rb = df.loc[df['ID_Linea'] == id50].iloc[0]
                    if str(ra[col_banco]).strip() == str(rb[col_banco]).strip():
                        continue  # esto ya lo cubriÃ³ Regla 1; aquÃ­ solo nos interesa banco distinto
                    ok, estado_final = clasificar_y_registrar(id40, id50, base_txt)
                    if ok:
                        usados.update([id40, id50])

            emparejar_reclasificacion(
                df_40, df_50, ['Abs_I', col_A], ['Abs_I', col_H],
                "Regla 6: AsignaciÃ³n (A) coincide con Referencia (H), pero el banco registrado difiere."
            )

            emparejar_reclasificacion(
                df_40[df_40['A_Limpia'] != ''], df_50[df_50['H_Limpia'] != ''],
                ['Abs_I', 'A_Limpia'], ['Abs_I', 'H_Limpia'],
                "Regla 6 (limpia): AsignaciÃ³n limpia coincide con Referencia limpia, pero el banco registrado difiere."
            )

            # =====================================================
            # 7. REGLA 2 â€” SECTORIZACIÃ“N (cuando A != H)
            # FIX: Permite el cruce FIFO si hay mÃºltiples registros en misma fecha (sin ambigÃ¼edad que bloquee)
            # =====================================================
            df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
            df_50 = df_50[~df_50['ID_Linea'].isin(usados)]

            d40_sect = df_40[df_40['Sector'] != 'Sin clasificar']
            d50_sect = df_50[df_50['Sector'] != 'Sin clasificar']

            if not d40_sect.empty and not d50_sect.empty:
                for grp, sub40 in d40_sect.groupby([col_banco, 'Abs_I', 'Sector']):
                    b, imp, sector = grp
                    sub50 = d50_sect[
                        (d50_sect[col_banco] == b) & (d50_sect['Abs_I'] == imp) & (d50_sect['Sector'] == sector)
                    ]
                    if sub50.empty:
                        continue
                    
                    s40_ord = sub40[~sub40['ID_Linea'].isin(usados)].sort_values('Fecha_F')
                    s50_ord = sub50[~sub50['ID_Linea'].isin(usados)].sort_values('Fecha_F')
                    if s40_ord.empty or s50_ord.empty:
                        continue

                    # Se eliminÃ³ la validaciÃ³n "es_unico_sin_ambiguedad". Ahora empareja directo por orden.
                    n_pares = min(len(s40_ord), len(s50_ord))
                    for i in range(n_pares):
                        r40, r50 = s40_ord.iloc[i], s50_ord.iloc[i]
                        id40, id50 = r40['ID_Linea'], r50['ID_Linea']
                        if id40 in usados or id50 in usados:
                            continue
                        ok, _ = clasificar_y_registrar(id40, id50, f"SectorizaciÃ³n ({sector})")
                        if ok: usados.update([id40, id50])

                    # Sobrantes desbalanceados -> listar candidatos (Regla 4)
                    rem40 = s40_ord[~s40_ord['ID_Linea'].isin(usados)]
                    rem50 = s50_ord[~s50_ord['ID_Linea'].isin(usados)]
                    if not rem40.empty or not rem50.empty:
                        docs50_txt = resumen_docs(rem50)
                        docs40_txt = resumen_docs(rem40)
                        for _, r in rem40.iterrows():
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Estado_Conciliacion'] = f"Sugerencia - SectorizaciÃ³n desbalanceada ({sector})"
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Candidatos_Conciliacion'] = f"{formato_linea(r['ID_Linea'])} | Candidatos posibles: {docs50_txt}"
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Comentario'] = f"Sector '{sector}' desbalanceado ({len(rem40)} vs {len(rem50)}). CrÃ©ditos candidatos: {docs50_txt}"
                        for _, r in rem50.iterrows():
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Estado_Conciliacion'] = f"Sugerencia - SectorizaciÃ³n desbalanceada ({sector})"
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Candidatos_Conciliacion'] = f"{formato_linea(r['ID_Linea'])} | Candidatos posibles: {docs40_txt}"
                            df.loc[df['ID_Linea'] == r['ID_Linea'], 'Comentario'] = f"Sector '{sector}' desbalanceado ({len(rem50)} vs {len(rem40)}). DÃ©bitos candidatos: {docs40_txt}"

            # =====================================================
            # 7B. SUGERENCIA SECTORIZACIÃ“N CON DIFERENCIA DE VALOR 
            # (Ej. Alerta faltante 50.000 u otras diferencias que comparten sector y fecha)
            # =====================================================
            rem40_sect = d40_sect[~d40_sect['ID_Linea'].isin(usados)]
            rem50_sect = d50_sect[~d50_sect['ID_Linea'].isin(usados)]

            if not rem40_sect.empty and not rem50_sect.empty:
                for (b, f, sec), g40 in rem40_sect.groupby([col_banco, 'Fecha_F', 'Sector']):
                    g50 = rem50_sect[(rem50_sect[col_banco] == b) & (rem50_sect['Fecha_F'] == f) & (rem50_sect['Sector'] == sec)]
                    if g50.empty: continue
                    
                    for _, r40 in g40.iterrows():
                        id40 = r40['ID_Linea']
                        if id40 in usados: continue
                        g50_disp = g50[~g50['ID_Linea'].isin(usados)].copy()
                        if g50_disp.empty: continue
                        
                        g50_disp['_dif'] = (g50_disp['Abs_I'] - r40['Abs_I']).abs()
                        best_50 = g50_disp.sort_values('_dif').head(3)
                        
                        if not best_50.empty:
                            docs50_txt = resumen_docs(best_50)
                            dif_min = best_50.iloc[0]['_dif']
                            df.loc[df['ID_Linea'] == id40, 'Estado_Conciliacion'] = f"Sugerencia - SectorizaciÃ³n dif. valor"
                            df.loc[df['ID_Linea'] == id40, 'Candidatos_Conciliacion'] = f"{formato_linea(id40)} | Candidatos posibles: {docs50_txt}"
                            df.loc[df['ID_Linea'] == id40, 'Comentario'] = f"Alerta Sector '{sec}': misma fecha pero difiere en valor (dif aprox ${dif_min:,.0f}). Candidatos: {docs50_txt}"

            # =====================================================
            # 8. EXCEPCIÃ“N NEQUI (A = "Nequi", C = DZ)
            # =====================================================
            df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
            df_50 = df_50[~df_50['ID_Linea'].isin(usados)]

            df_nequi_40 = df_40[df_40['Es_Nequi'] == True]

            for _, r40 in df_nequi_40.iterrows():
                id40 = r40['ID_Linea']
                if id40 in usados:
                    continue
                candidatos = df_50[
                    (df_50[col_banco] == r40[col_banco]) &
                    (~df_50['ID_Linea'].isin(usados))
                ].copy()
                if candidatos.empty:
                    continue

                if r40['Sector'] != 'Sin clasificar':
                    cf = candidatos[
                        (candidatos['Sector'] == r40['Sector']) | (candidatos['Sector'] == 'Sin clasificar')
                    ]
                    if not cf.empty:
                        candidatos = cf

                # Calcular la diferencia de dÃ­as
                candidatos['_dif_dias'] = candidatos['Fecha_F'].apply(
                    lambda x: abs((x - r40['Fecha_F']).days) if pd.notna(x) and pd.notna(r40['Fecha_F']) else 999
                )
                
                # APLICAMOS EL FILTRO ESTRICTO DE FECHAS ANTES DE EMPAREJAR
                candidatos = candidatos[candidatos['_dif_dias'] <= TOPE_DIAS_ALERTA]
                
                if candidatos.empty:
                    continue

                # Importes exactos (Ya filtrados por <= 4 dÃ­as)
                exactos = candidatos[candidatos['Abs_I'] == r40['Abs_I']].sort_values('_dif_dias')
                if not exactos.empty:
                    id50 = exactos.iloc[0]['ID_Linea']
                    ok, _ = clasificar_y_registrar(id40, id50, "ExcepciÃ³n Nequi (cruce importe exacto)")
                    if ok: usados.update([id40, id50])
                    continue
                    
                # Si no hay exacto Ãºnico, buscar dentro de tolerancia $500 (Regla morada)
                candidatos['_dif_val'] = (candidatos['Abs_I'] - r40['Abs_I']).abs()
                con_tol = candidatos[candidatos['_dif_val'] <= tol_valor_purpura].sort_values(['_dif_val', '_dif_dias'])
                if not con_tol.empty:
                    id50 = con_tol.iloc[0]['ID_Linea']
                    ok, _ = clasificar_y_registrar(id40, id50, "ExcepciÃ³n Nequi (con diferencia de valor)")
                    if ok: usados.update([id40, id50])
                    continue

            # =====================================================
            # 9. REGLA 4 â€” DOCUMENTOS DZ CON POSICIONES MÃšLTIPLES
            # =====================================================
            df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
            df_50 = df_50[~df_50['ID_Linea'].isin(usados)]

            dz_repetidos = df_40[(df_40['B_Repite'] == True) & (df_40['Candidatos_Conciliacion'] == '')]
            for b_doc, grupo in dz_repetidos.groupby(col_B):
                for _, linea in grupo.iterrows():
                    idl = linea['ID_Linea']
                    if idl in usados:
                        continue
                    candidatos = df_50[
                        (df_50[col_banco] == linea[col_banco]) &
                        (~df_50['ID_Linea'].isin(usados))
                    ].copy()
                    if candidatos.empty:
                        continue
                    candidatos['_dif_dias'] = candidatos['Fecha_F'].apply(
                        lambda f: abs((f - linea['Fecha_F']).days) if pd.notna(f) and pd.notna(linea['Fecha_F']) else 999
                    )
                    candidatos = candidatos[candidatos['_dif_dias'] <= TOPE_DIAS_ALERTA]
                    if candidatos.empty:
                        continue
                    candidatos['_dif_valor'] = (candidatos['Abs_I'] - linea['Abs_I']).abs()
                    candidatos['_mismo_sector'] = (candidatos['Sector'] == linea['Sector']) & (linea['Sector'] != 'Sin clasificar')
                    candidatos['_mismo_A_H'] = candidatos[col_H].astype(str).str.strip() == str(linea[col_A]).strip()
                    candidatos_ordenados = candidatos.sort_values(
                        by=['_mismo_A_H', '_mismo_sector', '_dif_valor', '_dif_dias'],
                        ascending=[False, False, True, True]
                    )
                    top_candidatos = candidatos_ordenados.head(5)
                    lineas_candidatos = []
                    for rank, (_, c) in enumerate(top_candidatos.iterrows(), start=1):
                        lineas_candidatos.append(
                            f"{rank}. Doc={int(c[col_B])}, H={c[col_H]}, I=${c['Abs_I']:,.0f}, "
                            f"F={c[col_F]}, Sector={c['Sector']}, dif_valor=${c['_dif_valor']:,.0f}, dif_dias={int(c['_dif_dias'])}"
                        )
                    texto_cand_final = f"LÃ­nea 40: Doc={int(linea[col_B])}, I=${linea['Abs_I']:,.0f}, F={linea[col_F]} || Candidatos: " + " ; ".join(lineas_candidatos)
                    df.loc[df['ID_Linea'] == idl, 'Estado_Conciliacion'] = 'Sugerencia - DZ posiciones mÃºltiples (verificar)'
                    df.loc[df['ID_Linea'] == idl, 'Candidatos_Conciliacion'] = texto_cand_final
                    df.loc[df['ID_Linea'] == idl, 'Comentario'] = (
                        f"Documento {int(b_doc)} tiene mÃºltiples posiciones/importes distintos. "
                        f"Se listan hasta 5 candidatos ordenados por calidad de coincidencia en la columna O."
                    )

            # =====================================================
            # 10. ÃšLTIMO RECURSO: FIFO CONTROLADO
            # =====================================================
            df_40 = df_40[~df_40['ID_Linea'].isin(usados)]
            df_50 = df_50[~df_50['ID_Linea'].isin(usados)]

            pendientes_40 = df[
                (df['ID_Linea'].isin(df_40['ID_Linea'])) &
                (~df['ID_Linea'].isin(usados))
            ]
            pendientes_50 = df[
                (df['ID_Linea'].isin(df_50['ID_Linea'])) &
                (~df['ID_Linea'].isin(usados))
            ]

            ind_fifo_ok = set()
            ind_fifo_verde_dz = set()

            for grp, sub40 in pendientes_40.groupby([col_banco, 'Abs_I', col_F, 'Sector']):
                b, imp, f, sector = grp
                sub50 = pendientes_50[
                    (pendientes_50[col_banco] == b) & (pendientes_50['Abs_I'] == imp) &
                    (pendientes_50[col_F] == f) & (pendientes_50['Sector'] == sector)
                ]
                if sub50.empty:
                    continue
                s40_ord = sub40[~sub40['ID_Linea'].isin(usados)].sort_values('ID_Linea')
                s50_ord = sub50[~sub50['ID_Linea'].isin(usados)].sort_values('ID_Linea')
                n_pares = min(len(s40_ord), len(s50_ord))
                for i in range(n_pares):
                    id40, id50 = s40_ord.iloc[i]['ID_Linea'], s50_ord.iloc[i]['ID_Linea']
                    texto_cand = f"{formato_linea(id40)} | {formato_linea(id50)}"
                    for idx in (id40, id50):
                        df.loc[df['ID_Linea'] == idx, 'Estado_Conciliacion'] = 'Conciliado - FIFO controlado (Ãºltima instancia)'
                        df.loc[df['ID_Linea'] == idx, 'Candidatos_Conciliacion'] = texto_cand
                        df.loc[df['ID_Linea'] == idx, 'Comentario'] = 'Emparejado por FIFO controlado: misma fecha F, mismo sector, banco e importe.'
                    ind_fifo_ok.update([id40, id50])
                    usados.update([id40, id50])

                # Sobrantes DZ con B repetido y sin pareja -> Regla verde
                sobrantes40 = s40_ord[~s40_ord['ID_Linea'].isin(usados)]
                for _, fila in sobrantes40.iterrows():
                    idl = fila['ID_Linea']
                    if bool(fila['B_Repite']):
                        df.loc[df['ID_Linea'] == idl, 'Estado_Conciliacion'] = 'Sugerencia - DZ multiposiciÃ³n sin cruce (verificar)'
                        df.loc[df['ID_Linea'] == idl, 'Comentario'] = (
                            f"Documento {int(fila[col_B])} tiene varias posiciones y Ã©sta no encontrÃ³ pareja "
                            f"exacta. Requiere descarte manual (Regla verde)."
                        )
                        ind_fifo_verde_dz.add(idl)

            # =====================================================
            # 11. CIERRE â€” Lo que sigue Pendiente
            # =====================================================
            sin_p = df['Estado_Conciliacion'] == 'Pendiente'
            if usar_ipcb:
                es_ip = df[col_C].astype(str).str.upper() == 'IP'
                df.loc[sin_p & es_ip & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia - PDV (requiere referencia homologada)'
                df.loc[sin_p & ~es_ip & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia ni sugerencia que cumpla reglas de seguridad - revisiÃ³n manual completa'
            else:
                df.loc[sin_p & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia ni sugerencia que cumpla reglas de seguridad - revisiÃ³n manual completa'

            # =====================================================
            # 12. EXPORTACIÃ“N
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
                if est in ('pendiente', '', 'nan'):
                    return [f'background-color: {COLOR_BLANCO}; color: black'] * len(row)
                if 'reclasificaciÃ³n' in est:
                    return [f'background-color: {COLOR_DURAZNO}; color: black'] * len(row)
                if 'diferencia de fecha' in est:
                    return [f'background-color: {COLOR_SALMON}; color: black'] * len(row)
                if 'diferencia de valor' in est:
                    return [f'background-color: {COLOR_MORADO}; color: black'] * len(row)
                if 'dz multiposiciÃ³n' in est or 'dz posiciones mÃºltiples' in est or 'sectorizaciÃ³n dif' in est:
                    return [f'background-color: {COLOR_VERDE}; color: black'] * len(row)
                if 'conciliado' in est:
                    return [f'background-color: {COLOR_AZUL}; color: black'] * len(row)
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
                total_azul = int(df_final['Estado_Conciliacion'].str.contains('Conciliado', na=False).sum())
                total_verde = int(
                    df_final['Estado_Conciliacion'].str.contains('DZ multiposiciÃ³n|DZ posiciones mÃºltiples|SectorizaciÃ³n dif', na=False, regex=True).sum()
                )
                total_salmon = int(df_final['Estado_Conciliacion'].str.contains('Diferencia de fecha', na=False).sum())
                total_morado = int(df_final['Estado_Conciliacion'].str.contains('Diferencia de valor', na=False).sum())
                total_durazno = int(df_final['Estado_Conciliacion'].str.contains('ReclasificaciÃ³n', na=False).sum())
                total_pendiente = int((df_final['Estado_Conciliacion'] == 'Pendiente').sum())

                resumen = pd.DataFrame({
                    "MÃ©trica": [
                        "Fecha de procesamiento", "Total filas procesadas",
                        "Azul - Conciliado (todas las reglas)",
                        "Verde - Sugerencias (DZ mÃºltiples / SectorizaciÃ³n Valor)",
                        "SalmÃ³n - Diferencia de fecha (Regla 7)",
                        "Morado - Diferencia de valor mÃ¡x $500 (Regla morado)",
                        "Durazno - ReclasificaciÃ³n de banco (Regla 6)",
                        "Blanco - Pendiente sin evidencia",
                        "IP conciliado exacto (Regla 3)", "IP con % de diferencia",
                        "Regla 8 Nequi - Azul (total y FIFO exacto)",
                        "Regla 8 Nequi - Sugerencia (grupo no cuadra exacto)",
                        "FIFO controlado (Ãºltima instancia)", "DZ verde sin cruce",
                        "Filas excluidas (sin doc/clave)", "Filas con NÂº doc. repetido",
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

                df_nov = df_final[df_final[col_G] == '40'].copy()
                estados_alerta = ['Diferencia de fecha', 'Diferencia de valor', 'ReclasificaciÃ³n de banco']
                mask_alerta = df_nov['Estado_Conciliacion'].isin(estados_alerta)
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
            st.success("Â¡ConciliaciÃ³n completada con el motor de reglas CLM v25!")
            if not cuadre_ok:
                st.warning("âš ï¸ Revisa la pestaÃ±a DESCARTADAS, el total de filas no coincide.")
            for adv in advertencias:
                st.warning(f"âš ï¸ {adv}")

            st.markdown(f"""
**Leyenda de colores:**
- <span style="background-color:{COLOR_AZUL}; padding:2px 8px;">Azul: Conciliado â€” cumple todas las reglas (A=H, misma F, mismo banco, importe exacto, sector coherente)</span>
- <span style="background-color:{COLOR_VERDE}; padding:2px 8px;">Verde: Sugerencia â€” sectorizaciÃ³n, DZ multiposiciÃ³n, Nequi ambiguo, IP con % de diferencia</span>
- <span style="background-color:{COLOR_SALMON}; padding:2px 8px;">SalmÃ³n: Diferencia de fecha F (hasta {TOPE_DIAS_ALERTA} dÃ­as)</span>
- <span style="background-color:{COLOR_MORADO}; padding:2px 8px;">Morado: Diferencia de valor (mÃ¡x ${tol_valor_purpura:.0f})</span>
- <span style="background-color:{COLOR_DURAZNO}; padding:2px 8px;">Durazno: ReclasificaciÃ³n de banco</span>
- <span style="background-color:{COLOR_BLANCO}; padding:2px 8px; border:1px solid #ccc;">Blanco: Pendiente sin evidencia suficiente</span>

**Reglas clave aplicadas:**
- La fecha D (periodo) es SOLO informativa. La fecha F es la ÃšNICA que valida conciliaciÃ³n.
- Un documento B con G=40 se empareja con un Ãºnico documento B con G=50 (Regla 1).
- Los IP solo concilian por referencia homologada de base de datos (Regla 3), nunca por fecha+importe simple.
- Los "Nequi" (Regla 3, excepciÃ³n) requieren candidato Ãºnico, mismo banco/fecha, sector coherente e importe exacto o diferencia â‰¤ ${tol_valor_purpura:.0f}.
- Documentos DZ repetidos (Regla 4) muestran hasta 5 candidatos ordenados en la columna O.
- El FIFO controlado es el ÃšLTIMO recurso: solo actÃºa sobre filas 100% vÃ­rgenes (sin estado, comentario ni candidatos previos), evitando el bug de versiones anteriores que sobrescribÃ­a bloqueos de seguridad.
""", unsafe_allow_html=True)

            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("Azul", total_azul)
            c2.metric("Verde", total_verde)
            c3.metric("SalmÃ³n", total_salmon)
            c4.metric("Morado", total_morado)
            c5.metric("Durazno", total_durazno)
            c6.metric("Pendiente", total_pendiente)

            if filas_excluidas > 0:
                st.warning(f"âš ï¸ Se excluyeron {filas_excluidas} filas vacÃ­as/totales.")

            st.download_button(
                label="ðŸ“¥ Descargar Excel con Resultados",
                data=output.getvalue(),
                file_name="Conciliacion_CLM_v25.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error tÃ©cnico detectado: {e}")
        st.exception(e)
