import streamlit as st
import pandas as pd
import numpy as np
import io
import re

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

st.title("🏦 Conciliación Integral Multibanco 🤖")
st.write(
    "Sube tu archivo consolidado. El sistema concilia de forma **conservadora**: "
    "solo marca como *Conciliado* (verde) lo que tiene evidencia inequívoca — referencia "
    "exacta, cruces múltiples de datáfonos (IP vs CB), sectorización por sede única, "
    "o cruce único sin ambigüedad. Todo lo ambiguo genera alertas de validación."
)

with st.expander("⚙️ Parámetros de tolerancia para sugerencias (alertas)"):
    tol_dias = st.slider("Días máximos de diferencia para alertar 'error de fecha' (dentro del mismo periodo)", 1, 15, 3)
    tol_valor_abs = st.number_input("Diferencia absoluta máxima de valor para alertar ($)", min_value=1, value=5000, step=100)
    tol_valor_pct = st.number_input("Diferencia relativa máxima de valor para alertar (%)", min_value=0.01, value=0.5, step=0.01) / 100
    multiplo_redondo = st.selectbox("Múltiplo para considerar un valor 'redondo' (alta ambigüedad)", [50000, 100000], index=1)

archivo_subido = st.file_uploader("Selecciona el archivo de Excel o CSV", type=['xlsx', 'csv'])

if archivo_subido is not None:
    try:
        with st.spinner("Leyendo archivo, unificando datos y aplicando reglas de conciliación..."):

            # =========================================================
            # 1. LECTURA A PRUEBA DE PESTAÑAS (Excel) O CSV
            # =========================================================
            if archivo_subido.name.lower().endswith('.csv'):
                df = pd.read_csv(archivo_subido)
            else:
                diccionario_hojas = pd.read_excel(archivo_subido, sheet_name=None)
                df = pd.concat(diccionario_hojas.values(), ignore_index=True)

            df.columns = df.columns.str.strip()

            # =========================================================
            # 2. MAPEO SEGURO Y DINÁMICO DE COLUMNAS
            # =========================================================
            col_asignacion = 'Asignación' if 'Asignación' in df.columns else 'Asignaión'
            col_referencia = 'Referencia'
            col_clave = 'Clave contabiliz.' if 'Clave contabiliz.' in df.columns else 'CT'
            col_fecha = 'Fecha valor' if 'Fecha valor' in df.columns else 'Fe-valor'
            col_importe = 'Importe en moneda local' if 'Importe en moneda local' in df.columns else 'Importe en ML'
            col_banco = 'Clave referencia 3'
            col_doc = 'Nº documento' if 'Nº documento' in df.columns else 'Nº doc.'
            col_texto = 'Texto' if 'Texto' in df.columns else None
            col_clase_doc = 'Clase de documento' if 'Clase de documento' in df.columns else 'Clase doc.'

            columnas_requeridas = [col_referencia, col_clave, col_fecha, col_importe, col_banco, col_doc, col_clase_doc]
            faltantes = [c for c in columnas_requeridas if c not in df.columns]
            if faltantes:
                st.error(f"No se encontraron estas columnas obligatorias en el archivo: {faltantes}")
                st.stop()

            # =========================================================
            # 3. AUTOCOMPLETADO DE BANCO POR CUENTA DE MAYOR
            # =========================================================
            mapeo_cuentas_banco = {
                "1110056101": "BANCO DE BOGOTA", "1110056201": "BANCO DAVIBANK S.A.",
                "1110056301": "BANCOLOMBIA S.A.", "1110056401": "BANCO CAJA SOCIAL S.",
                "1110056501": "BANCO DAVIVIENDA S.A", "1110056601": "BANCO BILBAO VIZCAYA",
                "1110056701": "BANCO AGRARIO DE COL", "1120055001": "BANCO COMERCIAL AV V",
                "1120055101": "BANCO DE OCCIDENTE", "1120055301": "BANCO GNB SUDAMERIS",
            }

            current_bank = None
            bancos_completados = []
            for _, row in df.iterrows():
                asig_val = str(row.get(col_asignacion, "")) if col_asignacion in df.columns else ""
                banco_val = row.get(col_banco, None)

                if "cuenta de mayor" in asig_val.lower():
                    match_cuenta = re.search(r'(\d{6,})', asig_val)
                    cuenta_num = match_cuenta.group(1) if match_cuenta else None
                    if cuenta_num:
                        current_bank = mapeo_cuentas_banco.get(cuenta_num, f"CUENTA {cuenta_num} (sin mapear)")

                if pd.notnull(banco_val) and str(banco_val).strip().lower() not in ("", "nan"):
                    current_bank = str(banco_val).strip()
                bancos_completados.append(current_bank)

            df[col_banco] = bancos_completados
            if col_asignacion in df.columns:
                df = df[~df[col_asignacion].astype(str).str.contains("cuenta de mayor", case=False, na=False)].copy()

            # =========================================================
            # 4. LIMPIEZA Y ORDENAMIENTO FIFO
            # =========================================================
            df[col_doc] = pd.to_numeric(df[col_doc], errors='coerce')
            filas_antes = len(df)
            filas_descartadas = df[df[col_doc].isna() | df[col_clave].isna()].copy()
            df = df.dropna(subset=[col_doc, col_clave]).reset_index(drop=True)
            filas_excluidas = filas_antes - len(df)

            df = df.sort_values(by=[col_doc], ascending=True).reset_index(drop=True)
            df['ID_Temp'] = df.index

            df[col_clave] = df[col_clave].astype(str).str.strip().str.replace('.0', '', regex=False)
            df[col_banco] = df[col_banco].astype(str).str.strip()
            df[col_importe] = pd.to_numeric(df[col_importe], errors='coerce').fillna(0)
            df['Abs_Importe'] = df[col_importe].abs()

            df['Fecha_Calc'] = pd.to_datetime(df[col_fecha], errors='coerce')
            df[col_fecha] = df['Fecha_Calc'].dt.date
            df['Estado_Conciliacion'] = 'Pendiente'
            df['Comentario'] = ''

            # =========================================================
            # 4B. DICCIONARIOS DE REFERENCIA (DATÁFONOS Y SECTORIZACIÓN)
            # =========================================================
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
                ref_limpia = re.sub(r'\.0$', '', ref_val)
                texto_val = str(row.get(col_texto, "")) if col_texto else ""
                asig_val = str(row.get(col_asignacion, "")) if col_asignacion else ""
                t_global = f"{ref_val} {texto_val} {asig_val}".upper()

                if ref_limpia in mapeo_datafono_ref:
                    ref_eval = mapeo_datafono_ref[ref_limpia]
                elif ref_limpia in mapeo_datafono_ref.values():
                    ref_eval = ref_limpia
                else:
                    ref_eval = ""

                if ref_eval or ref_limpia:
                    if ref_eval in ["3001","3002","3003","3004","3005","3006","3007","3008","3009","3010","3011","3012","3013","3200","3201","3202","2005","3203"]: return 'Dist Acopi'
                    if ref_eval in ["2001","2002","2003","2210"]: return 'Dist Buga'
                    if ref_eval in ["4002","4001","4003","4004","4005","4006","4008","4009","4010","4200","4253","4007","4203","4201"]: return 'Dist Dosquebradas'
                    if ref_eval in ["6101","6102","6103","6106","6108"]: return 'Dist Pasto'

                if 'DOSQ' in t_global or 'D504' in t_global: return 'Dist Dosquebradas'
                if 'ACOPI' in t_global or 'D503' in t_global: return 'Dist Acopi'
                if 'PASTO' in t_global or 'D505' in t_global: return 'Dist Pasto'
                if 'BUGA' in t_global or 'D502' in t_global: return 'Dist Buga'

                numeros = re.findall(r'\b\d{4}\b', t_global)
                for n in numeros:
                    num = int(n)
                    if 2000 <= num <= 2999: return 'Dist Buga'
                    if 3000 <= num <= 3999: return 'Dist Acopi'
                    if 4000 <= num <= 4999: return 'Dist Dosquebradas'
                    if 6000 <= num <= 6999: return 'Dist Pasto'

                return 'Sin clasificar'

            def obtener_ref_homologada(row):
                texto = f"{row.get(col_referencia,'')} {row.get(col_asignacion,'')}".upper()
                numeros_8 = re.findall(r'\b\d{8}\b', texto)
                for num in numeros_8:
                    if num in mapeo_datafono_ref:
                        return mapeo_datafono_ref[num]
                numeros_4 = re.findall(r'\b\d{4}\b', texto)
                for num in numeros_4:
                    if num in mapeo_datafono_ref.values():
                        return num
                return None

            df['Distribuidora'] = df.apply(clasificar_distribuidora, axis=1)

            # =========================================================
            # FUNCIONES AUXILIARES
            # =========================================================
            def set_estado(indices, estado):
                df.loc[df['ID_Temp'].isin(indices), 'Estado_Conciliacion'] = estado

            def set_comentarios(dic_comentarios):
                for id_temp, texto in dic_comentarios.items():
                    df.loc[df['ID_Temp'] == id_temp, 'Comentario'] = texto

            def resumen_docs(sub_df):
                return ", ".join(str(int(d)) for d in sub_df[col_doc].tolist())

            def es_valor_redondo(v):
                return (v % multiplo_redondo == 0) and v > 0

            df_40 = df[df[col_clave] == '40'].copy()
            df_50 = df[df[col_clave] == '50'].copy()

            # =========================================================
            # 5. NIVEL 1A — CRUCE EXACTO POR REFERENCIA
            # =========================================================
            df_40['Turno'] = df_40.groupby([col_banco, 'Abs_Importe', col_fecha, col_referencia]).cumcount()
            df_50['Turno'] = df_50.groupby([col_banco, 'Abs_Importe', col_fecha, col_referencia]).cumcount()
            c1 = pd.merge(df_40, df_50, on=[col_banco, 'Abs_Importe', col_fecha, col_referencia, 'Turno'], suffixes=('_40', '_50'))

            df_40['Turno2'] = df_40.groupby([col_banco, 'Abs_Importe', col_fecha, col_asignacion]).cumcount()
            df_50['Turno2'] = df_50.groupby([col_banco, 'Abs_Importe', col_fecha, col_referencia]).cumcount()
            c2 = pd.merge(df_40, df_50, left_on=[col_banco, 'Abs_Importe', col_fecha, col_asignacion, 'Turno2'], right_on=[col_banco, 'Abs_Importe', col_fecha, col_referencia, 'Turno2'], suffixes=('_40', '_50'))

            df_40['Turno3'] = df_40.groupby([col_banco, 'Abs_Importe', col_fecha, col_referencia]).cumcount()
            df_50['Turno3'] = df_50.groupby([col_banco, 'Abs_Importe', col_fecha, col_asignacion]).cumcount()
            c3 = pd.merge(df_40, df_50, left_on=[col_banco, 'Abs_Importe', col_fecha, col_referencia, 'Turno3'], right_on=[col_banco, 'Abs_Importe', col_fecha, col_asignacion, 'Turno3'], suffixes=('_40', '_50'))

            ind_r1 = (set(c1['ID_Temp_40']) | set(c1['ID_Temp_50']) | set(c2['ID_Temp_40']) | set(c2['ID_Temp_50']) | set(c3['ID_Temp_40']) | set(c3['ID_Temp_50']))
            set_estado(ind_r1, 'Conciliado - Cruce exacto (Referencia)')

            comentarios_r1 = {}
            for c in (c1, c2, c3):
                for _, r in c.iterrows():
                    comentarios_r1[r['ID_Temp_40']] = f"Cruce exacto con Doc. {int(r[col_doc + '_50'])} (misma referencia, banco e importe)"
                    comentarios_r1[r['ID_Temp_50']] = f"Cruce exacto con Doc. {int(r[col_doc + '_40'])} (misma referencia, banco e importe)"
            set_comentarios(comentarios_r1)

            # =========================================================
            # 6. NIVEL 1B — CRUCE POR REFERENCIA "LIMPIA"
            # =========================================================
            df_p0 = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_p0['Asig_limpia'] = df_p0[col_asignacion].astype(str).str.extract(r'(\d+)')[0]
            df_p0['Ref_limpia'] = df_p0[col_referencia].astype(str).str.extract(r'(\d+)')[0]

            d40b = df_p0[df_p0[col_clave] == '40'].drop(columns=['Ref_limpia'])
            d50b = df_p0[df_p0[col_clave] == '50'].drop(columns=['Asig_limpia'])
            d40b['Turno_b'] = d40b.groupby([col_banco, 'Abs_Importe', col_fecha, 'Asig_limpia']).cumcount()
            d50b['Turno_b'] = d50b.groupby([col_banco, 'Abs_Importe', col_fecha, 'Ref_limpia']).cumcount()

            c1b = pd.merge(d40b, d50b, left_on=[col_banco, 'Abs_Importe', col_fecha, 'Asig_limpia', 'Turno_b'], right_on=[col_banco, 'Abs_Importe', col_fecha, 'Ref_limpia', 'Turno_b'], suffixes=('_40', '_50'))
            c1b = c1b[c1b['Asig_limpia'].notna() & (c1b['Asig_limpia'] != '')]

            ind_r1b = set(c1b['ID_Temp_40']) | set(c1b['ID_Temp_50'])
            set_estado(ind_r1b, 'Conciliado - Cruce exacto (Ref. limpia)')

            comentarios_r1b = {}
            for _, r in c1b.iterrows():
                comentarios_r1b[r['ID_Temp_40']] = f"Cruce exacto con Doc. {int(r[col_doc + '_50'])} (referencia limpiada)"
                comentarios_r1b[r['ID_Temp_50']] = f"Cruce exacto con Doc. {int(r[col_doc + '_40'])} (referencia limpiada)"
            set_comentarios(comentarios_r1b)

            # =========================================================
            # 7. NUEVO NIVEL 1C — CRUCE MÚLTIPLE (1:N) PARA DATÁFONOS (IP vs CB)
            # =========================================================
            df_p_ipcb = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_p_ipcb['Ref_Homologada'] = df_p_ipcb.apply(obtener_ref_homologada, axis=1)
            df_validos_ipcb = df_p_ipcb[df_p_ipcb['Ref_Homologada'].notna()]

            df_ip = df_validos_ipcb[(df_validos_ipcb[col_clase_doc].astype(str).str.upper() == 'IP') & (df_validos_ipcb[col_clave] == '40')]
            df_cb = df_validos_ipcb[(df_validos_ipcb[col_clase_doc].astype(str).str.upper() == 'CB') & (df_validos_ipcb[col_clave] == '50')]

            ind_1c_ipcb = set()
            com_1c_ipcb = {}

            if not df_ip.empty and not df_cb.empty:
                grp_ip = df_ip.groupby([col_banco, col_fecha, 'Ref_Homologada'])['Abs_Importe'].sum().reset_index()
                grp_ip.rename(columns={'Abs_Importe': 'Suma_IP'}, inplace=True)

                merged_ipcb = pd.merge(df_cb, grp_ip, on=[col_banco, col_fecha, 'Ref_Homologada'])
                matches_ipcb = merged_ipcb[merged_ipcb['Abs_Importe'] == merged_ipcb['Suma_IP']]

                for _, m in matches_ipcb.iterrows():
                    cb_id = m['ID_Temp']
                    banco_val = m[col_banco]
                    fecha_val = m[col_fecha]
                    ref_val = m['Ref_Homologada']

                    ip_subset = df_ip[(df_ip[col_banco] == banco_val) & 
                                      (df_ip[col_fecha] == fecha_val) & 
                                      (df_ip['Ref_Homologada'] == ref_val)]

                    ip_ids = ip_subset['ID_Temp'].tolist()
                    docs_ip = resumen_docs(ip_subset)

                    ind_1c_ipcb.update([cb_id] + ip_ids)

                    com_1c_ipcb[cb_id] = f"Cruce múltiple Datáfono (Suma de {len(ip_ids)} IP). Ref. Interna: {ref_val}. Docs IP: {docs_ip}"
                    for ip_id in ip_ids:
                        com_1c_ipcb[ip_id] = f"Cruce múltiple Datáfono (Parte de consolidado CB). Ref. Interna: {ref_val}. Doc CB: {int(m[col_doc])}"

            set_estado(ind_1c_ipcb, 'Conciliado - Cruce múltiple Datáfono (IP vs CB)')
            set_comentarios(com_1c_ipcb)

            # =========================================================
            # 8. NIVEL 1D — SECTORIZACIÓN POR DISTRIBUIDORA
            # PREVENCIÓN: IP excluidos (solo cruzan en 1C)
            # =========================================================
            df_p1d = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_p1d = df_p1d[df_p1d[col_clase_doc].astype(str).str.upper() != 'IP'] # BLINDAJE IP
            df_p1d = df_p1d[df_p1d['Distribuidora'] != 'Sin clasificar']

            d40c = df_p1d[df_p1d[col_clave] == '40'].copy()
            d50c = df_p1d[df_p1d[col_clave] == '50'].copy()
            grp_dist = [col_banco, 'Abs_Importe', col_fecha, 'Distribuidora']

            ind_r1d = set(); com_r1d = {}
            ind_r1d_fuerte = set(); com_r1d_fuerte = {}
            ind_r1d_amb = set(); com_r1d_amb = {}

            if not d40c.empty and not d50c.empty:
                for grupo, sub40 in d40c.groupby(grp_dist):
                    banco_g, importe_g, fecha_g, dist_g = grupo
                    sub50 = d50c[(d50c[col_banco] == banco_g) & (d50c['Abs_Importe'] == importe_g) &
                                 (d50c[col_fecha] == fecha_g) & (d50c['Distribuidora'] == dist_g)]
                    if sub50.empty:
                        continue

                    sub40_ord = sub40.sort_values(col_doc)
                    sub50_ord = sub50.sort_values(col_doc)

                    if len(sub40_ord) == 1 and len(sub50_ord) == 1:
                        r40 = sub40_ord.iloc[0]
                        r50 = sub50_ord.iloc[0]
                        ind_r1d.update([r40['ID_Temp'], r50['ID_Temp']])
                        com_r1d[r40['ID_Temp']] = f"Único candidato por sede/rango ({dist_g}) - Doc. {int(r50[col_doc])}"
                        com_r1d[r50['ID_Temp']] = f"Único candidato por sede/rango ({dist_g}) - Doc. {int(r40[col_doc])}"
                    elif len(sub40_ord) == len(sub50_ord):
                        for (_, r40), (_, r50) in zip(sub40_ord.iterrows(), sub50_ord.iterrows()):
                            ind_r1d_fuerte.update([r40['ID_Temp'], r50['ID_Temp']])
                            txt = (f"Sede '{dist_g}' con misma cantidad de candidatos por lado; emparejado por FIFO")
                            com_r1d_fuerte[r40['ID_Temp']] = f"{txt}. Contraparte sugerida: Doc. {int(r50[col_doc])}"
                            com_r1d_fuerte[r50['ID_Temp']] = f"{txt}. Contraparte sugerida: Doc. {int(r40[col_doc])}"
                    else:
                        docs40 = resumen_docs(sub40_ord)
                        docs50 = resumen_docs(sub50_ord)
                        for _, r in sub40_ord.iterrows():
                            ind_r1d_amb.add(r['ID_Temp'])
                            com_r1d_amb[r['ID_Temp']] = f"Sede '{dist_g}': cantidades distintas ({len(sub40_ord)} vs {len(sub50_ord)}) - Docs créditos: {docs50}"
                        for _, r in sub50_ord.iterrows():
                            ind_r1d_amb.add(r['ID_Temp'])
                            com_r1d_amb[r['ID_Temp']] = f"Sede '{dist_g}': cantidades distintas ({len(sub50_ord)} vs {len(sub40_ord)}) - Docs débitos: {docs40}"

            set_estado(ind_r1d, 'Conciliado - Cruce por Sectorización (candidato único)')
            set_comentarios(com_r1d)
            set_estado(ind_r1d_fuerte, 'Sugerencia fuerte: Sectorización (FIFO, cantidades iguales)')
            set_comentarios(com_r1d_fuerte)
            set_estado(ind_r1d_amb, 'Sugerencia: Sectorización con candidatos desbalanceados')
            set_comentarios(com_r1d_amb)

            # =========================================================
            # 9. NIVEL 2 — CRUCE ÚNICO SIN REFERENCIA NI SECTORIZACIÓN
            # PREVENCIÓN: IP excluidos
            # =========================================================
            df_p = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_p = df_p[df_p[col_clase_doc].astype(str).str.upper() != 'IP'] # BLINDAJE IP
            
            grp_cols = [col_banco, 'Abs_Importe', col_fecha]

            df_p40 = df_p[df_p[col_clave] == '40'].copy()
            df_p50 = df_p[df_p[col_clave] == '50'].copy()

            df_p40['n40'] = df_p40.groupby(grp_cols)['ID_Temp'].transform('count')
            df_p50['n50'] = df_p50.groupby(grp_cols)['ID_Temp'].transform('count')

            unicos40 = df_p40[df_p40['n40'] == 1]
            unicos50 = df_p50[df_p50['n50'] == 1]
            c_unico = pd.merge(unicos40, unicos50, on=grp_cols, suffixes=('_40', '_50'))

            ind_r2 = set(c_unico['ID_Temp_40']) | set(c_unico['ID_Temp_50'])
            set_estado(ind_r2, 'Conciliado - Cruce unico sin referencia')

            comentarios_r2 = {}
            for _, r in c_unico.iterrows():
                comentarios_r2[r['ID_Temp_40']] = f"Cruce único con Doc. {int(r[col_doc + '_50'])} (mismo banco/fecha/importe, sin ambigüedad)"
                comentarios_r2[r['ID_Temp_50']] = f"Cruce único con Doc. {int(r[col_doc + '_40'])} (mismo banco/fecha/importe, sin ambigüedad)"
            set_comentarios(comentarios_r2)

            # =========================================================
            # 10. NIVEL 2B — DESEMPATE POR GRUPO CERRADO Y CANDIDATOS
            # =========================================================
            ambiguos40 = df_p40[~df_p40['ID_Temp'].isin(ind_r2)]
            ambiguos50 = df_p50[~df_p50['ID_Temp'].isin(ind_r2)]
            
            ind_r2d = set(); comentarios_r2d = {}
            ind_amb = set(); comentarios_amb = {}

            for grupo, sub40 in ambiguos40.groupby(grp_cols):
                banco_g, importe_g, fecha_g = grupo
                sub50 = ambiguos50[(ambiguos50[col_banco] == banco_g) &
                                    (ambiguos50['Abs_Importe'] == importe_g) &
                                    (ambiguos50[col_fecha] == fecha_g)]
                if sub50.empty:
                    continue

                if es_valor_redondo(importe_g):
                    sub40_ord = sub40.sort_values(col_doc)
                    sub50_ord = sub50.sort_values(col_doc)
                    if len(sub40_ord) == len(sub50_ord):
                        for (_, r40), (_, r50) in zip(sub40_ord.iterrows(), sub50_ord.iterrows()):
                            ind_r2d.update([r40['ID_Temp'], r50['ID_Temp']])
                            txt = f"Valor redondo (${importe_g:,.0f}) emparejado por FIFO (VERIFICAR)"
                            comentarios_r2d[r40['ID_Temp']] = f"{txt} - con Doc. {int(r50[col_doc])}"
                            comentarios_r2d[r50['ID_Temp']] = f"{txt} - con Doc. {int(r40[col_doc])}"
                    else:
                        docs50 = resumen_docs(sub50_ord)
                        docs40 = resumen_docs(sub40_ord)
                        for _, r in sub40_ord.iterrows():
                            ind_amb.add(r['ID_Temp'])
                            comentarios_amb[r['ID_Temp']] = f"Valor redondo ambiguo ({len(sub40_ord)} vs {len(sub50_ord)} docs) - Docs créditos: {docs50}"
                        for _, r in sub50_ord.iterrows():
                            ind_amb.add(r['ID_Temp'])
                            comentarios_amb[r['ID_Temp']] = f"Valor redondo ambiguo ({len(sub50_ord)} vs {len(sub40_ord)} docs) - Docs débitos: {docs40}"
                else:
                    docs40 = resumen_docs(sub40)
                    docs50 = resumen_docs(sub50)
                    for _, r in sub40.iterrows():
                        ind_amb.add(r['ID_Temp'])
                        comentarios_amb[r['ID_Temp']] = f"{len(sub50)} posibles cruces por importe/fecha (Docs candidatos: {docs50})"
                    for _, r in sub50.iterrows():
                        ind_amb.add(r['ID_Temp'])
                        comentarios_amb[r['ID_Temp']] = f"{len(sub40)} posibles cruces por importe/fecha (Docs candidatos: {docs40})"

            set_estado(ind_r2d, 'Sugerencia fuerte: Emparejado por FIFO en grupo cerrado (valor redondo)')
            set_comentarios(comentarios_r2d)
            set_estado(ind_amb, 'Sugerencia: Múltiples candidatos sin referencia')
            set_comentarios(comentarios_amb)

            # =========================================================
            # 11. NIVEL 3 — SUGERENCIAS BASADAS EN REFERENCIA VÁLIDA
            # PREVENCIÓN: IP excluidos
            # =========================================================
            df_pend = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_pend = df_pend[df_pend[col_clase_doc].astype(str).str.upper() != 'IP'] # BLINDAJE IP
            
            df_pend['Ref_Limpia'] = df_pend[col_referencia].astype(str).str.strip().str.lower()
            df_validos = df_pend[~df_pend['Ref_Limpia'].isin(['nan', '', 'none', '0', '/'])].copy()

            df_40n = df_validos[df_validos[col_clave] == '40'].copy()
            df_50n = df_validos[df_validos[col_clave] == '50'].copy()

            # --- 11A: Validar con error de fecha (distingue mismo periodo contable) ---
            sA = pd.merge(df_40n, df_50n, on=[col_banco, 'Abs_Importe', 'Ref_Limpia'], suffixes=('_40', '_50'))
            sA['Dif_Dias'] = (sA['Fecha_Calc_40'] - sA['Fecha_Calc_50']).dt.days.abs()
            sA = sA[sA['Dif_Dias'] > 0].sort_values('Dif_Dias').drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            ind_A = set(); com_A = {}
            for _, r in sA.iterrows():
                f40, f50 = r['Fecha_Calc_40'], r['Fecha_Calc_50']
                dif = int(r['Dif_Dias'])
                mismo_periodo = (f40.month == f50.month) and (f40.year == f50.year)

                if dif <= tol_dias:
                    estado_asignado = ' Diferencia de Fecha (Mismo Periodo)' if mismo_periodo else ' DIFERENTE PERIODO CONTABLE'
                    txt_estado = "mismo periodo" if mismo_periodo else "¡DIFERENTE MES/AÑO!"
                    com_txt = f"Misma referencia/banco/importe, pero difieren {dif} día(s) ({txt_estado})"
                elif mismo_periodo:
                    estado_asignado = ' Diferencia de Fecha MAYOR A 3 DIAS (Mismo Periodo)'
                    com_txt = f"Misma referencia/banco/importe, difiere {dif} días (supera tolerancia), pero es del mismo mes"
                else:
                    continue 

                ind_A.update([r['ID_Temp_40'], r['ID_Temp_50']])
                df.loc[df['ID_Temp'].isin([r['ID_Temp_40'], r['ID_Temp_50']]), 'Estado_Conciliacion'] = estado_asignado
                com_A[r['ID_Temp_40']] = f"{com_txt}. Doc: {int(r[col_doc+'_50'])}"
                com_A[r['ID_Temp_50']] = f"{com_txt}. Doc: {int(r[col_doc+'_40'])}"
            set_comentarios(com_A)

            df_40n = df_40n[~df_40n['ID_Temp'].isin(ind_A)]
            df_50n = df_50n[~df_50n['ID_Temp'].isin(ind_A)]

            # --- 11B: Sugerencia de reclasificación de banco ---
            sB = pd.merge(df_40n, df_50n, on=['Abs_Importe', col_fecha, 'Ref_Limpia'], suffixes=('_40', '_50'))
            sB = sB[sB[f'{col_banco}_40'] != sB[f'{col_banco}_50']]
            sB = sB.drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            ind_B = set(sB['ID_Temp_40']) | set(sB['ID_Temp_50'])
            set_estado(ind_B, ' Reclasificación de banco')
            com_B = {}
            for _, r in sB.iterrows():
                com_B[r['ID_Temp_40']] = f"Misma referencia/importe/fecha, pero en banco distinto '{r[col_banco+'_50']}'. Doc: {int(r[col_doc+'_50'])}"
                com_B[r['ID_Temp_50']] = f"Misma referencia/importe/fecha, pero en banco distinto '{r[col_banco+'_40']}'. Doc: {int(r[col_doc+'_40'])}"
            set_comentarios(com_B)

            df_40n = df_40n[~df_40n['ID_Temp'].isin(ind_B)]
            df_50n = df_50n[~df_50n['ID_Temp'].isin(ind_B)]

            # --- 11C: Revisar diferencia de valor ---
            sC = pd.merge(df_40n, df_50n, on=[col_banco, col_fecha, 'Ref_Limpia'], suffixes=('_40', '_50'))
            sC['Dif_Valor'] = (sC['Abs_Importe_40'] - sC['Abs_Importe_50']).abs()
            max_importe = sC[['Abs_Importe_40', 'Abs_Importe_50']].max(axis=1)
            sC['Dif_Pct'] = np.where(max_importe == 0, 0, sC['Dif_Valor'] / max_importe)

            sC = sC[(sC['Dif_Valor'] > 0) & ((sC['Dif_Valor'] <= tol_valor_abs) | (sC['Dif_Pct'] <= tol_valor_pct))]
            sC = sC.sort_values('Dif_Valor').drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            ind_C = set(sC['ID_Temp_40']) | set(sC['ID_Temp_50'])
            set_estado(ind_C, ' Diferencia de valor')
            com_C = {}
            for _, r in sC.iterrows():
                com_C[r['ID_Temp_40']] = f"Misma referencia/banco/fecha, con diferencia de ${r['Dif_Valor']:,.0f} ({r['Dif_Pct']*100:.2f}%). Doc: {int(r[col_doc+'_50'])}"
                com_C[r['ID_Temp_50']] = f"Misma referencia/banco/fecha, con diferencia de ${r['Dif_Valor']:,.0f} ({r['Dif_Pct']*100:.2f}%). Doc: {int(r[col_doc+'_40'])}"
            set_comentarios(com_C)

            # --- Lo que sigue pendiente ---
            sin_pista = df['Estado_Conciliacion'] == 'Pendiente'
            es_ip = df[col_clase_doc].astype(str).str.upper() == 'IP'
            
            # Asignación de comentario especial para IP
            df.loc[sin_pista & es_ip & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia - IP exclusivo requiere CB correspondiente'
            df.loc[sin_pista & ~es_ip & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia ni sugerencia encontrada - requiere revisión manual completa'

            # =========================================================
            # 12. CONTROL DE INTEGRIDAD
            # =========================================================
            total_filas_entrada = filas_antes
            total_filas_salida = len(df) + len(filas_descartadas)
            cuadre_ok = total_filas_entrada == total_filas_salida

            # =========================================================
            # 13. LIMPIEZA FINAL Y FORMATO DE FECHAS
            # =========================================================
            df_final = df.drop(columns=['ID_Temp', 'Abs_Importe', 'Fecha_Calc'], errors='ignore')
            columnas_fecha = [c for c in df_final.columns if 'fe.' in c.lower() or 'fecha' in c.lower() or 'fe-' in c.lower()]
            for col_f in columnas_fecha:
                df_final[col_f] = pd.to_datetime(df_final[col_f], errors='coerce').dt.strftime('%d/%m/%Y')

            # =========================================================
            # 14. PALETA DE COLORES PASTEL SEMÁNTICA
            # =========================================================
            def resaltar_conciliados(row):
                est = str(row['Estado_Conciliacion']).lower()

                if 'cruce exacto' in est or 'sectorización (candidato único)' in est or 'cruce múltiple' in est:
                    return ['background-color: #E2EFDA; color: black'] * len(row)
                elif 'cruce unico' in est:
                    return ['background-color: #DDEBF7; color: black'] * len(row)
                elif 'fifo en grupo cerrado' in est or 'sectorización (fifo' in est:
                    return ['background-color: #FFF2CC; color: black'] * len(row)
                elif 'alerta' in est:
                    return ['background-color: #FCE4D6; color: black'] * len(row)
                elif 'múltiples' in est or 'multiples' in est or 'desbalanceados' in est:
                    return ['background-color: #F5B7B1; color: black'] * len(row)
                elif 'pendiente' in est:
                    return ['background-color: #F2F2F2; color: black'] * len(row)
                return [''] * len(row)

            # =========================================================
            # 15. EXPORTACIÓN
            # =========================================================
            output = io.BytesIO()
            bancos_unicos = [b for b in df_final[col_banco].unique() if str(b).strip().lower() not in ('', 'nan')]

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                resumen = df_final['Estado_Conciliacion'].value_counts().reset_index()
                resumen.columns = ['Estado', 'Cantidad de registros']
                resumen.to_excel(writer, index=False, sheet_name='RESUMEN')

                for banco in bancos_unicos:
                    df_banco = df_final[df_final[col_banco] == banco].copy()
                    df_banco = df_banco.sort_values(by=col_importe, ascending=True)

                    nombre_pestana = str(banco)[:31]
                    for ch in ['/', '\\', ':', '?', '*', '[', ']']:
                        nombre_pestana = nombre_pestana.replace(ch, '-')
                    if not nombre_pestana.strip() or nombre_pestana.lower() == 'nan':
                        nombre_pestana = "Sin_Banco_Asignado"

                    styled_banco = df_banco.style.apply(resaltar_conciliados, axis=1)
                    styled_banco.to_excel(writer, index=False, sheet_name=nombre_pestana)

                df_novedades = df_final[~df_final['Estado_Conciliacion'].str.contains('Conciliado', na=False)].copy()
                if not df_novedades.empty:
                    df_novedades = df_novedades.sort_values(by=['Estado_Conciliacion', col_importe], ascending=[True, True])
                    styled_novedades = df_novedades.style.apply(resaltar_conciliados, axis=1)
                    styled_novedades.to_excel(writer, index=False, sheet_name='NOVEDADES_Y_ALERTAS')

                if not filas_descartadas.empty:
                    filas_descartadas.to_excel(writer, index=False, sheet_name='DESCARTADAS_SIN_DOC_O_CT')

            # =========================================================
            # 16. INTERFAZ Y DESCARGA
            # =========================================================
            st.success("¡Conciliación Integral terminada! Se integró validación múltiple (1:N) para Datáfonos y exclusividad en referencias IP.")

            if not cuadre_ok:
                st.warning("⚠️ Alerta de integridad: el total de filas de salida no coincide con el de entrada. Revisa la pestaña DESCARTADAS.")

            conciliados_exactos = len(ind_r1) + len(ind_r1b) + len(ind_r1d) + len(ind_1c_ipcb)
            conciliados_unicos = len(ind_r2)
            fuerte = len(ind_r1d_fuerte) + len(ind_r2d)
            alertas = len(ind_r1d_amb) + len(ind_amb) + len(ind_A) + len(ind_B) + len(ind_C)
            pendientes_sin_pista = len(df_final) - conciliados_exactos - conciliados_unicos - fuerte - alertas

            col1, col2, col3, col4, col5, col6 = st.columns(6)
            col1.metric("Bancos procesados", len(bancos_unicos))
            col2.metric("Conciliado exacto/múltiple", conciliados_exactos)
            col3.metric("Conciliado único", conciliados_unicos)
            col4.metric("Fuerte (verificar)", fuerte)
            col5.metric("Alertas (revisar)", alertas)
            col6.metric("Sin ninguna pista", pendientes_sin_pista)

            if filas_excluidas > 0:
                st.warning(f"⚠️ Se excluyeron {filas_excluidas} filas sin Nº de documento o sin clave contable (totales, subtotales o filas vacías).")

            st.download_button(
                label="📥 Descargar Excel con Resultados",
                data=output.getvalue(),
                file_name="Conciliacion_Integral_Resultados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error técnico detectado. Detalle técnico: {e}")
