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
    "solo marca como *Conciliado* lo que tiene evidencia inequívoca (referencia exacta, "
    "referencia limpia, cruce por sectorización de sede o cruce único sin ambigüedad). "
    "Los colores utilizan una **paleta de tonos pastel suaves** amigables con la vista, "
    "categorizados semánticamente para que puedas auditar el Excel rápidamente."
)

with st.expander("⚙️ Parámetros de tolerancia para sugerencias (alertas)"):
    tol_dias = st.slider("Días máximos de diferencia regular para alertar 'error de fecha'", 1, 15, 3)
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

            columnas_requeridas = [col_referencia, col_clave, col_fecha, col_importe, col_banco, col_doc]
            faltantes = [c for c in columnas_requeridas if c not in df.columns]
            if faltantes:
                st.error(f"No se encontraron estas columnas obligatorias en el archivo: {faltantes}")
                st.stop()
                
            if 'Asignación' not in df.columns and 'Asignaión' not in df.columns:
                st.error("No se encontró la columna de Asignación en el archivo.")
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
            # 4B. CLASIFICACIÓN DE DISTRIBUIDORAS INTEGRAL
            # =========================================================
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
                "17608605": "Dist Pasto",
                "17968405": "VENTA EN LINEA"
            }

            def clasificar_distribuidora(row):
                ref_val = str(row.get(col_referencia, "")).strip()
                ref_limpia = re.sub(r'\.0$', '', ref_val)
                
                # 1. Buscar en la base de datos de referencias
                if ref_limpia in mapeo_referencias_dist:
                    return mapeo_referencias_dist[ref_limpia]
                
                # 2. Buscar patrones combinando Texto, Asignación y Referencia
                texto_val = str(row.get(col_texto, "")) if col_texto else ""
                asig_val = str(row.get(col_asignacion, "")) if col_asignacion else ""
                t_global = f"{ref_val} {texto_val} {asig_val}".upper()
                
                if 'DOSQ' in t_global or 'D504' in t_global: return 'Dist Dosquebradas'
                if 'ACOPI' in t_global or 'D503' in t_global: return 'Dist Acopi'
                if 'PASTO' in t_global or 'D505' in t_global: return 'Dist Pasto'
                if 'BUGA' in t_global or 'D502' in t_global: return 'Dist Buga'
                
                # 3. Buscar rangos numéricos de 4 dígitos
                numeros = re.findall(r'\b\d{4}\b', t_global)
                for n in numeros:
                    num = int(n)
                    if 2000 <= num <= 2999: return 'Dist Buga'
                    if 3000 <= num <= 3999: return 'Dist Acopi'
                    if 4000 <= num <= 4999: return 'Dist Dosquebradas'
                    if 6000 <= num <= 6999: return 'Dist Pasto'
                        
                return 'Sin clasificar'

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
                    comentarios_r1[r['ID_Temp_40']] = f"Cruce exacto con Doc. {int(r[col_doc + '_50'])} (misma referencia)"
                    comentarios_r1[r['ID_Temp_50']] = f"Cruce exacto con Doc. {int(r[col_doc + '_40'])} (misma referencia)"
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
                comentarios_r1b[r['ID_Temp_40']] = f"Cruce exacto con Doc. {int(r[col_doc + '_50'])} (Ref. limpia)"
                comentarios_r1b[r['ID_Temp_50']] = f"Cruce exacto con Doc. {int(r[col_doc + '_40'])} (Ref. limpia)"
            set_comentarios(comentarios_r1b)

            # =========================================================
            # NUEVO NIVEL 1C — CRUCE POR SECTORIZACIÓN / DISTRIBUIDORA
            # =========================================================
            df_p1c = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_p1c = df_p1c[df_p1c['Distribuidora'] != 'Sin clasificar']
            
            d40c = df_p1c[df_p1c[col_clave] == '40']
            d50c = df_p1c[df_p1c[col_clave] == '50']
            
            d40c['Turno_c'] = d40c.groupby([col_banco, 'Abs_Importe', col_fecha, 'Distribuidora']).cumcount()
            d50c['Turno_c'] = d50c.groupby([col_banco, 'Abs_Importe', col_fecha, 'Distribuidora']).cumcount()
            
            c1c = pd.merge(d40c, d50c, on=[col_banco, 'Abs_Importe', col_fecha, 'Distribuidora', 'Turno_c'], suffixes=('_40', '_50'))
            
            ind_r1c = set(c1c['ID_Temp_40']) | set(c1c['ID_Temp_50'])
            set_estado(ind_r1c, 'Conciliado - Cruce por Sectorización/Rango')
            
            comentarios_r1c = {}
            for _, r in c1c.iterrows():
                com = f"Cruce resuelto por coincidencia de Sede/Rango ({r['Distribuidora']})"
                comentarios_r1c[r['ID_Temp_40']] = f"{com} - Doc: {int(r[col_doc + '_50'])}"
                comentarios_r1c[r['ID_Temp_50']] = f"{com} - Doc: {int(r[col_doc + '_40'])}"
            set_comentarios(comentarios_r1c)

            # =========================================================
            # 7. NIVEL 2 — CRUCE ÚNICO SIN REFERENCIA
            # =========================================================
            df_p = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
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
                comentarios_r2[r['ID_Temp_40']] = f"Cruce único con Doc. {int(r[col_doc + '_50'])} (sin ambigüedad)"
                comentarios_r2[r['ID_Temp_50']] = f"Cruce único con Doc. {int(r[col_doc + '_40'])} (sin ambigüedad)"
            set_comentarios(comentarios_r2)

            # --- Grupos AMBIGUOS ---
            ambiguos40 = df_p40[(df_p40['n40'] > 1) & (~df_p40['ID_Temp'].isin(ind_r2))]
            ambiguos50 = df_p50[(df_p50['n50'] > 1) & (~df_p50['ID_Temp'].isin(ind_r2))]

            # =========================================================
            # 8. NIVEL 2B — DESEMPATE POR GRUPO CERRADO
            # =========================================================
            ind_r2d = set(); comentarios_r2d = {}
            ind_amb = set(); comentarios_amb = {}

            for grupo, sub40 in ambiguos40.groupby(grp_cols):
                banco_g, importe_g, fecha_g = grupo
                sub50 = ambiguos50[(ambiguos50[col_banco] == banco_g) & (ambiguos50['Abs_Importe'] == importe_g) & (ambiguos50[col_fecha] == fecha_g)]
                if sub50.empty: continue

                if es_valor_redondo(importe_g):
                    sub40_ord = sub40.sort_values(col_doc)
                    sub50_ord = sub50.sort_values(col_doc)
                    if len(sub40_ord) == len(sub50_ord):
                        for (_, r40), (_, r50) in zip(sub40_ord.iterrows(), sub50_ord.iterrows()):
                            ind_r2d.update([r40['ID_Temp'], r50['ID_Temp']])
                            comentarios_r2d[r40['ID_Temp']] = f"Valor redondo empatado por FIFO. Verifica con Doc. {int(r50[col_doc])}"
                            comentarios_r2d[r50['ID_Temp']] = f"Valor redondo empatado por FIFO. Verifica con Doc. {int(r40[col_doc])}"
                    else:
                        docs50 = resumen_docs(sub50_ord)
                        docs40 = resumen_docs(sub40_ord)
                        for _, r in sub40_ord.iterrows():
                            ind_amb.add(r['ID_Temp'])
                            comentarios_amb[r['ID_Temp']] = f"Valor redondo ambiguo: {len(sub40_ord)} docs vs {len(sub50_ord)} (Créditos: {docs50})"
                        for _, r in sub50_ord.iterrows():
                            ind_amb.add(r['ID_Temp'])
                            comentarios_amb[r['ID_Temp']] = f"Valor redondo ambiguo: {len(sub50_ord)} docs vs {len(sub40_ord)} (Débitos: {docs40})"
                else:
                    docs40 = resumen_docs(sub40)
                    docs50 = resumen_docs(sub50)
                    for _, r in sub40.iterrows():
                        ind_amb.add(r['ID_Temp'])
                        comentarios_amb[r['ID_Temp']] = f"{len(sub50)} posibles cruces mismo importe/fecha (Docs: {docs50})"
                    for _, r in sub50.iterrows():
                        ind_amb.add(r['ID_Temp'])
                        comentarios_amb[r['ID_Temp']] = f"{len(sub40)} posibles cruces mismo importe/fecha (Docs: {docs40})"

            set_estado(ind_r2d, 'Sugerencia fuerte: Emparejado por FIFO en grupo cerrado')
            set_comentarios(comentarios_r2d)
            set_estado(ind_amb, 'Múltiples candidatos sin referencia')
            set_comentarios(comentarios_amb)

            # =========================================================
            # 9. NIVEL 3 — SUGERENCIAS BASADAS EN REFERENCIA VÁLIDA
            # =========================================================
            df_pend = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_pend['Ref_Limpia'] = df_pend[col_referencia].astype(str).str.strip().str.lower()
            df_validos = df_pend[~df_pend['Ref_Limpia'].isin(['nan', '', 'none', '0', '/'])].copy()

            df_40n = df_validos[df_validos[col_clave] == '40'].copy()
            df_50n = df_validos[df_validos[col_clave] == '50'].copy()

            # --- 9A: Validar con error de fecha (REGLA AMPLIADA PARA MISMO PERIODO) ---
            sA = pd.merge(df_40n, df_50n, on=[col_banco, 'Abs_Importe', 'Ref_Limpia'], suffixes=('_40', '_50'))
            sA['Dif_Dias'] = (sA['Fecha_Calc_40'] - sA['Fecha_Calc_50']).dt.days.abs()
            sA = sA[sA['Dif_Dias'] > 0].sort_values('Dif_Dias').drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            ind_A = set()
            com_A = {}
            for _, r in sA.iterrows():
                f40, f50 = r['Fecha_Calc_40'], r['Fecha_Calc_50']
                dif = int(r['Dif_Dias'])
                mismo_periodo = (f40.month == f50.month) and (f40.year == f50.year)
                
                if dif <= tol_dias:
                    estado_asignado = 'Alerta: Diferencia de Fecha (Mismo Periodo)' if mismo_periodo else 'Alerta: DIFERENTE PERIODO CONTABLE'
                    txt_estado = "mismo periodo" if mismo_periodo else "¡DIFERENTE MES/AÑO!"
                    com_txt = f"Coincide exacto, pero difieren {dif} día(s) ({txt_estado})."
                    
                    ind_A.update([r['ID_Temp_40'], r['ID_Temp_50']])
                    df.loc[df['ID_Temp'].isin([r['ID_Temp_40'], r['ID_Temp_50']]), 'Estado_Conciliacion'] = estado_asignado
                    com_A[r['ID_Temp_40']] = f"{com_txt} Doc: {int(r[col_doc+'_50'])}"
                    com_A[r['ID_Temp_50']] = f"{com_txt} Doc: {int(r[col_doc+'_40'])}"
                    
                elif dif > tol_dias and mismo_periodo:
                    estado_asignado = 'Alerta: Diferencia de Fecha EXTENDIDA (Mismo Periodo)'
                    com_txt = f"Coincide exacto, difiere {dif} días (supera tolerancia), pero es del mismo mes."
                    
                    ind_A.update([r['ID_Temp_40'], r['ID_Temp_50']])
                    df.loc[df['ID_Temp'].isin([r['ID_Temp_40'], r['ID_Temp_50']]), 'Estado_Conciliacion'] = estado_asignado
                    com_A[r['ID_Temp_40']] = f"{com_txt} Doc: {int(r[col_doc+'_50'])}"
                    com_A[r['ID_Temp_50']] = f"{com_txt} Doc: {int(r[col_doc+'_40'])}"

            set_comentarios(com_A)

            df_40n = df_40n[~df_40n['ID_Temp'].isin(ind_A)]
            df_50n = df_50n[~df_50n['ID_Temp'].isin(ind_A)]

            # --- 9B: Sugerencia de reclasificación de banco ---
            sB = pd.merge(df_40n, df_50n, on=['Abs_Importe', col_fecha, 'Ref_Limpia'], suffixes=('_40', '_50'))
            sB = sB[sB[f'{col_banco}_40'] != sB[f'{col_banco}_50']]
            sB = sB.drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            ind_B = set(sB['ID_Temp_40']) | set(sB['ID_Temp_50'])
            set_estado(ind_B, 'Alerta: Reclasificación de banco')
            com_B = {}
            for _, r in sB.iterrows():
                com_B[r['ID_Temp_40']] = f"Coincide en importe/fecha, pero en banco distinto '{r[col_banco+'_50']}'. Doc: {int(r[col_doc+'_50'])}"
                com_B[r['ID_Temp_50']] = f"Coincide en importe/fecha, pero en banco distinto '{r[col_banco+'_40']}'. Doc: {int(r[col_doc+'_40'])}"
            set_comentarios(com_B)

            df_40n = df_40n[~df_40n['ID_Temp'].isin(ind_B)]
            df_50n = df_50n[~df_50n['ID_Temp'].isin(ind_B)]

            # --- 9C: Revisar diferencia de valor ---
            sC = pd.merge(df_40n, df_50n, on=[col_banco, col_fecha, 'Ref_Limpia'], suffixes=('_40', '_50'))
            sC['Dif_Valor'] = (sC['Abs_Importe_40'] - sC['Abs_Importe_50']).abs()
            max_importe = sC[['Abs_Importe_40', 'Abs_Importe_50']].max(axis=1)
            sC['Dif_Pct'] = np.where(max_importe == 0, 0, sC['Dif_Valor'] / max_importe)
            
            sC = sC[(sC['Dif_Valor'] > 0) & ((sC['Dif_Valor'] <= tol_valor_abs) | (sC['Dif_Pct'] <= tol_valor_pct))]
            sC = sC.sort_values('Dif_Valor').drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            ind_C = set(sC['ID_Temp_40']) | set(sC['ID_Temp_50'])
            set_estado(ind_C, 'Alerta: Diferencia de valor')
            com_C = {}
            for _, r in sC.iterrows():
                com_C[r['ID_Temp_40']] = f"Diferencia de ${r['Dif_Valor']:,.0f} ({r['Dif_Pct']*100:.2f}%). Doc: {int(r[col_doc+'_50'])}"
                com_C[r['ID_Temp_50']] = f"Diferencia de ${r['Dif_Valor']:,.0f} ({r['Dif_Pct']*100:.2f}%). Doc: {int(r[col_doc+'_40'])}"
            set_comentarios(com_C)

            # --- Lo pendiente ---
            sin_pista = df['Estado_Conciliacion'] == 'Pendiente'
            df.loc[sin_pista & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia encontrada - requiere revisión manual'

            # =========================================================
            # 10. CONTROL DE INTEGRIDAD
            # =========================================================
            total_filas_entrada = filas_antes
            total_filas_salida = len(df) + len(filas_descartadas)
            cuadre_ok = total_filas_entrada == total_filas_salida

            # =========================================================
            # 11. LIMPIEZA FINAL Y FORMATO
            # =========================================================
            df_final = df.drop(columns=['ID_Temp', 'Abs_Importe', 'Fecha_Calc'], errors='ignore')
            columnas_fecha = [c for c in df_final.columns if 'fe.' in c.lower() or 'fecha' in c.lower() or 'fe-' in c.lower()]
            for col_f in columnas_fecha:
                df_final[col_f] = pd.to_datetime(df_final[col_f], errors='coerce').dt.strftime('%d/%m/%Y')

            # =========================================================
            # 12. NUEVA PALETA DE COLORES PASTEL SEMÁNTICA
            # =========================================================
            def resaltar_conciliados(row):
                est = str(row['Estado_Conciliacion']).lower()
                
                # 1. Verde Pastel Claro (Lo perfecto / Máxima seguridad)
                if 'cruce exacto' in est or 'cruce por sectorización' in est:
                    return ['background-color: #E2EFDA; color: black'] * len(row)
                # 2. Azul Pastel Claro (Cruce único / Muy seguro)
                elif 'cruce unico' in est:
                    return ['background-color: #DDEBF7; color: black'] * len(row)
                # 3. Amarillo Pastel Claro (Empate / Fuerte por FIFO)
                elif 'fifo en grupo' in est:
                    return ['background-color: #FFF2CC; color: black'] * len(row)
                # 4. Naranja/Melocotón Pastel (Todas las alertas que requieren atención)
                elif 'alerta' in est:
                    return ['background-color: #FCE4D6; color: black'] * len(row)
                # 5. Gris Muy Claro (Múltiples opciones o pendiente manual)
                elif 'múltiples' in est or 'multiples' in est or 'pendiente' in est:
                    return ['background-color: #F2F2F2; color: black'] * len(row)
                
                return [''] * len(row)

            # =========================================================
            # 13. EXPORTACIÓN
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
            # 14. INTERFAZ Y DESCARGA
            # =========================================================
            st.success("¡Conciliación Integral terminada! Se han aplicado cruces por sectorización y validaciones de periodo contable.")

            if not cuadre_ok:
                st.warning("⚠️ Alerta de integridad: el total de filas de salida no coincide con el de entrada. Revisa la pestaña DESCARTADAS.")

            conciliados_exactos = len(ind_r1) + len(ind_r1b) + len(ind_r1c)
            conciliados_unicos = len(ind_r2)
            fifo_grupo = len(ind_r2d)
            alertas = len(ind_amb) + len(ind_A) + len(ind_B) + len(ind_C)
            pendientes_sin_pista = len(df_final) - conciliados_exactos - conciliados_unicos - fifo_grupo - alertas

            col1, col2, col3, col4, col5, col6 = st.columns(6)
            col1.metric("Bancos procesados", len(bancos_unicos))
            col2.metric("Conciliado exacto/sede", conciliados_exactos)
            col3.metric("Conciliado único", conciliados_unicos)
            col4.metric("Fuerte (val. redondo)", fifo_grupo)
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
