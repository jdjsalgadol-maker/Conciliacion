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
    "Sube tu archivo consolidado. El sistema aplica reglas semaforizadas: "
    "Las coincidencias 100% exactas (por referencia, cruce único, o rango de sede) "
    "quedan en **Verde**. Las sugerencias por error de fecha, reclasificación de banco o FIFO "
    "quedan en **Amarillo**. Lo ambiguo o sin pista queda en **Rojo** para tu revisión."
)

with st.expander("⚙️ Parámetros de tolerancia para sugerencias (alertas)"):
    tol_dias = st.slider("Días máximos de diferencia para alertar 'error de fecha'", 1, 10, 3)
    tol_valor_abs = st.number_input("Diferencia absoluta máxima de valor para alertar ($)", min_value=1, value=5000, step=100)
    tol_valor_pct = st.number_input("Diferencia relativa máxima de valor para alertar (%)", min_value=0.01, value=0.5, step=0.01) / 100
    multiplo_redondo = st.selectbox("Múltiplo para considerar un valor 'redondo' (alta ambigüedad)", [50000, 100000], index=1)

archivo_subido = st.file_uploader("Selecciona el archivo de Excel o CSV", type=['xlsx', 'csv'])

if archivo_subido is not None:
    try:
        with st.spinner("Leyendo archivo, unificando datos y aplicando reglas de conciliación..."):

            if archivo_subido.name.lower().endswith('.csv'):
                df = pd.read_csv(archivo_subido)
            else:
                diccionario_hojas = pd.read_excel(archivo_subido, sheet_name=None)
                df = pd.concat(diccionario_hojas.values(), ignore_index=True)

            df.columns = df.columns.str.strip()

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

            # Autocompletado de banco por cuenta
            mapeo_cuentas_banco = {
                "1110056101": "BANCO DE BOGOTA", "1110056201": "BANCO DAVIBANK S.A.",
                "1110056301": "BANCOLOMBIA S.A.", "1110056401": "BANCO CAJA SOCIAL S.",
                "1110056501": "BANCO DAVIVIENDA S.A", "1110056601": "BANCO BILBAO VIZCAYA",
                "1110056701": "BANCO AGRARIO DE COL", "1120055001": "BANCO COMERCIAL AV V",
                "1120055101": "BANCO DE OCCIDENTE", "1120055301": "BANCO GNB SUDAMERIS",
            }

            bancos_completados = []
            for _, row in df.iterrows():
                asig_val = str(row.get(col_asignacion, ""))
                banco_val = row.get(col_banco, None)
                current_bank = None

                if "cuenta de mayor" in asig_val.lower():
                    match_cuenta = re.search(r'(\d{6,})', asig_val)
                    if match_cuenta:
                        current_bank = mapeo_cuentas_banco.get(match_cuenta.group(1), f"CUENTA {match_cuenta.group(1)}")

                if pd.notnull(banco_val) and str(banco_val).strip().lower() not in ("", "nan"):
                    current_bank = str(banco_val).strip()
                bancos_completados.append(current_bank)

            df[col_banco] = bancos_completados
            if col_asignacion in df.columns:
                df = df[~df[col_asignacion].astype(str).str.contains("cuenta de mayor", case=False, na=False)].copy()

            # Limpieza y preparación
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
            # CLASIFICACIÓN DE DISTRIBUIDORAS (NUEVA REGLA GLOBAL)
            # =========================================================
            mapeo_referencias_dist = {
                "11760923": "Dist Acopi", "11761277": "Dist Acopi", "11761350": "Dist Buga", 
                "11831583": "Dist Dosquebradas", "15536188": "Dist Pasto", "17968405": "VENTA EN LINEA"
                # (Se mantienen tus referencias fijas originales acortadas por brevedad en el bloque)
            }

            def clasificar_distribuidora(row):
                # Ahora escaneamos Referencia, Asignación y Texto simultáneamente
                ref_val = str(row.get(col_referencia, "")).strip().upper()
                asig_val = str(row.get(col_asignacion, "")).strip().upper()
                texto_val = str(row.get(col_texto, "")).strip().upper()
                
                texto_global = f"{ref_val} {asig_val} {texto_val}"
                ref_limpia = re.sub(r'\.0$', '', ref_val)
                
                if ref_limpia in mapeo_referencias_dist: return mapeo_referencias_dist[ref_limpia]
                
                if 'DOSQ' in texto_global or 'D504' in texto_global: return 'Dist Dosquebradas'
                if 'ACOPI' in texto_global or 'D503' in texto_global: return 'Dist Acopi'
                if 'PASTO' in texto_global or 'D505' in texto_global: return 'Dist Pasto'
                if 'BUGA' in texto_global or 'D502' in texto_global: return 'Dist Buga'
                
                # Extracción de rangos en CUALQUIER columna de texto
                numeros = re.findall(r'\b\d{4}\b', texto_global)
                for n in numeros:
                    num = int(n)
                    if 2000 <= num <= 2999: return 'Dist Buga'
                    if 3000 <= num <= 3999: return 'Dist Acopi'
                    if 4000 <= num <= 4999: return 'Dist Dosquebradas'
                    if 6000 <= num <= 6999: return 'Dist Pasto'
                        
                return 'Sin clasificar'

            # Aplicamos la clasificación a TODAS las filas (40 y 50) para cruzar rangos
            df['Distribuidora'] = df.apply(clasificar_distribuidora, axis=1)

            def set_estado(indices, estado): df.loc[df['ID_Temp'].isin(indices), 'Estado_Conciliacion'] = estado
            def set_comentarios(dic):
                for i, txt in dic.items(): df.loc[df['ID_Temp'] == i, 'Comentario'] = txt
            def resumen_docs(sub_df): return ", ".join(str(int(d)) for d in sub_df[col_doc].tolist())
            def es_valor_redondo(v): return (v % multiplo_redondo == 0) and v > 0

            df_40 = df[df[col_clave] == '40'].copy()
            df_50 = df[df[col_clave] == '50'].copy()

            # 1A: Cruce exacto por referencia
            df_40['T1'] = df_40.groupby([col_banco, 'Abs_Importe', col_fecha, col_referencia]).cumcount()
            df_50['T1'] = df_50.groupby([col_banco, 'Abs_Importe', col_fecha, col_referencia]).cumcount()
            c1 = pd.merge(df_40, df_50, on=[col_banco, 'Abs_Importe', col_fecha, col_referencia, 'T1'], suffixes=('_40', '_50'))

            df_40['T2'] = df_40.groupby([col_banco, 'Abs_Importe', col_fecha, col_asignacion]).cumcount()
            df_50['T2'] = df_50.groupby([col_banco, 'Abs_Importe', col_fecha, col_referencia]).cumcount()
            c2 = pd.merge(df_40, df_50, left_on=[col_banco, 'Abs_Importe', col_fecha, col_asignacion, 'T2'], right_on=[col_banco, 'Abs_Importe', col_fecha, col_referencia, 'T2'], suffixes=('_40', '_50'))

            ind_r1 = set(c1['ID_Temp_40']) | set(c1['ID_Temp_50']) | set(c2['ID_Temp_40']) | set(c2['ID_Temp_50'])
            set_estado(ind_r1, 'Conciliado - Cruce exacto (Referencia)')
            comentarios_r1 = {}
            for c in (c1, c2):
                for _, r in c.iterrows():
                    comentarios_r1[r['ID_Temp_40']] = f"Cruce exacto con Doc. {int(r[col_doc + '_50'])}"
                    comentarios_r1[r['ID_Temp_50']] = f"Cruce exacto con Doc. {int(r[col_doc + '_40'])}"
            set_comentarios(comentarios_r1)

            # 1B: Cruce exacto Ref. Limpia
            df_p0 = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_p0['Asig_limpia'] = df_p0[col_asignacion].astype(str).str.extract(r'(\d+)')[0]
            df_p0['Ref_limpia'] = df_p0[col_referencia].astype(str).str.extract(r'(\d+)')[0]

            d40b = df_p0[df_p0[col_clave] == '40'].drop(columns=['Ref_limpia'])
            d50b = df_p0[df_p0[col_clave] == '50'].drop(columns=['Asig_limpia'])
            d40b['T_b'] = d40b.groupby([col_banco, 'Abs_Importe', col_fecha, 'Asig_limpia']).cumcount()
            d50b['T_b'] = d50b.groupby([col_banco, 'Abs_Importe', col_fecha, 'Ref_limpia']).cumcount()
            
            c1b = pd.merge(d40b, d50b, left_on=[col_banco, 'Abs_Importe', col_fecha, 'Asig_limpia', 'T_b'], right_on=[col_banco, 'Abs_Importe', col_fecha, 'Ref_limpia', 'T_b'], suffixes=('_40', '_50'))
            c1b = c1b[c1b['Asig_limpia'].notna() & (c1b['Asig_limpia'] != '')]

            ind_r1b = set(c1b['ID_Temp_40']) | set(c1b['ID_Temp_50'])
            set_estado(ind_r1b, 'Conciliado - Cruce exacto (Ref. limpia)')
            comentarios_r1b = {}
            for _, r in c1b.iterrows():
                comentarios_r1b[r['ID_Temp_40']] = f"Cruce exacto con Doc. {int(r[col_doc + '_50'])} (Ref. limpia)"
                comentarios_r1b[r['ID_Temp_50']] = f"Cruce exacto con Doc. {int(r[col_doc + '_40'])} (Ref. limpia)"
            set_comentarios(comentarios_r1b)

            # =========================================================
            # NUEVO: 1C — CRUCE POR RANGO/DISTRIBUIDORA
            # =========================================================
            df_p1c = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_p1c = df_p1c[df_p1c['Distribuidora'] != 'Sin clasificar'] # Solo los que tienen sede identificada
            
            d40c = df_p1c[df_p1c[col_clave] == '40']
            d50c = df_p1c[df_p1c[col_clave] == '50']
            
            d40c['T_c'] = d40c.groupby([col_banco, 'Abs_Importe', col_fecha, 'Distribuidora']).cumcount()
            d50c['T_c'] = d50c.groupby([col_banco, 'Abs_Importe', col_fecha, 'Distribuidora']).cumcount()
            
            c1c = pd.merge(d40c, d50c, on=[col_banco, 'Abs_Importe', col_fecha, 'Distribuidora', 'T_c'], suffixes=('_40', '_50'))
            
            ind_r1c = set(c1c['ID_Temp_40']) | set(c1c['ID_Temp_50'])
            set_estado(ind_r1c, 'Conciliado - Cruce exacto por Rango/Sede')
            comentarios_r1c = {}
            for _, r in c1c.iterrows():
                com = f"Cruce resuelto por rango de Sede ({r['Distribuidora']})"
                comentarios_r1c[r['ID_Temp_40']] = f"{com} - Doc. asociado {int(r[col_doc + '_50'])}"
                comentarios_r1c[r['ID_Temp_50']] = f"{com} - Doc. asociado {int(r[col_doc + '_40'])}"
            set_comentarios(comentarios_r1c)

            # 2: Cruce único sin referencia
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
                comentarios_r2[r['ID_Temp_40']] = f"Cruce único con Doc. {int(r[col_doc + '_50'])}"
                comentarios_r2[r['ID_Temp_50']] = f"Cruce único con Doc. {int(r[col_doc + '_40'])}"
            set_comentarios(comentarios_r2)

            # 2B: Grupos Ambiguos
            ambiguos40 = df_p40[(df_p40['n40'] > 1) & (~df_p40['ID_Temp'].isin(ind_r2))]
            ambiguos50 = df_p50[(df_p50['n50'] > 1) & (~df_p50['ID_Temp'].isin(ind_r2))]
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
                            comentarios_r2d[r40['ID_Temp']] = f"Desempate por FIFO (Valor redondo) con Doc. {int(r50[col_doc])}"
                            comentarios_r2d[r50['ID_Temp']] = f"Desempate por FIFO (Valor redondo) con Doc. {int(r40[col_doc])}"
                    else:
                        for _, r in sub40_ord.iterrows():
                            ind_amb.add(r['ID_Temp']); comentarios_amb[r['ID_Temp']] = f"Valor redondo ambiguo: {len(sub40_ord)} docs vs {len(sub50_ord)}"
                        for _, r in sub50_ord.iterrows():
                            ind_amb.add(r['ID_Temp']); comentarios_amb[r['ID_Temp']] = f"Valor redondo ambiguo: {len(sub50_ord)} docs vs {len(sub40_ord)}"
                else:
                    for _, r in sub40.iterrows():
                        ind_amb.add(r['ID_Temp']); comentarios_amb[r['ID_Temp']] = f"{len(sub50)} posibles cruces de mismo banco/importe/fecha"
                    for _, r in sub50.iterrows():
                        ind_amb.add(r['ID_Temp']); comentarios_amb[r['ID_Temp']] = f"{len(sub40)} posibles cruces de mismo banco/importe/fecha"

            set_estado(ind_r2d, 'Sugerencia fuerte: Emparejado por FIFO')
            set_comentarios(comentarios_r2d)
            set_estado(ind_amb, 'Múltiples candidatos sin referencia clara')
            set_comentarios(comentarios_amb)

            # =========================================================
            # MEJORA: ALERTAS Y VALIDACIONES (FECHA, BANCO, VALOR)
            # =========================================================
            df_pend = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_pend['Ref_Limpia'] = df_pend[col_referencia].astype(str).str.strip().str.lower()
            df_validos = df_pend[~df_pend['Ref_Limpia'].isin(['nan', '', 'none', '0', '/'])].copy()

            df_40n = df_validos[df_validos[col_clave] == '40'].copy()
            df_50n = df_validos[df_validos[col_clave] == '50'].copy()

            # 9A: Alerta de fecha (Validando periodo contable)
            sA = pd.merge(df_40n, df_50n, on=[col_banco, 'Abs_Importe', 'Ref_Limpia'], suffixes=('_40', '_50'))
            sA['Dif_Dias'] = (sA['Fecha_Calc_40'] - sA['Fecha_Calc_50']).dt.days.abs()
            sA = sA[(sA['Dif_Dias'] > 0) & (sA['Dif_Dias'] <= tol_dias)]
            sA = sA.drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            ind_A = set(sA['ID_Temp_40']) | set(sA['ID_Temp_50'])
            set_estado(ind_A, 'Revisar - Diferencia de fecha valor')
            com_A = {}
            for _, r in sA.iterrows():
                f40, f50 = r['Fecha_Calc_40'], r['Fecha_Calc_50']
                # Validamos si el salto de días provocó un cambio de mes/año (Periodo Contable)
                mismo_periodo = (f40.month == f50.month) and (f40.year == f50.year)
                alerta_periodo = "MISMO PERIODO" if mismo_periodo else "¡ALERTA! DIFERENTE PERIODO CONTABLE"
                
                com_txt = f"Coincide Ref/Banco/Importe. Fechas difieren {int(r['Dif_Dias'])} día(s) ({alerta_periodo})."
                com_A[r['ID_Temp_40']] = f"{com_txt} Doc. sugerido: {int(r[col_doc+'_50'])}"
                com_A[r['ID_Temp_50']] = f"{com_txt} Doc. sugerido: {int(r[col_doc+'_40'])}"
            set_comentarios(com_A)

            df_40n = df_40n[~df_40n['ID_Temp'].isin(ind_A)]
            df_50n = df_50n[~df_50n['ID_Temp'].isin(ind_A)]

            # 9B: Banco y 9C: Valor
            sB = pd.merge(df_40n, df_50n, on=['Abs_Importe', col_fecha, 'Ref_Limpia'], suffixes=('_40', '_50'))
            sB = sB[sB[f'{col_banco}_40'] != sB[f'{col_banco}_50']].drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')
            ind_B = set(sB['ID_Temp_40']) | set(sB['ID_Temp_50'])
            set_estado(ind_B, 'Revisar - Reclasificación de banco')
            com_B = {}
            for _, r in sB.iterrows():
                com_B[r['ID_Temp_40']] = f"Coincide Ref/Importe/Fecha, pero difiere banco. Doc sugerido: {int(r[col_doc+'_50'])}"
                com_B[r['ID_Temp_50']] = f"Coincide Ref/Importe/Fecha, pero difiere banco. Doc sugerido: {int(r[col_doc+'_40'])}"
            set_comentarios(com_B)

            df_40n = df_40n[~df_40n['ID_Temp'].isin(ind_B)]
            df_50n = df_50n[~df_50n['ID_Temp'].isin(ind_B)]

            sC = pd.merge(df_40n, df_50n, on=[col_banco, col_fecha, 'Ref_Limpia'], suffixes=('_40', '_50'))
            sC['Dif_Valor'] = (sC['Abs_Importe_40'] - sC['Abs_Importe_50']).abs()
            max_importe = sC[['Abs_Importe_40', 'Abs_Importe_50']].max(axis=1)
            sC['Dif_Pct'] = np.where(max_importe == 0, 0, sC['Dif_Valor'] / max_importe)
            sC = sC[(sC['Dif_Valor'] > 0) & ((sC['Dif_Valor'] <= tol_valor_abs) | (sC['Dif_Pct'] <= tol_valor_pct))].sort_values('Dif_Valor').drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            ind_C = set(sC['ID_Temp_40']) | set(sC['ID_Temp_50'])
            set_estado(ind_C, 'Revisar - Diferencia de valor')
            com_C = {}
            for _, r in sC.iterrows():
                com_C[r['ID_Temp_40']] = f"Coincide Ref/Banco/Fecha, difiere ${r['Dif_Valor']:,.0f}. Doc: {int(r[col_doc+'_50'])}"
                com_C[r['ID_Temp_50']] = f"Coincide Ref/Banco/Fecha, difiere ${r['Dif_Valor']:,.0f}. Doc: {int(r[col_doc+'_40'])}"
            set_comentarios(com_C)

            # --- Lo pendiente ---
            sin_pista = df['Estado_Conciliacion'] == 'Pendiente'
            df.loc[sin_pista & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia ni pista encontrada - validación manual'

            total_filas_entrada = filas_antes
            total_filas_salida = len(df) + len(filas_descartadas)
            cuadre_ok = total_filas_entrada == total_filas_salida

            df_final = df.drop(columns=['ID_Temp', 'Abs_Importe', 'Fecha_Calc', 'Distribuidora'], errors='ignore')
            for col_f in [c for c in df_final.columns if 'fe.' in c.lower() or 'fecha' in c.lower() or 'fe-' in c.lower()]:
                df_final[col_f] = pd.to_datetime(df_final[col_f], errors='coerce').dt.strftime('%d/%m/%Y')

            # =========================================================
            # SEMÁFORO SIMPLIFICADO (3 NIVELES)
            # =========================================================
            def resaltar_conciliados(row):
                est = str(row['Estado_Conciliacion']).lower()
                # 🟢 VERDE: Todo lo que es cruce seguro
                if 'conciliado' in est:
                    return ['background-color: #D4EFDF'] * len(row)
                # 🟡 AMARILLO: Alertas con sugerencia clara (fecha, valor, banco, fifo)
                elif 'revisar' in est or 'sugerencia fuerte' in est:
                    return ['background-color: #FCF3CF'] * len(row)
                # 🔴 ROJO CLARO: Múltiples candidatos o sin pista (requiere revisión manual)
                elif 'múltiples' in est or 'multiples' in est or 'pendiente' in est:
                    return ['background-color: #F5B7B1'] * len(row)
                return [''] * len(row)

            # Exportación
            output = io.BytesIO()
            bancos_unicos = [b for b in df_final[col_banco].unique() if str(b).strip().lower() not in ('', 'nan')]

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                resumen = df_final['Estado_Conciliacion'].value_counts().reset_index()
                resumen.columns = ['Estado', 'Cantidad de registros']
                resumen.to_excel(writer, index=False, sheet_name='RESUMEN')

                for banco in bancos_unicos:
                    df_banco = df_final[df_final[col_banco] == banco].copy().sort_values(by=col_importe, ascending=True)
                    nombre_pestana = re.sub(r'[/\\:?*\[\]]', '-', str(banco)[:31])
                    if not nombre_pestana.strip() or nombre_pestana.lower() == 'nan': nombre_pestana = "Sin_Banco_Asignado"
                    df_banco.style.apply(resaltar_conciliados, axis=1).to_excel(writer, index=False, sheet_name=nombre_pestana)

                df_novedades = df_final[~df_final['Estado_Conciliacion'].str.contains('Conciliado', na=False)].copy()
                if not df_novedades.empty:
                    df_novedades.sort_values(by=['Estado_Conciliacion', col_importe], ascending=[True, True]).style.apply(resaltar_conciliados, axis=1).to_excel(writer, index=False, sheet_name='NOVEDADES_Y_ALERTAS')

                if not filas_descartadas.empty:
                    filas_descartadas.to_excel(writer, index=False, sheet_name='DESCARTADAS_SIN_DOC_O_CT')

            st.success("¡Conciliación terminada! Reglas de periodo aplicadas y ambigüedades resueltas por rango.")
            if not cuadre_ok: st.warning("⚠️ Alerta de integridad: revisa la pestaña DESCARTADAS.")

            st.download_button(
                label="📥 Descargar Excel con Resultados",
                data=output.getvalue(),
                file_name="Conciliacion_Integral_Resultados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error técnico detectado: {e}")
