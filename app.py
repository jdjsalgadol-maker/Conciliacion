import streamlit as st
import pandas as pd
import numpy as np
import io
import re

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Conciliación Integral @JuanS", layout="wide")

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
    "solo marca como *Conciliado* lo que tiene evidencia inequívoca (referencia exacta, o cruce "
    "único sin ambigüedad). Todo lo demás queda resaltado con color y una **sugerencia textual** "
    "para tu revisión manual, en lugar de arriesgarse a un cruce incorrecto."
)

# --- PARÁMETROS AJUSTABLES ---
with st.expander("⚙️ Parámetros de tolerancia para sugerencias (alertas)"):
    tol_dias = st.slider("Días máximos de diferencia para alertar 'error de fecha'", 1, 10, 3)
    tol_valor_abs = st.number_input("Diferencia absoluta máxima de valor para alertar ($)", min_value=1, value=5000, step=100)
    tol_valor_pct = st.number_input("Diferencia relativa máxima de valor para alertar (%)", min_value=0.01, value=0.5, step=0.01) / 100

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
            # 2. MAPEO SEGURO Y DINÁMICO DE COLUMNAS (acepta 2 nomenclaturas)
            # =========================================================
            col_asignacion = 'Asignación' if 'Asignación' in df.columns else 'Asignaión'
            col_referencia = 'Referencia'
            col_clave = 'Clave contabiliz.' if 'Clave contabiliz.' in df.columns else 'CT'
            col_fecha = 'Fecha valor' if 'Fecha valor' in df.columns else 'Fe-valor'
            col_importe = 'Importe en moneda local' if 'Importe en moneda local' in df.columns else 'Importe en ML'
            col_banco = 'Clave referencia 3'
            col_doc = 'Nº documento' if 'Nº documento' in df.columns else 'Nº doc.'

            columnas_requeridas = [col_referencia, col_clave, col_fecha, col_importe, col_banco, col_doc]
            faltantes = [c for c in columnas_requeridas if c not in df.columns]
            if faltantes:
                st.error(f"No se encontraron estas columnas obligatorias en el archivo: {faltantes}")
                st.stop()

            # =========================================================
            # 3. AUTOCOMPLETADO DE BANCO POR CUENTA DE MAYOR
            #    (se ejecuta ANTES de ordenar, para respetar los bloques
            #    de banco tal como vienen en el extracto original)
            # =========================================================
            mapeo_cuentas_banco = {
                "1110056101": "BANCO DE BOGOTA",
                "1110056201": "BANCO DAVIBANK S.A.",
                "1110056301": "BANCOLOMBIA S.A.",
                "1110056401": "BANCO CAJA SOCIAL S.",
                "1110056501": "BANCO DAVIVIENDA S.A",
                "1110056601": "BANCO BILBAO VIZCAYA",
                "1110056701": "BANCO AGRARIO DE COL",
                "1120055001": "BANCO COMERCIAL AV V",
                "1120055101": "BANCO DE OCCIDENTE",
                "1120055301": "BANCO GNB SUDAMERIS",
            }

            current_bank = None
            bancos_completados = []
            for _, row in df.iterrows():
                asig_val = str(row.get(col_asignacion, "")) if col_asignacion in df.columns else ""
                banco_val = row.get(col_banco, None)

                if "cuenta de mayor" in asig_val.lower():
                    match_cuenta = re.search(r'(\d{6,})', asig_val)
                    cuenta_num = match_cuenta.group(1) if match_cuenta else None
                    # Si la cuenta no está en el diccionario, NUNCA se asume el banco anterior
                    # a ciegas: se etiqueta explícitamente con el número de cuenta para que
                    # quede visible y no genere un error de clasificación silencioso.
                    if cuenta_num:
                        current_bank = mapeo_cuentas_banco.get(cuenta_num, f"CUENTA {cuenta_num} (sin mapear)")

                if pd.notnull(banco_val) and str(banco_val).strip().lower() not in ("", "nan"):
                    current_bank = str(banco_val).strip()
                    bancos_completados.append(current_bank)
                else:
                    bancos_completados.append(current_bank)

            df[col_banco] = bancos_completados

            if col_asignacion in df.columns:
                df = df[~df[col_asignacion].astype(str).str.contains("cuenta de mayor", case=False, na=False)].copy()

            # =========================================================
            # 4. LIMPIEZA Y ORDENAMIENTO FIFO
            # =========================================================
            df[col_doc] = pd.to_numeric(df[col_doc], errors='coerce')
            filas_antes = len(df)
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

            def set_estado(indices, estado):
                df.loc[df['ID_Temp'].isin(indices), 'Estado_Conciliacion'] = estado

            def set_comentarios(dic_comentarios):
                for id_temp, texto in dic_comentarios.items():
                    df.loc[df['ID_Temp'] == id_temp, 'Comentario'] = texto

            df_40 = df[df[col_clave] == '40']
            df_50 = df[df[col_clave] == '50']

            # =========================================================
            # 5. NIVEL 1 — CRUCE EXACTO POR REFERENCIA (100% seguro)
            #    Banco + Importe absoluto + Fecha + Referencia/Asignación
            #    cruzadas. Es el único tipo de cruce que se da por
            #    conciliado sin ninguna duda posible.
            # =========================================================
            c1 = pd.merge(df_40, df_50,
                          left_on=[col_banco, 'Abs_Importe', col_fecha, col_referencia],
                          right_on=[col_banco, 'Abs_Importe', col_fecha, col_referencia],
                          suffixes=('_40', '_50'))
            c2 = pd.merge(df_40, df_50,
                          left_on=[col_banco, 'Abs_Importe', col_fecha, col_asignacion],
                          right_on=[col_banco, 'Abs_Importe', col_fecha, col_referencia],
                          suffixes=('_40', '_50'))
            c3 = pd.merge(df_40, df_50,
                          left_on=[col_banco, 'Abs_Importe', col_fecha, col_referencia],
                          right_on=[col_banco, 'Abs_Importe', col_fecha, col_asignacion],
                          suffixes=('_40', '_50'))

            ind_r1 = (set(c1['ID_Temp_40']) | set(c1['ID_Temp_50']) |
                      set(c2['ID_Temp_40']) | set(c2['ID_Temp_50']) |
                      set(c3['ID_Temp_40']) | set(c3['ID_Temp_50']))
            set_estado(ind_r1, 'Conciliado - Cruce exacto (Referencia)')

            comentarios_r1 = {}
            for c in (c1, c2, c3):
                for _, r in c.iterrows():
                    comentarios_r1[r['ID_Temp_40']] = f"Cruce exacto con Doc. {int(r[col_doc + '_50'])} (misma referencia, banco e importe)"
                    comentarios_r1[r['ID_Temp_50']] = f"Cruce exacto con Doc. {int(r[col_doc + '_40'])} (misma referencia, banco e importe)"
            set_comentarios(comentarios_r1)

            # =========================================================
            # 6. NIVEL 2 — CRUCE ÚNICO SIN REFERENCIA
            #    Mismo banco + importe absoluto + fecha, y SOLO cuando
            #    existe exactamente un movimiento débito (40) y uno
            #    crédito (50) en ese grupo. Si hay más de un candidato
            #    posible, NO se concilia a ciegas (evita el error de
            #    emparejar el documento equivocado): se deja como
            #    sugerencia ambigua para revisión manual (nivel 2B).
            # =========================================================
            df_p = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            grp_cols = [col_banco, 'Abs_Importe', col_fecha]

            df_p40 = df_p[df_p[col_clave] == '40'].copy()
            df_p50 = df_p[df_p[col_clave] == '50'].copy()

            cnt40 = df_p40.groupby(grp_cols)['ID_Temp'].transform('count')
            cnt50 = df_p50.groupby(grp_cols)['ID_Temp'].transform('count')
            df_p40['n40'] = cnt40
            df_p50['n50'] = cnt50

            # Solo se cruzan grupos donde AMBOS lados tienen exactamente 1 movimiento
            unicos40 = df_p40[df_p40['n40'] == 1]
            unicos50 = df_p50[df_p50['n50'] == 1]
            c_unico = pd.merge(unicos40, unicos50, on=grp_cols, suffixes=('_40', '_50'))

            ind_r2 = set(c_unico['ID_Temp_40']) | set(c_unico['ID_Temp_50'])
            set_estado(ind_r2, 'Conciliado - Cruce unico sin referencia')

            comentarios_r2 = {}
            for _, r in c_unico.iterrows():
                comentarios_r2[r['ID_Temp_40']] = f"Cruce único con Doc. {int(r[col_doc + '_50'])} (mismo banco/fecha/importe, sin referencia disponible, sin ambigüedad)"
                comentarios_r2[r['ID_Temp_50']] = f"Cruce único con Doc. {int(r[col_doc + '_40'])} (mismo banco/fecha/importe, sin referencia disponible, sin ambigüedad)"
            set_comentarios(comentarios_r2)

            # --- Nivel 2B: grupos AMBIGUOS (más de un candidato posible) ---
            ambiguos40 = df_p40[(df_p40['n40'] > 1) & (~df_p40['ID_Temp'].isin(ind_r2))]
            ambiguos50 = df_p50[(df_p50['n50'] > 1) & (~df_p50['ID_Temp'].isin(ind_r2))]

            def resumen_docs(sub_df):
                return ", ".join(str(int(d)) for d in sub_df[col_doc].tolist())

            ind_amb = set()
            comentarios_amb = {}
            for grupo, sub40 in ambiguos40.groupby(grp_cols):
                sub50 = ambiguos50[(ambiguos50[col_banco] == grupo[0]) &
                                    (ambiguos50['Abs_Importe'] == grupo[1]) &
                                    (ambiguos50[col_fecha] == grupo[2])]
                if len(sub50) == 0:
                    continue
                docs40 = resumen_docs(sub40)
                docs50 = resumen_docs(sub50)
                for _, r in sub40.iterrows():
                    ind_amb.add(r['ID_Temp'])
                    comentarios_amb[r['ID_Temp']] = (f"Existen {len(sub50)} posibles cruces sin referencia con el mismo "
                                                      f"banco/fecha/importe (Docs. candidatos: {docs50}) - validar manualmente cuál corresponde")
                for _, r in sub50.iterrows():
                    ind_amb.add(r['ID_Temp'])
                    comentarios_amb[r['ID_Temp']] = (f"Existen {len(sub40)} posibles cruces sin referencia con el mismo "
                                                      f"banco/fecha/importe (Docs. candidatos: {docs40}) - validar manualmente cuál corresponde")

            set_estado(ind_amb, 'Sugerencia: Múltiples candidatos sin referencia')
            set_comentarios(comentarios_amb)

            # =========================================================
            # 7. NIVEL 3 — SUGERENCIAS BASADAS EN REFERENCIA VÁLIDA
            #    (para lo que sigue pendiente tras los niveles 1 y 2)
            # =========================================================
            df_pend = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_pend['Ref_Limpia'] = df_pend[col_referencia].astype(str).str.strip().str.lower()
            df_validos = df_pend[~df_pend['Ref_Limpia'].isin(['nan', '', 'none', '0', '/'])].copy()

            df_40n = df_validos[df_validos[col_clave] == '40'].copy()
            df_50n = df_validos[df_validos[col_clave] == '50'].copy()

            # --- 7A: Validar con error de fecha (misma ref/banco/importe, fecha distinta) ---
            sA = pd.merge(df_40n, df_50n, on=[col_banco, 'Abs_Importe', 'Ref_Limpia'], suffixes=('_40', '_50'))
            sA['Dif_Dias'] = (sA['Fecha_Calc_40'] - sA['Fecha_Calc_50']).dt.days.abs()
            sA = sA[(sA['Dif_Dias'] > 0) & (sA['Dif_Dias'] <= tol_dias)]
            sA = sA.drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            ind_A = set(sA['ID_Temp_40']) | set(sA['ID_Temp_50'])
            set_estado(ind_A, 'Validar con error de fecha')
            com_A = {}
            for _, r in sA.iterrows():
                com_A[r['ID_Temp_40']] = f"Misma referencia/banco/importe que Doc. {int(r[col_doc+'_50'])}, pero difieren {int(r['Dif_Dias'])} día(s) en la fecha"
                com_A[r['ID_Temp_50']] = f"Misma referencia/banco/importe que Doc. {int(r[col_doc+'_40'])}, pero difieren {int(r['Dif_Dias'])} día(s) en la fecha"
            set_comentarios(com_A)

            df_40n = df_40n[~df_40n['ID_Temp'].isin(ind_A)]
            df_50n = df_50n[~df_50n['ID_Temp'].isin(ind_A)]

            # --- 7B: Sugerencia de reclasificación de banco (misma ref/importe/fecha, banco distinto) ---
            sB = pd.merge(df_40n, df_50n, on=['Abs_Importe', col_fecha, 'Ref_Limpia'], suffixes=('_40', '_50'))
            sB = sB[sB[f'{col_banco}_40'] != sB[f'{col_banco}_50']]
            sB = sB.drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            ind_B = set(sB['ID_Temp_40']) | set(sB['ID_Temp_50'])
            set_estado(ind_B, 'Sugerencia: Reclasificación de banco')
            com_B = {}
            for _, r in sB.iterrows():
                com_B[r['ID_Temp_40']] = f"Misma referencia/importe/fecha que Doc. {int(r[col_doc+'_50'])}, pero aparece en banco '{r[col_banco+'_50']}' en vez de '{r[col_banco+'_40']}'"
                com_B[r['ID_Temp_50']] = f"Misma referencia/importe/fecha que Doc. {int(r[col_doc+'_40'])}, pero aparece en banco '{r[col_banco+'_40']}' en vez de '{r[col_banco+'_50']}'"
            set_comentarios(com_B)

            df_40n = df_40n[~df_40n['ID_Temp'].isin(ind_B)]
            df_50n = df_50n[~df_50n['ID_Temp'].isin(ind_B)]

            # --- 7C: Revisar diferencia de valor (misma ref/banco/fecha, importe distinto dentro de tolerancia) ---
            sC = pd.merge(df_40n, df_50n, on=[col_banco, col_fecha, 'Ref_Limpia'], suffixes=('_40', '_50'))
            sC['Dif_Valor'] = (sC['Abs_Importe_40'] - sC['Abs_Importe_50']).abs()
            sC['Dif_Pct'] = sC['Dif_Valor'] / sC[['Abs_Importe_40', 'Abs_Importe_50']].max(axis=1)
            sC = sC[(sC['Dif_Valor'] > 0) & ((sC['Dif_Valor'] <= tol_valor_abs) | (sC['Dif_Pct'] <= tol_valor_pct))]
            sC = sC.sort_values('Dif_Valor').drop_duplicates('ID_Temp_40').drop_duplicates('ID_Temp_50')

            ind_C = set(sC['ID_Temp_40']) | set(sC['ID_Temp_50'])
            set_estado(ind_C, 'Revisar diferencia de valor')
            com_C = {}
            for _, r in sC.iterrows():
                com_C[r['ID_Temp_40']] = f"Misma referencia/banco/fecha que Doc. {int(r[col_doc+'_50'])}, con diferencia de ${r['Dif_Valor']:,.0f} ({r['Dif_Pct']*100:.2f}%)"
                com_C[r['ID_Temp_50']] = f"Misma referencia/banco/fecha que Doc. {int(r[col_doc+'_40'])}, con diferencia de ${r['Dif_Valor']:,.0f} ({r['Dif_Pct']*100:.2f}%)"
            set_comentarios(com_C)

            # --- Lo que sigue pendiente sin ninguna pista ---
            sin_pista = df['Estado_Conciliacion'] == 'Pendiente'
            df.loc[sin_pista & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia ni sugerencia encontrada - requiere revisión manual completa'

            # =========================================================
            # 8. LIMPIEZA FINAL Y FORMATO DE FECHAS
            # =========================================================
            df_final = df.drop(columns=['ID_Temp', 'Abs_Importe', 'Fecha_Calc'], errors='ignore')

            columnas_fecha = [c for c in df_final.columns if 'fe.' in c.lower() or 'fecha' in c.lower() or 'fe-' in c.lower()]
            for col_f in columnas_fecha:
                df_final[col_f] = pd.to_datetime(df_final[col_f], errors='coerce').dt.strftime('%d/%m/%Y')

            # =========================================================
            # FUNCIÓN DE COLORES (resalta según el estado)
            # =========================================================
            def resaltar_conciliados(row):
                est = str(row['Estado_Conciliacion']).lower()
                if 'cruce exacto' in est:
                    return ['background-color: #D4EFDF'] * len(row)          # Verde - 100% seguro
                elif 'cruce unico' in est:
                    return ['background-color: #D6EAF8'] * len(row)          # Azul - seguro, sin referencia
                elif 'múltiples candidatos' in est or 'multiples candidatos' in est:
                    return ['background-color: #FDEBD0'] * len(row)          # Durazno - ambiguo, revisar
                elif 'error de fecha' in est:
                    return ['background-color: #FCF3CF'] * len(row)          # Amarillo - alerta fecha
                elif 'reclasificación' in est or 'reclasificacion' in est:
                    return ['background-color: #F5B7B1'] * len(row)          # Rojo claro - alerta banco
                elif 'diferencia de valor' in est:
                    return ['background-color: #FAD7A1'] * len(row)          # Naranja - alerta valor
                return [''] * len(row)

            # =========================================================
            # 9. EXPORTACIÓN: 1 PESTAÑA POR BANCO + PESTAÑA CONSOLIDADA
            # =========================================================
            output = io.BytesIO()
            bancos_unicos = df_final[col_banco].unique()

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
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

                # Pestaña consolidada de todo lo que NO quedó 100% conciliado
                df_novedades = df_final[~df_final['Estado_Conciliacion'].str.contains('Conciliado', na=False)].copy()
                if not df_novedades.empty:
                    df_novedades = df_novedades.sort_values(by=['Estado_Conciliacion', col_importe], ascending=[True, True])
                    styled_novedades = df_novedades.style.apply(resaltar_conciliados, axis=1)
                    styled_novedades.to_excel(writer, index=False, sheet_name='NOVEDADES_Y_ALERTAS')

                # Pestaña resumen
                resumen = df_final['Estado_Conciliacion'].value_counts().reset_index()
                resumen.columns = ['Estado', 'Cantidad de registros']
                resumen.to_excel(writer, index=False, sheet_name='RESUMEN')

            # =========================================================
            # 10. INTERFAZ Y DESCARGA
            # =========================================================
            st.success("¡Conciliación Integral terminada! Todo lo que no es 100% seguro quedó marcado con color y comentario.")

            conciliados_exactos = len(ind_r1)
            conciliados_unicos = len(ind_r2)
            alertas = len(ind_amb) + len(ind_A) + len(ind_B) + len(ind_C)
            pendientes_sin_pista = len(df_final) - conciliados_exactos - conciliados_unicos - alertas

            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("Bancos procesados", len(bancos_unicos))
            col2.metric("Conciliado exacto", conciliados_exactos)
            col3.metric("Conciliado único", conciliados_unicos)
            col4.metric("Alertas (revisar)", alertas)
            col5.metric("Sin ninguna pista", pendientes_sin_pista)

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
