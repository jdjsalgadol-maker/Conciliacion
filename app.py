import streamlit as st
import pandas as pd
import io

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Conciliación General @JuanS", layout="wide")

hide_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
"""
st.markdown(hide_style, unsafe_allow_html=True)

st.title("🏦 Conciliación General Multibanco 🤖")
st.write("Sube tu archivo consolidado. El sistema aplicará emparejamiento exacto y análisis integral de alertas (fechas, bancos y diferencias de valor por ajuste al peso).")

archivo_subido = st.file_uploader("Selecciona el archivo de Excel o CSV", type=['xlsx', 'csv'])

if archivo_subido is not None:
    try:
        with st.spinner("Leyendo archivo, unificando datos y analizando alertas..."):
            
            # --- 1. LECTURA DINÁMICA A PRUEBA DE PESTAÑAS ---
            if archivo_subido.name.lower().endswith('.csv'):
                df = pd.read_csv(archivo_subido)
            else:
                diccionario_hojas = pd.read_excel(archivo_subido, sheet_name=None)
                df = pd.concat(diccionario_hojas.values(), ignore_index=True)
            
            # Limpiar nombres de columnas
            df.columns = df.columns.str.strip()
            
            # --- 2. MAPEO SEGURO Y DINÁMICO DE COLUMNAS ---
            col_asignacion = 'Asignación' if 'Asignación' in df.columns else 'Asignaión' 
            col_referencia = 'Referencia'
            col_clave = 'Clave contabiliz.' if 'Clave contabiliz.' in df.columns else 'CT'
            col_fecha = 'Fecha valor' if 'Fecha valor' in df.columns else 'Fe-valor'
            col_importe = 'Importe en moneda local' if 'Importe en moneda local' in df.columns else 'Importe en ML'
            col_banco = 'Clave referencia 3'
            col_doc = 'Nº documento' if 'Nº documento' in df.columns else 'Nº doc.'

            if col_doc not in df.columns:
                st.error(f"No se encontró la columna de documento ({col_doc}) en ninguna de las pestañas.")
                st.stop()

            # --- 3. AUTOCOMPLETADO DE BANCOS POR CUENTA DE MAYOR ---
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

            for idx, row in df.iterrows():
                asig_val = str(row.get(col_asignacion, ""))
                banco_val = row.get(col_banco, None)

                if "Cuenta de mayor" in asig_val:
                    for cuenta, banco in mapeo_cuentas_banco.items():
                        if cuenta in asig_val:
                            current_bank = banco
                            break

                if pd.notnull(banco_val) and str(banco_val).strip() != "" and str(banco_val).strip().lower() != "nan":
                    current_bank = str(banco_val).strip()
                    bancos_completados.append(current_bank)
                else:
                    bancos_completados.append(current_bank)

            df[col_banco] = bancos_completados

            if col_asignacion in df.columns:
                df = df[~df[col_asignacion].astype(str).str.contains("Cuenta de mayor", case=False, na=False)].copy()

            # --- 4. LIMPIEZA Y ORDENAMIENTO (FIFO) ---
            df[col_doc] = pd.to_numeric(df[col_doc], errors='coerce')
            df = df.dropna(subset=[col_doc, col_clave]).reset_index(drop=True)
            df = df.sort_values(by=[col_doc], ascending=True).reset_index(drop=True)
            
            df['ID_Temp'] = df.index
            df[col_clave] = df[col_clave].astype(str).str.strip().str.replace('.0', '', regex=False)
            df[col_banco] = df[col_banco].astype(str).str.strip()
            
            # Dejar el importe como número puro (Genérico)
            df[col_importe] = pd.to_numeric(df[col_importe], errors='coerce').fillna(0)
            df['Abs_Importe'] = df[col_importe].abs() 
            
            # Guardamos la fecha original para cálculos de días
            df['Fecha_Calc'] = pd.to_datetime(df[col_fecha], errors='coerce')
            df[col_fecha] = df['Fecha_Calc'].dt.date
            df['Estado_Conciliacion'] = 'Pendiente'

            # --- 5. LÓGICA DE CASCADA VECTORIZADA (EXACTA) ---
            df_40 = df[df[col_clave] == '40']
            df_50 = df[df[col_clave] == '50']
            
            c1 = pd.merge(df_40, df_50, left_on=[col_banco, 'Abs_Importe', col_fecha, col_referencia], right_on=[col_banco, 'Abs_Importe', col_fecha, col_referencia], suffixes=('_40', '_50'))
            c2 = pd.merge(df_40, df_50, left_on=[col_banco, 'Abs_Importe', col_fecha, col_asignacion], right_on=[col_banco, 'Abs_Importe', col_fecha, col_referencia], suffixes=('_40', '_50'))
            c3 = pd.merge(df_40, df_50, left_on=[col_banco, 'Abs_Importe', col_fecha, col_referencia], right_on=[col_banco, 'Abs_Importe', col_fecha, col_asignacion], suffixes=('_40', '_50'))
            
            ind_r1 = set(c1['ID_Temp_40']).union(set(c1['ID_Temp_50'])).union(set(c2['ID_Temp_40'])).union(set(c2['ID_Temp_50'])).union(set(c3['ID_Temp_40'])).union(set(c3['ID_Temp_50']))
            df.loc[df['ID_Temp'].isin(ind_r1), 'Estado_Conciliacion'] = 'Conciliado Parte 1'

            df_p = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_p['Turno'] = df_p.groupby([col_banco, 'Abs_Importe', col_fecha, col_clave]).cumcount()
            
            c_n = pd.merge(df_p[df_p[col_clave]=='40'], df_p[df_p[col_clave]=='50'], on=[col_banco, 'Abs_Importe', col_fecha, 'Turno'], suffixes=('_4', '_5'))
            ind_r2 = set(c_n['ID_Temp_4']).union(set(c_n['ID_Temp_5']))
            df.loc[df['ID_Temp'].isin(ind_r2), 'Estado_Conciliacion'] = 'Conciliado Parte 2'

            # --- 6. ANÁLISIS INTEGRAL (NOVEDADES Y ALERTAS) ---
            df_pendientes = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            
            # Limpiamos referencias para evitar cruces falsos con celdas vacías
            df_pendientes['Ref_Limpia'] = df_pendientes[col_referencia].astype(str).str.strip().str.lower()
            df_validos = df_pendientes[~df_pendientes['Ref_Limpia'].isin(['nan', '', 'none', '0', '/'])].copy()
            
            df_40_nov = df_validos[df_validos[col_clave] == '40']
            df_50_nov = df_validos[df_validos[col_clave] == '50']

            # Alerta A: Validar con error de fecha (diferencia de 1 a 3 días, mismo banco e importe)
            s_date = pd.merge(df_40_nov, df_50_nov, on=[col_banco, 'Abs_Importe', 'Ref_Limpia'], suffixes=('_40', '_50'))
            s_date['Dif_Dias'] = (s_date['Fecha_Calc_40'] - s_date['Fecha_Calc_50']).dt.days.abs()
            s_date = s_date[(s_date['Dif_Dias'] > 0) & (s_date['Dif_Dias'] <= 3)]
            s_date = s_date.drop_duplicates(subset=['ID_Temp_40']).drop_duplicates(subset=['ID_Temp_50'])
            
            ind_date = set(s_date['ID_Temp_40']).union(set(s_date['ID_Temp_50']))
            df.loc[df['ID_Temp'].isin(ind_date), 'Estado_Conciliacion'] = 'Validar con error de fecha'

            # Actualizamos excluyendo los que ya marcamos para la siguiente regla
            df_40_nov = df_40_nov[~df_40_nov['ID_Temp'].isin(ind_date)]
            df_50_nov = df_50_nov[~df_50_nov['ID_Temp'].isin(ind_date)]

            # Alerta B: Sugerencia: Reclasificación de banco (mismo importe, fecha y ref, distinto banco)
            s_bank = pd.merge(df_40_nov, df_50_nov, on=['Abs_Importe', col_fecha, 'Ref_Limpia'], suffixes=('_40', '_50'))
            s_bank = s_bank[s_bank[f'{col_banco}_40'] != s_bank[f'{col_banco}_50']]
            s_bank = s_bank.drop_duplicates(subset=['ID_Temp_40']).drop_duplicates(subset=['ID_Temp_50'])
            
            ind_bank = set(s_bank['ID_Temp_40']).union(set(s_bank['ID_Temp_50']))
            df.loc[df['ID_Temp'].isin(ind_bank), 'Estado_Conciliacion'] = 'Sugerencia: Reclasificación de banco'
            
            # Actualizamos excluyendo los que ya marcamos para la siguiente regla
            df_40_nov = df_40_nov[~df_40_nov['ID_Temp'].isin(ind_bank)]
            df_50_nov = df_50_nov[~df_50_nov['ID_Temp'].isin(ind_bank)]

            # Alerta C: Revisar diferencia de valor (margen de hasta 100 pesos, mismo banco, fecha y ref)
            s_val = pd.merge(df_40_nov, df_50_nov, on=[col_banco, col_fecha, 'Ref_Limpia'], suffixes=('_40', '_50'))
            s_val['Dif_Valor'] = (s_val['Abs_Importe_40'] - s_val['Abs_Importe_50']).abs()
            s_val = s_val[(s_val['Dif_Valor'] > 0) & (s_val['Dif_Valor'] <= 100)]
            s_val = s_val.drop_duplicates(subset=['ID_Temp_40']).drop_duplicates(subset=['ID_Temp_50'])
            
            ind_val = set(s_val['ID_Temp_40']).union(set(s_val['ID_Temp_50']))
            df.loc[df['ID_Temp'].isin(ind_val), 'Estado_Conciliacion'] = 'Revisar diferencia de valor'

            # --- 7. LIMPIEZA FINAL Y FORMATO DE FECHAS ---
            df_final = df.drop(columns=['ID_Temp', 'Abs_Importe', 'Turno', 'Fecha_Calc', 'Ref_Limpia'], errors='ignore')

            # Aplicar formato de fecha corta (DD/MM/YYYY)
            columnas_fecha = [c for c in df_final.columns if 'fe.' in c.lower() or 'fecha' in c.lower() or 'fe-' in c.lower()]
            for col_f in columnas_fecha:
                df_final[col_f] = pd.to_datetime(df_final[col_f], errors='coerce').dt.strftime('%d/%m/%Y')

            # --- FUNCION DE COLORES ---
            def resaltar_conciliados(row):
                est = str(row['Estado_Conciliacion'])
                if 'Conciliado Parte 1' in est:
                    return ['background-color: #D4EFDF'] * len(row) # Verde Claro
                elif 'Conciliado Parte 2' in est:
                    return ['background-color: #D6EAF8'] * len(row) # Azul Claro
                elif 'error de fecha' in est.lower():
                    return ['background-color: #FCF3CF'] * len(row) # Amarillo (Alerta Fecha)
                elif 'reclasificación' in est.lower():
                    return ['background-color: #F5B7B1'] * len(row) # Rojo Claro (Alerta Banco)
                elif 'diferencia de valor' in est.lower():
                    return ['background-color: #FAD7A1'] * len(row) # Naranja Claro (Alerta Valor)
                return [''] * len(row)

            # --- 8. EXPORTACIÓN DIVIDIDA POR BANCO + PESTAÑA NOVEDADES ---
            output = io.BytesIO()
            bancos_unicos = df_final[col_banco].unique()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # 8.1 Exportar Bancos individualmente
                for banco in bancos_unicos:
                    df_banco = df_final[df_final[col_banco] == banco].copy()
                    df_banco = df_banco.sort_values(by=col_importe, ascending=True)
                    
                    nombre_pestana = str(banco)[:31].replace('/', '-').replace('\\', '-').replace(':', '').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
                    if not nombre_pestana.strip() or nombre_pestana.lower() == 'nan':
                        nombre_pestana = "Sin_Banco_Asignado"
                        
                    styled_banco = df_banco.style.apply(resaltar_conciliados, axis=1)
                    styled_banco.to_excel(writer, index=False, sheet_name=nombre_pestana)

                # 8.2 Crear pestaña consolidada de Novedades y Alertas
                # Se incluyen todas las condiciones que sean Validar, Sugerencia o Revisar
                df_novedades = df_final[df_final['Estado_Conciliacion'].str.contains('Validar|Sugerencia|Revisar', na=False, case=False)].copy()
                if not df_novedades.empty:
                    df_novedades = df_novedades.sort_values(by=['Estado_Conciliacion', col_importe], ascending=[True, True])
                    styled_novedades = df_novedades.style.apply(resaltar_conciliados, axis=1)
                    styled_novedades.to_excel(writer, index=False, sheet_name='NOVEDADES_Y_ALERTAS')

            # --- 9. INTERFAZ Y DESCARGA ---
            st.success("¡Conciliación Multibanco y Análisis Integral terminados!")
            
            # Métricas
            conciliados_exactos = len(ind_r1) + len(ind_r2)
            novedades_detectadas = len(ind_date) + len(ind_bank) + len(ind_val)
            pendientes = len(df_final) - (conciliados_exactos + novedades_detectadas)

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Bancos Procesados", len(bancos_unicos))
            col2.metric("Conciliados Exactos", conciliados_exactos)
            col3.metric("Novedades (Alertas)", novedades_detectadas)
            col4.metric("Aún Pendientes", pendientes)

            st.download_button(
                label="📥 Descargar Excel con Resultados",
                data=output.getvalue(),
                file_name="Conciliacion_Automatizada_Integral.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error técnico detectado. Detalle técnico: {e}")
