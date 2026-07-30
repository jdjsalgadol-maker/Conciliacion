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
st.write("Sube tu archivo consolidado. El sistema leerá todas las pestañas automáticamente, completará la marcación de bancos y aplicará reglas de emparejamiento (FIFO).")

archivo_subido = st.file_uploader("Selecciona el archivo de Excel o CSV", type=['xlsx', 'csv'])

if archivo_subido is not None:
    try:
        with st.spinner("Leyendo archivo y unificando datos..."):
            
            # --- 1. LECTURA DINÁMICA A PRUEBA DE PESTAÑAS ---
            if archivo_subido.name.lower().endswith('.csv'):
                df = pd.read_csv(archivo_subido)
            else:
                # sheet_name=None lee TODAS las pestañas como un diccionario.
                # Luego pd.concat las une todas en un solo DataFrame sin importar cómo se llamen.
                diccionario_hojas = pd.read_excel(archivo_subido, sheet_name=None)
                df = pd.concat(diccionario_hojas.values(), ignore_index=True)
            
            # Limpiar nombres de columnas para evitar fallos por espacios invisibles
            df.columns = df.columns.str.strip()
            
            # --- 2. MAPEO SEGURO Y DINÁMICO DE COLUMNAS ---
            col_asignacion = 'Asignación' if 'Asignación' in df.columns else 'Asignaión' 
            col_referencia = 'Referencia'
            col_clave = 'Clave contabiliz.' if 'Clave contabiliz.' in df.columns else 'CT'
            col_fecha = 'Fecha valor' if 'Fecha valor' in df.columns else 'Fe-valor'
            col_importe = 'Importe en moneda local' if 'Importe en moneda local' in df.columns else 'Importe en ML'
            col_banco = 'Clave referencia 3'
            col_doc = 'Nº documento' if 'Nº documento' in df.columns else 'Nº doc.'

            # Si después de leer las hojas no existe la columna documento, detenemos para no arrojar errores raros
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

            # Eliminar las filas de subtotales (Cuenta de mayor)
            if col_asignacion in df.columns:
                df = df[~df[col_asignacion].astype(str).str.contains("Cuenta de mayor", case=False, na=False)].copy()

            # --- 4. LIMPIEZA Y ORDENAMIENTO (FIFO) ---
            df[col_doc] = pd.to_numeric(df[col_doc], errors='coerce')
            
            # Limpiar filas basura que hayan quedado al unir pestañas (ej. filas vacías)
            df = df.dropna(subset=[col_doc, col_clave]).reset_index(drop=True)
            
            df = df.sort_values(by=[col_doc], ascending=True).reset_index(drop=True)
            
            df['ID_Temp'] = df.index
            df[col_clave] = df[col_clave].astype(str).str.strip().str.replace('.0', '', regex=False)
            df[col_banco] = df[col_banco].astype(str).str.strip()
            
            df[col_importe] = pd.to_numeric(df[col_importe], errors='coerce').fillna(0).round(2)
            df['Abs_Importe'] = df[col_importe].abs() 
            
            df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce').dt.date
            df['Estado_Conciliacion'] = 'Pendiente'

            # --- 5. LÓGICA DE CASCADA VECTORIZADA ---
            df_40 = df[df[col_clave] == '40']
            df_50 = df[df[col_clave] == '50']
            
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
            
            ind_r1 = set(c1['ID_Temp_40']).union(set(c1['ID_Temp_50'])) \
                     .union(set(c2['ID_Temp_40'])).union(set(c2['ID_Temp_50'])) \
                     .union(set(c3['ID_Temp_40'])).union(set(c3['ID_Temp_50']))
                     
            df.loc[df['ID_Temp'].isin(ind_r1), 'Estado_Conciliacion'] = 'Conciliado Parte 1'

            df_p = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_p['Turno'] = df_p.groupby([col_banco, 'Abs_Importe', col_fecha, col_clave]).cumcount()
            
            c_n = pd.merge(df_p[df_p[col_clave]=='40'], df_p[df_p[col_clave]=='50'], 
                           on=[col_banco, 'Abs_Importe', col_fecha, 'Turno'], 
                           suffixes=('_4', '_5'))
            
            ind_r2 = set(c_n['ID_Temp_4']).union(set(c_n['ID_Temp_5']))
            df.loc[df['ID_Temp'].isin(ind_r2), 'Estado_Conciliacion'] = 'Conciliado Parte 2'

            # --- 6. LIMPIEZA FINAL ---
            df_final = df.drop(columns=['ID_Temp', 'Abs_Importe'])

            # --- FUNCION DE COLOR ---
            def resaltar_conciliados(row):
                if 'Conciliado Parte 1' in str(row['Estado_Conciliacion']):
                    return ['background-color: #D4EFDF'] * len(row) 
                elif 'Conciliado Parte 2' in str(row['Estado_Conciliacion']):
                    return ['background-color: #D6EAF8'] * len(row) 
                return [''] * len(row)

            # --- 7. EXPORTACIÓN DIVIDIDA POR BANCO ---
            output = io.BytesIO()
            bancos_unicos = df_final[col_banco].unique()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for banco in bancos_unicos:
                    df_banco = df_final[df_final[col_banco] == banco]
                    
                    nombre_pestana = str(banco)[:31].replace('/', '-').replace('\\', '-').replace(':', '').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
                    if not nombre_pestana.strip() or nombre_pestana.lower() == 'nan':
                        nombre_pestana = "Sin_Banco_Asignado"
                        
                    styled_banco = df_banco.style.apply(resaltar_conciliados, axis=1)
                    styled_banco.to_excel(writer, index=False, sheet_name=nombre_pestana)

            # --- 8. INTERFAZ Y DESCARGA ---
            st.success("¡Conciliación Multibanco terminada exitosamente!")
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Bancos Procesados", len(bancos_unicos))
            col2.metric("Registros Conciliados", len(ind_r1) + len(ind_r2))
            col3.metric("Aún Pendientes", len(df_final) - (len(ind_r1) + len(ind_r2)))

            st.download_button(
                label="📥 Descargar Excel con Resultados",
                data=output.getvalue(),
                file_name="Conciliacion_Automatizada_Resultados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error técnico detectado. Detalle técnico: {e}")
