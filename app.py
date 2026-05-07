import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Conciliador GNB Multibanco", layout="wide")

hide_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
"""
st.markdown(hide_style, unsafe_allow_html=True)

st.title("🏦 Conciliador GNB - Multibanco")
st.write("Sube tu archivo consolidado. El sistema conciliará banco por banco y te devolverá un Excel con pestañas separadas.")

archivo_subido = st.file_uploader("Selecciona el archivo consolidado de Excel", type=['xlsx'])

if archivo_subido is not None:
    try:
        with st.spinner("Procesando todos los bancos de forma segura..."):
            df = pd.read_excel(archivo_subido, sheet_name="Sheet1")
            
            # --- MAPEO DE COLUMNAS ---
            col_asignacion = df.columns[0]   # Col A
            col_fecha = df.columns[5]        # Col F
            col_clave = df.columns[6]        # Col G
            col_referencia = df.columns[7]   # Col H
            col_importe = df.columns[8]      # Col I
            col_banco = df.columns[11]       # Col L (Clave referencia 3)

            # --- LIMPIEZA ---
            df['ID_Temp'] = df.index
            df[col_clave] = df[col_clave].astype(str).str.strip().str.replace('.0', '', regex=False)
            df[col_banco] = df[col_banco].astype(str).str.strip() # Limpiar nombre del banco
            df[col_importe] = pd.to_numeric(df[col_importe], errors='coerce').fillna(0).round(2)
            df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce').dt.date
            df['Estado_Conciliacion'] = 'Pendiente'

            # --- LÓGICA DE CASCADA (Ahora incluye el Banco en el cruce) ---
            
            # Regla 1 (Asignación/Referencia + Banco)
            df_40 = df[df[col_clave] == '40']
            df_50 = df[df[col_clave] == '50']
            
            c1 = pd.merge(df_40, df_50, 
                          left_on=[col_banco, col_importe, col_fecha, col_asignacion], 
                          right_on=[col_banco, col_importe, col_fecha, col_referencia], 
                          suffixes=('_40', '_50'))
            
            c2 = pd.merge(df_40, df_50, 
                          left_on=[col_banco, col_importe, col_fecha, col_referencia], 
                          right_on=[col_banco, col_importe, col_fecha, col_asignacion], 
                          suffixes=('_40', '_50'))
            
            ind_r1 = set(c1['ID_Temp_40']).union(set(c1['ID_Temp_50'])).union(set(c2['ID_Temp_40'])).union(set(c2['ID_Temp_50']))
            df.loc[df['ID_Temp'].isin(ind_r1), 'Estado_Conciliacion'] = 'Conciliado (Regla 1)'

            # Regla 2 (Nequi/Flexible + Banco)
            df_p = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            # Ahora el "Turno" se calcula por Banco también, para no mezclar saldos
            df_p['Turno'] = df_p.groupby([col_banco, col_importe, col_fecha, col_clave]).cumcount()
            
            c_n = pd.merge(df_p[df_p[col_clave]=='40'], df_p[df_p[col_clave]=='50'], 
                           on=[col_banco, col_importe, col_fecha, 'Turno'], 
                           suffixes=('_4', '_5'))
            
            ind_r2 = set(c_n['ID_Temp_4']).union(set(c_n['ID_Temp_5']))
            df.loc[df['ID_Temp'].isin(ind_r2), 'Estado_Conciliacion'] = 'Conciliado (Regla 2)'

            df_final = df.drop(columns=['ID_Temp'])

            # --- FUNCION DE COLOR ---
            def resaltar_conciliados(row):
                color = 'background-color: #D6EAF8' if 'Conciliado' in str(row['Estado_Conciliacion']) else ''
                return [color] * len(row)

            # --- EXPORTACIÓN DIVIDIDA POR BANCO ---
            output = io.BytesIO()
            bancos_unicos = df_final[col_banco].unique()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for banco in bancos_unicos:
                    # Filtrar los datos solo para este banco
                    df_banco = df_final[df_final[col_banco] == banco]
                    
                    # Limpiar el nombre del banco para que sea válido como nombre de pestaña en Excel
                    nombre_pestana = str(banco)[:31].replace('/', '-').replace('\\', '-').replace(':', '').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
                    if not nombre_pestana.strip():
                        nombre_pestana = "Sin_Banco"
                        
                    # Aplicar color y guardar en su pestaña respectiva
                    styled_banco = df_banco.style.apply(resaltar_conciliados, axis=1)
                    styled_banco.to_excel(writer, index=False, sheet_name=nombre_pestana)

            # --- INTERFAZ ---
            st.success("¡Conciliación Multibanco terminada!")
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Bancos Procesados", len(bancos_unicos))
            col2.metric("Registros Conciliados", len(ind_r1) + len(ind_r2))
            col3.metric("Aún Pendientes", len(df_final) - (len(ind_r1) + len(ind_r2)))

            st.download_button(
                label="📥 Descargar Excel por Pestañas",
                data=output.getvalue(),
                file_name="Conciliacion_Multibanco_Resaltado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error en el proceso: Revisa que la columna L sea la del banco. Detalle técnico: {e}")
