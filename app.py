import streamlit as st
import pandas as pd
import numpy as np
import io
import re

# ============================================================
# CONFIGURACION DE LA PAGINA
# ============================================================
st.set_page_config(page_title="Conciliacion Integral", layout="wide")

hide_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(hide_style, unsafe_allow_html=True)

st.title("🏦 Conciliacion Automatizada 🤖")
st.write("Sube tu archivo consolidado.")

with st.expander("⚙️ Parametros de tolerancia para sugerencias (alertas)"):
    tol_dias = st.slider("Dias maximos de diferencia para alertar 'error de fecha' (mismo periodo)", 1, 9, 9)
    tol_valor_abs = st.number_input("Diferencia absoluta maxima de valor para alertar ($)", min_value=1, value=5000, step=100)
    tol_valor_pct = st.number_input("Diferencia relativa maxima de valor para alertar (%)", min_value=0.01, value=0.5, step=0.01) / 100
    multiplo_redondo = st.selectbox("Multiplo para considerar un valor 'redondo' (alta ambiguedad)", [50000, 100000], index=1)
    tol_dias_cerrados = st.slider("Dias maximos de diferencia de Fecha valor DENTRO DEL MISMO PERIODO CONTABLE", 1, 9, 5)

archivo_subido = st.file_uploader("Selecciona el archivo de Excel o CSV", type=['xlsx', 'csv'])

# =========================================================
# FUNCION 1: DOCUMENTOS REPETIDOS 
# =========================================================
def analizar_documentos_repetidos(
    df, col_doc, col_clave, col_importe, col_clase_doc, usar_ipcb, col_estado="Estado_Conciliacion", col_comentario="Comentario"
):
    df = df.copy()

    if "ID_Temp" not in df.columns: df["ID_Temp"] = df.index

    df["Total_Posiciones_Doc"] = df.groupby(col_doc)[col_doc].transform("count")
    df["Tiene_Posiciones_Repetidas"] = df["Total_Posiciones_Doc"] > 1
    df["N_40_Doc"] = df.groupby(col_doc)[col_clave].transform(lambda x: (x == "40").sum())
    df["N_50_Doc"] = df.groupby(col_doc)[col_clave].transform(lambda x: (x == "50").sum())

    suma_40 = df[df[col_clave] == "40"].groupby(col_doc)[col_importe].sum()
    suma_50 = df[df[col_clave] == "50"].groupby(col_doc)[col_importe].sum()
    df["Suma_40_Doc"] = df[col_doc].map(suma_40).fillna(0)
    df["Suma_50_Doc"] = df[col_doc].map(suma_50).fillna(0)

    condiciones = [
        (df["N_40_Doc"] == 0) | (df["N_50_Doc"] == 0),
        (df["Suma_40_Doc"].abs().round(2) == df["Suma_50_Doc"].abs().round(2)),
    ]
    valores = ["Doc solo tiene un lado (40 o 50)", "Doc cruza exacto"]
    df["Cruce_Doc"] = np.select(condiciones, valores, default="Doc NO cruza - revisar linea")
    df["Detalle_Doc_Repetido"] = ""

    def format_doc_str(r):
        d = str(int(r[col_doc])) if pd.notna(r[col_doc]) else ""
        c = str(r[col_clase_doc]) if usar_ipcb and col_clase_doc in r and pd.notna(r.get(col_clase_doc)) else ""
        k = str(r[col_clave])
        return f"{d} ({c}={k})" if c and c.lower() != 'nan' else f"{d} (Clv {k})"

    for doc_val, grupo in df.groupby(col_doc):
        if len(grupo) <= 1: continue

        g40 = grupo[grupo[col_clave] == "40"].sort_values("ID_Temp")
        g50 = grupo[grupo[col_clave] == "50"].sort_values("ID_Temp")
        n_pares = min(len(g40), len(g50))

        ids_en_par = list(g40["ID_Temp"].iloc[:n_pares]) + list(g50["ID_Temp"].iloc[:n_pares])
        ids_sobrantes = list(g40["ID_Temp"].iloc[n_pares:]) + list(g50["ID_Temp"].iloc[n_pares:])

        # BLINDAJE COLUMNA O: Solo se escriben candidatos si es una relacion perfecta, sin desbalances
        if len(g40) == len(g50) and grupo["Cruce_Doc"].iloc[0] == "Doc cruza exacto":
            df.loc[df["ID_Temp"].isin(grupo["ID_Temp"]), "Detalle_Doc_Repetido"] = "En par dentro del documento"
            txt_candidatos = " | ".join(grupo.apply(format_doc_str, axis=1).tolist())
            df.loc[df["ID_Temp"].isin(grupo["ID_Temp"]), "Candidatos_Conciliacion"] = txt_candidatos
        else:
            df.loc[df["ID_Temp"].isin(ids_en_par), "Detalle_Doc_Repetido"] = "En par (Doc ambiguo o parcial)"
            df.loc[df["ID_Temp"].isin(ids_sobrantes), "Detalle_Doc_Repetido"] = "LÍNEA QUE NO CRUZA en el documento"

    if col_estado in df.columns:
        mask_pendiente = df[col_estado] == "Pendiente"
        mask_neto = (df["Cruce_Doc"] == "Doc cruza exacto") & (df["N_40_Doc"] > 0) & (df["N_50_Doc"] > 0)
        idx_ok = df[mask_pendiente & mask_neto].index
        df.loc[idx_ok, col_estado] = "Conciliado - Documento neto (mismo Nro documento)"
        df.loc[idx_ok, col_comentario] = "Documento " + df.loc[idx_ok, col_doc].astype(int).astype(str) + " con posiciones repetidas: netean exacto."

    return df

# =========================================================
# FUNCION 2: VALORES CERRADOS (BLINDADO CON DIST Y PERIODO)
# =========================================================
def conciliar_valores_cerrados(
    df, col_banco, col_importe, col_clave, col_fecha, col_fecha_contable, col_doc,
    col_clase_doc, usar_ipcb, col_estado="Estado_Conciliacion", col_comentario="Comentario",
    valores_cerrados=None, detectar_multiplos_de=50000, tol_dias_fecha=5, solo_pendientes=True
):
    df = df.copy()

    base = df[df[col_estado] == "Pendiente"].copy() if solo_pendientes and col_estado in df.columns else df.copy()
    if base.empty: return df

    base_cerrados = base[(base["Abs_Importe"] > 0) & (base["Abs_Importe"] % detectar_multiplos_de == 0)].copy()
    base_cerrados = base_cerrados[base_cerrados["Periodo_Contable"] != "SIN_FECHA_CONTABLE"]
    if base_cerrados.empty: return df

    resultados_estado = {}
    resultados_comentario = {}

    def format_doc_str(r):
        d = str(int(r[col_doc])) if pd.notna(r[col_doc]) else ""
        c = str(r[col_clase_doc]) if usar_ipcb and col_clase_doc in r and pd.notna(r.get(col_clase_doc)) else ""
        k = str(r[col_clave])
        return f"{d} ({c}={k})" if c and c.lower() != 'nan' else f"{d} (Clv {k})"

    # Agrupamos tambien por Distribuidora para no cruzar valores de distintas sedes
    for (banco, importe, periodo, dist), grupo in base_cerrados.groupby([col_banco, "Abs_Importe", "Periodo_Contable", "Distribuidora"]):
        lado_40 = grupo[grupo[col_clave] == "40"].sort_values(by=["Fecha_Calc", col_doc]).reset_index(drop=True)
        lado_50 = grupo[grupo[col_clave] == "50"].sort_values(by=["Fecha_Calc", col_doc]).reset_index(drop=True)
        if lado_40.empty or lado_50.empty: continue

        n_pares = min(len(lado_40), len(lado_50))
        es_ambiguo = len(lado_40) != len(lado_50) # BLINDAJE COLUMNA O

        for i in range(n_pares):
            fila_40, fila_50 = lado_40.iloc[i], lado_50.iloc[i]
            id_40, id_50 = fila_40["ID_Temp"], fila_50["ID_Temp"]
            f40, f50 = fila_40["Fecha_Calc"], fila_50["Fecha_Calc"]

            dif_dias = None if (pd.isna(f40) or pd.isna(f50)) else abs((f40 - f50).days)
            doc_40, doc_50 = fila_40[col_doc], fila_50[col_doc]

            if dif_dias is None:
                estado = "Sugerencia - Valor cerrado (fecha valor invalida)"
                detalle_dif = "fecha valor no disponible"
            elif dif_dias == 0:
                estado = "Conciliado - Valor cerrado (misma fecha, mismo periodo)"
                detalle_dif = "misma fecha valor"
            elif dif_dias <= tol_dias_fecha:
                estado = "Conciliado - Valor cerrado (mismo periodo contable)"
                detalle_dif = f"difieren {dif_dias} dia(s), MISMO periodo contable ({periodo})"
            elif dif_dias <= 9:
                estado = "Sugerencia - Valor cerrado (verificar fecha, mismo periodo)"
                detalle_dif = f"difieren {dif_dias} dia(s), MISMO periodo contable ({periodo})"
            else:
                continue # Limite estricto max 9 dias

            monto_txt = f"${importe:,.0f}"
            resultados_estado[id_40] = estado
            resultados_comentario[id_40] = f"Valor cerrado {monto_txt} emparejado FIFO Sede {dist} con Doc {int(doc_50)} ({detalle_dif})."
            resultados_estado[id_50] = estado
            resultados_comentario[id_50] = f"Valor cerrado {monto_txt} emparejado FIFO Sede {dist} con Doc {int(doc_40)} ({detalle_dif})."

            if not es_ambiguo: # Columna O limpia si hay desbalances
                txt_candidatos = format_doc_str(fila_40) + " | " + format_doc_str(fila_50)
                df.loc[[id_40, id_50], "Candidatos_Conciliacion"] = txt_candidatos

        for _, fila in lado_40.iloc[n_pares:].iterrows():
            resultados_estado[fila["ID_Temp"]] = "Pendiente - Valor cerrado sin par (mismo periodo)"
            resultados_comentario[fila["ID_Temp"]] = f"Valor cerrado ${importe:,.0f} sin contraparte en banco {banco}, per {periodo} ({dist})."
        for _, fila in lado_50.iloc[n_pares:].iterrows():
            resultados_estado[fila["ID_Temp"]] = "Pendiente - Valor cerrado sin par (mismo periodo)"
            resultados_comentario[fila["ID_Temp"]] = f"Valor cerrado ${importe:,.0f} sin contraparte en banco {banco}, per {periodo} ({dist})."

    for id_temp, estado in resultados_estado.items(): df.loc[df["ID_Temp"] == id_temp, col_estado] = estado
    for id_temp, comentario in resultados_comentario.items(): df.loc[df["ID_Temp"] == id_temp, col_comentario] = comentario
    return df

if archivo_subido is not None:
    try:
        with st.spinner("Ejecutando motor de reglas, M:N y clasificacion multibanco..."):

            # =========================================================
            # 1. LECTURA Y MAPEO DE COLUMNAS
            # =========================================================
            if archivo_subido.name.lower().endswith('.csv'): df = pd.read_csv(archivo_subido)
            else:
                diccionario_hojas = pd.read_excel(archivo_subido, sheet_name=None)
                df = pd.concat(diccionario_hojas.values(), ignore_index=True)

            df.columns = df.columns.str.strip()

            col_asignacion = 'Asignación' if 'Asignación' in df.columns else 'Asignacion' 
            col_referencia = 'Referencia'
            col_clave = 'Clave contabiliz.' if 'Clave contabiliz.' in df.columns else 'CT'
            col_fecha = 'Fecha valor' if 'Fecha valor' in df.columns else 'Fe-valor'
            col_fecha_contable = 'Fe.contabilización' if 'Fe.contabilización' in df.columns else ('Fecha de documento' if 'Fecha de documento' in df.columns else col_fecha)
            col_importe = 'Importe en moneda local' if 'Importe en moneda local' in df.columns else 'Importe en ML'
            col_banco = 'Clave referencia 3'
            col_doc = 'Nº documento' if 'Nº documento' in df.columns else 'Nº doc.'
            col_texto = 'Texto' if 'Texto' in df.columns else None
            col_clase_doc = 'Clase de documento' if 'Clase de documento' in df.columns else ('Clase doc.' if 'Clase doc.' in df.columns else None)

            columnas_requeridas = [col_referencia, col_clave, col_fecha, col_importe, col_banco, col_doc]
            faltantes = [c for c in columnas_requeridas if c not in df.columns]
            if faltantes: st.error(f"Faltan columnas: {faltantes}"); st.stop()
            if col_asignacion not in df.columns: st.error("No se encontró la columna de Asignación."); st.stop()

            usar_ipcb = col_clase_doc is not None

            # =========================================================
            # 2. AUTOCOMPLETADO DE BANCOS 
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
                asig_val, banco_val = str(row.get(col_asignacion, "")), row.get(col_banco, None)
                if "cuenta de mayor" in asig_val.lower():
                    match_cuenta = re.search(r'(\d{6,})', asig_val)
                    if match_cuenta: current_bank = mapeo_cuentas_banco.get(match_cuenta.group(1), f"CUENTA {match_cuenta.group(1)} (sin mapear)")
                if pd.notnull(banco_val) and str(banco_val).strip().lower() not in ("", "nan"): current_bank = str(banco_val).strip()
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
            
            # REGLA OBLIGATORIA: PERIODO CONTABLE
            df['Fecha_Contable_Calc'] = pd.to_datetime(df[col_fecha_contable], errors='coerce')
            df['Periodo_Contable'] = df['Fecha_Contable_Calc'].dt.to_period("M").astype(str)
            df.loc[df['Fecha_Contable_Calc'].isna(), 'Periodo_Contable'] = "SIN_FECHA_CONTABLE"

            # Blindaje Ref: Ceros vacios o nulos
            df[col_referencia] = df[col_referencia].astype(str).str.strip()
            df.loc[df[col_referencia].str.replace(r'^0+$', '', regex=True) == '', col_referencia] = np.nan
            df.loc[df[col_referencia].isin(['nan', 'None']), col_referencia] = np.nan

            df['Estado_Conciliacion'] = 'Pendiente'
            df['Comentario'] = ''
            df['Candidatos_Conciliacion'] = ''

            def set_estado(indices, estado):
                if indices: df.loc[df['ID_Temp'].isin(indices), 'Estado_Conciliacion'] = estado

            def set_comentarios(dic_comentarios):
                for id_temp, texto in dic_comentarios.items(): df.loc[df['ID_Temp'] == id_temp, 'Comentario'] = texto

            def registrar_candidatos(ids):
                ids = list(set(ids))
                if not ids: return
                sub = df.loc[ids].sort_values(col_clave)
                nombres = []
                for _, r in sub.iterrows():
                    d = str(int(r[col_doc])) if pd.notna(r[col_doc]) else ""
                    c = str(r[col_clase_doc]) if usar_ipcb and pd.notna(r.get(col_clase_doc)) else ""
                    k = str(r[col_clave])
                    if c and c.lower() != 'nan': nombres.append(f"{d} ({c}={k})")
                    else: nombres.append(f"{d} (Clv {k})")
                df.loc[ids, 'Candidatos_Conciliacion'] = " | ".join(nombres)

            # =========================================================
            # 3B. EJECUCIÓN DE DOCUMENTOS REPETIDOS
            # =========================================================
            df = analizar_documentos_repetidos(
                df, col_doc=col_doc, col_clave=col_clave, col_importe=col_importe,
                col_clase_doc=col_clase_doc, usar_ipcb=usar_ipcb,
                col_estado='Estado_Conciliacion', col_comentario='Comentario'
            )
            ind_doc_neto = set(df[df['Estado_Conciliacion'].astype(str).str.contains('Documento neto', na=False)]['ID_Temp'])

            # =========================================================
            # 4. CLASIFICACION DE DISTRIBUIDORAS Y HOMOLOGACION 
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
                asig_val = str(row.get(col_asignacion, "")) if col_asignacion else ""
                t = f"{texto_val} {asig_val} {ref_val}".upper()
                if 'DOSQ' in t or 'D504' in t: return 'Dist Dosquebradas'
                if 'ACOPI' in t or 'D503' in t: return 'Dist Acopi'
                if 'PASTO' in t or 'D505' in t: return 'Dist Pasto'
                if 'BUGA' in t or 'D502' in t: return 'Dist Buga'
                return 'Sin clasificar'

            def obtener_ref_homologada(row):
                texto = f"{row.get(col_referencia,'')} {row.get(col_asignacion,'')}".upper()
                for num in re.findall(r'\b\d{8}\b', texto):
                    if num in mapeo_datafono_ref: return mapeo_datafono_ref[num]
                for num in re.findall(r'\b\d{4}\b', texto):
                    if num in mapeo_datafono_ref.values(): return num
                return None

            df['Distribuidora'] = df.apply(clasificar_distribuidora, axis=1)
            def resumen_docs(sub_df): return ", ".join(str(int(d)) for d in sub_df[col_doc].tolist())
            def es_valor_redondo(v): return (v % multiplo_redondo == 0) and v > 0

            # =========================================================
            # NIVEL 1: CRUCES EXACTOS Y MULTIPLES (BLINDADO CON PERIODO)
            # =========================================================
            df_p1 = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_40, df_50 = df_p1[df_p1[col_clave] == '40'].copy(), df_p1[df_p1[col_clave] == '50'].copy()

            grp_base = [col_banco, 'Abs_Importe', col_fecha, 'Periodo_Contable']
            
            df_40['T'] = df_40.groupby(grp_base + [col_referencia]).cumcount()
            df_50['T'] = df_50.groupby(grp_base + [col_referencia]).cumcount()
            c1 = pd.merge(df_40, df_50, on=grp_base + [col_referencia, 'T'], suffixes=('_40', '_50'))

            df_40['T2'] = df_40.groupby(grp_base + [col_asignacion]).cumcount()
            df_50['T2'] = df_50.groupby(grp_base + [col_referencia]).cumcount()
            c2 = pd.merge(df_40, df_50, left_on=grp_base + [col_asignacion, 'T2'], right_on=grp_base + [col_referencia, 'T2'], suffixes=('_40', '_50'))

            df_40['T3'] = df_40.groupby(grp_base + [col_referencia]).cumcount()
            df_50['T3'] = df_50.groupby(grp_base + [col_asignacion]).cumcount()
            c3 = pd.merge(df_40, df_50, left_on=grp_base + [col_referencia, 'T3'], right_on=grp_base + [col_asignacion, 'T3'], suffixes=('_40', '_50'))

            ind_r1 = set(c1['ID_Temp_40']) | set(c1['ID_Temp_50']) | set(c2['ID_Temp_40']) | set(c2['ID_Temp_50']) | set(c3['ID_Temp_40']) | set(c3['ID_Temp_50'])
            set_estado(ind_r1, 'Conciliado - Cruce exacto')
            com_r1 = {}
            for c in (c1, c2, c3):
                for _, r in c.iterrows():
                    com_r1[r['ID_Temp_40']] = f"Cruce exacto (ref/banco/importe). Doc: {int(r[col_doc + '_50'])}"
                    com_r1[r['ID_Temp_50']] = f"Cruce exacto (ref/banco/importe). Doc: {int(r[col_doc + '_40'])}"
                    registrar_candidatos([r['ID_Temp_40'], r['ID_Temp_50']])
            set_comentarios(com_r1)

            # 1B: Cruce exacto (Referencia Limpia)
            df_p0 = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_p0['A_L'] = df_p0[col_asignacion].astype(str).str.extract(r'(\d+)')[0]
            df_p0['R_L'] = df_p0[col_referencia].astype(str).str.extract(r'(\d+)')[0]

            d40b, d50b = df_p0[df_p0[col_clave] == '40'].drop(columns=['R_L']).copy(), df_p0[df_p0[col_clave] == '50'].drop(columns=['A_L']).copy()
            d40b['Tb'] = d40b.groupby(grp_base + ['A_L']).cumcount()
            d50b['Tb'] = d50b.groupby(grp_base + ['R_L']).cumcount()

            c1b = pd.merge(d40b, d50b, left_on=grp_base + ['A_L', 'Tb'], right_on=grp_base + ['R_L', 'Tb'], suffixes=('_40', '_50'))
            c1b = c1b[c1b['A_L'].notna() & (c1b['A_L'] != '')]

            ind_r1b = set(c1b['ID_Temp_40']) | set(c1b['ID_Temp_50'])
            set_estado(ind_r1b, 'Conciliado - Cruce exacto (Ref limpia)')
            com_r1b = {}
            for _, r in c1b.iterrows():
                com_r1b[r['ID_Temp_40']] = f"Cruce ref limpiada. Doc: {int(r[col_doc + '_50'])}"
                com_r1b[r['ID_Temp_50']] = f"Cruce ref limpiada. Doc: {int(r[col_doc + '_40'])}"
                registrar_candidatos([r['ID_Temp_40'], r['ID_Temp_50']])
            set_comentarios(com_r1b)

            # 1C: Cruce Multiple Datafonos M:N (BLINDADO CON PERIODO Y CANTIDADES)
            ind_1c_ipcb = set(); com_1c_ipcb = {}
            if usar_ipcb:
                df_p_ipcb = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
                df_p_ipcb['Ref_H'] = df_p_ipcb.apply(obtener_ref_homologada, axis=1)
                df_vp = df_p_ipcb[df_p_ipcb['Ref_H'].notna()]

                df_ip = df_vp[(df_vp[col_clase_doc].astype(str).str.upper() == 'IP') & (df_vp[col_clave] == '40')]
                df_cb = df_vp[(df_vp[col_clase_doc].astype(str).str.upper() == 'CB') & (df_vp[col_clave] == '50')]

                if not df_ip.empty and not df_cb.empty:
                    grp_ip = df_ip.groupby([col_banco, col_fecha, 'Periodo_Contable', 'Ref_H'])['Abs_Importe'].sum().reset_index(name='S_IP')
                    grp_cb = df_cb.groupby([col_banco, col_fecha, 'Periodo_Contable', 'Ref_H'])['Abs_Importe'].sum().reset_index(name='S_CB')

                    m_ipcb = pd.merge(grp_cb, grp_ip, on=[col_banco, col_fecha, 'Periodo_Contable', 'Ref_H'])
                    matches_ipcb = m_ipcb[m_ipcb['S_CB'] == m_ipcb['S_IP']] # Garantiza balance exacto M:N

                    for _, m in matches_ipcb.iterrows():
                        b, f, p, rh = m[col_banco], m[col_fecha], m['Periodo_Contable'], m['Ref_H']
                        sub_ip = df_ip[(df_ip[col_banco] == b) & (df_ip[col_fecha] == f) & (df_ip['Periodo_Contable'] == p) & (df_ip['Ref_H'] == rh)]
                        sub_cb = df_cb[(df_cb[col_banco] == b) & (df_cb[col_fecha] == f) & (df_cb['Periodo_Contable'] == p) & (df_cb['Ref_H'] == rh)]

                        ip_ids, cb_ids = sub_ip['ID_Temp'].tolist(), sub_cb['ID_Temp'].tolist()
                        ind_1c_ipcb.update(ip_ids + cb_ids)

                        txt = f"Cruce multiple IP/CB. Ref: {rh}."
                        for cb_id in cb_ids: com_1c_ipcb[cb_id] = f"{txt} Docs IP: {resumen_docs(sub_ip)}"
                        for ip_id in ip_ids: com_1c_ipcb[ip_id] = f"{txt} Docs CB: {resumen_docs(sub_cb)}"
                        registrar_candidatos(ip_ids + cb_ids)

                set_estado(ind_1c_ipcb, 'Conciliado - Cruce multiple')
                set_comentarios(com_1c_ipcb)

            # =========================================================
            # NIVEL 1D: VALORES CERRADOS REPETIDOS
            # =========================================================
            df = conciliar_valores_cerrados(
                df, col_banco=col_banco, col_importe=col_importe, col_clave=col_clave,
                col_fecha=col_fecha, col_fecha_contable=col_fecha_contable, col_doc=col_doc,
                col_clase_doc=col_clase_doc, usar_ipcb=usar_ipcb, col_estado='Estado_Conciliacion',
                col_comentario='Comentario', detectar_multiplos_de=multiplo_redondo,
                tol_dias_fecha=tol_dias_cerrados, solo_pendientes=True
            )
            ind_cerrados = set(df[df['Estado_Conciliacion'].astype(str).str.contains('Valor cerrado', na=False)]['ID_Temp'])

            # =========================================================
            # NIVEL 2: SUGERENCIAS MULTIPLES Y SECTORIZACION
            # =========================================================
            df_p1d = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            if usar_ipcb: df_p1d = df_p1d[df_p1d[col_clase_doc].astype(str).str.upper() != 'IP']

            df_sect = df_p1d[df_p1d['Distribuidora'] != 'Sin clasificar']
            d40c, d50c = df_sect[df_sect[col_clave] == '40'].copy(), df_sect[df_sect[col_clave] == '50'].copy()

            ind_r1d = set(); ind_r1d_f = set(); ind_r1d_a = set()
            com_r1d = {}; com_r1d_f = {}; com_r1d_a = {}

            if not d40c.empty and not d50c.empty:
                for grp, sub40 in d40c.groupby([col_banco, 'Abs_Importe', col_fecha, 'Periodo_Contable', 'Distribuidora']):
                    b, imp, f, per, dist = grp
                    sub50 = d50c[(d50c[col_banco] == b) & (d50c['Abs_Importe'] == imp) & (d50c[col_fecha] == f) & (d50c['Periodo_Contable'] == per) & (d50c['Distribuidora'] == dist)]
                    if sub50.empty: continue

                    s40_ord, s50_ord = sub40.sort_values(col_doc), sub50.sort_values(col_doc)

                    if len(s40_ord) == 1 and len(s50_ord) == 1:
                        r40, r50 = s40_ord.iloc[0], s50_ord.iloc[0]
                        ind_r1d.update([r40['ID_Temp'], r50['ID_Temp']])
                        com_r1d[r40['ID_Temp']] = f"Candidato unico sede ({dist}). Doc: {int(r50[col_doc])}"
                        com_r1d[r50['ID_Temp']] = f"Candidato unico sede ({dist}). Doc: {int(r40[col_doc])}"
                        registrar_candidatos([r40['ID_Temp'], r50['ID_Temp']])
                    elif len(s40_ord) == len(s50_ord):
                        for (_, r40), (_, r50) in zip(s40_ord.iterrows(), s50_ord.iterrows()):
                            ind_r1d_f.update([r40['ID_Temp'], r50['ID_Temp']])
                            txt = f"Sede '{dist}' emparejado FIFO."
                            com_r1d_f[r40['ID_Temp']] = f"{txt} Doc: {int(r50[col_doc])}"
                            com_r1d_f[r50['ID_Temp']] = f"{txt} Doc: {int(r40[col_doc])}"
                            registrar_candidatos([r40['ID_Temp'], r50['ID_Temp']])
                    else:
                        for _, r in s40_ord.iterrows():
                            ind_r1d_a.add(r['ID_Temp']); com_r1d_a[r['ID_Temp']] = f"Sede '{dist}' desbalance ({len(s40_ord)} vs {len(s50_ord)}). Creditos: {resumen_docs(s50_ord)}"
                        for _, r in s50_ord.iterrows():
                            ind_r1d_a.add(r['ID_Temp']); com_r1d_a[r['ID_Temp']] = f"Sede '{dist}' desbalance ({len(s50_ord)} vs {len(s40_ord)}). Debitos: {resumen_docs(s40_ord)}"
                        # BLINDAJE COLUMNA O: NO se llama a registrar_candidatos si hay desbalance.

            set_estado(ind_r1d, 'Conciliado - Cruce Distribuidora')
            set_comentarios(com_r1d)
            set_estado(ind_r1d_f, 'Sugerencia fuerte: Sectorizacion (FIFO)')
            set_comentarios(com_r1d_f)
            set_estado(ind_r1d_a, 'Sugerencia: Sugerencia por Distribuidora Multiples')
            set_comentarios(com_r1d_a)

            # Cruce Unico sin referencia (BLINDADO CON DISTRIBUIDORA Y PERIODO)
            df_p = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            if usar_ipcb: df_p = df_p[df_p[col_clase_doc].astype(str).str.upper() != 'IP']
            grp_c = [col_banco, 'Abs_Importe', col_fecha, 'Distribuidora', 'Periodo_Contable']

            df_p40, df_p50 = df_p[df_p[col_clave] == '40'].copy(), df_p[df_p[col_clave] == '50'].copy()
            df_p40['n40'] = df_p40.groupby(grp_c)['ID_Temp'].transform('count')
            df_p50['n50'] = df_p50.groupby(grp_c)['ID_Temp'].transform('count')

            u40, u50 = df_p40[df_p40['n40'] == 1], df_p50[df_p50['n50'] == 1]
            c_un = pd.merge(u40, u50, on=grp_c, suffixes=('_40', '_50'))

            ind_r2 = set(c_un['ID_Temp_40']) | set(c_un['ID_Temp_50'])
            set_estado(ind_r2, 'Conciliado - Cruce unico')
            com_r2 = {}
            for _, r in c_un.iterrows():
                com_r2[r['ID_Temp_40']] = f"Unico sin referencia. Doc: {int(r[col_doc + '_50'])}"
                com_r2[r['ID_Temp_50']] = f"Unico sin referencia. Doc: {int(r[col_doc + '_40'])}"
                registrar_candidatos([r['ID_Temp_40'], r['ID_Temp_50']])
            set_comentarios(com_r2)

            # Desempate Grupo Cerrado (FIFO y Ambiguos)
            rem40, rem50 = df_p40[~df_p40['ID_Temp'].isin(ind_r2)], df_p50[~df_p50['ID_Temp'].isin(ind_r2)]
            ind_r2d = set(); com_r2d = {}
            ind_amb = set(); com_amb = {}

            for grp, sub40 in rem40.groupby(grp_c):
                b, imp, f, dist, per = grp
                sub50 = rem50[(rem50[col_banco] == b) & (rem50['Abs_Importe'] == imp) & (rem50[col_fecha] == f) & (rem50['Distribuidora'] == dist) & (rem50['Periodo_Contable'] == per)]
                if sub50.empty: continue

                if es_valor_redondo(imp):
                    s40_ord, s50_ord = sub40.sort_values(col_doc), sub50.sort_values(col_doc)
                    if len(s40_ord) == len(s50_ord):
                        for (_, r40), (_, r50) in zip(s40_ord.iterrows(), s50_ord.iterrows()):
                            ind_r2d.update([r40['ID_Temp'], r50['ID_Temp']])
                            com_r2d[r40['ID_Temp']] = f"Valor redondo (${imp:,.0f}) FIFO Sede {dist}. Doc: {int(r50[col_doc])}"
                            com_r2d[r50['ID_Temp']] = f"Valor redondo (${imp:,.0f}) FIFO Sede {dist}. Doc: {int(r40[col_doc])}"
                            registrar_candidatos([r40['ID_Temp'], r50['ID_Temp']])
                    else:
                        for _, r in s40_ord.iterrows():
                            ind_amb.add(r['ID_Temp']); com_amb[r['ID_Temp']] = f"Confuso ({len(s40_ord)} vs {len(s50_ord)}). Creditos: {resumen_docs(sub50)}"
                        for _, r in s50_ord.iterrows():
                            ind_amb.add(r['ID_Temp']); com_amb[r['ID_Temp']] = f"Confuso ({len(s50_ord)} vs {len(s40_ord)}). Debitos: {resumen_docs(sub40)}"
                        # BLINDAJE COLUMNA O: NO se registra. Columna limpia.
                else:
                    for _, r in sub40.iterrows():
                        ind_amb.add(r['ID_Temp']); com_amb[r['ID_Temp']] = f"{len(sub50)} posibles cruces. Docs: {resumen_docs(sub50)}"
                    for _, r in sub50.iterrows():
                        ind_amb.add(r['ID_Temp']); com_amb[r['ID_Temp']] = f"{len(sub40)} posibles cruces. Docs: {resumen_docs(sub40)}"
                    # BLINDAJE COLUMNA O: NO se registra.

            set_estado(ind_r2d, 'Sugerencia fuerte')
            set_comentarios(com_r2d)
            set_estado(ind_amb, 'Sugerencia: Solicitar soporte')
            set_comentarios(com_amb)

            # =========================================================
            # NIVEL 3: ALERTAS DE FECHA Y VALOR 
            # =========================================================
            df_pend = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            df_pend['Regex'] = df_pend[col_referencia].astype(str).str.extract(r'(\d+)')[0]
            df_v = df_pend[df_pend['Regex'].notna() & (df_pend['Regex'].str.replace('0', '') != '')].copy()

            df_4n, df_5n = df_v[df_v[col_clave] == '40'].copy(), df_v[df_v[col_clave] == '50'].copy()

            # 3A: Alertas de Fecha (BLINDADO 1 a 1 ESTRICTO)
            sA = pd.merge(df_4n, df_5n, on=[col_banco, 'Abs_Importe', 'Regex', 'Periodo_Contable'], suffixes=('_40', '_50'))
            sA['Dif'] = (sA['Fecha_Calc_40'] - sA['Fecha_Calc_50']).dt.days.abs()
            sA = sA[(sA['Dif'] > 0) & (sA['Dif'] <= tol_dias)]
            sA['n4'] = sA.groupby('ID_Temp_40')['ID_Temp_50'].transform('count')
            sA['n5'] = sA.groupby('ID_Temp_50')['ID_Temp_40'].transform('count')
            sA = sA[(sA['n4'] == 1) & (sA['n5'] == 1)] # Regla Cero Ambigüedades

            ind_A = set(); com_A = {}
            for _, r in sA.iterrows():
                dif = int(r['Dif'])
                ids = [r['ID_Temp_40'], r['ID_Temp_50']]
                ind_A.update(ids)
                df.loc[df['ID_Temp'].isin(ids), 'Estado_Conciliacion'] = 'Alerta - Diferencia Fecha (Mismo Periodo)'
                com_A[r['ID_Temp_40']] = f"ALERTA: Difiere {dif} dia(s) (mismo periodo). Doc: {int(r[col_doc+'_50'])}"
                com_A[r['ID_Temp_50']] = f"ALERTA: Difiere {dif} dia(s) (mismo periodo). Doc: {int(r[col_doc+'_40'])}"
                registrar_candidatos([r['ID_Temp_40'], r['ID_Temp_50']])
            set_comentarios(com_A)

            df_4n, df_5n = df_4n[~df_4n['ID_Temp'].isin(ind_A)], df_5n[~df_5n['ID_Temp'].isin(ind_A)]

            # 3B: Reclasificacion de Banco (BLINDADO 1 a 1 ESTRICTO)
            sB = pd.merge(df_4n, df_5n, on=['Abs_Importe', col_fecha, 'Regex', 'Periodo_Contable'], suffixes=('_40', '_50'))
            sB = sB[sB[f'{col_banco}_40'] != sB[f'{col_banco}_50']]
            sB['n4'] = sB.groupby('ID_Temp_40')['ID_Temp_50'].transform('count')
            sB['n5'] = sB.groupby('ID_Temp_50')['ID_Temp_40'].transform('count')
            sB = sB[(sB['n4'] == 1) & (sB['n5'] == 1)]

            ind_B = set(sB['ID_Temp_40']) | set(sB['ID_Temp_50'])
            set_estado(ind_B, 'Reclasificacion de banco')
            com_B = {}
            for _, r in sB.iterrows():
                com_B[r['ID_Temp_40']] = f"Registrado en banco '{r[col_banco+'_50']}'. Doc: {int(r[col_doc+'_50'])}"
                com_B[r['ID_Temp_50']] = f"Registrado en banco '{r[col_banco+'_40']}'. Doc: {int(r[col_doc+'_40'])}"
                registrar_candidatos([r['ID_Temp_40'], r['ID_Temp_50']])
            set_comentarios(com_B)

            df_4n, df_5n = df_4n[~df_4n['ID_Temp'].isin(ind_B)], df_5n[~df_5n['ID_Temp'].isin(ind_B)]

            # 3C: Diferencia de Valor con Referencia (BLINDADO 1 a 1 ESTRICTO)
            sC = pd.merge(df_4n, df_5n, on=[col_banco, col_fecha, 'Regex', 'Periodo_Contable'], suffixes=('_40', '_50'))
            sC['DifV'] = (sC['Abs_Importe_40'] - sC['Abs_Importe_50']).abs()
            max_imp = sC[['Abs_Importe_40', 'Abs_Importe_50']].max(axis=1)
            sC['Pct'] = np.where(max_imp == 0, 0, sC['DifV'] / max_imp)
            sC = sC[(sC['DifV'] > 0) & ((sC['DifV'] <= tol_valor_abs) | (sC['Pct'] <= tol_valor_pct))]
            sC['n4'] = sC.groupby('ID_Temp_40')['ID_Temp_50'].transform('count')
            sC['n5'] = sC.groupby('ID_Temp_50')['ID_Temp_40'].transform('count')
            sC = sC[(sC['n4'] == 1) & (sC['n5'] == 1)]

            ind_C = set(sC['ID_Temp_40']) | set(sC['ID_Temp_50'])
            set_estado(ind_C, 'Diferencia en valor')
            com_C = {}
            for _, r in sC.iterrows():
                com_C[r['ID_Temp_40']] = f"Diferencia de ${r['DifV']:,.0f} ({r['Pct']*100:.2f}%). Doc: {int(r[col_doc+'_50'])}"
                com_C[r['ID_Temp_50']] = f"Diferencia de ${r['DifV']:,.0f} ({r['Pct']*100:.2f}%). Doc: {int(r[col_doc+'_40'])}"
                registrar_candidatos([r['ID_Temp_40'], r['ID_Temp_50']])
            set_comentarios(com_C)

            # 3D: Diferencia de Valor SIN Referencia - NEQUI (BLINDADO CON DISTRIBUIDORA Y PERIODO)
            df_p11d = df[df['Estado_Conciliacion'] == 'Pendiente'].copy()
            if usar_ipcb: df_p11d = df_p11d[df_p11d[col_clase_doc].astype(str).str.upper() != 'IP']

            d4d, d5d = df_p11d[df_p11d[col_clave] == '40'], df_p11d[df_p11d[col_clave] == '50']
            ind_D = set()
            if not d4d.empty and not d5d.empty:
                sD = pd.merge(d4d, d5d, on=[col_banco, col_fecha, 'Distribuidora', 'Periodo_Contable'], suffixes=('_40', '_50'))
                sD['DifV'] = (sD['Abs_Importe_40'] - sD['Abs_Importe_50']).abs()
                max_impD = sD[['Abs_Importe_40', 'Abs_Importe_50']].max(axis=1)
                sD['Pct'] = np.where(max_impD == 0, 0, sD['DifV'] / max_impD)
                sDt = sD[(sD['DifV'] > 0) & ((sD['DifV'] <= tol_valor_abs) | (sD['Pct'] <= tol_valor_pct))].copy()
                if not sDt.empty:
                    sDt['n4'] = sDt.groupby('ID_Temp_40')['ID_Temp_50'].transform('count')
                    sDt['n5'] = sDt.groupby('ID_Temp_50')['ID_Temp_40'].transform('count')
                    sDu = sDt[(sDt['n4'] == 1) & (sDt['n5'] == 1)]
                    ind_D = set(sDu['ID_Temp_40']) | set(sDu['ID_Temp_50'])
                    set_estado(ind_D, 'Diferencia en valor (NEQUI)')
                    com_D = {}
                    for _, r in sDu.iterrows():
                        com_D[r['ID_Temp_40']] = f"Candidato unico ({r['Distribuidora']}) con dif. de ${r['DifV']:,.0f}. Doc: {int(r[col_doc+'_50'])}"
                        com_D[r['ID_Temp_50']] = f"Candidato unico ({r['Distribuidora']}) con dif. de ${r['DifV']:,.0f}. Doc: {int(r[col_doc+'_40'])}"
                        registrar_candidatos([r['ID_Temp_40'], r['ID_Temp_50']])
                    set_comentarios(com_D)

            # =========================================================
            # PENDIENTES FINALES
            # =========================================================
            sin_p = df['Estado_Conciliacion'] == 'Pendiente'
            if usar_ipcb:
                es_ip = df[col_clase_doc].astype(str).str.upper() == 'IP'
                df.loc[sin_p & es_ip & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia - PDV'
                df.loc[sin_p & ~es_ip & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia segura encontrada - requiere revision manual'
            else:
                df.loc[sin_p & (df['Comentario'] == ''), 'Comentario'] = 'Sin coincidencia segura encontrada - requiere revision manual'

            # =========================================================
            # LIMPIEZA FINAL Y FORMATO (ESTRUCTURA ORIGINAL INTACTA)
            # =========================================================
            cuadre_ok = filas_antes == (len(df) + len(filas_descartadas))
            
            columnas_a_borrar = [
                'ID_Temp', 'Abs_Importe', 'Fecha_Calc', 'Periodo_Contable',
                'Total_Posiciones_Doc', 'Tiene_Posiciones_Repetidas', 
                'N_40_Doc', 'N_50_Doc', 'Suma_40_Doc', 'Suma_50_Doc', 'Cruce_Doc', 'Fecha_Contable_Calc'
            ]
            df_final = df.drop(columns=columnas_a_borrar, errors='ignore')
            
            for col_f in [c for c in df_final.columns if 'fe.' in c.lower() or 'fecha' in c.lower() or 'fe-' in c.lower()]:
                df_final[col_f] = pd.to_datetime(df_final[col_f], errors='coerce').dt.strftime('%d/%m/%Y')

            def resaltar_conciliados(row):
                est = str(row['Estado_Conciliacion']).strip().lower()

                if est == 'pendiente' or est == '' or est == 'nan':
                    return [''] * len(row)

                if 'valor cerrado' in est:
                    if 'sugerencia' in est or 'sin par' in est:
                        return ['background-color: #FFE699; color: black'] * len(row)
                    return ['background-color: #A9D18E; color: black'] * len(row)

                if 'documento neto' in est:
                    return ['background-color: #9CC2E5; color: black'] * len(row)

                if ('cruce exacto' in est or 'cruce multiple' in est or 'cruce unico' in est
                    or 'cruce distribuidora' in est):
                    return ['background-color: #C5D9F1; color: black'] * len(row)

                if ('fifo' in est or 'multiples' in est
                    or 'sectorizacion' in est or 'solicitar soporte' in est):
                    return ['background-color: #FFF2CC; color: black'] * len(row)

                if 'fecha' in est or 'periodo' in est or 'alerta' in est:
                    return ['background-color: #FDEBD0; color: black'] * len(row)

                if 'reclasificacion' in est or 'otro banco' in est:
                    return ['background-color: #D7BDE2; color: black'] * len(row)

                if 'valor' in est:
                    return ['background-color: #F5B7B1; color: black'] * len(row)

                return [''] * len(row)

            # =========================================================
            # EXPORTACION (PESTAÑAS SEPARADAS ORIGINALES)
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
                banco_str = str(banco_str).strip()
                if banco_str in nombres_ordenados: return nombres_ordenados.index(banco_str)
                for i, acc in enumerate(orden_cuentas):
                    if acc in banco_str: return i
                return 999

            b_unicos = sorted(b_unicos, key=get_bank_order)

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_nov = df_final[~df_final['Estado_Conciliacion'].str.contains('Conciliado|exacto|unico|multiple|Sectorizacion|Documento neto', case=False, na=False)].copy()
                df_nov = df_nov[df_nov[col_clave] == '40']

                if not df_nov.empty:
                    df_nov = df_nov.sort_values(by=['Estado_Conciliacion', col_importe])
                    df_nov.style.apply(resaltar_conciliados, axis=1).to_excel(writer, index=False, sheet_name='NOVEDADES_Y_PENDIENTES_40')
                else:
                    pd.DataFrame(columns=df_final.columns).to_excel(writer, index=False, sheet_name='NOVEDADES_Y_PENDIENTES_40')

                for banco in b_unicos:
                    df_b = df_final[df_final[col_banco] == banco].copy().sort_values(by=col_importe, ascending=True)
                    n_pestana = re.sub(r'[\\/*?:\[\]]', '-', str(banco)[:31])
                    if not n_pestana.strip() or n_pestana.lower() == 'nan':
                        n_pestana = "Sin_Banco"
                    df_b.style.apply(resaltar_conciliados, axis=1).to_excel(writer, index=False, sheet_name=n_pestana)

                if not filas_descartadas.empty:
                    filas_descartadas.to_excel(writer, index=False, sheet_name='DESCARTADAS_SIN_DOC_O_CT')

            # =========================================================
            # INTERFAZ
            # =========================================================
            st.success("¡Conciliacion Integral y Segura terminada! Pestañas ordenadas secuencialmente.")
            if not cuadre_ok:
                st.warning("⚠️ Revisa la pestaña DESCARTADAS, el total de filas no coincide.")

            c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
            c1.metric("Seguras (Azul)", len(ind_r1 | ind_r1b | ind_1c_ipcb | ind_r2 | ind_r1d))
            c2.metric("Documento neto (Nro doc)", len(ind_doc_neto))
            c3.metric("Valores cerrados", len(ind_cerrados))
            c4.metric("Multiples/FIFO (Amarillo)", len(ind_r1d_f | ind_r1d_a | ind_r2d | ind_amb))
            c5.metric("Reclasificar (Lila)", len(ind_B))
            c6.metric("Diferencias Fe/Val (Rojo)", len(ind_A | ind_C | ind_D))
            c7.metric("Pendientes (Sin Color)", len(df_final[df_final['Estado_Conciliacion'] == 'Pendiente']))

            if filas_excluidas > 0:
                st.warning(f"⚠️ Se excluyeron {filas_excluidas} filas vacias/totales.")

            st.download_button(label="📥 Descargar Excel Original Blindado", data=output.getvalue(), file_name="Conciliacion_V1_Segura.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error tecnico detectado: {e}")
