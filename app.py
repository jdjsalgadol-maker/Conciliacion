# app_conciliacion_definitiva.py
#
# Version definitiva: combina el gate de seguridad universal (candidato_seguro)
# con un emparejador generico (emparejar_por_grupo) que elimina la duplicacion
# de codigo entre niveles. Un solo punto de entrada escribe en la columna O
# (Candidatos_Conciliacion): escribir_candidato(). Ninguna regla puede saltarlo.
#
# Reglas obligatorias para CUALQUIER relacion en la columna O:
#   1. Mismo Periodo_Contable (no negociable)
#   2. Fecha valida en ambos lados
#   3. Diferencia de fecha <= dias_alerta_texto (0 = seguro, 1-9 = alerta de texto)
#   4. Mismo banco (Clave referencia 3)
#   5. Mismo importe o dentro de tolerancia declarada
#   6. Sectorizacion por Distribuidora / punto de venta (si aplica en ambos lados)
#   7. Referencia limpia, Asignacion limpia o Referencia homologada coincidente
#      (o ausente en ambos lados, caso "cruce unico sin referencia")
#   8. NEQUI: se acepta como candidato pero siempre marcado para verificacion manual
#   9. Ninguna fila se reutiliza en dos relaciones distintas
#  10. Si el candidato no pasa alguna regla, la columna O queda vacia y el motivo
#      queda explicado en Comentario
#
# Ejecutar con: streamlit run app_conciliacion_definitiva.py

import io
import re

import numpy as np
import pandas as pd
import streamlit as st


st.set_page_config(page_title="Conciliacion Integral Definitiva", layout="wide")

st.title("Conciliacion Automatizada - Version Definitiva")
st.caption(
    "Toda relacion en Candidatos_Conciliacion pasa por un unico gate de seguridad: "
    "periodo, fecha, banco, importe, sectorizacion, referencia y NEQUI."
)

with st.expander("Parametros", expanded=True):
    dias_alerta_texto = st.number_input(
        "Maximo de dias para alerta textual dentro del mismo periodo",
        min_value=1, max_value=9, value=9, step=1
    )
    multiplo_redondo = st.selectbox(
        "Multiplo para considerar un valor cerrado/redondo",
        [50000, 100000], index=1
    )
    tolerancia_valor_abs = st.number_input(
        "Tolerancia absoluta para diferencias de valor ($)",
        min_value=1, value=5000, step=100
    )
    tolerancia_valor_pct = st.number_input(
        "Tolerancia porcentual para diferencias de valor (%)",
        min_value=0.01, value=0.50, step=0.01
    ) / 100

archivo = st.file_uploader("Sube el archivo Excel o CSV", type=["xlsx", "csv"])


# =========================================================
# MAPEOS DE NEGOCIO
# =========================================================

MAPEO_CUENTAS = {
    "1110056001": "CUENTA 1110056001",
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

MAPEO_DISTRIBUIDORA = {
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
    "17608605": "Dist Pasto", "17968405": "VENTA EN LINEA",
}

MAPEO_DATAFONO = {
    "11760923": "3001", "11761277": "3002", "11761293": "3003",
    "11761327": "3004", "11761301": "3005", "12273934": "3006",
    "11761319": "3007", "12273900": "3008", "12273926": "3009",
    "14632012": "3010", "15186547": "3011", "13048756": "3012",
    "15186539": "3013", "16219602": "3200", "16591240": "3201",
    "16634586": "3202", "14885164": "2005", "19827765": "3203",
    "11761350": "2001", "12161154": "2002", "14294946": "2003",
    "15926645": "2210", "11831583": "4002", "12161162": "4001",
    "12161121": "4003", "12161139": "4004", "12874475": "4005",
    "15190309": "4006", "14468144": "4006", "12500773": "4008",
    "14468151": "4009", "14651459": "4010", "15444946": "4200",
    "16062176": "4253", "20836698": "4007", "72806854": "4203",
    "20719829": "4201", "15536188": "6101", "12637294": "6102",
    "11844685": "6103", "15536170": "6106", "17549197": "6108",
}


# =========================================================
# FUNCIONES DE APOYO
# =========================================================

def normalizar_texto(valor) -> str:
    return re.sub(r"\s+", " ", str(valor).strip())


def numero_limpio(valor) -> str:
    if pd.isna(valor):
        return ""
    texto = normalizar_texto(valor).upper()
    texto = re.sub(r"\.0$", "", texto)
    numeros = re.findall(r"\d+", texto)
    return numeros[0] if numeros else ""


def periodo_de(fecha) -> str:
    if pd.isna(fecha):
        return "SIN_FECHA_CONTABLE"
    return fecha.strftime("%Y-%m")


def clasificar_distribuidora(fila):
    ref = numero_limpio(fila.get("Referencia", ""))
    if ref in MAPEO_DISTRIBUIDORA:
        return MAPEO_DISTRIBUIDORA[ref]
    texto = " ".join(
        str(fila.get(c, "")) for c in ["Texto", "Asignación", "Referencia"]
    ).upper()
    if "DOSQ" in texto or "D504" in texto:
        return "Dist Dosquebradas"
    if "ACOPI" in texto or "D503" in texto:
        return "Dist Acopi"
    if "PASTO" in texto or "D505" in texto:
        return "Dist Pasto"
    if "BUGA" in texto or "D502" in texto:
        return "Dist Buga"
    return "Sin clasificar"


def homologar_datafono(fila):
    texto = f"{fila.get('Referencia', '')} {fila.get('Asignación', '')}"
    for numero in re.findall(r"\d{8}", texto):
        if numero in MAPEO_DATAFONO:
            return MAPEO_DATAFONO[numero]
    for numero in re.findall(r"\d{4}", texto):
        if numero in MAPEO_DATAFONO.values():
            return numero
    return ""


def es_nequi(fila) -> bool:
    texto = f"{fila.get('Texto', '')} {fila.get('Asignación', '')} {fila.get('Referencia', '')}".upper()
    return "NEQUI" in texto


def es_valor_redondo(valor) -> bool:
    return valor > 0 and (valor % multiplo_redondo == 0)


# =========================================================
# GATE UNICO DE SEGURIDAD (unico punto que decide si algo entra en la columna O)
# =========================================================

def candidato_seguro(df, id_a, id_b):
    """
    Valida un par de filas contra TODAS las reglas de seguridad antes de
    permitir que se escriban en Candidatos_Conciliacion (columna O).
    Devuelve (ok: bool, motivo: str, es_alerta_texto: bool)
    """
    ra = df.loc[id_a]
    rb = df.loc[id_b]

    periodo_a = ra["Periodo_Contable"]
    periodo_b = rb["Periodo_Contable"]
    if periodo_a == "SIN_FECHA_CONTABLE" or periodo_b == "SIN_FECHA_CONTABLE":
        return False, "Sin periodo contable valido", False
    if periodo_a != periodo_b:
        return False, f"Periodo distinto ({periodo_a} vs {periodo_b})", False

    fa = ra["Fecha_Calc"]
    fb = rb["Fecha_Calc"]
    if pd.isna(fa) or pd.isna(fb):
        return False, "Fecha valor invalida", False

    diferencia_dias = abs((fa - fb).days)
    es_alerta_texto = diferencia_dias > 0
    if diferencia_dias > dias_alerta_texto:
        return False, f"Diferencia de fecha fuera de rango ({diferencia_dias} dias)", False

    banco_a = str(ra["Clave referencia 3"]).strip()
    banco_b = str(rb["Clave referencia 3"]).strip()
    if banco_a != banco_b:
        return False, f"Banco distinto ({banco_a} vs {banco_b})", False

    imp_a = abs(ra["Importe en moneda local"])
    imp_b = abs(rb["Importe en moneda local"])
    dif_valor = abs(imp_a - imp_b)
    max_imp = max(imp_a, imp_b, 1)
    dentro_tolerancia = (
        dif_valor == 0
        or dif_valor <= tolerancia_valor_abs
        or (dif_valor / max_imp) <= tolerancia_valor_pct
    )
    if not dentro_tolerancia:
        return False, f"Diferencia de valor fuera de tolerancia (${dif_valor:,.0f})", False

    dist_a = str(ra.get("Distribuidora", "")).strip()
    dist_b = str(rb.get("Distribuidora", "")).strip()
    if dist_a not in ("", "Sin clasificar") and dist_b not in ("", "Sin clasificar"):
        if dist_a != dist_b:
            return False, f"Distribuidora distinta ({dist_a} vs {dist_b})", False

    ref_a = str(ra.get("Referencia_Limpia", "")).strip()
    ref_b = str(rb.get("Referencia_Limpia", "")).strip()
    asi_a = str(ra.get("Asignacion_Limpia", "")).strip()
    asi_b = str(rb.get("Asignacion_Limpia", "")).strip()
    refh_a = str(ra.get("Ref_Homologada", "")).strip()
    refh_b = str(rb.get("Ref_Homologada", "")).strip()

    referencia_valida = (
        (ref_a and ref_b and ref_a == ref_b)
        or (asi_a and ref_b and asi_a == ref_b)
        or (ref_a and asi_b and ref_a == asi_b)
        or (asi_a and asi_b and asi_a == asi_b)
        or (refh_a and refh_b and refh_a == refh_b)
    )
    sin_referencia_en_ambos = not (ref_a or asi_a or refh_a) and not (ref_b or asi_b or refh_b)

    if not referencia_valida and not sin_referencia_en_ambos:
        return False, "Referencia/Asignacion no coincide", False

    if es_nequi(ra) or es_nequi(rb):
        return True, "Candidato NEQUI: requiere verificacion manual", es_alerta_texto

    if sin_referencia_en_ambos:
        return True, "Candidato sin referencia (cruce unico banco/importe/fecha)", es_alerta_texto

    return True, "Candidato valido (todas las reglas cumplidas)", es_alerta_texto


def escribir_candidato(df, id_a, id_b, estado_si_ok, prefijo_comentario):
    """Punto UNICO de escritura hacia la columna O."""
    ok, motivo, es_alerta = candidato_seguro(df, id_a, id_b)

    if not ok:
        for idx in (id_a, id_b):
            anterior = str(df.at[idx, "Comentario"])
            bloqueo = f"Relacion bloqueada por regla de seguridad: {motivo}."
            df.at[idx, "Comentario"] = bloqueo if not anterior else anterior + " | " + bloqueo
        return False

    def fmt(idx):
        doc = int(df.at[idx, "Nº documento"]) if pd.notna(df.at[idx, "Nº documento"]) else 0
        clave = str(df.at[idx, "Clave contabil."])
        clase = str(df.at[idx, "Clase de documento"]) if "Clase de documento" in df.columns else ""
        if clase and clase.lower() != "nan":
            return f"{doc} ({clase}={clave})"
        return f"{doc} (Clv {clave})"

    candidatos_texto = f"{fmt(id_a)} | {fmt(id_b)}"
    estado_final = "Alerta de texto - Diferencia de fecha (mismo periodo)" if es_alerta else estado_si_ok
    comentario_final = f"{prefijo_comentario} {motivo}."

    for idx in (id_a, id_b):
        if df.at[idx, "Estado_Conciliacion"] == "Pendiente":
            df.at[idx, "Estado_Conciliacion"] = estado_final
        df.at[idx, "Candidatos_Conciliacion"] = candidatos_texto
        anterior = str(df.at[idx, "Comentario"])
        df.at[idx, "Comentario"] = comentario_final if not anterior else anterior + " | " + comentario_final

    return True


# =========================================================
# EMPAREJADOR GENERICO (elimina duplicacion entre niveles)
# =========================================================

def emparejar_por_grupo(df, usados, llaves, estado_unico, estado_multiple, comentario_base, filtro=None, ambiguo_si_no_redondo=False):
    """
    Agrupa filas Pendiente por 'llaves' (deben incluir siempre banco+importe+periodo)
    y empareja FIFO entre Clave contabil. 40 y 50 dentro de cada grupo.
    Cada par pasa por escribir_candidato (que aplica candidato_seguro).

    Si ambiguo_si_no_redondo=True: cuando un grupo tiene mas de un candidato en
    cada lado y el importe NO es multiplo redondo, NO se fuerza un emparejamiento
    FIFO; se deja solo un comentario informativo sin llenar la columna O.
    """
    pend = df[df["Estado_Conciliacion"] == "Pendiente"].copy()
    if filtro is not None:
        pend = filtro(pend)

    deb = pend[pend["Clave contabil."] == "40"]
    cre = pend[pend["Clave contabil."] == "50"]

    if deb.empty or cre.empty:
        return

    for grp_val, sub40 in deb.groupby(llaves):
        valores_grp = grp_val if isinstance(grp_val, tuple) else (grp_val,)
        sub50 = cre
        for k, v in zip(llaves, valores_grp):
            sub50 = sub50[sub50[k] == v]
        if sub50.empty:
            continue

        s40 = sub40[~sub40["ID_Temp"].isin(usados)].sort_values(["Fecha_Calc", "Nº documento"])
        s50 = sub50[~sub50["ID_Temp"].isin(usados)].sort_values(["Fecha_Calc", "Nº documento"])
        if s40.empty or s50.empty:
            continue

        n = min(len(s40), len(s50))

        if ambiguo_si_no_redondo and n > 1:
            importe_grp = s40["Abs_Importe"].iloc[0]
            if not es_valor_redondo(importe_grp):
                docs40 = ", ".join(str(int(d)) for d in s40["Nº documento"].tolist())
                docs50 = ", ".join(str(int(d)) for d in s50["Nº documento"].tolist())
                for idx in list(s40["ID_Temp"]) + list(s50["ID_Temp"]):
                    anterior = str(df.at[idx, "Comentario"])
                    info = (
                        f"Multiples candidatos ambiguos (importe no redondo), sin cruce automatico. "
                        f"Debitos: {docs40}. Creditos: {docs50}. Solicitar soporte."
                    )
                    df.at[idx, "Comentario"] = info if not anterior else anterior + " | " + info
                    df.at[idx, "Estado_Conciliacion"] = "Sugerencia - Solicitar soporte"
                continue

        estado = estado_unico if n == 1 else estado_multiple
        for i in range(n):
            ida = s40.iloc[i]["ID_Temp"]
            idb = s50.iloc[i]["ID_Temp"]
            if ida in usados or idb in usados:
                continue
            if escribir_candidato(df, ida, idb, estado, comentario_base):
                usados.update([ida, idb])


# =========================================================
# LECTURA Y PREPARACION DEL ARCHIVO
# =========================================================

def preparar_archivo(archivo):
    if archivo.name.lower().endswith(".csv"):
        df = pd.read_csv(archivo)
    else:
        hojas = pd.read_excel(archivo, sheet_name=None)
        validas = [h for h in hojas.values() if not h.dropna(how="all").empty]
        if not validas:
            raise ValueError("El archivo no contiene hojas con datos.")
        df = pd.concat(validas, ignore_index=True)

    df.columns = [normalizar_texto(c) for c in df.columns]

    alias = {
        "Asignacion": "Asignación",
        "N documento": "Nº documento",
        "N doc.": "Nº documento",
        "Clave contabiliz.": "Clave contabil.",
        "CT": "Clave contabil.",
        "Importe en ML": "Importe en moneda local",
        "Fe-valor": "Fecha valor",
        "Fecha de documento": "Fe.contabilización",
        "Clase doc.": "Clase de documento",
    }
    for viejo, nuevo in alias.items():
        if viejo in df.columns and nuevo not in df.columns:
            df.rename(columns={viejo: nuevo}, inplace=True)

    requeridas = [
        "Asignación", "Nº documento", "Clave contabil.", "Referencia",
        "Fe.contabilización", "Fecha valor", "Importe en moneda local",
        "Clave referencia 3",
    ]
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas obligatorias: {faltantes}")

    df = df.copy()
    df["Nº documento"] = pd.to_numeric(df["Nº documento"], errors="coerce")
    df["Clave contabil."] = (
        df["Clave contabil."].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    )
    df["Importe en moneda local"] = pd.to_numeric(df["Importe en moneda local"], errors="coerce")
    df["Fecha valor"] = pd.to_datetime(df["Fecha valor"], errors="coerce")
    df["Fe.contabilización"] = pd.to_datetime(df["Fe.contabilización"], errors="coerce")
    df["Clave referencia 3"] = df["Clave referencia 3"].ffill().astype(str).str.strip()

    filas_antes = len(df)
    filas_descartadas = df[
        df["Nº documento"].isna()
        | df["Clave contabil."].isna()
        | df["Fecha valor"].isna()
    ].copy()
    df = df.dropna(subset=["Nº documento", "Clave contabil.", "Fecha valor"]).reset_index(drop=True)
    filas_excluidas = filas_antes - len(df)

    df["Periodo_Contable"] = df["Fe.contabilización"].map(periodo_de)
    df["Fecha_Calc"] = df["Fecha valor"]
    df["Abs_Importe"] = df["Importe en moneda local"].abs()
    df["Referencia_Limpia"] = df["Referencia"].map(numero_limpio)
    df["Asignacion_Limpia"] = df["Asignación"].map(numero_limpio)
    df["Distribuidora"] = df.apply(clasificar_distribuidora, axis=1)
    df["Ref_Homologada"] = df.apply(homologar_datafono, axis=1)
    df["Estado_Conciliacion"] = "Pendiente"
    df["Comentario"] = ""
    df["Candidatos_Conciliacion"] = ""
    df["ID_Temp"] = df.index

    return df, filas_descartadas, filas_excluidas


# =========================================================
# MOTOR DE CONCILIACION
# =========================================================

def conciliacion_definitiva(df):
    usar_ipcb = "Clase de documento" in df.columns
    usados = set()

    # ---------------------------------------------------
    # Nivel 0: documentos con posiciones repetidas que netean
    # exacto dentro del mismo Nº documento (misma transaccion).
    # ---------------------------------------------------
    for doc_val, grupo in df.groupby("Nº documento"):
        if len(grupo) <= 1:
            continue
        g40 = grupo[grupo["Clave contabil."] == "40"]
        g50 = grupo[grupo["Clave contabil."] == "50"]
        if g40.empty or g50.empty:
            continue
        if round(g40["Importe en moneda local"].abs().sum(), 2) != round(g50["Importe en moneda local"].abs().sum(), 2):
            continue
        ids = grupo["ID_Temp"].tolist()
        texto = " | ".join(
            f"{int(r['Nº documento'])} (Clv {r['Clave contabil.']})" for _, r in grupo.iterrows()
        )
        for idx in ids:
            if df.at[idx, "Estado_Conciliacion"] == "Pendiente":
                df.at[idx, "Estado_Conciliacion"] = "Conciliado - Documento neto (mismo Nro documento)"
            df.at[idx, "Candidatos_Conciliacion"] = texto
            df.at[idx, "Comentario"] = f"Documento {int(doc_val)}: clave 40 y 50 netean exacto."
        usados.update(ids)

    filtro_no_ip = (lambda d: d[d["Clase de documento"].astype(str).str.upper() != "IP"]) if usar_ipcb else None

    # ---------------------------------------------------
    # Nivel 1: cruce exacto por Referencia_Limpia
    # (banco + importe + periodo + referencia_limpia)
    # ---------------------------------------------------
    emparejar_por_grupo(
        df, usados,
        llaves=["Clave referencia 3", "Abs_Importe", "Periodo_Contable", "Referencia_Limpia"],
        estado_unico="Conciliado - Cruce exacto",
        estado_multiple="Conciliado - Cruce exacto (multiple)",
        comentario_base="Cruce por banco, importe, fecha, periodo y referencia limpia.",
        filtro=lambda d: d[d["Referencia_Limpia"] != ""],
    )

    # ---------------------------------------------------
    # Nivel 1B: cruce Asignacion_Limpia (lado 40) contra
    # Referencia_Limpia (lado 50). Requiere merge cruzado
    # porque las columnas de agrupacion son distintas por lado.
    # ---------------------------------------------------
    pend = df[df["Estado_Conciliacion"] == "Pendiente"]
    deb = pend[(pend["Clave contabil."] == "40") & (pend["Asignacion_Limpia"] != "")].copy()
    cre = pend[(pend["Clave contabil."] == "50") & (pend["Referencia_Limpia"] != "")].copy()
    if not deb.empty and not cre.empty:
        deb["n"] = deb.groupby(["Clave referencia 3", "Abs_Importe", "Periodo_Contable", "Asignacion_Limpia"]).cumcount()
        cre["n"] = cre.groupby(["Clave referencia 3", "Abs_Importe", "Periodo_Contable", "Referencia_Limpia"]).cumcount()
        cruce = deb.merge(
            cre,
            left_on=["Clave referencia 3", "Abs_Importe", "Periodo_Contable", "Asignacion_Limpia", "n"],
            right_on=["Clave referencia 3", "Abs_Importe", "Periodo_Contable", "Referencia_Limpia", "n"],
            suffixes=("_40", "_50"),
        )
        for _, r in cruce.iterrows():
            ida, idb = r["ID_Temp_40"], r["ID_Temp_50"]
            if ida in usados or idb in usados:
                continue
            if escribir_candidato(
                df, ida, idb,
                "Conciliado - Cruce exacto (Referencia limpia)",
                "Cruce por asignacion limpia contra referencia limpia, banco/importe/fecha/periodo iguales."
            ):
                usados.update([ida, idb])

    # ---------------------------------------------------
    # Nivel 1C: cruce multiple IP vs CB con referencia homologada,
    # validado posicion por posicion antes de aceptar el agregado.
    # ---------------------------------------------------
    if usar_ipcb:
        pend = df[df["Estado_Conciliacion"] == "Pendiente"]
        ip = pend[
            (pend["Clase de documento"].astype(str).str.upper() == "IP")
            & (pend["Clave contabil."] == "40")
            & (pend["Ref_Homologada"] != "")
        ].copy()
        cb = pend[
            (pend["Clase de documento"].astype(str).str.upper() == "CB")
            & (pend["Clave contabil."] == "50")
            & (pend["Ref_Homologada"] != "")
        ].copy()

        if not ip.empty and not cb.empty:
            grp_ip = ip.groupby(["Clave referencia 3", "Periodo_Contable", "Ref_Homologada"])["Abs_Importe"].sum().reset_index(name="S_IP")
            grp_cb = cb.groupby(["Clave referencia 3", "Periodo_Contable", "Ref_Homologada"])["Abs_Importe"].sum().reset_index(name="S_CB")
            coincide = grp_ip.merge(grp_cb, on=["Clave referencia 3", "Periodo_Contable", "Ref_Homologada"])
            coincide = coincide[coincide["S_IP"].round(2) == coincide["S_CB"].round(2)]

            for _, r in coincide.iterrows():
                sub_ip = ip[
                    (ip["Clave referencia 3"] == r["Clave referencia 3"])
                    & (ip["Periodo_Contable"] == r["Periodo_Contable"])
                    & (ip["Ref_Homologada"] == r["Ref_Homologada"])
                ]
                sub_cb = cb[
                    (cb["Clave referencia 3"] == r["Clave referencia 3"])
                    & (cb["Periodo_Contable"] == r["Periodo_Contable"])
                    & (cb["Ref_Homologada"] == r["Ref_Homologada"])
                ]
                ids_ip = [i for i in sub_ip["ID_Temp"].tolist() if i not in usados]
                ids_cb = [i for i in sub_cb["ID_Temp"].tolist() if i not in usados]
                if not ids_ip or not ids_cb:
                    continue

                pares_validos = [
                    (a, b) for a in ids_ip for b in ids_cb
                    if candidato_seguro(df, a, b)[0]
                ]
                if not pares_validos:
                    for idx in ids_ip + ids_cb:
                        anterior = str(df.at[idx, "Comentario"])
                        bloqueo = "Cruce IP/CB agregado bloqueado: ninguna posicion individual cumple reglas de seguridad."
                        df.at[idx, "Comentario"] = bloqueo if not anterior else anterior + " | " + bloqueo
                    continue

                texto = " | ".join(
                    f"{int(df.at[i, 'Nº documento'])} (IP=40)" for i in ids_ip
                ) + " | " + " | ".join(
                    f"{int(df.at[i, 'Nº documento'])} (CB=50)" for i in ids_cb
                )
                for idx in ids_ip + ids_cb:
                    if df.at[idx, "Estado_Conciliacion"] == "Pendiente":
                        df.at[idx, "Estado_Conciliacion"] = "Conciliado - Cruce multiple IP/CB"
                    df.at[idx, "Candidatos_Conciliacion"] = texto
                    anterior = str(df.at[idx, "Comentario"])
                    comentario = f"Cruce agregado IP/CB validado posicion por posicion. Ref {r['Ref_Homologada']}, periodo {r['Periodo_Contable']}."
                    df.at[idx, "Comentario"] = comentario if not anterior else anterior + " | " + comentario
                usados.update(ids_ip + ids_cb)

    # ---------------------------------------------------
    # Nivel 1D: sectorizacion por Distribuidora / punto de venta
    # ---------------------------------------------------
    def filtro_distribuidora(d):
        d = d[d["Distribuidora"] != "Sin clasificar"]
        if usar_ipcb:
            d = d[d["Clase de documento"].astype(str).str.upper() != "IP"]
        return d

    emparejar_por_grupo(
        df, usados,
        llaves=["Clave referencia 3", "Abs_Importe", "Periodo_Contable", "Distribuidora"],
        estado_unico="Conciliado - Cruce Distribuidora",
        estado_multiple="Sugerencia fuerte - Sectorizacion FIFO",
        comentario_base="Sectorizacion validada por banco/importe/fecha/periodo/distribuidora.",
        filtro=filtro_distribuidora,
    )

    # ---------------------------------------------------
    # Nivel 2: cruce unico sin referencia (banco+importe+periodo)
    # y valor redondo con multiples candidatos ambiguos.
    # ---------------------------------------------------
    emparejar_por_grupo(
        df, usados,
        llaves=["Clave referencia 3", "Abs_Importe", "Periodo_Contable"],
        estado_unico="Conciliado - Cruce unico",
        estado_multiple="Sugerencia fuerte - Valor redondo FIFO",
        comentario_base="Cruce por banco/importe/fecha/periodo, sin referencia distintiva o valor redondo.",
        filtro=filtro_no_ip,
        ambiguo_si_no_redondo=True,
    )

    # ---------------------------------------------------
    # Nivel 3: alertas textuales de fecha (0 < diferencia <= dias_alerta_texto)
    # y diferencias de valor con referencia, ambas dentro del mismo periodo.
    # candidato_seguro ya marca automaticamente "alerta de texto" cuando
    # diferencia_dias > 0, asi que basta reutilizar emparejar_por_grupo.
    # ---------------------------------------------------
    emparejar_por_grupo(
        df, usados,
        llaves=["Clave referencia 3", "Referencia_Limpia", "Periodo_Contable"],
        estado_unico="Alerta de texto - Diferencia de fecha o valor",
        estado_multiple="Alerta de texto - Diferencia de fecha o valor (multiple)",
        comentario_base="Misma referencia y periodo, con diferencia de fecha o valor dentro de tolerancia.",
        filtro=lambda d: d[d["Referencia_Limpia"] != ""],
    )

    # ---------------------------------------------------
    # Cierre: todo lo que sigue Pendiente sin candidatos queda documentado.
    # ---------------------------------------------------
    pendientes = df["Estado_Conciliacion"] == "Pendiente"
    sin_candidato = pendientes & (df["Candidatos_Conciliacion"] == "")
    if usar_ipcb:
        es_ip = df["Clase de documento"].astype(str).str.upper() == "IP"
        df.loc[sin_candidato & es_ip & (df["Comentario"] == ""), "Comentario"] = (
            "Sin coincidencia - PDV, requiere revision manual."
        )
        df.loc[sin_candidato & ~es_ip & (df["Comentario"] == ""), "Comentario"] = (
            "Sin coincidencia ni sugerencia que cumpla las reglas de seguridad. Revision manual completa."
        )
    else:
        df.loc[sin_candidato & (df["Comentario"] == ""), "Comentario"] = (
            "Sin coincidencia ni sugerencia que cumpla las reglas de seguridad. Revision manual completa."
        )

    return df


# =========================================================
# ESTILO VISUAL
# =========================================================

def estilo(fila):
    estado = str(fila["Estado_Conciliacion"]).lower()
    if "documento neto" in estado:
        color = "#9CC2E5"
    elif "cruce exacto" in estado or "cruce multiple" in estado or "cruce unico" in estado or "cruce distribuidora" in estado:
        color = "#C5D9F1"
    elif "sugerencia" in estado or "fifo" in estado:
        color = "#FFF2CC"
    elif "alerta" in estado or "diferencia" in estado:
        color = "#FDEBD0"
    else:
        color = ""
    return [f"background-color: {color}"] * len(fila)


# =========================================================
# EJECUCION PRINCIPAL
# =========================================================

if archivo is not None:
    try:
        with st.spinner("Procesando con validacion universal de la columna O..."):
            df, filas_descartadas, filas_excluidas = preparar_archivo(archivo)
            df = conciliacion_definitiva(df)

            columnas_tecnicas = ["ID_Temp", "Abs_Importe", "Fecha_Calc"]
            final = df.drop(columns=columnas_tecnicas, errors="ignore").copy()

            for col in ["Fe.contabilización", "Fecha valor"]:
                if col in final.columns:
                    final[col] = pd.to_datetime(final[col], errors="coerce").dt.strftime("%d/%m/%Y")

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final.style.apply(estilo, axis=1).to_excel(writer, index=False, sheet_name="CONCILIACION_FINAL")
                final[final["Estado_Conciliacion"].eq("Pendiente")].to_excel(writer, index=False, sheet_name="PENDIENTES")
                final[
                    final["Estado_Conciliacion"].str.contains("Alerta|Sugerencia|Diferencia", case=False, na=False)
                ].to_excel(writer, index=False, sheet_name="ALERTAS_Y_SUGERENCIAS")
                if not filas_descartadas.empty:
                    filas_descartadas.to_excel(writer, index=False, sheet_name="DESCARTADAS_SIN_DOC_O_FECHA")

            st.success(
                "Procesamiento terminado. La columna Candidatos_Conciliacion solo "
                "contiene relaciones que pasaron TODAS las reglas de seguridad."
            )

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Conciliadas", int(final["Estado_Conciliacion"].str.contains("Conciliado", na=False).sum()))
            c2.metric(
                "Alertas/Sugerencias/Diferencias",
                int(final["Estado_Conciliacion"].str.contains("Alerta|Sugerencia|Diferencia", case=False, na=False).sum()),
            )
            c3.metric("Pendientes", int(final["Estado_Conciliacion"].eq("Pendiente").sum()))
            c4.metric("Filas descartadas", len(filas_descartadas))

            if filas_excluidas > 0:
                st.warning(f"Se excluyeron {filas_excluidas} filas sin documento, clave o fecha valida.")

            st.download_button(
                "Descargar Excel corregido",
                data=output.getvalue(),
                file_name="Conciliacion_integral_definitiva.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as error:
        st.error(f"Error controlado: {error}")
        st.exception(error)
