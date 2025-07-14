import os
import sys
import re
import csv
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import unicodedata
import subprocess
import threading

# ------------- Configuración y variables fijas ------------------

RUTA_TXT = r"C:\Users\perei\OneDrive\EJPC\RUBY REAL ESTATE\QUINTA KO'OK TANIL\REPORTES\ASISTENCIA\_chat.txt"
CARPETA_SALIDA = r"C:\Users\perei\OneDrive\EJPC\RUBY REAL ESTATE\QUINTA KO'OK TANIL\REPORTES\ASISTENCIA"
NOMBRE_CSV_ASISTENCIA = "asistencia_empleados.csv"
RUTA_CSV_ASISTENCIA = os.path.join(CARPETA_SALIDA, NOMBRE_CSV_ASISTENCIA)
RUTA_EXCEL = r"D:\Quinta Koox Tanil\SISTEMA INTEGRAL ADMINISTRATIVO.xlsm"
NOMBRE_CSV_BASE1 = "asistencia_base1.csv"
RUTA_CSV_BASE1 = os.path.join(CARPETA_SALIDA, NOMBRE_CSV_BASE1)
NOMBRE_CSV_FUSIONADO = "asistencia_empleados_fusionado.csv"
RUTA_CSV_FUSIONADO = os.path.join(CARPETA_SALIDA, NOMBRE_CSV_FUSIONADO)
NOMBRE_EXCEL_FUSIONADO = "Nomina.xlsx"
RUTA_EXCEL_FUSIONADO = os.path.join(CARPETA_SALIDA, NOMBRE_EXCEL_FUSIONADO)
NOMBRE_CSV_SEMANAL = "Nomina.csv"
RUTA_CSV_SEMANAL = os.path.join(CARPETA_SALIDA, NOMBRE_CSV_SEMANAL)

EMPLEADOS = {
    "Arcadio Pech May": {
        "id": "E001",
        "hora_entrada": "08:00",
        "tolerancia": 10,
        "cuota": 37.50
    },
    "Kevin Abisai Montuy Canche": {  # Nombre estandarizado para Kevin
        "id": "E002",
        "hora_entrada": "07:30",
        "tolerancia": 10,
        "cuota": 50.00
    }
}

EMPLEADOS_INFO = {
    "ARCADIO PECH MAY": {
        "ID": "E001",
        "cuota_hora": 34.7,
        "hora_entrada": "08:00:00",
        "tolerancia_min": 10
    },
    "KEVIN ABISAI MONTUY CANCHE": {  # Nombre estandarizado para Kevin
        "ID": "E002",
        "cuota_hora": 50.00,
        "hora_entrada": "07:30:00",
        "tolerancia_min": 10
    },
    "LUIS": {
        "ID": "E003",
        "cuota_hora": 34.7,
        "hora_entrada": "08:00:00",
        "tolerancia_min": 10
    }
}

# Mapeo de nombres alternativos a nombres estandarizados
NOMBRES_ESTANDARIZADOS = {
    "KEVIN ABISAI CANCHE MONTUY": "KEVIN ABISAI MONTUY CANCHE",
    "KEVIN ABISAI": "KEVIN ABISAI MONTUY CANCHE",
    "KEVIN MONTUY": "KEVIN ABISAI MONTUY CANCHE",
    "KEVIN CANCHE": "KEVIN ABISAI MONTUY CANCHE",
    "KEVIN": "KEVIN ABISAI MONTUY CANCHE"
}

# Mapeo inverso de IDs a nombres estandarizados
ID_A_NOMBRE = {
    "E001": "ARCADIO PECH MAY",
    "E002": "KEVIN ABISAI MONTUY CANCHE",
    "E003": "LUIS"
}

COLUMNS_FINAL = [
    "Día", "ID", "ID EMPLEADO", "Nombre",
    "Hora de entrada", "Hora de salida", "Horas trabajadas",
    "Estatus", "Cuota del día", "Día (semana)", "editado_manual"
]

REGEX_MSG = r"\[(\d{2}/\d{2}/\d{2,4}), (\d{1,2}:\d{2}:\d{2})\s?([ap]\.m\.)?\] ([^:]+): (.*)"
REGEX_INICIO = r"📕\s?\*?INICIO\*?"
REGEX_SALIDA = r"📒\s?\*?SALIDA\*?"

# ------------- Funciones de procesamiento ------------------

def quitar_acentos(cadena):
    if pd.isnull(cadena):
        return ""
    nfkd = unicodedata.normalize('NFKD', str(cadena))
    return "".join([c for c in nfkd if not unicodedata.combining(c)])

def obtener_dia_semana(dt):
    try:
        dias = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo']
        idx = dt.weekday()
        return dias[idx].capitalize()
    except Exception:
        return ""

def safe_float(val):
    try:
        return float(val)
    except Exception:
        return 0.0

def normalizar_nombre(nombre):
    """Normaliza el nombre para que siempre use el formato estándar"""
    if pd.isnull(nombre) or nombre == "":
        return ""
    nombre_norm = quitar_acentos(nombre).upper().strip()
    
    # Verificar si es uno de los nombres alternativos conocidos
    if nombre_norm in NOMBRES_ESTANDARIZADOS:
        return NOMBRES_ESTANDARIZADOS[nombre_norm]
    
    # Verificar palabras claves para identificar a los empleados
    if "KEVIN" in nombre_norm or "ABISAI" in nombre_norm or "MONTUY" in nombre_norm or "CANCHE" in nombre_norm:
        return "KEVIN ABISAI MONTUY CANCHE"
    if "ARCADIO" in nombre_norm or "PECH" in nombre_norm:
        return "ARCADIO PECH MAY"
    if "LUIS" in nombre_norm:
        return "LUIS"
    
    return nombre_norm

def get_empleado_info(nombre):
    """Obtiene la información del empleado basado en su nombre normalizado"""
    nombre = normalizar_nombre(nombre)
    
    # Búsqueda directa
    if nombre in EMPLEADOS_INFO:
        return EMPLEADOS_INFO[nombre]
    
    # Si no encuentra, buscar por coincidencias parciales
    for key in EMPLEADOS_INFO.keys():
        if key.split()[0] in nombre:
            return EMPLEADOS_INFO[key]
    
    return None

def calcular_horas_trabajadas(entrada, salida, nombre=None):
    try:
        if not entrada or not salida:
            return 0.0
        h1 = datetime.strptime(entrada.zfill(8), "%H:%M:%S")
        h2 = datetime.strptime(salida.zfill(8), "%H:%M:%S")
        if nombre:
            info = get_empleado_info(nombre)
            if info:
                hora_oficial = datetime.strptime(info["hora_entrada"], "%H:%M:%S")
                h1 = max(h1, hora_oficial)
        diff = (h2 - h1).total_seconds()/3600
        if diff < 0:
            diff += 24
        return round(diff, 2)
    except Exception:
        return 0.0

def calcular_cuota_dia(nombre, horas_trabajadas):
    info = get_empleado_info(nombre)
    if not info or horas_trabajadas <= 0:
        return 0.0
    cuota = horas_trabajadas * info["cuota_hora"]
    return round(cuota, 2)

def marcar_estatus(nombre, entrada, salida):
    ent = (entrada or "").strip()
    sal = (salida or "").strip()
    if ent in ["00:00:00", "0:00:00", "n/a", ""] and sal in ["00:00:00", "0:00:00", "n/a", ""]:
        return "Falta"
    if ent in ["n/a", "", None, "00:00:00", "0:00:00"]:
        return "Falta"
    info = get_empleado_info(nombre)
    if not info:
        return "Falta"
    try:
        hora_real = datetime.strptime(ent.zfill(8), "%H:%M:%S")
        hora_ref = datetime.strptime(info["hora_entrada"], "%H:%M:%S") + timedelta(minutes=info["tolerancia_min"])
        if hora_real <= hora_ref:
            return "A tiempo"
        else:
            return "Retardo"
    except Exception:
        return "A tiempo"

def limpiar_y_formatear_fechas(df, columna_fecha="Día"):
    def filtrar_fecha(dt):
        if pd.isnull(dt) or dt == "":
            return pd.NaT
        try:
            dt_str = str(dt)
            if " " in dt_str:
                dt_str = dt_str.split(" ")[0]
            m = re.match(r"(\d{4})[-/](\d{2})[-/](\d{2})", dt_str)
            if m:
                dt_str = f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
            return pd.to_datetime(dt_str, dayfirst=True, errors='coerce')
        except Exception:
            return pd.NaT
    if df.empty or columna_fecha not in df.columns:
        df['Día_datetime'] = pd.NaT
        df['Día'] = ""
        df['Día (semana)'] = ""
        return df
    df = df.copy()
    df['Día_datetime'] = df[columna_fecha].apply(filtrar_fecha)
    df = df[df['Día_datetime'].notnull()].copy()
    df['Día'] = df['Día_datetime'].dt.strftime("%d/%m/%Y")
    df['Día (semana)'] = df['Día_datetime'].apply(lambda x: quitar_acentos(obtener_dia_semana(x)) if not pd.isnull(x) else '')
    return df

def transformar_base1(df):
    df_nuevo = pd.DataFrame()
    df_nuevo["Día"] = df["FECHA"]
    
    # Normalizar los nombres para resolver el problema de los nombres diferentes
    df_nuevo["Nombre"] = df["EMPLEADO"].astype(str).str.title().apply(quitar_acentos).apply(normalizar_nombre)
    df_nuevo["ID EMPLEADO"] = df_nuevo["Nombre"].apply(lambda x: get_empleado_info(x)["ID"] if get_empleado_info(x) else "")
    
    # Estandarizar nombres basados en ID
    df_nuevo["Nombre"] = df_nuevo["ID EMPLEADO"].apply(lambda x: ID_A_NOMBRE.get(x, x))
    
    df_nuevo["Hora de entrada"] = df["ENTRADA"].fillna("").astype(str)
    df_nuevo["Hora de salida"] = df["SALIDA"].fillna("").astype(str)
    df_nuevo["Horas trabajadas"] = [
        calcular_horas_trabajadas(ent, sal, nombre)
        for ent, sal, nombre in zip(df_nuevo["Hora de entrada"], df_nuevo["Hora de salida"], df_nuevo["Nombre"])
    ]
    df_nuevo = limpiar_y_formatear_fechas(df_nuevo, "Día")
    
    # Generar ID único para cada registro (ID_EMPLEADO + FECHA)
    df_nuevo['ID'] = df_nuevo.apply(lambda row: f"{row['ID EMPLEADO']}_{row['Día_datetime'].strftime('%d%m%Y')}" if pd.notnull(row['Día_datetime']) else "", axis=1)
    
    df_nuevo["Cuota del día"] = [
        calcular_cuota_dia(row["Nombre"], safe_float(row["Horas trabajadas"]))
        for _, row in df_nuevo.iterrows()
    ]
    df_nuevo["Estatus"] = [
        marcar_estatus(row["Nombre"], row["Hora de entrada"], row["Hora de salida"])
        for _, row in df_nuevo.iterrows()
    ]
    df_nuevo["editado_manual"] = 0
    for col in COLUMNS_FINAL:
        if col not in df_nuevo.columns:
            df_nuevo[col] = ""
    return df_nuevo[COLUMNS_FINAL]

def transformar_empleados(df):
    df_nuevo = pd.DataFrame()
    df_nuevo["Día"] = df["Día"]
    
    # Normalizar los nombres para resolver el problema de los nombres diferentes
    df_nuevo["Nombre"] = df["Nombre"].astype(str).str.title().apply(quitar_acentos).apply(normalizar_nombre)
    df_nuevo["ID EMPLEADO"] = df["ID EMPLEADO"] if "ID EMPLEADO" in df.columns else df_nuevo["Nombre"].apply(lambda x: get_empleado_info(x)["ID"] if get_empleado_info(x) else "")
    
    # Estandarizar nombres basados en ID
    df_nuevo["Nombre"] = df_nuevo["ID EMPLEADO"].apply(lambda x: ID_A_NOMBRE.get(x, x))
    
    df_nuevo["Hora de entrada"] = df["Hora de entrada"].fillna("").astype(str)
    df_nuevo["Hora de salida"] = df["Hora de salida"].fillna("").astype(str)
    if "Horas trabajadas" in df.columns:
        df_nuevo["Horas trabajadas"] = [safe_float(x) for x in df["Horas trabajadas"]]
    else:
        df_nuevo["Horas trabajadas"] = [
            calcular_horas_trabajadas(ent, sal, nombre)
            for ent, sal, nombre in zip(df_nuevo["Hora de entrada"], df_nuevo["Hora de salida"], df_nuevo["Nombre"])
        ]
    df_nuevo = limpiar_y_formatear_fechas(df_nuevo, "Día")
    
    # Generar ID único para cada registro (ID_EMPLEADO + FECHA)
    df_nuevo['ID'] = df_nuevo.apply(lambda row: f"{row['ID EMPLEADO']}_{row['Día_datetime'].strftime('%d%m%Y')}" if pd.notnull(row['Día_datetime']) else "", axis=1)
    
    df_nuevo["Cuota del día"] = [
        calcular_cuota_dia(row["Nombre"], safe_float(row["Horas trabajadas"]))
        for _, row in df_nuevo.iterrows()
    ]
    df_nuevo["Estatus"] = [
        marcar_estatus(row["Nombre"], row["Hora de entrada"], row["Hora de salida"])
        for _, row in df_nuevo.iterrows()
    ]
    df_nuevo["editado_manual"] = 0
    for col in COLUMNS_FINAL:
        if col not in df_nuevo.columns:
            df_nuevo[col] = ""
    return df_nuevo[COLUMNS_FINAL]

def procesar_excel_a_csv(ruta_excel, ruta_csv_base1):
    if not os.path.exists(ruta_excel):
        print(f"Advertencia: El archivo Excel '{ruta_excel}' no fue encontrado. Saltando este paso.")
        return False
    try:
        df_hojaA = pd.read_excel(ruta_excel, sheet_name="A")
        df_hojaA.to_csv(ruta_csv_base1, index=False, encoding="utf-8")
        print(f"Guardada hoja 'A' de Excel como '{ruta_csv_base1}'.")
        return True
    except Exception as e:
        print(f"Error procesando Excel: {e}")
        return False

def generar_resumen_semanal_return_df(df):
    df['FECHA'] = pd.to_datetime(df['FECHA'], format='%d/%m/%Y', errors='coerce')
    df = df[df['FECHA'].notnull()]
    df['JORNADA'] = pd.to_numeric(df['JORNADA'], errors='coerce').fillna(0)
    df['CUOTA'] = pd.to_numeric(df['CUOTA'], errors='coerce').fillna(0)
    df['Semana'] = df['FECHA'].dt.isocalendar().week
    df['Año'] = df['FECHA'].dt.year
    df['Dia_inicio_semana'] = df['FECHA'] - pd.to_timedelta(df['FECHA'].dt.weekday, unit='D')
    df['Dia_fin_semana'] = df['Dia_inicio_semana'] + pd.Timedelta(days=6)
    df['EsFalta'] = df['ESTATUS'].str.lower() == 'falta'
    df['EsRetardo'] = df['ESTATUS'].str.lower() == 'retardo'
    df['EsLaborado'] = df['JORNADA'] > 0

    # Normalizar los nombres antes de agrupar
    df['NOMBRE_NORM'] = df['NOMBRE'].apply(normalizar_nombre)
    
    # Usar los nombres normalizados para agrupar
    resumen = df.groupby(['Año', 'Semana', 'ID EMPLEADO', 'NOMBRE_NORM', 'Dia_inicio_semana', 'Dia_fin_semana']).agg(
        Dias_laborados=('EsLaborado', 'sum'),
        Horas_trabajadas=('JORNADA', 'sum'),
        Pago_semana=('CUOTA', 'sum'),
        Faltas=('EsFalta', 'sum'),
        Retardos=('EsRetardo', 'sum')
    ).reset_index()

    # Renombrar la columna normalizada de vuelta a NOMBRE
    resumen = resumen.rename(columns={'NOMBRE_NORM': 'NOMBRE'})
    
    # Estandarizar nombres basados en ID
    resumen['NOMBRE'] = resumen['ID EMPLEADO'].apply(lambda x: ID_A_NOMBRE.get(x, x))

    # Generar ID único para cada registro semanal (ID_EMPLEADO + SEMANA + AÑO)
    resumen['ID'] = resumen.apply(lambda row: f"{row['ID EMPLEADO']}_{row['Semana']:02d}_{row['Año']}", axis=1)
    
    resumen = resumen.reset_index(drop=True)
    return resumen

def autoajustar_columnas(writer, df, sheet_name):
    worksheet = writer.sheets[sheet_name]
    for idx, col in enumerate(df.columns):
        max_len = max(
            df[col].astype(str).map(len).max(),
            len(str(col))
        ) + 2
        worksheet.set_column(idx, idx, max_len)

def fusionar_asistencias(ruta_base1, ruta_empleados, ruta_fusionado, ruta_excel_fusionado, ruta_resumen_semana):
    if os.path.exists(ruta_fusionado):
        df_fusionado_prev = pd.read_csv(ruta_fusionado, encoding='utf-8')
        if 'editado_manual' not in df_fusionado_prev.columns:
            df_fusionado_prev['editado_manual'] = 0
    else:
        df_fusionado_prev = pd.DataFrame()

    if not os.path.exists(ruta_base1):
        print(f"Advertencia: No se encontró '{ruta_base1}'. Solo se usará el archivo de empleados.")
        df_base1 = pd.DataFrame()
    else:
        df_base1 = pd.read_csv(ruta_base1, encoding='utf-8')

    if not os.path.exists(ruta_empleados):
        print(f"Error: No se encontró '{ruta_empleados}'. No se puede fusionar asistencias.")
        return False

    df_empleados = pd.read_csv(ruta_empleados, encoding='utf-8')

    if not df_base1.empty:
        df_base1 = transformar_base1(df_base1)
    else:
        df_base1 = pd.DataFrame(columns=COLUMNS_FINAL)
    df_empleados = transformar_empleados(df_empleados)

    df_total = pd.concat([df_empleados, df_base1], ignore_index=True)
    df_total = df_total[df_total['Día'].notnull() & (df_total['Día'] != "")]
    df_total['Día_datetime'] = pd.to_datetime(df_total['Día'], format='%d/%m/%Y', errors='coerce')
    df_total = df_total[df_total['Día_datetime'].notnull()]
    df_total = df_total.sort_values(by=["Día_datetime"], ascending=False).reset_index(drop=True)
    df_final = df_total[COLUMNS_FINAL]
    df_final = df_final[df_final['Día'].notnull() & (df_final['Día'] != "")]
    df_final["Día"] = pd.to_datetime(df_final["Día"], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y')

    column_rename = {
        "Día": "FECHA",
        "ID": "ID",
        "ID EMPLEADO": "ID EMPLEADO",
        "Nombre": "NOMBRE",
        "Hora de entrada": "ENTRADA",
        "Hora de salida": "SALIDA",
        "Horas trabajadas": "JORNADA",
        "Estatus": "ESTATUS",
        "Cuota del día": "CUOTA",
        "Día (semana)": "DIA",
        "editado_manual": "EDITADO"
    }
    df_final = df_final.rename(columns=column_rename)
    
    # Normalizar los nombres para unificar registros de Kevin
    df_final['NOMBRE'] = df_final['NOMBRE'].apply(normalizar_nombre)
    
    # Estandarizar nombres basados en ID
    df_final['NOMBRE'] = df_final['ID EMPLEADO'].apply(lambda x: ID_A_NOMBRE.get(x, x))

    resumen = generar_resumen_semanal_return_df(df_final)
    resumen_rename = {
        "ID": "ID",
        "Semana": "SEMANA",
        "ID EMPLEADO": "ID EMPLEADO",
        "NOMBRE": "NOMBRE",
        "Dia_inicio_semana": "INICIO",
        "Dia_fin_semana": "FIN",
        "Dias_laborados": "DIAS LABORADOS",
        "Horas_trabajadas": "JORNADA",
        "Pago_semana": "PAGO",
        "Faltas": "FALTAS",
        "Retardos": "RETARDOS",
        "Año": "AÑO"
    }
    resumen = resumen.rename(columns=resumen_rename)

    if "INICIO" in resumen.columns:
        resumen = resumen.sort_values(by="INICIO", ascending=False).reset_index(drop=True)

    df_final_export = df_final.copy()
    if "FECHA" in df_final_export.columns:
        df_final_export["FECHA"] = pd.to_datetime(df_final_export["FECHA"], errors='coerce').dt.strftime('%d/%m/%Y')

    resumen_export = resumen.copy()
    for col in ["INICIO", "FIN"]:
        if col in resumen_export.columns:
            resumen_export[col] = pd.to_datetime(resumen_export[col], errors='coerce').dt.strftime('%d/%m/%Y')

    df_final_export.to_csv(ruta_fusionado, index=False, encoding='utf-8')
    print(f"Archivo fusionado generado: {ruta_fusionado}")
    resumen_export.to_csv(ruta_resumen_semana, index=False, encoding='utf-8')
    print(f"¡Resumen semanal generado: {ruta_resumen_semana}")

    import xlsxwriter
    with pd.ExcelWriter(ruta_excel_fusionado, engine='xlsxwriter') as writer:
        df_final_export.to_excel(writer, index=False, sheet_name="RD")
        resumen_export.to_excel(writer, index=False, sheet_name="RS")

        workbook = writer.book
        text_fmt = workbook.add_format({'num_format': '@'})
        ws_rd = writer.sheets["RD"]
        ws_rs = writer.sheets["RS"]
        ws_rd.set_column('A:A', 12, text_fmt)
        ws_rs.set_column('E:F', 12, text_fmt)
        autoajustar_columnas(writer, df_final_export, "RD")
        autoajustar_columnas(writer, resumen_export, "RS")

    print(f"Archivo fusionado Excel generado: {ruta_excel_fusionado}")
    return True

def abrir_archivo(ruta):
    if sys.platform.startswith('win') and os.path.exists(ruta):
        os.startfile(ruta)
    elif sys.platform.startswith('darwin'):
        subprocess.call(('open', ruta))
    elif sys.platform.startswith('linux'):
        subprocess.call(('xdg-open', ruta))

def parsear_whatsapp_txt(archivo_txt):
    registros = []
    with open(archivo_txt, encoding="utf-8") as f:
        for linea in f:
            m = re.match(REGEX_MSG, linea)
            if m:
                fecha, hora, ampm, remitente, mensaje = m.groups()
                nombre = remitente.strip()
                if nombre not in EMPLEADOS:
                    continue
                tipo = None
                if re.search(REGEX_INICIO, mensaje, re.IGNORECASE):
                    tipo = "entrada"
                elif re.search(REGEX_SALIDA, mensaje, re.IGNORECASE):
                    tipo = "salida"
                if tipo:
                    if ampm:
                        ampm_str = ampm.lower().replace("a.m.", "AM").replace("p.m.", "PM")
                        dt = datetime.strptime(f"{fecha} {hora} {ampm_str}", "%d/%m/%y %I:%M:%S %p")
                    else:
                        try:
                            dt = datetime.strptime(f"{fecha} {hora}", "%d/%m/%y %H:%M:%S")
                        except Exception:
                            dt = datetime.strptime(f"{fecha} {hora}", "%d/%m/%Y %H:%M:%S")
                    hoy = datetime.now()
                    if not (datetime(2020, 1, 1) <= dt <= hoy):
                        continue
                    registros.append({
                        "fecha": dt.date(),
                        "hora": dt.time(),
                        "nombre": nombre,
                        "tipo": tipo,
                        "dt": dt
                    })
    print(f"[parsear_whatsapp_txt] Registros encontrados: {len(registros)}")
    return registros

def construir_tabla_asistencia(registros):
    asistencia = {}
    for r in registros:
        fecha = r["fecha"]
        nombre = r["nombre"]
        if fecha not in asistencia:
            asistencia[fecha] = {}
        if nombre not in asistencia[fecha]:
            asistencia[fecha][nombre] = {"entrada": None, "salida": None}
        if r["tipo"] == "entrada":
            if asistencia[fecha][nombre]["entrada"] is None or r["dt"] < asistencia[fecha][nombre]["entrada"]:
                asistencia[fecha][nombre]["entrada"] = r["dt"]
        elif r["tipo"] == "salida":
            if asistencia[fecha][nombre]["salida"] is None or r["dt"] > asistencia[fecha][nombre]["salida"]:
                asistencia[fecha][nombre]["salida"] = r["dt"]
    return asistencia

def exportar_csv_asistencia(asistencia, archivo_out):
    if not asistencia:
        return
    fechas = list(asistencia.keys())
    fecha_min = min(fechas)
    fecha_max = max(fechas)
    rango_fechas = pd.date_range(fecha_min, fecha_max)

    filas = []
    for fecha in rango_fechas:
        fecha_dt = fecha.date()
        for nombre in EMPLEADOS:
            datos = asistencia.get(fecha_dt, {}).get(nombre, {"entrada": None, "salida": None})
            entrada = datos["entrada"]
            salida = datos["salida"]
            entrada_str = entrada.strftime("%H:%M:%S") if entrada else "00:00:00"
            salida_str = salida.strftime("%H:%M:%S") if salida else "00:00:00"
            estatus = marcar_estatus(nombre, entrada_str, salida_str)
            horas_trab = 0
            if entrada and salida:
                try:
                    horas_trab = calcular_horas_trabajadas(entrada_str, salida_str, nombre)
                except Exception:
                    horas_trab = 0
            cuota = round(horas_trab * EMPLEADOS[nombre]["cuota"], 2)
            
            # Generar ID único para registro diario
            fecha_str = fecha_dt.strftime("%d%m%Y")
            id_empleado = EMPLEADOS[nombre]["id"]
            id_registro = f"{id_empleado}_{fecha_str}"
            
            # Normalizar el nombre para garantizar consistencia
            nombre_normalizado = nombre
            
            fila = [
                fecha_dt.strftime("%d/%m/%Y") if fecha_dt is not None else "n/a",
                id_registro,  # Nuevo campo ID
                id_empleado,  # ID EMPLEADO
                nombre_normalizado if nombre_normalizado is not None else "n/a",
                entrada_str,
                salida_str,
                horas_trab if horas_trab is not None else "n/a",
                estatus if estatus is not None else "n/a",
                cuota if cuota is not None else "n/a"
            ]
            fila = ["n/a" if str(x).lower() == "null" else x for x in fila]
            filas.append(fila)
    filas_ordenadas = sorted(filas, key=lambda x: datetime.strptime(x[0], '%d/%m/%Y'), reverse=True)
    with open(archivo_out, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Día", "ID", "ID EMPLEADO", "Nombre", "Hora de entrada", "Hora de salida", "Horas trabajadas", "Estatus", "Cuota del día"])
        writer.writerows(filas_ordenadas)

def procesar_asistencia_original():
    print("---------- INICIANDO PROCESO DE ASISTENCIA Y NÓMINA ----------")

    if not os.path.exists(RUTA_TXT):
        print(f"Error: No se encontró el archivo de WhatsApp '{RUTA_TXT}'.")
        return
    print("1) Procesando archivo de WhatsApp...")
    registros = parsear_whatsapp_txt(RUTA_TXT)
    tabla = construir_tabla_asistencia(registros)
    exportar_csv_asistencia(tabla, RUTA_CSV_ASISTENCIA)
    print(f"   > Asistencia de empleados exportada a '{RUTA_CSV_ASISTENCIA}'.")

    print("2) Procesando hoja 'A' del Excel administrativo (si existe)...")
    procesar_excel_a_csv(RUTA_EXCEL, RUTA_CSV_BASE1)

    print("3) Fusionando asistencias (WhatsApp + Excel)...")
    fusion_ok = fusionar_asistencias(RUTA_CSV_BASE1, RUTA_CSV_ASISTENCIA, RUTA_CSV_FUSIONADO, RUTA_EXCEL_FUSIONADO, RUTA_CSV_SEMANAL)
    if not fusion_ok:
        print("   > ERROR: No se pudo fusionar las asistencias. Deteniendo el proceso.")
        return

    print("4) Abriendo resumen semanal generado...")
    abrir_archivo(RUTA_CSV_SEMANAL)

    print("\n¡PROCESO COMPLETO! Todos los archivos fueron generados correctamente.")

# ------------- Clase de la interfaz gráfica ------------------

class AsistenciaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Asistencia y Nómina - Quinta Ko'ok Tanil")
        self.root.geometry("1200x700")  # Ventana más grande para acomodar los filtros
        self.root.iconbitmap("icono.ico") if os.path.exists("icono.ico") else None
        
        # Variables para almacenar datos
        self.df_asistencia = None
        self.df_semanal = None
        
        # Crear notebook (pestañas)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Pestaña de inicio
        self.tab_inicio = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_inicio, text="Inicio")
        
        # Pestaña de asistencia diaria
        self.tab_asistencia = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_asistencia, text="Asistencia Diaria")
        
        # Pestaña de resumen semanal
        self.tab_semanal = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_semanal, text="Resumen Semanal")
        
        # Configurar cada pestaña
        self._configurar_tab_inicio()
        self._configurar_tab_asistencia()
        self._configurar_tab_semanal()
        
        # Cargar datos si existen
        self.cargar_datos_existentes()

    def _configurar_tab_inicio(self):
        # Titulo
        titulo = ttk.Label(self.tab_inicio, text="Sistema de Gestión de Asistencia y Nómina", 
                         font=("Arial", 16, "bold"))
        titulo.pack(pady=20)
        
        # Marco para botones
        frame_botones = ttk.Frame(self.tab_inicio)
        frame_botones.pack(pady=20)
        
        # Botón para procesar asistencia
        btn_procesar = ttk.Button(frame_botones, text="Procesar Asistencia", 
                                 command=self.procesar_asistencia, width=30)
        btn_procesar.grid(row=0, column=0, padx=10, pady=10)
        
        # Botón para abrir el Excel generado
        btn_excel = ttk.Button(frame_botones, text="Abrir Excel Generado", 
                              command=lambda: abrir_archivo(RUTA_EXCEL_FUSIONADO), width=30)
        btn_excel.grid(row=0, column=1, padx=10, pady=10)
        
        # Botón para actualizar datos
        btn_actualizar = ttk.Button(frame_botones, text="Actualizar Datos", 
                                   command=self.cargar_datos_existentes, width=30)
        btn_actualizar.grid(row=1, column=0, padx=10, pady=10, columnspan=2)
        
        # Información de archivos
        frame_info = ttk.LabelFrame(self.tab_inicio, text="Rutas de archivos")
        frame_info.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Mostrar rutas
        ttk.Label(frame_info, text=f"Archivo de chat: {RUTA_TXT}").pack(anchor="w", padx=10, pady=5)
        ttk.Label(frame_info, text=f"Carpeta de salida: {CARPETA_SALIDA}").pack(anchor="w", padx=10, pady=5)
        ttk.Label(frame_info, text=f"Excel fusionado: {RUTA_EXCEL_FUSIONADO}").pack(anchor="w", padx=10, pady=5)
        ttk.Label(frame_info, text=f"Resumen semanal: {RUTA_CSV_SEMANAL}").pack(anchor="w", padx=10, pady=5)

    def _configurar_tab_asistencia(self):
        # Marco para filtros
        frame_filtros = ttk.Frame(self.tab_asistencia)
        frame_filtros.pack(fill="x", padx=10, pady=10)
        
        # Primera fila de filtros
        ttk.Label(frame_filtros, text="Filtrar por empleado:").grid(row=0, column=0, padx=5, pady=5)
        self.combo_empleado = ttk.Combobox(frame_filtros, width=25)
        self.combo_empleado.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frame_filtros, text="Año:").grid(row=0, column=2, padx=5, pady=5)
        self.combo_anio_diario = ttk.Combobox(frame_filtros, width=6)
        self.combo_anio_diario.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(frame_filtros, text="Mes:").grid(row=0, column=4, padx=5, pady=5)
        self.combo_mes = ttk.Combobox(frame_filtros, width=4)
        self.combo_mes.grid(row=0, column=5, padx=5, pady=5)
        
        # Segunda fila de filtros
        ttk.Label(frame_filtros, text="Día:").grid(row=1, column=0, padx=5, pady=5)
        self.combo_dia = ttk.Combobox(frame_filtros, width=4)
        self.combo_dia.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(frame_filtros, text="Estatus:").grid(row=1, column=2, padx=5, pady=5)
        self.combo_estatus = ttk.Combobox(frame_filtros, width=10)
        self.combo_estatus.grid(row=1, column=3, padx=5, pady=5)
        
        # Botones de control
        ttk.Button(frame_filtros, text="Aplicar Filtros", 
                  command=self.aplicar_filtro_asistencia).grid(row=1, column=4, padx=5, pady=5)
        
        ttk.Button(frame_filtros, text="Limpiar Filtros", 
                  command=self.limpiar_filtro_asistencia).grid(row=1, column=5, padx=5, pady=5)
        
        ttk.Button(frame_filtros, text="Exportar a Excel", 
                  command=lambda: self.exportar_tabla(self.df_asistencia_filtrado if hasattr(self, 'df_asistencia_filtrado') else self.df_asistencia, "asistencia_diaria")).grid(row=0, column=6, padx=5, pady=5, rowspan=2)
        
        # Crear treeview para tabla de asistencia
        frame_tabla = ttk.Frame(self.tab_asistencia)
        frame_tabla.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.tree_asistencia = ttk.Treeview(frame_tabla)
        self.tree_asistencia.pack(side="left", fill="both", expand=True)
        
        # Configurar scrollbar vertical
        scrolly = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tree_asistencia.yview)
        self.tree_asistencia.configure(yscrollcommand=scrolly.set)
        scrolly.pack(side="right", fill="y")
        
        # Configurar scrollbar horizontal
        frame_scrollx = ttk.Frame(self.tab_asistencia)
        frame_scrollx.pack(fill="x", padx=10, pady=(0, 5))
        
        scrollx = ttk.Scrollbar(frame_scrollx, orient="horizontal", command=self.tree_asistencia.xview)
        self.tree_asistencia.configure(xscrollcommand=scrollx.set)
        scrollx.pack(fill="x")

    def _configurar_tab_semanal(self):
        # Marco para filtros
        frame_filtros = ttk.Frame(self.tab_semanal)
        frame_filtros.pack(fill="x", padx=10, pady=10)
        
        # Filtros de resumen semanal
        ttk.Label(frame_filtros, text="Filtrar por empleado:").grid(row=0, column=0, padx=5, pady=5)
        self.combo_empleado_semanal = ttk.Combobox(frame_filtros, width=25)
        self.combo_empleado_semanal.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frame_filtros, text="Año:").grid(row=0, column=2, padx=5, pady=5)
        self.combo_anio_semanal = ttk.Combobox(frame_filtros, width=6)
        self.combo_anio_semanal.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(frame_filtros, text="Semana:").grid(row=0, column=4, padx=5, pady=5)
        self.combo_semana = ttk.Combobox(frame_filtros, width=6)
        self.combo_semana.grid(row=0, column=5, padx=5, pady=5)
        
        ttk.Button(frame_filtros, text="Aplicar Filtros", 
                  command=self.aplicar_filtro_semanal).grid(row=0, column=6, padx=5, pady=5)
        
        ttk.Button(frame_filtros, text="Limpiar Filtros", 
                  command=self.limpiar_filtro_semanal).grid(row=0, column=7, padx=5, pady=5)
        
        ttk.Button(frame_filtros, text="Exportar a Excel", 
                  command=lambda: self.exportar_tabla(self.df_semanal_filtrado if hasattr(self, 'df_semanal_filtrado') else self.df_semanal, "resumen_semanal")).grid(row=0, column=8, padx=5, pady=5)
        
        # Crear treeview para tabla de resumen semanal
        frame_tabla = ttk.Frame(self.tab_semanal)
        frame_tabla.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.tree_semanal = ttk.Treeview(frame_tabla)
        self.tree_semanal.pack(side="left", fill="both", expand=True)
        
        # Configurar scrollbar vertical
        scrolly = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tree_semanal.yview)
        self.tree_semanal.configure(yscrollcommand=scrolly.set)
        scrolly.pack(side="right", fill="y")
        
        # Configurar scrollbar horizontal
        frame_scrollx = ttk.Frame(self.tab_semanal)
        frame_scrollx.pack(fill="x", padx=10, pady=(0, 5))
        
        scrollx = ttk.Scrollbar(frame_scrollx, orient="horizontal", command=self.tree_semanal.xview)
        self.tree_semanal.configure(xscrollcommand=scrollx.set)
        scrollx.pack(fill="x")

    def cargar_datos_existentes(self):
        """Cargar datos de archivos existentes si están disponibles"""
        try:
            # Cargar datos de asistencia diaria
            if os.path.exists(RUTA_CSV_FUSIONADO):
                self.df_asistencia = pd.read_csv(RUTA_CSV_FUSIONADO, encoding='utf-8')
                
                # Normalizar los nombres para unificar registros de Kevin
                self.df_asistencia['NOMBRE'] = self.df_asistencia['NOMBRE'].apply(normalizar_nombre)
                
                # Estandarizar nombres basados en ID
                self.df_asistencia['NOMBRE'] = self.df_asistencia['ID EMPLEADO'].apply(lambda x: ID_A_NOMBRE.get(x, x))
                
                # Crear columnas de año, mes y día para filtrado
                if 'FECHA' in self.df_asistencia.columns:
                    self.df_asistencia['FECHA_DT'] = pd.to_datetime(self.df_asistencia['FECHA'], format='%d/%m/%Y', errors='coerce')
                    self.df_asistencia['AÑO'] = self.df_asistencia['FECHA_DT'].dt.year
                    self.df_asistencia['MES'] = self.df_asistencia['FECHA_DT'].dt.month
                    self.df_asistencia['DIA'] = self.df_asistencia['FECHA_DT'].dt.day
                
                self.actualizar_tabla_asistencia(self.df_asistencia)
                
                # Actualizar combobox de filtros diarios
                empleados = sorted(self.df_asistencia['NOMBRE'].unique().tolist())
                self.combo_empleado['values'] = empleados
                
                # Actualizar combobox de año, mes, día y estatus para filtros diarios
                if 'AÑO' in self.df_asistencia.columns:
                    anios = sorted(self.df_asistencia['AÑO'].dropna().unique().tolist())
                    self.combo_anio_diario['values'] = anios
                
                if 'MES' in self.df_asistencia.columns:
                    meses = sorted(self.df_asistencia['MES'].dropna().unique().tolist())
                    self.combo_mes['values'] = meses
                
                if 'DIA' in self.df_asistencia.columns:
                    dias = sorted(self.df_asistencia['DIA'].dropna().unique().tolist())
                    self.combo_dia['values'] = dias
                
                if 'ESTATUS' in self.df_asistencia.columns:
                    estatus = sorted(self.df_asistencia['ESTATUS'].dropna().unique().tolist())
                    self.combo_estatus['values'] = estatus
            
            # Cargar datos de resumen semanal
            if os.path.exists(RUTA_CSV_SEMANAL):
                self.df_semanal = pd.read_csv(RUTA_CSV_SEMANAL, encoding='utf-8')
                
                # Normalizar los nombres para unificar registros de Kevin
                if 'NOMBRE' in self.df_semanal.columns:
                    self.df_semanal['NOMBRE'] = self.df_semanal['NOMBRE'].apply(normalizar_nombre)
                    # Estandarizar nombres basados en ID
                    self.df_semanal['NOMBRE'] = self.df_semanal['ID EMPLEADO'].apply(lambda x: ID_A_NOMBRE.get(x, x))
                
                self.actualizar_tabla_semanal(self.df_semanal)
                
                # Actualizar combobox de filtros semanales
                if 'NOMBRE' in self.df_semanal.columns:
                    empleados_semanal = sorted(self.df_semanal['NOMBRE'].unique().tolist())
                    self.combo_empleado_semanal['values'] = empleados_semanal
                    
                if 'AÑO' in self.df_semanal.columns:
                    anios_semanal = sorted(self.df_semanal['AÑO'].dropna().unique().tolist())
                    self.combo_anio_semanal['values'] = anios_semanal
                    
                if 'SEMANA' in self.df_semanal.columns:
                    semanas = sorted(self.df_semanal['SEMANA'].dropna().unique().tolist())
                    self.combo_semana['values'] = semanas
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar datos: {str(e)}")

    def actualizar_tabla_asistencia(self, df):
        """Actualizar la tabla de asistencia diaria con los datos del DataFrame"""
        # Limpiar tabla actual
        for item in self.tree_asistencia.get_children():
            self.tree_asistencia.delete(item)
        
        if df is None or df.empty:
            return
        
        # Configurar columnas
        columnas = df.columns.tolist()
        # Eliminar columnas auxiliares para la visualización
        if 'FECHA_DT' in columnas:
            columnas.remove('FECHA_DT')
        
        self.tree_asistencia['columns'] = columnas
        self.tree_asistencia['show'] = 'headings'  # Ocultar columna de ID
        
        # Configurar encabezados
        for col in columnas:
            self.tree_asistencia.heading(col, text=col)
            width = 100 if col in ['FECHA', 'ENTRADA', 'SALIDA'] else 80
            self.tree_asistencia.column(col, width=width)
        
        # Agregar filas
        for _, row in df.iterrows():
            valores = [row[col] for col in columnas]
            self.tree_asistencia.insert('', 'end', values=valores)

    def actualizar_tabla_semanal(self, df):
        """Actualizar la tabla de resumen semanal con los datos del DataFrame"""
        # Limpiar tabla actual
        for item in self.tree_semanal.get_children():
            self.tree_semanal.delete(item)
        
        if df is None or df.empty:
            return
        
        # Configurar columnas
        columnas = df.columns.tolist()
        self.tree_semanal['columns'] = columnas
        self.tree_semanal['show'] = 'headings'  # Ocultar columna de ID
        
        # Configurar encabezados
        for col in columnas:
            self.tree_semanal.heading(col, text=col)
            width = 100 if col in ['INICIO', 'FIN'] else 80
            self.tree_semanal.column(col, width=width)
        
        # Agregar filas
        for _, row in df.iterrows():
            valores = [row[col] for col in columnas]
            self.tree_semanal.insert('', 'end', values=valores)

    def procesar_asistencia(self):
        """Ejecutar el proceso de asistencia en un hilo separado"""
        def ejecutar():
            try:
                procesar_asistencia_original()
                # Actualizar la interfaz después del proceso
                self.root.after(0, self.cargar_datos_existentes)
                self.root.after(0, lambda: messagebox.showinfo("Éxito", "Proceso completado con éxito"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", f"Error en el proceso: {str(e)}"))
        
        # Mostrar mensaje de proceso
        messagebox.showinfo("Procesando", "Iniciando procesamiento de asistencia...\nEsto puede tomar unos momentos.")
        
        # Ejecutar en un hilo separado para no bloquear la interfaz
        threading.Thread(target=ejecutar, daemon=True).start()

    def aplicar_filtro_asistencia(self):
        """Aplicar todos los filtros seleccionados a la tabla de asistencia"""
        if self.df_asistencia is None or self.df_asistencia.empty:
            return
        
        # Obtener valores de los filtros
        empleado = self.combo_empleado.get()
        anio = self.combo_anio_diario.get()
        mes = self.combo_mes.get()
        dia = self.combo_dia.get()
        estatus = self.combo_estatus.get()
        
        # Aplicar filtros
        df_filtrado = self.df_asistencia.copy()
        
        if empleado:
            df_filtrado = df_filtrado[df_filtrado['NOMBRE'] == empleado]
        
        if anio:
            df_filtrado = df_filtrado[df_filtrado['AÑO'] == int(anio)]
        
        if mes:
            df_filtrado = df_filtrado[df_filtrado['MES'] == int(mes)]
        
        if dia:
            df_filtrado = df_filtrado[df_filtrado['DIA'] == int(dia)]
        
        if estatus:
            df_filtrado = df_filtrado[df_filtrado['ESTATUS'] == estatus]
        
        # Guardar el DataFrame filtrado y actualizar la tabla
        self.df_asistencia_filtrado = df_filtrado
        self.actualizar_tabla_asistencia(df_filtrado)
        
        # Mostrar cantidad de registros encontrados
        self.mostrar_contador_registros(df_filtrado, "asistencia")

    def limpiar_filtro_asistencia(self):
        """Limpiar todos los filtros de asistencia diaria"""
        if self.df_asistencia is not None:
            # Limpiar comboboxes
            self.combo_empleado.set('')
            self.combo_anio_diario.set('')
            self.combo_mes.set('')
            self.combo_dia.set('')
            self.combo_estatus.set('')
            
            # Actualizar tabla con todos los datos
            self.actualizar_tabla_asistencia(self.df_asistencia)
            
            # Eliminar el DataFrame filtrado
            if hasattr(self, 'df_asistencia_filtrado'):
                delattr(self, 'df_asistencia_filtrado')
            
            # Mostrar cantidad de registros
            self.mostrar_contador_registros(self.df_asistencia, "asistencia")

    def aplicar_filtro_semanal(self):
        """Aplicar filtros a la tabla de resumen semanal"""
        if self.df_semanal is None or self.df_semanal.empty:
            return
            
        # Obtener valores de los filtros
        empleado = self.combo_empleado_semanal.get()
        anio = self.combo_anio_semanal.get()
        semana = self.combo_semana.get()
        
        # Aplicar filtros
        df_filtrado = self.df_semanal.copy()
        
        if empleado:
            df_filtrado = df_filtrado[df_filtrado['NOMBRE'] == empleado]
        
        if anio:
            df_filtrado = df_filtrado[df_filtrado['AÑO'] == int(anio)]
            
        if semana:
            df_filtrado = df_filtrado[df_filtrado['SEMANA'] == int(semana)]
        
        # Guardar el DataFrame filtrado y actualizar la tabla
        self.df_semanal_filtrado = df_filtrado
        self.actualizar_tabla_semanal(df_filtrado)
        
        # Mostrar cantidad de registros encontrados
        self.mostrar_contador_registros(df_filtrado, "semanal")

    def limpiar_filtro_semanal(self):
        """Limpiar todos los filtros del resumen semanal"""
        if self.df_semanal is not None:
            # Limpiar comboboxes
            self.combo_empleado_semanal.set('')
            self.combo_anio_semanal.set('')
            self.combo_semana.set('')
            
            # Actualizar tabla con todos los datos
            self.actualizar_tabla_semanal(self.df_semanal)
            
            # Eliminar el DataFrame filtrado
            if hasattr(self, 'df_semanal_filtrado'):
                delattr(self, 'df_semanal_filtrado')
            
            # Mostrar cantidad de registros
            self.mostrar_contador_registros(self.df_semanal, "semanal")

    def mostrar_contador_registros(self, df, tipo):
        """Muestra un mensaje con la cantidad de registros encontrados"""
        if df is None or df.empty:
            mensaje = f"No hay registros que coincidan con los filtros aplicados."
        else:
            mensaje = f"Se encontraron {len(df)} registros."
        
        if tipo == "asistencia":
            tab_index = 1  # Índice de la pestaña de asistencia
        else:
            tab_index = 2  # Índice de la pestaña de resumen semanal
        
        # Mostrar mensaje en barra de estado o en una etiqueta
        print(mensaje)  # Por ahora solo imprimimos en consola
        messagebox.showinfo("Registros encontrados", mensaje)

    def exportar_tabla(self, df, nombre_base):
        """Exportar datos de una tabla a Excel"""
        if df is None or df.empty:
            messagebox.showinfo("Información", "No hay datos para exportar")
            return
            
        try:
            # Solicitar ubicación para guardar
            fecha_actual = datetime.now().strftime('%Y%m%d_%H%M')
            nombre_archivo = f"{nombre_base}_{fecha_actual}.xlsx"
            ruta_archivo = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=nombre_archivo
            )
            
            if not ruta_archivo:
                return
            
            # Si hay una columna FECHA_DT auxiliar, eliminarla antes de exportar
            df_export = df.copy()
            if 'FECHA_DT' in df_export.columns:
                df_export = df_export.drop(columns=['FECHA_DT'])
                
            # Exportar a Excel
            df_export.to_excel(ruta_archivo, index=False)
            messagebox.showinfo("Éxito", f"Datos exportados a {ruta_archivo}")
            
            # Abrir el archivo exportado
            abrir_archivo(ruta_archivo)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar: {str(e)}")

# ------------- Punto de entrada principal ------------------

if __name__ == "__main__":
    # Verificar si se quiere usar la interfaz gráfica o el modo consola
    if len(sys.argv) > 1 and sys.argv[1].lower() == "consola":
        # Modo consola
        procesar_asistencia_original()
    else:
        # Modo interfaz gráfica
        root = tk.Tk()
        app = AsistenciaApp(root)
        root.mainloop()