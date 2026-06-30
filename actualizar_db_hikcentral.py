"""
===========================================================
 ACTUALIZAR HIKCENTRAL CON RRHH + CLASIFICAR DEPARTAMENTOS
===========================================================

Fusiona en un solo paso lo que antes eran dos scripts:

  1) cruce_tablas_OK.py        -> agrega a HikCentral las
                                   personas nuevas que están
                                   en la nómina de RRHH
  2) fix_department_mejorado.py -> clasifica el Department de
                                   TODAS las filas (las que ya
                                   estaban y las nuevas) en uno
                                   de los 13 departamentos
                                   oficiales

Por qué en un solo paso: si se corren por separado, las filas
nuevas se agregan con un Department "crudo" (ej. "CURF/Medicina
Adulto") y recién quedan bien clasificadas si después te acordás
de correr el segundo script. Acá la clasificación se aplica
SIEMPRE, sobre el archivo final completo, en la misma corrida.

Clasificación (igual criterio que antes):
  1) Busca el SED_Sector en el diccionario de 138 sectores ->
     departamento oficial.
  2) Si no está, usa como respaldo el campo Department/departamento
     crudo, traducido (DEPT_FALLBACK).
  3) Si ninguna de las dos funciona, queda "CURF/Sin Clasificar"
     y se lista aparte para revisión manual.

Genera:
- HIKCENTRAL NOW ACTUALIZADO.xlsx
    hoja "Resultado"       -> todo el padrón ya cruzado y clasificado
    hoja "Sin Clasificar"  -> filas (nuevas o viejas) sin clasificar
- LOG_AGREGADOS.txt        -> personas agregadas desde RRHH

Uso:
python actualizar_hikcentral.py
===========================================================
"""

import pandas as pd
from datetime import datetime
import sys
import time

# ==========================================================
# CONFIGURACIÓN
# ==========================================================

ARCHIVO_HIK = "HIKCENTRAL NOW.xlsx"
ARCHIVO_RRHH = "NOMINA RRHH NOW.xlsx"
ARCHIVO_SALIDA = "HIKCENTRAL NOW ACTUALIZADO.xlsx"
ARCHIVO_LOG = "LOG_AGREGADOS.txt"

ID_HIK = "ID"
ID_RRHH = "documento"

COLUMNAS_RRHH_REQUERIDAS = [
    "documento", "nombres", "apellido", "sexo",
    "puesto", "departamento", "sector", "conductor1",
]

PERSON_TYPE_DEFAULT = "Normal Person"
SIN_CLASIFICAR = "Sin Clasificar"

DEPARTAMENTOS_OFICIALES = [
    "Administración", "Atención al Paciente", "Cirugía",
    "Diagnóstico y Laboratorio", "Dirección", "Enfermería", "Farmacia",
    "Investigación y Docencia", "Medicina", "Proveedores Externos",
    "Rehabilitación", "Seguridad", "Soporte Hospitalario",
]

# ==========================================================
# MAPEO SED_Sector -> Departamento oficial (138 sectores)
# ==========================================================

SECTOR_TO_DEPT = {
    "Ginecología y Obstetricia": "Medicina", "Guardia Pediátrica": "Medicina",
    "Cardiología": "Medicina", "Guardia Central": "Medicina", "Clinica Medica": "Medicina",
    "Neurología": "Medicina", "Neonatología": "Medicina",
    "Unidad de Terapia Intensiva": "Medicina", "Pediatría": "Medicina",
    "Dermatología": "Medicina", "Clinica Pediátrica": "Medicina",
    "UTI Pediátrica": "Medicina", "Salud Mental": "Medicina",
    "Diabetología  y Nutrición": "Medicina", "Gastroenterología": "Medicina",
    "Alergia e Inmunología": "Medicina", "Internado Pediatría": "Medicina",
    "Neurología Pediátrica": "Medicina", "Endocrinología": "Medicina",
    "Nutrición": "Medicina", "Infectología": "Medicina", "Oncología": "Medicina",
    "Coordinadores": "Medicina", "Salud Femenina y Perinatal": "Medicina",
    "Cardiología Pediátrica": "Medicina", "Internado Adultos": "Medicina",
    "Neumonología": "Medicina", "UCO": "Medicina",
    "Unidad de Medicina del Deporte": "Medicina", "Nefrología": "Medicina",
    "Reumatología": "Medicina", "Neumonología Pediátrica": "Medicina",
    "Nutrición Pediátrica": "Medicina", "Gastroenterología Pediátrica": "Medicina",
    "Hemodinamia": "Medicina", "Nefrología pediátrica": "Medicina",
    "Dermatología Pediátrica": "Medicina", "Endocrinología Pediátrica": "Medicina",
    "Nutrición Pediátrica Médica": "Medicina", "Infectología Pediátrica": "Medicina",
    "Enfermedades Metabólicas": "Medicina", "Reumatología Pediátrica": "Medicina",
    "Cuidados Paliativos Adultos": "Medicina", "Oncohematología Pediátrica": "Medicina",
    "Neurocirugía Pediátrica": "Medicina", "UCO y Cardiología Ambulatoria": "Medicina",
    "Diabetología Pediátrica": "Medicina", "Toxicología": "Medicina",
    "Cirugía Cardiovascular Pediátrica": "Medicina",
    "Cuidado Paliativo Pediátricos": "Medicina", "Genetica": "Medicina",
    "Hepatología": "Medicina",

    "Instrumentación Quirúrgica": "Cirugía", "Anestesia": "Cirugía",
    "Central de Esterilización": "Cirugía", "Oftalmología": "Cirugía",
    "Traumatología": "Cirugía", "ORL": "Cirugía", "Cirugía General": "Cirugía",
    "Urología": "Cirugía", "Odontología": "Cirugía",
    "Secretaría Quirurgica": "Cirugía", "Neurocirugía": "Cirugía",
    "Cirugía Plástica": "Cirugía", "Cirugía General Pediátrica": "Cirugía",
    "Cirugía Vascular Periférica": "Cirugía", "Cabeza y Cuello": "Cirugía",
    "Cirugía Cardiovascular Adultos": "Cirugía", "Cirugía de Tórax": "Cirugía",
    "Traumatología Pediátrica": "Cirugía", "Urología Pediátrica": "Cirugía",
    "Quirófano": "Cirugía", "Cirugía": "Cirugía",

    "Enfermería Área Crítica": "Enfermería", "Enfermería General": "Enfermería",
    "Enfermería Área Pediátrica": "Enfermería",
    "Enfermería Área Ambulatoria": "Enfermería",
    "Enfermería Área Gineco - Obstetricia": "Enfermería",
    "Enfermería Área Neonatología": "Enfermería", "Camilleros": "Enfermería",
    "Enfermería": "Enfermería",

    "Secretaría": "Atención al Paciente", "Centro de Contacto": "Atención al Paciente",
    "Guardia Administrativa": "Atención al Paciente", "Admisión": "Atención al Paciente",
    "Mega Salud": "Atención al Paciente",
    "Secretaría de Presupuestos": "Atención al Paciente",
    "Atención al Pte Internado y Guardia": "Atención al Paciente",

    "Diagnóstico por Imágenes": "Diagnóstico y Laboratorio",
    "Laboratorio Central": "Diagnóstico y Laboratorio",
    "Microbiología": "Diagnóstico y Laboratorio",
    "Hematología": "Diagnóstico y Laboratorio",
    "Anatomía Patológica": "Diagnóstico y Laboratorio",
    "Hemoterapia": "Diagnóstico y Laboratorio",
    "Biología Molecular": "Diagnóstico y Laboratorio",
    "Hemostasia": "Diagnóstico y Laboratorio",
    "Anatomía Patológica A": "Diagnóstico y Laboratorio",

    "Limpieza": "Soporte Hospitalario",
    "Limpro - Servicio tercerizado de limpieza": "Soporte Hospitalario",
    "Mantenimiento": "Soporte Hospitalario", "Sistemas": "Soporte Hospitalario",
    "Calidad y Operaciones": "Soporte Hospitalario", "Bioingeniería": "Soporte Hospitalario",
    "Archivo": "Soporte Hospitalario", "Obras": "Soporte Hospitalario",
    "Personal de Limpieza Limpro": "Soporte Hospitalario",

    "Facturación": "Administración", "RRHH": "Administración",
    "Comunicación e Imagen Institucional": "Administración", "Compras": "Administración",
    "Higiene y Seguridad": "Administración", "Finanzas": "Administración",
    "Convenios": "Administración", "Auditoría Medica": "Administración",
    "Contaduría": "Administración",
    "Comité Institucional de Ética en Invest. en Salud": "Administración",
    "Administración y Finanzas": "Administración",

    "Fisioterapia consultorios externos": "Rehabilitación",
    "Fonoaudiología": "Rehabilitación", "Fisioterapia Internados": "Rehabilitación",
    "Fisioterapia Internados y Consultorios externos": "Rehabilitación",
    "Audiología": "Rehabilitación", "Terapia Ocupacional": "Rehabilitación",
    "Psicopedagogía": "Rehabilitación", "Rehabilitación": "Rehabilitación",
    "Osteopatía": "Rehabilitación",

    "Bar": "Proveedores Externos",
    "Okinet - Servicio de impresoras externo": "Proveedores Externos",

    "Optima Seguridad": "Seguridad", "Policias": "Seguridad",

    "Farmacia": "Farmacia",

    "Comisión Directiva de Fund. para el Prog. de UCC": "Dirección",
    "Dirección": "Dirección", "Comité de Bioética": "Dirección",
    "Prestaciones Médicas": "Dirección", "Sin Datos": "Dirección",

    "Docencia e Investigación": "Investigación y Docencia",
    "Unidad de Investigación Clínica": "Investigación y Docencia",
}

DEPT_FALLBACK = {
    "Medicina": "Medicina", "Medicina Adulto": "Medicina", "Pediatría": "Medicina",
    "Salud Femenina y Perinatal": "Medicina", "Cirugía": "Cirugía",
    "Enfermería": "Enfermería", "Atención al Paciente": "Atención al Paciente",
    "Diagnóstico y Laboratorio": "Diagnóstico y Laboratorio",
    "Diagnostico por Imagenes": "Diagnóstico y Laboratorio",
    "Laboratorios": "Diagnóstico y Laboratorio",
    "Soporte Hospitalario": "Soporte Hospitalario", "Administración": "Administración",
    "RRHH": "Administración", "CIEIS": "Administración", "Rehabilitación": "Rehabilitación",
    "Proveedores Externos": "Proveedores Externos", "BAR": "Proveedores Externos",
    "Seguridad": "Seguridad", "Farmacia": "Farmacia", "Dirección": "Dirección",
    "Investigación y Docencia": "Investigación y Docencia",
    "Docencia e Investigación": "Investigación y Docencia",
}


def normalizar(serie):
    return (
        serie.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.upper()
    )


def limpiar_sector(valor):
    valor = str(valor).strip()
    if valor.upper() == "BAR":
        return "Bar"
    return valor


def limpiar_department_crudo(valor):
    """Quita CURF/ repetidos y se queda con el primer nivel."""
    valor = str(valor).strip()
    while valor.startswith("CURF/"):
        valor = valor[5:]
    if "/" in valor:
        valor = valor.split("/")[0]
    return valor.strip()


def clasificar_departamento(sector, department_crudo):
    sector = limpiar_sector(sector)
    if sector and sector in SECTOR_TO_DEPT:
        return SECTOR_TO_DEPT[sector]

    dept_crudo = limpiar_department_crudo(department_crudo)
    if dept_crudo in DEPT_FALLBACK:
        return DEPT_FALLBACK[dept_crudo]

    return SIN_CLASIFICAR


def main():
    inicio = time.time()

    print("=" * 60)
    print(" ACTUALIZAR HIKCENTRAL CON RRHH + CLASIFICAR DEPARTAMENTOS")
    print("=" * 60)

    # ------------------------------------------------------
    # FASE 1: leer archivos
    # ------------------------------------------------------
    print("\n[1/7] Leyendo archivos...")

    try:
        hik = pd.read_excel(ARCHIVO_HIK, dtype=str).fillna("")
    except FileNotFoundError:
        print(f"\n❌ No se encontró el archivo '{ARCHIVO_HIK}'.")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Error leyendo '{ARCHIVO_HIK}':\n{e}")
        sys.exit(1)

    try:
        rrhh = pd.read_excel(ARCHIVO_RRHH, dtype=str).fillna("")
    except FileNotFoundError:
        print(f"\n❌ No se encontró el archivo '{ARCHIVO_RRHH}'.")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Error leyendo '{ARCHIVO_RRHH}':\n{e}")
        sys.exit(1)

    print(f"✔ {ARCHIVO_HIK}: {len(hik)} registros")
    print(f"✔ {ARCHIVO_RRHH}: {len(rrhh)} registros")

    if ID_HIK not in hik.columns:
        print(f"\n❌ '{ARCHIVO_HIK}' no tiene la columna '{ID_HIK}'.")
        sys.exit(1)
    if "SED_Sector" not in hik.columns:
        print(f"\n❌ '{ARCHIVO_HIK}' no tiene la columna 'SED_Sector'.")
        sys.exit(1)
    if "Department" not in hik.columns:
        print(f"\n❌ '{ARCHIVO_HIK}' no tiene la columna 'Department'.")
        sys.exit(1)

    faltantes = [c for c in COLUMNAS_RRHH_REQUERIDAS if c not in rrhh.columns]
    if faltantes:
        print(f"\n❌ '{ARCHIVO_RRHH}' no tiene estas columnas requeridas: {faltantes}")
        sys.exit(1)

    # ------------------------------------------------------
    # FASE 2: normalizar IDs y deduplicar RRHH
    # ------------------------------------------------------
    print("\n[2/7] Normalizando IDs...")

    hik[ID_HIK] = normalizar(hik[ID_HIK])
    rrhh[ID_RRHH] = normalizar(rrhh[ID_RRHH])

    antes = len(rrhh)
    rrhh = rrhh.drop_duplicates(subset=[ID_RRHH], keep="last")
    duplicados = antes - len(rrhh)
    if duplicados > 0:
        print(f"⚠ {duplicados} documentos duplicados en RRHH. Se conservó la última fila de cada uno.")

    # ------------------------------------------------------
    # FASE 3: detectar personas nuevas
    # ------------------------------------------------------
    print("\n[3/7] Detectando personas nuevas...")

    ids_hik = set(hik[ID_HIK])
    nuevos = rrhh[
        (~rrhh[ID_RRHH].isin(ids_hik))
        & (rrhh[ID_RRHH] != "")
    ].copy()

    print(f"✔ Personas en HikCentral : {len(hik)}")
    print(f"✔ Personas en RRHH       : {len(rrhh)}")
    print(f"✔ Nuevos encontrados     : {len(nuevos)}")

    # ------------------------------------------------------
    # FASE 4: armar registros nuevos con formato HikCentral
    # ------------------------------------------------------
    print("\n[4/7] Armando registros nuevos...")

    hoy = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    registros = []

    for _, fila in nuevos.iterrows():
        registro = {c: "" for c in hik.columns}

        registro["ID"] = fila["documento"]
        registro["First Name"] = fila["nombres"]
        registro["Last Name"] = fila["apellido"]
        registro["Gender"] = fila["sexo"]
        registro["Position"] = fila["puesto"]

        if "Person Type" in hik.columns:
            registro["Person Type"] = PERSON_TYPE_DEFAULT
        if "Enrollment Date" in hik.columns:
            registro["Enrollment Date"] = hoy

        departamento = str(fila["departamento"]).strip()
        sector = str(fila["sector"]).strip()

        # Department "crudo": sirve de respaldo en la clasificación de la FASE 6
        if departamento == sector or sector == "":
            registro["Department"] = f"CURF/{departamento}"
        else:
            registro["Department"] = f"CURF/{departamento}/{sector}"

        if "SED_Sector" in hik.columns:
            registro["SED_Sector"] = sector
        if "SED_Conductor" in hik.columns:
            registro["SED_Conductor"] = fila["conductor1"]

        registros.append(registro)

    df_nuevos = pd.DataFrame(registros, columns=hik.columns)

    # ------------------------------------------------------
    # FASE 5: unir tablas
    # ------------------------------------------------------
    print("\n[5/7] Uniendo tablas...")

    resultado = pd.concat([hik, df_nuevos], ignore_index=True)

    # ------------------------------------------------------
    # FASE 6: clasificar Department en TODO el padrón
    # (las que ya estaban + las recién agregadas)
    # ------------------------------------------------------
    print("\n[6/7] Clasificando Department en los 13 departamentos oficiales...")

    resultado["Department"] = resultado.apply(
        lambda row: f"CURF/{clasificar_departamento(row['SED_Sector'], row['Department'])}",
        axis=1,
    )

    sin_clasificar = resultado[resultado["Department"] == f"CURF/{SIN_CLASIFICAR}"]

    conteo = resultado["Department"].value_counts().sort_index()
    print("-" * 60)
    print("Departamentos encontrados")
    print("-" * 60)
    for departamento, cantidad in conteo.items():
        print(f"{departamento:<40} {cantidad:>5}")
    print("-" * 60)
    print(f"Total departamentos distintos: {len(conteo)}")

    if len(sin_clasificar) > 0:
        print(f"\n⚠ {len(sin_clasificar)} filas sin clasificar (revisar manualmente):")
        cols_mostrar = [c for c in ["ID", "First Name", "Last Name", "SED_Sector"] if c in resultado.columns]
        print(sin_clasificar[cols_mostrar].to_string(index=False))
    else:
        print("\n✔ Sin filas pendientes de revisión.")

    # ------------------------------------------------------
    # FASE 7: guardar Excel y LOG
    # ------------------------------------------------------
    print("\n[7/7] Guardando archivos...")

    try:
        with pd.ExcelWriter(ARCHIVO_SALIDA, engine="openpyxl") as writer:
            resultado.to_excel(writer, sheet_name="Resultado", index=False)
            if len(sin_clasificar) > 0:
                sin_clasificar.to_excel(writer, sheet_name="Sin Clasificar", index=False)
    except Exception as e:
        print(f"\n❌ Error guardando '{ARCHIVO_SALIDA}':\n{e}")
        sys.exit(1)

    with open(ARCHIVO_LOG, "w", encoding="utf-8") as log:
        log.write("PERSONAS AGREGADAS\n")
        log.write("===============================\n\n")
        for _, fila in nuevos.iterrows():
            log.write(f'{fila["documento"]} - {fila["apellido"]}, {fila["nombres"]}\n')

        if len(sin_clasificar) > 0:
            log.write("\n\nPENDIENTES DE CLASIFICAR (revisar Department a mano)\n")
            log.write("===============================\n\n")
            for _, fila in sin_clasificar.iterrows():
                log.write(f'{fila["ID"]} - {fila.get("Last Name","")}, {fila.get("First Name","")}\n')

    fin = time.time()

    print("\n" + "=" * 60)
    print("✅ PROCESO FINALIZADO CORRECTAMENTE")
    print("=" * 60)
    print(f"Registros originales      : {len(hik)}")
    print(f"Registros agregados       : {len(df_nuevos)}")
    print(f"Total final               : {len(resultado)}")
    print(f"Departamentos únicos      : {len(conteo)}")
    print(f"Sin clasificar (revisar)  : {len(sin_clasificar)}")
    print(f"Tiempo de ejecución       : {fin - inicio:.2f} segundos")
    print("")
    print(f"Excel generado : {ARCHIVO_SALIDA}")
    print(f"Log generado   : {ARCHIVO_LOG}")
    print("=" * 60)


if __name__ == "__main__":
    main()
