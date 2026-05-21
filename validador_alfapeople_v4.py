import pandas as pd
import logging
from datetime import datetime

# =========================================================
# CONFIGURACIÓN
# =========================================================
ARCHIVO_ENTRADA = "alfapeople_fiel_v2.xlsx"
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
ARCHIVO_SALIDA = f"resultado_validacion_alfapeople_{timestamp}.xlsx"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(f"validacion_{timestamp}.log"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)

HOJAS_REQUERIDAS = {
    "Producto": [
        "Nº", "Descripción", "Unidad medida base", "Costo unitario",
        "Sistema", "Cód. cate", "Nº ruta", "Nº L.M. producción"
    ],
    "Cab. L.M. producción": [
        "Nº", "Descripción", "Cód. unidad medida"
    ],
    "Línea L.M. producción": [
        "Nº L.M. producción", "Cód. versión", "Nº línea", "Tipo", "Nº",
        "Descripción", "Cód. unidad medida", "Cantidad",
        "Cód. conexión ruta", "Cantidad por"
    ],
    "Cab. ruta": [
        "Nº", "Descripción"
    ],
    "Línea ruta": [
        "Nº ruta", "Cód. versión", "Nº operación", "Nº operación siguiente",
        "Nº operación anterior", "Tipo", "Nº", "N° centro trabajo",
        "Descripcion", "tiempo preparacion", "Tiempo ejecución",
        "Tiempo espera", "Tamaño lote", "% Factor rechazo",
        "Cód. unidad medida tiempo prep", "Cód. unidad medida tiempo ejec.",
        "Cód. unidad medida tiempo espera", "Capacidades concurrentes",
        "Cant. adelantamiento", "Cód. conexión ruta"
    ]
}

hallazgos = []


# =========================================================
# FUNCIONES UTILITARIAS
# =========================================================
def limpiar_texto(valor):
    if pd.isna(valor):
        return None
    texto = str(valor).strip()
    return texto if texto != "" else None


def normalizar_id(valor):
    """
    Normaliza IDs evitando falsos positivos por tipo de dato,
    pero sin destruir códigos con ceros a la izquierda.

    Ejemplos:
    - FINELI-0001 queda FINELI-0001
    - ME-25745 queda ME-25745
    - 1000.0 queda 1000
    - 00123 queda 00123
    """
    v = limpiar_texto(valor)

    if v is None:
        return None

    # Si tiene letras, guiones o ceros a la izquierda, se conserva como texto
    if not v.replace(".", "", 1).isdigit():
        return v

    if v.startswith("0") and len(v) > 1:
        return v

    try:
        numero = float(v)
        if numero.is_integer():
            return str(int(numero))
        return str(numero)
    except (ValueError, TypeError):
        return v


def convertir_numero(valor):
    return pd.to_numeric(valor, errors="coerce")


def agregar_hallazgo(nivel, hoja, tipo, detalle, fila_excel=None, codigo=None):
    hallazgos.append({
        "Nivel": nivel,
        "Hoja": hoja,
        "Tipo Hallazgo": tipo,
        "Detalle": detalle,
        "Fila Excel": fila_excel,
        "Código / Referencia": codigo
    })


def validar_obligatorio(row, hoja, campo, fila_excel):
    if campo in row and limpiar_texto(row[campo]) is None:
        agregar_hallazgo(
            "ALTO",
            hoja,
            f"{campo} vacío",
            f"El campo obligatorio '{campo}' está vacío.",
            fila_excel=fila_excel
        )


# =========================================================
# 1. CARGAR ARCHIVO
# =========================================================
try:
    xls = pd.ExcelFile(ARCHIVO_ENTRADA)
    log.info("Archivo cargado correctamente")
    log.info(f"Hojas detectadas: {xls.sheet_names}")
except Exception as e:
    log.error(f"Error al cargar archivo: {e}")
    raise


# =========================================================
# 2. VALIDAR EXISTENCIA DE HOJAS Y COLUMNAS
# =========================================================
data = {}

for hoja, columnas_requeridas in HOJAS_REQUERIDAS.items():

    if hoja not in xls.sheet_names:
        agregar_hallazgo(
            "ALTO",
            hoja,
            "Hoja faltante",
            f"No existe la hoja requerida: {hoja}"
        )
        continue

    df = pd.read_excel(ARCHIVO_ENTRADA, sheet_name=hoja)
    df.columns = [str(c).strip() for c in df.columns]

    faltantes = [c for c in columnas_requeridas if c not in df.columns]

    if faltantes:
        agregar_hallazgo(
            "ALTO",
            hoja,
            "Columnas faltantes",
            f"Faltan columnas obligatorias: {faltantes}"
        )
        log.warning(f"[{hoja}] Columnas faltantes: {faltantes}")
        log.warning(f"[{hoja}] Columnas detectadas: {df.columns.tolist()}")

    df = df.reset_index(drop=True)
    df["Fila Excel"] = df.index + 2
    data[hoja] = df


df_producto   = data.get("Producto", pd.DataFrame())
df_cab_bom    = data.get("Cab. L.M. producción", pd.DataFrame())
df_linea_bom  = data.get("Línea L.M. producción", pd.DataFrame())
df_cab_ruta   = data.get("Cab. ruta", pd.DataFrame())
df_linea_ruta = data.get("Línea ruta", pd.DataFrame())


# =========================================================
# 3. NORMALIZAR CAMPOS TEXTO E IDs
# =========================================================
# Columnas que funcionan como IDs de referencia cruzada
COLUMNAS_ID = {
    "Producto":               ["Nº", "Nº ruta", "Nº L.M. producción"],
    "Cab. L.M. producción":   ["Nº"],
    "Línea L.M. producción":  ["Nº L.M. producción", "Nº"],
    "Cab. ruta":              ["Nº"],
    "Línea ruta":             ["Nº ruta", "Nº"],
}

for hoja, df in data.items():
    for col in df.columns:
        if col == "Fila Excel":
            continue
        if col in COLUMNAS_ID.get(hoja, []):
            df[col] = df[col].apply(normalizar_id)
        else:
            df[col] = df[col].apply(limpiar_texto)


# =========================================================
# 4. VALIDACIONES HOJA PRODUCTO
# =========================================================
if not df_producto.empty:
    for _, row in df_producto.iterrows():
        fila       = row["Fila Excel"]
        nro        = row.get("Nº")
        sistema    = row.get("Sistema")
        costo      = convertir_numero(row.get("Costo unitario"))
        bom_ref    = row.get("Nº L.M. producción")
        ruta_ref   = row.get("Nº ruta")

        for campo in ["Nº", "Descripción", "Unidad medida base", "Sistema"]:
            validar_obligatorio(row, "Producto", campo, fila)

        if sistema == "Ord. prod.":
            if limpiar_texto(bom_ref) is None:
                agregar_hallazgo(
                    "ALTO", "Producto",
                    "Producto fabricado sin BOM",
                    "El producto tiene Sistema = Ord. prod. pero no tiene Nº L.M. producción.",
                    fila_excel=fila, codigo=nro
                )

            if limpiar_texto(ruta_ref) is None:
                agregar_hallazgo(
                    "ALTO", "Producto",
                    "Producto fabricado sin ruta",
                    "El producto tiene Sistema = Ord. prod. pero no tiene Nº ruta.",
                    fila_excel=fila, codigo=nro
                )

            # Costo unitario no debe estar fijo en productos fabricados
            # BC lo calcula por BOM+Ruta; un valor manual puede generar desvíos
            if pd.notna(costo) and costo > 0:
                agregar_hallazgo(
                    "MEDIO", "Producto",
                    "Costo manual en producto fabricado",
                    f"El producto fabricado tiene Costo unitario = {costo}. "
                    "En BC el costo se calcula desde BOM y Ruta. Validar si es intencional.",
                    fila_excel=fila, codigo=nro
                )

    # Duplicados
    duplicados = df_producto[df_producto.duplicated(subset=["Nº"], keep=False)]
    for _, row in duplicados.iterrows():
        agregar_hallazgo(
            "ALTO", "Producto",
            "Producto duplicado",
            "El código de producto aparece más de una vez en la hoja Producto.",
            fila_excel=row["Fila Excel"], codigo=row.get("Nº")
        )


# =========================================================
# 5. VALIDACIONES CAB. L.M. PRODUCCIÓN
# =========================================================
if not df_cab_bom.empty:
    for _, row in df_cab_bom.iterrows():
        fila = row["Fila Excel"]
        for campo in ["Nº", "Descripción", "Cód. unidad medida"]:
            validar_obligatorio(row, "Cab. L.M. producción", campo, fila)

    duplicados = df_cab_bom[df_cab_bom.duplicated(subset=["Nº"], keep=False)]
    for _, row in duplicados.iterrows():
        agregar_hallazgo(
            "ALTO", "Cab. L.M. producción",
            "BOM duplicado",
            "El Nº de L.M. producción aparece duplicado.",
            fila_excel=row["Fila Excel"], codigo=row.get("Nº")
        )

    # Validar que cada cabecera BOM tenga al menos una línea asociada
    if not df_linea_bom.empty and "Nº L.M. producción" in df_linea_bom.columns:
        boms_con_lineas = set(df_linea_bom["Nº L.M. producción"].dropna().unique())

        for _, row in df_cab_bom.iterrows():
            bom_id = row.get("Nº")

            if limpiar_texto(bom_id) is not None and bom_id not in boms_con_lineas:
                agregar_hallazgo(
                    "ALTO",
                    "Cab. L.M. producción",
                    "BOM sin líneas",
                    "La cabecera de L.M. producción no tiene ninguna línea asociada en Línea L.M. producción.",
                    fila_excel=row["Fila Excel"],
                    codigo=bom_id
                )

# =========================================================
# 6. VALIDACIONES LÍNEA L.M. PRODUCCIÓN
# =========================================================
if not df_linea_bom.empty:
    cab_bom_ids   = set(df_cab_bom["Nº"].dropna().unique()) if not df_cab_bom.empty else set()
    productos_ids = set(df_producto["Nº"].dropna().unique()) if not df_producto.empty else set()

    for _, row in df_linea_bom.iterrows():
        fila       = row["Fila Excel"]
        bom_id     = row.get("Nº L.M. producción")
        componente = row.get("Nº")

        for campo in ["Nº L.M. producción", "Nº línea", "Tipo", "Nº",
                      "Descripción", "Cód. unidad medida", "Cantidad"]:
            validar_obligatorio(row, "Línea L.M. producción", campo, fila)

        if limpiar_texto(bom_id) is not None and bom_id not in cab_bom_ids:
            agregar_hallazgo(
                "ALTO", "Línea L.M. producción",
                "BOM inexistente",
                f"La línea referencia una L.M. producción que no existe: {bom_id}",
                fila_excel=fila, codigo=bom_id
            )

        cantidad = convertir_numero(row.get("Cantidad"))

        if pd.isna(cantidad):
            agregar_hallazgo(
                "ALTO", "Línea L.M. producción",
                "Cantidad no numérica",
                f"La cantidad no es válida: {row.get('Cantidad')}",
                fila_excel=fila, codigo=componente
            )
        elif cantidad <= 0:
            agregar_hallazgo(
                "ALTO", "Línea L.M. producción",
                "Cantidad no positiva",
                f"La cantidad debe ser mayor a cero. Valor encontrado: {cantidad}",
                fila_excel=fila, codigo=componente
            )

        if limpiar_texto(componente) is not None and componente not in productos_ids:
            agregar_hallazgo(
                "MEDIO", "Línea L.M. producción",
                "Componente no encontrado en Producto",
                f"El componente {componente} no existe en la hoja Producto. "
                "Validar si falta en el maestro.",
                fila_excel=fila, codigo=componente
            )


# =========================================================
# 7. VALIDACIONES CAB. RUTA
# =========================================================
if not df_cab_ruta.empty:
    for _, row in df_cab_ruta.iterrows():
        fila = row["Fila Excel"]
        for campo in ["Nº", "Descripción"]:
            validar_obligatorio(row, "Cab. ruta", campo, fila)

    duplicados = df_cab_ruta[df_cab_ruta.duplicated(subset=["Nº"], keep=False)]
    for _, row in duplicados.iterrows():
        agregar_hallazgo(
            "ALTO", "Cab. ruta",
            "Ruta duplicada",
            "El Nº de ruta aparece duplicado.",
            fila_excel=row["Fila Excel"], codigo=row.get("Nº")
        )

    # Validar que cada ruta tenga al menos una línea
    if not df_linea_ruta.empty:
        rutas_con_linea = set(df_linea_ruta["Nº ruta"].dropna().unique())
        for _, row in df_cab_ruta.iterrows():
            if row.get("Nº") not in rutas_con_linea:
                agregar_hallazgo(
                    "ALTO", "Cab. ruta",
                    "Ruta sin líneas",
                    "La ruta no tiene ninguna operación asociada en Línea ruta.",
                    fila_excel=row["Fila Excel"], codigo=row.get("Nº")
                )


# =========================================================
# 8. VALIDACIONES LÍNEA RUTA
# =========================================================
if not df_linea_ruta.empty:
    ruta_ids = set(df_cab_ruta["Nº"].dropna().unique()) if not df_cab_ruta.empty else set()

    for _, row in df_linea_ruta.iterrows():
        fila = row["Fila Excel"]
        ruta = row.get("Nº ruta")

        for campo in ["Nº ruta", "Nº operación", "Tipo", "Nº",
                      "N° centro trabajo", "Descripcion", "Cód. conexión ruta"]:
            validar_obligatorio(row, "Línea ruta", campo, fila)

        if limpiar_texto(ruta) is not None and ruta not in ruta_ids:
            agregar_hallazgo(
                "ALTO", "Línea ruta",
                "Ruta inexistente",
                f"La línea referencia una ruta que no existe en Cab. ruta: {ruta}",
                fila_excel=fila, codigo=ruta
            )

        # Tiempos: no numérico y negativo
        for campo_tiempo in ["tiempo preparacion", "Tiempo ejecución", "Tiempo espera"]:
            valor = convertir_numero(row.get(campo_tiempo))
            if pd.isna(valor):
                agregar_hallazgo(
                    "ALTO", "Línea ruta",
                    f"{campo_tiempo} no numérico",
                    f"El campo {campo_tiempo} no es numérico: {row.get(campo_tiempo)}",
                    fila_excel=fila, codigo=ruta
                )
            elif valor < 0:
                agregar_hallazgo(
                    "ALTO", "Línea ruta",
                    f"{campo_tiempo} negativo",
                    f"El campo {campo_tiempo} no puede ser negativo. Valor: {valor}",
                    fila_excel=fila, codigo=ruta
                )

        # Tamaño de lote: 0 genera división por cero en BC al calcular tiempos
        tamanio_lote = convertir_numero(row.get("Tamaño lote"))
        if pd.notna(tamanio_lote) and tamanio_lote <= 0:
            agregar_hallazgo(
                "ALTO", "Línea ruta",
                "Tamaño lote inválido",
                f"El Tamaño lote debe ser mayor a cero. Valor: {tamanio_lote}. "
                "Esto genera división por cero en BC al calcular tiempos de operación.",
                fila_excel=fila, codigo=ruta
            )


# =========================================================
# 9. DETECCIÓN DE CIRCULARIDAD EN BOM
# =========================================================
def detectar_circularidad(df_linea_bom, df_producto):
    """
    Detecta componentes que referencian su propio padre o cadenas circulares.
    En farma con semielaborados esto suele ocurrir.
    """
    if df_linea_bom.empty or df_producto.empty:
        return

    # Mapa: bom_id -> set de componentes
    grafo = {}
    for _, row in df_linea_bom.iterrows():
        bom    = row.get("Nº L.M. producción")
        comp   = row.get("Nº")
        if bom and comp:
            grafo.setdefault(bom, set()).add(comp)

    # Mapa: producto -> bom
    producto_a_bom = {}
    for _, row in df_producto.iterrows():
        nro = row.get("Nº")
        bom = row.get("Nº L.M. producción")
        if nro and bom:
            producto_a_bom[nro] = bom

    def tiene_ciclo(bom_inicio, visitados):
        componentes = grafo.get(bom_inicio, set())
        for comp in componentes:
            bom_comp = producto_a_bom.get(comp)
            if bom_comp:
                if bom_comp in visitados:
                    return True, comp
                visitados.add(bom_comp)
                resultado, nodo = tiene_ciclo(bom_comp, visitados)
                if resultado:
                    return True, nodo
        return False, None

    for bom_id in grafo:
        ciclo, nodo = tiene_ciclo(bom_id, {bom_id})
        if ciclo:
            agregar_hallazgo(
                "ALTO", "Circularidad BOM",
                "Circularidad detectada",
                f"La L.M. producción {bom_id} genera una referencia circular "
                f"a través del componente/BOM {nodo}.",
                codigo=bom_id
            )


detectar_circularidad(df_linea_bom, df_producto)


# =========================================================
# 10. VALIDACIONES CRUZADAS BOM ↔ RUTA
# =========================================================
if not df_linea_bom.empty and not df_linea_ruta.empty:
    conexiones_bom  = set(df_linea_bom["Cód. conexión ruta"].dropna().unique())
    conexiones_ruta = set(df_linea_ruta["Cód. conexión ruta"].dropna().unique())

    for conexion in conexiones_bom - conexiones_ruta:
        agregar_hallazgo(
            "ALTO", "Cruce BOM-Ruta",
            "Conexión de BOM sin operación",
            f"La conexión '{conexion}' existe en Línea L.M. producción "
            "pero no existe en Línea ruta.",
            codigo=conexion
        )

    for conexion in conexiones_ruta - conexiones_bom:
        agregar_hallazgo(
            "INFO", "Cruce BOM-Ruta",
            "Operación sin consumo asociado",
            f"La conexión '{conexion}' existe en Línea ruta pero no tiene "
            "componentes asociados en Línea L.M. producción.",
            codigo=conexion
        )


# =========================================================
# 11. VALIDACIONES CRUZADAS PRODUCTO ↔ BOM / RUTA
# =========================================================
if not df_producto.empty:
    cab_bom_ids = set(df_cab_bom["Nº"].dropna().unique()) if not df_cab_bom.empty else set()
    ruta_ids    = set(df_cab_ruta["Nº"].dropna().unique()) if not df_cab_ruta.empty else set()

    for _, row in df_producto.iterrows():
        fila     = row["Fila Excel"]
        producto = row.get("Nº")
        sistema  = row.get("Sistema")
        bom_ref  = row.get("Nº L.M. producción")
        ruta_ref = row.get("Nº ruta")

        if sistema == "Ord. prod.":
            if limpiar_texto(bom_ref) is not None and bom_ref not in cab_bom_ids:
                agregar_hallazgo(
                    "ALTO", "Cruce Producto-BOM",
                    "Producto referencia BOM inexistente",
                    f"El producto referencia una L.M. producción que no existe: {bom_ref}",
                    fila_excel=fila, codigo=producto
                )

            if limpiar_texto(ruta_ref) is not None and ruta_ref not in ruta_ids:
                agregar_hallazgo(
                    "ALTO", "Cruce Producto-Ruta",
                    "Producto referencia ruta inexistente",
                    f"El producto referencia una ruta que no existe: {ruta_ref}",
                    fila_excel=fila, codigo=producto
                )


# =========================================================
# 12. RESÚMENES, SEMÁFORO Y EXPORTACIÓN
# =========================================================
df_hallazgos = pd.DataFrame(hallazgos)

if df_hallazgos.empty:
    df_hallazgos = pd.DataFrame(columns=[
        "Nivel", "Hoja", "Tipo Hallazgo", "Detalle", "Fila Excel", "Código / Referencia"
    ])

resumen_hallazgos = (
    df_hallazgos.groupby(["Nivel", "Hoja", "Tipo Hallazgo"], dropna=False)
    .size()
    .reset_index(name="Cantidad")
    .sort_values(by=["Nivel", "Hoja", "Cantidad"], ascending=[True, True, False])
)

resumen_por_hoja = (
    df_hallazgos.groupby(["Hoja", "Nivel"], dropna=False)
    .size()
    .reset_index(name="Cantidad")
    .sort_values(by=["Hoja", "Nivel"])
)

# Semáforo por hoja
# ROJO   = tiene hallazgos ALTO  → blocker para carga
# AMARILLO = solo MEDIO/INFO     → carga con observación
# VERDE  = sin hallazgos         → apto para carga
hojas_con_alto  = set(df_hallazgos[df_hallazgos["Nivel"] == "ALTO"]["Hoja"].unique())
hojas_con_medio = set(df_hallazgos[df_hallazgos["Nivel"] == "MEDIO"]["Hoja"].unique())
todas_las_hojas = list(HOJAS_REQUERIDAS.keys()) + ["Cruce BOM-Ruta", "Cruce Producto-BOM",
                                                    "Cruce Producto-Ruta", "Circularidad BOM"]

semaforo_rows = []
for hoja in todas_las_hojas:
    if hoja in hojas_con_alto:
        estado = "🔴 BLOCKER"
        descripcion = "Tiene hallazgos de nivel ALTO. No apto para carga."
    elif hoja in hojas_con_medio:
        estado = "🟡 CON OBSERVACIÓN"
        descripcion = "Solo hallazgos MEDIO/INFO. Carga posible con revisión previa."
    else:
        estado = "🟢 APTO"
        descripcion = "Sin hallazgos. Apto para carga."
    semaforo_rows.append({"Hoja": hoja, "Estado": estado, "Descripción": descripcion})

# KPIs generales
total_registros = sum(len(df) for df in data.values())
total_hallazgos = len(df_hallazgos)
total_alto      = len(df_hallazgos[df_hallazgos["Nivel"] == "ALTO"])
total_medio     = len(df_hallazgos[df_hallazgos["Nivel"] == "MEDIO"])
total_info      = len(df_hallazgos[df_hallazgos["Nivel"] == "INFO"])
pct_limpio      = round((1 - total_alto / max(total_registros, 1)) * 100, 1)

kpis = pd.DataFrame([
    {"Indicador": "Total registros analizados",   "Valor": total_registros},
    {"Indicador": "Total hallazgos",              "Valor": total_hallazgos},
    {"Indicador": "Hallazgos ALTO (blockers)",    "Valor": total_alto},
    {"Indicador": "Hallazgos MEDIO",              "Valor": total_medio},
    {"Indicador": "Hallazgos INFO",               "Valor": total_info},
    {"Indicador": "% registros sin blocker",      "Valor": f"{pct_limpio}%"},
    {"Indicador": "Archivo fuente",               "Valor": ARCHIVO_ENTRADA},
    {"Indicador": "Fecha validación",             "Valor": datetime.now().strftime("%Y-%m-%d %H:%M")},
])

df_semaforo = pd.DataFrame(semaforo_rows)

# Exportar todo
with pd.ExcelWriter(ARCHIVO_SALIDA, engine="openpyxl") as writer:
    # KPIs primero para que sea la primera hoja visible
    kpis.to_excel(writer, sheet_name="KPIs", index=False)
    df_semaforo.to_excel(writer, sheet_name="ESTADO_CARGA", index=False)
    df_hallazgos.to_excel(writer, sheet_name="HALLAZGOS", index=False)
    resumen_hallazgos.to_excel(writer, sheet_name="RESUMEN_HALLAZGOS", index=False)
    resumen_por_hoja.to_excel(writer, sheet_name="RESUMEN_POR_HOJA", index=False)

    for hoja, df in data.items():
        nombre_hoja = f"DATA_{hoja}"[:31]
        df.to_excel(writer, sheet_name=nombre_hoja, index=False)

log.info("\n=== RESUMEN DE VALIDACIÓN ===")
log.info(f"Total registros analizados : {total_registros}")
log.info(f"Hallazgos ALTO (blockers)  : {total_alto}")
log.info(f"Hallazgos MEDIO            : {total_medio}")
log.info(f"Hallazgos INFO             : {total_info}")
log.info(f"% registros sin blocker    : {pct_limpio}%")
log.info(f"Archivo generado           : {ARCHIVO_SALIDA}")

if not resumen_hallazgos.empty:
    log.info("\nDetalle por hoja y tipo:\n" + resumen_hallazgos.to_string(index=False))