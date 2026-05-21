# Validador BOM AlfaPeople - Business Central

Script en Python para validar archivos Excel de carga relacionados con productos, listas de materiales de producción y rutas de fabricación en un contexto de implementación ERP con Microsoft Dynamics 365 Business Central.

## Objetivo

Detectar errores de estructura, datos incompletos, inconsistencias, referencias cruzadas inválidas y riesgos de carga antes de utilizar el archivo Excel en un mock, demo o proceso de migración/carga de datos.

## Archivo de entrada esperado

El script espera encontrar en la misma carpeta el archivo:

alfapeople_fiel_v2.xlsx

Este nombre está definido en el script en la variable:

ARCHIVO_ENTRADA = "alfapeople_fiel_v2.xlsx"

## Hojas requeridas

El Excel debe contener las siguientes hojas:

- Producto
- Cab. L.M. producción
- Línea L.M. producción
- Cab. ruta
- Línea ruta

## Validaciones incluidas

### Producto

- Campos obligatorios vacíos
- Producto fabricado sin BOM
- Producto fabricado sin ruta
- Costo unitario manual en producto fabricado
- Producto duplicado

### Cabecera de L.M. producción

- Campos obligatorios vacíos
- BOM duplicado
- BOM sin líneas asociadas

### Línea de L.M. producción

- Campos obligatorios vacíos
- BOM inexistente
- Cantidad no numérica
- Cantidad menor o igual a cero
- Componente no encontrado en Producto

### Cabecera de ruta

- Campos obligatorios vacíos
- Ruta duplicada
- Ruta sin líneas asociadas

### Línea de ruta

- Campos obligatorios vacíos
- Ruta inexistente
- Tiempos no numéricos
- Tiempos negativos
- Tamaño de lote inválido

### Validaciones cruzadas

- Circularidad en BOM
- Conexión de BOM sin operación en ruta
- Operación de ruta sin consumo asociado
- Producto que referencia BOM inexistente
- Producto que referencia ruta inexistente

## Resultado generado

El script genera un archivo Excel con timestamp, por ejemplo:

resultado_validacion_alfapeople_20260521_1143.xlsx

También genera un archivo log, por ejemplo:

validacion_20260521_1143.log

## Hojas del archivo de salida

- KPIs
- ESTADO_CARGA
- HALLAZGOS
- RESUMEN_HALLAZGOS
- RESUMEN_POR_HOJA
- DATA_Producto
- DATA_Cab. L.M. producción
- DATA_Línea L.M. producción
- DATA_Cab. ruta
- DATA_Línea ruta

## Semáforo de carga

| Estado | Significado |
|---|---|
| BLOCKER | Tiene hallazgos ALTO. No apto para carga. |
| CON OBSERVACIÓN | Tiene hallazgos MEDIO o INFO. Requiere revisión. |
| APTO | Sin hallazgos relevantes. Apto para carga. |

## Requisitos

Instalar dependencias:

pip install pandas openpyxl

## Ejecución

Desde la carpeta del proyecto:

python3 validador_alfapeople_v4.py

## Recomendación para GitHub

No subir archivos reales del cliente ni resultados generados con información sensible.

Se recomienda excluir:

- Archivos .xlsx
- Archivos .log
- Carpeta .venv
- Carpetas __pycache__
- Archivos .pyc
- Archivos .DS_Store

## Autor

Fiore Mosqueira