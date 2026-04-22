# BOTS DE EXTRACCION

Proyecto para extraer:
- Fechas de ingreso por radicado
- Numero de documento por radicado

## 1) Requisitos para otro equipo

- Windows 10/11 (recomendado)
- Python 3.12+
- Git
- Acceso a la ruta de datos donde estan los radicados

Opcional:
- Poppler (solo si quieres render PDF con pdf2image; el bot FECHAS ya tiene fallback con pypdfium2)

## 2) Instalacion desde cero

1. Clonar repositorio

```powershell
git clone https://github.com/tiquesebastian/BOTS_DE_EXTRACION.git
cd BOTS_DE_EXTRACION
```

2. Crear y activar entorno virtual

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Si PowerShell bloquea scripts:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

3. Instalar dependencias

```powershell
pip install -r requirements.txt
```

## 3) Estructura principal

- FECHAS/fechas.py: extractor de fechas (CLI)
- FECHAS/interfaz_fechas.py: interfaz grafica retro para FECHAS
- FECHAS/abrir_interfaz_fechas.bat: lanzador rapido de la interfaz
- documentos/documentos.py: extractor mixto (TIF/PDF/imagenes)
- documentos/documentos_pdf.py: extractor solo PDF
- EXTRATOR/extractor_documentos.js: extractor alterno en JS

## 4) Uso rapido del bot FECHAS

### Opcion A: Interfaz grafica (recomendada)

Ejecuta:

```powershell
FECHAS\abrir_interfaz_fechas.bat
```

En la interfaz:
- Define ruta raiz de radicados
- Define archivo con lista de radicados
- Define salida CSV y salida Excel
- Ajusta limites de paginas/archivos si hace falta
- Define keywords de archivo (minimo 3) para priorizar nombres como otros,factura,resumen,tapa,urgencias
- Define keywords del dato para buscar la fecha cerca de palabras como ingreso,atencion,urgencias
- Clic en EJECUTAR

La interfaz tiene estilo retro terminal verde y muestra log en vivo.

Durante la ejecucion muestra tambien:
- Encontrados exactos
- No encontrados hasta el momento
- Restantes por OCR
- Tiempo transcurrido
- ETA estimada
- Lectura de velocidad: muy rapido, rapido, medio o lento

### Opcion B: Consola (CLI)

```powershell
python FECHAS\fechas.py \
	--root "Z:\IA 10\NUEVO\PARTE3\890701715" \
	--radicados "c:\ruta\numeros_pendientes.txt" \
	--out-csv "c:\ruta\890701715_fechas_ingreso.csv" \
	--out-excel "c:\ruta\890701715_fechas_ingreso_detalle.xlsx" \
	--max-pdf-text-pages 7 \
	--max-ocr-pages 7 \
	--max-files 12
```

## 5) Uso rapido del bot DOCUMENTOS

Mixto (TIF/PDF/imagenes):

```powershell
python documentos\documentos.py
```

Solo PDF:

```powershell
python documentos\documentos_pdf.py
```

## 6) Salidas esperadas

- CSV resumen con columnas principales
- Excel detalle con metodo, archivo, pagina, score y motivo

## 7) Recomendaciones operativas

- No correr 2 procesos del mismo bot sobre la misma salida al tiempo.
- Guardar numerospendientes por lote para trazabilidad.
- Si hay red inestable en unidad compartida, reintentar por lotes pequenos.

## 8) Problemas comunes

- Error de permisos al activar venv:
	- usar Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
- Error por ruta de red no accesible:
	- validar que la unidad este montada y con permisos
- OCR lento en lotes grandes:
	- bajar max paginas OCR o max archivos por radicado

## 9) Git y limpieza

El repositorio ignora salidas generadas y dependencias pesadas (csv/xlsx/log/zip y node_modules) para mantener commits limpios.
