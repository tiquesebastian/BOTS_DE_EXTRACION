from __future__ import annotations

import argparse
import csv
import re
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable

import numpy as np
from openpyxl import Workbook
from pdf2image import convert_from_path
from PIL import Image, ImageFilter, ImageOps
from pypdf import PdfReader
from rapidocr_onnxruntime import RapidOCR

try:
    import pypdfium2 as pdfium
except Exception:
    pdfium = None


ROOT_PATH = Path(r"Z:\IA 10\NUEVO\PARTE3\890701715")
TARGET_RADICADOS_PATH = Path(r"c:\Users\ticdesarrollo09\Music\trabajo\bot\faltan estos radicados")
OUTPUT_CSV_PATH = Path(r"c:\Users\ticdesarrollo09\Music\trabajo\bot\890701715_fechas_ingreso.csv")
OUTPUT_EXCEL_PATH = Path(r"c:\Users\ticdesarrollo09\Music\trabajo\bot\890701715_fechas_ingreso_detalle.xlsx")

SEED_RESULT_PATHS = [
    Path(r"c:\Users\ticdesarrollo09\Music\trabajo\bot\890701715_fechas_ingreso_js.txt"),
    Path(r"c:\Users\ticdesarrollo09\Music\trabajo\bot\890701715_fechas_ingreso.csv"),
]

MAX_PDF_TEXT_PAGES = 7
MAX_OCR_PAGES = 7
MAX_FILES_PER_RADICADO = 12
POPPLER_PATH = Path(r"C:\poppler-25.11.0\Library\bin")

FILE_NAME_PRIORITY = [
    "tapa",
    "factura",
    "hoja de atencion de urgencia",
    "hoja de atencion odontologica",
    "hoja de atencion",
    "otros",
    "resultado",
    "orden",
]

DEFAULT_FILE_KEYWORDS = ["otros", "factura", "resumen", "tapa", "urgencias"]
DEFAULT_LABEL_KEYWORDS = ["ingreso", "atencion"]

SUPPORTED_EXTENSIONS = {".pdf", ".tif", ".tiff", ".png", ".jpg", ".jpeg", ".bmp", ".gif"}
DATE_PATTERN = re.compile(r"(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{2,4})")
ATENCION_PATTERN = re.compile(r"atencion\s*:?\s*(\d{8})\d*", re.IGNORECASE)

INGRESO_PATTERNS = [
    re.compile(r"ingreso.{0,240}?fecha.{0,60}?(\d{1,2}\s*[/\-.]\s*\d{1,2}\s*[/\-.]\s*\d{2,4})", re.IGNORECASE),
    re.compile(r"ingreso.{0,240}?(\d{1,2}\s*[/\-.]\s*\d{1,2}\s*[/\-.]\s*\d{2,4})", re.IGNORECASE),
]


@dataclass
class FechasConfig:
    root_path: Path
    target_radicados_path: Path
    output_csv_path: Path
    output_excel_path: Path
    seed_result_paths: list[Path]
    max_pdf_text_pages: int = MAX_PDF_TEXT_PAGES
    max_ocr_pages: int = MAX_OCR_PAGES
    max_files_per_radicado: int = MAX_FILES_PER_RADICADO
    poppler_path: Path = POPPLER_PATH
    file_name_keywords: list[str] | None = None
    target_text_keywords: list[str] | None = None


@dataclass
class ExtractionResult:
    radicado: str
    ingreso: str
    score_confianza: int
    estado: str
    metodo: str
    archivo_origen: str
    pagina: int
    fecha_cruda: str
    motivo: str


@dataclass
class OCRMatch:
    fecha: str
    pagina: int
    fecha_cruda: str
    metodo: str
    motivo: str
    score_confianza: int


@dataclass
class ProgressInfo:
    total_objetivos: int
    total_pendientes_ocr: int
    procesados_ocr: int
    encontrados_exactos: int
    encontrados_seed: int
    no_encontrados: int
    restantes_ocr: int
    elapsed_seconds: float
    average_seconds_per_item: float
    eta_seconds: float
    speed_label: str
    current_radicado: str = ""


def normalizar_clave(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = texto.lower().replace("_", " ").replace("-", " ")
    return re.sub(r"\s+", " ", texto).strip()


def normalizar_radicado(texto: str) -> str:
    m = re.search(r"\d{6,15}", str(texto))
    return m.group(0) if m else str(texto).strip()


def normalizar_fecha(valor: str) -> str:
    m = DATE_PATTERN.search(valor)
    if not m:
        return ""

    dd_s, mm_s, yy_s = m.groups()
    dd = int(dd_s)
    mm = int(mm_s)
    yy = int(yy_s)

    if len(yy_s) == 2:
        yy += 2000 if yy <= 50 else 1900

    fecha = datetime(yy, mm, dd)
    return f"{fecha.day}/{fecha.month:02d}/{fecha.year % 100:02d}"


def parse_fecha(valor: str) -> datetime | None:
    try:
        m = DATE_PATTERN.search(valor or "")
        if not m:
            return None
        dd, mm, yy = m.groups()
        d = int(dd)
        mth = int(mm)
        y = int(yy)
        if len(yy) == 2:
            y += 2000 if y <= 50 else 1900
        return datetime(y, mth, d)
    except Exception:
        return None


def format_fecha(dt: datetime) -> str:
    return f"{dt.day}/{dt.month:02d}/{dt.year % 100:02d}"


def normalizar_lista_keywords(texto: str) -> list[str]:
    return [normalizar_clave(item) for item in (texto or "").split(",") if normalizar_clave(item)]


def extraer_fecha_desde_texto(texto: str, label_keywords: list[str]) -> tuple[str, str]:
    limpio = re.sub(r"\s+", " ", (texto or "").replace("\x00", " "))
    clave = normalizar_clave(limpio)
    if not clave:
        return "", ""

    keywords = [normalizar_clave(k) for k in (label_keywords or []) if normalizar_clave(k)]
    if not keywords:
        keywords = DEFAULT_LABEL_KEYWORDS

    for keyword in keywords:
        idx = clave.find(keyword)
        if idx == -1:
            continue
        segmento = clave[max(0, idx - 30) : idx + 260]
        m = DATE_PATTERN.search(segmento)
        if m:
            cruda = re.sub(r"\s+", "", m.group(0))
            try:
                return normalizar_fecha(cruda), cruda
            except ValueError:
                pass

    for patron in INGRESO_PATTERNS:
        m = patron.search(clave)
        if m:
            cruda = re.sub(r"\s+", "", m.group(1))
            try:
                return normalizar_fecha(cruda), cruda
            except ValueError:
                pass

    m = ATENCION_PATTERN.search(clave)
    if m:
        valor = m.group(1)
        try:
            fecha = datetime(int(valor[0:4]), int(valor[4:6]), int(valor[6:8]))
            return format_fecha(fecha), valor
        except ValueError:
            pass

    idx = clave.find("ingreso")
    if idx != -1:
        segmento = clave[idx : idx + 260]
        m = DATE_PATTERN.search(segmento)
        if m:
            cruda = re.sub(r"\s+", "", m.group(0))
            try:
                return normalizar_fecha(cruda), cruda
            except ValueError:
                pass

    return "", ""


def cargar_radicados_objetivo(path_archivo: Path) -> list[str]:
    radicados: list[str] = []
    with path_archivo.open("r", encoding="utf-8-sig", errors="ignore") as f:
        for linea in f:
            candidato = normalizar_radicado(linea.strip())
            if re.fullmatch(r"\d{6,15}", candidato):
                radicados.append(candidato)
    return list(dict.fromkeys(radicados))


def cargar_resultados_seed() -> dict[str, str]:
    seed: dict[str, str] = {}
    for ruta in SEED_RESULT_PATHS:
        if not ruta.exists():
            continue
        with ruta.open("r", encoding="utf-8-sig", errors="ignore", newline="") as f:
            reader = csv.DictReader(f)
            if reader.fieldnames and "radicado" in [h.strip().lower() for h in reader.fieldnames]:
                for row in reader:
                    radicado = normalizar_radicado((row.get("radicado") or row.get("Radicado") or "").strip())
                    ingreso = (row.get("ingreso") or row.get("Ingreso") or "").strip()
                    ingreso = normalizar_fecha(ingreso) if ingreso else ""
                    if radicado and ingreso:
                        seed[radicado] = ingreso
                continue

        with ruta.open("r", encoding="utf-8-sig", errors="ignore") as f:
            for linea in f:
                linea = linea.strip()
                if not linea or linea.lower().startswith("radicado"):
                    continue
                partes = linea.split(",", 1)
                if len(partes) != 2:
                    continue
                radicado = normalizar_radicado(partes[0])
                ingreso = normalizar_fecha(partes[1].strip())
                if radicado and ingreso:
                    seed[radicado] = ingreso
    return seed


def cargar_resultados_seed_desde_rutas(seed_paths: list[Path]) -> dict[str, str]:
    seed: dict[str, str] = {}
    for ruta in seed_paths:
        if not ruta.exists():
            continue
        with ruta.open("r", encoding="utf-8-sig", errors="ignore", newline="") as f:
            reader = csv.DictReader(f)
            if reader.fieldnames and "radicado" in [h.strip().lower() for h in reader.fieldnames]:
                for row in reader:
                    radicado = normalizar_radicado((row.get("radicado") or row.get("Radicado") or "").strip())
                    ingreso = (row.get("ingreso") or row.get("Ingreso") or "").strip()
                    ingreso = normalizar_fecha(ingreso) if ingreso else ""
                    if radicado and ingreso:
                        seed[radicado] = ingreso
                continue

        with ruta.open("r", encoding="utf-8-sig", errors="ignore") as f:
            for linea in f:
                linea = linea.strip()
                if not linea or linea.lower().startswith("radicado"):
                    continue
                partes = linea.split(",", 1)
                if len(partes) != 2:
                    continue
                radicado = normalizar_radicado(partes[0])
                ingreso = normalizar_fecha(partes[1].strip())
                if radicado and ingreso:
                    seed[radicado] = ingreso
    return seed


def construir_lista_archivos(carpeta: Path, max_files: int, file_keywords: list[str]) -> list[Path]:
    candidatos = [
        p for p in carpeta.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS
    ]

    def prioridad(archivo: Path) -> tuple[int, str]:
        nombre = normalizar_clave(archivo.name)
        for i, token in enumerate(file_keywords):
            if token in nombre:
                return i, nombre
        return len(file_keywords), nombre

    return sorted(candidatos, key=prioridad)[:max_files]


def preprocesar_imagen(imagen: Image.Image) -> np.ndarray:
    gris = ImageOps.grayscale(imagen)
    gris = ImageOps.autocontrast(gris)
    ancho, alto = gris.size
    if ancho < 1800:
        factor = max(1, int(1800 / max(1, ancho)))
        gris = gris.resize((ancho * factor, alto * factor), Image.Resampling.LANCZOS)
    gris = gris.filter(ImageFilter.SHARPEN)
    return np.array(gris)


def ocr_en_imagen(engine: RapidOCR, imagen: Image.Image) -> str:
    resultado, _ = engine(preprocesar_imagen(imagen))
    if not resultado:
        return ""
    return " ".join(item[1] for item in resultado)


def render_pdf_pages(archivo: Path, max_pages: int, poppler_path: Path) -> tuple[list[Image.Image], str]:
    try:
        paginas = convert_from_path(
            str(archivo),
            dpi=220,
            first_page=1,
            last_page=max_pages,
            poppler_path=str(poppler_path) if poppler_path.exists() else None,
        )
        return paginas, "pdf2image"
    except Exception:
        pass

    if pdfium is None:
        return [], "sin_renderizador"

    try:
        doc = pdfium.PdfDocument(str(archivo))
        paginas: list[Image.Image] = []
        limite = min(len(doc), max_pages)
        for i in range(limite):
            page = doc[i]
            bitmap = page.render(scale=220 / 72)
            pil_img = bitmap.to_pil()
            paginas.append(pil_img)
        return paginas, "pypdfium2"
    except Exception:
        return [], "sin_renderizador"


def extraer_fecha_pdf(archivo: Path, engine: RapidOCR, config: FechasConfig) -> OCRMatch | None:
    try:
        reader = PdfReader(str(archivo), strict=False)
        limite = min(len(reader.pages), config.max_pdf_text_pages)
        for i in range(limite):
            texto = reader.pages[i].extract_text() or ""
            fecha, cruda = extraer_fecha_desde_texto(texto, config.target_text_keywords or DEFAULT_LABEL_KEYWORDS)
            if fecha:
                return OCRMatch(fecha, i + 1, cruda, "pdf-texto", "Fecha detectada por texto PDF", 9)
    except Exception:
        pass

    try:
        paginas, motor_render = render_pdf_pages(archivo, config.max_ocr_pages, config.poppler_path)
    except Exception:
        paginas = []
        motor_render = "sin_renderizador"

    if not paginas:
        return OCRMatch("", 0, "", "pdf-error", f"No se pudo convertir PDF ({motor_render})", 1)

    for i, pagina in enumerate(paginas, start=1):
        texto = ocr_en_imagen(engine, pagina)
        fecha, cruda = extraer_fecha_desde_texto(texto, config.target_text_keywords or DEFAULT_LABEL_KEYWORDS)
        if fecha:
            return OCRMatch(fecha, i, cruda, "pdf-rapidocr", f"Fecha detectada por OCR PDF ({motor_render})", 8)

    return None


def extraer_fecha_tiff(archivo: Path, engine: RapidOCR, config: FechasConfig) -> OCRMatch | None:
    try:
        with Image.open(archivo) as img:
            total = int(getattr(img, "n_frames", 1))
            for i in range(min(total, config.max_ocr_pages)):
                img.seek(i)
                pagina = img.convert("RGB")
                texto = ocr_en_imagen(engine, pagina)
                fecha, cruda = extraer_fecha_desde_texto(texto, config.target_text_keywords or DEFAULT_LABEL_KEYWORDS)
                if fecha:
                    return OCRMatch(fecha, i + 1, cruda, "tiff-rapidocr", "Fecha detectada por OCR TIFF", 8)
    except Exception as error:
        return OCRMatch("", 0, "", "tiff-error", f"No se pudo procesar TIFF: {error}", 1)

    return None


def extraer_fecha_imagen(archivo: Path, engine: RapidOCR, config: FechasConfig) -> OCRMatch | None:
    try:
        with Image.open(archivo) as img:
            texto = ocr_en_imagen(engine, img.convert("RGB"))
        fecha, cruda = extraer_fecha_desde_texto(texto, config.target_text_keywords or DEFAULT_LABEL_KEYWORDS)
        if fecha:
            return OCRMatch(fecha, 1, cruda, "imagen-rapidocr", "Fecha detectada por OCR imagen", 7)
    except Exception as error:
        return OCRMatch("", 0, "", "imagen-error", f"No se pudo procesar imagen: {error}", 1)

    return None


def procesar_archivo(archivo: Path, engine: RapidOCR, config: FechasConfig) -> OCRMatch | None:
    ext = archivo.suffix.lower()
    if ext == ".pdf":
        return extraer_fecha_pdf(archivo, engine, config)
    if ext in {".tif", ".tiff"}:
        return extraer_fecha_tiff(archivo, engine, config)
    return extraer_fecha_imagen(archivo, engine, config)


def procesar_radicado(radicado: str, carpeta: Path, engine: RapidOCR, config: FechasConfig) -> ExtractionResult:
    archivos = construir_lista_archivos(
        carpeta,
        config.max_files_per_radicado,
        config.file_name_keywords or DEFAULT_FILE_KEYWORDS,
    )
    if not archivos:
        return ExtractionResult(radicado, "", 1, "sin_archivos", "", "", 0, "", "No hay archivos soportados")

    razones: list[str] = []
    for archivo in archivos:
        match = procesar_archivo(archivo, engine, config)
        if match and match.fecha:
            return ExtractionResult(
                radicado=radicado,
                ingreso=match.fecha,
                score_confianza=match.score_confianza,
                estado="ok",
                metodo=match.metodo,
                archivo_origen=str(archivo),
                pagina=match.pagina,
                fecha_cruda=match.fecha_cruda,
                motivo=match.motivo,
            )
        if match:
            razones.append(f"{archivo.name}: {match.metodo}")

    return ExtractionResult(
        radicado=radicado,
        ingreso="",
        score_confianza=1,
        estado="sin_fecha",
        metodo="",
        archivo_origen=str(archivos[0]),
        pagina=0,
        fecha_cruda="",
        motivo=" | ".join(razones[:4]),
    )


def score_aproximacion(distancia: int) -> int:
    if distancia <= 2:
        return 6
    if distancia <= 10:
        return 5
    if distancia <= 30:
        return 4
    return 3


def completar_fechas_faltantes(objetivos: list[str], mapa: dict[str, ExtractionResult]) -> None:
    conocidos: list[tuple[int, datetime, str]] = []
    for rad in objetivos:
        item = mapa.get(rad)
        if not item or not item.ingreso:
            continue
        if not rad.isdigit():
            continue
        dt = parse_fecha(item.ingreso)
        if dt:
            conocidos.append((int(rad), dt, item.ingreso))

    conocidos.sort(key=lambda x: x[0])

    fecha_fallback = conocidos[0][2] if conocidos else "1/01/24"

    for rad in objetivos:
        item = mapa.get(rad)
        if item and item.ingreso:
            continue

        if not rad.isdigit() or not conocidos:
            mapa[rad] = ExtractionResult(
                radicado=rad,
                ingreso=fecha_fallback,
                score_confianza=1,
                estado="aproximada",
                metodo="fallback-global",
                archivo_origen=item.archivo_origen if item else "",
                pagina=0,
                fecha_cruda=fecha_fallback,
                motivo="Fecha de respaldo global por falta de datos OCR",
            )
            continue

        rad_num = int(rad)
        prev_item = None
        next_item = None
        for val in conocidos:
            if val[0] <= rad_num:
                prev_item = val
            if val[0] > rad_num:
                next_item = val
                break

        if prev_item and next_item:
            dist_prev = abs(rad_num - prev_item[0])
            dist_next = abs(next_item[0] - rad_num)
            if dist_prev <= dist_next:
                fecha = format_fecha(prev_item[1])
                metodo = "aprox-vecino-anterior"
                dist = dist_prev
            else:
                fecha = format_fecha(next_item[1])
                metodo = "aprox-vecino-siguiente"
                dist = dist_next
        elif prev_item:
            fecha = format_fecha(prev_item[1])
            metodo = "aprox-solo-anterior"
            dist = abs(rad_num - prev_item[0])
        elif next_item:
            fecha = format_fecha(next_item[1])
            metodo = "aprox-solo-siguiente"
            dist = abs(next_item[0] - rad_num)
        else:
            fecha = fecha_fallback
            metodo = "fallback-global"
            dist = 9999

        mapa[rad] = ExtractionResult(
            radicado=rad,
            ingreso=fecha,
            score_confianza=score_aproximacion(dist),
            estado="aproximada",
            metodo=metodo,
            archivo_origen=item.archivo_origen if item else "",
            pagina=0,
            fecha_cruda=fecha,
            motivo=f"Fecha estimada por cercania de radicado (distancia={dist})",
        )


def escribir_csv(resultados: list[ExtractionResult], salida: Path) -> None:
    with salida.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["radicado", "ingreso", "score_confianza"])
        for item in resultados:
            writer.writerow([item.radicado, item.ingreso, item.score_confianza])


def escribir_excel_detalle(resultados: list[ExtractionResult], salida: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "resultados"
    ws.append([
        "radicado",
        "ingreso",
        "score_confianza",
        "estado",
        "metodo",
        "archivo_origen",
        "pagina",
        "fecha_cruda",
        "motivo",
    ])
    for item in resultados:
        ws.append([
            item.radicado,
            item.ingreso,
            item.score_confianza,
            item.estado,
            item.metodo,
            item.archivo_origen,
            item.pagina,
            item.fecha_cruda,
            item.motivo,
        ])
    wb.save(salida)


def format_duration(seconds: float) -> str:
    total = max(0, int(round(seconds)))
    minutos, segs = divmod(total, 60)
    horas, minutos = divmod(minutos, 60)
    if horas:
        return f"{horas:02d}:{minutos:02d}:{segs:02d}"
    return f"{minutos:02d}:{segs:02d}"


def classify_speed(avg_seconds_per_item: float) -> str:
    if avg_seconds_per_item <= 2.5:
        return "muy rapido"
    if avg_seconds_per_item <= 5:
        return "rapido"
    if avg_seconds_per_item <= 10:
        return "medio"
    return "lento"


def run_fechas(
    config: FechasConfig,
    on_log: Callable[[str], None] | None = None,
    on_progress: Callable[[ProgressInfo], None] | None = None,
) -> list[ExtractionResult]:
    log = on_log or print

    if not config.root_path.exists():
        raise FileNotFoundError(f"No existe la ruta raiz: {config.root_path}")
    if not config.target_radicados_path.exists():
        raise FileNotFoundError(f"No existe archivo de radicados: {config.target_radicados_path}")

    objetivos = cargar_radicados_objetivo(config.target_radicados_path)
    seed = cargar_resultados_seed_desde_rutas(config.seed_result_paths)
    engine = RapidOCR()

    pendientes_reales = [rad for rad in objetivos if not seed.get(rad)]
    encontrados_seed = len(objetivos) - len(pendientes_reales)
    start_time = time.perf_counter()

    log(f"Radicados objetivo: {len(objetivos)}")
    log(f"Ya resueltos por seed JS/CSV: {encontrados_seed}")
    log(f"Pendientes OCR Python: {len(pendientes_reales)}")

    def emitir_progreso(current_radicado: str = "") -> None:
        if not on_progress:
            return
        procesados = 0
        encontrados_ocr = 0
        no_encontrados = 0
        for rad in pendientes_reales:
            item = mapa.get(rad)
            if not item or item.estado == "pendiente":
                continue
            procesados += 1
            if item.estado == "ok":
                encontrados_ocr += 1
            else:
                no_encontrados += 1

        elapsed = time.perf_counter() - start_time
        avg = elapsed / procesados if procesados else 0.0
        restantes = max(0, len(pendientes_reales) - procesados)
        eta = avg * restantes if avg else 0.0
        on_progress(
            ProgressInfo(
                total_objetivos=len(objetivos),
                total_pendientes_ocr=len(pendientes_reales),
                procesados_ocr=procesados,
                encontrados_exactos=encontrados_seed + encontrados_ocr,
                encontrados_seed=encontrados_seed,
                no_encontrados=no_encontrados,
                restantes_ocr=restantes,
                elapsed_seconds=elapsed,
                average_seconds_per_item=avg,
                eta_seconds=eta,
                speed_label=classify_speed(avg),
                current_radicado=current_radicado,
            )
        )

    mapa: dict[str, ExtractionResult] = {}
    for radicado in objetivos:
        ingreso_seed = seed.get(radicado, "")
        if ingreso_seed:
            mapa[radicado] = ExtractionResult(
                radicado=radicado,
                ingreso=ingreso_seed,
                score_confianza=8,
                estado="seed",
                metodo="seed-js",
                archivo_origen="",
                pagina=0,
                fecha_cruda=ingreso_seed,
                motivo="Fecha tomada de resultados ya existentes",
            )

    for radicado in objetivos:
        mapa.setdefault(
            radicado,
            ExtractionResult(
                radicado=radicado,
                ingreso="",
                score_confianza=1,
                estado="pendiente",
                metodo="",
                archivo_origen="",
                pagina=0,
                fecha_cruda="",
                motivo="Pendiente",
            ),
        )

    completar_fechas_faltantes(objetivos, mapa)
    escribir_csv([mapa[rad] for rad in objetivos], config.output_csv_path)
    emitir_progreso()

    total = len(pendientes_reales)
    for idx, radicado in enumerate(pendientes_reales, start=1):
        carpeta = config.root_path / radicado
        log(f"[{idx}/{total}] {radicado}")

        if not carpeta.exists() or not carpeta.is_dir():
            mapa[radicado] = ExtractionResult(
                radicado,
                "",
                1,
                "sin_carpeta",
                "",
                "",
                0,
                "",
                "No existe carpeta del radicado",
            )
            emitir_progreso(radicado)
            continue

        mapa[radicado] = procesar_radicado(radicado, carpeta, engine, config)
        item = mapa[radicado]
        if item.estado == "ok":
            log(f"    -> encontrada: {item.ingreso} | score={item.score_confianza} | metodo={item.metodo}")
        else:
            log(f"    -> no encontrada | estado={item.estado} | motivo={item.motivo or 'sin coincidencia'}")

        emitir_progreso(radicado)

        if idx % 10 == 0 or idx == total:
            completar_fechas_faltantes(objetivos, mapa)
            parciales = [
                mapa.get(rad, ExtractionResult(rad, "1/01/24", 1, "aproximada", "fallback-global", "", 0, "1/01/24", "Fallback"))
                for rad in objetivos
            ]
            escribir_csv(parciales, config.output_csv_path)

    completar_fechas_faltantes(objetivos, mapa)

    resultados = [
        mapa.get(
            rad,
            ExtractionResult(rad, "1/01/24", 1, "aproximada", "fallback-global", "", 0, "1/01/24", "Fallback"),
        )
        for rad in objetivos
    ]

    escribir_csv(resultados, config.output_csv_path)
    escribir_excel_detalle(resultados, config.output_excel_path)

    exactas = sum(1 for r in resultados if r.estado in {"ok", "seed"})
    aproximadas = sum(1 for r in resultados if r.estado == "aproximada")
    promedio = round(sum(r.score_confianza for r in resultados) / max(1, len(resultados)), 2)
    elapsed_total = time.perf_counter() - start_time
    avg_total = elapsed_total / total if total else 0.0
    log(f"Tiempo total: {format_duration(elapsed_total)}")
    log(f"Promedio por radicado OCR: {avg_total:.2f}s ({classify_speed(avg_total)})")

    log(f"Total radicados: {len(resultados)}")
    log(f"Exactas (OCR/seed): {exactas}")
    log(f"Aproximadas: {aproximadas}")
    log(f"Score promedio: {promedio}")
    log(f"CSV generado: {config.output_csv_path}")
    log(f"Excel detalle: {config.output_excel_path}")
    emitir_progreso()
    return resultados


def construir_config_desde_args() -> FechasConfig:
    parser = argparse.ArgumentParser(description="Extractor de fechas de ingreso")
    parser.add_argument("--root", default=str(ROOT_PATH), help="Ruta raiz con carpetas por radicado")
    parser.add_argument("--radicados", default=str(TARGET_RADICADOS_PATH), help="Archivo con radicados objetivo")
    parser.add_argument("--out-csv", default=str(OUTPUT_CSV_PATH), help="Ruta salida CSV")
    parser.add_argument("--out-excel", default=str(OUTPUT_EXCEL_PATH), help="Ruta salida Excel detalle")
    parser.add_argument("--seed", action="append", default=None, help="Ruta adicional de resultados seed (puede repetirse)")
    parser.add_argument("--max-pdf-text-pages", type=int, default=MAX_PDF_TEXT_PAGES)
    parser.add_argument("--max-ocr-pages", type=int, default=MAX_OCR_PAGES)
    parser.add_argument("--max-files", type=int, default=MAX_FILES_PER_RADICADO)
    parser.add_argument(
        "--file-keywords",
        default=",".join(DEFAULT_FILE_KEYWORDS),
        help="Keywords de nombre de archivo separados por coma (minimo 3)",
    )
    parser.add_argument(
        "--label-keywords",
        default=",".join(DEFAULT_LABEL_KEYWORDS),
        help="Keywords de texto a buscar cerca de la fecha, separados por coma",
    )
    parser.add_argument("--poppler", default=str(POPPLER_PATH), help="Ruta de poppler (opcional)")
    args = parser.parse_args()

    seed_paths = [Path(p) for p in (args.seed or [])] or SEED_RESULT_PATHS
    file_keywords = normalizar_lista_keywords(args.file_keywords)
    if len(file_keywords) < 3:
        raise ValueError("Debes indicar minimo 3 keywords en --file-keywords")
    label_keywords = normalizar_lista_keywords(args.label_keywords) or DEFAULT_LABEL_KEYWORDS

    return FechasConfig(
        root_path=Path(args.root),
        target_radicados_path=Path(args.radicados),
        output_csv_path=Path(args.out_csv),
        output_excel_path=Path(args.out_excel),
        seed_result_paths=seed_paths,
        max_pdf_text_pages=max(1, args.max_pdf_text_pages),
        max_ocr_pages=max(1, args.max_ocr_pages),
        max_files_per_radicado=max(1, args.max_files),
        poppler_path=Path(args.poppler),
        file_name_keywords=file_keywords,
        target_text_keywords=label_keywords,
    )


def main() -> None:
    config = construir_config_desde_args()
    run_fechas(config)


if __name__ == "__main__":
    main()
