from __future__ import annotations

import csv
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import numpy as np
from openpyxl import Workbook
from pdf2image import convert_from_path
from PIL import Image, ImageFilter, ImageOps
from pypdf import PdfReader
from rapidocr_onnxruntime import RapidOCR


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

SUPPORTED_EXTENSIONS = {".pdf", ".tif", ".tiff", ".png", ".jpg", ".jpeg", ".bmp", ".gif"}
DATE_PATTERN = re.compile(r"(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{2,4})")
ATENCION_PATTERN = re.compile(r"atencion\s*:?\s*(\d{8})\d*", re.IGNORECASE)

INGRESO_PATTERNS = [
    re.compile(r"ingreso.{0,240}?fecha.{0,60}?(\d{1,2}\s*[/\-.]\s*\d{1,2}\s*[/\-.]\s*\d{2,4})", re.IGNORECASE),
    re.compile(r"ingreso.{0,240}?(\d{1,2}\s*[/\-.]\s*\d{1,2}\s*[/\-.]\s*\d{2,4})", re.IGNORECASE),
]


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


def extraer_fecha_desde_texto(texto: str) -> tuple[str, str]:
    limpio = re.sub(r"\s+", " ", (texto or "").replace("\x00", " "))
    clave = normalizar_clave(limpio)
    if not clave:
        return "", ""

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


def construir_lista_archivos(carpeta: Path) -> list[Path]:
    candidatos = [
        p for p in carpeta.rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS
    ]

    def prioridad(archivo: Path) -> tuple[int, str]:
        nombre = normalizar_clave(archivo.name)
        for i, token in enumerate(FILE_NAME_PRIORITY):
            if token in nombre:
                return i, nombre
        return len(FILE_NAME_PRIORITY), nombre

    return sorted(candidatos, key=prioridad)[:MAX_FILES_PER_RADICADO]


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


def extraer_fecha_pdf(archivo: Path, engine: RapidOCR) -> OCRMatch | None:
    try:
        reader = PdfReader(str(archivo), strict=False)
        limite = min(len(reader.pages), MAX_PDF_TEXT_PAGES)
        for i in range(limite):
            texto = reader.pages[i].extract_text() or ""
            fecha, cruda = extraer_fecha_desde_texto(texto)
            if fecha:
                return OCRMatch(fecha, i + 1, cruda, "pdf-texto", "Fecha detectada por texto PDF", 9)
    except Exception:
        pass

    try:
        paginas = convert_from_path(
            str(archivo),
            dpi=220,
            first_page=1,
            last_page=MAX_OCR_PAGES,
            poppler_path=str(POPPLER_PATH) if POPPLER_PATH.exists() else None,
        )
    except Exception as error:
        return OCRMatch("", 0, "", "pdf-error", f"No se pudo convertir PDF: {error}", 1)

    for i, pagina in enumerate(paginas, start=1):
        texto = ocr_en_imagen(engine, pagina)
        fecha, cruda = extraer_fecha_desde_texto(texto)
        if fecha:
            return OCRMatch(fecha, i, cruda, "pdf-rapidocr", "Fecha detectada por OCR PDF", 8)

    return None


def extraer_fecha_tiff(archivo: Path, engine: RapidOCR) -> OCRMatch | None:
    try:
        with Image.open(archivo) as img:
            total = int(getattr(img, "n_frames", 1))
            for i in range(min(total, MAX_OCR_PAGES)):
                img.seek(i)
                pagina = img.convert("RGB")
                texto = ocr_en_imagen(engine, pagina)
                fecha, cruda = extraer_fecha_desde_texto(texto)
                if fecha:
                    return OCRMatch(fecha, i + 1, cruda, "tiff-rapidocr", "Fecha detectada por OCR TIFF", 8)
    except Exception as error:
        return OCRMatch("", 0, "", "tiff-error", f"No se pudo procesar TIFF: {error}", 1)

    return None


def extraer_fecha_imagen(archivo: Path, engine: RapidOCR) -> OCRMatch | None:
    try:
        with Image.open(archivo) as img:
            texto = ocr_en_imagen(engine, img.convert("RGB"))
        fecha, cruda = extraer_fecha_desde_texto(texto)
        if fecha:
            return OCRMatch(fecha, 1, cruda, "imagen-rapidocr", "Fecha detectada por OCR imagen", 7)
    except Exception as error:
        return OCRMatch("", 0, "", "imagen-error", f"No se pudo procesar imagen: {error}", 1)

    return None


def procesar_archivo(archivo: Path, engine: RapidOCR) -> OCRMatch | None:
    ext = archivo.suffix.lower()
    if ext == ".pdf":
        return extraer_fecha_pdf(archivo, engine)
    if ext in {".tif", ".tiff"}:
        return extraer_fecha_tiff(archivo, engine)
    return extraer_fecha_imagen(archivo, engine)


def procesar_radicado(radicado: str, carpeta: Path, engine: RapidOCR) -> ExtractionResult:
    archivos = construir_lista_archivos(carpeta)
    if not archivos:
        return ExtractionResult(radicado, "", 1, "sin_archivos", "", "", 0, "", "No hay archivos soportados")

    razones: list[str] = []
    for archivo in archivos:
        match = procesar_archivo(archivo, engine)
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


def main() -> None:
    if not ROOT_PATH.exists():
        raise FileNotFoundError(f"No existe la ruta raiz: {ROOT_PATH}")
    if not TARGET_RADICADOS_PATH.exists():
        raise FileNotFoundError(f"No existe archivo de radicados: {TARGET_RADICADOS_PATH}")

    objetivos = cargar_radicados_objetivo(TARGET_RADICADOS_PATH)
    seed = cargar_resultados_seed()
    engine = RapidOCR()

    pendientes_reales = [rad for rad in objetivos if not seed.get(rad)]

    print(f"Radicados objetivo: {len(objetivos)}")
    print(f"Ya resueltos por seed JS/CSV: {len(objetivos) - len(pendientes_reales)}")
    print(f"Pendientes OCR Python: {len(pendientes_reales)}")

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
    escribir_csv([mapa[rad] for rad in objetivos], OUTPUT_CSV_PATH)

    total = len(pendientes_reales)
    for idx, radicado in enumerate(pendientes_reales, start=1):
        carpeta = ROOT_PATH / radicado
        print(f"[{idx}/{total}] {radicado}")

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
            continue

        mapa[radicado] = procesar_radicado(radicado, carpeta, engine)

        if idx % 10 == 0 or idx == total:
            completar_fechas_faltantes(objetivos, mapa)
            parciales = [
                mapa.get(rad, ExtractionResult(rad, "1/01/24", 1, "aproximada", "fallback-global", "", 0, "1/01/24", "Fallback"))
                for rad in objetivos
            ]
            escribir_csv(parciales, OUTPUT_CSV_PATH)

    completar_fechas_faltantes(objetivos, mapa)

    resultados = [
        mapa.get(
            rad,
            ExtractionResult(rad, "1/01/24", 1, "aproximada", "fallback-global", "", 0, "1/01/24", "Fallback"),
        )
        for rad in objetivos
    ]

    escribir_csv(resultados, OUTPUT_CSV_PATH)
    escribir_excel_detalle(resultados, OUTPUT_EXCEL_PATH)

    exactas = sum(1 for r in resultados if r.estado in {"ok", "seed"})
    aproximadas = sum(1 for r in resultados if r.estado == "aproximada")
    promedio = round(sum(r.score_confianza for r in resultados) / max(1, len(resultados)), 2)

    print(f"Total radicados: {len(resultados)}")
    print(f"Exactas (OCR/seed): {exactas}")
    print(f"Aproximadas: {aproximadas}")
    print(f"Score promedio: {promedio}")
    print(f"CSV generado: {OUTPUT_CSV_PATH}")
    print(f"Excel detalle: {OUTPUT_EXCEL_PATH}")


if __name__ == "__main__":
    main()
