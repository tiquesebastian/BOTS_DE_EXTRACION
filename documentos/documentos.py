from __future__ import annotations

import csv
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path

import numpy as np
from openpyxl import Workbook
from pdf2image import convert_from_path
from PIL import Image, ImageFilter, ImageOps
from pypdf import PdfReader
from rapidocr_onnxruntime import RapidOCR


ROOT_PATH = Path(r"Z:\IA 10\NUEVO\PARTE3\890701715")
TARGET_RADICADOS_PATH = Path(r"c:\Users\ticdesarrollo09\Music\trabajo\bot\numeros_pendientes")
OUTPUT_CSV_PATH = Path(r"c:\Users\ticdesarrollo09\Music\trabajo\bot\890701715_documentos.csv")
OUTPUT_EXCEL_PATH = Path(r"c:\Users\ticdesarrollo09\Music\trabajo\bot\890701715_documentos_detalle.xlsx")

SEED_RESULT_PATHS = [OUTPUT_CSV_PATH]

MAX_PDF_TEXT_PAGES = 0  # 0 = todas las paginas
MAX_OCR_PAGES = 0  # 0 = todas las paginas
MAX_FILES_PER_RADICADO = 0  # 0 = todos los archivos soportados del radicado
FAST_SCAN_PAGES = 10
POPPLER_PATH = Path(r"C:\poppler-25.11.0\Library\bin")

FILE_NAME_PRIORITY = [
    "imprimir liquidacion",
    "resumen de atencion",
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

DOCUMENTOS_BLOQUEADOS = {
    "890701715",  # NIT del hospital, no documento del paciente
}

PALABRAS_CONTEXTO_INSTITUCIONAL = (
    "nit",
    "hospital",
    "ips",
    "prestador",
    "razon social",
    "empresa",
)

LABELLED_DOC_PATTERN = re.compile(
    r"(?:paciente|usuario|afiliado|documento(?:\s+de\s+identidad)?|identificacion|id)"
    r"\s*[:\-]?\s*(?:tipo\s*[:\-]?\s*)?(?:cc|ti|ce|rc|pa|pep|ppt|nit|dni)?\s*[-:]?\s*(\d{5,15})",
    re.IGNORECASE,
)

IDENTIFICACION_PATTERN = re.compile(
    r"(?:tipo\s+)?identificacion\s*[:\-]?\s*(\d{5,15})",
    re.IGNORECASE,
)

USUARIO_CERCA_IDENTIFICACION_PATTERN = re.compile(
    r"(?:nombre\s+usuario|usuario|estado\s+afiliacion\s+usuario).{0,260}?identificacion\s*[:\-]?\s*(\d{5,15})",
    re.IGNORECASE,
)

USUARIO_CERCA_CEDULA_PATTERN = re.compile(
    r"(?:nombre\s+del?\s+usuario|nombre\s+usuario|usuario).{0,220}?"
    r"(?:numero\s+de\s+cedula|n[úu]mero\s+de\s+c[eé]dula|cedula)\s*[:\-]?\s*(?:cc\s*)?(\d{6,11})",
    re.IGNORECASE,
)

CEDULA_PATTERN = re.compile(
    r"(?:numero\s+de\s+cedula|n[úu]mero\s+de\s+c[eé]dula|cedula(?:\s+de\s+ciudadania)?)\s*[:\-]?\s*(?:cc\s*)?(\d{6,11})",
    re.IGNORECASE,
)

TYPE_DOC_PATTERN = re.compile(
    r"\b(?:cc|ti|ce|rc|pa|pep|ppt|nit|dni)\s*[-:]?\s*(\d{5,15})\b",
    re.IGNORECASE,
)

GENERIC_DOC_PATTERN = re.compile(r"\b(\d{7,15})\b")


@dataclass
class ExtractionResult:
    radicado: str
    numero_documento: str
    score_confianza: int
    estado: str
    metodo: str
    archivo_origen: str
    pagina: int
    doc_crudo: str
    motivo: str


@dataclass
class OCRMatch:
    numero_documento: str
    pagina: int
    doc_crudo: str
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


def normalizar_documento(valor: str) -> str:
    m = re.search(r"\d{5,15}", valor or "")
    return m.group(0) if m else ""


def compactar_digitos(texto: str) -> str:
    return re.sub(r"(?<=\d)\s+(?=\d)", "", texto or "")


def parece_fecha_numerica(valor: str) -> bool:
    if len(valor) != 8 or not valor.isdigit():
        return False

    anio = int(valor[0:4])
    mes = int(valor[4:6])
    dia = int(valor[6:8])
    if 1900 <= anio <= 2099 and 1 <= mes <= 12 and 1 <= dia <= 31:
        return True

    dia2 = int(valor[0:2])
    mes2 = int(valor[2:4])
    anio2 = int(valor[4:8])
    return 1 <= dia2 <= 31 and 1 <= mes2 <= 12 and 1900 <= anio2 <= 2099


def esta_en_contexto_institucional(texto: str, numero: str) -> bool:
    idx = texto.find(numero)
    if idx == -1:
        return False
    inicio = max(0, idx - 80)
    fin = min(len(texto), idx + len(numero) + 80)
    ventana = texto[inicio:fin]
    return any(palabra in ventana for palabra in PALABRAS_CONTEXTO_INSTITUCIONAL)


def es_documento_valido(doc: str, radicado: str, texto: str) -> bool:
    if not doc:
        return False
    if doc == radicado:
        return False
    if doc in DOCUMENTOS_BLOQUEADOS:
        return False
    if parece_fecha_numerica(doc):
        return False
    if esta_en_contexto_institucional(texto, doc):
        return False
    return True


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
                    numero = normalizar_documento(
                        (row.get("numero_documento") or row.get("documento") or row.get("paciente") or "").strip()
                    )
                    if radicado and numero and numero not in DOCUMENTOS_BLOQUEADOS:
                        seed[radicado] = numero
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

    ordenados = sorted(candidatos, key=prioridad)
    if MAX_FILES_PER_RADICADO and MAX_FILES_PER_RADICADO > 0:
        return ordenados[:MAX_FILES_PER_RADICADO]
    return ordenados


def preprocesar_imagen(imagen: Image.Image) -> np.ndarray:
    gris = ImageOps.grayscale(imagen)
    gris = ImageOps.autocontrast(gris)
    ancho, alto = gris.size
    # Evita cuelgues del motor OCR con paginas demasiado grandes.
    max_lado = 2200
    if max(ancho, alto) > max_lado:
        escala = max_lado / max(ancho, alto)
        nuevo_ancho = max(1, int(ancho * escala))
        nuevo_alto = max(1, int(alto * escala))
        gris = gris.resize((nuevo_ancho, nuevo_alto), Image.Resampling.LANCZOS)
        ancho, alto = gris.size
    if ancho < 1800:
        factor = max(1, int(1800 / max(1, ancho)))
        gris = gris.resize((ancho * factor, alto * factor), Image.Resampling.LANCZOS)
    gris = gris.filter(ImageFilter.SHARPEN)
    return np.array(gris)


def ocr_en_imagen(engine: RapidOCR, imagen: Image.Image) -> tuple[str, float]:
    try:
        resultado, _ = engine(preprocesar_imagen(imagen))
    except Exception:
        return "", 0.0
    if not resultado:
        return "", 0.0

    textos = []
    scores = []
    for item in resultado:
        textos.append(item[1])
        if len(item) >= 3 and isinstance(item[2], (float, int)):
            scores.append(float(item[2]))

    promedio = sum(scores) / len(scores) if scores else 0.0
    return " ".join(textos), promedio


def puntuar_ocr(base: int, ocr_score: float) -> int:
    if ocr_score >= 0.92:
        return min(10, base + 2)
    if ocr_score >= 0.84:
        return min(10, base + 1)
    if ocr_score > 0:
        return base
    return max(1, base - 1)


def extraer_documento_desde_texto(texto: str, radicado: str) -> tuple[str, str, int, str]:
    if not texto:
        return "", "", 1, "Texto vacio"

    normal = compactar_digitos(re.sub(r"\s+", " ", texto.replace("\x00", " ")))
    clave = normalizar_clave(normal)
    if not clave:
        return "", "", 1, "Texto no util"

    m = USUARIO_CERCA_CEDULA_PATTERN.search(clave)
    if m:
        doc = normalizar_documento(m.group(1))
        if es_documento_valido(doc, radicado, clave):
            return doc, "usuario-cedula", 10, "Documento por contexto usuario + numero de cedula"

    m = USUARIO_CERCA_IDENTIFICACION_PATTERN.search(clave)
    if m:
        doc = normalizar_documento(m.group(1))
        if es_documento_valido(doc, radicado, clave):
            return doc, "usuario-identificacion", 10, "Documento por contexto usuario + identificacion"

    m = CEDULA_PATTERN.search(clave)
    if m:
        doc = normalizar_documento(m.group(1))
        if es_documento_valido(doc, radicado, clave):
            return doc, "cedula", 9, "Documento por etiqueta numero de cedula"

    m = IDENTIFICACION_PATTERN.search(clave)
    if m:
        doc = normalizar_documento(m.group(1))
        if es_documento_valido(doc, radicado, clave):
            score = 10 if "imprimir liquidacion" in clave else 9
            return doc, "identificacion", score, "Documento por etiqueta identificacion"

    m = LABELLED_DOC_PATTERN.search(clave)
    if m:
        doc = normalizar_documento(m.group(1))
        if es_documento_valido(doc, radicado, clave):
            return doc, "patron-etiquetado", 9, "Documento por etiqueta paciente/documento"

    m = TYPE_DOC_PATTERN.search(clave)
    if m:
        doc = normalizar_documento(m.group(1))
        if es_documento_valido(doc, radicado, clave):
            return doc, "patron-tipo-doc", 8, "Documento por tipo (cc/ti/ce/etc)"

    idx = clave.find("paciente")
    if idx != -1:
        segmento = clave[idx : idx + 180]
        m = GENERIC_DOC_PATTERN.search(segmento)
        if m:
            doc = normalizar_documento(m.group(1))
            if es_documento_valido(doc, radicado, clave):
                return doc, "paciente-cercano", 7, "Documento cercano a palabra paciente"

    m = GENERIC_DOC_PATTERN.search(clave)
    if m:
        doc = normalizar_documento(m.group(1))
        if es_documento_valido(doc, radicado, clave):
            return doc, "generico", 5, "Documento por patron numerico generico"

    return "", "", 1, "No se detecto numero de documento"


def extraer_documento_pdf(archivo: Path, radicado: str, engine: RapidOCR, page_limit: int = 0) -> OCRMatch | None:
    total_paginas = 0
    try:
        reader = PdfReader(str(archivo), strict=False)
        total_paginas = len(reader.pages)
        base_limit = total_paginas if MAX_PDF_TEXT_PAGES <= 0 else min(total_paginas, MAX_PDF_TEXT_PAGES)
        limite = min(base_limit, page_limit) if page_limit > 0 else base_limit
        for i in range(limite):
            texto = reader.pages[i].extract_text() or ""
            doc, metodo, score, motivo = extraer_documento_desde_texto(texto, radicado)
            if doc:
                return OCRMatch(doc, i + 1, doc, f"pdf-texto-{metodo}", motivo, score)
    except Exception:
        pass

    limite_ocr = total_paginas if (total_paginas > 0 and MAX_OCR_PAGES <= 0) else MAX_OCR_PAGES
    if total_paginas == 0 and MAX_OCR_PAGES <= 0:
        limite_ocr = 120
    if page_limit > 0:
        limite_ocr = min(limite_ocr, page_limit)

    try:
        for i in range(1, max(1, limite_ocr) + 1):
            paginas = convert_from_path(
                str(archivo),
                dpi=220,
                first_page=i,
                last_page=i,
                poppler_path=str(POPPLER_PATH) if POPPLER_PATH.exists() else None,
            )
            if not paginas:
                break
            texto, ocr_score = ocr_en_imagen(engine, paginas[0])
            doc, metodo, base_score, motivo = extraer_documento_desde_texto(texto, radicado)
            if doc:
                return OCRMatch(doc, i, doc, f"pdf-rapidocr-{metodo}", motivo, puntuar_ocr(base_score, ocr_score))
    except Exception as error:
        return OCRMatch("", 0, "", "pdf-error", f"No se pudo convertir PDF: {error}", 1)

    return None


def extraer_documento_tiff(archivo: Path, radicado: str, engine: RapidOCR, page_limit: int = 0) -> OCRMatch | None:
    try:
        with Image.open(archivo) as img:
            total = int(getattr(img, "n_frames", 1))
            limite = total if MAX_OCR_PAGES <= 0 else min(total, MAX_OCR_PAGES)
            if page_limit > 0:
                limite = min(limite, page_limit)
            for i in range(limite):
                img.seek(i)
                pagina = img.convert("RGB")
                texto, ocr_score = ocr_en_imagen(engine, pagina)
                doc, metodo, base_score, motivo = extraer_documento_desde_texto(texto, radicado)
                if doc:
                    return OCRMatch(doc, i + 1, doc, f"tiff-rapidocr-{metodo}", motivo, puntuar_ocr(base_score, ocr_score))
    except Exception as error:
        return OCRMatch("", 0, "", "tiff-error", f"No se pudo procesar TIFF: {error}", 1)

    return None


def extraer_documento_imagen(archivo: Path, radicado: str, engine: RapidOCR, page_limit: int = 0) -> OCRMatch | None:
    try:
        with Image.open(archivo) as img:
            texto, ocr_score = ocr_en_imagen(engine, img.convert("RGB"))
        doc, metodo, base_score, motivo = extraer_documento_desde_texto(texto, radicado)
        if doc:
            return OCRMatch(doc, 1, doc, f"imagen-rapidocr-{metodo}", motivo, puntuar_ocr(base_score, ocr_score))
    except Exception as error:
        return OCRMatch("", 0, "", "imagen-error", f"No se pudo procesar imagen: {error}", 1)

    return None


def procesar_archivo(archivo: Path, radicado: str, engine: RapidOCR, page_limit: int = 0) -> OCRMatch | None:
    ext = archivo.suffix.lower()
    if ext == ".pdf":
        return extraer_documento_pdf(archivo, radicado, engine, page_limit=page_limit)
    if ext in {".tif", ".tiff"}:
        return extraer_documento_tiff(archivo, radicado, engine, page_limit=page_limit)
    return extraer_documento_imagen(archivo, radicado, engine, page_limit=page_limit)


def procesar_radicado(radicado: str, carpeta: Path, engine: RapidOCR) -> ExtractionResult:
    archivos = construir_lista_archivos(carpeta)
    if not archivos:
        return ExtractionResult(radicado, "SIN_DATO", 1, "sin_archivos", "", "", 0, "", "No hay archivos soportados")

    razones: list[str] = []

    # Fase 1: barrido rapido para dar resultados tempranos.
    for archivo in archivos:
        match = procesar_archivo(archivo, radicado, engine, page_limit=FAST_SCAN_PAGES)
        if match and match.numero_documento:
            return ExtractionResult(
                radicado=radicado,
                numero_documento=match.numero_documento,
                score_confianza=match.score_confianza,
                estado="ok",
                metodo=match.metodo,
                archivo_origen=str(archivo),
                pagina=match.pagina,
                doc_crudo=match.doc_crudo,
                motivo=f"[FAST] {match.motivo}",
            )
        if match:
            razones.append(f"{archivo.name}: {match.metodo}")

    # Fase 2: barrido completo de todas las paginas si no se hallo en fase rapida.
    for archivo in archivos:
        match = procesar_archivo(archivo, radicado, engine, page_limit=0)
        if match and match.numero_documento:
            return ExtractionResult(
                radicado=radicado,
                numero_documento=match.numero_documento,
                score_confianza=match.score_confianza,
                estado="ok",
                metodo=match.metodo,
                archivo_origen=str(archivo),
                pagina=match.pagina,
                doc_crudo=match.doc_crudo,
                motivo=match.motivo,
            )
        if match:
            razones.append(f"{archivo.name}: {match.metodo}")

    return ExtractionResult(
        radicado=radicado,
        numero_documento="SIN_DATO",
        score_confianza=1,
        estado="sin_documento",
        metodo="",
        archivo_origen=str(archivos[0]),
        pagina=0,
        doc_crudo="",
        motivo=" | ".join(razones[:4]) if razones else "No se detecto documento",
    )


def escribir_csv(resultados: list[ExtractionResult], salida: Path) -> None:
    with salida.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["radicado", "numero_documento", "score_confianza"])
        for item in resultados:
            writer.writerow([item.radicado, item.numero_documento, item.score_confianza])


def escribir_excel_detalle(resultados: list[ExtractionResult], salida: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "resultados"
    ws.append([
        "radicado",
        "numero_documento",
        "score_confianza",
        "estado",
        "metodo",
        "archivo_origen",
        "pagina",
        "doc_crudo",
        "motivo",
    ])
    for item in resultados:
        ws.append([
            item.radicado,
            item.numero_documento,
            item.score_confianza,
            item.estado,
            item.metodo,
            item.archivo_origen,
            item.pagina,
            item.doc_crudo,
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
    print(f"Ya resueltos por seed CSV: {len(objetivos) - len(pendientes_reales)}")
    print(f"Pendientes OCR documentos: {len(pendientes_reales)}")

    mapa: dict[str, ExtractionResult] = {}
    for radicado in objetivos:
        documento_seed = seed.get(radicado, "")
        if documento_seed:
            mapa[radicado] = ExtractionResult(
                radicado=radicado,
                numero_documento=documento_seed,
                score_confianza=8,
                estado="seed",
                metodo="seed-csv",
                archivo_origen="",
                pagina=0,
                doc_crudo=documento_seed,
                motivo="Documento tomado de salida existente",
            )

    for radicado in objetivos:
        mapa.setdefault(
            radicado,
            ExtractionResult(
                radicado=radicado,
                numero_documento="PENDIENTE",
                score_confianza=1,
                estado="pendiente",
                metodo="",
                archivo_origen="",
                pagina=0,
                doc_crudo="",
                motivo="Pendiente",
            ),
        )

    escribir_csv([mapa[rad] for rad in objetivos], OUTPUT_CSV_PATH)

    total = len(pendientes_reales)
    for idx, radicado in enumerate(pendientes_reales, start=1):
        carpeta = ROOT_PATH / radicado
        print(f"[{idx}/{total}] {radicado}")

        try:
            existe_carpeta = carpeta.exists()
            es_directorio = carpeta.is_dir() if existe_carpeta else False
        except OSError as error:
            mapa[radicado] = ExtractionResult(
                radicado,
                "SIN_DATO",
                1,
                "error_red",
                "",
                "",
                0,
                "",
                f"Error de red al validar carpeta: {error}",
            )
            print(f"    -> error de red ({error})")
            parciales = [
                mapa.get(rad, ExtractionResult(rad, "PENDIENTE", 1, "pendiente", "", "", 0, "", "Pendiente"))
                for rad in objetivos
            ]
            escribir_csv(parciales, OUTPUT_CSV_PATH)
            continue

        if not existe_carpeta or not es_directorio:
            mapa[radicado] = ExtractionResult(
                radicado,
                "SIN_DATO",
                1,
                "sin_carpeta",
                "",
                "",
                0,
                "",
                "No existe carpeta del radicado",
            )
            continue

        try:
            mapa[radicado] = procesar_radicado(radicado, carpeta, engine)
        except OSError as error:
            mapa[radicado] = ExtractionResult(
                radicado,
                "SIN_DATO",
                1,
                "error_red",
                "",
                str(carpeta),
                0,
                "",
                f"Error de red procesando radicado: {error}",
            )
        except Exception as error:
            mapa[radicado] = ExtractionResult(
                radicado,
                "SIN_DATO",
                1,
                "error_proceso",
                "",
                str(carpeta),
                0,
                "",
                f"Error inesperado procesando radicado: {error}",
            )

        resultado = mapa[radicado]
        if resultado.numero_documento not in {"", "PENDIENTE", "SIN_DATO"}:
            print(f"    -> doc={resultado.numero_documento} score={resultado.score_confianza} metodo={resultado.metodo}")
        else:
            print(f"    -> sin documento ({resultado.estado})")

        if idx % 1 == 0 or idx == total:
            parciales = [
                mapa.get(rad, ExtractionResult(rad, "PENDIENTE", 1, "pendiente", "", "", 0, "", "Pendiente"))
                for rad in objetivos
            ]
            escribir_csv(parciales, OUTPUT_CSV_PATH)

    resultados = [
        mapa.get(rad, ExtractionResult(rad, "SIN_DATO", 1, "sin_documento", "", "", 0, "", "Sin dato"))
        for rad in objetivos
    ]

    escribir_csv(resultados, OUTPUT_CSV_PATH)
    escribir_excel_detalle(resultados, OUTPUT_EXCEL_PATH)

    encontrados = sum(1 for r in resultados if r.estado in {"ok", "seed"} and r.numero_documento)
    sin_doc = sum(1 for r in resultados if r.numero_documento in {"", "SIN_DATO", "PENDIENTE"})
    promedio = round(sum(r.score_confianza for r in resultados) / max(1, len(resultados)), 2)

    print(f"Total radicados: {len(resultados)}")
    print(f"Con documento: {encontrados}")
    print(f"Sin documento: {sin_doc}")
    print(f"Score promedio: {promedio}")
    print(f"CSV generado: {OUTPUT_CSV_PATH}")
    print(f"Excel detalle: {OUTPUT_EXCEL_PATH}")


if __name__ == "__main__":
    main()
