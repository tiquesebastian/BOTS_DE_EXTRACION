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
import pypdfium2 as pdfium
from rapidocr_onnxruntime import RapidOCR


ROOT_PATH = Path(r"Z:\IA 10\NUEVO\PARTE3\890701715")
TARGET_RADICADOS_PATH = Path(r"c:\Users\ticdesarrollo09\Music\trabajo\bot\numeros_pendientes")
OUTPUT_CSV_PATH = Path(r"c:\Users\ticdesarrollo09\Music\trabajo\bot\890701715_documentos_pdf.csv")
OUTPUT_EXCEL_PATH = Path(r"c:\Users\ticdesarrollo09\Music\trabajo\bot\890701715_documentos_pdf_detalle.xlsx")

PDF_REQUIRED_KEYWORDS = ("otros", "tapa", "resumen de atencion", "epicrisis")
POPPLER_PATH = Path(r"C:\poppler-25.11.0\Library\bin")
MAX_PDF_TEXT_PAGES = 0  # 0 = todas
MAX_OCR_PAGES = 0  # 0 = todas

DOCUMENTOS_BLOQUEADOS = {"890701715"}
PALABRAS_CONTEXTO_INSTITUCIONAL = (
    "nit",
    "hospital",
    "ips",
    "prestador",
    "razon social",
    "empresa",
)

IDENT_USER_PATTERN = re.compile(
    r"(?:nombre\s+usuario|usuario|estado\s+afiliacion\s+usuario).{0,260}?identif[a-z]{2,18}\s*[:\-]?\s*(\d{6,11})",
    re.IGNORECASE,
)
CEDULA_USER_PATTERN = re.compile(
    r"(?:nombre\s+del?\s+usuario|nombre\s+usuario|usuario).{0,220}?"
    r"(?:numero\s+de\s+cedula|n[úu]mero\s+de\s+c[eé]dula|cedula)\s*[:\-]?\s*(?:cc\s*)?(\d{6,11})",
    re.IGNORECASE,
)
CEDULA_PATTERN = re.compile(
    r"(?:numero\s+de\s+cedula|n[úu]mero\s+de\s+c[eé]dula|cedula(?:\s+de\s+ciudadania)?)\s*[:\-]?\s*(?:cc\s*)?(\d{6,11})",
    re.IGNORECASE,
)
IDENT_PATTERN = re.compile(r"(?:tipo\s+)?identif[a-z]{2,18}\s*[:\-]?\s*(\d{6,11})", re.IGNORECASE)
TYPE_DOC_PATTERN = re.compile(r"\b(?:cc|ti|ce|rc|pa|pep|ppt|dni)\s*[-:]?\s*(\d{6,11})\b", re.IGNORECASE)
GENERIC_PATTERN = re.compile(r"\b(\d{6,11})\b")
PATIENT_HINTS = (
    "usuario",
    "paciente",
    "identif",
    "cedula",
    "historia clinica",
    "datos del paciente",
    "tipo documento",
)


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
    m = re.search(r"\d{6,11}", valor or "")
    return m.group(0) if m else ""


def compactar_digitos(texto: str) -> str:
    return re.sub(r"(?<=\d)\s+(?=\d)", "", texto or "")


def parece_fecha_numerica(valor: str) -> bool:
    if len(valor) != 8 or not valor.isdigit():
        return False

    y = int(valor[0:4])
    m = int(valor[4:6])
    d = int(valor[6:8])
    if 1900 <= y <= 2099 and 1 <= m <= 12 and 1 <= d <= 31:
        return True

    d2 = int(valor[0:2])
    m2 = int(valor[2:4])
    y2 = int(valor[4:8])
    return 1 <= d2 <= 31 and 1 <= m2 <= 12 and 1900 <= y2 <= 2099


def esta_en_contexto_institucional(texto: str, numero: str) -> bool:
    idx = texto.find(numero)
    if idx == -1:
        return False
    ini = max(0, idx - 90)
    fin = min(len(texto), idx + len(numero) + 90)
    ventana = texto[ini:fin]
    return any(pal in ventana for pal in PALABRAS_CONTEXTO_INSTITUCIONAL)


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


def puntuar_ocr(base: int, ocr_score: float) -> int:
    if ocr_score >= 0.92:
        return min(10, base + 1)
    if ocr_score > 0:
        return base
    return max(1, base - 1)


def cargar_radicados_objetivo(path_archivo: Path) -> list[str]:
    radicados: list[str] = []
    with path_archivo.open("r", encoding="utf-8-sig", errors="ignore") as f:
        for linea in f:
            candidato = normalizar_radicado(linea.strip())
            if re.fullmatch(r"\d{6,15}", candidato):
                radicados.append(candidato)
    return list(dict.fromkeys(radicados))


def construir_lista_pdfs(carpeta: Path) -> list[Path]:
    archivos = [p for p in carpeta.rglob("*.pdf") if p.is_file()]

    keywords = [normalizar_clave(k) for k in PDF_REQUIRED_KEYWORDS if normalizar_clave(k)]

    def contiene_keyword(nombre_normalizado: str) -> bool:
        return any(k in nombre_normalizado for k in keywords)

    def prioridad_pdf(pdf_path: Path) -> tuple[int, str]:
        nombre = normalizar_clave(pdf_path.name)
        if "resumen de atencion" in nombre:
            prioridad = 0
        elif "tapa" in nombre:
            prioridad = 1
        elif "otros" in nombre:
            prioridad = 2
        else:
            prioridad = 3 if contiene_keyword(nombre) else 4
        return prioridad, nombre

    filtrados = [p for p in archivos if contiene_keyword(normalizar_clave(p.name))]
    return sorted(filtrados, key=prioridad_pdf)


def elegir_mejor_candidato(texto: str, radicado: str) -> str:
    clave = normalizar_clave(compactar_digitos(texto or ""))
    if not clave:
        return ""

    candidatos: dict[str, int] = {}
    for m in GENERIC_PATTERN.finditer(clave):
        doc = normalizar_documento(m.group(1))
        if not es_documento_valido(doc, radicado, clave):
            continue

        inicio = max(0, m.start() - 100)
        fin = min(len(clave), m.end() + 100)
        ventana = clave[inicio:fin]

        score = 1
        if any(h in ventana for h in PATIENT_HINTS):
            score += 4
        if re.search(r"\b(?:cc|ti|ce|rc|pa|dni)\s*[:\-]?\s*" + re.escape(doc), ventana):
            score += 3
        if "imprimir liquidacion" in ventana:
            score += 2
        if esta_en_contexto_institucional(clave, doc):
            score -= 3

        candidatos[doc] = max(candidatos.get(doc, -999), score)

    if not candidatos:
        return ""

    return sorted(candidatos.items(), key=lambda x: (-x[1], -len(x[0])))[0][0]


def preprocesar_imagen(imagen: Image.Image) -> np.ndarray:
    gris = ImageOps.grayscale(imagen)
    gris = ImageOps.autocontrast(gris)
    ancho, alto = gris.size

    max_lado = 2200
    if max(ancho, alto) > max_lado:
        escala = max_lado / max(ancho, alto)
        gris = gris.resize((max(1, int(ancho * escala)), max(1, int(alto * escala))), Image.Resampling.LANCZOS)
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


def render_pdf_page(pdf_path: Path, page_number_1: int) -> Image.Image | None:
    # page_number_1 es 1-based
    if POPPLER_PATH.exists():
        try:
            pages = convert_from_path(
                str(pdf_path),
                dpi=220,
                first_page=page_number_1,
                last_page=page_number_1,
                poppler_path=str(POPPLER_PATH),
            )
            if pages:
                return pages[0]
        except Exception:
            pass

    # Fallback robusto sin poppler
    try:
        doc = pdfium.PdfDocument(str(pdf_path))
        page_index = page_number_1 - 1
        if page_index < 0 or page_index >= len(doc):
            return None
        page = doc[page_index]
        bitmap = page.render(scale=2.0)
        return bitmap.to_pil()
    except Exception:
        return None


def extraer_documento_desde_texto(texto: str, radicado: str) -> tuple[str, str, int, str]:
    if not texto:
        return "", "", 1, "Texto vacio"

    limpio = compactar_digitos(re.sub(r"\s+", " ", texto.replace("\x00", " ")))
    clave = normalizar_clave(limpio)
    if not clave:
        return "", "", 1, "Texto no util"

    m = CEDULA_USER_PATTERN.search(clave)
    if m:
        doc = normalizar_documento(m.group(1))
        if es_documento_valido(doc, radicado, clave):
            return doc, "usuario-cedula", 10, "Documento por contexto usuario + numero de cedula"

    m = CEDULA_PATTERN.search(clave)
    if m:
        doc = normalizar_documento(m.group(1))
        if es_documento_valido(doc, radicado, clave):
            return doc, "cedula", 9, "Documento por etiqueta numero de cedula"

    m = IDENT_USER_PATTERN.search(clave)
    if m:
        doc = normalizar_documento(m.group(1))
        if es_documento_valido(doc, radicado, clave):
            return doc, "usuario-identificacion", 10, "Documento por contexto usuario + identificacion"

    m = IDENT_PATTERN.search(clave)
    if m:
        doc = normalizar_documento(m.group(1))
        if es_documento_valido(doc, radicado, clave):
            score = 10 if "imprimir liquidacion" in clave else 9
            return doc, "identificacion", score, "Documento por etiqueta identificacion"

    m = TYPE_DOC_PATTERN.search(clave)
    if m:
        doc = normalizar_documento(m.group(1))
        if es_documento_valido(doc, radicado, clave):
            return doc, "tipo-doc", 8, "Documento por tipo (cc/ti/ce/etc)"

    doc_rankeado = elegir_mejor_candidato(clave, radicado)
    if doc_rankeado:
        return doc_rankeado, "generico-rankeado", 6, "Documento por ranking de contexto"

    return "", "", 1, "No se detecto numero de documento"


def extraer_documento_pdf(archivo: Path, radicado: str, engine: RapidOCR) -> OCRMatch | None:
    total_paginas = 0

    try:
        reader = PdfReader(str(archivo), strict=False)
        total_paginas = len(reader.pages)
        limite = total_paginas if MAX_PDF_TEXT_PAGES <= 0 else min(total_paginas, MAX_PDF_TEXT_PAGES)
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

    try:
        for i in range(1, max(1, limite_ocr) + 1):
            pagina = render_pdf_page(archivo, i)
            if pagina is None:
                break

            texto, ocr_score = ocr_en_imagen(engine, pagina)
            doc, metodo, base_score, motivo = extraer_documento_desde_texto(texto, radicado)
            if doc:
                return OCRMatch(doc, i, doc, f"pdf-rapidocr-{metodo}", motivo, puntuar_ocr(base_score, ocr_score))
    except Exception as error:
        return OCRMatch("", 0, "", "pdf-error", f"No se pudo procesar PDF: {error}", 1)

    return None


def procesar_radicado(radicado: str, carpeta: Path, engine: RapidOCR) -> ExtractionResult:
    archivos = construir_lista_pdfs(carpeta)
    if not archivos:
        return ExtractionResult(
            radicado,
            "SIN_DATO",
            1,
            "sin_archivos_pdf_otros_tapa",
            "",
            "",
            0,
            "",
            "No hay PDF con keyword 'otros', 'tapa', 'resumen de atencion' o 'epicrisis'",
        )

    razones: list[str] = []
    for archivo in archivos:
        match = extraer_documento_pdf(archivo, radicado, engine)
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
    engine = RapidOCR()

    print(f"Radicados objetivo: {len(objetivos)}")
    print(f"Filtro tipo archivo: PDF")
    print("Filtro keyword nombre archivo: otros / tapa / resumen de atencion / epicrisis")

    mapa: dict[str, ExtractionResult] = {}
    for radicado in objetivos:
        mapa[radicado] = ExtractionResult(
            radicado=radicado,
            numero_documento="PENDIENTE",
            score_confianza=1,
            estado="pendiente",
            metodo="",
            archivo_origen="",
            pagina=0,
            doc_crudo="",
            motivo="Pendiente",
        )

    escribir_csv([mapa[rad] for rad in objetivos], OUTPUT_CSV_PATH)

    total = len(objetivos)
    for idx, radicado in enumerate(objetivos, start=1):
        carpeta = ROOT_PATH / radicado
        print(f"[{idx}/{total}] {radicado}")

        try:
            existe = carpeta.exists()
            es_dir = carpeta.is_dir() if existe else False
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
            continue

        if not existe or not es_dir:
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

        escribir_csv([mapa[rad] for rad in objetivos], OUTPUT_CSV_PATH)

    resultados = [mapa.get(rad) for rad in objetivos]

    escribir_csv(resultados, OUTPUT_CSV_PATH)
    escribir_excel_detalle(resultados, OUTPUT_EXCEL_PATH)

    encontrados = sum(1 for r in resultados if r and r.numero_documento not in {"", "PENDIENTE", "SIN_DATO"})
    sin_doc = sum(1 for r in resultados if r and r.numero_documento in {"", "PENDIENTE", "SIN_DATO"})
    promedio = round(sum((r.score_confianza if r else 1) for r in resultados) / max(1, len(resultados)), 2)

    print(f"Total radicados: {len(resultados)}")
    print(f"Con documento: {encontrados}")
    print(f"Sin documento: {sin_doc}")
    print(f"Score promedio: {promedio}")
    print(f"CSV generado: {OUTPUT_CSV_PATH}")
    print(f"Excel detalle: {OUTPUT_EXCEL_PATH}")


if __name__ == "__main__":
    main()
