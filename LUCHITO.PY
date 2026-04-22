# -*- coding: utf-8 -*-
"""
Extractor Paciente + Ingreso + Autorización
- PDF / TIFF / MultiTIFF
- Filtro por radicados
- Filtro por palabra clave
- Salida CSV delimitado por |
- Multihilo
- Interfaz Tkinter

Requisitos:
pip install pillow pytesseract pdf2image pandas

Además:
- Instalar Tesseract OCR
- Instalar Poppler para PDF
"""

import os
import re
import csv
import threading
import queue
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageSequence, ImageOps, ImageFilter, ImageEnhance

import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# =========================================================
# CONFIGURACIÓN
# =========================================================

# Ajusta estas rutas según tu equipo
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH = r"C:\poppler-23.11.0\Library\bin"

# Extensiones válidas
VALID_EXTENSIONS = {".pdf", ".tif", ".tiff"}

# Tipos de archivo soportados
FILETYPE_BOTH = "Ambos"
FILETYPE_PDF = "PDF"
FILETYPE_TIFF = "TIFF"

# OCR configs
OCR_CONFIGS = [
    r'--oem 3 --psm 6 -l spa',
    r'--oem 3 --psm 11 -l spa',
    r'--oem 3 --psm 4 -l spa',
]


# =========================================================
# UTILIDADES
# =========================================================

def safe_str(value) -> str:
    return "" if value is None else str(value).strip()


def normalize_spaces(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_radicados(text: str) -> list[str]:
    """
    Recibe radicados uno por línea.
    """
    lines = [x.strip() for x in text.splitlines()]
    return [x for x in lines if x]


def is_valid_file(filename: str, mode: str, keyword: str) -> bool:
    ext = os.path.splitext(filename)[1].lower()

    if ext not in VALID_EXTENSIONS:
        return False

    if mode == FILETYPE_PDF and ext != ".pdf":
        return False

    if mode == FILETYPE_TIFF and ext not in {".tif", ".tiff"}:
        return False

    if keyword:
        if keyword.lower() not in filename.lower():
            return False

    return True


def preprocess_image(img: Image.Image) -> list[Image.Image]:
    """
    Devuelve varias versiones de la imagen para intentar OCR robusto.
    """
    variants = []

    # Original escala de grises
    base = img.convert("L")
    variants.append(base)

    # Más contraste
    contrast = ImageEnhance.Contrast(base).enhance(2.0)
    variants.append(contrast)

    # Autocontrast
    auto = ImageOps.autocontrast(base)
    variants.append(auto)

    # Binarizada suave
    bin_img = ImageOps.autocontrast(base)
    bin_img = bin_img.point(lambda p: 255 if p > 170 else 0)
    variants.append(bin_img)

    # Sharpen
    sharp = auto.filter(ImageFilter.SHARPEN)
    variants.append(sharp)

    return variants


def ocr_image_multi(img: Image.Image) -> str:
    """
    Ejecuta OCR con varias variantes de imagen y configs.
    Retorna el mejor texto por longitud.
    """
    best_text = ""
    variants = preprocess_image(img)

    for variant in variants:
        for cfg in OCR_CONFIGS:
            try:
                text = pytesseract.image_to_string(variant, config=cfg)
                text = normalize_spaces(text)
                if len(text) > len(best_text):
                    best_text = text
            except Exception:
                continue

    return best_text


# =========================================================
# EXTRACCIÓN
# =========================================================

def extract_paciente_ingreso_autorizacion(text: str) -> tuple[str, str, str]:
    """
    Extrae:
    - Paciente
    - Ingreso
    - N° Autorización

    Ejemplo esperado:
    Paciente: CC-29614680 ...
    Ingreso: 25/01/2024 ...
    N° Autorización: 227599046
    """
    if not text:
        return "", "", ""

    text = normalize_spaces(text)

    # -------------------------
    # PACIENTE
    # Ej:
    # Paciente: CC-29614680
    # Paciente CC29614680
    # Paciente: TI-12345678
    # -------------------------
    paciente = ""
    m_p = re.search(
        r"Paciente\s*[:\-]?\s*([A-Z]{2,5}\s*-\s*\d{5,20}|[A-Z]{2,5}\d{5,20})",
        text,
        flags=re.IGNORECASE
    )
    if m_p:
        paciente = safe_str(m_p.group(1))
        paciente = re.sub(r"\s+", "", paciente)

    # -------------------------
    # INGRESO
    # Ej:
    # Ingreso: 25/01/2024
    # -------------------------
    ingreso = ""
    m_i = re.search(
        r"Ingreso\s*[:\-]?\s*(\d{1,2}/\d{1,2}/\d{4})",
        text,
        flags=re.IGNORECASE
    )
    if m_i:
        ingreso = safe_str(m_i.group(1))

    # -------------------------
    # AUTORIZACIÓN
    # Ej:
    # N° Autorización: 227599046
    # No Autorizacion: 227599046
    # Número Autorización: 227599046
    # -------------------------
    autorizacion = ""
    m_a = re.search(
        r"(?:N[°ºo]\s*Autorizaci[oó]n|No\s*Autorizaci[oó]n|Numero\s*Autorizaci[oó]n|Número\s*Autorizaci[oó]n)\s*[:\-]?\s*(\d{5,20})",
        text,
        flags=re.IGNORECASE
    )
    if m_a:
        autorizacion = safe_str(m_a.group(1))

    return paciente, ingreso, autorizacion


# =========================================================
# CARGA DE PÁGINAS
# =========================================================

def load_pdf_pages(pdf_path: str) -> list[Image.Image]:
    """
    Convierte PDF a imágenes PIL, página por página.
    """
    pages = convert_from_path(
        pdf_path,
        dpi=300,
        poppler_path=POPPLER_PATH
    )
    return pages


def load_tiff_pages(tiff_path: str) -> list[Image.Image]:
    """
    Lee TIFF uni o multipágina.
    """
    pages = []
    with Image.open(tiff_path) as img:
        for frame in ImageSequence.Iterator(img):
            pages.append(frame.copy())
    return pages


# =========================================================
# PROCESAMIENTO DE ARCHIVOS
# =========================================================

def process_file(file_path: str, radicado: str, stop_event: threading.Event) -> dict:
    """
    Procesa un archivo y retorna un dict resultado.
    Se detiene al encontrar Paciente / Ingreso / Autorización.
    """
    result = {
        "radicado": radicado,
        "carpeta": radicado,
        "archivo": os.path.basename(file_path),
        "pagina": "",
        "paciente": "",
        "ingreso": "",
        "autorizacion": "",
        "estado": "no_encontrado",
        "observacion": "",
    }

    if stop_event.is_set():
        result["estado"] = "detenido"
        result["observacion"] = "Proceso detenido por el usuario"
        return result

    ext = os.path.splitext(file_path)[1].lower()

    try:
        if ext == ".pdf":
            pages = load_pdf_pages(file_path)
        elif ext in {".tif", ".tiff"}:
            pages = load_tiff_pages(file_path)
        else:
            result["estado"] = "omitido"
            result["observacion"] = "Extensión no soportada"
            return result

        for i, page_img in enumerate(pages, start=1):
            if stop_event.is_set():
                result["estado"] = "detenido"
                result["observacion"] = "Proceso detenido por el usuario"
                return result

            text = ocr_image_multi(page_img)
            paciente, ingreso, autorizacion = extract_paciente_ingreso_autorizacion(text)

            if paciente or ingreso or autorizacion:
                result["pagina"] = i
                result["paciente"] = paciente
                result["ingreso"] = ingreso
                result["autorizacion"] = autorizacion
                result["estado"] = "encontrado"
                result["observacion"] = ""
                return result

        result["estado"] = "no_encontrado"
        result["observacion"] = "No se halló patrón de Paciente / Ingreso / Autorización"
        return result

    except Exception as e:
        result["estado"] = "error"
        result["observacion"] = str(e)
        return result


def collect_files(base_folder: str, radicados: list[str], keyword: str, mode: str) -> list[tuple[str, str]]:
    """
    Devuelve lista de tuplas: (radicado, ruta_archivo)
    """
    files_to_process = []

    for rad in radicados:
        rad = rad.strip()
        if not rad:
            continue

        folder_path = os.path.join(base_folder, rad)
        if not os.path.isdir(folder_path):
            continue

        for root, _, files in os.walk(folder_path):
            for fname in files:
                if is_valid_file(fname, mode, keyword):
                    files_to_process.append((rad, os.path.join(root, fname)))

    return files_to_process


def save_results_csv(results: list[dict], output_csv: str) -> None:
    """
    Guarda CSV delimitado por |
    """
    columns = [
        "radicado", "carpeta", "archivo", "pagina",
        "paciente", "ingreso", "autorizacion",
        "estado", "observacion"
    ]
    df = pd.DataFrame(results, columns=columns)
    df.to_csv(output_csv, sep="|", index=False, encoding="utf-8-sig", quoting=csv.QUOTE_MINIMAL)


# =========================================================
# INTERFAZ
# =========================================================

class ExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Extractor Paciente + Ingreso + Autorización - PDF / TIFF")
        self.root.geometry("1100x760")

        self.stop_event = threading.Event()
        self.log_queue = queue.Queue()
        self.results = []
        self.worker_thread = None

        self._build_ui()
        self._poll_log_queue()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)

        # ---------------------------
        # Entrada
        # ---------------------------
        frm_input = ttk.LabelFrame(main, text="Entrada", padding=10)
        frm_input.pack(fill="x", padx=5, pady=5)

        ttk.Label(frm_input, text="Carpeta base:").grid(row=0, column=0, sticky="w")
        self.var_base = tk.StringVar()
        ttk.Entry(frm_input, textvariable=self.var_base, width=90).grid(row=0, column=1, sticky="we", padx=5)
        ttk.Button(frm_input, text="Examinar", command=self.select_base_folder).grid(row=0, column=2, padx=5)

        ttk.Label(frm_input, text="Lista de radicados (uno por línea):").grid(row=1, column=0, sticky="nw", pady=(10, 0))
        self.txt_radicados = tk.Text(frm_input, width=40, height=10)
        self.txt_radicados.grid(row=1, column=1, sticky="we", padx=5, pady=(10, 0))

        side_filters = ttk.Frame(frm_input)
        side_filters.grid(row=1, column=2, sticky="n", padx=5, pady=(10, 0))

        ttk.Label(side_filters, text="Palabra clave archivo:").pack(anchor="w")
        self.var_keyword = tk.StringVar()
        ttk.Entry(side_filters, textvariable=self.var_keyword, width=25).pack(anchor="w", pady=(0, 10))

        ttk.Label(side_filters, text="Tipo de archivo:").pack(anchor="w")
        self.var_filetype = tk.StringVar(value=FILETYPE_BOTH)
        ttk.Combobox(
            side_filters,
            textvariable=self.var_filetype,
            values=[FILETYPE_BOTH, FILETYPE_PDF, FILETYPE_TIFF],
            state="readonly",
            width=22
        ).pack(anchor="w", pady=(0, 10))

        ttk.Label(side_filters, text="Hilos / workers:").pack(anchor="w")
        self.var_workers = tk.StringVar(value=str(max(2, (os.cpu_count() or 4) // 2)))
        workers_frame = ttk.Frame(side_filters)
        workers_frame.pack(anchor="w")
        ttk.Entry(workers_frame, textvariable=self.var_workers, width=8).pack(side="left")
        ttk.Button(workers_frame, text="Auto detectar", command=self.auto_detect_workers).pack(side="left", padx=5)

        frm_input.columnconfigure(1, weight=1)

        # ---------------------------
        # Salida
        # ---------------------------
        frm_output = ttk.LabelFrame(main, text="Salida", padding=10)
        frm_output.pack(fill="x", padx=5, pady=5)

        ttk.Label(frm_output, text="Salida CSV:").grid(row=0, column=0, sticky="w")
        self.var_output = tk.StringVar()
        ttk.Entry(frm_output, textvariable=self.var_output, width=90).grid(row=0, column=1, sticky="we", padx=5)
        ttk.Button(frm_output, text="Ruta", command=self.select_output_file).grid(row=0, column=2, padx=5)

        frm_output.columnconfigure(1, weight=1)

        # ---------------------------
        # Botones
        # ---------------------------
        frm_actions = ttk.Frame(main)
        frm_actions.pack(fill="x", padx=5, pady=5)

        ttk.Button(frm_actions, text="Procesar", command=self.start_processing).pack(side="left", padx=5)
        ttk.Button(frm_actions, text="Detener", command=self.stop_processing).pack(side="left", padx=5)
        ttk.Button(frm_actions, text="Limpiar", command=self.clear_all).pack(side="left", padx=5)

        # ---------------------------
        # Progreso
        # ---------------------------
        frm_progress = ttk.LabelFrame(main, text="Progreso", padding=10)
        frm_progress.pack(fill="x", padx=5, pady=5)

        self.progress = ttk.Progressbar(frm_progress, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", pady=5)

        self.var_status = tk.StringVar(value="Listo")
        ttk.Label(frm_progress, textvariable=self.var_status).pack(anchor="w")

        self.var_summary = tk.StringVar(value="Procesados: 0 | Encontrados: 0 | No encontrados: 0 | Errores: 0")
        ttk.Label(frm_progress, textvariable=self.var_summary).pack(anchor="w", pady=(5, 0))

        # ---------------------------
        # Log
        # ---------------------------
        frm_log = ttk.LabelFrame(main, text="Log / Resultados", padding=10)
        frm_log.pack(fill="both", expand=True, padx=5, pady=5)

        self.txt_log = tk.Text(frm_log, wrap="word")
        self.txt_log.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(frm_log, orient="vertical", command=self.txt_log.yview)
        scroll.pack(side="right", fill="y")
        self.txt_log.configure(yscrollcommand=scroll.set)

    # ---------------------------
    # UI helpers
    # ---------------------------

    def log(self, message: str):
        self.log_queue.put(("log", message))

    def set_status(self, message: str):
        self.log_queue.put(("status", message))

    def set_progress(self, value: int, maximum: int):
        self.log_queue.put(("progress", (value, maximum)))

    def set_summary(self, processed: int, found: int, not_found: int, errors: int):
        self.log_queue.put(("summary", (processed, found, not_found, errors)))

    def _poll_log_queue(self):
        try:
            while True:
                item_type, payload = self.log_queue.get_nowait()

                if item_type == "log":
                    self.txt_log.insert("end", payload + "\n")
                    self.txt_log.see("end")

                elif item_type == "status":
                    self.var_status.set(payload)

                elif item_type == "progress":
                    value, maximum = payload
                    self.progress["maximum"] = max(1, maximum)
                    self.progress["value"] = value

                elif item_type == "summary":
                    processed, found, not_found, errors = payload
                    self.var_summary.set(
                        f"Procesados: {processed} | Encontrados: {found} | No encontrados: {not_found} | Errores: {errors}"
                    )

        except queue.Empty:
            pass

        self.root.after(150, self._poll_log_queue)

    def select_base_folder(self):
        folder = filedialog.askdirectory(title="Selecciona la carpeta base")
        if folder:
            self.var_base.set(folder)

    def select_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="Guardar CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")]
        )
        if file_path:
            self.var_output.set(file_path)

    def auto_detect_workers(self):
        cpu = os.cpu_count() or 4
        suggested = max(2, min(8, cpu // 2 if cpu > 2 else 2))
        self.var_workers.set(str(suggested))
        self.log(f"[INFO] Workers sugeridos: {suggested}")

    def clear_all(self):
        self.txt_log.delete("1.0", "end")
        self.var_status.set("Listo")
        self.var_summary.set("Procesados: 0 | Encontrados: 0 | No encontrados: 0 | Errores: 0")
        self.progress["value"] = 0

    def stop_processing(self):
        self.stop_event.set()
        self.log("[INFO] Se solicitó detener el proceso...")

    # ---------------------------
    # Validación
    # ---------------------------

    def validate_inputs(self):
        base = self.var_base.get().strip()
        output_csv = self.var_output.get().strip()
        radicados = parse_radicados(self.txt_radicados.get("1.0", "end"))

        if not base or not os.path.isdir(base):
            messagebox.showerror("Error", "Debes seleccionar una carpeta base válida.")
            return None

        if not output_csv:
            messagebox.showerror("Error", "Debes indicar la ruta del CSV de salida.")
            return None

        if not radicados:
            messagebox.showerror("Error", "Debes ingresar al menos un radicado.")
            return None

        try:
            workers = int(self.var_workers.get().strip())
            if workers < 1:
                raise ValueError
        except Exception:
            messagebox.showerror("Error", "El número de workers/hilos no es válido.")
            return None

        return {
            "base": base,
            "output_csv": output_csv,
            "radicados": radicados,
            "keyword": self.var_keyword.get().strip(),
            "mode": self.var_filetype.get().strip(),
            "workers": workers,
        }

    # ---------------------------
    # Proceso principal
    # ---------------------------

    def start_processing(self):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("En proceso", "Ya hay un proceso en ejecución.")
            return

        params = self.validate_inputs()
        if not params:
            return

        self.stop_event.clear()
        self.results = []
        self.clear_all()

        self.worker_thread = threading.Thread(
            target=self.run_process,
            args=(params,),
            daemon=True
        )
        self.worker_thread.start()

    def run_process(self, params: dict):
        try:
            base = params["base"]
            output_csv = params["output_csv"]
            radicados = params["radicados"]
            keyword = params["keyword"]
            mode = params["mode"]
            workers = params["workers"]

            self.set_status("Recolectando archivos...")
            files = collect_files(base, radicados, keyword, mode)

            if not files:
                self.log("[INFO] No se encontraron archivos para procesar con esos filtros.")
                self.set_status("Sin archivos para procesar")
                return

            total = len(files)
            self.log(f"[INFO] Total de archivos a procesar: {total}")
            self.set_progress(0, total)

            processed = 0
            found = 0
            not_found = 0
            errors = 0

            with ThreadPoolExecutor(max_workers=workers) as executor:
                future_map = {
                    executor.submit(process_file, file_path, radicado, self.stop_event): (radicado, file_path)
                    for radicado, file_path in files
                }

                for future in as_completed(future_map):
                    if self.stop_event.is_set():
                        self.log("[INFO] Proceso detenido por el usuario.")
                        break

                    radicado, file_path = future_map[future]
                    fname = os.path.basename(file_path)

                    try:
                        result = future.result()
                        self.results.append(result)

                        processed += 1
                        if result["estado"] == "encontrado":
                            found += 1
                            self.log(
                                f"[OK] {result['radicado']} | {result['archivo']} | "
                                f"pag {result['pagina']} | paciente={result['paciente']} | "
                                f"ingreso={result['ingreso']} | autorizacion={result['autorizacion']}"
                            )
                        elif result["estado"] == "no_encontrado":
                            not_found += 1
                            self.log(
                                f"[NO] {result['radicado']} | {result['archivo']} | {result['observacion']}"
                            )
                        elif result["estado"] == "error":
                            errors += 1
                            self.log(
                                f"[ERR] {result['radicado']} | {result['archivo']} | {result['observacion']}"
                            )
                        elif result["estado"] == "detenido":
                            self.log(
                                f"[STOP] {result['radicado']} | {result['archivo']} | detenido"
                            )
                        else:
                            self.log(
                                f"[INFO] {result['radicado']} | {result['archivo']} | {result['estado']}"
                            )

                    except Exception as e:
                        processed += 1
                        errors += 1
                        self.results.append({
                            "radicado": radicado,
                            "carpeta": radicado,
                            "archivo": fname,
                            "pagina": "",
                            "paciente": "",
                            "ingreso": "",
                            "autorizacion": "",
                            "estado": "error",
                            "observacion": str(e),
                        })
                        self.log(f"[ERR] {radicado} | {fname} | {e}")

                    self.set_progress(processed, total)
                    self.set_summary(processed, found, not_found, errors)
                    self.set_status(f"Procesando... {processed}/{total}")

            if self.results:
                save_results_csv(self.results, output_csv)
                self.log(f"[INFO] CSV guardado en: {output_csv}")

            self.set_status("Proceso finalizado")

        except Exception as e:
            self.log("[ERR] Falla general del proceso")
            self.log(str(e))
            self.log(traceback.format_exc())
            self.set_status("Error general")


# =========================================================
# MAIN
# =========================================================

def main():
    root = tk.Tk()
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    app = ExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()