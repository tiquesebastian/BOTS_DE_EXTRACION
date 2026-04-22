from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from fechas import (
    DEFAULT_FILE_KEYWORDS,
    DEFAULT_LABEL_KEYWORDS,
    FechasConfig,
    ProgressInfo,
    SEED_RESULT_PATHS,
    format_duration,
    normalizar_lista_keywords,
    run_fechas,
)


BG = "#03140A"
PANEL = "#062313"
FG = "#7CFF6B"
FG_DIM = "#4ACB5F"
ACCENT = "#B8FFAE"
ENTRY_BG = "#021008"
BTN_BG = "#103B1F"
BTN_ACTIVE = "#185A30"
MONO = ("Consolas", 10)
MONO_BOLD = ("Consolas", 11, "bold")


class AppFechas(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("BOT FECHAS // RETRO TERMINAL")
        self.geometry("1040x720")
        self.configure(bg=BG)

        self.log_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None

        self.root_var = tk.StringVar(value=r"Z:\IA 10\NUEVO\PARTE3\890701715")
        self.radicados_var = tk.StringVar(value=r"c:\Users\ticdesarrollo09\Music\trabajo\bot\faltan estos radicados")
        self.out_csv_var = tk.StringVar(value=r"c:\Users\ticdesarrollo09\Music\trabajo\bot\890701715_fechas_ingreso.csv")
        self.out_excel_var = tk.StringVar(value=r"c:\Users\ticdesarrollo09\Music\trabajo\bot\890701715_fechas_ingreso_detalle.xlsx")
        self.max_text_pages_var = tk.StringVar(value="7")
        self.max_ocr_pages_var = tk.StringVar(value="7")
        self.max_files_var = tk.StringVar(value="12")
        self.file_keywords_var = tk.StringVar(value=",".join(DEFAULT_FILE_KEYWORDS))
        self.label_keywords_var = tk.StringVar(value=",".join(DEFAULT_LABEL_KEYWORDS))
        self.metrics_var = tk.StringVar(value="ENCONTRADOS: 0 | NO ENCONTRADOS: 0 | RESTANTES: 0")
        self.timing_var = tk.StringVar(value="TIEMPO: 00:00 | ETA: 00:00 | VELOCIDAD: SIN DATOS")

        self._build_ui()
        self.after(150, self._drain_logs)

    def _build_ui(self) -> None:
        container = tk.Frame(self, bg=BG, padx=12, pady=12)
        container.pack(fill="both", expand=True)

        header = tk.Frame(container, bg=PANEL, bd=1, relief="solid", highlightbackground=FG_DIM, highlightthickness=1)
        header.grid(row=0, column=0, columnspan=5, sticky="ew", pady=(0, 10))
        tk.Label(
            header,
            text="BOT FECHAS :: MODO EXTRACCION",
            bg=PANEL,
            fg=ACCENT,
            font=("Consolas", 13, "bold"),
            padx=10,
            pady=8,
            anchor="w",
        ).pack(fill="x")

        subtitle = tk.Label(
            container,
            text="Configura rutas y ejecuta. El log se actualiza en vivo.",
            bg=BG,
            fg=FG_DIM,
            font=MONO,
            anchor="w",
        )
        subtitle.grid(row=1, column=0, columnspan=5, sticky="w", pady=(0, 10))

        self._row_path(container, 2, "RUTA RAIZ", self.root_var, self._pick_root)
        self._row_path(container, 3, "ARCHIVO RADICADOS", self.radicados_var, self._pick_radicados)
        self._row_path(container, 4, "SALIDA CSV", self.out_csv_var, self._pick_out_csv)
        self._row_path(container, 5, "SALIDA EXCEL", self.out_excel_var, self._pick_out_excel)

        params = tk.Frame(container, bg=PANEL, bd=1, relief="solid", highlightbackground=FG_DIM, highlightthickness=1, padx=10, pady=8)
        params.grid(row=6, column=0, columnspan=5, sticky="ew", pady=(8, 0))

        tk.Label(params, text="MAX PAGINAS TEXTO PDF", bg=PANEL, fg=FG, font=MONO).grid(row=0, column=0, sticky="w")
        self.max_text_entry = self._retro_entry(params, self.max_text_pages_var, width=8)
        self.max_text_entry.grid(row=0, column=1, sticky="w", padx=(10, 22))

        tk.Label(params, text="MAX PAGINAS OCR", bg=PANEL, fg=FG, font=MONO).grid(row=0, column=2, sticky="w")
        self.max_ocr_entry = self._retro_entry(params, self.max_ocr_pages_var, width=8)
        self.max_ocr_entry.grid(row=0, column=3, sticky="w", padx=(10, 22))

        tk.Label(params, text="MAX ARCHIVOS/RAD", bg=PANEL, fg=FG, font=MONO).grid(row=0, column=4, sticky="w")
        self.max_files_entry = self._retro_entry(params, self.max_files_var, width=8)
        self.max_files_entry.grid(row=0, column=5, sticky="w", padx=(10, 0))

        tk.Label(params, text="KEYWORDS ARCHIVO (min 3)", bg=PANEL, fg=FG, font=MONO).grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.file_keywords_entry = self._retro_entry(params, self.file_keywords_var)
        self.file_keywords_entry.grid(row=1, column=1, columnspan=5, sticky="ew", padx=(10, 0), pady=(10, 0))

        tk.Label(params, text="KEYWORDS DATO", bg=PANEL, fg=FG, font=MONO).grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.label_keywords_entry = self._retro_entry(params, self.label_keywords_var)
        self.label_keywords_entry.grid(row=2, column=1, columnspan=5, sticky="ew", padx=(10, 0), pady=(8, 0))

        buttons = tk.Frame(container, bg=BG)
        buttons.grid(row=7, column=0, columnspan=5, sticky="w", pady=(12, 0))

        self.run_btn = tk.Button(
            buttons,
            text="EJECUTAR",
            command=self._run,
            bg=BTN_BG,
            fg=ACCENT,
            activebackground=BTN_ACTIVE,
            activeforeground=ACCENT,
            font=MONO_BOLD,
            relief="flat",
            bd=1,
            padx=16,
            pady=6,
            highlightbackground=FG_DIM,
            highlightthickness=1,
            cursor="hand2",
        )
        self.run_btn.pack(side="left")

        self.clear_btn = tk.Button(
            buttons,
            text="LIMPIAR LOG",
            command=self._clear_log,
            bg=BTN_BG,
            fg=FG,
            activebackground=BTN_ACTIVE,
            activeforeground=ACCENT,
            font=MONO,
            relief="flat",
            bd=1,
            padx=14,
            pady=6,
            highlightbackground=FG_DIM,
            highlightthickness=1,
            cursor="hand2",
        )
        self.clear_btn.pack(side="left", padx=(10, 0))

        metrics = tk.Frame(container, bg=PANEL, bd=1, relief="solid", highlightbackground=FG_DIM, highlightthickness=1, padx=10, pady=8)
        metrics.grid(row=8, column=0, columnspan=5, sticky="ew", pady=(12, 0))

        tk.Label(metrics, textvariable=self.metrics_var, bg=PANEL, fg=ACCENT, font=MONO_BOLD, anchor="w").grid(row=0, column=0, sticky="ew")
        tk.Label(metrics, textvariable=self.timing_var, bg=PANEL, fg=FG, font=MONO, anchor="w").grid(row=1, column=0, sticky="ew", pady=(6, 0))

        tk.Label(container, text="LOG DE EJECUCION", bg=BG, fg=FG, font=MONO_BOLD, anchor="w").grid(
            row=9, column=0, columnspan=5, sticky="w", pady=(14, 6)
        )

        self.log_text = tk.Text(
            container,
            wrap="word",
            height=22,
            bg=ENTRY_BG,
            fg=FG,
            insertbackground=ACCENT,
            font=MONO,
            bd=1,
            relief="solid",
            highlightbackground=FG_DIM,
            highlightthickness=1,
            padx=8,
            pady=8,
        )
        self.log_text.grid(row=10, column=0, columnspan=4, sticky="nsew")

        scroll = tk.Scrollbar(container, orient="vertical", command=self.log_text.yview, bg=PANEL, troughcolor=ENTRY_BG)
        scroll.grid(row=10, column=4, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)

        self.status_var = tk.StringVar(value="STATUS: IDLE")
        status = tk.Label(
            container,
            textvariable=self.status_var,
            bg=PANEL,
            fg=FG,
            font=MONO,
            anchor="w",
            padx=8,
            pady=5,
            bd=1,
            relief="solid",
            highlightbackground=FG_DIM,
            highlightthickness=1,
        )
        status.grid(row=11, column=0, columnspan=5, sticky="ew", pady=(8, 0))

        container.columnconfigure(1, weight=1)
        container.columnconfigure(3, weight=1)
        container.rowconfigure(10, weight=1)

    def _retro_entry(self, parent: tk.Widget, var: tk.StringVar, width: int | None = None) -> tk.Entry:
        return tk.Entry(
            parent,
            textvariable=var,
            width=width,
            bg=ENTRY_BG,
            fg=FG,
            insertbackground=ACCENT,
            font=MONO,
            relief="solid",
            bd=1,
            highlightbackground=FG_DIM,
            highlightthickness=1,
        )

    def _retro_button(self, parent: tk.Widget, text: str, command) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=BTN_BG,
            fg=FG,
            activebackground=BTN_ACTIVE,
            activeforeground=ACCENT,
            font=MONO,
            relief="flat",
            bd=1,
            highlightbackground=FG_DIM,
            highlightthickness=1,
            padx=10,
            pady=4,
            cursor="hand2",
        )

    def _row_path(self, parent: tk.Frame, row: int, label: str, var: tk.StringVar, browse_cmd) -> None:
        tk.Label(parent, text=label, bg=BG, fg=FG, font=MONO_BOLD).grid(row=row, column=0, sticky="w", pady=5)
        entry = self._retro_entry(parent, var)
        entry.grid(row=row, column=1, columnspan=3, sticky="ew", pady=5, padx=(8, 8))
        self._retro_button(parent, "BUSCAR", browse_cmd).grid(row=row, column=4, sticky="e", pady=5)

    def _pick_root(self) -> None:
        path = filedialog.askdirectory(title="Selecciona la ruta raiz")
        if path:
            self.root_var.set(path)

    def _pick_radicados(self) -> None:
        path = filedialog.askopenfilename(title="Selecciona archivo de radicados")
        if path:
            self.radicados_var.set(path)

    def _pick_out_csv(self) -> None:
        path = filedialog.asksaveasfilename(title="Guardar CSV", defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if path:
            self.out_csv_var.set(path)

    def _pick_out_excel(self) -> None:
        path = filedialog.asksaveasfilename(title="Guardar Excel", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.out_excel_var.set(path)

    def _clear_log(self) -> None:
        self.log_text.delete("1.0", tk.END)
        self.metrics_var.set("ENCONTRADOS: 0 | NO ENCONTRADOS: 0 | RESTANTES: 0")
        self.timing_var.set("TIEMPO: 00:00 | ETA: 00:00 | VELOCIDAD: SIN DATOS")
        self.status_var.set("STATUS: LOG LIMPIO")

    def _append_log(self, text: str) -> None:
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)

    def _apply_progress(self, info: ProgressInfo) -> None:
        self.metrics_var.set(
            f"ENCONTRADOS: {info.encontrados_exactos} | NO ENCONTRADOS: {info.no_encontrados} | RESTANTES: {info.restantes_ocr}"
        )
        self.timing_var.set(
            f"TIEMPO: {format_duration(info.elapsed_seconds)} | ETA: {format_duration(info.eta_seconds)} | VELOCIDAD: {info.speed_label.upper()}"
        )
        if info.current_radicado:
            self.status_var.set(
                f"STATUS: RUNNING | RADICADO: {info.current_radicado} | OCR {info.procesados_ocr}/{info.total_pendientes_ocr}"
            )

    def _drain_logs(self) -> None:
        while True:
            try:
                kind, payload = self.log_queue.get_nowait()
            except queue.Empty:
                break
            if kind == "log":
                self._append_log(str(payload))
            elif kind == "progress" and isinstance(payload, ProgressInfo):
                self._apply_progress(payload)
        self.after(150, self._drain_logs)

    def _run(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("En ejecucion", "Ya hay una ejecucion en curso")
            return

        try:
            file_keywords = normalizar_lista_keywords(self.file_keywords_var.get().strip())
            if len(file_keywords) < 3:
                raise ValueError("Debes ingresar minimo 3 keywords de archivo")
            label_keywords = normalizar_lista_keywords(self.label_keywords_var.get().strip()) or DEFAULT_LABEL_KEYWORDS

            config = FechasConfig(
                root_path=Path(self.root_var.get().strip()),
                target_radicados_path=Path(self.radicados_var.get().strip()),
                output_csv_path=Path(self.out_csv_var.get().strip()),
                output_excel_path=Path(self.out_excel_var.get().strip()),
                seed_result_paths=[Path(self.out_csv_var.get().strip())] + list(SEED_RESULT_PATHS),
                max_pdf_text_pages=max(1, int(self.max_text_pages_var.get().strip())),
                max_ocr_pages=max(1, int(self.max_ocr_pages_var.get().strip())),
                max_files_per_radicado=max(1, int(self.max_files_var.get().strip())),
                file_name_keywords=file_keywords,
                target_text_keywords=label_keywords,
            )
        except Exception as error:
            messagebox.showerror("Error de parametros", str(error))
            return

        self.run_btn.configure(state="disabled")
        self.status_var.set("STATUS: RUNNING")
        self.log_queue.put(("log", "Iniciando extraccion de fechas..."))

        def emit_log(message: str) -> None:
            self.log_queue.put(("log", message))

        def emit_progress(info: ProgressInfo) -> None:
            self.log_queue.put(("progress", info))

        def worker() -> None:
            try:
                run_fechas(config, on_log=emit_log, on_progress=emit_progress)
                self.log_queue.put(("log", "Proceso finalizado con exito"))
                self.after(0, lambda: self.status_var.set("STATUS: DONE"))
            except Exception as error:
                self.log_queue.put(("log", f"ERROR: {error}"))
                self.after(0, lambda: self.status_var.set("STATUS: ERROR"))
            finally:
                self.after(0, lambda: self.run_btn.configure(state="normal"))

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()


if __name__ == "__main__":
    app = AppFechas()
    app.mainloop()
