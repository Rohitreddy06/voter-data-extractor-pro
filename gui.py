from __future__ import annotations
"""
gui.py
======
CustomTkinter desktop GUI for Voter Data Extractor Pro.

All PDF/OCR/Excel work happens on background threads (via
ThreadPoolExecutor) so the Tk main loop never freezes. Worker threads
communicate progress back to the GUI through a thread-safe queue that
is drained on a `root.after()` polling loop — this is the standard
safe pattern for Tkinter, since Tk widgets must only be touched from
the main thread.
"""


import queue
import re
import threading
import traceback
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Optional

import customtkinter as ctk
from tkinter import filedialog, messagebox

import config
import utils
from database import VoterDatabase
from extractor import PDFExtractor, ExtractionError, EncryptedPDFError, CorruptPDFError
from excel_writer import ExcelDataError, run_matching_pipeline, ExcelColumnError
from translator import translate_excel_file
from utils import logger, remember_last_folder, get_last_folder, load_resume_state

# Optional drag-and-drop support.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _DND_AVAILABLE = True
except ImportError:
    _DND_AVAILABLE = False


ctk.set_appearance_mode(config.THEME_MODE)
ctk.set_default_color_theme(config.THEME_COLOR)


class _CTkDnD(ctk.CTk):
    """CTk root that also behaves like a TkinterDnD.Tk when tkinterdnd2
    is installed, enabling drag-and-drop of PDF/Excel files onto the
    window. Falls back to a plain CTk root otherwise."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._dnd_enabled = False
        if _DND_AVAILABLE:
            try:
                self.TkdndVersion = TkinterDnD._require(self)
                self._dnd_enabled = hasattr(self, "drop_target_register") and hasattr(self, "dnd_bind")
            except Exception as exc:
                logger.warning("Drag-and-drop init failed, continuing without it: %s", exc)


class VoterExtractorApp(_CTkDnD):
    def __init__(self) -> None:
        super().__init__()
        self.title(config.APP_TITLE)
        self.geometry(config.APP_GEOMETRY)
        self.minsize(900, 650)

        self.pdf_path: Optional[str] = None
        self.excel_path: Optional[str] = None
        self.output_folder: str = get_last_folder()
        self.current_pdf_base: str = ""
        self.last_output_path: Optional[str] = None
        self.db: Optional[VoterDatabase] = None

        self.cancel_event = threading.Event()
        self.translate_spinner_active = False
        self.executor = ThreadPoolExecutor(max_workers=2)
        self.msg_queue: "queue.Queue[tuple]" = queue.Queue()
        self.stopwatch = utils.Stopwatch()

        self._build_layout()
        self._poll_queue()

    # ------------------------------------------------------------------ layout
    def _build_layout(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        self._build_file_selection_frame()
        self._build_options_frame()
        self._build_action_frame()
        self._build_progress_frame()
        self._build_log_frame()
        self._build_search_frame()

    def _build_file_selection_frame(self) -> None:
        frame = ctk.CTkFrame(self, fg_color="#1f2433", corner_radius=16)
        frame.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))
        frame.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(frame, text="Select PDF", command=self._select_pdf, width=140).grid(
            row=0, column=0, padx=8, pady=8
        )
        self.pdf_label = ctk.CTkLabel(frame, text="No PDF selected", anchor="w")
        self.pdf_label.grid(row=0, column=1, sticky="ew", padx=8)

        ctk.CTkButton(frame, text="Select Excel", command=self._select_excel, width=140).grid(
            row=1, column=0, padx=8, pady=8
        )
        self.excel_label = ctk.CTkLabel(frame, text="No Excel selected", anchor="w")
        self.excel_label.grid(row=1, column=1, sticky="ew", padx=8)

        ctk.CTkButton(frame, text="Output Folder", command=self._select_output_folder, width=140).grid(
            row=2, column=0, padx=8, pady=8
        )
        self.output_label = ctk.CTkLabel(frame, text=self.output_folder, anchor="w")
        self.output_label.grid(row=2, column=1, sticky="ew", padx=8)

        ctk.CTkLabel(frame, text="Output Filename Prefix:").grid(row=3, column=0, padx=8, pady=(8, 8), sticky="w")
        self.output_prefix_var = ctk.StringVar(value="")
        self.output_prefix_entry = ctk.CTkEntry(frame, textvariable=self.output_prefix_var, width=360)
        self.output_prefix_entry.grid(row=3, column=1, padx=8, pady=(8, 8), sticky="ew")

        if getattr(self, "_dnd_enabled", False):
            try:
                self.drop_target_register(DND_FILES)
                self.dnd_bind("<<Drop>>", self._on_drop)
            except Exception as exc:
                logger.warning("Drag-and-drop feature unavailable, continuing without it: %s", exc)

    def _build_options_frame(self) -> None:
        frame = ctk.CTkFrame(self, fg_color="#1f2433", corner_radius=16)
        frame.grid(row=1, column=0, sticky="ew", padx=16, pady=8)

        self.translate_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(frame, text="Translate Telugu -> English", variable=self.translate_var).pack(
            side="left", padx=12, pady=10
        )

        self.ocr_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(frame, text="Use OCR for scanned pages", variable=self.ocr_var).pack(
            side="left", padx=12, pady=10
        )

    def _build_action_frame(self) -> None:
        frame = ctk.CTkFrame(self, fg_color="#1f2433", corner_radius=16)
        frame.grid(row=2, column=0, sticky="ew", padx=16, pady=8)

        self.start_button = ctk.CTkButton(
            frame, text="Start", fg_color="#1f8a3d", hover_color="#166b2f",
            command=self._on_start_clicked, width=140,
        )
        self.start_button.pack(side="left", padx=8, pady=10)

        self.cancel_button = ctk.CTkButton(
            frame, text="Cancel", fg_color="#a33", hover_color="#822",
            command=self._on_cancel_clicked, width=140, state="disabled",
        )
        self.cancel_button.pack(side="left", padx=8, pady=10)

        self.translate_button = ctk.CTkButton(
            frame, text="Translate to English", fg_color="#3f6ad8", hover_color="#3454b3",
            command=self._on_translate_clicked, width=180, state="disabled",
        )
        self.translate_button.pack(side="left", padx=8, pady=10)

        self.translate_state_label = ctk.CTkLabel(frame, text="", anchor="w")
        self.translate_state_label.pack(side="left", padx=(16, 4))

        self.status_label = ctk.CTkLabel(frame, text="Idle", anchor="w")
        self.status_label.pack(side="left", padx=16)

    def _build_progress_frame(self) -> None:
        frame = ctk.CTkFrame(self, fg_color="#1f2433", corner_radius=16)
        frame.grid(row=3, column=0, sticky="nsew", padx=16, pady=8)
        frame.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(frame)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))

        stats_frame = ctk.CTkFrame(frame, fg_color="transparent")
        stats_frame.grid(row=1, column=0, sticky="ew", padx=8, pady=4)

        self.elapsed_label = ctk.CTkLabel(stats_frame, text="Elapsed: 00:00")
        self.elapsed_label.pack(side="left", padx=8)
        self.found_label = ctk.CTkLabel(stats_frame, text="Records found: 0")
        self.found_label.pack(side="left", padx=8)
        self.matched_label = ctk.CTkLabel(stats_frame, text="Matched: 0")
        self.matched_label.pack(side="left", padx=8)
        self.remaining_label = ctk.CTkLabel(stats_frame, text="Remaining: 0")
        self.remaining_label.pack(side="left", padx=8)

    def _build_log_frame(self) -> None:
        frame = ctk.CTkFrame(self, fg_color="#1f2433", corner_radius=16)
        frame.grid(row=4, column=0, sticky="nsew", padx=16, pady=8)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=2)

        self.log_box = ctk.CTkTextbox(frame, height=180, state="disabled")
        self.log_box.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

    def _build_search_frame(self) -> None:
        frame = ctk.CTkFrame(self, fg_color="#1f2433", corner_radius=16)
        frame.grid(row=5, column=0, sticky="ew", padx=16, pady=(8, 16))

        ctk.CTkLabel(frame, text="Search:").pack(side="left", padx=(8, 4))
        self.search_entry = ctk.CTkEntry(frame, placeholder_text="EPIC / Name / House No", width=250)
        self.search_entry.pack(side="left", padx=4)

        self.search_mode = ctk.StringVar(value="EPIC")
        ctk.CTkOptionMenu(frame, values=["EPIC", "Name", "House No"], variable=self.search_mode).pack(
            side="left", padx=4
        )
        ctk.CTkButton(frame, text="Search", width=100, command=self._on_search_clicked).pack(
            side="left", padx=8
        )
        ctk.CTkButton(frame, text="Export CSV", width=110, command=self._on_export_csv).pack(
            side="left", padx=4
        )
        ctk.CTkButton(frame, text="Export DB Copy", width=130, command=self._on_export_db).pack(
            side="left", padx=4
        )

    # ------------------------------------------------------------------ file selection
    def _select_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Electoral Roll PDF",
            initialdir=self.output_folder,
            filetypes=[("PDF files", "*.pdf")],
        )
        if path:
            self._set_pdf(path)

    def _select_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Excel Template",
            initialdir=self.output_folder,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self._set_excel(path)

    def _select_output_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select Output Folder", initialdir=self.output_folder)
        if folder:
            self.output_folder = folder
            self.output_label.configure(text=folder)
            remember_last_folder(folder)

    def _on_drop(self, event) -> None:
        paths = self.tk.splitlist(event.data)
        for p in paths:
            if p.lower().endswith(".pdf"):
                self._set_pdf(p)
            elif p.lower().endswith(".xlsx"):
                self._set_excel(p)

    def _set_pdf(self, path: str) -> None:
        self.pdf_path = path
        self.pdf_label.configure(text=path)
        self.current_pdf_base = Path(path).stem
        if not self.output_prefix_var.get() or self.output_prefix_var.get() == self.current_pdf_base:
            self.output_prefix_var.set(self.current_pdf_base)
        remember_last_folder(str(Path(path).parent))

    def _set_excel(self, path: str) -> None:
        self.excel_path = path
        self.excel_label.configure(text=path)

    # ------------------------------------------------------------------ logging helpers
    def _log(self, message: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _set_status(self, text: str) -> None:
        self.status_label.configure(text=text)

    # ------------------------------------------------------------------ start / cancel
    def _on_start_clicked(self) -> None:
        if not self.pdf_path:
            messagebox.showerror("Missing PDF", "Please select an Electoral Roll PDF first.")
            return
        if not self.excel_path:
            messagebox.showerror("Missing Excel", "Please select an Excel template first.")
            return

        resume_state = load_resume_state(self.pdf_path)
        resume_from = 0
        if resume_state:
            answer = messagebox.askyesno(
                "Resume previous run?",
                f"A previous run stopped at page {resume_state['last_page_completed']}. "
                f"Resume from there instead of starting over?",
            )
            if answer:
                resume_from = resume_state["last_page_completed"]

        self.cancel_event.clear()
        self.translate_button.configure(state="disabled")
        self.translate_state_label.configure(text="")
        self.start_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.progress_bar.set(0)
        self.stopwatch.start()
        self._set_status("Processing...")
        self._log(f"Starting extraction from '{self.pdf_path}'")

        self.executor.submit(self._run_pipeline, resume_from)
        self._tick_elapsed()

    def _on_cancel_clicked(self) -> None:
        self.cancel_event.set()
        self._set_status("Cancelling...")
        self._log("Cancel requested by user. Finishing current page then stopping.")

    def _tick_elapsed(self) -> None:
        if self.start_button.cget("state") == "disabled":
            self.elapsed_label.configure(text=f"Elapsed: {self.stopwatch.elapsed_str()}")
            self.after(1000, self._tick_elapsed)

    # ------------------------------------------------------------------ background pipeline
    def _sanitize_filename(self, raw: str) -> str:
        cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", raw or "").strip("_. -")
        return cleaned or self.current_pdf_base or "Completed"

    def _run_pipeline(self, resume_from: int) -> None:
        """Runs entirely on a worker thread. Never touches Tk widgets
        directly — everything goes through self.msg_queue."""
        try:
            self.db = VoterDatabase()

            def extraction_progress(cur: int, total: int, msg: str) -> None:
                self.msg_queue.put(("extract_progress", cur, total, msg))

            extractor = PDFExtractor(
                pdf_path=self.pdf_path,
                db=self.db,
                use_ocr=self.ocr_var.get(),
                translate_telugu=self.translate_var.get(),
                progress_callback=extraction_progress,
                cancel_event=self.cancel_event,
            )
            self.msg_queue.put(("log", f"PDF opened. {extractor.page_count} pages detected."))
            extract_stats = extractor.run(resume_from_page=resume_from)
            extractor.close()

            if self.cancel_event.is_set():
                self.msg_queue.put(("cancelled", extract_stats))
                return

            self.msg_queue.put(("log", "Extraction complete. Starting Excel matching..."))

            def match_progress(cur: int, total: int) -> None:
                self.msg_queue.put(("match_progress", cur, total))

            output_name = self._sanitize_filename(self.output_prefix_var.get()) + ".xlsx"
            out_path, match_stats = run_matching_pipeline(
                self.excel_path,
                self.db,
                self.output_folder,
                output_filename=output_name,
                progress_callback=match_progress,
            )

            self.msg_queue.put(("done", extract_stats, match_stats, out_path))

        except (EncryptedPDFError, CorruptPDFError, ExtractionError, ExcelColumnError, ExcelDataError) as exc:
            self.msg_queue.put(("error", str(exc)))
        except Exception as exc:
            logger.error("Unexpected pipeline failure: %s\n%s", exc, traceback.format_exc())
            self.msg_queue.put(("error", f"Unexpected error: {exc}"))

    # ------------------------------------------------------------------ queue polling (main thread)
    def _poll_queue(self) -> None:
        try:
            while True:
                item = self.msg_queue.get_nowait()
                self._handle_queue_item(item)
        except queue.Empty:
            pass
        self.after(150, self._poll_queue)

    def _handle_queue_item(self, item: tuple) -> None:
        kind = item[0]

        if kind == "log":
            self._log(item[1])

        elif kind == "extract_progress":
            _, cur, total, msg = item
            self.progress_bar.set(cur / total if total else 0)
            self.found_label.configure(text=f"Records found: {self.db.count() if self.db else 0}")
            self._set_status(msg)

        elif kind == "match_progress":
            _, cur, total = item
            self.progress_bar.set(cur / total if total else 0)
            self.remaining_label.configure(text=f"Remaining: {max(0, total - cur)}")
            self._set_status(f"Matching Excel rows: {cur}/{total}")

        elif kind == "cancelled":
            self.stopwatch.stop()
            self._set_status("Cancelled")
            self._log("Processing cancelled. Progress saved — you can resume next run.")
            self._reset_buttons()

        elif kind == "error":
            self.stopwatch.stop()
            self._set_status("Error")
            self._log(f"ERROR: {item[1]}")
            messagebox.showerror("Processing Error", item[1])
            self._reset_buttons()

        elif kind == "translate_done":
            self.translate_spinner_active = False
            self.translate_state_label.configure(text="Translation complete")
            self.translate_button.configure(state="normal")
            self._log(f"Translation finished and saved to: {item[1]}")
            messagebox.showinfo("Translation Complete", f"Translated workbook saved to:\n{item[1]}")

        elif kind == "translate_error":
            self.translate_spinner_active = False
            self.translate_state_label.configure(text="Translation failed")
            self.translate_button.configure(state="normal")
            self._log(f"Translation ERROR: {item[1]}")
            messagebox.showerror("Translation Error", item[1])

        elif kind == "done":
            _, extract_stats, match_stats, out_path = item
            self.last_output_path = out_path
            self.translate_button.configure(state="normal")
            self.stopwatch.stop()
            self._set_status("Completed")
            self.matched_label.configure(text=f"Matched: {match_stats.matched}")
            self.remaining_label.configure(text="Remaining: 0")
            self._log(
                f"DONE. Extracted {extract_stats['total_records']} voter records from "
                f"{extract_stats['pages_processed']} pages. "
                f"Matched {match_stats.matched}/{match_stats.total} Excel rows "
                f"({match_stats.not_found} not found, {match_stats.duplicates} duplicates). "
                f"Saved: {out_path}"
            )
            messagebox.showinfo("Completed", f"Processing finished.\nSaved to:\n{out_path}")
            self._reset_buttons()

    def _reset_buttons(self) -> None:
        self.start_button.configure(state="normal")
        self.cancel_button.configure(state="disabled")

    def _on_translate_clicked(self) -> None:
        if not self.last_output_path:
            messagebox.showwarning("No output file", "Run the extraction first to produce an output workbook.")
            return

        self.translate_button.configure(state="disabled")
        self.translate_state_label.configure(text="Translating...")
        self.translate_spinner_active = True
        self._animate_translate_spinner(0)
        self.executor.submit(self._run_translation, self.last_output_path)

    def _run_translation(self, input_path: str) -> None:
        try:
            output_path = Path(input_path).with_name(
                f"{Path(input_path).stem}_English.xlsx"
            )
            translated_path = translate_excel_file(input_path, str(output_path))
            self.msg_queue.put(("translate_done", translated_path))
        except Exception as exc:
            self.msg_queue.put(("translate_error", str(exc)))

    def _animate_translate_spinner(self, step: int) -> None:
        if not self.translate_spinner_active:
            return
        dots = (step % 4) * "."
        self.translate_state_label.configure(text=f"Translating{dots}")
        self.after(400, lambda: self._animate_translate_spinner(step + 1))

    # ------------------------------------------------------------------ search / export
    def _on_search_clicked(self) -> None:
        if not self.db:
            messagebox.showwarning("No database yet", "Run extraction first, or it will be empty.")
            return
        query = self.search_entry.get().strip()
        if not query:
            return

        mode = self.search_mode.get()
        if mode == "EPIC":
            rows = self.db.find_by_epic(query)
        elif mode == "Name":
            rows = self.db.search_by_name(query)
        else:
            rows = self.db.search_by_house_no(query)

        if not rows:
            self._log(f"Search '{query}' ({mode}): no results.")
            return

        self._log(f"Search '{query}' ({mode}): {len(rows)} result(s):")
        for r in rows[:20]:
            self._log(
                f"  EPIC={r['EPIC']} Name={r['Name']} House={r['HouseNo']} "
                f"Area={r['Area']} Page={r['PageNo']}"
            )

    def _on_export_csv(self) -> None:
        if not self.db:
            messagebox.showwarning("No database yet", "Run extraction first.")
            return
        out = Path(self.output_folder) / "voters_export.csv"
        self.db.export_csv(out)
        self._log(f"Exported CSV to {out}")

    def _on_export_db(self) -> None:
        if not self.db:
            messagebox.showwarning("No database yet", "Run extraction first.")
            return
        out = Path(self.output_folder) / "voters_export.db"
        self.db.export_sqlite_copy(out)
        self._log(f"Exported database copy to {out}")

    # ------------------------------------------------------------------ lifecycle
    def on_close(self) -> None:
        self.cancel_event.set()
        self.executor.shutdown(wait=False, cancel_futures=True)
        self.destroy()
