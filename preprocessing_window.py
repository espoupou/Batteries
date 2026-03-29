
import tkinter as tk
from pathlib import Path
from datetime import datetime
from tkinter import filedialog, messagebox

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from eis_core import (
    preprocess_eclab_text_file,
    preprocess_eclab_text_to_dataframe,
    remove_warburg_wo, is_excel_file, read_eis_excel_workbook, preprocess_eis_dataframe, export_processed_eis_workbook,
    export_processed_sheets_over_original_workbook,
)


class EISPreprocessingWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent_app = parent

        self.title("EIS Preprocessing Studio")
        self.geometry("1400x850")
        self.minsize(1200, 720)

        self.input_path = tk.StringVar(value="")
        self.output_path = tk.StringVar(value="")

        # paramètres conversion texte
        self.skip_first_line = tk.BooleanVar(value=True)
        self.replace_comma = tk.BooleanVar(value=True)
        self.keep_n_cols = tk.IntVar(value=3)
        self.negate_col3 = tk.BooleanVar(value=False)

        self.use_stop_first_col = tk.BooleanVar(value=False)
        self.stop_first_col = tk.StringVar(value="")

        # paramètres Wo
        self.apply_wo = tk.BooleanVar(value=False)
        self.wor = tk.StringVar(value="")
        self.wot = tk.StringVar(value="")
        self.wop = tk.StringVar(value="")
        self.use_fmin = tk.BooleanVar(value=False)
        self.fmin = tk.StringVar(value="")
        self.trim_pos_im = tk.BooleanVar(value=False)

        self.raw_df = None
        self.proc_df = None

        self.excel_sheets = {}
        self.sheet_names = []
        self.current_sheet_index = 0
        self.current_sheet_name = None
        self.sheet_nav_var = tk.StringVar(value="Text / single file")

        self.sheet_cache = {}  # sheet_name -> {"raw_df", "proc_df", "params", "updated_at"}
        self.sheet_saved_params = {}
        self.cache_workbook_path = None

        self._build_ui()
        self._update_sheet_nav_state()
        self.lift()
        self.focus_force()

    def _build_ui(self):
        main = tk.Frame(self)
        main.pack(fill="both", expand=True, padx=10, pady=10)

        main.columnconfigure(0, weight=0)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(0, weight=1)

        # =============== LEFT PANEL ===============
        left = tk.LabelFrame(main, text="Controls")
        left.grid(row=0, column=0, sticky="nsw", padx=(0, 8), pady=0)

        # fichier entrée
        tk.Label(left, text="Input file").grid(row=0, column=0, sticky="w", padx=8, pady=(8, 4))
        tk.Entry(left, textvariable=self.input_path, width=42, state="readonly").grid(
            row=1, column=0, columnspan=3, sticky="we", padx=8
        )
        tk.Button(left, text="Choose file...", command=self.choose_input_file).grid(
            row=2, column=0, columnspan=3, sticky="we", padx=8, pady=(4, 8)
        )

        nav = tk.LabelFrame(left, text="Workbook navigation")
        nav.grid(row=3, column=0, columnspan=3, sticky="we", padx=8, pady=6)
        nav.columnconfigure(1, weight=1)

        self.prev_sheet_btn = tk.Button(nav, text="←", width=4, command=self.go_to_previous_sheet)
        self.prev_sheet_btn.grid(row=0, column=0, padx=(8, 4), pady=8)

        tk.Label(nav, textvariable=self.sheet_nav_var, anchor="w", justify="left").grid(
            row=0, column=1, sticky="we", padx=4, pady=8
        )

        self.next_sheet_btn = tk.Button(nav, text="→", width=4, command=self.go_to_next_sheet)
        self.next_sheet_btn.grid(row=0, column=2, padx=(4, 8), pady=8)

        # conversion simple
        block1 = tk.LabelFrame(left, text="Text preprocessing")
        block1.grid(row=4, column=0, columnspan=3, sticky="we", padx=8, pady=6)
        block1.columnconfigure(0, weight=0)
        block1.columnconfigure(1, weight=0)
        block1.columnconfigure(2, weight=0)

        tk.Checkbutton(block1, text="Skip first line", variable=self.skip_first_line).grid(
            row=0, column=0, sticky="w", padx=8, pady=4
        )
        tk.Checkbutton(block1, text="Replace comma by dot", variable=self.replace_comma).grid(
            row=0, column=1, columnspan=2, sticky="w", padx=8, pady=4
        )

        tk.Checkbutton(block1, text="Negate 3rd column", variable=self.negate_col3).grid(
            row=1, column=0, sticky="w", padx=8, pady=4
        )
        tk.Label(block1, text="Keep first N columns").grid(
            row=1, column=1, sticky="w", padx=(16, 8), pady=4
        )
        tk.Entry(block1, textvariable=self.keep_n_cols, width=8).grid(
            row=1, column=2, sticky="w", padx=8, pady=4
        )

        tk.Checkbutton(
            block1,
            text="Min freq",
            variable=self.use_stop_first_col
        ).grid(row=2, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 4))
        tk.Entry(block1, textvariable=self.stop_first_col, width=12).grid(
            row=2, column=2, sticky="w", padx=8, pady=(8, 4)
        )

        # Wo remove
        block2 = tk.LabelFrame(left, text="Wo diffusion removal")
        block2.grid(row=5, column=0, columnspan=3, sticky="we", padx=8, pady=6)
        block2.columnconfigure(0, weight=0)
        block2.columnconfigure(1, weight=0)
        block2.columnconfigure(2, weight=0)
        block2.columnconfigure(3, weight=0)

        tk.Checkbutton(block2, text="Apply Wo removal", variable=self.apply_wo).grid(
            row=0, column=0, columnspan=4, sticky="w", padx=8, pady=4
        )

        tk.Label(block2, text="Wo-R").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        tk.Entry(block2, textvariable=self.wor, width=10).grid(row=1, column=1, sticky="w", padx=8, pady=4)

        tk.Label(block2, text="Wo-T").grid(row=1, column=2, sticky="w", padx=(16, 8), pady=4)
        tk.Entry(block2, textvariable=self.wot, width=10).grid(row=1, column=3, sticky="w", padx=8, pady=4)

        tk.Label(block2, text="Wo-P").grid(row=2, column=0, sticky="w", padx=8, pady=4)
        tk.Entry(block2, textvariable=self.wop, width=10).grid(row=2, column=1, sticky="w", padx=8, pady=4)

        tk.Checkbutton(block2, text="Use minimum frequency", variable=self.use_fmin).grid(
            row=3, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 4)
        )
        tk.Entry(block2, textvariable=self.fmin, width=10).grid(
            row=3, column=2, sticky="w", padx=8, pady=(8, 4)
        )

        tk.Checkbutton(block2, text="Drop corrected points with Im > 0", variable=self.trim_pos_im).grid(
            row=4, column=0, columnspan=4, sticky="w", padx=8, pady=4
        )

        # actions
        block3 = tk.LabelFrame(left, text="Actions")
        block3.grid(row=6, column=0, columnspan=3, sticky="we", padx=8, pady=6)
        block3.columnconfigure(0, weight=1)
        block3.columnconfigure(1, weight=1)

        tk.Button(block3, text="Preview", command=self.preview_processing).grid(
            row=0, column=0, sticky="we", padx=(8, 4), pady=(8, 4)
        )
        tk.Button(block3, text="Update current sheet", command=self.update_current_sheet_cache).grid(
            row=0, column=1, sticky="we", padx=(4, 8), pady=(8, 4)
        )

        tk.Button(block3, text="Export processed file...", command=self.export_processed_file).grid(
            row=1, column=0, columnspan=2, sticky="we", padx=8, pady=4
        )
        tk.Button(block3, text="Close", command=self.destroy).grid(
            row=2, column=0, columnspan=2, sticky="we", padx=8, pady=(4, 8)
        )

        # infos
        block4 = tk.LabelFrame(left, text="Info")
        block4.grid(row=7, column=0, columnspan=3, sticky="nsew", padx=8, pady=6)

        self.info_text = tk.Text(block4, width=42, height=16, wrap="word")
        self.info_text.pack(fill="both", expand=True, padx=8, pady=8)

        # =============== RIGHT PANEL ===============
        right = tk.Frame(main)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.columnconfigure(1, weight=1)
        right.rowconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        # Raw Nyquist
        raw_frame = tk.LabelFrame(right, text="Raw Nyquist")
        raw_frame.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        self.raw_fig = Figure(figsize=(5, 4), dpi=100)
        self.raw_ax = self.raw_fig.add_subplot(111)
        self.raw_ax.set_xlabel("Re(Z) [Ohm]")
        self.raw_ax.set_ylabel("-Im(Z) [Ohm]")
        self.raw_ax.grid(True)

        self.raw_canvas = FigureCanvasTkAgg(self.raw_fig, master=raw_frame)
        self.raw_canvas.get_tk_widget().pack(fill="both", expand=True)
        raw_toolbar = NavigationToolbar2Tk(self.raw_canvas, raw_frame)
        raw_toolbar.update()
        raw_toolbar.pack(fill="x")

        # Processed Nyquist
        proc_frame = tk.LabelFrame(right, text="Processed Nyquist")
        proc_frame.grid(row=0, column=1, sticky="nsew", padx=6, pady=6)

        self.proc_fig = Figure(figsize=(5, 4), dpi=100)
        self.proc_ax = self.proc_fig.add_subplot(111)
        self.proc_ax.set_xlabel("Re(Z) [Ohm]")
        self.proc_ax.set_ylabel("-Im(Z) [Ohm]")
        self.proc_ax.grid(True)

        self.proc_canvas = FigureCanvasTkAgg(self.proc_fig, master=proc_frame)
        self.proc_canvas.get_tk_widget().pack(fill="both", expand=True)
        proc_toolbar = NavigationToolbar2Tk(self.proc_canvas, proc_frame)
        proc_toolbar.update()
        proc_toolbar.pack(fill="x")

        # Frequency plot
        freq_frame = tk.LabelFrame(right, text="Frequency vs -Im")
        freq_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=6, pady=6)

        self.freq_fig = Figure(figsize=(10, 4), dpi=100)
        self.freq_ax = self.freq_fig.add_subplot(111)
        self.freq_ax.set_xscale("log")
        self.freq_ax.set_xlabel("Frequency [Hz]")
        self.freq_ax.set_ylabel("-Im(Z) [Ohm]")
        self.freq_ax.grid(True)

        self.freq_canvas = FigureCanvasTkAgg(self.freq_fig, master=freq_frame)
        self.freq_canvas.get_tk_widget().pack(fill="both", expand=True)
        freq_toolbar = NavigationToolbar2Tk(self.freq_canvas, freq_frame)
        freq_toolbar.update()
        freq_toolbar.pack(fill="x")

    def choose_input_file(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="Choose input file",
            filetypes=[
                ("Supported files", "*.txt *.dat *.csv *.tsv *.xlsx *.xlsm *.xltx *.xltm *.xls"),
                ("Text files", "*.txt *.dat *.csv *.tsv"),
                ("Excel files", "*.xlsx *.xlsm *.xltx *.xltm *.xls"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        self.input_path.set(path)
        self.raw_df = None
        self.proc_df = None

        self.sheet_cache = {}
        self.sheet_saved_params = {}
        self.cache_workbook_path = None

        try:
            if self._is_excel_input(path):
                self._load_excel_workbook(path)
                suggested = f"{Path(path).stem}_PROCESSED.xlsx"
                self.output_path.set(str(Path(path).with_name(suggested)))
                self.cache_workbook_path = str(Path(path).with_name(f"{Path(path).stem}__PREPROCESS_CACHE.xlsx"))
                self._on_sheet_changed()
                self._write_info(
                    f"Loaded Excel workbook:\n{path}\n\n"
                    f"Sheets detected: {len(self.sheet_names)}\n"
                    f"Current sheet: {self.current_sheet_name or '-'}\n"
                    f"Cache file: {self.cache_workbook_path}"
                )
            else:
                self._clear_excel_workbook_state()
                suggested = f"{Path(path).stem}_OUTPUT{Path(path).suffix}"
                self.output_path.set(str(Path(path).with_name(suggested)))
                self._write_info(f"Loaded input file:\n{path}")

            self._update_sheet_nav_state()

        except Exception as e:
            self._clear_excel_workbook_state()
            self.input_path.set("")
            messagebox.showerror("Load error", f"{e}", parent=self)

    def _is_excel_input(self, path=None):
        candidate = path if path is not None else self.input_path.get().strip()
        return bool(candidate) and is_excel_file(candidate)

    def _clear_excel_workbook_state(self):
        self.excel_sheets = {}
        self.sheet_names = []
        self.current_sheet_index = 0
        self.current_sheet_name = None
        self.sheet_nav_var.set("Text / single file")

    def _load_excel_workbook(self, path):
        self.excel_sheets = read_eis_excel_workbook(path)
        self.sheet_names = list(self.excel_sheets.keys())
        self.current_sheet_index = 0
        self.current_sheet_name = self.sheet_names[0] if self.sheet_names else None
        self._update_sheet_nav_state()

    def _update_sheet_nav_state(self):
        if self.sheet_names:
            total = len(self.sheet_names)
            index = self.current_sheet_index + 1
            current = self.sheet_names[self.current_sheet_index]
            self.current_sheet_name = current
            self.sheet_nav_var.set(f"Sheet {index}/{total}\n{current}")
            prev_state = tk.NORMAL if total > 1 else tk.DISABLED
            next_state = tk.NORMAL if total > 1 else tk.DISABLED
        else:
            self.current_sheet_name = None
            self.sheet_nav_var.set("Text / single file")
            prev_state = tk.DISABLED
            next_state = tk.DISABLED

        self.prev_sheet_btn.config(state=prev_state)
        self.next_sheet_btn.config(state=next_state)

    def go_to_previous_sheet(self):
        if not self.sheet_names:
            return
        self.current_sheet_index = (self.current_sheet_index - 1) % len(self.sheet_names)
        self._on_sheet_changed()
        self.preview_processing()

    def go_to_next_sheet(self):
        if not self.sheet_names:
            return
        self.current_sheet_index = (self.current_sheet_index + 1) % len(self.sheet_names)
        self._on_sheet_changed()
        self.preview_processing()

    def _get_stop_first_col_value(self):
        if not self.use_stop_first_col.get():
            return None
        txt = self.stop_first_col.get().strip()
        if not txt:
            raise ValueError("Stop value is enabled but empty.")
        return float(txt.replace(",", "."))

    def _get_wo_params(self):
        if not self.apply_wo.get():
            return None

        wor = float(self.wor.get().strip().replace(",", "."))
        wot = float(self.wot.get().strip().replace(",", "."))
        wop = float(self.wop.get().strip().replace(",", "."))

        fmin = None
        if self.use_fmin.get():
            txt = self.fmin.get().strip()
            if not txt:
                raise ValueError("Minimum frequency is enabled but empty.")
            fmin = float(txt.replace(",", "."))

        return {
            "wor": wor,
            "wot": wot,
            "wop": wop,
            "fmin": fmin,
            "trim_pos_im": self.trim_pos_im.get(),
        }

    def _apply_processing_pipeline(self, raw_df):
        if raw_df.empty:
            raise ValueError("No valid points after preprocessing.")

        proc_df = raw_df.copy()
        wo_params = self._get_wo_params()
        if wo_params is not None:
            proc_df = remove_warburg_wo(
                proc_df,
                wor=wo_params["wor"],
                wot=wo_params["wot"],
                wop=wo_params["wop"],
                fmin=wo_params["fmin"],
                trim_pos_im=wo_params["trim_pos_im"],
            )[["freq", "R", "Im"]].copy()

        return raw_df, proc_df

    def _preview_text_input(self, path, stop_first_col):
        raw_df = preprocess_eclab_text_to_dataframe(
            input_path=path,
            skip_first_line=self.skip_first_line.get(),
            replace_comma=self.replace_comma.get(),
            keep_n_cols=self.keep_n_cols.get(),
            negate_col3=self.negate_col3.get(),
            stop_first_col=stop_first_col,
        )
        return self._apply_processing_pipeline(raw_df)

    def _preview_excel_input(self, stop_first_col):
        if not self.sheet_names:
            raise ValueError("Aucune feuille Excel chargée.")

        sheet_name = self.sheet_names[self.current_sheet_index]
        sheet_df = self.excel_sheets[sheet_name]
        raw_df = preprocess_eis_dataframe(
            sheet_df,
            keep_n_cols=self.keep_n_cols.get(),
            negate_col3=self.negate_col3.get(),
            stop_first_col=stop_first_col,
        )
        return sheet_name, *self._apply_processing_pipeline(raw_df)

    def preview_processing(self):
        path = self.input_path.get().strip()
        if not path:
            messagebox.showwarning("No file", "Choose an input file first.", parent=self)
            return

        try:
            stop_first_col = self._get_stop_first_col_value()

            if self._is_excel_input(path):
                sheet_name, raw_df, proc_df = self._preview_excel_input(stop_first_col)
                self.current_sheet_name = sheet_name
            else:
                raw_df, proc_df = self._preview_text_input(path, stop_first_col)
                self.current_sheet_name = None

            self.raw_df = raw_df
            self.proc_df = proc_df

            self._draw_plots()
            self._update_sheet_nav_state()
            self._update_info_panel()

        except Exception as e:
            messagebox.showerror("Preview error", f"{e}", parent=self)

    def _export_text_file(self, path, stop_first_col):
        default_name = f"{Path(path).stem}_OUTPUT{Path(path).suffix}"
        save_path = filedialog.asksaveasfilename(
            parent=self,
            title="Export processed text file",
            defaultextension=Path(path).suffix or ".txt",
            initialfile=default_name,
            filetypes=[
                ("Text files", "*.txt *.dat *.csv *.tsv"),
                ("All files", "*.*"),
            ],
        )
        if not save_path:
            return None

        out_path = preprocess_eclab_text_file(
            input_path=path,
            output_path=save_path,
            skip_first_line=self.skip_first_line.get(),
            replace_comma=self.replace_comma.get(),
            keep_n_cols=self.keep_n_cols.get(),
            negate_col3=self.negate_col3.get(),
            stop_first_col=stop_first_col,
        )
        return out_path

    def _export_excel_workbook(self, path, stop_first_col):
        if not self.sheet_names:
            raise ValueError("Aucune feuille Excel chargée.")

        default_name = f"{Path(path).stem}_PROCESSED.xlsx"
        save_path = filedialog.asksaveasfilename(
            parent=self,
            title="Export processed Excel workbook",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[
                ("Excel workbook", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if not save_path:
            return None

        processed_sheets = {}
        for sheet_name in self.sheet_names:
            raw_df = preprocess_eis_dataframe(
                self.excel_sheets[sheet_name],
                keep_n_cols=self.keep_n_cols.get(),
                negate_col3=self.negate_col3.get(),
                stop_first_col=stop_first_col,
            )
            _, proc_df = self._apply_processing_pipeline(raw_df)
            processed_sheets[sheet_name] = proc_df

        return export_processed_eis_workbook(processed_sheets, save_path)

    def export_processed_file(self):
        path = self.input_path.get().strip()
        if not path:
            messagebox.showwarning("No file", "Choose an input file first.", parent=self)
            return

        try:
            # mode Excel = export à partir du cache uniquement
            if self._is_excel_input(path):
                if not self.sheet_cache:
                    messagebox.showwarning(
                        "Nothing to export",
                        "No processed sheet has been cached yet.\nUse 'Update current sheet' first.",
                        parent=self,
                    )
                    return

                default_name = f"{Path(path).stem}_PROCESSED.xlsx"
                save_path = filedialog.asksaveasfilename(
                    parent=self,
                    title="Export processed Excel workbook",
                    defaultextension=".xlsx",
                    initialfile=default_name,
                    filetypes=[
                        ("Excel workbook", "*.xlsx"),
                        ("All files", "*.*"),
                    ],
                )
                if not save_path:
                    return

                processed_sheets = {
                    sheet_name: entry["proc_df"]
                    for sheet_name, entry in self.sheet_cache.items()
                    if entry.get("proc_df") is not None
                }

                out_path = export_processed_sheets_over_original_workbook(
                    source_workbook_path=path,
                    processed_sheets=processed_sheets,
                    out_path=save_path,
                    metadata_rows=self._build_cache_metadata_rows(),
                )

                self._write_info(
                    f"Processed workbook exported:\n{out_path}\n\n"
                    f"Updated sheets exported: {len(processed_sheets)}\n"
                    f"Unprocessed sheets kept intact from original workbook."
                )
                messagebox.showinfo("Export done", f"File created:\n{out_path}", parent=self)
                return

            # mode texte = comportement normal
            stop_first_col = self._get_stop_first_col_value()

            default_name = f"{Path(path).stem}_OUTPUT{Path(path).suffix}"
            save_path = filedialog.asksaveasfilename(
                parent=self,
                title="Export processed text file",
                defaultextension=Path(path).suffix or ".txt",
                initialfile=default_name,
                filetypes=[
                    ("Text files", "*.txt *.dat *.csv *.tsv"),
                    ("All files", "*.*"),
                ],
            )
            if not save_path:
                return

            out_path = preprocess_eclab_text_file(
                input_path=path,
                output_path=save_path,
                skip_first_line=self.skip_first_line.get(),
                replace_comma=self.replace_comma.get(),
                keep_n_cols=self.keep_n_cols.get(),
                negate_col3=self.negate_col3.get(),
                stop_first_col=stop_first_col,
            )

            self._write_info(f"Processed file exported:\n{out_path}")
            messagebox.showinfo("Export done", f"File created:\n{out_path}", parent=self)

        except Exception as e:
            messagebox.showerror("Export error", f"{e}", parent=self)

    def _draw_plots(self):
        # raw
        self.raw_ax.clear()
        self.raw_ax.set_xlabel("Re(Z) [Ohm]")
        self.raw_ax.set_ylabel("-Im(Z) [Ohm]")
        self.raw_ax.grid(True)

        if self.raw_df is not None and not self.raw_df.empty:
            self.raw_ax.plot(self.raw_df["R"], -self.raw_df["Im"], linewidth=1.0)
            self.raw_ax.scatter(self.raw_df["R"], -self.raw_df["Im"], s=10)
            self.raw_ax.set_aspect("equal", adjustable="datalim")

        self.raw_fig.tight_layout()
        self.raw_canvas.draw()

        # processed
        self.proc_ax.clear()
        self.proc_ax.set_xlabel("Re(Z) [Ohm]")
        self.proc_ax.set_ylabel("-Im(Z) [Ohm]")
        self.proc_ax.grid(True)

        if self.proc_df is not None and not self.proc_df.empty:
            self.proc_ax.plot(self.proc_df["R"], -self.proc_df["Im"], linewidth=1.0)
            self.proc_ax.scatter(self.proc_df["R"], -self.proc_df["Im"], s=10)
            self.proc_ax.set_aspect("equal", adjustable="datalim")

        self.proc_fig.tight_layout()
        self.proc_canvas.draw()

        # frequency plot
        self.freq_ax.clear()
        self.freq_ax.set_xscale("log")
        self.freq_ax.set_xlabel("Frequency [Hz]")
        self.freq_ax.set_ylabel("-Im(Z) [Ohm]")
        self.freq_ax.grid(True)

        if self.raw_df is not None and not self.raw_df.empty:
            self.freq_ax.plot(self.raw_df["freq"], -self.raw_df["Im"], linewidth=1.0, label="Raw")
            self.freq_ax.scatter(self.raw_df["freq"], -self.raw_df["Im"], s=10, label="Raw")
        if self.proc_df is not None and not self.proc_df.empty:
            self.freq_ax.plot(self.proc_df["freq"], -self.proc_df["Im"], linewidth=1.0, label="Processed")
            self.freq_ax.scatter(self.proc_df["freq"], -self.proc_df["Im"], s=10, label="Processed")

        self.freq_ax.invert_xaxis()

        if self.raw_df is not None or self.proc_df is not None:
            self.freq_ax.legend()

        self.freq_fig.tight_layout()
        self.freq_canvas.draw()

    def _update_info_panel(self):
        self.info_text.delete("1.0", tk.END)

        self.info_text.insert(tk.END, "Input file\n")
        self.info_text.insert(tk.END, f"{self.input_path.get()}\n\n")

        if self.sheet_names:
            self.info_text.insert(tk.END, "Workbook mode\n")
            self.info_text.insert(tk.END, f"Sheets in workbook: {len(self.sheet_names)}\n")
            self.info_text.insert(tk.END, f"Current sheet: {self.current_sheet_name}\n")
            self.info_text.insert(tk.END, f"Cached processed sheets: {len(self.sheet_cache)}\n")
            self.info_text.insert(tk.END, f"Current sheet cached: {self.current_sheet_name in self.sheet_cache}\n")
            if self.cache_workbook_path:
                self.info_text.insert(tk.END, f"Cache file: {self.cache_workbook_path}\n")
            self.info_text.insert(tk.END, "\n\n")

        if self.raw_df is not None and not self.raw_df.empty:
            self.info_text.insert(tk.END, "Raw data\n")
            self.info_text.insert(tk.END, f"Points: {len(self.raw_df)}\n")
            self.info_text.insert(tk.END, f"f_start: {self.raw_df['freq'].iloc[0]:.6g} Hz\n")
            self.info_text.insert(tk.END, f"f_end:   {self.raw_df['freq'].iloc[-1]:.6g} Hz\n")
            self.info_text.insert(tk.END, f"R range: [{self.raw_df['R'].min():.6g}, {self.raw_df['R'].max():.6g}]\n")
            self.info_text.insert(tk.END, f"Im range: [{self.raw_df['Im'].min():.6g}, {self.raw_df['Im'].max():.6g}]\n\n")

        if self.proc_df is not None and not self.proc_df.empty:
            self.info_text.insert(tk.END, "Processed data\n")
            self.info_text.insert(tk.END, f"Points: {len(self.proc_df)}\n")
            self.info_text.insert(tk.END, f"f_start: {self.proc_df['freq'].iloc[0]:.6g} Hz\n")
            self.info_text.insert(tk.END, f"f_end:   {self.proc_df['freq'].iloc[-1]:.6g} Hz\n")
            self.info_text.insert(tk.END, f"R range: [{self.proc_df['R'].min():.6g}, {self.proc_df['R'].max():.6g}]\n")
            self.info_text.insert(tk.END, f"Im range: [{self.proc_df['Im'].min():.6g}, {self.proc_df['Im'].max():.6g}]\n\n")

        self.info_text.insert(tk.END, "Applied options\n")
        self.info_text.insert(tk.END, f"- Skip first line: {self.skip_first_line.get()}\n")
        self.info_text.insert(tk.END, f"- Replace comma: {self.replace_comma.get()}\n")
        self.info_text.insert(tk.END, f"- Keep first N columns: {self.keep_n_cols.get()}\n")
        self.info_text.insert(tk.END, f"- Negate 3rd column: {self.negate_col3.get()}\n")
        self.info_text.insert(tk.END, f"- Stop at first-column value: {self.use_stop_first_col.get()}\n")

        if self.use_stop_first_col.get():
            self.info_text.insert(tk.END, f"  value = {self.stop_first_col.get()}\n")

        self.info_text.insert(tk.END, f"- Apply Wo removal: {self.apply_wo.get()}\n")
        if self.apply_wo.get():
            self.info_text.insert(tk.END, f"  Wo-R = {self.wor.get()}\n")
            self.info_text.insert(tk.END, f"  Wo-T = {self.wot.get()}\n")
            self.info_text.insert(tk.END, f"  Wo-P = {self.wop.get()}\n")
            if self.use_fmin.get():
                self.info_text.insert(tk.END, f"  fmin = {self.fmin.get()}\n")
            self.info_text.insert(tk.END, f"  trim Im > 0 = {self.trim_pos_im.get()}\n")

    def _write_info(self, text):
        self.info_text.delete("1.0", tk.END)
        self.info_text.insert(tk.END, text)

    def _collect_current_params(self):
        return {
            "skip_first_line": self.skip_first_line.get(),
            "replace_comma": self.replace_comma.get(),
            "keep_n_cols": self.keep_n_cols.get(),
            "negate_col3": self.negate_col3.get(),
            "use_stop_first_col": self.use_stop_first_col.get(),
            "stop_first_col": self.stop_first_col.get(),
            "apply_wo": self.apply_wo.get(),
            "wor": self.wor.get(),
            "wot": self.wot.get(),
            "wop": self.wop.get(),
            "use_fmin": self.use_fmin.get(),
            "fmin": self.fmin.get(),
            "trim_pos_im": self.trim_pos_im.get(),
        }

    def _apply_params_to_ui(self, params: dict):
        self.skip_first_line.set(params.get("skip_first_line", True))
        self.replace_comma.set(params.get("replace_comma", True))
        self.keep_n_cols.set(params.get("keep_n_cols", 3))
        self.negate_col3.set(params.get("negate_col3", False))
        self.use_stop_first_col.set(params.get("use_stop_first_col", False))
        self.stop_first_col.set(params.get("stop_first_col", ""))
        self.apply_wo.set(params.get("apply_wo", False))
        self.wor.set(params.get("wor", ""))
        self.wot.set(params.get("wot", ""))
        self.wop.set(params.get("wop", ""))
        self.use_fmin.set(params.get("use_fmin", False))
        self.fmin.set(params.get("fmin", ""))
        self.trim_pos_im.set(params.get("trim_pos_im", False))

    def _restore_params_for_current_sheet(self):
        if not self.current_sheet_name:
            return
        params = self.sheet_saved_params.get(self.current_sheet_name)
        if params:
            self._apply_params_to_ui(params)

    def _load_cached_view_for_current_sheet(self):
        if not self.current_sheet_name:
            self.raw_df = None
            self.proc_df = None
            self._draw_plots()
            self._update_info_panel()
            return

        entry = self.sheet_cache.get(self.current_sheet_name)
        if entry is None:
            self.raw_df = None
            self.proc_df = None
            self._draw_plots()
            self._update_info_panel()
            return

        self.raw_df = entry["raw_df"].copy()
        self.proc_df = entry["proc_df"].copy()
        self._draw_plots()
        self._update_info_panel()

    def _on_sheet_changed(self):
        self._update_sheet_nav_state()
        self._restore_params_for_current_sheet()
        self._load_cached_view_for_current_sheet()

    def _build_cache_metadata_rows(self):
        rows = []
        for sheet_name, entry in self.sheet_cache.items():
            params = entry.get("params", {})
            proc_df = entry.get("proc_df")
            rows.append({
                "source_sheet": sheet_name,
                "updated_at": entry.get("updated_at", ""),
                "n_points": int(len(proc_df)) if proc_df is not None else 0,
                "keep_n_cols": params.get("keep_n_cols"),
                "negate_col3": params.get("negate_col3"),
                "use_stop_first_col": params.get("use_stop_first_col"),
                "stop_first_col": params.get("stop_first_col"),
                "apply_wo": params.get("apply_wo"),
                "wor": params.get("wor"),
                "wot": params.get("wot"),
                "wop": params.get("wop"),
                "use_fmin": params.get("use_fmin"),
                "fmin": params.get("fmin"),
                "trim_pos_im": params.get("trim_pos_im"),
            })
        return rows

    def _write_cache_workbook(self):
        if not self._is_excel_input():
            return
        if not self.cache_workbook_path:
            return

        processed_sheets = {
            sheet_name: entry["proc_df"]
            for sheet_name, entry in self.sheet_cache.items()
            if entry.get("proc_df") is not None
        }

        export_processed_sheets_over_original_workbook(
            source_workbook_path=self.input_path.get().strip(),
            processed_sheets=processed_sheets,
            out_path=self.cache_workbook_path,
            metadata_rows=self._build_cache_metadata_rows(),
        )

    def update_current_sheet_cache(self):
        path = self.input_path.get().strip()
        if not path:
            messagebox.showwarning("No file", "Choose an input file first.", parent=self)
            return

        if not self._is_excel_input(path):
            messagebox.showinfo(
                "Text mode",
                "The cache/update workflow is only used for Excel workbooks.\nUse Preview or Export directly for text files.",
                parent=self,
            )
            return

        try:
            stop_first_col = self._get_stop_first_col_value()
            sheet_name, raw_df, proc_df = self._preview_excel_input(stop_first_col)

            params = self._collect_current_params()

            self.sheet_saved_params[sheet_name] = params.copy()
            self.sheet_cache[sheet_name] = {
                "raw_df": raw_df.copy(),
                "proc_df": proc_df.copy(),
                "params": params.copy(),
                "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }

            self.raw_df = raw_df
            self.proc_df = proc_df

            self._write_cache_workbook()
            self._draw_plots()
            self._update_info_panel()

            messagebox.showinfo(
                "Sheet updated",
                f"Current sheet cached successfully:\n{sheet_name}\n\nCache file:\n{self.cache_workbook_path}",
                parent=self,
            )

        except Exception as e:
            messagebox.showerror("Update error", f"{e}", parent=self)
