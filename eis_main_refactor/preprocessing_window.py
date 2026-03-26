
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from eis_core import (
    preprocess_eclab_text_file,
    preprocess_eclab_text_to_dataframe,
    remove_warburg_wo,
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

        self._build_ui()
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
            row=1, column=0, columnspan=2, sticky="we", padx=8
        )
        tk.Button(left, text="Choose file...", command=self.choose_input_file).grid(
            row=2, column=0, columnspan=2, sticky="we", padx=8, pady=(4, 8)
        )

        # conversion simple
        block1 = tk.LabelFrame(left, text="Text preprocessing")
        block1.grid(row=3, column=0, columnspan=2, sticky="we", padx=8, pady=6)

        tk.Checkbutton(block1, text="Skip first line", variable=self.skip_first_line).grid(
            row=0, column=0, sticky="w", padx=8, pady=4
        )
        tk.Checkbutton(block1, text="Replace comma by dot", variable=self.replace_comma).grid(
            row=1, column=0, sticky="w", padx=8, pady=4
        )
        tk.Checkbutton(block1, text="Negate 3rd column", variable=self.negate_col3).grid(
            row=2, column=0, sticky="w", padx=8, pady=4
        )

        tk.Label(block1, text="Keep first N columns").grid(row=3, column=0, sticky="w", padx=8, pady=(6, 2))
        tk.Entry(block1, textvariable=self.keep_n_cols, width=10).grid(row=3, column=1, sticky="w", padx=8, pady=(6, 2))

        tk.Checkbutton(
            block1,
            text="Stop when first column reaches",
            variable=self.use_stop_first_col
        ).grid(row=4, column=0, sticky="w", padx=8, pady=(8, 4))
        tk.Entry(block1, textvariable=self.stop_first_col, width=12).grid(
            row=4, column=1, sticky="w", padx=8, pady=(8, 4)
        )

        # Wo remove
        block2 = tk.LabelFrame(left, text="Wo diffusion removal")
        block2.grid(row=4, column=0, columnspan=2, sticky="we", padx=8, pady=6)

        tk.Checkbutton(block2, text="Apply Wo removal", variable=self.apply_wo).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=8, pady=4
        )

        tk.Label(block2, text="Wo-R").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        tk.Entry(block2, textvariable=self.wor, width=12).grid(row=1, column=1, sticky="w", padx=8, pady=4)

        tk.Label(block2, text="Wo-T").grid(row=2, column=0, sticky="w", padx=8, pady=4)
        tk.Entry(block2, textvariable=self.wot, width=12).grid(row=2, column=1, sticky="w", padx=8, pady=4)

        tk.Label(block2, text="Wo-P").grid(row=3, column=0, sticky="w", padx=8, pady=4)
        tk.Entry(block2, textvariable=self.wop, width=12).grid(row=3, column=1, sticky="w", padx=8, pady=4)

        tk.Checkbutton(block2, text="Use minimum frequency", variable=self.use_fmin).grid(
            row=4, column=0, sticky="w", padx=8, pady=(8, 4)
        )
        tk.Entry(block2, textvariable=self.fmin, width=12).grid(
            row=4, column=1, sticky="w", padx=8, pady=(8, 4)
        )

        tk.Checkbutton(block2, text="Drop corrected points with Im > 0", variable=self.trim_pos_im).grid(
            row=5, column=0, columnspan=2, sticky="w", padx=8, pady=4
        )

        # actions
        block3 = tk.LabelFrame(left, text="Actions")
        block3.grid(row=5, column=0, columnspan=2, sticky="we", padx=8, pady=6)

        tk.Button(block3, text="Preview", command=self.preview_processing).pack(fill="x", padx=8, pady=(8, 4))
        tk.Button(block3, text="Export processed file...", command=self.export_processed_file).pack(
            fill="x", padx=8, pady=4
        )
        tk.Button(block3, text="Close", command=self.destroy).pack(fill="x", padx=8, pady=(4, 8))

        # infos
        block4 = tk.LabelFrame(left, text="Info")
        block4.grid(row=6, column=0, columnspan=2, sticky="nsew", padx=8, pady=6)

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
            title="Choose input text file",
            filetypes=[
                ("Text files", "*.txt *.dat *.csv *.tsv"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        self.input_path.set(path)
        suggested = f"{Path(path).stem}_OUTPUT{Path(path).suffix}"
        self.output_path.set(str(Path(path).with_name(suggested)))
        self._write_info(f"Loaded input file:\n{path}")

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

    def preview_processing(self):
        path = self.input_path.get().strip()
        if not path:
            messagebox.showwarning("No file", "Choose an input file first.", parent=self)
            return

        try:
            stop_first_col = self._get_stop_first_col_value()

            raw_df = preprocess_eclab_text_to_dataframe(
                input_path=path,
                skip_first_line=self.skip_first_line.get(),
                replace_comma=self.replace_comma.get(),
                keep_n_cols=self.keep_n_cols.get(),
                negate_col3=self.negate_col3.get(),
                stop_first_col=stop_first_col,
            )

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

            self.raw_df = raw_df
            self.proc_df = proc_df

            self._draw_plots()
            self._update_info_panel()

        except Exception as e:
            messagebox.showerror("Preview error", f"{e}", parent=self)

    def export_processed_file(self):
        path = self.input_path.get().strip()
        if not path:
            messagebox.showwarning("No file", "Choose an input file first.", parent=self)
            return

        try:
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

    def _write_info(self, text: str):
        self.info_text.delete("1.0", tk.END)
        self.info_text.insert(tk.END, text)
