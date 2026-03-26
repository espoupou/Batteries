
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from dialogs import DRTDialog
from drt_core import load_drt_measures, run_drt_workbook


class DRTPlotterWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent_app = parent

        self.title("DRT Plotter")
        self.geometry("1400x900")
        self.minsize(1200, 760)

        self.current_path = ""
        self.current_measures = []
        self.current_summary = None

        self.build_ui()
        self.lift()
        self.focus_force()

    def build_ui(self):
        main = tk.Frame(self)
        main.pack(fill="both", expand=True, padx=10, pady=10)

        # grille 2 x 3
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)
        main.columnconfigure(2, weight=1)
        main.rowconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        # panneau de contrôle
        ctrl = tk.LabelFrame(main, text="Contrôles / infos")
        ctrl.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        self.info_text = tk.Text(ctrl, height=16, wrap="word")
        self.info_text.pack(fill="both", expand=True, padx=8, pady=8)

        btns = tk.Frame(ctrl)
        btns.pack(fill="x", padx=8, pady=(0, 8))

        tk.Button(btns, text="Ouvrir DRT Excel...", command=self.load_drt_workbook).pack(fill="x", pady=3)
        tk.Button(btns, text="Recalculer DRT...", command=self.recalculate_drt).pack(fill="x", pady=3)
        tk.Button(btns, text="Actualiser affichage", command=self.refresh_plots).pack(fill="x", pady=3)

        self.figures = []
        self.axes = []
        self.canvases = []

        positions = [(0, 1), (0, 2), (1, 0), (1, 1), (1, 2)]

        for row, col in positions:
            frame = tk.LabelFrame(main, text="Mesure")
            frame.grid(row=row, column=col, sticky="nsew", padx=6, pady=6)

            fig = Figure(figsize=(4, 3), dpi=100)
            ax = fig.add_subplot(111)
            ax.set_xscale("log")
            ax.set_xlabel("Tau (s)")
            ax.set_ylabel("Gamma")
            ax.grid(True)

            canvas = FigureCanvasTkAgg(fig, master=frame)
            canvas.get_tk_widget().pack(fill="both", expand=True)

            toolbar = NavigationToolbar2Tk(canvas, frame)
            toolbar.update()
            toolbar.pack(fill="x")

            self.figures.append(fig)
            self.axes.append(ax)
            self.canvases.append(canvas)

    def load_drt_workbook(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="Choisir un classeur DRT",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")],
        )
        if not path:
            return

        try:
            measures, summary = load_drt_measures(path)
            if not measures:
                messagebox.showwarning("Aucune DRT", "Aucune feuille DRT exploitable trouvée.", parent=self)
                return

            self.current_path = path
            self.current_measures = measures[:5]
            self.current_summary = summary
            self.refresh_plots()

        except Exception as e:
            messagebox.showerror("Erreur DRT Plotter", f"{e}", parent=self)

    def refresh_plots(self):
        for ax, canvas in zip(self.axes, self.canvases):
            ax.clear()
            ax.set_xscale("log")
            ax.set_xlabel("Tau (s)")
            ax.set_ylabel("Gamma")
            ax.grid(True)

        for i, m in enumerate(self.current_measures[:5]):
            ax = self.axes[i]
            ax.plot(m["tau"], m["gamma"], linewidth=1.2)
            ax.set_title(m["sheet_name"])

            meta = m.get("meta", {})
            subtitle_parts = []
            if "lambda_optimal" in meta:
                subtitle_parts.append(f"λ={meta['lambda_optimal']:.3e}")
            if "R_ohm" in meta:
                subtitle_parts.append(f"R={meta['R_ohm']:.3g}")
            if "L_H" in meta:
                subtitle_parts.append(f"L={meta['L_H']:.3g}")
            if subtitle_parts:
                ax.text(
                    0.02, 0.98,
                    "\n".join(subtitle_parts),
                    transform=ax.transAxes,
                    va="top", ha="left",
                    fontsize=8,
                    bbox=dict(boxstyle="round", alpha=0.15)
                )

        for fig, canvas in zip(self.figures, self.canvases):
            fig.tight_layout()
            canvas.draw()

        self.update_info_panel()

    def update_info_panel(self):
        self.info_text.delete("1.0", tk.END)

        if not self.current_path:
            self.info_text.insert(tk.END, "Aucun classeur DRT chargé.")
            return

        self.info_text.insert(tk.END, f"Fichier : {self.current_path}\n\n")
        self.info_text.insert(tk.END, f"Nombre de mesures détectées : {len(self.current_measures)}\n\n")

        for i, m in enumerate(self.current_measures, start=1):
            self.info_text.insert(tk.END, f"{i}. {m['sheet_name']}\n")
            meta = m.get("meta", {})
            for k in ["Potential_V", "R_ohm", "L_H", "lambda_optimal"]:
                if k in meta:
                    self.info_text.insert(tk.END, f"   - {k}: {meta[k]}\n")
            self.info_text.insert(tk.END, "\n")

    def recalculate_drt(self):
        source_path = filedialog.askopenfilename(
            parent=self,
            title="Choisir le classeur Excel source pour recalcul DRT",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")],
        )
        if not source_path:
            return

        dlg = DRTDialog(self)
        params = dlg.result
        if not params:
            return

        self.parent_app.status.set("DRT en cours...")

        def job():
            return run_drt_workbook(
                excel_file_path=source_path,
                rbf_type=params["rbf_type"],
                data_used=params["data_used"],
                induct_used=params["induct_used"],
                der_used=params["der_used"],
                cv_type=params["cv_type"],
                reg_param=params["reg_param"],
                shape_control=params["shape_control"],
                coeff=params["coeff"],
                run_twice=params["run_twice"],
            )

        def on_done(result_path):
            self.parent_app.status.set(f"DRT recalculée: {result_path}")
            self.current_path = result_path
            self.current_measures, self.current_summary = load_drt_measures(result_path)
            self.current_measures = self.current_measures[:5]
            self.refresh_plots()
            messagebox.showinfo("DRT terminée", f"Fichier créé :\n{result_path}", parent=self)

        def worker():
            try:
                result_path = job()
                self.after(0, lambda: on_done(result_path))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Erreur DRT", str(e), parent=self))

        threading.Thread(target=worker, daemon=True).start()
