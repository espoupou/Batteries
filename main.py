import os
import re
import threading
import pandas as pd
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from eis_core import (
    read_freq_R_Im,
    remove_warburg_wo,
    export_freq_R_Im,
    read_eis_raw_table,
    split_raw_eis_by_start_freq,
    export_raw_eis_batches_to_excel,
    collect_excel_workbooks,
    preprocess_eclab_text_file,
    preprocess_eclab_text_to_dataframe,
)

from drt_core import (
    savgol_workbook,
    linkk_workbook,
    run_drt_workbook,
    extract_eis_from_sheet,
    load_excel_measures,
    gradient_colors, load_drt_measures
)

class WarburgDialog(simpledialog.Dialog):
    def __init__(self, parent, available_slots):
        self.available_slots = available_slots
        self.result = None
        super().__init__(parent, title="Suppression diffusion Warburg (Wo)")

    def body(self, master):
        tk.Label(master, text="Slot cible :").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.slot_var = tk.StringVar(value=str(self.available_slots[0]))
        tk.OptionMenu(master, self.slot_var, *[str(i) for i in self.available_slots]).grid(
            row=0, column=1, sticky="we", padx=6, pady=6
        )

        tk.Label(master, text="Wo-R :").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.wor_entry = tk.Entry(master)
        self.wor_entry.grid(row=1, column=1, sticky="we", padx=6, pady=6)

        tk.Label(master, text="Wo-T :").grid(row=2, column=0, sticky="w", padx=6, pady=6)
        self.wot_entry = tk.Entry(master)
        self.wot_entry.grid(row=2, column=1, sticky="we", padx=6, pady=6)

        tk.Label(master, text="Wo-P :").grid(row=3, column=0, sticky="w", padx=6, pady=6)
        self.wop_entry = tk.Entry(master)
        self.wop_entry.grid(row=3, column=1, sticky="we", padx=6, pady=6)

        tk.Label(master, text="fmin (optionnel) :").grid(row=4, column=0, sticky="w", padx=6, pady=6)
        self.fmin_entry = tk.Entry(master)
        self.fmin_entry.grid(row=4, column=1, sticky="we", padx=6, pady=6)

        self.trim_pos_im = tk.BooleanVar(value=False)
        tk.Checkbutton(
            master,
            text="Supprimer les points corrigés avec Im > 0",
            variable=self.trim_pos_im
        ).grid(row=5, column=0, columnspan=2, sticky="w", padx=6, pady=6)

        master.columnconfigure(1, weight=1)
        return self.wor_entry

    def validate(self):
        try:
            wor = float(self.wor_entry.get().strip().replace(",", "."))
            wot = float(self.wot_entry.get().strip().replace(",", "."))
            wop = float(self.wop_entry.get().strip().replace(",", "."))

            fmin_txt = self.fmin_entry.get().strip()
            fmin = None if not fmin_txt else float(fmin_txt.replace(",", "."))

            self.result = {
                "slot_idx": int(self.slot_var.get()) - 1,
                "wor": wor,
                "wot": wot,
                "wop": wop,
                "fmin": fmin,
                "trim_pos_im": self.trim_pos_im.get(),
            }
            return True
        except Exception as e:
            messagebox.showerror("Paramètres invalides", f"Valeurs incorrectes.\n\nDétail :\n{e}")
            return False

class BatchExportDialog(simpledialog.Dialog):
    def __init__(self, parent, available_slots):
        self.available_slots = available_slots
        self.result = None
        super().__init__(parent, title="Batch loops -> Export Excel")

    def body(self, master):
        self.after(50, self._raise_front)

        tk.Label(master, text="Slot cible :").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.slot_var = tk.StringVar(value=str(self.available_slots[0]))
        tk.OptionMenu(master, self.slot_var, *[str(i) for i in self.available_slots]).grid(
            row=0, column=1, sticky="we", padx=6, pady=6
        )

        tk.Label(master, text="Fréquence de départ (Hz) :").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.start_freq_entry = tk.Entry(master)
        self.start_freq_entry.insert(0, "1000000")
        self.start_freq_entry.grid(row=1, column=1, sticky="we", padx=6, pady=6)

        tk.Label(master, text="Tolérance relative :").grid(row=2, column=0, sticky="w", padx=6, pady=6)
        self.rel_tol_entry = tk.Entry(master)
        self.rel_tol_entry.insert(0, "0.15")
        self.rel_tol_entry.grid(row=2, column=1, sticky="we", padx=6, pady=6)

        tk.Label(master, text="Jump ratio :").grid(row=3, column=0, sticky="w", padx=6, pady=6)
        self.jump_ratio_entry = tk.Entry(master)
        self.jump_ratio_entry.insert(0, "5.0")
        self.jump_ratio_entry.grid(row=3, column=1, sticky="we", padx=6, pady=6)

        tk.Label(master, text="Min points / mesure :").grid(row=4, column=0, sticky="w", padx=6, pady=6)
        self.min_points_entry = tk.Entry(master)
        self.min_points_entry.insert(0, "10")
        self.min_points_entry.grid(row=4, column=1, sticky="we", padx=6, pady=6)

        master.columnconfigure(1, weight=1)
        return self.start_freq_entry

    def _raise_front(self):
        try:
            self.lift()
            self.focus_force()
            self.attributes("-topmost", True)
            self.after(150, lambda: self.attributes("-topmost", False))
        except Exception:
            pass

    def validate(self):
        try:
            self.result = {
                "slot_idx": int(self.slot_var.get()) - 1,
                "start_freq_hz": float(self.start_freq_entry.get().strip().replace(",", ".")),
                "rel_tol": float(self.rel_tol_entry.get().strip().replace(",", ".")),
                "jump_ratio": float(self.jump_ratio_entry.get().strip().replace(",", ".")),
                "min_points": int(self.min_points_entry.get().strip()),
            }
            return True
        except Exception as e:
            messagebox.showerror(
                "Paramètres invalides",
                f"Valeurs incorrectes.\n\nDétail :\n{e}",
                parent=self
            )
            return False

class MergeSheetsDialog(simpledialog.Dialog):
    def __init__(self, parent, items):
        self.items = items
        self.result = None
        super().__init__(parent, title="Fusionner des sheets Excel")

    def body(self, master):
        self.after(50, self._raise_front)

        tk.Label(master, text="Choisis les sheets à fusionner :").pack(anchor="w", padx=8, pady=(8, 4))

        self.listbox = tk.Listbox(master, selectmode=tk.EXTENDED, width=90, height=18)
        self.listbox.pack(fill="both", expand=True, padx=8, pady=8)

        for item in self.items:
            label = f"{item['file_name']}  |  {item['sheet_name']}"
            self.listbox.insert(tk.END, label)

        opts = tk.Frame(master)
        opts.pack(fill="x", padx=8, pady=(0, 8))

        tk.Label(opts, text="Nom du grand sheet :").pack(side="left")
        self.name_entry = tk.Entry(opts, width=28)
        self.name_entry.insert(0, "Merged_Measures")
        self.name_entry.pack(side="left", padx=6)

        self.add_source_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            opts,
            text="Ajouter colonnes source_file / source_sheet",
            variable=self.add_source_var
        ).pack(side="left", padx=12)

        return self.listbox

    def _raise_front(self):
        try:
            self.lift()
            self.focus_force()
            self.attributes("-topmost", True)
            self.after(150, lambda: self.attributes("-topmost", False))
        except Exception:
            pass

    def validate(self):
        idxs = list(self.listbox.curselection())
        if not idxs:
            messagebox.showwarning("Aucune sélection", "Choisis au moins une sheet.", parent=self)
            return False

        self.result = {
            "selections": [
                (self.items[i]["file_path"], self.items[i]["sheet_name"])
                for i in idxs
            ],
            "merged_sheet_name": self.name_entry.get().strip() or "Merged_Measures",
            "add_source_columns": self.add_source_var.get(),
        }
        return True

class WorkbookPlotterWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Workbook Nyquist Plotter")
        self.geometry("1100x720")
        self.minsize(980, 620)

        self.neg_im = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value="Prêt.")
        self.show_legend = tk.BooleanVar(value=True)

        # 3 slots max
        default_colors = ["tab:blue", "tab:orange", "tab:green"]

        self.slots = []
        for i in range(3):
            self.slots.append({
                "path": tk.StringVar(value=""),
                "label": tk.StringVar(value=f"Workbook {i+1}"),
                "enabled": tk.BooleanVar(value=True),
                "base_color": default_colors[i],
                "measures": [],   # liste de dicts {sheet_name, df}
            })

        self.build_ui()
        self.lift()
        self.focus_force()

    def build_ui(self):
        top = tk.Frame(self)
        top.pack(fill="x", padx=12, pady=10)

        tk.Label(
            top,
            text="Sélectionne jusqu'à 3 classeurs Excel et superpose toutes les mesures EIS détectées :"
        ).grid(row=0, column=0, columnspan=6, sticky="w")

        for i in range(3):
            row = i + 1

            tk.Label(top, text=f"Classeur {i+1}:").grid(row=row, column=0, sticky="w", pady=(6, 0))

            tk.Entry(
                top,
                textvariable=self.slots[i]["path"],
                state="readonly",
                width=70
            ).grid(row=row, column=1, sticky="we", padx=(6, 6), pady=(6, 0))

            tk.Button(
                top,
                text="Choisir...",
                command=lambda idx=i: self.choose_excel_workbook(idx)
            ).grid(row=row, column=2, padx=(0, 6), pady=(6, 0))

            tk.Button(
                top,
                text="X",
                width=3,
                command=lambda idx=i: self.clear_workbook(idx)
            ).grid(row=row, column=3, padx=(0, 8), pady=(6, 0))

            tk.Entry(
                top,
                textvariable=self.slots[i]["label"],
                width=18
            ).grid(row=row, column=4, sticky="we", pady=(6, 0))

            tk.Checkbutton(
                top,
                text="Actif",
                variable=self.slots[i]["enabled"],
                command=self.refresh_plot_if_possible
            ).grid(row=row, column=5, sticky="w", padx=(8, 0), pady=(6, 0))

        top.columnconfigure(1, weight=1)

        opts = tk.Frame(self)
        opts.pack(fill="x", padx=12, pady=(0, 6))

        tk.Checkbutton(
            opts,
            text="Tracer -Im (décoché = tracer Im)",
            variable=self.neg_im,
            command=self.refresh_plot_if_possible
        ).pack(side="left")

        tk.Checkbutton(
            opts,
            text="Afficher légende",
            variable=self.show_legend,
            command=self.refresh_plot_if_possible
        ).pack(side="left", padx=16)

        tk.Button(
            opts,
            text="Tracer / Mettre à jour",
            width=18,
            command=self.run_plot
        ).pack(side="right")

        tk.Label(self, textvariable=self.status, anchor="w").pack(fill="x", padx=12, pady=(0, 6))

        plot_frame = tk.Frame(self, bd=1, relief="groove")
        plot_frame.pack(fill="both", expand=True, padx=12, pady=10)

        self.fig = Figure(figsize=(7, 5), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_title("Nyquist (Excel workbooks)")
        self.ax.set_xlabel("Re(Z) [Ohm]")
        self.ax.set_ylabel("-Im(Z) [Ohm]")
        self.ax.grid(True)

        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        toolbar = NavigationToolbar2Tk(self.canvas, plot_frame)
        toolbar.update()
        toolbar.pack(fill="x")

        self.canvas.draw()

    def choose_excel_workbook(self, idx: int):
        path = filedialog.askopenfilename(
            parent=self,
            title=f"Sélectionner le classeur {idx+1}",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("Tous les fichiers", "*.*")]
        )
        if not path:
            return

        self.slots[idx]["path"].set(path)

        # label par défaut = nom de fichier sans dernière extension
        base = os.path.splitext(os.path.basename(path))[0]
        if not self.slots[idx]["label"].get().strip() or self.slots[idx]["label"].get().startswith("Workbook "):
            self.slots[idx]["label"].set(base)

        self.slots[idx]["measures"] = []
        self.status.set("Classeur sélectionné. Clique sur Tracer / Mettre à jour.")

    def clear_workbook(self, idx: int):
        self.slots[idx]["path"].set("")
        self.slots[idx]["measures"] = []
        self.status.set(f"Classeur {idx+1} supprimé.")
        self.refresh_plot_if_possible()

    def run_plot(self):
        any_loaded = False
        msgs = []

        for i, slot in enumerate(self.slots):
            path = slot["path"].get().strip()
            if not path:
                slot["measures"] = []
                continue

            try:
                measures = load_excel_measures(path)
                slot["measures"] = measures

                if measures:
                    any_loaded = True
                    msgs.append(f"C{i+1}: {len(measures)} mesure(s)")
                else:
                    msgs.append(f"C{i+1}: aucune feuille EIS valide")

            except Exception as e:
                slot["measures"] = []
                msgs.append(f"C{i+1}: erreur ({e})")

        if not any_loaded:
            messagebox.showwarning("Aucune donnée", "Aucun classeur exploitable chargé.", parent=self)
            self.clear_plot()
            self.status.set("Aucune donnée valide.")
            return

        self.draw_overlay()
        self.status.set(" | ".join(msgs) + " — Tracé OK.")

    def clear_plot(self):
        self.ax.clear()
        use_neg = self.neg_im.get()
        self.ax.set_title("Nyquist (Excel workbooks)")
        self.ax.set_xlabel("Re(Z) [Ohm]")
        self.ax.set_ylabel("-Im(Z) [Ohm]" if use_neg else "Im(Z) [Ohm]")
        self.ax.grid(True)
        self.canvas.draw()

    def draw_overlay(self):
        use_neg = self.neg_im.get()
        self.ax.clear()

        plotted = 0

        for i, slot in enumerate(self.slots):
            if not slot["enabled"].get():
                continue

            measures = slot["measures"]
            if not measures:
                continue

            label_prefix = slot["label"].get().strip() or f"Workbook {i+1}"
            colors = gradient_colors(slot["base_color"], min(len(measures), 6))

            for j, item in enumerate(measures):
                df = item["df"]
                sheet_name = item["sheet_name"]

                color = colors[min(j, len(colors) - 1)]
                y = -df["Im"] if use_neg else df["Im"]

                self.ax.plot(
                    df["R"],
                    y,
                    linewidth=1.1,
                    color=color,
                    label=f"{label_prefix} | {sheet_name}"
                )
                self.ax.scatter(
                    df["R"],
                    y,
                    s=10,
                    color=color
                )
                plotted += 1

        self.ax.set_title("Nyquist (Excel workbooks)")
        self.ax.set_xlabel("Re(Z) [Ohm]")
        self.ax.set_ylabel("-Im(Z) [Ohm]" if use_neg else "Im(Z) [Ohm]")
        self.ax.grid(True)

        if plotted > 0 and self.show_legend.get():
            self.ax.legend(fontsize=8)

        self.ax.set_aspect("equal", adjustable="datalim")
        self.fig.tight_layout()
        self.canvas.draw()

    def refresh_plot_if_possible(self):
        if any(len(slot["measures"]) > 0 for slot in self.slots):
            self.draw_overlay()

class SavgolDialog(simpledialog.Dialog):
    def __init__(self, parent):
        self.result = None
        super().__init__(parent, title="Paramètres Savitzky-Golay")

    def body(self, master):
        self.after(50, self._raise_front)

        tk.Label(master, text="Window size :").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.window_entry = tk.Entry(master)
        self.window_entry.insert(0, "11")
        self.window_entry.grid(row=0, column=1, sticky="we", padx=6, pady=6)

        tk.Label(master, text="Polynomial order :").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.poly_entry = tk.Entry(master)
        self.poly_entry.insert(0, "3")
        self.poly_entry.grid(row=1, column=1, sticky="we", padx=6, pady=6)

        master.columnconfigure(1, weight=1)
        return self.window_entry

    def _raise_front(self):
        try:
            self.lift()
            self.focus_force()
            self.attributes("-topmost", True)
            self.after(150, lambda: self.attributes("-topmost", False))
        except Exception:
            pass

    def validate(self):
        try:
            self.result = {
                "window_size": int(self.window_entry.get().strip()),
                "poly_order": int(self.poly_entry.get().strip()),
            }
            return True
        except Exception as e:
            messagebox.showerror("Erreur", f"Paramètres invalides.\n\n{e}", parent=self)
            return False

class LinKKDialog(simpledialog.Dialog):
    def __init__(self, parent):
        self.result = None
        super().__init__(parent, title="Paramètres LinKK")

    def body(self, master):
        self.after(50, self._raise_front)

        tk.Label(master, text="fit_type :").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.fit_type_var = tk.StringVar(value="complex")
        tk.OptionMenu(master, self.fit_type_var, "complex", "real", "imag").grid(
            row=0, column=1, sticky="we", padx=6, pady=6
        )

        tk.Label(master, text="M_max :").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.mmax_entry = tk.Entry(master)
        self.mmax_entry.insert(0, "100")
        self.mmax_entry.grid(row=1, column=1, sticky="we", padx=6, pady=6)

        tk.Label(master, text="c :").grid(row=2, column=0, sticky="w", padx=6, pady=6)
        self.c_entry = tk.Entry(master)
        self.c_entry.insert(0, "0.85")
        self.c_entry.grid(row=2, column=1, sticky="we", padx=6, pady=6)

        master.columnconfigure(1, weight=1)
        return self.mmax_entry

    def _raise_front(self):
        try:
            self.lift()
            self.focus_force()
            self.attributes("-topmost", True)
            self.after(150, lambda: self.attributes("-topmost", False))
        except Exception:
            pass

    def validate(self):
        try:
            self.result = {
                "fit_type": self.fit_type_var.get(),
                "M_max": int(self.mmax_entry.get().strip()),
                "c": float(self.c_entry.get().strip().replace(",", ".")),
            }
            return True
        except Exception as e:
            messagebox.showerror("Erreur", f"Paramètres invalides.\n\n{e}", parent=self)
            return False

class DRTDialog(simpledialog.Dialog):
    def __init__(self, parent):
        self.result = None
        super().__init__(parent, title="DRT parameters")

    def body(self, master):
        self.after(50, self._raise_front)

        tk.Label(master, text="Method of discretisation:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.rbf_var = tk.StringVar(value="Gaussian")
        tk.OptionMenu(master, self.rbf_var, "Gaussian", "C0 Matern", "C2 Matern", "C4 Matern", "C6 Matern").grid(
            row=0, column=1, sticky="we", padx=6, pady=6
        )

        tk.Label(master, text="Data used:").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.data_used_var = tk.StringVar(value="Combined Re-Im Data")
        tk.OptionMenu(master, self.data_used_var, "Combined Re-Im Data", "Re Data", "Im Data").grid(
            row=1, column=1, sticky="we", padx=6, pady=6
        )

        tk.Label(master, text="Inductance setting:").grid(row=2, column=0, sticky="w", padx=6, pady=6)
        self.induct_var = tk.StringVar(value="2")
        tk.OptionMenu(master, self.induct_var, "0", "1", "2").grid(
            row=2, column=1, sticky="we", padx=6, pady=6
        )

        tk.Label(master, text="Regularization derivative:").grid(row=3, column=0, sticky="w", padx=6, pady=6)
        self.der_var = tk.StringVar(value="2nd order")
        tk.OptionMenu(master, self.der_var, "1st order", "2nd order").grid(
            row=3, column=1, sticky="we", padx=6, pady=6
        )

        tk.Label(master, text="Parameter selection method:").grid(row=4, column=0, sticky="w", padx=6, pady=6)
        self.cv_var = tk.StringVar(value="GCV")
        tk.OptionMenu(master, self.cv_var, "GCV", "mGCV", "rGCV", "LC").grid(
            row=4, column=1, sticky="we", padx=6, pady=6
        )

        tk.Label(master, text="Regularization parameter:").grid(row=5, column=0, sticky="w", padx=6, pady=6)
        self.reg_entry = tk.Entry(master)
        self.reg_entry.insert(0, "1e-3")
        self.reg_entry.grid(row=5, column=1, sticky="we", padx=6, pady=6)

        self.run_twice_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            master,
            text="Run second pass with optimal parameter",
            variable=self.run_twice_var
        ).grid(row=6, column=0, columnspan=2, sticky="w", padx=6, pady=6)

        master.columnconfigure(1, weight=1)
        return self.reg_entry

    def _raise_front(self):
        try:
            self.lift()
            self.focus_force()
            self.attributes("-topmost", True)
            self.after(150, lambda: self.attributes("-topmost", False))
        except Exception:
            pass

    def validate(self):
        try:
            self.result = {
                "rbf_type": self.rbf_var.get(),
                "data_used": self.data_used_var.get(),
                "induct_used": int(self.induct_var.get()),
                "der_used": self.der_var.get(),
                "cv_type": self.cv_var.get(),
                "reg_param": float(self.reg_entry.get().strip().replace(",", ".")),
                "shape_control": "FWHM Coefficient",
                "coeff": 0.5,
                "run_twice": self.run_twice_var.get(),
            }
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Invalid parameters.\n\n{e}", parent=self)
            return False

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

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.build_menus()
        self.title("Comparaison Nyquist (jusqu'à 3 machines)")
        self.geometry("980x650")
        self.minsize(880, 560)

        self.neg_im = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value="Prêt.")

        # 3 slots: path, label, enabled, df cache
        self.slots = []
        for _ in range(3):
            self.slots.append({
                "path": tk.StringVar(value=""),
                "label": tk.StringVar(value=""),
                "enabled": tk.BooleanVar(value=True),
                "df": None
            })

        # ===== Controls =====
        top = tk.Frame(self)
        top.pack(fill="x", padx=12, pady=10)

        tk.Label(top, text="Sélectionne jusqu'à 3 fichiers (freq R Im) et superpose les plots :").grid(
            row=0, column=0, columnspan=5, sticky="w"
        )

        for i in range(3):
            row = 1 + i
            tk.Label(top, text=f"Fichier {i+1}:").grid(row=row, column=0, sticky="w", pady=(6, 0))

            entry = tk.Entry(top, textvariable=self.slots[i]["path"], state="readonly", width=70)
            entry.grid(row=row, column=1, sticky="we", padx=(6, 6), pady=(6, 0))

            tk.Button(top, text="Choisir...", command=lambda idx=i: self.choose_file(idx)).grid(
                row=row, column=2, padx=(0, 6), pady=(6, 0)
            )

            tk.Button(top, text="X", width=3, command=lambda idx=i: self.clear_file(idx)).grid(
                row=row, column=3, padx=(0, 10), pady=(6, 0)
            )

            lbl_entry = tk.Entry(top, textvariable=self.slots[i]["label"], width=18)
            lbl_entry.grid(row=row, column=4, sticky="e", pady=(6, 0))
            if not self.slots[i]["label"].get():
                self.slots[i]["label"].set(f"Machine {i+1}")

            tk.Checkbutton(top, text="Actif", variable=self.slots[i]["enabled"],
                           command=self.refresh_plot_if_possible).grid(
                row=row, column=5, sticky="w", padx=(8, 0), pady=(6, 0)
            )

        top.columnconfigure(1, weight=1)

        opts = tk.Frame(self)
        opts.pack(fill="x", padx=12, pady=(0, 6))

        tk.Checkbutton(
            opts,
            text="Tracer -Im (décoché = tracer Im)",
            variable=self.neg_im,
            command=self.refresh_plot_if_possible
        ).pack(side="left")

        tk.Button(opts, text="Tracer / Mettre à jour", width=18, command=self.run_plot).pack(side="right")
        tk.Button(opts, text="Exporter PNG", width=14, command=self.export_png).pack(side="right", padx=8)

        tk.Label(self, textvariable=self.status, anchor="w").pack(fill="x", padx=12, pady=(0, 6))

        # ===== Plot embed =====
        plot_frame = tk.Frame(self, bd=1, relief="groove")
        plot_frame.pack(fill="both", expand=True, padx=12, pady=10)

        self.fig = Figure(figsize=(6, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_title("Nyquist (scatter) : Re vs -Im")
        self.ax.set_xlabel("Re (R)")
        self.ax.set_ylabel("-Im")
        self.ax.grid(True)

        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        toolbar = NavigationToolbar2Tk(self.canvas, plot_frame)
        toolbar.update()
        toolbar.pack(fill="x")

        self.canvas.draw()

        # Markers différents (sans imposer de couleurs)
        self.markers = ["o", "s", "^"]

    def build_menus(self):
        menubar = tk.Menu(self)

        menu_file = tk.Menu(menubar, tearoff=0)
        menu_file.add_command(label="Ouvrir fichier 1...", command=lambda: self.choose_file(0))
        menu_file.add_command(label="Ouvrir fichier 2...", command=lambda: self.choose_file(1))
        menu_file.add_command(label="Ouvrir fichier 3...", command=lambda: self.choose_file(2))

        menu_file.add_separator()
        menu_file.add_command(label="Exporter slot 1...", command=lambda: self.export_slot_data(0))
        menu_file.add_command(label="Exporter slot 2...", command=lambda: self.export_slot_data(1))
        menu_file.add_command(label="Exporter slot 3...", command=lambda: self.export_slot_data(2))

        menu_file.add_separator()
        menu_file.add_command(label="Exporter PNG...", command=self.export_png)
        menu_file.add_separator()
        menu_file.add_command(label="Quitter", command=self.destroy)

        menu_tools = tk.Menu(menubar, tearoff=0)
        menu_tools.add_command(label="Supprimer diffusion Warburg (Wo)...", command=self.run_warburg_tool)
        menu_tools.add_separator()
        menu_tools.add_command(label="EIS Preprocessing Studio...", command=self.open_preprocessing_studio)
        menu_tools.add_separator()
        menu_tools.add_command(label="Batch loops -> Export Excel...", command=self.run_batch_export_tool)
        menu_tools.add_command(label="Rassembler des classeurs Excel...", command=self.run_collect_workbooks_tool)
        menu_tools.add_separator()
        menu_tools.add_command(label="Savitzky-Golay sur un classeur Excel...", command=self.run_savgol_tool)
        menu_tools.add_command(label="LinKK sur un classeur Excel...", command=self.run_linkk_tool)
        menu_tools.add_command(label="DRT sur un classeur Excel...", command=self.run_drt_tool)
        menu_tools.add_separator()
        menu_tools.add_command(label="Plot classeurs Excel...", command=self.open_workbook_plotter)
        menu_tools.add_command(label="Plot DRT Excel...", command=self.open_drt_plotter)

        menubar.add_cascade(label="Fichier", menu=menu_file)
        menubar.add_cascade(label="Tools", menu=menu_tools)

        self.config(menu=menubar)

    def choose_file(self, idx: int):
        path = filedialog.askopenfilename(
            title=f"Sélectionner le fichier {idx+1}",
            filetypes=[
                ("Fichiers texte", "*.txt *.dat *.csv *.tsv"),
                ("Tous les fichiers", "*.*")
            ],
        )
        if path:
            self.slots[idx]["path"].set(path)
            if not self.slots[idx]["label"].get().strip():
                self.slots[idx]["label"].set(os.path.splitext(os.path.basename(path))[0])
            self.slots[idx]["df"] = None  # reset cache
            self.status.set("Fichier sélectionné. Clique sur Tracer / Mettre à jour.")

    def choose_excel_workbook(self, title="Choisir un fichier Excel"):
        self.bring_to_front()
        return filedialog.askopenfilename(
            parent=self,
            title=title,
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")],
        )

    def clear_file(self, idx: int):
        self.slots[idx]["path"].set("")
        self.slots[idx]["df"] = None
        self.status.set(f"Fichier {idx+1} supprimé. Clique sur Tracer / Mettre à jour.")
        self.refresh_plot_if_possible()

    def run_plot(self):
        any_loaded = False
        msgs = []

        # Load / cache DFs
        for i, slot in enumerate(self.slots):
            path = slot["path"].get().strip()
            if not path:
                slot["df"] = None
                continue

            try:
                df = read_freq_R_Im(path)
                if df.empty:
                    slot["df"] = None
                    msgs.append(f"F{i+1}: aucun point valide")
                else:
                    slot["df"] = df
                    any_loaded = True
                    msgs.append(f"F{i+1}: {len(df)} points")
            except Exception as e:
                slot["df"] = None
                msgs.append(f"F{i+1}: erreur lecture ({e})")

        if not any_loaded:
            messagebox.showwarning("Aucune donnée", "Aucun fichier valide chargé.")
            self.clear_plot()
            self.status.set("Aucune donnée valide.")
            return

        self.draw_overlay()
        self.status.set(" | ".join(msgs) + "  —  Tracé OK.")

    def clear_plot(self):
        self.ax.clear()
        use_neg = self.neg_im.get()
        self.ax.set_title("Nyquist (scatter) : Re vs " + ("-Im" if use_neg else "Im"))
        self.ax.set_xlabel("Re(Z) [Ohm]")
        self.ax.set_ylabel("-Im(Z) [Ohm]" if use_neg else "Im(Z) [Ohm]")
        self.ax.grid(True)
        self.canvas.draw()

    def draw_overlay(self):
        use_neg = self.neg_im.get()
        self.ax.clear()

        plotted = 0
        for i, slot in enumerate(self.slots):
            if not slot["enabled"].get():
                continue
            df = slot["df"]
            if df is None or df.empty:
                continue

            y = -df["Im"] if use_neg else df["Im"]
            label = slot["label"].get().strip() or f"Machine {i + 1}"

            self.ax.scatter(df["R"], y, s=14, marker=self.markers[i], label=label)
            plotted += 1

        self.ax.set_title("Nyquist (scatter) : Re vs " + ("-Im" if use_neg else "Im"))
        self.ax.set_xlabel("Re(Z) [Ohm]")
        self.ax.set_ylabel("-Im(Z) [Ohm]" if use_neg else "Im(Z) [Ohm]")
        self.ax.grid(True)

        if plotted > 0:
            self.ax.legend()

        self.ax.set_aspect("equal", adjustable="datalim")
        self.fig.tight_layout()
        self.canvas.draw()

    def refresh_plot_if_possible(self):
        # Redessine si au moins un df est chargé (sans recharger les fichiers)
        if any(slot["df"] is not None and not slot["df"].empty for slot in self.slots):
            self.draw_overlay()

    def export_png(self):
        if not any(slot["df"] is not None and not slot["df"].empty and slot["enabled"].get() for slot in self.slots):
            messagebox.showinfo("Rien à exporter", "Veuillez d'abord tracer au moins un plot actif.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Enregistrer en PNG",
            defaultextension=".png",
            filetypes=[("Image PNG", "*.png")],
            initialfile="nyquist_compare.png",
        )
        if not save_path:
            return

        try:
            self.fig.savefig(save_path, dpi=300, bbox_inches="tight")
            self.status.set(f"Export PNG OK: {save_path}")
            messagebox.showinfo("Export réussi", "Le PNG a été enregistré avec succès.")
        except Exception as e:
            messagebox.showerror("Erreur export", f"Impossible d'exporter en PNG.\n\nDétail:\n{e}")

    def run_warburg_tool(self):
        available = [
            i + 1
            for i, slot in enumerate(self.slots)
            if slot["df"] is not None and not slot["df"].empty
        ]

        if not available:
            messagebox.showwarning(
                "Aucune donnée",
                "Charge d'abord au moins un fichier valide avant d'utiliser l'outil Warburg."
            )
            return

        dlg = WarburgDialog(self, available)
        params = dlg.result
        if not params:
            return

        idx = params["slot_idx"]
        slot = self.slots[idx]

        if slot["df"] is None or slot["df"].empty:
            messagebox.showwarning("Slot vide", f"Le slot {idx + 1} ne contient aucune donnée.")
            return

        try:
            df_corr = remove_warburg_wo(
                slot["df"],
                wor=params["wor"],
                wot=params["wot"],
                wop=params["wop"],
                fmin=params["fmin"],
                trim_pos_im=params["trim_pos_im"],
            )

            if df_corr.empty:
                messagebox.showwarning(
                    "Résultat vide",
                    "La correction a produit un tableau vide. Vérifie les paramètres et filtres."
                )
                return

            # On garde seulement les colonnes normalisées pour l'affichage
            slot["df"] = df_corr[["freq", "R", "Im"]].copy()

            old_label = slot["label"].get().strip() or f"Machine {idx + 1}"
            if "noWo" not in old_label:
                slot["label"].set(old_label + " | noWo")

            self.draw_overlay()
            self.status.set(
                f"Suppression Wo appliquée au fichier {idx + 1} "
                f"(Wo-R={params['wor']}, Wo-T={params['wot']}, Wo-P={params['wop']})."
            )

        except Exception as e:
            messagebox.showerror("Erreur correction Warburg", f"Impossible d'appliquer la correction.\n\nDétail:\n{e}")

    def run_batch_export_tool(self):
        available = [
            i + 1
            for i, slot in enumerate(self.slots)
            if slot["path"].get().strip()
        ]

        if not available:
            messagebox.showwarning(
                "Aucune donnée",
                "Charge d'abord au moins un fichier valide.",
                parent=self
            )
            return

        self.bring_to_front()
        dlg = BatchExportDialog(self, available)
        params = dlg.result
        if not params:
            return

        idx = params["slot_idx"]
        path = self.slots[idx]["path"].get().strip()

        if not path:
            messagebox.showwarning(
                "Slot vide",
                f"Le slot {idx + 1} n'a pas de fichier associé.",
                parent=self
            )
            return

        base_name = Path(path).stem  # enlève seulement la dernière extension
        safe_base_name = "".join(
            c if c.isalnum() or c in ("-", "_", " ", "(", ")") else "_"
            for c in base_name
        ).strip()

        if not safe_base_name:
            safe_base_name = f"slot_{idx + 1}"

        self.bring_to_front()
        save_path = filedialog.asksaveasfilename(
            parent=self,
            title="Exporter les mesures en Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"{safe_base_name}_batched.xlsx",
        )
        if not save_path:
            return

        try:
            raw_df, meta = read_eis_raw_table(path)

            batches = split_raw_eis_by_start_freq(
                raw_df=raw_df,
                freq_col=meta["freq_col"],
                start_freq_hz=params["start_freq_hz"],
                rel_tol=params["rel_tol"],
                jump_ratio=params["jump_ratio"],
                min_points=params["min_points"],
            )

            if not batches:
                messagebox.showwarning(
                    "Aucune mesure détectée",
                    "Aucune boucle n'a été détectée avec ces paramètres.",
                    parent=self
                )
                return

            export_raw_eis_batches_to_excel(
                batches=batches,
                out_path=save_path,
                source_path=path,
                freq_col=meta["freq_col"],
            )

            self.status.set(f"Batch export OK: {len(batches)} mesures exportées vers {save_path}")
            messagebox.showinfo(
                "Export réussi",
                f"{len(batches)} mesures détectées et exportées.\n\n{save_path}",
                parent=self
            )

        except Exception as e:
            messagebox.showerror(
                "Erreur batch export",
                f"Impossible de découper/exporter les mesures.\n\nDétail:\n{e}",
                parent=self
            )

    def run_collect_workbooks_tool(self):
        self.bring_to_front()

        file_paths = filedialog.askopenfilenames(
            parent=self,
            title="Choisir les fichiers Excel à rassembler",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")],
        )
        if not file_paths:
            return

        self.bring_to_front()
        save_path = filedialog.asksaveasfilename(
            parent=self,
            title="Enregistrer le classeur rassemblé",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="collected_workbooks.xlsx",
        )
        if not save_path:
            return

        try:
            collect_excel_workbooks(
                file_paths=list(file_paths),
                out_path=save_path,
                copy_styles=True,
            )

            self.status.set(f"Rassemblement Excel OK: {save_path}")
            messagebox.showinfo(
                "Rassemblement réussi",
                f"Toutes les feuilles ont été regroupées dans un seul classeur.\n\n{save_path}",
                parent=self
            )

        except Exception as e:
            messagebox.showerror(
                "Erreur rassemblement Excel",
                f"Impossible de rassembler les classeurs.\n\nDétail:\n{e}",
                parent=self
            )

    def bring_to_front(self):
        try:
            self.lift()
            self.focus_force()
            self.attributes("-topmost", True)
            self.after(150, lambda: self.attributes("-topmost", False))
        except Exception:
            pass

    def run_background_job(self, job_func, done_message: str, error_title: str):
        """
        Exécute un traitement long dans un thread séparé
        puis met à jour l'UI proprement.
        """

        def worker():
            try:
                result = job_func()

                def on_success():
                    self.status.set(done_message.format(result=result))
                    messagebox.showinfo("Terminé", done_message.format(result=result), parent=self)

                self.after(0, on_success)

            except Exception as e:
                def on_error():
                    self.status.set(f"Erreur: {e}")
                    messagebox.showerror(error_title, f"{e}", parent=self)

                self.after(0, on_error)

        threading.Thread(target=worker, daemon=True).start()

    def run_savgol_tool(self):
        path = self.choose_excel_workbook("Choisir le classeur Excel à lisser")
        if not path:
            return

        self.bring_to_front()
        dlg = SavgolDialog(self)
        params = dlg.result
        if not params:
            return

        self.status.set("Savitzky-Golay en cours...")

        def job():
            return savgol_workbook(
                excel_file_path=path,
                window_size=params["window_size"],
                poly_order=params["poly_order"],
            )

        self.run_background_job(
            job_func=job,
            done_message="Lissage terminé.\n\nFichier créé :\n{result}",
            error_title="Erreur Savitzky-Golay",
        )

    def run_linkk_tool(self):
        path = self.choose_excel_workbook("Choisir le classeur Excel pour LinKK")
        if not path:
            return

        self.bring_to_front()
        dlg = LinKKDialog(self)
        params = dlg.result
        if not params:
            return

        self.status.set("LinKK en cours...")

        def job():
            return linkk_workbook(
                excel_file_path=path,
                fit_type=params["fit_type"],
                M_max=params["M_max"],
                c=params["c"],
            )

        self.run_background_job(
            job_func=job,
            done_message="Validation LinKK terminée.\n\nFichier créé :\n{result}",
            error_title="Erreur LinKK",
        )

    def run_drt_tool(self):
        path = self.choose_excel_workbook("Choisir le classeur Excel pour la DRT")
        if not path:
            return

        self.bring_to_front()
        dlg = DRTDialog(self)
        params = dlg.result
        if not params:
            return

        self.status.set("DRT en cours...")

        def job():
            return run_drt_workbook(
                excel_file_path=path,
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

        self.run_background_job(
            job_func=job,
            done_message="DRT terminée.\n\nFichier créé :\n{result}",
            error_title="Erreur DRT",
        )

    def open_workbook_plotter(self):
        WorkbookPlotterWindow(self)

    def open_drt_plotter(self):
        DRTPlotterWindow(self)

    def open_preprocessing_studio(self):
        EISPreprocessingWindow(self)

if __name__ == "__main__":
    App().mainloop()
