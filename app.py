
import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from dialogs import (
    BatchExportDialog,
    DRTDialog,
    LinKKDialog,
    SavgolDialog,
    WarburgDialog,
)
from drt_plotter_window import DRTPlotterWindow
from preprocessing_window import EISPreprocessingWindow
from workbook_plotter_window import WorkbookPlotterWindow

from drt_core import linkk_workbook, run_drt_workbook, savgol_workbook
from eis_core import (
    collect_excel_workbooks,
    export_freq_R_Im,
    export_raw_eis_batches_to_excel,
    read_eis_raw_table,
    read_freq_R_Im,
    remove_warburg_wo,
    split_raw_eis_by_start_freq,
)


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
