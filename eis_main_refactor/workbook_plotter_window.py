
import os
import tkinter as tk
from tkinter import filedialog, messagebox

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from drt_core import gradient_colors, load_excel_measures


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
