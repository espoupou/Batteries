import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk


def read_3col_file(path: str) -> pd.DataFrame:
    """Lit un fichier avec 3 colonnes: freq R Im (espaces/tabs), tolère header, virgules décimales."""
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        first = f.readline().strip().split()

    def is_num(x: str) -> bool:
        try:
            float(x.replace(",", "."))
            return True
        except Exception:
            return False

    has_header = not (len(first) >= 3 and all(is_num(x) for x in first[:3]))

    df = pd.read_csv(
        path,
        sep=r"\s+",
        header=0 if has_header else None,
        names=["freq", "R", "Im"] if not has_header else None,
        dtype=str,
        engine="python",
    )

    df = df.iloc[:, :3]
    df.columns = ["freq", "R", "Im"]

    for c in ["freq", "R", "Im"]:
        df[c] = df[c].astype(str).str.replace(",", ".", regex=False)
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df.dropna(subset=["R", "Im"]).reset_index(drop=True)
    return df


class App(tk.Tk):
    def __init__(self):
        super().__init__()
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
                df = read_3col_file(path)
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
        self.ax.set_title("Nyquist (scatter) : Re vs " + ("Im" if use_neg else "-Im"))
        self.ax.set_xlabel("Re (R)")
        self.ax.set_ylabel("-Im" if use_neg else "Im")
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
            label = slot["label"].get().strip() or f"Machine {i+1}"

            self.ax.scatter(df["R"], y, s=14, marker=self.markers[i], label=label)
            plotted += 1

        self.ax.set_title("Nyquist (scatter) : Re vs " + ("Im" if use_neg else "-Im"))
        self.ax.set_xlabel("Re (R)")
        self.ax.set_ylabel("-Im" if use_neg else "Im")
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


if __name__ == "__main__":
    App().mainloop()
