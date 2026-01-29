import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

import matplotlib
matplotlib.use("TkAgg")  # force un rendu intégré Tkinter (évite les docks PyCharm)
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
        self.title("Nyquist intégré - Re vs (+/-)Im")
        self.geometry("900x600")
        self.minsize(800, 520)

        self.file_path = tk.StringVar(value="Aucun fichier sélectionné")
        self.neg_im = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value="Prêt.")
        self.last_df = None

        # ====== Haut (contrôles) ======
        top = tk.Frame(self)
        top.pack(fill="x", padx=12, pady=10)

        tk.Label(top, text="Fichier (freq R Im):").grid(row=0, column=0, sticky="w")
        entry = tk.Entry(top, textvariable=self.file_path, state="readonly", width=80)
        entry.grid(row=1, column=0, sticky="we", pady=(4, 0))
        tk.Button(top, text="Choisir...", command=self.choose_file).grid(row=1, column=1, padx=(10, 0), pady=(4, 0))

        tk.Checkbutton(
            top,
            text="Tracer -Im (décoché = tracer Im)",
            variable=self.neg_im,
            command=self.refresh_plot_if_possible
        ).grid(row=2, column=0, sticky="w", pady=(8, 0))

        btns = tk.Frame(top)
        btns.grid(row=2, column=1, sticky="e", pady=(8, 0))
        tk.Button(btns, text="Tracer", width=12, command=self.run_plot).pack(side="left")
        tk.Button(btns, text="Exporter PNG", width=12, command=self.export_png).pack(side="left", padx=8)

        top.columnconfigure(0, weight=1)

        tk.Label(self, textvariable=self.status, anchor="w").pack(fill="x", padx=12)

        # ====== Zone graphique (embed) ======
        plot_frame = tk.Frame(self, bd=1, relief="groove")
        plot_frame.pack(fill="both", expand=True, padx=12, pady=10)

        self.fig = Figure(figsize=(6, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_title("Nyquist (scatter)")
        self.ax.set_xlabel("Re (R)")
        self.ax.set_ylabel("-Im")
        self.ax.grid(True)

        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        toolbar = NavigationToolbar2Tk(self.canvas, plot_frame)
        toolbar.update()

        # Astuce: rendre le layout plus propre
        toolbar.pack(fill="x")

        self.canvas.draw()

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="Sélectionner un fichier",
            filetypes=[
                ("Fichiers texte", "*.txt *.dat *.csv *.tsv"),
                ("Tous les fichiers", "*.*")
            ],
        )
        if path:
            self.file_path.set(path)
            self.status.set("Fichier sélectionné. Cliquez sur Tracer.")
            self.last_df = None

    def run_plot(self):
        path = self.file_path.get()
        if not path or path == "Aucun fichier sélectionné":
            messagebox.showwarning("Fichier manquant", "Veuillez sélectionner un fichier.")
            return

        try:
            df = read_3col_file(path)
            if df.empty:
                messagebox.showerror("Erreur", "Aucune donnée valide détectée (R/Im).")
                return

            self.last_df = df
            self.draw_nyquist(df)
            self.status.set(f"Tracé OK ({len(df)} points).")

        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire/Tracer le fichier.\n\nDétail:\n{e}")

    def draw_nyquist(self, df: pd.DataFrame):
        use_neg = self.neg_im.get()
        y = -df["Im"] if use_neg else df["Im"]

        self.ax.clear()
        self.ax.scatter(df["R"], y, s=14)
        self.ax.set_xlabel("Re (R)")
        self.ax.set_ylabel("-Im" if use_neg else "Im")
        self.ax.set_title("Nyquist (scatter) : Re vs " + ("-Im" if use_neg else "Im"))
        self.ax.grid(True)
        self.ax.set_aspect("equal", adjustable="datalim")

        self.fig.tight_layout()
        self.canvas.draw()

    def refresh_plot_if_possible(self):
        if self.last_df is not None and not self.last_df.empty:
            self.draw_nyquist(self.last_df)

    def export_png(self):
        if self.last_df is None:
            messagebox.showinfo("Rien à exporter", "Veuillez d'abord tracer le graphique.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Enregistrer en PNG",
            defaultextension=".png",
            filetypes=[("Image PNG", "*.png")],
            initialfile="nyquist.png",
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
