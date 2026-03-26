import tkinter as tk
from tkinter import messagebox, simpledialog


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

