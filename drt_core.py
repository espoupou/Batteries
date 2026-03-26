import re

import numpy as np
import pandas as pd
from pathlib import Path
from typing import Optional, Dict, Any
from matplotlib import colors as mcolors

from scipy.signal import savgol_filter

from impedance.validation import linKK
from pyDRTtools.runs import EIS_object, simple_run


# ---------------------------------------------------------
# Helpers
# ---------------------------------------------------------

def _clean_header(s: str) -> str:
    s = str(s).strip().lower()
    s = s.replace("−", "-").replace("–", "-")
    return s


def _find_sheet_eis_columns(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Détecte les colonnes utiles dans une feuille Excel:
        freq, R, Im
    Retourne aussi le signe de la colonne imaginaire:
        im_sign = -1 si la colonne est '-Im'
        im_sign = +1 si la colonne est 'Im'
    """
    freq_col = None
    r_col = None
    im_col = None
    im_sign = +1

    for col in df.columns:
        raw = str(col).strip()
        s = _clean_header(raw)

        # fréquence
        if freq_col is None and ("freq" in s or s == "f"):
            freq_col = col
            continue

        # partie réelle
        if r_col is None:
            if (
                "re(z" in s
                or "real(z" in s
                or "z'" in s
                or "zreal" in s
                or "rez" in s
            ):
                r_col = col
                continue

        # partie imaginaire
        if im_col is None:
            if "-im(z" in s or s.startswith("-im") or "-z''" in s:
                im_col = col
                im_sign = -1
                continue
            if "im(z" in s or "imag(z" in s or "z''" in s or "zimag" in s:
                im_col = col
                im_sign = +1
                continue

    return {
        "freq_col": freq_col,
        "R_col": r_col,
        "Im_col": im_col,
        "im_sign": im_sign,
        "valid": all(v is not None for v in [freq_col, r_col, im_col]),
    }


def extract_eis_from_sheet(df: pd.DataFrame):
    """
    Retourne:
        freq, Z_real, Z_imag, meta

    Convention interne:
        Z_imag = vraie partie imaginaire Im(Z)
    """
    meta = _find_sheet_eis_columns(df)
    if not meta["valid"]:
        raise ValueError("Colonnes EIS introuvables dans cette feuille.")

    freq = pd.to_numeric(
        df[meta["freq_col"]].astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    )
    z_real = pd.to_numeric(
        df[meta["R_col"]].astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    )
    z_im_col = pd.to_numeric(
        df[meta["Im_col"]].astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    )

    valid = freq.notna() & z_real.notna() & z_im_col.notna()

    freq = freq[valid].to_numpy(dtype=float)
    z_real = z_real[valid].to_numpy(dtype=float)
    z_im_col = z_im_col[valid].to_numpy(dtype=float)

    # conversion vers Im(Z) physique
    if meta["im_sign"] == -1:
        z_imag = -z_im_col
    else:
        z_imag = z_im_col

    return freq, z_real, z_imag, meta


def try_get_potential(df: pd.DataFrame) -> Optional[float]:
    candidates = ["<Ewe>/V", "Ewe/V", "Potential/V", "potential"]
    for c in candidates:
        if c in df.columns:
            s = pd.to_numeric(df[c].astype(str).str.replace(",", ".", regex=False), errors="coerce").dropna()
            if not s.empty:
                return float(s.iloc[0])
    return None


def is_valid_savgol_window(n: int, window_size: int, poly_order: int) -> int:
    """
    Ajuste la fenêtre SG pour éviter les erreurs.
    """
    w = int(window_size)
    p = int(poly_order)

    if w < 3:
        w = 3
    if w % 2 == 0:
        w += 1
    if w > n:
        w = n if n % 2 == 1 else n - 1
    if w < 3:
        raise ValueError("Pas assez de points pour Savitzky-Golay.")
    if p >= w:
        p = w - 1
    return w, p


def blend_with_white(color, t: float):
    """
    t=0 -> presque blanc
    t=1 -> couleur de base
    """
    r, g, b = mcolors.to_rgb(color)
    return (
        1 - (1 - r) * t,
        1 - (1 - g) * t,
        1 - (1 - b) * t,
    )


def gradient_colors(base_color: str, n: int):
    if n <= 1:
        return [base_color]
    # du plus clair au plus soutenu
    ts = [0.35 + 0.55 * i / (n - 1) for i in range(n)]
    return [blend_with_white(base_color, t) for t in ts]


def measure_sort_key(sheet_name: str):
    """
    Trie intelligemment les feuilles du type ..._M1, ..._M2, etc.
    Sinon garde un tri texte.
    """
    m = re.search(r"_M(\d+)$", str(sheet_name), flags=re.IGNORECASE)
    if m:
        return (0, int(m.group(1)), str(sheet_name).lower())
    return (1, 10**9, str(sheet_name).lower())


def load_excel_measures(path: str):
    """
    Lit un classeur Excel et retourne uniquement les feuilles exploitables en EIS.
    Chaque mesure est normalisée en colonnes freq, R, Im.
    """
    xls = pd.ExcelFile(path)
    measures = []

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet_name)

        try:
            freq, z_real, z_imag, meta = extract_eis_from_sheet(df)
        except Exception:
            continue

        if len(freq) == 0:
            continue

        measures.append({
            "sheet_name": sheet_name,
            "df": pd.DataFrame({
                "freq": freq,
                "R": z_real,
                "Im": z_imag,
            })
        })

    measures.sort(key=lambda x: measure_sort_key(x["sheet_name"]))
    return measures


def load_drt_measures(path: str):
    xls = pd.ExcelFile(path)
    measures = []
    summary_df = None

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet_name)

        if sheet_name.lower() == "summary":
            summary_df = df
            continue

        if not {"tau_s", "gamma"}.issubset(df.columns):
            continue

        tau = pd.to_numeric(df["tau_s"], errors="coerce")
        gamma = pd.to_numeric(df["gamma"], errors="coerce")
        valid = tau.notna() & gamma.notna()

        if valid.sum() == 0:
            continue

        meta = {}
        for c in ["Potential_V", "R_ohm", "L_H", "lambda_optimal"]:
            if c in df.columns:
                s = pd.to_numeric(df[c], errors="coerce").dropna()
                if not s.empty:
                    meta[c] = float(s.iloc[0])

        measures.append({
            "sheet_name": sheet_name,
            "tau": tau[valid].to_numpy(),
            "gamma": gamma[valid].to_numpy(),
            "meta": meta,
        })

    measures.sort(key=lambda x: measure_sort_key(x["sheet_name"]))
    return measures, summary_df

# ---------------------------------------------------------
# DRTTOOLS
# ---------------------------------------------------------


def savgol_workbook(
    excel_file_path,
    window_size: int = 11,
    poly_order: int = 3,
):
    excel_file_path = Path(excel_file_path)
    xls = pd.ExcelFile(excel_file_path)
    output_file = excel_file_path.with_name(f"{excel_file_path.stem}_Smoothed.xlsx")

    with pd.ExcelWriter(output_file) as writer:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

            try:
                freq, z_real, z_imag, meta = extract_eis_from_sheet(df)
            except Exception:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                continue

            w, p = is_valid_savgol_window(len(freq), window_size, poly_order)

            real_smooth = savgol_filter(z_real, w, p)
            imag_smooth = savgol_filter(z_imag, w, p)

            out = df.copy()
            out["Smoothed_Re(Z)_Ohm"] = np.nan
            out["Smoothed_Im(Z)_Ohm"] = np.nan

            valid = (
                pd.to_numeric(df[meta["freq_col"]].astype(str).str.replace(",", ".", regex=False), errors="coerce").notna()
                & pd.to_numeric(df[meta["R_col"]].astype(str).str.replace(",", ".", regex=False), errors="coerce").notna()
                & pd.to_numeric(df[meta["Im_col"]].astype(str).str.replace(",", ".", regex=False), errors="coerce").notna()
            )

            out.loc[valid, "Smoothed_Re(Z)_Ohm"] = real_smooth
            out.loc[valid, "Smoothed_Im(Z)_Ohm"] = imag_smooth

            # si la feuille source était en -Im, on ajoute aussi cette version
            if meta["im_sign"] == -1:
                out["Smoothed_-Im(Z)_Ohm"] = np.nan
                out.loc[valid, "Smoothed_-Im(Z)_Ohm"] = -imag_smooth

            out.to_excel(writer, sheet_name=sheet_name, index=False)

    return str(output_file)


def linkk_workbook(
    excel_file_path,
    fit_type: str = "complex",
    M_max: int = 100,
    c: float = 0.85,
):
    excel_file_path = Path(excel_file_path)
    xls = pd.ExcelFile(excel_file_path)
    output_file = excel_file_path.with_name(f"{excel_file_path.stem}_LinKK.xlsx")

    with pd.ExcelWriter(output_file) as writer:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

            try:
                freq, z_real, z_imag, meta = extract_eis_from_sheet(df)
            except Exception:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                continue

            z_complex = z_real + 1j * z_imag

            M, mu, Z_fit, residuals_real, residuals_imag = linKK(
                np.array(freq),
                np.array(z_complex),
                c=c,
                max_M=M_max,
                fit_type=fit_type,
                add_cap=True
            )

            out = df.copy()
            out["LinKK_Re(Z)_fit_Ohm"] = np.nan
            out["LinKK_Im(Z)_fit_Ohm"] = np.nan
            out["LinKK_Re_residuals_pct"] = np.nan
            out["LinKK_Im_residuals_pct"] = np.nan
            out["LinKK_Number_of_RC_used"] = np.nan
            out["LinKK_Fitting_coefficient"] = np.nan

            valid = (
                pd.to_numeric(df[meta["freq_col"]].astype(str).str.replace(",", ".", regex=False), errors="coerce").notna()
                & pd.to_numeric(df[meta["R_col"]].astype(str).str.replace(",", ".", regex=False), errors="coerce").notna()
                & pd.to_numeric(df[meta["Im_col"]].astype(str).str.replace(",", ".", regex=False), errors="coerce").notna()
            )

            out.loc[valid, "LinKK_Re(Z)_fit_Ohm"] = Z_fit.real
            out.loc[valid, "LinKK_Im(Z)_fit_Ohm"] = Z_fit.imag
            out.loc[valid, "LinKK_Re_residuals_pct"] = residuals_real
            out.loc[valid, "LinKK_Im_residuals_pct"] = residuals_imag
            out.loc[valid, "LinKK_Number_of_RC_used"] = M
            out.loc[valid, "LinKK_Fitting_coefficient"] = mu

            out.to_excel(writer, sheet_name=sheet_name, index=False)

    return str(output_file)


def run_drt_workbook(
    excel_file_path,
    rbf_type="Gaussian",
    data_used="Combined Re-Im Data",
    induct_used=2,
    der_used="2nd order",
    cv_type="GCV",
    reg_param=1e-3,
    shape_control="FWHM Coefficient",
    coeff=0.5,
    run_twice=False,
):
    excel_file_path = Path(excel_file_path)
    xls = pd.ExcelFile(excel_file_path)
    output_file = excel_file_path.with_name(f"{excel_file_path.stem}_DRT.xlsx")

    summary_rows = []

    with pd.ExcelWriter(output_file) as writer:
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

            try:
                freq, z_real, z_imag, meta = extract_eis_from_sheet(df)
            except Exception:
                continue

            potential = try_get_potential(df)

            entry = EIS_object(freq, z_real, z_imag)

            results = simple_run(
                entry,
                rbf_type=rbf_type,
                data_used=data_used,
                induct_used=induct_used,
                der_used=der_used,
                cv_type=cv_type,
                reg_param=reg_param,
                shape_control=shape_control,
                coeff=coeff
            )

            opt_lambda = results.lambda_value

            if run_twice and opt_lambda != reg_param:
                results = simple_run(
                    entry,
                    rbf_type=rbf_type,
                    data_used=data_used,
                    induct_used=induct_used,
                    der_used=der_used,
                    cv_type=cv_type,
                    reg_param=opt_lambda,
                    shape_control=shape_control,
                    coeff=coeff
                )

            results_dict = {
                "Potential_V": potential,
                "R_ohm": results.R,
                "L_H": results.L,
                "tau_s": results.out_tau_vec,
                "gamma": results.gamma,
                "method": results.method,
                "lambda_optimal": results.lambda_value,
                "frequency_Hz": results.freq,
                "z_real_ohm": results.Z_prime,
                "z_imag_ohm": results.Z_double_prime,
                "z_real_fit_ohm": results.mu_Z_re,
                "z_imag_fit_ohm": results.mu_Z_im,
                "Z_re_res": results.res_re,
                "Z_im_res": results.res_im,
            }

            max_len = max(
                len(v) if isinstance(v, (list, np.ndarray)) else 1
                for v in results_dict.values()
            )

            out = pd.DataFrame({
                k: (
                    np.pad(v, (0, max_len - len(v)), constant_values=np.nan)
                    if isinstance(v, (list, np.ndarray))
                    else [v] * max_len
                )
                for k, v in results_dict.items()
            })

            out.to_excel(writer, sheet_name=sheet_name[:31], index=False)

            summary_rows.append({
                "sheet_name": sheet_name,
                "n_points": len(freq),
                "potential_V": potential,
                "lambda_optimal": results.lambda_value,
                "R_ohm": results.R,
                "L_H": results.L,
            })

        if summary_rows:
            pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)

    return str(output_file)