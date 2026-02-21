#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Remove a ZView "Wo" (Finite Length Warburg - Open Circuit Terminus) contribution
from EIS data, using the fitted parameters Wo-R, Wo-T, Wo-P.

ZView manual (Wo element):   Z_Wo = Wo-R * coth( (j*Wo-T*ω)^(Wo-P) ) / ( (j*Wo-T*ω)^(Wo-P) )
with ω = 2πf

Typical use:
    python remove_warburg_wo.py --input data.mpt.txt --output data_noWo.csv \
        --wor 22.65 --wot 26.79 --wop 0.23327 --im-is-neg

If your preprocessor already outputs 3 columns (freq, Re, Im or -Im), it will work too.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Tuple, Optional

import numpy as np

try:
    import pandas as pd
except Exception as e:  # pragma: no cover
    pd = None


def coth(z: np.ndarray) -> np.ndarray:
    # coth(z) = 1 / tanh(z)
    return 1.0 / np.tanh(z)


def z_wo(freq_hz: np.ndarray, wor: float, wot: float, wop: float) -> np.ndarray:
    """
    Compute Z of ZView Wo element (Finite Length Warburg - Open Circuit Terminus)
    for each frequency in Hz.
    """
    omega = 2.0 * np.pi * freq_hz.astype(float)
    arg = (1j * wot * omega) ** wop
    return wor * coth(arg) / arg


def _guess_delim(sample_line: str) -> Optional[str]:
    # Prefer explicit delimiters; avoid comma because it may be decimal separator
    candidates = ['\t', ';']
    counts = {c: sample_line.count(c) for c in candidates}
    best = max(counts, key=counts.get)
    if counts[best] > 0:
        return best
    # fallback: whitespace
    return None


def _safe_float(s: str) -> float:
    # convert European decimal comma to dot, keep exponent
    s = s.strip()
    if not s:
        raise ValueError("empty numeric field")
    s = s.replace(',', '.')
    return float(s)


def load_eis_3cols(path: Path, freq_col: str | None, re_col: str | None, im_col: str | None) -> Tuple[np.ndarray, np.ndarray, np.ndarray, bool]:
    """
    Load EIS data (freq, Re, Im) from either:
    - a BioLogic .mpt(.txt) table (many columns, header row contains freq/Hz, Re(Z), -Im(Z))
    - a simple 3-column file (freq, Re, Im or -Im)

    Returns: freq_hz, re_ohm, im_raw, im_is_neg
      im_raw is the 3rd column numeric values from the file
      im_is_neg indicates whether the file column is "-Im" (common in EIS exports)
    """
    text = path.read_text(encoding="utf-8", errors="ignore").splitlines()
    # Find a header line containing 'freq' if present
    header_idx = None
    header_line = None
    for i, line in enumerate(text[:400]):
        if re.search(r'\bfreq', line, re.IGNORECASE):
            # likely the column header row
            if ('Re' in line or 'Im' in line or 'Z' in line):
                header_idx = i
                header_line = line
                break

    if pd is not None and header_idx is not None:
        delim = _guess_delim(header_line)
        try:
            df = pd.read_csv(
                path,
                sep=delim if delim else r"\s+",
                engine="python",
                header=header_idx,
                decimal=",",
                encoding="utf-8",
            )
        except UnicodeDecodeError:
            df = pd.read_csv(
                path,
                sep=delim if delim else r"\s+",
                engine="python",
                header=header_idx,
                decimal=",",
                encoding="latin1",
            )

        # Normalize column names
        cols = {c: str(c).strip() for c in df.columns}
        df.rename(columns=cols, inplace=True)
        cn = list(df.columns)

        def find_col(patterns):
            for pat in patterns:
                for c in cn:
                    if re.search(pat, c, re.IGNORECASE):
                        return c
            return None

        fcol = freq_col or find_col([r"^freq", r"frequency"])
        rcol = re_col or find_col([r"Re\(Z\)", r"\bZ'\b", r"Zre", r"Re\s*Z"])
        icol = im_col or find_col([r"-Im\(Z\)", r"Im\(Z\)", r"\bZ''\b", r"Zim", r"Im\s*Z"])

        if fcol is None or rcol is None or icol is None:
            raise ValueError(
                f"Could not locate columns. Found: {cn[:15]}... "
                "Use --freq-col/--re-col/--im-col to specify explicitly."
            )

        # infer sign convention: if header contains "-Im" then values are -Im (positive)
        im_is_neg = bool(re.search(r"-\s*Im", str(icol), re.IGNORECASE))

        # Coerce to numeric safely (handle European decimal commas AND scientific notation)
        def _to_num(s):
            # Keep NaNs as-is; operate on strings
            return pd.to_numeric(s.astype(str).str.replace(',', '.', regex=False), errors="coerce")

        freq = _to_num(df[fcol])
        re_z = _to_num(df[rcol])
        im_z = _to_num(df[icol])

        good = freq.notna() & re_z.notna() & im_z.notna()
        freq = freq[good].to_numpy(dtype=float)
        re_z = re_z[good].to_numpy(dtype=float)
        im_z = im_z[good].to_numpy(dtype=float)
        return freq, re_z, im_z, im_is_neg

    # Fallback: simple parsing (3 columns)
    freq_list, re_list, im_list = [], [], []
    for line in text:
        line = line.strip()
        if not line:
            continue
        if line.startswith("#"):
            continue
        if re.search(r"[a-zA-Z]", line):
            # skip non-numeric lines
            continue
        # split on whitespace, tab, or semicolon
        parts = re.split(r"[\t; ]+", line)
        parts = [p for p in parts if p]
        if len(parts) < 3:
            continue
        try:
            f = _safe_float(parts[0])
            r = _safe_float(parts[1])
            im = _safe_float(parts[2])
        except Exception:
            continue
        freq_list.append(f)
        re_list.append(r)
        im_list.append(im)

    if len(freq_list) < 5:
        raise ValueError("Could not parse enough numeric rows. Provide a 3-column file or specify columns for a table export.")

    # With 3-column file we cannot know if it's Im or -Im; default: assume it's -Im (common)
    return np.array(freq_list, float), np.array(re_list, float), np.array(im_list, float), True


def main():
    ap = argparse.ArgumentParser(description="Subtract ZView Wo diffusion element from EIS data.")
    ap.add_argument("--input", "-i", required=True, help="Input file (BioLogic .mpt.txt or 3-column file).")
    ap.add_argument("--output", "-o", required=True, help="Output CSV/TSV path.")
    ap.add_argument("--wor", type=float, required=True, help="Wo-R (Ohm).")
    ap.add_argument("--wot", type=float, required=True, help="Wo-T (s).")
    ap.add_argument("--wop", type=float, required=True, help="Wo-P (dimensionless exponent).")
    ap.add_argument("--freq-col", default=None, help="(Optional) explicit frequency column name.")
    ap.add_argument("--re-col", default=None, help="(Optional) explicit Re(Z) column name.")
    ap.add_argument("--im-col", default=None, help="(Optional) explicit Im(Z) or -Im(Z) column name.")
    ap.add_argument("--im-is-neg", action="store_true",
                    help="Force the 3rd column to be treated as -Im(Z) (positive).")
    ap.add_argument("--im-is-im", action="store_true",
                    help="Force the 3rd column to be treated as Im(Z) (can be negative).")
    ap.add_argument("--tsv", action="store_true", help="Write TSV instead of CSV.")
    ap.add_argument("--plot", action="store_true", help="Save Nyquist plots (before/after) next to output.")
    ap.add_argument("--fmin", type=float, default=None,
                    help="Keep only frequencies >= fmin (Hz) after correction.")
    ap.add_argument("--trim-pos-im", action="store_true",
                    help="Drop points where corrected Im(Z) > 0 (inductive-looking hook) after correction.")
    args = ap.parse_args()

    in_path = Path(args.input)
    out_path = Path(args.output)

    freq, re_z, im_raw, inferred_is_neg = load_eis_3cols(in_path, args.freq_col, args.re_col, args.im_col)

    if args.im_is_neg and args.im_is_im:
        raise SystemExit("Choose only one of --im-is-neg or --im-is-im.")

    im_is_neg = inferred_is_neg
    if args.im_is_neg:
        im_is_neg = True
    if args.im_is_im:
        im_is_neg = False

    # Build measured complex Z
    # If file gives -Im(Z) (positive), then Im(Z) = -(-Im) = -im_raw
    im_meas = (-im_raw) if im_is_neg else im_raw
    z_meas = re_z + 1j * im_meas

    # Compute Wo and subtract
    zwo = z_wo(freq, args.wor, args.wot, args.wop)
    z_corr = z_meas - zwo

    # Convert back to convenient columns:
    re_corr = z_corr.real
    im_corr = z_corr.imag
    im_corr_out = (-im_corr) if im_is_neg else im_corr  # keep same convention as input for convenience

    # Optional: also export Wo contribution for checking
    zwo_re = zwo.real
    zwo_im = zwo.imag
    zwo_im_out = (-zwo_im) if im_is_neg else zwo_im

    # --- Optional trimming / cleaning (low-frequency artifacts) ---
    mask = np.ones_like(freq, dtype=bool)

    if args.fmin is not None:
        mask &= (freq >= args.fmin)

    if args.trim_pos_im:
        # keep only capacitive-looking points after correction (Im <= 0)
        mask &= (z_corr.imag <= 0)

    # Apply mask consistently
    freq = freq[mask]
    z_meas = z_meas[mask]
    z_corr = z_corr[mask]

    re_corr = z_corr.real
    im_corr = z_corr.imag

    # if pd is not None:
    #     df_out = pd.DataFrame({
    #         "freq_Hz": freq,
    #         "Re_meas_Ohm": re_z,
    #         ("-Im_meas_Ohm" if im_is_neg else "Im_meas_Ohm"): im_raw,
    #         "Wo_Re_Ohm": zwo_re,
    #         ("Wo_-Im_Ohm" if im_is_neg else "Wo_Im_Ohm"): zwo_im_out,
    #         "Re_corr_Ohm": re_corr,
    #         ("-Im_corr_Ohm" if im_is_neg else "Im_corr_Ohm"): im_corr_out,
    #     })
    #     sep = "\t" if args.tsv or out_path.suffix.lower() in {".tsv", ".txt"} else ","
    #     df_out.to_csv(out_path, index=False, sep=sep)
    # else:
    #     sep = "\t" if args.tsv or out_path.suffix.lower() in {".tsv", ".txt"} else ","
    #     header = [
    #         "freq_Hz", "Re_meas_Ohm", ("-Im_meas_Ohm" if im_is_neg else "Im_meas_Ohm"),
    #         "Wo_Re_Ohm", ("Wo_-Im_Ohm" if im_is_neg else "Wo_Im_Ohm"),
    #         "Re_corr_Ohm", ("-Im_corr_Ohm" if im_is_neg else "Im_corr_Ohm")
    #     ]
    #     with out_path.open("w", encoding="utf-8") as f:
    #         f.write(sep.join(header) + "\n")
    #         for i in range(len(freq)):
    #             row = [freq[i], re_z[i], im_raw[i], zwo_re[i], zwo_im_out[i], re_corr[i], im_corr_out[i]]
    #             f.write(sep.join(f"{v:.12g}" for v in row) + "\n")

    # --- Export only corrected Z' and Z'' ---
    # Note: Z'' here is the true imaginary part Im(Z) (usually negative in Nyquist),
    # not "-Im(Z)". That's what you want if you write Z' / Z''.
    if pd is not None:
        df_out = pd.DataFrame({
            "freq_Hz": freq,
            "Zprime_Ohm": re_corr,  # Z'  = Re(Z)
            "Zdoubleprime_Ohm": im_corr,  # Z'' = Im(Z)
        })
        sep = "\t" if args.tsv or out_path.suffix.lower() in {".tsv", ".txt"} else ","
        df_out.to_csv(out_path, index=False, sep=sep)
    else:
        sep = "\t" if args.tsv or out_path.suffix.lower() in {".tsv", ".txt"} else ","
        header = ["freq_Hz", "Zprime_Ohm", "Zdoubleprime_Ohm"]
        with out_path.open("w", encoding="utf-8") as f:
            #f.write(sep.join(header) + "\n")
            for i in range(len(freq)):
                row = [freq[i], re_corr[i], im_corr[i]]
                f.write(sep.join(f"{v:.12g}" for v in row) + "\n")

    if args.plot:
        try:
            import matplotlib.pyplot as plt

            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4))

            # --- Measured ---
            ax1.plot(z_meas.real, -z_meas.imag, marker='.', linestyle='None')
            ax1.set_xlabel("Z' (Ohm)")
            ax1.set_ylabel("-Z'' (Ohm)")
            ax1.set_title("Nyquist (measured)")
            ax1.grid(True, which="both", linestyle=":", linewidth=0.7)
            ax1.set_aspect("equal", adjustable="box")
            ax1.axis("equal")

            # --- Corrected (Wo removed) ---
            ax2.plot(z_corr.real, -z_corr.imag, marker='.', linestyle='None')
            ax2.set_xlabel("Z' (Ohm)")
            ax2.set_ylabel("-Z'' (Ohm)")
            ax2.set_title("Nyquist (Wo removed)")
            ax2.grid(True, which="both", linestyle=":", linewidth=0.7)
            ax2.set_aspect("equal", adjustable="box")
            ax2.axis("equal")

            fig.tight_layout()
            plt.show()

        except Exception as e:
            print(f"[WARN] Plot failed: {e}")

    print(f"Done. Wrote: {out_path}")


if __name__ == "__main__":
    main()


#C:\Users\synux\AppData\Local\Programs\Python\Python313\python.exe remove_warburg_wo.py \
"""  --input "ton_fichier.txt" \
  --output "ton_fichier_noWo.csv" \
  --wor 22.65 --wot 26.79 --wop 0.23327 \
  --im-is-neg --plot"""