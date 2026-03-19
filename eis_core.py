from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd


# =========================
# Lecture EIS
# =========================

def _clean_header(s: str) -> str:
    s = str(s).strip().lower()
    s = s.replace("−", "-").replace("–", "-")
    s = s.replace("[", "(").replace("]", ")")
    s = s.replace("{", "(").replace("}", ")")
    s = re.sub(r"\s+", "", s)
    return s


def _classify_header(col: str):
    """
    Retourne (role, sign_im)
    role in {"freq", "R", "Im", None}
    sign_im utile seulement pour Im:
        +1 si la colonne contient Im(Z)
        -1 si la colonne contient -Im(Z)
    """
    raw = str(col).strip().lower().replace("−", "-").replace("–", "-")
    s = _clean_header(col)

    # fréquence
    if "freq" in s or s in {"f", "frequency", "frequency(hz)"}:
        return "freq", None

    # partie réelle de Z
    if (
        "re(z" in raw
        or "real(z" in raw
        or "rez" in s
        or "zreal" in s
        or s == "z'"
        or s.startswith("z'")
    ):
        return "R", None

    # partie imaginaire -Im(Z)
    if (
        "-im(z" in raw
        or raw.startswith("-im")
        or "-z''" in raw
        or "-zimag" in s
    ):
        return "Im", -1

    # partie imaginaire Im(Z)
    if (
        "im(z" in raw
        or "imag(z" in raw
        or "imz" in s
        or "zimag" in s
        or "z''" in raw
    ):
        return "Im", +1

    return None, None


def _detect_header_and_sep(path: Path, max_lines: int = 120):
    """
    Cherche une ligne d'en-tête contenant freq + R + Im.
    Retourne (header_line_idx, sep)
    sep ∈ {'\\t', ';', r'\\s+'}
    """
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()

    for i, line in enumerate(lines[:max_lines]):
        stripped = line.strip()
        if not stripped:
            continue

        if "\t" in line:
            parts = [p.strip() for p in line.rstrip("\n\r").split("\t")]
            sep = "\t"
        elif ";" in line:
            parts = [p.strip() for p in line.rstrip("\n\r").split(";")]
            sep = ";"
        else:
            parts = re.split(r"\s+", stripped)
            sep = r"\s+"

        found = set()
        for p in parts:
            role, _ = _classify_header(p)
            if role:
                found.add(role)

        if {"freq", "R", "Im"}.issubset(found):
            return i, sep

    return None, None


def read_freq_R_Im(path: str | Path) -> pd.DataFrame:
    """
    Retourne un DataFrame normalisé avec colonnes:
        freq, R, Im

    Convention interne:
        Im = vraie partie imaginaire de Z
    Donc si le fichier source contient '-Im(Z)', on reconstruit:
        Im = -(-Im)
    """
    path = Path(path)

    header_idx, sep = _detect_header_and_sep(path)

    if header_idx is not None:
        df = pd.read_csv(
            path,
            sep=sep,
            skiprows=header_idx,
            header=0,
            dtype=str,
            engine="python",
            encoding="utf-8",
        )

        selected = {}
        im_sign = +1

        for col in df.columns:
            role, sign = _classify_header(col)
            if role and role not in selected:
                selected[role] = col
                if role == "Im":
                    im_sign = sign if sign is not None else +1

        missing = [k for k in ["freq", "R", "Im"] if k not in selected]
        if missing:
            raise ValueError(f"Colonnes introuvables: {missing}")

        out = df[[selected["freq"], selected["R"], selected["Im"]]].copy()
        out.columns = ["freq", "R", "Im"]

        for c in ["freq", "R", "Im"]:
            out[c] = out[c].astype(str).str.replace(",", ".", regex=False)
            out[c] = pd.to_numeric(out[c], errors="coerce")

        # si la colonne source était -Im, on revient à Im physique
        if im_sign == -1:
            out["Im"] = -out["Im"]

        out = out.dropna(subset=["freq", "R", "Im"]).reset_index(drop=True)
        return out

    # fallback: fichier simple 3 colonnes
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

    df = df.iloc[:, :3].copy()
    df.columns = ["freq", "R", "Im"]

    for c in ["freq", "R", "Im"]:
        df[c] = df[c].astype(str).str.replace(",", ".", regex=False)
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df.dropna(subset=["freq", "R", "Im"]).reset_index(drop=True)
    return df


# =========================
# Warburg Wo
# =========================

def coth(z: np.ndarray) -> np.ndarray:
    return 1.0 / np.tanh(z)


def z_wo(freq_hz: np.ndarray, wor: float, wot: float, wop: float) -> np.ndarray:
    omega = 2.0 * np.pi * freq_hz.astype(float)
    arg = (1j * wot * omega) ** wop
    return wor * coth(arg) / arg


def remove_warburg_wo(
    df: pd.DataFrame,
    wor: float,
    wot: float,
    wop: float,
    fmin: Optional[float] = None,
    trim_pos_im: bool = False,
) -> pd.DataFrame:
    """
    Applique la suppression du terme Wo sur un DataFrame normalisé:
        colonnes attendues: freq, R, Im

    Retourne un DataFrame avec:
        freq, R, Im, R_meas, Im_meas, Wo_R, Wo_Im
    """
    required = {"freq", "R", "Im"}
    if not required.issubset(df.columns):
        raise ValueError("Le DataFrame doit contenir: freq, R, Im")

    freq = df["freq"].to_numpy(dtype=float)
    re_meas = df["R"].to_numpy(dtype=float)
    im_meas = df["Im"].to_numpy(dtype=float)

    z_meas = re_meas + 1j * im_meas
    zwo = z_wo(freq, wor, wot, wop)
    z_corr = z_meas - zwo

    out = pd.DataFrame({
        "freq": freq,
        "R": z_corr.real,
        "Im": z_corr.imag,
        "R_meas": re_meas,
        "Im_meas": im_meas,
        "Wo_R": zwo.real,
        "Wo_Im": zwo.imag,
    })

    mask = np.ones(len(out), dtype=bool)

    if fmin is not None:
        mask &= (out["freq"].to_numpy() >= float(fmin))

    if trim_pos_im:
        mask &= (out["Im"].to_numpy() <= 0)

    out = out.loc[mask].reset_index(drop=True)
    return out


def export_freq_R_Im(df: pd.DataFrame, out_path: str | Path, sep: Optional[str] = None):
    out_path = Path(out_path)

    if sep is None:
        sep = "\t" if out_path.suffix.lower() in {".tsv", ".txt"} else ","

    out = df[["freq", "R", "Im"]].copy()
    out.columns = ["freq_Hz", "Zprime_Ohm", "Zdoubleprime_Ohm"]
    out.to_csv(out_path, index=False, sep=sep)