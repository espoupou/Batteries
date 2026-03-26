from pathlib import Path
from pathlib import Path
import math


def process_file(
    input_path: Path,
    output_dir: Path | None = None,
    *,
    skip_first_line: bool = True,
    replace_comma: bool = True,
    keep_n_cols: int = 3,
    negate_col3: bool = False,
    stop_first_col: float | None = None,
    encoding: str = "utf-8",
) -> Path:
    """
    Lit un fichier, garde les N premières colonnes, et écrit le résultat
    dans un fichier *_OUTPUT.txt.
    """

    if output_dir is None:
        output_dir = input_path.parent

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{input_path.stem}_OUTPUT{input_path.suffix}"

    with input_path.open("r", encoding=encoding) as fin, output_path.open("w", encoding=encoding) as fout:
        if skip_first_line:
            next(fin, None)

        start_line = 2 if skip_first_line else 1

        for line_num, line in enumerate(fin, start=start_line):
            line = line.strip()
            if not line:
                continue

            if replace_comma:
                line = line.replace(",", ".")

            parts = line.split()

            if len(parts) < keep_n_cols:
                raise ValueError(
                    f"{input_path.name} | line {line_num} does not have {keep_n_cols} columns: {line!r}"
                )

            parts = parts[:keep_n_cols]

            # arrêt si la 1re colonne atteint une certaine valeur
            if stop_first_col is not None:
                try:
                    first_val = float(parts[0])
                except ValueError as e:
                    raise ValueError(
                        f"{input_path.name} | line {line_num} has non-numeric 1st column: {parts[0]!r}"
                    ) from e

                if math.isclose(first_val, stop_first_col, rel_tol=0, abs_tol=1e-12):
                    break

            # inversion du signe de la 3e colonne
            if negate_col3:
                try:
                    parts[2] = str(-float(parts[2]))
                except ValueError as e:
                    raise ValueError(
                        f"{input_path.name} | line {line_num} has non-numeric 3rd column: {parts[2]!r}"
                    ) from e

            fout.write("\t".join(parts) + "\n")

    return output_path


def process_folder(
    folder: Path,
    output_dir: Path | None = None,
    *,
    recursive: bool = False,
    patterns: tuple[str, ...] = ("*.txt",),
    skip_first_line: bool = True,
    replace_comma: bool = True,
    keep_n_cols: int = 3,
    negate_col3: bool = False,
    stop_first_col: float | None = None,
    name_contains: str | None = None,
) -> None:
    """
    Traite tous les fichiers correspondants dans un dossier.
    """

    if output_dir is None:
        output_dir = folder / "converted"

    files = []
    for pattern in patterns:
        iterator = folder.rglob(pattern) if recursive else folder.glob(pattern)
        files.extend(p for p in iterator if p.is_file())

    # enlève les doublons + tri
    files = sorted(set(files))

    # ignore les fichiers déjà générés
    files = [p for p in files if "_OUTPUT" not in p.stem]

    # filtre optionnel sur le nom
    if name_contains is not None:
        files = [p for p in files if name_contains in p.name]

    if not files:
        print("Aucun fichier trouvé.")
        return

    print(f"{len(files)} fichier(s) trouvé(s).")

    ok = 0
    errors = 0

    for file_path in files:
        try:
            out = process_file(
                file_path,
                output_dir=output_dir,
                skip_first_line=skip_first_line,
                replace_comma=replace_comma,
                keep_n_cols=keep_n_cols,
                negate_col3=negate_col3,
                stop_first_col=stop_first_col,
            )
            print(f"[OK] {file_path.name} -> {out.name}")
            ok += 1
        except Exception as e:
            print(f"[ERROR] {file_path.name} -> {e}")
            errors += 1

    print(f"\nTerminé : {ok} succès, {errors} erreur(s).")

process_folder(
    Path("datas/draft"),
    negate_col3=True
)

print(f"Done. Wrote")
