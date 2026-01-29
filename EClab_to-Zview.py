from pathlib import Path

input_path = Path("input.txt")
output_path = Path("output.txt")

with input_path.open("r", encoding="utf-8") as fin, output_path.open("w", encoding="utf-8") as fout:
    _ = fin.readline()

    for line_num, line in enumerate(fin, start=2):
        line = line.strip()
        if not line:
            continue

        line = line.replace(",", ".")

        parts = line.split()
        if len(parts) < 3:
            raise ValueError(f"Line {line_num} does not have 3 columns: {line!r}")

        try:
            parts[2] = str(-float(parts[2]))
        except ValueError as e:
            raise ValueError(f"Line {line_num} has non-numeric 3rd column: {parts[2]!r}") from e

        fout.write("\t".join(parts) + "\n")

print(f"Done. Wrote: {output_path}")
