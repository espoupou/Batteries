from pathlib import Path

input_path = Path("datas/AB044 NR37 14-02-2025 VARTA pulse discharge 120mA dts= 10s for 30min rest (loop 720)_01_PEIS_C02_.mpt.txt")
output_path = Path("datas/AB044 NR37 14-02-2025 VARTA pulse discharge 120mA dts= 10s for 30min rest (loop 720)_01_PEIS_C02_.mpt_OUTPUT_direct_delete.txt")

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

        parts = parts[:3]

        if parts[0] == 2.6930346E+000 or str(parts[0]) == "2.6930346E+000":
            break
        try:
            parts[2] = str(-float(parts[2]))
        except ValueError as e:
            raise ValueError(f"Line {line_num} has non-numeric 3rd column: {parts[2]!r}") from e

        fout.write("\t".join(parts) + "\n")

print(f"Done. Wrote: {output_path}")
