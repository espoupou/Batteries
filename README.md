
# EIS main refactor

This folder contains a structural refactor of the GUI entrypoint.
The goal was to keep the same logic and behavior while separating:

- dialogs
- workbook plotter window
- DRT plotter window
- preprocessing window
- main application shell

Original processing cores are kept as-is in `eis_core.py` and `drt_core.py`.
Run with:

```bash
python main.py
```
