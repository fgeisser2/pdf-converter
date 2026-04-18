#!/bin/python

from pathlib import Path

main_py = Path("./src/main.py").read_text()
macro_py = Path("./src/macro.py").read_text()

dist_py = main_py.replace("{{LIBREOFFICE_MACRO}}", macro_py)

Path("pdf-converter.py").write_text(dist_py)
