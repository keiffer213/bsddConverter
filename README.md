# Excel2bSDD Converter

A desktop Python application that converts a specially structured Excel file into a valid bSDD-compliant JSON document.

## Features

- Built-in GUI using Tkinter
- Optional null-field removal
- No dependencies required for users (compiled to .exe with PyInstaller)
- Modular, readable Python codebase

## Pyinstaller Packaging
Run Command Prompt: 

pyinstaller --onefile --name bsddconverter --windowed --paths src --add-data "templates;templates" --add-data "data;data" --hidden-import openpyxl --hidden-import et_xmlfile  --hidden-import jdcal src/bsddconverter/gui.py

pyinstaller --onefile --name bsddconverter --windowed --paths src --add-data "templates;templates" --add-data "data;data" --collect-submodules openpyxl src/bsddconverter/gui.py


<!-- Build Ouput Path -->
pyinstaller --onefile --name bsddconverter --windowed --paths src --collect-submodules openpyxl src/bsddconverter/gui.py --distpath build_output/dist --workpath build_output/build --specpath build_output/spec

## Run as Python Module
Navigate to the root folder then to \src and use:

```bash 
python -m bsddconverter.gui

```markdown
📁 Recommended Project Structure:
.
├── src/
│   └── bsddconverter/
│       ├── gui.py
│       ├── mapper.py
│       ├── __init__.py
├── templates/
├── data/
├── requirements.txt
├── pyproject.toml

## Requirements (for development)

- Python 3.8+
- pandas
- openpyxl
- tqdm
- numpy

Install all dependencies:
```bash
pip install -r requirements.txt