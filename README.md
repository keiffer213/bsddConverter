# Excel2bSDD Converter

A desktop Python application that converts structured Excel files into valid bSDD-compliant JSON documents. Built with Tkinter for GUI and supports packaging into a standalone `.exe`.

## Features

- 🖼️ GUI-based interface using Tkinter
- ⚙️ Converts Excel (based on bSDD template) to JSON
- 🧹 Option to remove `null` fields
- 🧪 Unit tests comparing GUI and CLI outputs
- 📦 Easy packaging with PyInstaller into a .exe

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
├── tests/
│   ├── test_converter_output.py
│   ├── data/
│   │   ├── test_excel_dd.xlsx
│   │   └── bsdd-import-model.json
│   └── expected_result.json
├── README.md
├── requirements.txt
├── pyproject.toml
```

## Pyinstaller Packaging
Run Command Prompt: 

pyinstaller --onefile --name bsddconverter --windowed --paths src --add-data "templates;templates" --add-data "data;data" --hidden-import openpyxl --hidden-import et_xmlfile  --hidden-import jdcal src/bsddconverter/gui.py

pyinstaller --onefile --name bsddconverter --windowed --paths src --add-data "templates;templates" --add-data "data;data" --collect-submodules openpyxl src/bsddconverter/gui.py

<!-- Build Ouput Path excluding tests -->
pyinstaller --onefile --name bsddconverter --windowed --paths src --collect-submodules openpyxl src/bsddconverter/gui.py --distpath build_output/dist --workpath build_output/build --exclude-module tests --exclude-module pytest --specpath build_output/spec

<!-- Build Ouput Path excluding tests with console -->
pyinstaller --onefile --name bsddconverter --windowed --paths src --collect-submodules openpyxl src/bsddconverter/gui.py --distpath build_output/dist --workpath build_output/build --exclude-module tests --exclude-module pytest --specpath build_output/spec --console

## Run the GUI
Navigate to the root folder then use:

```bash 
cd src
python -m bsddconverter.gui
```

## Run pytest
```bash
pytest tests/
```
or a specific file
```bash
pytest tests/test_converter_output.py
```

🧪 Test Flow
Tests validate your GUI converter by comparing its JSON output to the known-good result from the original Excel2bSDD_converter.py CLI tool.

## Requirements (for development)

- Python 3.8+
- pandas
- openpyxl
- tqdm
- numpy
- pytest

Install all dependencies:
```bash
pip install -r requirements.txt


📬 Contact
Maintained by @keiffer213.
For bSDD spec questions, visit buildingsmart.org.