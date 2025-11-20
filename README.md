📘 DOCX/RTF Descriptive Memorial Generator

A robust and flexible Python-based tool that automates the generation of Descriptive Memorial documents from a Microsoft Word (.docx or .rtf) template and structured tabular data (.csv) containing general information and perimeter point data.

This tool is designed for seamless integration with surveying workflows (e.g., TopoCAD 2000) and can operate both in Command Line Mode and Graphical Mode (GUI).

🚀 Features

Supports both .DOCX and .RTF templates.

Automatic detection of General Data CSV and Point Data CSV on the same file.

Preserves font styles, sizes, and bold formatting from the template.

Offers CLI (command-line) and GUI (interactive) execution modes.

Generates a fully formatted descriptive memorial automatically.

⚙️ Execution Modes
🖥️ 1. Graphical Interface (Default Mode)

If you run the script without arguments, it automatically launches an interactive graphical window that allows you to manually select:

The template file (.docx or .rtf)

The general data CSV file

The point data CSV file

The destination path for the generated document

py gerador_rtf.py
# or
gerador_rtf.exe

💻 2. Command Line Mode

If you prefer automation or need to integrate the generator with external systems (like TopoCAD), you can execute it directly with arguments.

Syntax
my_script.exe <PATH_TO_TEMPLATE> <CSV_BASE_PATH> [options]

Example
my_script.exe "C:\Projects\Template.docx" "C:\Projects\Data"


The script will automatically read:

C:\Projects\Data1.csv → General data

C:\Projects\Data2.csv → Perimeter point data

And generate:

C:\Projects\Data.docx → Final descriptive memorial document.

⚙️ Optional Flags
Argument	Description	Default
--nao_repetir_confrontante	Disables automatic replacement of repeated confronting names (“o mesmo”).	False
--x	Forces CLI mode even if no arguments are detected (useful for testing).	False
Example with both options:
my_script.exe "C:\Users\lipin\Downloads\MODELO (1).docx" "C:\Users\lipin\Downloads\TESTE-MOD" --x --nao_repetir_confrontante


This command:

Uses the specified DOCX template.

Reads data from:

TESTE-MOD1.csv (general data)

TESTE-MOD2.csv (point data)

Generates the memorial document automatically in the same directory.

🧱 Input File Requirements
1. Template File (.docx or .rtf)

Contains substitution placeholders enclosed in < > brackets.
Example:

<IMOVEL> - <MUNICIPIO>
Azimuth: <AZIMUTE> - Distance: <DISTANCIA>m - Point: <PONTO>
Confrontant: <CONFRONTANTE>
<***>
... repeating section ...
<***>

Supported Placeholders
Category	Example	Description
General	<IMOVEL>, <PERIMETRO>, <MUNICIPIO>, <RESPONSAVEL>	Header and document metadata
Point Data	<PONTO>, <AZIMUTE>, <DISTANCIA>, <CONFRONTANTE>, <UTMX>, <UTMY>	Used in repetition block
Repetition Block	<***>	Defines section to repeat for each segment
2. General Data CSV (<BASE_PATH>1.csv)

Format: KEY;VALUE
Example:

<IMOVEL>;SÍTIO ALEGRIA
<AREA_M2>;19742,5
<DATA>;23/07/2025

3. Point Data CSV (<BASE_PATH>2.csv)

Contains the list of perimeter points, one per row.
Example:

<PONTO>;<AZIMUTE>;<DISTANCIA>;<CONFRONTANTE>;<UTMX>;<UTMY>
M01;45°30'00";125.6;JOÃO SILVA;756321.44;9102234.55
M02;130°45'15";210.2;O MESMO;756444.11;9102445.10
...

🧩 Document Generation Details

The section between <***> markers is repeated for each perimeter segment (N-1 times).

The final paragraph is automatically closed with the return to the first point:

... until point M01, where this description began.


If --nao_repetir_confrontante is not passed, the script replaces repeated confronting names with “o mesmo” automatically.

🧠 Technical Notes

The generator ensures font consistency by cloning styles (family, size, and bold attributes) from the reference paragraph in the DOCX template.

RTF templates are supported with minimal formatting guarantees.

The script works fully offline and requires no installation if distributed as a .exe file.

📦 Distribution (Optional)

If you plan to share the tool:

Build it into a standalone .exe with PyInstaller:

pyinstaller --onefile --clean --noconfirm gerador_rtf.py


(Optional) Compress it using UPX
 for a smaller size:

pyinstaller --onefile --clean --noconfirm --upx-dir C:\path\to\upx gerador_rtf.py


Distribute the .exe along with this README and example CSV/template files.

🧾 Example Full Command
gerador_rtf.exe "C:\Users\lipin\Downloads\MODELO (1).docx" "C:\Users\lipin\Downloads\TESTE-MOD" --x


✅ This will generate:

C:\Users\lipin\Downloads\TESTE-MOD.docx


Using:

Template → MODELO (1).docx

General Data → TESTE-MOD1.csv

Point Data → TESTE-MOD2.csv
