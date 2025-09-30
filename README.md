DOCX/RTF Descriptive Memorial Generator
A robust Python script designed to automate the creation of Descriptive Memorial documents from a Microsoft Word template (.docx, .rtf) and structured tabular data (.csv) containing point information and general details.

The tool is designed to be executed via the command line, making it ideal for integration with surveying systems like TopoCAD 2000.

🚀 How to Use (Via Command Line)
The standard execution of the script requires you to pass the file path of the document template and the base path of the data files (CSVs) as arguments.

Prerequisites
No Python installation or external library installation is required for the end-user if you are distributing a bundled executable (.exe).

Execution Syntax (Recommended)
Run the script (e.g., memorial_generator.exe or memorial_generator.py) followed by the two required arguments:

Bash

# Executing the script via console (TopoCAD, Prompt, etc.)
my_script.exe <PATH_TO_TEMPLATE> <CSV_BASE_PATH>
Practical Example:

If your files are:

Template: C:\Projects\Template.docx

General Data: C:\Projects\Data_1.csv

Point Data: C:\Projects\Data_2.csv

You should use the path to the template and the common path prefix of the CSV files (C:\Projects\Data):

Bash

my_script.exe C:\Projects\Template.docx C:\Projects\Data
The script will automatically look for C:\Projects\Data1.csv and C:\Projects\Data2.csv, and save the resulting document as C:\Projects\Data.docx.

📂 Required Input File Structure
The script expects your input data to be organized into three main files:

1. Document Template (.docx or .rtf)
This file must contain the placeholders (substitution markers) in <KEY> format:

Data Type	Placeholder Example	Usage
General	<IMOVEL>, <PERIMETRO>, <MUNICIPIO>, <RESPONSAVEL>	Header and Footer information.
Point Data	<PONTO>, <AZIMUTE>, <DISTANCIA>, <CONFRONTANTE>, <UTMX>, <UTMY>	Repetition Block, Opening, and Closing description.
Repetition Block	<***>	Marks the start and end of the text block that will be repeated for each segment of the perimeter.

Exportar para Sheets
2. General Data CSV (<CSV_BASE_PATH>1.csv)
Format: <KEY>;<VALUE>

Column A (Key)	Column B (Value)
<IMOVEL>	SITIO ALEGRIA
<AREAM2>	19,7425
<DATA>	23/07/2025

Exportar para Sheets
3. Point Data CSV (<CSV_BASE_PATH>2.csv)
This file must contain the header row (<PONTO>;<DISTANCIA>;<AZIMUTE>;...) followed by the list of vertices. The script automatically manages the order of the confrontations.

🛠️ Technical Details
Block Processing: The section between <***> and <***> is repeated for each perimeter segment (N-1 repetitions, where N is the total number of points).

Formatting (Bold): The script implements manual reconstruction of runs to ensure that point data (<PONTO>) is preserved in bold across all generated paragraphs (Opening, Repetition Blocks, and Closing), provided the original placeholder in the model was formatted as such.

Repeated Confrontants: The logic automatically detects repeated confronting properties, replacing the name with the expression "o mesmo" (the same).

Final Output: The final document ends after the last successfully generated repetition block.
