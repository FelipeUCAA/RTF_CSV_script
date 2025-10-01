DOCX/RTF Descriptive Memorial Generator
A robust Python script designed to automate the creation of Descriptive Memorial documents from a Microsoft Word template (.docx, .rtf) and structured tabular data (.csv) containing point information and general details.

The tool is designed to be executed via the command line, making it ideal for integration with surveying systems like TopoCAD 2000.

🚀 How to Use (Via Command Line)
The standard execution requires you to pass the file path of the document template and the base path of the data files (CSVs) as arguments.

Prerequisites
No Python installation or external library installation is required for the end-user if you are distributing a bundled executable (.exe).

Execution Syntax (Recommended)
Run the script (e.g., memorial_generator.exe or memorial_generator.py) followed by the required arguments and any optional flags:

# Execution Syntax (Standard/Default)
my_script.exe <PATH_TO_TEMPLATE> <CSV_BASE_PATH>

# Execution with Optional Flag (Disabling repetition of 'o mesmo')
my_script.exe <PATH_TO_TEMPLATE> <CSV_BASE_PATH> --nao_repetir_confrontante

Practical Example:

If your files are: C:\Projects\Template.docx, C:\Projects\Data_1.csv, and C:\Projects\Data_2.csv.

my_script.exe C:\Projects\Template.docx C:\Projects\Data

The script will automatically look for C:\Projects\Data1.csv and C:\Projects\Data2.csv, and save the resulting document as C:\Projects\Data.docx.

⚙️ Technical Details and Command Options
Control of Repeated Confrontants (Optional)
The handling of repetitive adjacent confronting properties is conditional, controlled by the main argument:

Behavior

Description

Argument

Default

If the confrontant name is the same as the previous segment, it is replaced by the phrase "o mesmo".

None

Optional

The full confrontant name from the CSV is repeated for every segment, regardless of similarity.

--nao_repetir_confrontante

Formatting Assurance (DOCX Only)
The code ensures formatting integrity in the generated paragraphs:

Font and Size: The script copies the exact font family and font size from the reference paragraph in the template and applies them to all newly generated text runs.

Boldness: Point data (<PONTO>) is explicitly preserved in bold across all generated blocks, aligning with the template's requirements.

Other Technical Details
Block Processing: The section between <***> and <***> is repeated for each perimeter segment (N-1 repetitions).

Final Output: The final document ends after the last successfully generated repetition block.

📂 Required Input File Structure
The script expects your input data to be organized into three main files:

1. Document Template (.docx or .rtf)
This file must contain the placeholders (substitution markers) in <KEY> format:

Data Type

Placeholder Example

Usage

General

<IMOVEL>, <PERIMETRO>, <MUNICIPIO>, <RESPONSAVEL>

Header and Footer information.

Point Data

<PONTO>, <AZIMUTE>, <DISTANCIA>, <CONFRONTANTE>, <UTMX>, <UTMY>

Repetition Block, Opening, and Closing description.

Repetition Block

<***>

Marks the start and end of the text block that will be repeated for each segment of the perimeter.

2. General Data CSV (<CSV_BASE_PATH>1.csv)
Format: <KEY>;<VALUE>

Column A (Key)

Column B (Value)

<IMOVEL>

SITIO ALEGRIA

<AREAM2>

19,7425

<DATA>

23/07/2025

3. Point Data CSV (<CSV_BASE_PATH>2.csv)
This file must contain the header row (<PONTO>;<DISTANCIA>;<AZIMUTE>;...) followed by the list of vertices. The script automatically manages the order of the confrontations.
