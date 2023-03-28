# KiCad BOM Generator

Simple .xlsx format BOM generator for KiCad 7.0
 
Input parameters are the same as for default .csv script.

## Instalation
1. Open KiCad BOM generation window.
2. Click + icon to add new generator
3. Select bom_excel_PSD.py script
4. Add .xlsx to output parameter in command line. It should look like this: `python "Path\to\script\bom_excel_PSD.py" "%I" "%O.xlsx"`
5. Open KiCad command promt.
6. Install python packages with `pip install -r requirements.txt`
7. Generator is ready to go
