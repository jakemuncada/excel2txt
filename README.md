# excel2txt
Python module to convert excel files to txt files.

## Main Purpose:
As the saying goes, necessity is the mother of invention. 

I had a project where the client provides (via the Project Leader) an excel document. Sometimes changes would be made to the document, but since Git does not show specific changes to an excel file, I had no way to know which sheets/modules had changed.

I wanted to be able to use a Diff tool to compare sheets to quickly determine (or just to verify) which rows had changed.

Thus, I had the idea to create a python script to convert it into a txt file. It was also a good opportunity to practice using `openpyxl` since this was my first time using it.

---

## Installation:
Install the module requirements using the code below:
```
pip install -r requirements.txt
```

## Usage:
```
python -m excel2txt INPUT_FILEPATH [-s SHEET_NAME] [-o OUTPUT_FILEPATH]
```
- `INPUT_FILEPATH`: The path of the excel file to be converted. Requried.
- `-s SHEET_NAME`: Used to specify a single excel sheet to be converted. Optional. If not specified, the entire excel along with all its sheets will be converted.
- `-o OUTPUT_PATH`: Used to specify the output txt file path. Optional. If not specified, the output directory will be the same directory as the script.
