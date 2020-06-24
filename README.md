# batch pdf form filler

This project was originally designed as a gift to my wife to automate her workflow by converting Excel or CSV data sheets into a predesigned fillable PDF form template. I want to further expand the functions and make it a handy tool for every one.

## Requirements

Python 3.6+ environment

~Install [PDF tk free version](https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/)~

```shell
    pip install -r requirements
```

## Usage

You are suggested to install `pyinstaller` to generate executable file to run the program.

```shell
  pyinstaller -w -i .\office.ico --hidden-import xlrd --hidden-import pdfrw --noupx .\main.py
```


## Acknowledgement

Thanks to [pdfrw](https://github.com/pmaupin/pdfrw) for the pdf rendering and thanks to [PySimpleGUI library](https://pysimplegui.readthedocs.io/en/latest/) for the GUI.

## Roadmap

- [x] Basic functions of reading Excel/CSV files
- [x] Batch generating PDFs on both Windows and Linux
- [x] Basic and ugly GUI
- [x] Solve Chinese character path issues
- [x] Use pdfrw instead of PDF tk tool to accelerate the process
- [ ] More options for template and parameters
- [ ] Inspect new PDF form and generate new template based on it
- [ ] Able to import new template and design new pattern to make the project more general
- [ ] Nicer appearance (current layout is too ugly)
- [ ] Able to stop generating process in the middle
- [x] Dataview pane for the imported data to preview
- [ ] Change the column number to be imported and viewed
- [ ] PDF Preview
- [ ] Support more PDF form field types
  - [x] Textbox field
  - [ ] Checkbox
  - [ ] Radio Button
  - [ ] List of choices
  - [ ] Dropdown list
  - [ ] Signature (not planning to support)