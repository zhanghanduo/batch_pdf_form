# batch pdf form filler

This project was originally designed as a gift to my wife to automate her workflow by converting Excel or CSV data sheets into a predesigned fillable PDF form template. I want to further extand the functions and make it a handy tool for every one.

## Requirements

Python 3.6+ environment

Install [PDF tk free version](https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/)

```
    pip install -r requirements
```

## Usage



## Acknowledgement

Thanks to [pdfforms](https://github.com/altaurog/pdfforms) for the original idea and thanks to [PySimpleGUI library](https://pysimplegui.readthedocs.io/en/latest/) for the GUI.

## Roadmap

- [x] Basic functions of reading Excel/CSV files
- [x] Batch generating PDFs on Windows and Linux
- [x] Basic and ugly GUI
- [ ] Solve Chinese character path issues
- [ ] More elegant way of working with PDF tk tool or get independent from it
- [ ] Inspect new PDF form and generate new template based on it
- [ ] Able to import new template and design new pattern to make the project more general
- [ ] Nicer appearance (current layout is too ugly)
- [ ] Able to stop generating process in the middle
- [ ] Support more PDF form field types
  - [x] Textbox field
  - [ ] Checkbox
  - [ ] Radio Button
  - [ ] List of choices
  - [ ] Dropdown list
  - [ ] Signature (not planning to support)