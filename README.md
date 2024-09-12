# PDF Parser
A program to parse PDF document containing several purchase orders of an organization and splits them into individual purchase order PDF documents.
Each purchase order document can be saved and shared accordingly.

This program does the following:
- Parses the given PDF
- Splits each purchase order into individual PDF documents
- Names each purchase order PDF document according to its purchase order number
- Exports the purchase order date, purchase order number, and vendor number to an excel sheet

## Requirements
Python 3

A python IDE

## Instructions
- Run the pdf-parser.py file `python pdf_parser.py`
- Enter the file path of the document e.g. `input/byu.pdf`
- Enter the path of the folder where the new pdf documents will be placed e.g. `output`
- Press Enter
