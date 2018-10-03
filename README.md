# Name Book Parser

Takes a book of names sorted by state, city, and high school and parses into a usable Excel workbook.
Names are sourced from a book with specific formatting and read to a text file using Tesseract OCR.

Included Python file uses the aforementioned file to create a spreadsheet by parsing recognized tokens in the text,
and populates CEEB codes for high schools when possible by checking against a nested dictionary
