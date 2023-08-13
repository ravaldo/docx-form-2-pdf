### What is this?

When it comes to providing forms for customers/clients/users to fill out and return, inarguably the best approach is to issue PDFs with interactive form fields. However, the process of creating these forms can be quite tedious, even with paid solutions that often require manual placement and alignment of fields.

This script allows you to create a form within Microsoft Word, using all the familiar tools and formatting options and then generate the PDF automatically with form fields correctly placed. Any red boxes will be replaced with a text field, and any green boxes will be replaced with a checkbox. Note that text and checkbox fields are the only two field types handled; if you need a more complex form with more controls (or you need labelling for mail merge), then you will still need to use something like Acrobat Pro.

An example application form is included for reference.

### How to use
This is a python script which uses some 3rd party packages, a pdf printer driver and ghostscript.

* ```pip3 install -r requirements.txt```
* install [Bullzip PDF Printer](https://www.bullzip.com/products/pdf/info.php), a feature-laden printer driver that supports cli control.
* install [Ghostscript](https://www.ghostscript.com/releases/index.html), the de facto open source pdf renderer.

You can run it from the command line ```python docxform2pdf.py "<input.dox>"``` or drag and drop your word file onto the script's icon.


### How does it work?
1. generate an input pdf from the docx file
2. use *ghostscript* on the input pdf to generate PNGs for each page
3. use the *opencv* imaging library to locate the coloured boxes within the PNGs
4. use *reportlab* to create a blank pdf canvas and position form fields corresponding to each coloured box
5. use *docx* to edit the original file, remove the red and green shading and save a "clean" docx
6. generate a clean pdf from the cleaned docx file
7. use *pdfrw* to merge the interactive canvas over the clean pdf

Steps 1 to 3 are required because docx files are effectively XML and have no notion of coordinates.

![steps](https://i.imgur.com/UHq5Oa0.png)

