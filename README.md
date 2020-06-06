# pdf2docx
A Python script that converts pdf files to docx. 

## How it works
The script ```pdf2docx.py``` takes a pdf file as input. It converts the pdf file to png files then inserts each png file to a docx file. This means that the output cannot be edited since it is just a series of pasted png files. 

## Prerequisites 
To use the script, the following programs should be installed in your terminal: 
- Python 3
- [PyPDF2](https://pypi.org/project/PyPDF2/) 
- [Wand](http://docs.wand-py.org/en/0.6.1/) 
- [python-docx](https://github.com/python-openxml/python-docx) 

## Installation 
Just clone this repo in any directory you choose, then just run the ```pdf2docx.py``` script. 

## Usage
The script ```pdf2docx.py``` takes one required parameter which is the filename of the input Pdf file. The optional arguments are pagesize, orientation, and output filename. 

usage: pdf2docx.py [-h] [-s PAPERSIZE]
                   [-r ORIENTATION] [-o OUTPUT]
                   [-v]
                   inputPdf

positional arguments:
  inputPdf              Filename of the input
                        PDF.

optional arguments:
  -h, --help            show this help message
                        and exit
  -s PAPERSIZE, --papersize PAPERSIZE
                        Papersize. Valid values
                        are a4, letter, and
                        legal.
  -r ORIENTATION, --orientation ORIENTATION
                        Page orientation. Valid
                        values are portrait and
                        landscape.
  -o OUTPUT, --output OUTPUT
                        Output filename. Must be
                        docx.
  -v, --version         show program's version
                        number and exit


