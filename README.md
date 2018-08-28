# StevensSpad
Python script that extracts data from PDFS and places it into an Excel sheet for Stevens accredidation purposes

## Prerequisits
GhostScript Version 9.23

```
wget https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs923/ghostscript-9.23.tar.gz
tar xvf ghostscript-9.23.tar.gz
cd ghostscript-9.23
./configure
make
make install
```

Pandas for python 

```
pip install pandas
```

xlsxwriter

```
pip install xlsxwriter
```

xlrd
```
pip install xlrd
```

## Pre Processing

Create a folder for whichever year you are covering and move all of the Pdfs you are using into that folder

run the shell script to convert all of the pdfs into texts using ghostscript

``` sh PdfToTxt.sh pdfDirectory ```

modify the constants at the top of the two python files TXTDIRECTORY and PDFDIRECTORY to whichever your directory name is

## Execution

Execute ExcelSetup.py

Exectue pdfParser.py
