#!/bin/bash
#Usage sh PdfToText.sh PdfDirectory
if [ "$#" -ne 1 ]; then
    echo "Usage sh PdfToText.sh PdfDirectory"
	exit 1
fi

cd $1

for file in *.pdf;
	do gs -dBATCH -dNOPAUSE -sDEVICE=txtwrite -sOutputFile="$file.txt" "$file";
	done; 
for file in *.pdf.txt;
	do mv $file `basename $file .pdf.txt`.txt;
	done;