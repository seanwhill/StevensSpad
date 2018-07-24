#!/bin/bash
for file in *.pdf; do gs -dBATCH -dNOPAUSE -sDEVICE=txtwrite -sOutputFile="$file.txt" "$file"; done; for file in *.pdf.txt; do mv $file `basename $file .pdf.txt`.txt; done;