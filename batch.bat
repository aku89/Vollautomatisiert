@echo off
"C:\Program Files\R\R-3.6.3\bin\R.exe" CMD BATCH Workbook2.R
md Einzelteile
mv cars.xlsx Einzelteile/
mv CO2.xlsx Einzelteile/
mv Emissions.xlsx Einzelteile/
mv cars.png Einzelteile/
mv plants.png Einzelteile/
mv geographic.png Einzelteile/
