# CTM

## Main files in order of operation:
### 1) CTMdata.ahk - extract distributor data from the CTM tool
### 2) ctmdatahandler.pyw - process the extracted data into a usable table for easy copying into the work table. (Automatically executed at end of CTMdata.ahk)

_______________________________

## Required files for operation:
### * distalias.txt - full list of distributors and their alias to be used as a final table substitute
### * distignore.txt - list of distributors to be ignored and not put in the final table
### * distspecial.txt - list of distributors to be marked as special case in the final table
### * CTMdataStorage.xlsx - empty excel file
