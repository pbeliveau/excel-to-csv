# Excel to CSV
# Why convert to CSV?

Sometime, I come upon a dataset that are only available in .XLSX format (or similar
Excel format). There is a few reason why converting it to a text file is beneficial:

1. Text files are forever! They can be edited easily without specific softwares,
without Excel in this case;
2. Some SaS languages don't have the ability to pull data from Excel files, and/or, 
it's just easier to use a text file format;
3. Have you tried doing version control of a dataset in an Excel format? Using git
with a .CSV file is straightforward.

## Can't I just export to .CSV using the built-in Excel way?

Yes, you can. If you have a large dataset, try this code. It will be faster.
As a matter of fact, using vba to export is faster than comparable languages
like Python.

Not enough to convince you? You can alter this code to export in various ways.
Don't want a comma to separate column? Change it to a "|" or whatever.
 
# How-to use this 

1. Open the Excel workbook you want to use this code on;
2. Press ALT+F11 to open the Visual Basic Editor in Excel;
3. In the menu, select "Insert" and "Module";
4. Copy-paste the content of the file "csv_module.bas" to the module in Excel;
5. From the Visual Basic Editor, select the module, press F5 to use it.