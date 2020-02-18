# Worksheets Consolidator
Script to consolidate excel sheets

Usage: <exec> -i <input files path> -o <output file>

I would like a script to consolidate tabs in several excel files to one excel file. I would like to be able to provide the sharepoint locations of all the excel files, run the script to create a single file that includes all sheets from on the excel files.

The final file should be a single excel file with several sheets.
An example is below with 4 input files and a total of 10 sheets. The output should be 1 files and 10 sheets.

excelFile1: EF1_Sheet1, EF1_Sheet2, EF1_Sheet3
excelFile2: EF2_Sheet1, EF2_Sheet2
excelFile3: EF3_Sheet1, EF3_Sheet2, EF3_Sheet3, EF3_sheet4
excelFile4: EF4_Sheet1

NEW_File: EF1_Sheet1, EF1_Sheet3, EF1_Sheet3, EF2_Sheet1, EF2_Sheet2, EF3_Sheet1, EF3_Sheet2, EF3_Sheet3, EF3_sheet4, EF4_Sheet1

The locations of the input excel files will be different from file to file. The input files are stored on a sharepoint server. The sharepoint link will be provided for each file to run the script. The number of input excel files could vary.

Need completed within 24hrs
Should work on Mac or PC
