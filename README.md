# VBA-challenge
Module 2 Challenge 

'The files within this repository are described below:

'StockComparisonSummary.vb'
    'The this VB file contains the main calculation subroutine which creates summary data from a dataset
    'of stocks, creates a separated table with this summarized data, then creates another separated table showing the ticker of the
    'stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume. 

        'This VB file is assigned to a button called "Run Summary Comparison" which is present on each of the worksheets in this workbook.

        'Within this code, I did utitlize the Stack Overflow and MrExcel websites when syntax was not functioning properly for some of the functions I was calling.
    'Functions borrowed (but not copied verbatim) from the Stack Overflow and MrExcel websites are:
        'line 39 - inserting entire columns
        'line 93 - Formatting columns and cells to correct numberformat
        'lines 112,117,122 - using the Application.Max and Application.Min functions to find largest and smallest numbers in a range
        'lines 131-135 - formatting column headers to bold font, changing font size, and surrounding ranges with borders

'ClearData.vb'
    'The other VB file in this repository is titled "ClearData.vb".  This is assigned to a buttom called "Clear Data" on each of the worksheets in the workbook.
    'Pressing this button and calling this subroutine clears the summary data on all four worksheets

    'Within this code, I did utilize the Stack Overflow website when syntax wasn't working correctly in my code.
    'Functions utilized (but not copied verbatim) include:
        'line 9 - Using the .Delete function to delete entire columns in a range
        'line 12 - properly using the .ColumnWidth function to resize a column

'Multiple_year_stock_data.xlsm'
    'This file is the excel worksheet which contains the data, visual basic code, and buttons to run this code to process/summarize the data 
      
