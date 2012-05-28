xls2csv
=======

A script to convert excel spreadsheets to CSV files on OSX, using AppleScript
and Excel.

Usage
-----

1. Open the Excel document that you want to convert to `CSV`
2. Make sure that the excel document only contains a header (if needed) and the data
3. Open the script `xls2csv.applescript` in the "AppleScript Editor" application
4. Choose "run"
5. Follow the dialog boxes

Third Parties
-------------

### Adobe Indesign - Data Merge
When doing data merges in Indesign the encoding needs to be `UTF-16` to handle
Unicode characters correctly.

### Perl module - Tie::Handle::CSV
This one uses `UTF-8` without a problem.