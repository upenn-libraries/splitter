## README for `splitter`

This is a small, command-line-based utility for splitting large Excel or CSV files into multiple, smaller XSLX files, to create Colenda-friendly spreadsheets for work with [Bulwark](https://gitlab.library.upenn.edu/repo/bulwark).

## Requirements
This script requires Ruby 2.2.5 or later, and the [rubyXL gem](https://github.com/weshatheleopard/rubyXL).

## Instructions for Use

To use this script, navigate to the directory containing the script and execute the following command:
```bash
ruby splitter PATH_TO_CSV_OR_XLSX INDEX_OF_IDENTIFIER BATCH_SIZE
```
The values of the command-line arguments should correspond to the following:
* ```PATH_TO_CSV_OR_XLSX``` - The [absolute file path](https://www.computerhope.com/jargon/a/absopath.htm) to the Excel or CSV file to be processed.
* ```INDEX_OF_IDENTIFIER``` - The column number in the CSV file that contains the value that should be considered the row identifier.  This is used to generate the XLSX filename.
* ```BATCH_SIZE``` - The number of rows per XLSX file to be written.