# tangible
C# library for exporting data tables to an Office Open XML spreadsheet.
It also contains a TSV format reader useful for testing.
The TSV reader can identify column data types.

The library is dependent upon these packages:
1. DocumentFormat.OpenXml  (built initially with version 2.15.0)
2. LumenWorksCsvReader     (built initially with version 4.0.0)

## `WorkbookWriter`

Office Open XML spreadsheet writer.
Writes multiple `System.Data.DataTable` tables each
to a unique worksheet.

## `TsvReader`

Tab-separated-variable text file reader. 
Reader identifies column data types.
