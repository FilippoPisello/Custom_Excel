# AIM OF THE PROJECT
This module wants to provide a compact way to export a Pandas' data frame to
Excel, directly formatting the sheet. The goal is to obtain immediately a file
with a pleasant look which can be easily consulted.

This tool needs to offer a standard solution requiring as little code as
possible. Moreover, it should accommodate more demanding users through
a significant number of optional parameters, allowing them to tailor the
formatting to their needs.     

The formatting process will mainly affect the table's heading and index, while
adjusting columns' width and text orientation.

## V 0.0
### Code's design
The idea is to structure the code as follows:
- **One main function**: it will be the one to be used once the program is
complete. It will mainly work as the architecture collecting the sub-functions and
regulating their alternation. It should end up providing an easy overlook on the
processes.
- **A number of small sub-functions**: these functions will be conceived to serve
a single purpose, thus being as concise as possible.

To hope is to enhance in this way code readability and to make future edits
easier.

### Logic of the program
- main function:
    1. Intake pandas data frame to be transformed in a formatted excel
        1. Extract information on:
            1. Headers' length & depth
            1. Index length & depth
        1. Intake from the user information on:
            1. Final file name
            1. Final sheet name (these two options should allow the user to add
            sheets to existing workbooks)
            1. Relevance of the index
    1. Export the data frame to a row excel using pandas.to_excel
    1. Open back the excel with openpyxl
        1. Create the lists appending the cells (ex: [A1, A2, ...]) which form:
            1. The header:
                1. If multicolumns of level n, include n rows
                1. If index is present start from column m, where m is the level
                of multiindex
                    1. Otherwise start from column 0 (A).
                1. Given index of lenght i, include until column i+m.
            1. The index:
                1. Check if index is present, if yes:
                    1. If multiindex of level n, include m columns
                    1. If header is present start from row n, where n is the level
                    of multicolumns
                        1. Otherwise start from row 0 (1).
                    1. Given columns of length c, include until row c+n.
                1. If no, check if it is still relevant (ex: contains table's
                key and needs to be highlighted):
                    1. Same as above
                1. If no, skip.
            1. The body
                1. Given multicolumns of level n, start from row n+1.
                1. Given multiindex of lvel m, start from row m+1.
                1. Given columns of length c, include until row c+n.
                1. Given index of lenght i, include until column i+m.
        1. Apply formatting to:
            1. Columns (section to be improved):
                1. Set a pre-established width (can be made adaptive)
                1. Align center
            1. The header:
                1. Colour
                1. Bold
                1. Align center
            1. The index if present as index:
                1. Same as previous
            1. The index if present as first highlighted column:
                1. Lighter colour
                1. Align center
        1. Save the workbook
        1. Option to save the table as image/pdf (?)
