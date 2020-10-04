# AIM OF THE PROJECT
This module wants to provide a compact way to export a Pandas' data frame to 
Excel, directly formatting the sheet. The goal is to obtain immediately a file
with a pleasant look which can be easily consulted.

This tool needs to offer a standard solution requiring as little code as
possible. At the same time, it should accommodate more demanding users through a significant number of optional parameters, allowing them to tailor the formatting to their needs.     

The formatting process will mainly affect the table's heading and index, while
adjusting columns' width and text orientation.

## V 1.0
### Code's design
Unlike for version 0.0, the class structure was used. The class's methods can
be seen as organized in four levels:
- **The method "to_custom_excel"**: it is the main method and the only one
serving a complete purpose.
- **Methods referred as "building blocks"**: methods conceived to 
serve a single purpose to enhance code simplicity.
- **Methods used in building blocks**: methods which are used across multiple
building blocks. They are sort of the building blocks of the building blocks.
- **Methods for debugging**: mainly methods which print out a group of attributes
just to check their values in case of undesired behaviors.

This levels are clearly denoted in the code by comments forming some sort of headings enclosed in dashed lines.

### Logic of the program
I briefly describe the sequence of methods within the main method "to_custom_excel".
#### File saving process
In this portion of the process a pandas data frame is taken and saved to excel.
This customize option goes beyond the "dataframe.to_excel" pandas method since
it automatically replaces/appends sheets to excel files based on existence logics.
1. Check if file name is acceptable
    - If extension is missing, adding and notify the user
1. Check if in the current directory a workbook with the same file name as the 
one provided exists
    - If so, check if it already contains a sheet named as the provided sheet name
1. Save the file:
    - If the workbook did not exist, create it
    - If it existed but it contained only a sheet called as the provided sheet
    name, replace it
    - If it existed and it contained no sheet called as the provided sheet name,
    add a sheet to the workbook.
#### Table dimensions analysis
In this segment the pandas data frame is analyzed to deduce which set of cells of
the sheet will form the header, which the index and which the body, including the
potential presence of multicolumns/multiindexes. A list of cells in the form 
["A1", "A2"] is created for each of the mentioned components.
1. Depth of multindex/multicolumns is calculated
1. For header, index and body:
    1. Imagining it as a rectangle of cells, the coordinates of the top left cell
    and of the bottom left are found.
    1. All the other cells of the rectangle are deduced.
        - (_Ex: If a rectangle extends from A1 to B2, no other information is
        required to know that it contains the cells A1, A2, B1, B2_)
#### Table formatting
Finally, the desired formatting is applied to the three areas. The large number
of parameters accepted by the class is justified by the generous degree of customization allowed. The logic of this process is based on a three-levels system
of styles: main, light formatting and body. The parameters of these styles can
be modified and then matched to header, index and body.
1. Based on user inputs the styling attributes are updated.
1. For all the options:
    1. Check if they are enabled by the user
    1. Apply the formatting to the portion of table considered
1. Save the workbook

### Potential features to be added in the future
1. Export table to jpg/png
1. Automatic column width

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
