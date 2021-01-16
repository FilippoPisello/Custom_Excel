# Table of Contents <!-- omit in toc -->
- [1. Overview](#1-overview)
  - [1.1. Goal of the project](#11-goal-of-the-project)
  - [1.2. Code structure](#12-code-structure)
- [2. Required packages](#2-required-packages)
- [3. Class elements](#3-class-elements)
  - [3.1. Arguments](#31-arguments)
  - [3.2. Attributes](#32-attributes)
  - [3.3. Class variables](#33-class-variables)
  - [3.4. Properties](#34-properties)
  - [3.5. Methods](#35-methods)
    - [3.5.1. to_custom_excel()](#351-to_custom_excel)
    - [3.5.2. new_style()](#352-new_style)
  - [3.6. ExcelStyle objects](#36-excelstyle-objects)

# 1. Overview
The CustomExcel class allows the user to export a pandas data frame in excel directly applying some styling. The logic of the formatting is the following: some styles objects are created and applied to the different parts of the table. These are: index, header and body.

CustomExcel is a subclass of the Spreadsheet class, whose code and doc can be found [at this link](https://github.com/FilippoPisello/Spreadsheet).

## 1.1. Goal of the project
The CustomExcel class wants to provide a compact way to export a Pandas' data frame to Excel, directly formatting the sheet. The goal is to obtain immediately a file with a pleasant look which can be easily consulted. To understand its use, one can see this as an enhanced form of the pandas built-in method df.to_excel().

This tool offers a standard solution requiring very little code. At the same time, it accommodates more demanding users through a significant number of optional parameters, allowing them to tailor the formatting to their needs.

## 1.2. Code structure
The project is made of two files:
- **excel_formatting.py**: the main module containing the CustomExcel class.
- **excel_style.py**: an helper module, containing the ExcelStyle class which is invoked within the CustomExcel one.

Within the main file, the code is designed to convey a hierarchical division of the class' methods. Different code portions are introduced by comment blocks which are made of two compact lines of "#" having a number in between.

The sections are structured as follows:
- **Part 1**: main methods
- **Part 2**: worker methods
  - **Part 2.1**: methods used in attributes
  - **Part 2.2**: chain of methods used in to_custom_excel() for the file saving process
  - **Part 2.3**: chain of methods used in to_custom_excel() for formatting

The methods' are ordered so that if method B is invoked by method A, then B will be below A. This structure should hopefully help the reader to understand how the simple pieces are assembled to construct more complex items.

The individual methods are designed to follow as closely as possible the **single-responsibility principle**. Some of them are **protected** - their name is preceded by an underscore. This is done for two main reasons. First, not to clutter excessively the help text of the class, since protected methods are not displayed in this output. This allows the focus to be kept on the most important elements. Second, protection is in place to clearly signal which are the methods meant to be used only internally.

# 2. Required packages
CustomExcel requires the following custom module created by me:
- **Spreadsheet** [_link_](https://github.com/FilippoPisello/Spreadsheet)

This class relies on the following built-in packages:
- **os**
- **typing**
- **string** _[by Spreadsheet class]_

And on the following additional packages:
- **pandas**
- **openpyxl**
- **numpy** _[by Spreadsheet class]_

# 3. Class elements
## 3.1. Arguments
The class inherits five arguments from the Spreadsheet class:
- **dataframe** : pandas data frame object (mandatory)
  - The pandas data frame to be considered.
- **keep_index** : Bool, default=False
  - If True, it is taken into account that the first column of the spreadsheet will be occupied by the index. All the dimensions will be adjusted as consequence.
- **starting_cell**: str, default="A1"
  - The cell where it will be placed the top left corner of the data frame.
- **correct_lists**: Bool, default=False
  - If True, the lists stored as the data frame entries are modified to be more readable in the traditional spreadsheet softwares. This happens in four ways. (1) Empty lists are replaced by missing values. (2) Missing values are removed from within the lists. (3) Lists of len 1 are replaced by the single element they contain. (4) Lists are replaced by str formed by their elements separated by commas.

There are then five native arguments of the class:
- **file_name** : str (mandatory)
  - Name of the file to be exported.
- **sheet_name**: str, default="Sheet1"
  - Label of the sheet which should be target of the data upload. If a sheet with the provided name does not exist, it will be created.
- **header_style**: str or ExcelStyle object, default="strong"
  - The style to be given to the table's header. If str it should be one of the following keywords: "strong", "light", "plain". Custom ExcelStyle obj can be created through the new_style() method.
- **index_style**: str or ExcelStyle object, default="light"
  - The style to be given to the table's index. If str it should be one of the following keywords: "strong", "light", "plain". Custom ExcelStyle obj can be created through the new_style() method.
- **body_style**: str or ExcelStyle object, default="plain"
  - The style to be given to the table's body. If str it should be one of the following keywords: "strong", "light", "plain". Custom ExcelStyle obj can be created through the new_style() method.

## 3.2. Attributes
The CustomExcel object inherits four attributes from the Spreadsheet class:
- **self.df** : pandas data frame object
- **self.keep_index** : Bool
- **self.skip_rows**: int
- **self.skip_cols**: int

There are then seven native attributes:
- **self.file_name**: str
- **self.sheet_name**: str
- **self.workbook**: None or openpyxl.workbook object
- **self.sheet**: None or openpyxl.sheet object
- **self.header_style**: ExcelStyle object
- **self.index_style**: ExcelStyle object
- **self.body_style**: ExcelStyle object

## 3.3. Class variables
There are three built-in class variables which correspond to the three predefined styles:
- **self.strong_formatting**: ExcelStyle object
- **self.light_formatting**: ExcelStyle object
- **self.plain_formatting**: ExcelStyle object

More on this on section 3.6.

## 3.4. Properties
The CustomExcel object inherits eight properties from the Spreadsheet class:
- **self.indexes_depth** : [int, int]
- **self.header_coordinates** : [[int, int], [int, int]]
- **self.index_coordinates**: [[int, int], [int, int]]
- **self.body_coordinates**: [[int, int], [int, int]]
- **self.header**: SpreadsheetElement object
- **self.index**: SpreadsheetElement object
- **self.body**: SpreadsheetElement object
- **self.table**: SpreadsheetElement object

Details on these can be found at this [_link_](https://github.com/FilippoPisello/Spreadsheet).

## 3.5. Methods
This section just includes the methods which are meant to be accessed by the user, thus in part 1. For further info on the worker methods please consult their docstrings.

### 3.5.1. to_custom_excel()
Exports pandas data frame to Excel with some formatting.

First, the pandas data frame is saved to Excel with the given file and sheet name. Three scenarios for this process: (1) if no workbook under the given name exists, it gets created. (2) If the workbook exists but there is no sheet with given name, a new sheet gets appended. (3) If both workbook and sheet exists, the latter gets overwritten.

Then the table gets formatted: the styles for header, index and body are applied. These are chosen through the object attributes obj.header_style, obj.index_style and obj.body_style.

**Arguments**
- **custom_width**: int, default=20
  - Width to be set for all the columns of the spreadsheet, expressed in points.
- **check_file_name**: Bool, default=True
  - If True, the program will try to add the ".xlsx" file extension at the end of the file name if it is not present, to avoid errors. If False no control is carried out.

### 3.5.2. new_style()
Returns an ExcelStyle object with the formatting properties chosen by the user.

This object can then be assigned to the desired table part, modifying the value of the parameters self.header_style, self.index_style and self.body_style.

**Arguments**
- **fill_color**: str, default=None
  - Fill color of the cells. If None, no fill color is applied.
- **font_color**: str, default="000000"
  - Font color of the cells. Default color is black.
- **font_size**: int, default=11
  - Size of the cell font.
- **font_bold**: Bool, default=False
  - If True cell text is bold.
- **alignment**: str, default="center"
  - Horizontal alignment of the text content. It can be either "center",
"right" or "left".

## 3.6. ExcelStyle objects
The ExcelStyle objects are created to concisely bring together all the formatting properties to be applied to a single portion of the table, defining a style.

The ExcelStyle object has five attributes:
- **self.fill_color**: str, default=None
  - Fill color of the cells. If None, no fill color is applied.
- **self.font_color**: str, default="000000"
  - Font color of the cells. Default color is black.
- **self.font_size**: int, default=11
  - Size of the cell font.
- **self.font_bold**: Bool, default=False
  - If True cell text is bold.
- **self.alignment**: str, default="center"
  - Horizontal alignment of the text content. It can be either "center",
"right" or "left".

There are three predefined styles which are built-in into the CustomExcel class. Their details are the following:

||self.strong_formatting|self.light_formatting|self.plain_formatting|
|---|---|---|---|
|**Keyword**|"strong"|"light"|"plain"|
|**Fill color**|Blue (#0066cc)|Grey (#b2beb5)|None|
|**Font color**|White|Black|Black|
|**Font size**|12|11|11|
|**Bold**|Yes|No|No|
|**Alignment**|center|left|center|